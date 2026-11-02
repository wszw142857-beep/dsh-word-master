/**
 * dsh-word-master — Word Master 的 DSH 原生封装。
 *
 * 本包做三件事，且都只在「选中 Word Master 预设的会话」里生效：
 *
 *   1. **工具集**：把 skill 里的 Python 脚本包成 DSH 工具（`word_env` /
 *      `word_spec` / `word_build` / `word_gate` / `word_render` / `word_export` …），
 *      模型直接调用，不用先读 SKILL.md 再拼命令行。
 *
 *   2. **skill provider**：把 SKILL.md 挂进 skill 目录，提供工具**无法**承载的
 *      部分 —— 规格冻结闸门、逐节推进、质检阻断、看图验收这些流程纪律。
 *
 *   3. **偏好记忆**：本地偏好库 + 每轮系统提示词注入。这是与 ppt-master 的关键
 *      差别：写作约束不能靠模型"想起来"，必须每轮都在场。
 *
 * 关于 `${SKILL_DIR}`
 * -------------------
 * DSH 的 filesystem skill provider 解析 frontmatter 后原样返回正文，不做变量替换。
 * 而 SKILL.md 里要写 `${SKILL_DIR}/scripts/xxx.py` 这种可复制的命令行，所以本
 * provider 在 `get()` 阶段把它替换成装好之后的真实绝对路径。
 *
 * 关于 rank
 * ---------
 * 默认 600 = `BUNDLED_SKILL_RANK`，低于用户级 `~/.agents/skills`（500）。机器上若
 * 已有手工调试副本，那份继续优先 —— 刻意如此。想让插件压过手工副本，把
 * config.rank 设为小于 500。
 *
 * @module dsh-word-master
 */

import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { createToolDefinitions } from './tools.js';
import { loadPreferences, renderPreferencesSection, preferencesPath } from './preferences.js';

/** 本 provider 在 `ctx.skills` 里注册的名字。 */
const PROVIDER_NAME = 'word-master';

/** skill 载荷在包内的位置（`<pkg>/skills/`）。 */
const SKILLS_ROOT_URL = new URL('../skills/', import.meta.url);

/** 随包提供的两个 skill。 */
const SKILLS = [
	{ name: 'word-master', dir: 'word-master' },
	{ name: 'word-master-dsh', dir: 'word-master-dsh' },
];

/** 官方 `BUNDLED_SKILL_RANK`。 */
const BUNDLED_SKILL_RANK = 600;

/**
 * skill 根目录的绝对路径，统一成正斜杠。
 *
 * @param dir skill 在包内的目录名。
 * @returns 该 skill 的根目录，末尾不带分隔符。
 */
function resolveSkillDir(dir = 'word-master') {
	return fileURLToPath(new URL(`${dir}/`, SKILLS_ROOT_URL)).replace(/[\\/]+$/, '').replaceAll('\\', '/');
}

/** 上游 skill 载荷的绝对路径（`word-master-dsh` 里用它引用详细规范）。 */
function resolvePayloadDir() {
	return resolveSkillDir('word-master');
}

/**
 * 解析 YAML frontmatter 的最小子集。
 *
 * 刻意不依赖 yaml 包：本包零依赖，装到任何 profile 都不会因为缺第三方模块而
 * 加载失败。SKILL.md 的 frontmatter 只用到 `key: value`、折叠块与一层嵌套 map。
 *
 * @param text frontmatter 内部文本（不含 `---` 分隔行）。
 * @returns 解析出的键值对象。
 */
function parseYamlSubset(text) {
	const root = {};
	const stack = [{ indent: -1, node: root }];
	const lines = text.split(/\r?\n/);

	for (let i = 0; i < lines.length; i += 1) {
		const line = lines[i];
		if (line.trim() === '' || line.trimStart().startsWith('#')) continue;

		const indent = line.length - line.trimStart().length;
		const content = line.trim();
		while (stack.length > 1 && indent <= stack[stack.length - 1].indent) stack.pop();
		const parent = stack[stack.length - 1].node;

		const colon = content.indexOf(':');
		if (colon < 0) continue;
		const key = content.slice(0, colon).trim();
		const rest = content.slice(colon + 1).trim();
		if (key === '') continue;

		if (/^[>|][+-]?$/.test(rest)) {
			const folded = rest[0] === '>';
			const block = [];
			let blockIndent = -1;
			let j = i + 1;
			for (; j < lines.length; j += 1) {
				const raw = lines[j];
				if (raw.trim() === '') {
					block.push('');
					continue;
				}
				const currentIndent = raw.length - raw.trimStart().length;
				if (currentIndent <= indent) break;
				if (blockIndent < 0) blockIndent = currentIndent;
				block.push(raw.slice(Math.min(blockIndent, currentIndent)));
			}
			i = j - 1;
			const merged = folded ? block.join(' ').replace(/\s+/g, ' ').trim() : block.join('\n').trim();
			parent[key] = merged;
			continue;
		}

		if (rest === '') {
			const child = {};
			parent[key] = child;
			stack.push({ indent, node: child });
			continue;
		}

		if (rest.startsWith('"') && rest.endsWith('"') && rest.length >= 2) {
			parent[key] = rest.slice(1, -1);
			continue;
		}
		if (rest.startsWith("'") && rest.endsWith("'") && rest.length >= 2) {
			parent[key] = rest.slice(1, -1);
			continue;
		}
		parent[key] = rest;
	}

	return root;
}

/**
 * 拆出 frontmatter 与正文。
 *
 * @param raw SKILL.md 原文。
 * @returns `{ attributes, body }`；没有 frontmatter 时 attributes 为空对象。
 */
function splitFrontmatter(raw) {
	const match = raw.match(/^---\r?\n([\s\S]*?)\r?\n---\r?\n?/);
	if (match === null) return { attributes: {}, body: raw };
	return { attributes: parseYamlSubset(match[1]), body: raw.slice(match[0].length) };
}

/**
 * 把正文里的占位符展开成真实绝对路径。
 *
 * @param body SKILL.md 正文。
 * @param skillDir 本 skill 根目录绝对路径（正斜杠）。
 * @param payloadDir 上游载荷根目录绝对路径；省略时现场解析。
 * @returns 展开后的正文。
 */
function expandSkillDir(body, skillDir, payloadDir) {
	const payload = payloadDir ?? resolvePayloadDir();
	return body.replaceAll('${SKILL_DIR}', () => skillDir).replaceAll('${PAYLOAD_DIR}', () => payload);
}

/**
 * 读取并解析包内携带的某个 skill。
 *
 * @param dir skill 在包内的目录名。
 * @returns skill 行对象。
 */
async function loadSkill(dir) {
	const fileUrl = new URL(`${dir}/SKILL.md`, SKILLS_ROOT_URL);
	const raw = await readFile(fileUrl, 'utf8');
	const { attributes, body } = splitFrontmatter(raw);
	const skillDir = resolveSkillDir(dir);

	const name = typeof attributes.name === 'string' && attributes.name !== '' ? attributes.name : dir;
	const description = typeof attributes.description === 'string' ? attributes.description : '';

	return {
		dir,
		name,
		description,
		whenToUse: typeof attributes.whenToUse === 'string' ? attributes.whenToUse : undefined,
		metadata: typeof attributes.metadata === 'object' && attributes.metadata !== null ? attributes.metadata : undefined,
		content: expandSkillDir(body, skillDir),
		path: fileURLToPath(fileUrl),
		resourceBase: { kind: 'directory', path: skillDir },
		locator: fileUrl,
	};
}

/**
 * 构造 skill provider。
 *
 * @param providerName 注册到 `ctx.skills` 的 provider 名。
 * @param rank 优先级；数字小的赢。
 * @returns provider 对象。
 */
function createProvider(providerName, rank) {
	let candidates = null;

	return {
		name: providerName,

		async list() {
			const rows = await Promise.all(SKILLS.map((entry) => loadSkill(entry.dir)));
			candidates = new Map(rows.map((row) => [row.name, row]));
			return rows.map((row) => ({
				name: row.name,
				description: row.description,
				invocation: { modelInvocable: true, userInvocable: true },
				provider: providerName,
				source: 'bundled',
				resourceBase: row.resourceBase,
				rank,
				locator: row.locator,
				metadata: row.metadata,
			}));
		},

		/**
		 * 加载正文。
		 *
		 * ⚠ 调用约定：registry 传进来的第一个参数是**候选对象**（list() 返回的那条），
		 * 不是名字字符串。这里同时容忍字符串，便于手工调用与测试。
		 *
		 * @param candidate 候选对象或 skill 名。
		 * @returns 完整 skill 定义，或 undefined。
		 */
		async get(candidate) {
			const wanted =
				typeof candidate === 'string'
					? candidate
					: typeof candidate?.name === 'string' && candidate.name !== ''
						? candidate.name
						: SKILLS[0].name;

			let row = candidates?.get(wanted);
			if (row === undefined) {
				const rows = await Promise.all(SKILLS.map((entry) => loadSkill(entry.dir)));
				candidates = new Map(rows.map((r) => [r.name, r]));
				row = candidates.get(wanted);
			}
			if (row === undefined) return undefined;

			return {
				name: row.name,
				description: row.description,
				whenToUse: row.whenToUse,
				metadata: row.metadata,
				invocation: { modelInvocable: true, userInvocable: true },
				provider: providerName,
				source: 'bundled',
				resourceBase: row.resourceBase,
				content: row.content,
				path: row.path,
			};
		},
	};
}

/** Cordis 插件名。 */
export const name = 'dsh-word-master';

/**
 * 依赖注入：工具注册表 + skill 注册表 + 系统提示词。
 *
 * `systemPrompt` 是可选依赖：装了才有偏好自动注入，没装时工具与 skill 照常工作。
 * 所以不写进 inject（那样没 systemPrompt 的部署会一直 PENDING），改用 ctx.inject
 * 的延迟注册。
 */
export const inject = ['tools', 'skills'];

/**
 * 注册工具集、skill provider，以及偏好注入。
 *
 * @param ctx Cordis 上下文。
 * @param config 配置。
 * @param config.providerName 覆盖 skill 注册名（默认 `word-master`）。
 * @param config.rank 覆盖 skill 优先级（默认 600）。
 * @param config.tools 是否注册工具集（默认 true）。
 * @param config.skill 是否注册 skill provider（默认 true）。
 * @param config.injectPreferences 是否每轮注入偏好（默认 true）。
 * @param config.preferenceOrder 偏好小节在系统提示词里的排序权重（默认 3200）。
 */
export function apply(ctx, config = {}) {
	const wantTools = config.tools !== false;
	const wantSkill = config.skill !== false;
	const wantPreferences = config.injectPreferences !== false;

	if (wantTools) {
		const skillDir = resolveSkillDir();
		for (const definition of createToolDefinitions({ skillDir })) {
			ctx.tools.register(definition);
		}
	}

	if (wantSkill) {
		const providerName =
			typeof config.providerName === 'string' && config.providerName !== '' ? config.providerName : PROVIDER_NAME;
		const rank = Number.isInteger(config.rank) && config.rank > 0 ? config.rank : BUNDLED_SKILL_RANK;
		ctx.skills.registerProvider(() => createProvider(providerName, rank));
	}

	if (wantPreferences) {
		const order = Number.isInteger(config.preferenceOrder) ? config.preferenceOrder : 3200;
		ctx.inject(['systemPrompt'], (promptCtx) => {
			promptCtx.effect(() =>
				promptCtx.systemPrompt.section({
					name: 'word-master:preferences',
					order,
					interpolate: false,
					text: () => {
						const loaded = loadPreferences();
						return renderPreferencesSection(loaded.preferences);
					},
				}),
			);
		});
	}
}

/** 供诊断使用：暴露偏好路径，便于排查"为什么注入没生效"。 */
export const diagnostics = { preferencesPath, loadPreferences };

/** 供测试与自检使用。 */
export const __internals = {
	parseYamlSubset,
	splitFrontmatter,
	expandSkillDir,
	resolveSkillDir,
	loadSkill,
	createToolDefinitions,
	BUNDLED_SKILL_RANK,
};
