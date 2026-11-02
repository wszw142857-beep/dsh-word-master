/**
 * Word Master 工具集：把 skill 里的 Python 脚本逐个包成 DSH 工具。
 *
 * 移植自 dsh-ppt-master 的工具层约定
 * ---------------------------------
 * 1. **一个脚本一个工具，参数直映射真实 CLI**。
 * 2. **`extra_args` 逃生口**：凡这里没建模的旗标都能原样透传，脚本能力不会因为
 *    封装而丢失。
 * 3. **每个工具都带 `cwd`**：脚本可能在别的项目目录里跑，不能假设进程 CWD。
 * 4. **不重写脚本逻辑**：这层只做「参数 → argv」翻译与执行，脚本是唯一事实来源。
 * 5. **不做流程编排**：流程纪律留在 SKILL.md，工具只负责"把这一步做掉"。
 * 6. **零运行时依赖**：不 import 任何 `@deepseek-ai/dsh-*`（见 README 说明）。
 *
 * @module dsh-word-master/tools
 */

import { resolve } from 'node:path';
import { runScript, runPythonCode, environmentReport, formatEnvironmentReport } from './runner.js';
import { loadPreferences, renderPreferencesSection, preferencesPath } from './preferences.js';

/**
 * 把「作者向」schema 编译成原生 JSON Schema。
 *
 * 两条硬规则（原生 schema 比作者向写法严格）：
 *   1. `required` 必须是对象级字符串数组，不能写在单个属性上；
 *   2. `additionalProperties` 出现时必须是显式布尔值。
 *
 * @param node 作者向 schema 节点。
 * @returns 原生 JSON Schema 节点。
 */
function toRawSchema(node) {
	const out = {};
	if (typeof node.description === 'string') out.description = node.description;
	if (typeof node.title === 'string') out.title = node.title;

	if (Array.isArray(node.enum)) out.enum = [...node.enum];
	if (Object.hasOwn(node, 'const')) out.const = node.const;

	if (node.type === 'object' || node.properties !== undefined) {
		out.type = 'object';
		out.additionalProperties = false;
		const properties = {};
		const required = [];
		for (const [key, child] of Object.entries(node.properties ?? {})) {
			properties[key] = toRawSchema(child);
			if (child.required === true) required.push(key);
		}
		out.properties = properties;
		if (required.length > 0) out.required = required;
		return out;
	}

	if (node.type === 'array') {
		out.type = 'array';
		if (node.items !== undefined) out.items = toRawSchema(node.items);
		return out;
	}

	out.type = node.type ?? 'string';
	return out;
}

/** 把参数字典编译成原生对象 schema（不封口）。 */
function toRawParameters(params) {
	const properties = {};
	const required = [];
	for (const [key, child] of Object.entries(params)) {
		properties[key] = toRawSchema(child);
		if (child.required === true) required.push(key);
	}
	const schema = { type: 'object', properties };
	if (required.length > 0) schema.required = required;
	return schema;
}

/** 统一的文本输出。 */
function textOutput(text) {
	return [{ type: 'text', text: text === '' ? '(无输出)' : text }];
}

/** 追加 `--flag value`。 */
function push(argv, flag, value) {
	if (typeof value === 'string' && value.trim() !== '') argv.push(flag, value);
}

/** 追加布尔旗标。 */
function pushFlag(argv, flag, value) {
	if (value === true) argv.push(flag);
}

/** 追加数字参数。 */
function pushNum(argv, flag, value) {
	if (Number.isFinite(value)) argv.push(flag, String(value));
}

/** 追加 `extra_args` 透传项。 */
function pushExtra(argv, extra) {
	if (Array.isArray(extra)) {
		for (const item of extra) {
			if (item !== null && item !== undefined) argv.push(String(item));
		}
	}
}

/** 通用的 cwd / extra_args 参数定义。 */
const COMMON_PARAMS = {
	cwd: { type: 'string', description: '执行时的工作目录；省略则用进程当前目录。' },
	extra_args: {
		type: 'array',
		items: { type: 'string' },
		description: '原样透传给脚本的参数（未在上面建模的旗标）。',
	},
};

/**
 * 构造全部工具定义。
 *
 * @param options 选项。
 * @param options.skillDir skill 载荷根目录绝对路径（正斜杠，末尾无分隔符）。
 * @returns 工具定义数组。
 */
export function createToolDefinitions({ skillDir }) {
	const defs = [];
	const scriptPath = (name) => resolve(skillDir, 'scripts', name);

	// ── 项目管理 ────────────────────────────────────────────────────────

	defs.push({
		name: 'word_project',
		description:
			'项目生命周期：init 建骨架、import-sources 收素材并登记、validate 按阶段查漏、info 看进度。' +
			'validate 会推断当前阶段（empty/draft/frozen/built/reviewed/delivered）并给出下一步该做什么，' +
			'还会检出「质检有 P0 未清」「draft 比质检报告新（报告过期）」「素材没转换」这几类隐性漏步。' +
			'exit 0=干净 / 2=项目不存在 / 3=有错误 / 4=只有警告。',
		parameters: toRawParameters({
			action: { type: 'string', enum: ['init', 'import-sources', 'validate', 'info'], required: true },
			project: { type: 'string', required: true, description: '项目目录绝对路径。' },
			title: { type: 'string', description: 'action=init 时的标题。' },
			sources: {
				type: 'array',
				items: { type: 'string' },
				description: 'action=import-sources 时的素材路径或 URL 列表。',
			},
			move: { type: 'boolean', description: 'action=import-sources 时移动而非复制本地文件。' },
			force: { type: 'boolean', description: 'action=init 时即使已存在也重建（不会清空 spec/draft）。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = [args.action, args.project];
			if (args.action === 'init') {
				push(argv, '--title', args.title);
				pushFlag(argv, '--force', args.force);
			} else if (args.action === 'import-sources') {
				const list = Array.isArray(args.sources) ? args.sources.filter((s) => typeof s === 'string' && s !== '') : [];
				if (list.length === 0) return '[ERR] action=import-sources 需要 sources 数组';
				argv.push(...list);
				pushFlag(argv, '--move', args.move);
			}
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('project.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 180000,
			});
			return result.output;
		},
	});

	// ── 素材转换 ────────────────────────────────────────────────────────

	defs.push({
		name: 'word_convert',
		description:
			'把素材转成 Markdown：PDF / Word / Excel / PowerPoint / HTML / 网页 URL。' +
			'写作起点往往不是 Markdown，而是别人给的 PDF 或旧 Word——先转再写。' +
			'PDF 按字号推断标题层级（可能误判，转完建议先扫一遍标题）；' +
			'docx 保留三级标题、列表、表格与图片；xlsx 每 sheet 一节；pptx 每页一节。',
		parameters: toRawParameters({
			input: { type: 'string', required: true, description: '源文件绝对路径，或网页 URL。' },
			out_dir: { type: 'string', required: true, description: '输出目录（Markdown 落在里面）。' },
			kind: {
				type: 'string',
				enum: ['auto', 'pdf', 'docx', 'xlsx', 'pptx', 'web', 'html'],
				description: '转换类型，默认 auto 按扩展名/URL 判断。',
			},
			stem: { type: 'string', description: '输出文件名（不含扩展名）；省略则按来源生成。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--input', args.input, '--out-dir', args.out_dir];
			push(argv, '--kind', args.kind);
			push(argv, '--stem', args.stem);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('convert.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 300000,
			});
			return result.output;
		},
	});

	// ── 视觉判据 ────────────────────────────────────────────────────────

	defs.push({
		name: 'word_visual',
		description:
			'视觉判据（排版 rubric）：从渲染后的 PDF 上量行距过紧、文字越界、段落重叠、标题孤行、字号漂移。' +
			'与 word_gate 分工：gate 查内容规则，本工具查排版观感。' +
			'退出码 0=干净 / 2=无 PDF / 3=有 P0 破损 / 4=只有 P1。' +
			'⚠️ 机器判据只覆盖可量化部分，**仍须 word_render 后逐页看图**。',
		parameters: toRawParameters({
			project: { type: 'string', required: true },
			pdf: { type: 'string', description: 'PDF 相对路径，默认 render/output.pdf。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--project', args.project];
			push(argv, '--pdf', args.pdf);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('visual_check.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 120000,
			});
			const verdict =
				result.exitCode === 0 ? 'PASS' : result.exitCode === 3 ? 'BLOCK' : result.exitCode === 4 ? 'WARN' : 'ERROR';
			return `[${verdict}] 退出码 ${result.exitCode}\n${result.output}`;
		},
	});

	// ── 模板 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_template',
		description:
			'套用现成 .docx 的版式（字体、字号、行距、页边距、标题样式）。' +
			'Word 的"模板"是样式集合而非版式骨架——真实需求是"单位/学校有规范文档，新文档要长得一样"。' +
			'四个动作：inspect 读画像、save 存进库、list 列出、apply 写进项目规格。' +
			'apply 之后 word_build 就会用这套版式，且 spec freeze 不会覆盖它（模板优先于偏好）。',
		parameters: toRawParameters({
			action: { type: 'string', enum: ['inspect', 'save', 'list', 'apply'], required: true },
			template: { type: 'string', description: 'action=inspect/save 时的 .docx 路径。' },
			id: { type: 'string', description: 'action=save/apply 的模板 id。' },
			project: { type: 'string', description: 'action=apply 时的项目目录。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = [args.action];
			if (args.action === 'inspect' || args.action === 'save') {
				if (typeof args.template !== 'string' || args.template.trim() === '') {
					return '[ERR] 需要 template 参数（.docx 路径）';
				}
				argv.push('--template', args.template);
			}
			if (args.action === 'save' || args.action === 'apply') {
				if (typeof args.id !== 'string' || args.id.trim() === '') {
					return '[ERR] 需要 id 参数（模板名）';
				}
				argv.push('--id', args.id);
			}
			if (args.action === 'apply') {
				if (typeof args.project !== 'string' || args.project.trim() === '') {
					return '[ERR] 需要 project 参数（项目目录）';
				}
				argv.push('--project', args.project);
			}
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('template.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
			});
			return result.output;
		},
	});

	// ── 体检 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_env',
		description:
			'体检 Word Master 环境：Python 解释器、python-docx/lxml、LibreOffice 渲染器，并给出缺失项的补齐命令。' +
			'开工前先跑这个，不要靠试错定位环境问题。',
		parameters: toRawParameters({}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute() {
			return formatEnvironmentReport(await environmentReport());
		},
	});

	// ── 偏好 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_pref',
		description:
			'读写用户写作偏好（标点、禁用词、段落推进、章节首句、方法章约束、版式、引文格式）。' +
			'用户提出任何"以后都这样写"的要求时用 set 落盘——只在对话里答应是没用的，' +
			'下一轮注入的还是旧偏好。开工前用 read 确认当前约束。',
		parameters: toRawParameters({
			action: {
				type: 'string',
				enum: ['read', 'set', 'memory_text', 'reset'],
				required: true,
				description: 'read 读全部；set 合并补丁；memory_text 生成可写入记忆空间的文字；reset 恢复内置默认。',
			},
			patch: {
				type: 'object',
				description: 'action=set 的补丁，与偏好文件同构；设 null 表示显式关闭某项。',
			},
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			if (args.action === 'read') {
				const loaded = loadPreferences();
				return JSON.stringify(
					{ ok: true, source: loaded.source, path: loaded.path, warning: loaded.warning, preferences: loaded.preferences },
					null,
					2,
				);
			}
			if (args.action === 'memory_text') {
				return renderPreferencesSection(loadPreferences().preferences, { maxChars: 12000 }).replace(
					/\n\n写作前先调.*$/s,
					'\n',
				);
			}
			if (args.action === 'reset') {
				const result = await runScript(scriptPath('prefs.py'), ['reset'], { cwd: args.cwd, signal: exec?.signal });
				return result.output;
			}
			if (typeof args.patch !== 'object' || args.patch === null || Array.isArray(args.patch)) {
				return '[ERR] action=set 需要 patch 对象。';
			}
			const argv = ['set', '--patch', JSON.stringify(args.patch)];
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('prefs.py'), argv, { cwd: args.cwd, signal: exec?.signal });
			return result.output;
		},
	});

	defs.push({
		name: 'word_pref_section',
		description: '把当前已生效的写作偏好渲染成系统提示词片段原文，用于核对"这些约束现在真的在生效"。',
		parameters: toRawParameters({}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute() {
			const loaded = loadPreferences();
			const section = renderPreferencesSection(loaded.preferences);
			return `偏好文件：${preferencesPath()}\n来源：${loaded.source}\n\n${section === '' ? '(自动注入已关闭)' : section}`;
		},
	});

	// ── 规格 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_spec',
		description:
			'写作规格：init 建草稿、read 读、freeze 冻结。冻结后构建与质检都以它为准，并快照当时的偏好与版式。' +
			'这是流水线第一道闸门——冻结前必须把标题、目标读者、章节结构跟用户确认清楚。',
		parameters: toRawParameters({
			action: { type: 'string', enum: ['init', 'read', 'freeze'], required: true },
			project: { type: 'string', required: true, description: '项目目录绝对路径。' },
			title: { type: 'string', description: 'action=init 时的文档标题。' },
			audience: { type: 'string', description: '目标读者。' },
			length: { type: 'number', description: '目标字数。' },
			tone: { type: 'string', description: '语气，如"正式、克制"。' },
			genre: { type: 'string', description: '文类：论文/报告/公文/说明书等。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = [args.action, '--project', args.project];
			if (args.action === 'init') {
				if (typeof args.title !== 'string' || args.title.trim() === '') {
					return '[ERR] action=init 需要 title。';
				}
				argv.push('--title', args.title);
				push(argv, '--audience', args.audience);
				pushNum(argv, '--length', args.length);
				push(argv, '--tone', args.tone);
				push(argv, '--genre', args.genre);
			}
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('spec.py'), argv, { cwd: args.cwd, signal: exec?.signal });
			return result.output;
		},
	});

	// ── 构建 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_build',
		description:
			'把项目里的 draft.md（Markdown 稿）构建成排版规范的 DOCX。支持三级标题、列表、引用、三线表、' +
			'粗体斜体、LaTeX 块级公式（转 Word 原生可编辑公式）。版式取自规格快照，脚本不写死字体字号。',
		parameters: toRawParameters({
			project: { type: 'string', required: true, description: '项目目录绝对路径。' },
			input: { type: 'string', description: '草稿文件名，默认 draft.md。' },
			output: { type: 'string', description: '输出文件名，默认 output.docx。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--project', args.project];
			push(argv, '--input', args.input);
			push(argv, '--output', args.output);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('build_docx.py'), argv, { cwd: args.cwd, signal: exec?.signal });
			return result.output;
		},
	});

	// ── 质检门 ──────────────────────────────────────────────────────────

	defs.push({
		name: 'word_gate',
		description:
			'质检闸门（唯一有否决权的一步）：全角括号、禁用词、标题跳级、章节覆盖、节首句、方法章具体数值。' +
			'docx 不存在时自动改查 draft.md，写作途中也能用。' +
			'退出码 0=通过 / 2=无法检查 / 3=有 P0 阻断 / 4=只有 P1。P0 未清零不得交付。',
		parameters: toRawParameters({
			project: { type: 'string', required: true },
			input: { type: 'string', description: '待检 docx 文件名，默认 output.docx。' },
			source: { type: 'string', description: '回落检查的源稿，默认 draft.md。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--project', args.project];
			push(argv, '--input', args.input);
			push(argv, '--source', args.source);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('gate.py'), argv, { cwd: args.cwd, signal: exec?.signal });
			const verdict =
				result.exitCode === 0 ? 'PASS' : result.exitCode === 3 ? 'BLOCK' : result.exitCode === 4 ? 'WARN' : 'ERROR';
			return `[${verdict}] 退出码 ${result.exitCode}\n${result.output}`;
		},
	});

	// ── 渲染闸门 ────────────────────────────────────────────────────────

	defs.push({
		name: 'word_render',
		description:
			'渲染闸门：DOCX → PDF → 逐页 PNG。结构全过但排版崩掉很常见（字体回退、表格错位、标题孤行），' +
			'交付前必须真的看过图。缺 LibreOffice 时自动跳过并明确说明，不假装通过。',
		parameters: toRawParameters({
			project: { type: 'string', required: true },
			input: { type: 'string', description: '待渲文件名，默认 output.docx。' },
			dpi: { type: 'number', description: '预览图 DPI，默认 110。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--project', args.project];
			push(argv, '--input', args.input);
			pushNum(argv, '--dpi', args.dpi);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('render.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 300000,
			});
			return result.output;
		},
	});

	// ── 导出 ────────────────────────────────────────────────────────────

	defs.push({
		name: 'word_export',
		description:
			'交付导出：补核心属性（标题/作者/语言）、回读校验可打开性、可选转 PDF。' +
			'导出前必须先过 word_gate 且无 P0。',
		parameters: toRawParameters({
			project: { type: 'string', required: true },
			input: { type: 'string', description: '源文件名，默认 output.docx。' },
			name: { type: 'string', description: '交付文件名，默认 final.docx。' },
			pdf: { type: 'boolean', description: '同时导出 PDF。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const argv = ['--project', args.project];
			push(argv, '--input', args.input);
			push(argv, '--name', args.name);
			pushFlag(argv, '--pdf', args.pdf);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('export.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 300000,
			});
			return result.output;
		},
	});

	// ── 审阅界面 ────────────────────────────────────────────────────────

	defs.push({
		name: 'word_review',
		description:
			'打开标注式 HTML 审阅界面（本地浏览器）：正文可改、逐段选处理方式（chips）、可写自由文本要求、' +
			'质检发现可定位/一键修复、整篇要求可多选。用户的改动写回 draft.md，决定写进 review-notes.json。' +
			'交接判据以 review-result.json 为准。用完必须 action=stop，别留占端口的僵尸进程。',
		parameters: toRawParameters({
			action: {
				type: 'string',
				enum: ['open', 'stop', 'status', 'notes'],
				required: true,
				description: 'open 启动并返回 URL；stop 关闭；status 查状态；notes 读收集到的决定。',
			},
			project: { type: 'string', required: true, description: '项目目录绝对路径。' },
			port: { type: 'number', description: '首选端口，默认 5070；被占用会自动往后找。' },
			timeout: { type: 'number', description: '空闲超时秒数，默认 1800；0 表示不超时。' },
			cwd: COMMON_PARAMS.cwd,
			extra_args: COMMON_PARAMS.extra_args,
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const project = args.project;

			if (args.action === 'notes') {
				const code = [
					'import json, sys',
					'from pathlib import Path',
					'p = Path(sys.argv[1]) / "review-notes.json"',
					'if not p.is_file():',
					'    print(json.dumps({"ok": True, "decisions": [], "globalIntents": [], "globalNote": "", "note": "还没有决定"}, ensure_ascii=False))',
					'else:',
					'    print(p.read_text(encoding="utf-8"))',
				].join('\n');
				const result = await runPythonCode(code, { argv: [project], signal: exec?.signal });
				return result.output;
			}

			if (args.action === 'status') {
				const code = [
					'import json, sys, urllib.request',
					'from pathlib import Path',
					'lock = Path(sys.argv[1]) / ".review-lock.json"',
					'if not lock.is_file():',
					'    print(json.dumps({"ok": True, "running": False}))',
					'    sys.exit(0)',
					'info = json.loads(lock.read_text(encoding="utf-8"))',
					'try:',
					'    with urllib.request.urlopen(f"http://127.0.0.1:{info[\'port\']}/api/health", timeout=2) as r:',
					'        health = json.loads(r.read().decode("utf-8"))',
					'    print(json.dumps({"ok": True, "running": True, "port": info["port"], "pid": info.get("pid"), "url": f"http://127.0.0.1:{info[\'port\']}/"}, ensure_ascii=False))',
					'except Exception as exc:',
					'    print(json.dumps({"ok": True, "running": False, "staleLock": True, "error": str(exc)[:120]}, ensure_ascii=False))',
				].join('\n');
				const result = await runPythonCode(code, { argv: [project], signal: exec?.signal });
				return result.output;
			}

			if (args.action === 'stop') {
				const code = [
					'import json, sys, time, urllib.request',
					'from pathlib import Path',
					'lock = Path(sys.argv[1]) / ".review-lock.json"',
					'if not lock.is_file():',
					'    print(json.dumps({"ok": True, "stopped": False, "note": "没有运行中的服务器"}))',
					'    sys.exit(0)',
					'info = json.loads(lock.read_text(encoding="utf-8"))',
					'port = info.get("port")',
					'try:',
					'    req = urllib.request.Request(f"http://127.0.0.1:{port}/api/shutdown", data=b"{}", headers={"Content-Type": "application/json"}, method="POST")',
					'    urllib.request.urlopen(req, timeout=3)',
					'except Exception:',
					'    pass',
					'time.sleep(0.8)',
					'try:',
					'    lock.unlink()',
					'except OSError:',
					'    pass',
					'print(json.dumps({"ok": True, "stopped": True, "port": port}, ensure_ascii=False))',
				].join('\n');
				const result = await runPythonCode(code, { argv: [project], signal: exec?.signal });
				return result.output;
			}

			// open
			const argv = [project, '--daemon'];
			pushNum(argv, '--port', args.port);
			pushNum(argv, '--timeout', args.timeout);
			pushExtra(argv, args.extra_args);
			const result = await runScript(scriptPath('review_server.py'), argv, {
				cwd: args.cwd,
				signal: exec?.signal,
				timeoutMs: 60000,
			});
			return result.output;
		},
	});

	// ── 读文档 ──────────────────────────────────────────────────────────

	defs.push({
		name: 'word_read',
		description:
			'读取现有 DOCX 的正文结构（标题层级、段落、表格、字符数），用于改稿前先了解现状。' +
			'返回结构化摘要而非全文，长文档也不会撑爆上下文。',
		parameters: toRawParameters({
			path: { type: 'string', required: true, description: 'DOCX 绝对路径。' },
			maxParagraphs: { type: 'number', description: '最多返回多少段，默认 200。' },
		}),
		output: { schema: { type: 'string' }, render: (_a, v) => textOutput(String(v)) },
		async execute(args, exec) {
			const code = [
				'import json, sys',
				'from docx import Document',
				'path = sys.argv[1]',
				'limit = int(sys.argv[2])',
				'doc = Document(path)',
				'rows = []',
				'total = 0',
				'for p in doc.paragraphs:',
				'    text = p.text.strip()',
				'    if not text:',
				'        continue',
				'    total += len(text)',
				'    if len(rows) < limit:',
				'        rows.append({"style": (p.style.name or ""), "text": text[:400]})',
				'print(json.dumps({"paragraphs": len(doc.paragraphs), "tables": len(doc.tables), "charCount": total, "items": rows}, ensure_ascii=False))',
			].join('\n');
			const result = await runPythonCode(code, {
				argv: [args.path, String(Number.isFinite(args.maxParagraphs) ? args.maxParagraphs : 200)],
				signal: exec?.signal,
				timeoutMs: 60000,
			});
			return result.ok ? result.output : `[ERR] 读取失败：${result.output}`;
		},
	});

	return defs;
}

/** 供测试与自检使用。 */
export const __internals = { toRawSchema, toRawParameters, pushExtra };
