/**
 * Word Master 偏好记忆：本地偏好库 + 每轮自动注入。
 *
 * 为什么要自己做一层，而不是直接依赖记忆空间
 * ------------------------------------------
 * 记忆空间是「按需检索」的：模型得**想起来**调用召回才会有内容。写作偏好这
 * 类约束恰恰不能靠自觉 —— 忘了召回一次，那一节就违反规定了。所以这里做两层：
 *
 *   1. **本地偏好库**（`$DSH_HOME/word-master/preferences.json`）是唯一事实
 *      来源，结构明确、可直接编辑、可版本化；
 *   2. **每轮注入**：偏好被渲染成一段系统提示词小节，通过
 *      `ctx.systemPrompt.section()` 挂在每一次请求上 —— 不依赖模型是否记得查。
 *
 * 记忆空间那侧作为**跨机器备份与长期档案**：`word_pref_memory_text` 把偏好库
 * 渲染成一段可直接交给 `mnemon_remember` 的文字，由用户说"记住"时写入。这样
 * 既不重复造记忆系统，也不让关键约束悬在一次召回上。
 *
 * 偏好分三态
 * ----------
 * - `default`：内置默认值，用户还没改过；
 * - `user`：用户明确要求过的（优先级最高，导出时只带这一类 + default 中未被覆盖的）；
 * - 被 `null` 显式关闭的项不注入。
 *
 * @module dsh-word-master/preferences
 */

import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';

/** 偏好库文件名。 */
const FILE_NAME = 'preferences.json';

/** 偏好库 schema 版本；结构变更时用它做迁移判断。 */
const SCHEMA_VERSION = 1;

/**
 * 解析 DSH 主目录。
 *
 * @returns `$DSH_HOME` 或 `~/.dsh`。
 */
function dshHome() {
	const env = process.env.DSH_HOME?.trim();
	return env !== undefined && env !== '' ? env : join(homedir(), '.dsh');
}

/**
 * 偏好库的绝对路径。
 *
 * @returns 路径字符串。
 */
export function preferencesPath() {
	return join(dshHome(), 'word-master', FILE_NAME);
}

/**
 * 内置默认偏好。
 *
 * 这些值不是你机器上的猜测，而是从既有写作规范里固化下来的可执行条目：
 * 标点、禁用表述、段落推进、章节首句、方法章约束、版式。用户可以随时改。
 */
export const DEFAULT_PREFERENCES = {
	punctuation: {
		halfWidthBracketsOnly: true,
		note: '全文一律半角括号 (),禁止全角（）,表头也不例外',
	},
	bannedPhrases: {
		defensive: ['不能', '不宜', '不应', '并非', '无法', '缺乏', '不代表', '不等于'],
		delivery: ['待补', '待核', '待定', '待授权', '台账', '核对状态', '证据缺口', '占位', '工作稿', '候选版'],
		meta: ['需要说明的是', '上述份额分别对应', '以便读者', '本文不再重复'],
		process: ['经核实', '逐条记录', '纳入N条记录'],
	},
	paragraph: {
		oneTopicPerParagraph: true,
		mergeAdjacentShortParagraphs: true,
		progression: '给结果 → 说不确定性 → 给判断',
	},
	sectionOpener: {
		enabled: true,
		pattern: '为了 + 本节要解决的问题,需要 + 相应动作',
		methodChapterExempt: true,
	},
	methodChapter: {
		noConcreteNumbers: true,
		keepOnly: ['公式常数', '术语内数字（HC5/EC50）', '引用年份', '表图号'],
		allowSimulationSettings: true,
		noTablePointers: true,
		keepFormulaExplanations: true,
		noOpenerSentence: true,
	},
	resultsChapter: {
		noForwardReference: true,
		noRepeatFromMethod: true,
		oneNumberPerExperiment: true,
	},
	terminology: {
		plainLanguage: true,
		avoidLiteraryWords: ['时机', '历程', '特征', '相对'],
		significanceNeedsTest: true,
		explainFirstOccurrence: true,
	},
	formatting: {
		bodyFont: '宋体',
		latinFont: 'Times New Roman',
		bodySizePt: 12,
		lineSpacing: 1.5,
		firstLineIndentChars: 2,
		marginCm: { top: 2.54, bottom: 2.54, left: 3.17, right: 3.17 },
		headingStyleBuiltin: true,
	},
	citation: {
		style: 'GB/T 7714',
		numbering: '顺序编码制',
	},
	autoInjection: {
		enabled: true,
		maxChars: 4000,
	},
	notes: [],
};

/** 深拷贝（偏好对象只有 JSON 值）。 */
function clone(value) {
	return JSON.parse(JSON.stringify(value));
}

/** 判断是否为普通对象。 */
function isPlainObject(value) {
	return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * 把补丁深合并进目标。
 *
 * 规则：对象递归合并；数组与标量整体替换；`null` 表示**显式关闭**，保留 null。
 *
 * @param target 目标对象（会被修改）。
 * @param patch 补丁对象。
 * @returns 合并后的目标。
 */
function deepMerge(target, patch) {
	for (const [key, value] of Object.entries(patch)) {
		if (isPlainObject(value) && isPlainObject(target[key])) {
			deepMerge(target[key], value);
		} else {
			target[key] = value;
		}
	}
	return target;
}

/**
 * 读取偏好库；文件不存在时返回内置默认值。
 *
 * 读取失败（JSON 损坏）时**不抛异常**：返回默认值并在 `_warning` 里说明，
 * 避免一个坏文件把整个预设装载搞崩。
 *
 * @returns `{ preferences, source, warning? }`。
 */
export function loadPreferences() {
	const path = preferencesPath();
	const preferences = clone(DEFAULT_PREFERENCES);
	if (!existsSync(path)) {
		return { preferences, source: 'default', path };
	}
	try {
		const raw = readFileSync(path, 'utf8').replace(/^\uFEFF/, '');
		const parsed = JSON.parse(raw);
		const merged = deepMerge(preferences, isPlainObject(parsed.values) ? parsed.values : parsed);
		return { preferences: merged, source: 'file', path };
	} catch (error) {
		return {
			preferences,
			source: 'default',
			path,
			warning: `偏好库读取失败，已回退到默认值：${String(error?.message ?? error)}`,
		};
	}
}

/**
 * 写入偏好库（原子替换：先写临时文件再 rename）。
 *
 * @param patch 要合并进去的补丁。
 * @returns `{ ok, path, preferences, error? }`。
 */
export function savePreferences(patch) {
	const path = preferencesPath();
	try {
		const current = loadPreferences();
		const values = deepMerge(current.preferences, patch ?? {});
		mkdirSync(dirname(path), { recursive: true });
		const payload = {
			schemaVersion: SCHEMA_VERSION,
			updatedAt: new Date().toISOString(),
			values,
		};
		const tmp = `${path}.tmp`;
		writeFileSync(tmp, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
		renameSync(tmp, path);
		return { ok: true, path, preferences: values };
	} catch (error) {
		return { ok: false, path, error: String(error?.message ?? error) };
	}
}

/**
 * 把偏好渲染成系统提示词小节。
 *
 * 只列出**可执行**的约束，不写解释性长文：注入的每一行都该是模型能直接照做的。
 *
 * @param preferences 偏好对象。
 * @param options 选项。
 * @param options.maxChars 字符上限；超出时按段落截断并标注。
 * @returns 提示词文本；`autoInjection.enabled === false` 时返回空串。
 */
export function renderPreferencesSection(preferences, options = {}) {
	const maxChars = Number.isInteger(options.maxChars) ? options.maxChars : preferences?.autoInjection?.maxChars ?? 4000;
	if (preferences?.autoInjection?.enabled === false) return '';

	const lines = ['## 用户写作偏好（已生效，优先级高于通用写作习惯）', ''];

	const punct = preferences.punctuation ?? {};
	if (punct.halfWidthBracketsOnly === true) {
		lines.push('- 标点：全文一律半角括号 `()`，禁止全角 `（）`，表头也不例外的规则同样适用。');
	}

	const banned = preferences.bannedPhrases ?? {};
	const bannedLabels = [
		['defensive', '防御性表述'],
		['delivery', '交付性用语'],
		['meta', '元表述'],
		['process', '过程性叙述'],
	];
	for (const [key, label] of bannedLabels) {
		const words = Array.isArray(banned[key]) ? banned[key].filter((w) => typeof w === 'string' && w !== '') : [];
		if (words.length > 0) lines.push(`- 禁止${label}：${words.join('、')}`);
	}

	const paragraph = preferences.paragraph ?? {};
	if (paragraph.oneTopicPerParagraph === true) {
		lines.push('- 段落：一段承担一个完整论题，同论题的相邻短段合并。');
	}
	if (typeof paragraph.progression === 'string' && paragraph.progression !== '') {
		lines.push(`- 段内推进：${paragraph.progression}。`);
	}

	const opener = preferences.sectionOpener ?? {};
	if (opener.enabled === true && typeof opener.pattern === 'string') {
		lines.push(
			`- 每节首句统一句式：${opener.pattern}；说明这一节为何存在，不复述方法${opener.methodChapterExempt === true ? '；方法章节除外' : ''}。`,
		);
	}

	const method = preferences.methodChapter ?? {};
	if (method.noConcreteNumbers === true) {
		const keep = Array.isArray(method.keepOnly) ? method.keepOnly.join('、') : '';
		lines.push(`- 方法章节只说明方法，不出现具体数值${keep === '' ? '' : `；仅保留${keep}`}。`);
	}
	if (method.noTablePointers === true) lines.push('- 方法章节不指向具体表格，用"按情景登记""单列"这类方法表述。');
	if (method.noOpenerSentence === true) lines.push('- 方法章节不加"为了……"总起句。');

	const results = preferences.resultsChapter ?? {};
	if (results.noForwardReference === true) lines.push('- 结果章节不提前引用后文才推导的阈值。');
	if (results.noRepeatFromMethod === true) lines.push('- 方法章节已讲过的定义、公式、参数表，结果章节删除，只留指向。');
	if (results.oneNumberPerExperiment === true) lines.push('- 同一实验的数值只在一处讲，另一处用"X节所述的"建立指代。');

	const term = preferences.terminology ?? {};
	if (term.plainLanguage === true) lines.push('- 术语用白话，首次出现给一句解释，之后直接用。');
	const literary = Array.isArray(term.avoidLiteraryWords) ? term.avoidLiteraryWords.filter((w) => typeof w === 'string') : [];
	if (literary.length > 0) lines.push(`- 避免文学化模糊词：${literary.join('、')}。`);
	if (term.significanceNeedsTest === true) lines.push('- 用"显著"必须有统计检验支撑，否则删掉。');

	const fmt = preferences.formatting ?? {};
	const fmtBits = [];
	if (typeof fmt.bodyFont === 'string') fmtBits.push(`中文${fmt.bodyFont}`);
	if (typeof fmt.latinFont === 'string') fmtBits.push(`西文${fmt.latinFont}`);
	if (Number.isFinite(fmt.bodySizePt)) fmtBits.push(`正文 ${fmt.bodySizePt}pt`);
	if (Number.isFinite(fmt.lineSpacing)) fmtBits.push(`${fmt.lineSpacing} 倍行距`);
	if (Number.isFinite(fmt.firstLineIndentChars)) fmtBits.push(`首行缩进 ${fmt.firstLineIndentChars} 字符`);
	if (fmtBits.length > 0) lines.push(`- 版式：${fmtBits.join('，')}。`);

	const citation = preferences.citation ?? {};
	if (typeof citation.style === 'string' && citation.style !== '') {
		lines.push(`- 引文格式：${citation.style}${typeof citation.numbering === 'string' ? `，${citation.numbering}` : ''}。`);
	}

	const notes = Array.isArray(preferences.notes) ? preferences.notes.filter((n) => typeof n === 'string' && n.trim() !== '') : [];
	for (const note of notes) lines.push(`- ${note}`);

	if (lines.length <= 2) return '';

	let text = `${lines.join('\n')}\n\n写作前先调 word_spec 冻结规格，交付前必须过 word_gate。`;
	if (text.length > maxChars) {
		text = `${text.slice(0, maxChars)}\n（偏好已截断，完整内容见 ${preferencesPath()}）`;
	}
	return text;
}

/**
 * 把偏好渲染成可直接写入记忆空间的一段文字。
 *
 * 只输出被用户显式改过或语义关键的条目，避免把默认值当作"用户要求"存进长期记忆。
 *
 * @param preferences 偏好对象。
 * @returns 文本。
 */
export function renderMemoryText(preferences) {
	const lines = ['用户写作偏好（Word Master，长期有效）：', ''];
	const section = renderPreferencesSection(preferences, { maxChars: 12000 });
	const body = section.replace(/^## .*\n\n/, '').replace(/\n\n写作前先调.*$/s, '');
	lines.push(body.trim());
	return lines.join('\n');
}

/** 供测试与自检使用。 */
export const __internals = { deepMerge, dshHome, SCHEMA_VERSION };
