#!/usr/bin/env node
/**
 * dsh-word-master 自检。
 *
 * 覆盖四类检查：
 *   1. DSH 插件契约 —— 复刻 `installBundle` 的硬校验（不满足会装不上）；
 *   2. skill 载荷 —— frontmatter 可解析、`${SKILL_DIR}` 全部可展开；
 *   3. Python 脚本 —— 语法可编译、依赖可导入、`--help` 能起来；
 *   4. 发布卫生 —— 无本机专属路径、无凭据、无垃圾文件。
 *
 * 无外部依赖：`node verify.mjs`。
 */

import { execFile } from 'node:child_process';
import { readFile, readdir, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join, relative } from 'node:path';
import { promisify } from 'node:util';

const run = promisify(execFile);
const PKG_DIR = dirname(fileURLToPath(import.meta.url));
const results = [];

/**
 * 记录一条检查结果。
 * @param {'PASS'|'FAIL'|'WARN'} status 状态。
 * @param {string} label 检查项名称。
 * @param {string} detail 说明。
 */
function check(status, label, detail) {
	results.push([status, label, detail]);
}

/** 读取并解析包根 package.json。 */
async function readManifest() {
	const raw = await readFile(join(PKG_DIR, 'package.json'), 'utf8');
	return JSON.parse(raw.replace(/^\uFEFF/, ''));
}

/** 递归列出文件（跳过 node_modules 与 .git）。 */
async function listFiles(root, base = root) {
	const out = [];
	let entries;
	try {
		entries = await readdir(root, { withFileTypes: true });
	} catch {
		return out;
	}
	for (const entry of entries) {
		if (entry.name === 'node_modules' || entry.name === '.git' || entry.name === '__pycache__') continue;
		const full = join(root, entry.name);
		if (entry.isDirectory()) {
			out.push(...(await listFiles(full, base)));
		} else if (entry.isFile()) {
			out.push(relative(base, full).replaceAll('\\', '/'));
		}
	}
	return out;
}

// ── 1. 插件契约 ────────────────────────────────────────────────────────────

const manifest = await readManifest();
check(manifest.name === 'dsh-word-master' ? 'PASS' : 'FAIL', '包名', manifest.name);
check(manifest.type === 'module' ? 'PASS' : 'FAIL', 'ESM 模块', `type=${manifest.type}`);
check(manifest.main === 'lib/index.js' ? 'PASS' : 'FAIL', '入口', String(manifest.main));
check(
	manifest.dsh?.bundle?.patch === './cordis.patch.yml' ? 'PASS' : 'FAIL',
	'dsh.bundle 声明',
	JSON.stringify(manifest.dsh?.bundle ?? null),
);

const filesField = Array.isArray(manifest.files) ? manifest.files : [];
for (const required of ['lib', 'skills', 'cordis.patch.yml']) {
	check(filesField.includes(required) ? 'PASS' : 'FAIL', `files 含 ${required}`, filesField.join(', '));
}

// 包内不能声明的字段：preset 是 bundle 的事，不是 plugin 的事
check(
	manifest.dsh?.profile === undefined ? 'PASS' : 'FAIL',
	'不同时声明 profile',
	'组合包与 profile 不能是同一个包',
);

// ── 2. 入口导出的契约 ──────────────────────────────────────────────────────

const mod = await import(new URL('./lib/index.js', import.meta.url));
check(mod.name === 'dsh-word-master' ? 'PASS' : 'FAIL', '插件 name', mod.name);
check(typeof mod.apply === 'function' ? 'PASS' : 'FAIL', '导出 apply', typeof mod.apply);
check(Array.isArray(mod.inject) ? 'PASS' : 'FAIL', '导出 inject', JSON.stringify(mod.inject));
check(
	Array.isArray(mod.inject) && mod.inject.includes('tools') && mod.inject.includes('skills')
		? 'PASS'
		: 'FAIL',
	'inject 含 tools/skills',
	JSON.stringify(mod.inject),
);
check(
	Array.isArray(mod.inject) && !mod.inject.includes('systemPrompt') ? 'PASS' : 'WARN',
	'systemPrompt 是可选依赖',
	'写进 inject 会让无 systemPrompt 的部署一直 PENDING',
);

// 工具集能构造出来，且每个工具都有合法 schema
const skillDir = mod.__internals.resolveSkillDir();
let definitions = [];
try {
	definitions = mod.__internals.createToolDefinitions({ skillDir });
	check('PASS', '工具集构造', `${definitions.length} 个工具`);
} catch (error) {
	check('FAIL', '工具集构造', String(error?.message ?? error));
}

const toolNames = new Set();
for (const def of definitions) {
	const bad = [];
	if (typeof def.name !== 'string' || !/^[a-z][a-z0-9_]*$/.test(def.name)) bad.push('name');
	if (typeof def.description !== 'string' || def.description.length < 20) bad.push('description');
	if (def.parameters?.type !== 'object') bad.push('parameters');
	if (def.arguments !== undefined && def.parameters?.required !== undefined) bad.push('legacy-required');
	if (!def.output || typeof def.output.render !== 'function') bad.push('output.render');
	if (toolNames.has(def.name)) bad.push('duplicate');
	toolNames.add(def.name);
	check(bad.length === 0 ? 'PASS' : 'FAIL', `工具 ${def.name}`, bad.length === 0 ? 'schema 合法' : bad.join(', '));
}

// ── 3. skill 载荷 ──────────────────────────────────────────────────────────

const SKILL_DIRS = ['word-master', 'word-master-dsh'];
for (const dir of SKILL_DIRS) {
	let row;
	try {
		row = await mod.__internals.loadSkill(dir);
		check('PASS', `skill ${dir} 可加载`, `${row.content.length} 字符`);
	} catch (error) {
		check('FAIL', `skill ${dir} 可加载`, String(error?.message ?? error));
		continue;
	}
	check(row.name === dir ? 'PASS' : 'WARN', `skill ${dir} 名称`, row.name);
	check(row.description !== '' ? 'PASS' : 'FAIL', `skill ${dir} 描述`, row.description.slice(0, 40));
	check(
		!row.content.includes('${SKILL_DIR}') ? 'PASS' : 'FAIL',
		`skill ${dir} 占位符展开`,
		'${SKILL_DIR} 应全部替换为绝对路径',
	);
}

// 上游载荷引用的文件都得真实存在
const flowPath = join(PKG_DIR, 'skills', 'word-master', 'references', 'flow.md');
try {
	await stat(flowPath);
	check('PASS', 'references/flow.md 存在', 'SKILL.md 引用了它');
} catch {
	check('FAIL', 'references/flow.md 存在', 'SKILL.md 引用但文件缺失');
}

// ── 4. Python 脚本 ─────────────────────────────────────────────────────────

const SCRIPTS = [
	'spec.py',
	'build_docx.py',
	'gate.py',
	'render.py',
	'export.py',
	'prefs.py',
	'review_server.py',
	'convert.py',
	'template.py',
	'visual_check.py',
	'project.py',
	'wm_common.py',
	'wm_checks.py',
	'draft_model.py',
	'omml.py',
];
const pyDir = join(PKG_DIR, 'skills', 'word-master', 'scripts');

// 审阅界面的静态资源必须齐全，缺一个页面就白屏
for (const asset of ['index.html', 'style.css', 'app.js', 'catalogs.json']) {
	try {
		const info = await stat(join(pyDir, 'review_ui', asset));
		check(info.size > 200 ? 'PASS' : 'FAIL', `审阅界面 ${asset}`, `${info.size} 字节`);
	} catch {
		check('FAIL', `审阅界面 ${asset}`, '缺失');
	}
}

// catalogs 结构：verdicts / global_intents 都得是非空数组且 id 唯一
try {
	const catalogsRaw = await readFile(join(pyDir, 'review_ui', 'catalogs.json'), 'utf8');
	const catalogs = JSON.parse(catalogsRaw);
	for (const key of ['verdicts', 'global_intents']) {
		const list = catalogs[key];
		const ok =
			Array.isArray(list) && list.length >= 6 && list.every((o) => o && typeof o.id === 'string' && o.label_zh);
		check(ok ? 'PASS' : 'FAIL', `catalogs.${key}`, Array.isArray(list) ? `${list.length} 项` : '不是数组');
		const ids = Array.isArray(list) ? list.map((o) => o.id) : [];
		check(new Set(ids).size === ids.length ? 'PASS' : 'FAIL', `catalogs.${key} id 唯一`, `${ids.length} 个`);
	}
} catch (error) {
	check('FAIL', 'catalogs 可解析', String(error?.message ?? error).slice(0, 120));
}

let python;
for (const candidate of ['python', 'python3', 'py']) {
	try {
		const { stdout } = await run(candidate, ['-c', 'import sys;print(sys.executable)'], { timeout: 15000 });
		if (stdout.trim() !== '') {
			python = candidate;
			check('PASS', 'Python 解释器', stdout.trim());
			break;
		}
	} catch {
		/* 试下一个 */
	}
}
if (python === undefined) check('FAIL', 'Python 解释器', '找不到可执行的 python');

if (python !== undefined) {
	// 语法编译
	try {
		await run(python, ['-m', 'py_compile', ...SCRIPTS.map((s) => join(pyDir, s))], { timeout: 60000 });
		check('PASS', 'Python 语法', `${SCRIPTS.length} 个脚本全部可编译`);
	} catch (error) {
		check('FAIL', 'Python 语法', String(error?.stdout ?? error?.message ?? error).slice(0, 200));
	}

	// 依赖导入
	for (const moduleName of ['docx', 'lxml']) {
		try {
			await run(python, ['-c', `import ${moduleName}`], { timeout: 30000 });
			check('PASS', `依赖 ${moduleName}`, '可导入');
		} catch {
			check('FAIL', `依赖 ${moduleName}`, `pip install ${moduleName === 'docx' ? 'python-docx' : moduleName}`);
		}
	}

	// 素材转换的可选依赖：缺了只是不支持那种格式，不影响主流程
	for (const [moduleName, pkg, label] of [
		['fitz', 'pymupdf', 'PDF'],
		['mammoth', 'mammoth', 'Word'],
		['openpyxl', 'openpyxl', 'Excel'],
		['pptx', 'python-pptx', 'PPT'],
		['bs4', 'beautifulsoup4', '网页'],
	]) {
		try {
			await run(python, ['-c', `import ${moduleName}`], { timeout: 30000 });
			check('PASS', `转换 ${label} 依赖`, moduleName);
		} catch {
			check('WARN', `转换 ${label} 依赖`, `缺 ${pkg}（该格式不可用，pip install ${pkg}）`);
		}
	}

	// 四个主要脚本 --help 能起来（证明 argparse 结构没问题）
	for (const script of ['spec.py', 'build_docx.py', 'gate.py', 'export.py']) {
		try {
			const { stdout } = await run(python, [join(pyDir, script), '--help'], { timeout: 30000 });
			check(stdout.includes('usage') ? 'PASS' : 'FAIL', `${script} --help`, 'argparse 正常');
		} catch (error) {
			check('FAIL', `${script} --help`, String(error?.stderr ?? error?.message ?? error).slice(0, 160));
		}
	}

	// gate 的退出码语义：无文件时应为 2（不是崩）
	try {
		await run(python, [join(pyDir, 'gate.py'), '--project', join(PKG_DIR, 'nonexistent-project')], {
			timeout: 30000,
		});
		check('WARN', 'gate 缺文件行为', '本应非零退出，实际为 0');
	} catch (error) {
		const code = error?.code;
		check(code === 2 ? 'PASS' : 'WARN', 'gate 缺文件行为', `退出码 ${code}（期望 2）`);
	}

	// 内置默认偏好必须与 JS 侧一致（否则"注入的"和"查的"是两套规则）
	// 注意：必须显式让 Python 以 UTF-8 输出，否则 Windows 管道走 GBK，中文会读成乱码。
	try {
		const probe = [
			'import json,sys',
			'sys.path.insert(0, sys.argv[1])',
			'import wm_common',
			'print(json.dumps(sorted(wm_common.DEFAULT_PREFERENCES.get("bannedPhrases",{}).get("defensive",[])), ensure_ascii=True))',
		].join('\n');
		const { stdout } = await run(python, ['-c', probe, pyDir], {
			timeout: 30000,
			env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
		});
		const pyList = JSON.parse(stdout.trim());
		const jsList = [...(mod.diagnostics.loadPreferences().preferences?.bannedPhrases?.defensive ?? [])].sort();
		const same = JSON.stringify([...pyList].sort()) === JSON.stringify(jsList);
		check(same ? 'PASS' : 'FAIL', '默认偏好两侧一致', same ? '禁用词表相同' : `js=${jsList} py=${pyList}`);
	} catch (error) {
		check('WARN', '默认偏好两侧一致', String(error?.message ?? error).slice(0, 160));
	}
}

// ── 5. 发布卫生 ────────────────────────────────────────────────────────────

const allFiles = await listFiles(PKG_DIR);
check(allFiles.length > 0 ? 'PASS' : 'FAIL', '文件清单', `${allFiles.length} 个文件`);

const junk = allFiles.filter((f) => /(__pycache__|\.pyc$|\.tmp$|\.log$|node_modules)/.test(f));
check(junk.length === 0 ? 'PASS' : 'FAIL', '无垃圾文件', junk.length === 0 ? '干净' : junk.join(', '));

const TEXT_EXT = /\.(js|mjs|json|yml|yaml|md|py|txt)$/;
const textFiles = allFiles.filter((f) => TEXT_EXT.test(f));
const localPaths = [];
const secrets = [];
for (const rel of textFiles) {
	let text;
	try {
		text = await readFile(join(PKG_DIR, rel), 'utf8');
	} catch {
		continue;
	}
	// 本机专属路径（文档里的示例路径允许 C:\ 但必须是通用形态）
	for (const match of text.matchAll(/[A-Z]:[\\/](?:Users|home)[\\/][^\\/"'\s]+/g)) {
		localPaths.push(`${rel}: ${match[0]}`);
	}
	for (const match of text.matchAll(/(gh[pousr]_[A-Za-z0-9_]{20,}|github_pat_[A-Za-z0-9_]{20,}|sk-[A-Za-z0-9]{20,})/g)) {
		secrets.push(`${rel}: ${match[1].slice(0, 8)}…`);
	}
}
check(localPaths.length === 0 ? 'PASS' : 'FAIL', '无本机专属路径', localPaths.length === 0 ? `扫了 ${textFiles.length} 个文本文件` : localPaths.slice(0, 3).join('; '));
check(secrets.length === 0 ? 'PASS' : 'FAIL', '无凭据泄漏', secrets.length === 0 ? '干净' : secrets.join('; '));

// patch 文件必须能被 YAML 结构识别（顶层是 insert 列表）
let patchText = '';
try {
	patchText = await readFile(join(PKG_DIR, 'cordis.patch.yml'), 'utf8');
	check(patchText.includes('- insert:') ? 'PASS' : 'FAIL', 'patch 结构', '顶层 insert 列表');
	check(
		patchText.includes("@deepseek-ai/dsh-agent-preset") ? 'PASS' : 'FAIL',
		'patch 声明 preset',
		'用官方 preset 声明行',
	);
	check(
		patchText.includes('name: dsh-word-master') ? 'PASS' : 'FAIL',
		'patch 引用本包',
		'预设内组装行按包名引用',
	);
	check(patchText.includes('\t') ? 'FAIL' : 'PASS', 'patch 无制表符', 'YAML 只用空格缩进');
} catch {
	check('FAIL', 'patch 结构', 'cordis.patch.yml 缺失');
}

// ── 输出 ───────────────────────────────────────────────────────────────────

const width = Math.max(...results.map(([, label]) => label.length));
for (const [status, label, detail] of results) {
	console.log(`${status}  ${label.padEnd(width)}  ${detail}`);
}
const failed = results.filter(([s]) => s === 'FAIL').length;
const warned = results.filter(([s]) => s === 'WARN').length;
console.log(`\n${results.length - failed}/${results.length} 通过${warned > 0 ? `（${warned} 条 WARN）` : ''}`);
process.exit(failed === 0 ? 0 : 1);
