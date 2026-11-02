/**
 * Word Master 运行时：解释器探测、脚本执行、内联 Python。
 *
 * 三条设计原则
 * ------------
 * 1. **不经 shell**。所有命令走 `spawn(bin, argv)` 数组形式，路径里有空格、
 *    中文、引号都不需要转义，也不会被 shell 元字符破坏。
 *
 * 2. **解释器必须真的能跑**。Windows 上 `python` / `python3` 经常是微软商店的
 *    应用执行别名（AppInstaller 重定向器），它**不执行脚本，直接返回 49**。
 *    只看 `which` 命中就采用会得到「脚本没跑但退出码非零」的假故障，因此这里
 *    对每个候选做一次真实探测：`-c "import sys;print(sys.executable)"`，
 *    能打印出路径才算可用。
 *
 * 3. **不写字节码**。脚本住在已安装的包目录里，Python 默认会在旁边生成
 *    `__pycache__/`，把用户的 node_modules 弄脏。子进程统一注入
 *    `PYTHONDONTWRITEBYTECODE=1`，同时给 `PYTHONIOENCODING=utf-8` 保证中文
 *    在 GBK 控制台下不乱码。
 *
 * @module dsh-word-master/runner
 */

import { spawn } from 'node:child_process';
import { existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';

/** 探测超时：解释器起不来通常是秒级失败，不需要长等。 */
const PROBE_TIMEOUT_MS = 8000;

/** 解释器探测结果缓存（成功才缓存，失败允许用户改环境后重试）。 */
let cachedInterpreter;

/** 子进程环境：禁字节码 + UTF-8 输出 + 关掉 Python 的写入缓冲。 */
function childEnv(extra) {
	return {
		...process.env,
		PYTHONDONTWRITEBYTECODE: '1',
		PYTHONIOENCODING: 'utf-8',
		PYTHONUNBUFFERED: '1',
		...extra,
	};
}

/**
 * 跑一个外部命令。
 *
 * @param bin 可执行文件（PATH 名或绝对路径）。
 * @param args 参数数组。
 * @param options 选项。
 * @param options.cwd 工作目录。
 * @param options.timeoutMs 超时；到点 SIGKILL。
 * @param options.signal AbortSignal，取消时立刻杀掉子进程。
 * @param options.env 追加环境变量。
 * @returns `{ ok, exitCode, output, timedOut }`。
 */
export function runCommand(bin, args, options = {}) {
	const { cwd, timeoutMs = 120000, signal, env } = options;
	return new Promise((resolve) => {
		let child;
		try {
			child = spawn(bin, args, {
				cwd,
				env: childEnv(env),
				stdio: ['ignore', 'pipe', 'pipe'],
				windowsHide: true,
			});
		} catch (error) {
			resolve({ ok: false, exitCode: -1, output: `无法启动 ${bin}：${String(error?.message ?? error)}`, timedOut: false });
			return;
		}

		let out = '';
		let err = '';
		let settled = false;
		let timedOut = false;

		const finish = (code) => {
			if (settled) return;
			settled = true;
			clearTimeout(timer);
			if (signal !== undefined) signal.removeEventListener('abort', onAbort);
			const text = [out.trim(), err.trim()].filter((s) => s !== '').join('\n');
			resolve({ ok: code === 0 && !timedOut, exitCode: code ?? -1, output: text, timedOut });
		};

		const timer = setTimeout(() => {
			timedOut = true;
			try {
				child.kill('SIGKILL');
			} catch {
				/* 已经退出 */
			}
		}, timeoutMs);

		const onAbort = () => {
			try {
				child.kill('SIGKILL');
			} catch {
				/* 已经退出 */
			}
		};
		if (signal !== undefined) {
			if (signal.aborted) {
				onAbort();
			} else {
				signal.addEventListener('abort', onAbort, { once: true });
			}
		}

		child.stdout?.on('data', (chunk) => {
			out += chunk.toString('utf8');
		});
		child.stderr?.on('data', (chunk) => {
			err += chunk.toString('utf8');
		});
		child.on('error', (error) => {
			err += String(error?.message ?? error);
			finish(-1);
		});
		child.on('close', (code) => finish(code));
	});
}

/**
 * 探测一个解释器候选是否真的能执行代码。
 *
 * @param name 可执行名（`python` / `python3` / `py` 或绝对路径）。
 * @returns 可执行文件的绝对路径，或 undefined。
 */
async function probeInterpreter(name) {
	const result = await runCommand(name, ['-c', 'import sys;print(sys.executable)'], { timeoutMs: PROBE_TIMEOUT_MS });
	if (!result.ok) return undefined;
	const line = result.output.split(/\r?\n/).map((s) => s.trim()).find((s) => s !== '');
	if (line === undefined || !existsSync(line)) return undefined;
	return line;
}

/**
 * 解析出可用的 Python 解释器。
 *
 * 优先级：`DSH_WORD_PYTHON` 显式指定 → PATH 上的候选（Windows 先 `python`
 * 再 `python3`，因为商店别名通常占着 `python3`）→ 常见安装位置。
 *
 * @returns `{ path }` 或 `{ error, tried }`。
 */
export async function resolveInterpreter() {
	if (cachedInterpreter !== undefined) return cachedInterpreter;

	const override = process.env.DSH_WORD_PYTHON?.trim();
	const candidates = [];
	if (override !== undefined && override !== '') candidates.push(override);
	if (process.platform === 'win32') {
		candidates.push('python', 'python3', 'py');
	} else {
		candidates.push('python3', 'python');
	}

	const tried = [];
	for (const name of candidates) {
		const found = await probeInterpreter(name);
		if (found !== undefined) {
			cachedInterpreter = { path: found };
			return cachedInterpreter;
		}
		tried.push(name);
	}

	// 最后再试几个固定位置（PATH 被裁掉的部署）。
	const fixed = [];
	const home = homedir();
	if (process.platform === 'win32') {
		fixed.push(
			join(home, 'AppData', 'Local', 'Programs', 'Python', 'Python313', 'python.exe'),
			join(home, 'AppData', 'Local', 'Programs', 'Python', 'Python312', 'python.exe'),
			join(home, 'AppData', 'Local', 'Programs', 'Python', 'Python311', 'python.exe'),
		);
	} else {
		fixed.push('/usr/bin/python3', '/usr/local/bin/python3', '/opt/homebrew/bin/python3');
	}
	for (const candidate of fixed) {
		if (!existsSync(candidate)) continue;
		const found = await probeInterpreter(candidate);
		if (found !== undefined) {
			cachedInterpreter = { path: found };
			return cachedInterpreter;
		}
		tried.push(candidate);
	}

	return { error: 'no-runnable-python', tried };
}

/** 清掉探测缓存（用户装了 Python 或改了环境变量后重试用）。 */
export function resetInterpreterCache() {
	cachedInterpreter = undefined;
}

/**
 * 用解析出的解释器跑一个脚本。
 *
 * @param scriptPath 脚本绝对路径。
 * @param args 传给脚本的参数数组。
 * @param options 选项（cwd / timeoutMs / signal / env）。
 * @returns `{ ok, exitCode, output, timedOut }`。
 */
export async function runScript(scriptPath, args = [], options = {}) {
	const interpreter = await resolveInterpreter();
	if (interpreter.error !== undefined) {
		return {
			ok: false,
			exitCode: 127,
			output:
				`[ERR] 找不到可用的 Python 解释器（试过：${interpreter.tried.join(', ')}）。\n` +
				'请安装 Python 3.10+，或用环境变量 DSH_WORD_PYTHON 指定解释器的绝对路径。\n' +
				'注意：Windows 上 WindowsApps 目录里的 python/python3 是商店占位别名，不能用来跑脚本。',
			timedOut: false,
		};
	}
	return runCommand(interpreter.path, [scriptPath, ...args], options);
}

/**
 * 用 `python -c` 跑一段内联代码，取回 stdout。
 *
 * 用途：把多次脚本调用压进一个进程（如批量检查多份文档）。
 *
 * @param code Python 源码。
 * @param options 选项；`argv` 追加为 `sys.argv[1:]`。
 * @returns `{ ok, exitCode, output, timedOut }`。
 */
export async function runPythonCode(code, options = {}) {
	const interpreter = await resolveInterpreter();
	if (interpreter.error !== undefined) {
		return { ok: false, exitCode: 127, output: '[ERR] 找不到可用的 Python 解释器。', timedOut: false };
	}
	const argv = Array.isArray(options.argv) ? options.argv : [];
	return runCommand(interpreter.path, ['-c', code, ...argv], options);
}

/** 环境体检用的依赖清单（模块名 → 用途 → 安装命令）。 */
export const DEPENDENCIES = [
	{ module: 'docx', label: 'Word 读写（python-docx）', install: 'python-docx', required: true },
	{ module: 'lxml', label: 'OOXML 结构校验（lxml）', install: 'lxml', required: true },
];

/** 可选外部工具：LibreOffice 用于渲染闸门，缺失时降级为纯结构检查。 */
export const EXTERNAL_TOOLS = [
	{
		id: 'soffice',
		label: 'LibreOffice（docx → PDF → PNG 渲染闸门）',
		candidates:
			process.platform === 'win32'
				? [
						'C:/Program Files/LibreOffice/program/soffice.exe',
						'C:/Program Files (x86)/LibreOffice/program/soffice.exe',
					]
				: ['/usr/bin/soffice', '/usr/local/bin/soffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice'],
		required: false,
	},
];

/**
 * 体检：解释器 + Python 依赖 + 外部工具。
 *
 * @returns 结构化报告。
 */
export async function environmentReport() {
	const interpreter = await resolveInterpreter();
	const report = {
		interpreter: interpreter.error === undefined ? { ok: true, path: interpreter.path } : { ok: false, tried: interpreter.tried },
		pythonDependencies: [],
		externalTools: [],
		ok: false,
	};

	if (interpreter.error === undefined) {
		const code = [
			'import json, importlib.util',
			'names = json.loads(__import_json__)',
			'out = {}',
			'for name in names:',
			'    out[name] = importlib.util.find_spec(name) is not None',
			'print(json.dumps(out))',
		].join('\n');
		const probe = await runPythonCode(code, {
			argv: [JSON.stringify(DEPENDENCIES.map((d) => d.module))],
			timeoutMs: 30000,
			env: { __import_json__: JSON.stringify(DEPENDENCIES.map((d) => d.module)) },
		});
		let found = {};
		if (probe.ok) {
			const line = probe.output.split(/\r?\n/).map((s) => s.trim()).filter((s) => s.startsWith('{')).pop();
			try {
				found = line === undefined ? {} : JSON.parse(line);
			} catch {
				found = {};
			}
		}
		for (const dep of DEPENDENCIES) {
			report.pythonDependencies.push({
				module: dep.module,
				label: dep.label,
				install: dep.install,
				required: dep.required,
				present: found[dep.module] === true,
			});
		}
	}

	for (const tool of EXTERNAL_TOOLS) {
		const hit = tool.candidates.find((candidate) => existsSync(candidate));
		report.externalTools.push({
			id: tool.id,
			label: tool.label,
			required: tool.required,
			present: hit !== undefined,
			path: hit,
		});
	}

	const pythonOk = interpreter.error === undefined;
	const depsOk = report.pythonDependencies.every((d) => d.present || d.required === false);
	report.ok = pythonOk && depsOk;
	return report;
}

/**
 * 生成给模型看的体检摘要。
 *
 * @param report `environmentReport()` 的返回值。
 * @returns 纯文本。
 */
export function formatEnvironmentReport(report) {
	const lines = ['# Word Master 环境体检', ''];

	if (report.interpreter?.ok === true) {
		lines.push(`- 解释器：${report.interpreter.path}`);
	} else {
		lines.push(`- 解释器：**未找到**（试过 ${(report.interpreter?.tried ?? []).join(', ')}）`);
		lines.push('  - 装 Python 3.10+，或设置环境变量 DSH_WORD_PYTHON 指向 python.exe。');
		lines.push('  - Windows 的 WindowsApps\\python.exe 是商店占位别名，不能用于执行脚本（跑了直接返回 49）。');
	}

	const missing = report.pythonDependencies.filter((d) => !d.present);
	for (const dep of report.pythonDependencies) {
		lines.push(
			`- ${dep.label}：${dep.present ? '已装' : '**缺失**'}${dep.present ? '' : `（安装：pip install ${dep.install}）`}`,
		);
	}

	for (const tool of report.externalTools) {
		lines.push(
			`- ${tool.label}：${tool.present ? `已装（${tool.path}）` : '未装（渲染闸门将降级为结构检查）'}`,
		);
	}

	if (missing.length > 0) {
		// 给出可直接复制的完整命令，而不是让模型自己拼
		const pkgs = [...new Set(missing.map((d) => d.install))].join(' ');
		lines.push('');
		lines.push('## 补齐命令');
		lines.push('```');
		lines.push(`pip install ${pkgs}`);
		lines.push('```');
		if (report.interpreter?.ok === true) {
			lines.push(`（如果 PATH 上有多个 Python，用完整路径：\n\`"${report.interpreter.path}" -m pip install ${pkgs}\`）`);
		}
	}

	lines.push('');
	lines.push(report.ok ? '结论：**可以开工**。' : '结论：**先补齐缺失项**，否则生成或质检会失败。');
	return lines.join('\n');
}
