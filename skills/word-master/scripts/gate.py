#!/usr/bin/env python3
"""Word Master — 质检闸门（gate）。

这是流水线里唯一有**否决权**的一步。判据全部来自 `wm_checks`（与审阅界面共用
同一套规则，避免两套实现漂移），本文件只负责：读哪份文件、汇总、写报告、定
退出码。

退出码
------
- `0`：全部通过
- `2`：无法检查（文件缺失/损坏）
- `3`：存在 P0（阻断交付）
- `4`：只有 P1（建议改，不阻断）

用法
----
    python gate.py --project DIR [--input output.docx] [--source draft.md]
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, load_preferences, read_json, write_json  # noqa: E402
import wm_checks  # noqa: E402
import draft_model  # noqa: E402


def run(project: Path, input_name: str, source_name: str | None) -> dict:
    """跑全部检查。

    优先查 docx（那是真正要交付的东西）。docx 不存在时回落到 draft.md —— 这样
    用户在写作途中也能先跑一次看问题，不必等构建。
    """
    docx_path = project / input_name
    draft_path = project / (source_name or "draft.md")

    blocks = None
    checked = None
    if docx_path.is_file():
        try:
            blocks = draft_model.blocks_from_docx(docx_path)
            checked = str(docx_path)
        except Exception as exc:  # noqa: BLE001
            return {"ok": False, "reason": "unreadable", "message": f"docx 读取失败：{exc}"}
    elif draft_path.is_file():
        blocks = draft_model.blocks_from_markdown(draft_path)
        checked = str(draft_path)
    else:
        return {
            "ok": False,
            "reason": "missing-document",
            "message": f"既没有 {docx_path.name} 也没有 {draft_path.name}，先调 word_build 或写出 draft.md。",
        }

    preferences = load_preferences()
    spec = read_json(project / "spec.json", default={})
    if not isinstance(spec, dict):
        spec = {}
    gate_cfg = spec.get("gate") if isinstance(spec.get("gate"), dict) else {}

    result = wm_checks.run_all(blocks, preferences=preferences, spec=spec, gate_cfg=gate_cfg)
    char_count = sum(len(b["text"]) for b in blocks)

    report = {
        "ok": not result["p0"],
        "checkedAt": datetime.now(timezone.utc).isoformat(),
        "checked": checked,
        "source": "docx" if docx_path.is_file() else "draft",
        "charCount": char_count,
        "p0": result["p0"],
        "p1": result["p1"],
        "counts": result["counts"],
        "verdict": result["verdict"],
    }
    write_json(project / "gate-report.json", report)
    return report


def format_report(report: dict) -> str:
    """渲染成给模型读的文本。"""
    if report.get("verdict") is None:
        return f"质检无法进行：{report.get('message') or report.get('reason')}"
    lines = [
        f"# 质检结果：{report['verdict']}",
        "",
        f"- 检查对象：{report['checked']}（{report['source']}）",
        f"- 字符数：{report['charCount']}",
    ]
    for key, label in (("p0", "P0 阻断项"), ("p1", "P1 建议项")):
        items = report.get(key) or []
        if not items:
            lines.append(f"- {label}：无")
            continue
        lines.append(f"- {label}：{len(items)} 条")
        for item in items[:12]:
            where = f"第{item['line']}行" if item.get("line") else ""
            lines.append(f"  - [{item['id']}] {where} {item['message']}")
            if item.get("snippet"):
                lines.append(f"    - {item['snippet']}")
            if item.get("suggestion"):
                lines.append(f"    - 建议：{item['suggestion']}")
        if len(items) > 12:
            lines.append(f"  - …… 还有 {len(items) - 12} 条，见 gate-report.json")

    lines.append("")
    if report["verdict"] == "BLOCK":
        lines.append("结论：**不能交付**，改完 P0 再重跑 word_gate。")
    elif report["verdict"] == "WARN":
        lines.append("结论：可以交付，建议处理 P1。")
    else:
        lines.append("结论：**通过**，可以进入导出。")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 质检闸门")
    parser.add_argument("--project", required=True)
    parser.add_argument("--input", default="output.docx")
    parser.add_argument("--source", default="draft.md")
    parser.add_argument("--json", action="store_true", help="只输出机器可读 JSON")
    args = parser.parse_args(argv)

    report = run(Path(args.project).expanduser(), args.input, args.source)
    emit(report)
    if not args.json:
        print(format_report(report), file=sys.stderr)
    if report.get("verdict") is None:
        return 2
    if report.get("p0"):
        return 3
    if report.get("p1"):
        return 4
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
