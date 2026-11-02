#!/usr/bin/env python3
"""Word Master — 导出（DOCX 定稿 + 可选 PDF）。

导出不是简单复制：它做三件事
1. **补核心属性**（标题、作者、语言），避免打开时「文档属性」一片空白；
2. **校验可打开性**：重新读一遍产物，确认 zip 结构完好、段落数非零；
3. **可选转 PDF**（走 LibreOffice，与渲染闸门同一套引擎）。

用法
----
    python export.py --project DIR [--input output.docx] [--name 定稿.docx] [--pdf]
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, iter_paragraph_texts, read_json  # noqa: E402

try:
    from docx import Document
except ImportError:  # pragma: no cover
    print("[ERR] 缺少 python-docx，先执行：pip install python-docx", file=sys.stderr)
    raise SystemExit(127)


SOFFICE_CANDIDATES = (
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/usr/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


def find_soffice() -> str | None:
    """定位 LibreOffice。"""
    found = shutil.which("soffice")
    if found:
        return found
    for candidate in SOFFICE_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


def export(project: Path, input_name: str, name: str, want_pdf: bool) -> dict:
    """执行导出。"""
    source = project / input_name
    if not source.is_file():
        return {"ok": False, "reason": "missing-document", "message": f"找不到 {source}，先调 word_build。"}

    spec = read_json(project / "spec.json", default={})
    title = spec.get("title") if isinstance(spec, dict) else None

    document = Document(str(source))
    properties = document.core_properties
    if isinstance(title, str) and title.strip() != "":
        properties.title = title.strip()
    if properties.author in (None, ""):
        properties.author = "Word Master"
    properties.comments = ""
    properties.language = "zh-CN"

    target = project / name
    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target))

    # 回读校验：能重新打开、且确实有内容
    try:
        reopened = Document(str(target))
        paragraphs = len(iter_paragraph_texts(reopened))
        tables = len(reopened.tables)
        sections = len(reopened.sections)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "reason": "unreadable-output", "message": f"导出的文件无法重新打开：{exc}"}

    result = {
        "ok": True,
        "docx": str(target),
        "sizeBytes": target.stat().st_size,
        "paragraphs": paragraphs,
        "tables": tables,
        "sections": sections,
    }

    if paragraphs == 0:
        result.update({"ok": False, "reason": "empty-output", "message": "导出的文档没有段落，检查构建输入。"})
        return result

    if want_pdf:
        soffice = find_soffice()
        if soffice is None:
            result["pdf"] = None
            result["pdfNote"] = "未安装 LibreOffice，跳过 PDF。"
        else:
            try:
                completed = subprocess.run(
                    [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(project), str(target)],
                    capture_output=True,
                    timeout=180,
                    check=False,
                )
                pdf = project / (target.stem + ".pdf")
                if pdf.is_file():
                    result["pdf"] = str(pdf)
                    result["pdfSizeBytes"] = pdf.stat().st_size
                else:
                    result["pdf"] = None
                    result["pdfNote"] = (completed.stderr or completed.stdout).decode("utf-8", errors="replace")[:400]
            except (OSError, subprocess.TimeoutExpired) as exc:
                result["pdf"] = None
                result["pdfNote"] = str(exc)

    return result


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 导出")
    parser.add_argument("--project", required=True)
    parser.add_argument("--input", default="output.docx")
    parser.add_argument("--name", default="final.docx")
    parser.add_argument("--pdf", action="store_true")
    args = parser.parse_args(argv)

    result = export(Path(args.project).expanduser(), args.input, args.name, args.pdf)
    emit(result)
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
