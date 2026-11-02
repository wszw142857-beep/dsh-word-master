#!/usr/bin/env python3
"""Word Master — 渲染闸门（docx → PDF → PNG 预览）。

质检门（gate.py）查的是结构，这一层查的是**看起来对不对**：字体有没有回退、
表格线有没有错位、标题有没有孤行、中英混排有没有挤在一起。结构全过但排版崩掉
的文档很常见，所以交付前必须真的看一眼。

为什么走 LibreOffice 而不是在 Python 里画：只有真正的 Word 排版引擎（LibreOffice
的 DOCX 渲染器已足够接近）才能暴露字体回退和分页问题。缺失 LibreOffice 时降级为
"跳过渲染"，并在报告里明确说明，**不**假装通过。

用法
----
    python render.py --project DIR [--input output.docx] [--dpi 110]
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, read_json, write_json  # noqa: E402


SOFFICE_CANDIDATES = (
    "C:/Program Files/LibreOffice/program/soffice.exe",
    "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


def find_soffice() -> str | None:
    """定位 LibreOffice 可执行文件。"""
    override = os.environ.get("DSH_WORD_SOFFICE", "").strip()
    if override != "" and Path(override).is_file():
        return override
    found = shutil.which("soffice")
    if found:
        return found
    for candidate in SOFFICE_CANDIDATES:
        if Path(candidate).is_file():
            return candidate
    return None


def run(cmd: list[str], timeout: int = 180) -> tuple[int, str]:
    """跑一个命令，返回 (退出码, 合并输出)。"""
    try:
        completed = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return -1, str(exc)
    out = completed.stdout.decode("utf-8", errors="replace")
    err = completed.stderr.decode("utf-8", errors="replace")
    return completed.returncode, (out + "\n" + err).strip()


def convert_pdf(soffice: str, source: Path, out_dir: Path) -> tuple[int, str, Path | None]:
    """docx → pdf。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    code, output = run(
        [
            soffice,
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(source),
        ]
    )
    pdf = out_dir / (source.stem + ".pdf")
    return code, output, (pdf if pdf.is_file() else None)


def rasterize(pdf: Path, out_dir: Path, dpi: int) -> dict:
    """pdf → png 逐页。

    优先 PyMuPDF（项目依赖清单里已有），退回 pdftoppm。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    images: list[str] = []
    try:
        import fitz  # type: ignore
    except ImportError:
        fitz = None

    if fitz is not None:
        with fitz.open(str(pdf)) as doc:
            for index in range(doc.page_count):
                page = doc.load_page(index)
                zoom = dpi / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                target = out_dir / f"page-{index + 1:02d}.png"
                pix.save(str(target))
                images.append(str(target))
        return {"ok": True, "engine": "pymupdf", "images": images}

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm is None:
        return {"ok": False, "engine": "none", "images": [], "message": "既没有 PyMuPDF 也没有 pdftoppm，无法生成预览图"}

    code, output = run([pdftoppm, "-png", "-r", str(dpi), str(pdf), str(out_dir / "page")])
    if code != 0:
        return {"ok": False, "engine": "pdftoppm", "images": [], "message": output}
    images = sorted(str(p) for p in out_dir.glob("page*.png"))
    return {"ok": True, "engine": "pdftoppm", "images": images}


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 渲染闸门")
    parser.add_argument("--project", required=True)
    parser.add_argument("--input", default="output.docx")
    parser.add_argument("--dpi", type=int, default=110)
    args = parser.parse_args(argv)

    project = Path(args.project).expanduser()
    source = project / args.input
    if not source.is_file():
        emit({"ok": False, "reason": "missing-document", "message": f"找不到 {source}"})
        return 2

    soffice = find_soffice()
    if soffice is None:
        report = {
            "ok": True,
            "skipped": True,
            "reason": "no-libreoffice",
            "message": "未安装 LibreOffice，渲染闸门跳过（结构质检仍有效）。装好后重跑可得到逐页预览图。",
        }
        write_json(project / "render-report.json", report)
        emit(report)
        return 0

    pdf_dir = project / "render"
    code, output, pdf = convert_pdf(soffice, source, pdf_dir)
    if pdf is None:
        report = {
            "ok": False,
            "reason": "convert-failed",
            "exitCode": code,
            "message": output[:1200],
            "soffice": soffice,
        }
        write_json(project / "render-report.json", report)
        emit(report)
        return 3

    pages_dir = pdf_dir / "pages"
    raster = rasterize(pdf, pages_dir, args.dpi)
    report = {
        "ok": True,
        "soffice": soffice,
        "pdf": str(pdf),
        "pages": len(raster.get("images", [])),
        "engine": raster.get("engine"),
        "images": raster.get("images", []),
        "message": "逐页预览已生成，必须真的看过图再宣布完成。" if raster.get("ok") else raster.get("message"),
    }
    write_json(project / "render-report.json", report)

    # 供模型直接读的清单
    if raster.get("ok"):
        manifest = project / "render" / "pages.txt"
        manifest.write_text("\n".join(raster["images"]) + "\n", encoding="utf-8")
        report["manifest"] = str(manifest)

    emit(report)
    return 0 if raster.get("ok") else 4


if __name__ == "__main__":
    raise SystemExit(main())
