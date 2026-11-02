#!/usr/bin/env python3
"""Word Master — 素材转换：PDF / Word / Excel / PPT / 网页 → Markdown。

为什么必须有这一层
------------------
写作的起点通常不是一段 Markdown，而是一份别人给的 PDF、一个旧 Word、一个网页。
没有转换层，用户只能自己先转好再喂进来 —— 那等于把最脏的一步推给用户。

与 ppt-master `source_to_md/` 的关系
-----------------------------------
同样按「每种输入一个转换器、输出统一为 Markdown + 资源目录」组织，但做减法：
它支持 pandoc 回退（.doc/.odt/.rtf/.tex），我们**只做纯 Python 能覆盖的格式**。
理由：pandoc 是外部依赖，用户机器上装没装不可控；而 docx/pdf/xlsx/pptx/html
恰好是写论文报告最高频的五种输入，纯 Python 全都能吃下。

输出约定
--------
    <out>/<name>.md          Markdown 正文
    <out>/<name>_files/      抽出的图片等资源（相对引用）

五个转换器
----------
- `pdf`   PyMuPDF 提取文本；按字号聚类推断标题层级
- `docx`  mammoth 转 HTML 再转 Markdown；保留三级标题、列表、表格
- `xlsx`  openpyxl 读每个 sheet，输出为 Markdown 表格
- `pptx`  python-pptx 逐页取文本框，输出为分节 Markdown
- `web`   BeautifulSoup 提正文（去脚本/导航/页脚）+ markdownify

用法
----
    python convert.py --kind pdf  --input a.pdf  --out-dir out/
    python convert.py --kind auto --input a.docx --out-dir out/
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, write_json  # noqa: E402


# ── 公共工具 ──────────────────────────────────────────────────────────────


def _clean_markdown(text: str) -> str:
    """压掉转换器常见的噪声：多余空行、孤立空白、行尾空格。"""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+$", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip() + "\n"


def _safe_stem(path: Path) -> str:
    """从输入文件取输出 stem，**带上来源后缀**避免同名覆盖。

    坑：同为 `test.docx` / `test.pdf` / `test.xlsx` 时，只按 stem 输出会全部写成
    `test.md`，后一次覆盖前一次 —— 批量转换时静默丢数据。所以固定加来源扩展名：
    `test.docx` → `test_docx`，`test.pdf` → `test_pdf`。
    """
    raw = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", path.stem).strip("_") or "converted"
    suffix = path.suffix.lower().lstrip(".").replace("htm", "html")
    return f"{raw}_{suffix}" if suffix else raw


def _write(out_dir: Path, stem: str, markdown: str) -> Path:
    """写 Markdown 文件并返回路径。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / f"{stem}.md"
    target.write_text(_clean_markdown(markdown), encoding="utf-8")
    return target


# ── PDF ───────────────────────────────────────────────────────────────────


def convert_pdf(source: Path, out_dir: Path, stem: str | None = None) -> dict:
    """PDF → Markdown。

    难点是**标题层级**：PDF 没有语义标签，只能按字号推断。做法是统计全文的
    字号分布，取最频繁的作为正文基准，明显大于基准的按大小分档 → 各级标题。

    这是启发式，不完美（封面、页码、页眉会被误收），但比"全是一堆散段落"
    有用得多；转换报告里会明确标注哪些行被判为标题，方便人工复核。
    """
    try:
        import pymupdf as fitz  # 新版推荐名
    except ImportError:
        try:
            import fitz  # 旧版
        except ImportError:
            return {"ok": False, "reason": "missing-dep", "message": "缺少 PyMuPDF，先 pip install pymupdf"}

    doc = fitz.open(str(source))
    pages: list[list[dict]] = []
    sizes: dict[float, int] = {}

    for page in doc:
        blocks = []
        for block in page.get_text("dict")["blocks"]:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                spans = line.get("spans", [])
                if not spans:
                    continue
                text = "".join(s.get("text", "") for s in spans).strip()
                if text == "":
                    continue
                size = max((s.get("size", 0) for s in spans), default=0)
                blocks.append({"text": text, "size": round(size, 1)})
                sizes[round(size, 1)] = sizes.get(round(size, 1), 0) + len(text)
        if blocks:
            pages.append(blocks)
    total_pages = doc.page_count
    doc.close()

    if not pages:
        return {"ok": False, "reason": "empty-pdf", "message": "PDF 里没有提取到文本（可能是扫描件，需要 OCR）"}

    # 出现字符数最多的字号 = 正文基准
    body_size = max(sizes, key=sizes.get) if sizes else 10.0
    # 明显大于正文的字号，按大小降序分档 → h1/h2/h3
    bigger = sorted({s for s in sizes if s > body_size * 1.12}, reverse=True)
    level_of = {size: min(index + 1, 3) for index, size in enumerate(bigger)}

    lines_out: list[str] = []
    heading_hits: list[str] = []
    for page_index, blocks in enumerate(pages):
        if page_index > 0:
            lines_out.append("")
        # 合并同段落的连续行：PDF 按视觉行切，同一段会被拆成多行
        buffer: list[str] = []
        for block in blocks:
            text = block["text"]
            level = level_of.get(block["size"])
            if level is not None and len(text) <= 80:
                if buffer:
                    lines_out.append(_join_paragraph(buffer))
                    buffer = []
                lines_out.append("")
                lines_out.append("#" * level + " " + text)
                lines_out.append("")
                heading_hits.append(f"h{level}: {text[:50]}")
            else:
                buffer.append(text)
                # 行尾有句末标点就认为段落结束
                if re.search(r"[。！？.!?:：]$", text):
                    lines_out.append(_join_paragraph(buffer))
                    buffer = []
                    lines_out.append("")
        if buffer:
            lines_out.append(_join_paragraph(buffer))
            lines_out.append("")

    markdown = "\n".join(lines_out)
    target = _write(out_dir, stem or _safe_stem(source), markdown)
    return {
        "ok": True,
        "kind": "pdf",
        "markdown": str(target),
        "pages": total_pages,
        "bodyFontSize": body_size,
        "headingsDetected": len(heading_hits),
        "headingPreview": heading_hits[:8],
        "note": "标题层级按字号推断，可能有误判；建议转换后先扫一遍标题",
    }


def _join_paragraph(lines: list[str]) -> str:
    """把 PDF 的多行合并成一段。中文行之间不插空格，英文插。"""
    if not lines:
        return ""
    out = lines[0]
    for nxt in lines[1:]:
        need_space = (
            bool(re.search(r"[A-Za-z0-9,;]$", out))
            and bool(re.match(r"^[A-Za-z0-9]", nxt))
        )
        out += (" " if need_space else "") + nxt
    return out


# ── DOCX ──────────────────────────────────────────────────────────────────


def convert_docx(source: Path, out_dir: Path, stem: str | None = None) -> dict:
    """Word → Markdown。

    走 mammoth（docx → HTML）再 markdownify（HTML → Markdown）这条链，而不是
    直接用 python-docx 遍历段落：mammoth 能正确保留列表编号、嵌套、表格结构与
    粗斜体，自己遍历要重新实现一遍这些语义，且很容易漏。

    ⚠ **必须自定义 style_map**。mammoth 默认把 Word 的 `Title` 与 `Heading 1`
    都映射成 `<h1>` —— 于是"文档大标题"和"第一章"变成同一级，层级被压平，
    审阅与重建都会受影响。这里把 Title 单独映射成 h1、Heading 1/2/3 顺延成
    h2/h3/h4，恢复"标题 > 章 > 节"的真实结构。
    """
    try:
        import mammoth
        from markdownify import markdownify as md
    except ImportError as exc:
        return {"ok": False, "reason": "missing-dep", "message": f"缺少依赖：{exc}（pip install mammoth markdownify）"}

    # Word 样式 → HTML 标签。左侧是样式名，右侧是目标标签。
    style_map = (
        "p[style-name='Title'] => h1:fresh\n"
        "p[style-name='标题'] => h1:fresh\n"
        "p[style-name='Subtitle'] => p:fresh\n"
        "p[style-name='Heading 1'] => h2:fresh\n"
        "p[style-name='标题 1'] => h2:fresh\n"
        "p[style-name='Heading 2'] => h3:fresh\n"
        "p[style-name='标题 2'] => h3:fresh\n"
        "p[style-name='Heading 3'] => h4:fresh\n"
        "p[style-name='标题 3'] => h4:fresh\n"
        "p[style-name='Heading 4'] => h5:fresh\n"
        "p[style-name='标题 4'] => h5:fresh\n"
    )

    base_stem = stem or _safe_stem(source)
    resource_dir = out_dir / f"{base_stem}_files"
    resource_dir.mkdir(parents=True, exist_ok=True)
    images: list[str] = []

    def save_image(image):
        """mammoth 回调：把内嵌图片落盘并返回相对路径。"""
        with image.open() as fh:
            data = fh.read()
        name = f"image-{len(images) + 1}{_guess_ext(image.content_type)}"
        (resource_dir / name).write_bytes(data)
        images.append(name)
        return {"src": f"{base_stem}_files/{name}"}

    with source.open("rb") as fh:
        result = mammoth.convert_to_html(
            fh,
            convert_image=mammoth.images.img_element(save_image),
            style_map=style_map,
        )
    html = result.value
    messages = [str(m) for m in result.messages][:10]

    markdown = md(html, heading_style="ATX", strip=["script", "style"])
    markdown = _fix_docx_markdown(markdown)
    target = _write(out_dir, stem or _safe_stem(source), markdown)

    if not images:
        try:
            resource_dir.rmdir()
        except OSError:
            pass

    return {
        "ok": True,
        "kind": "docx",
        "markdown": str(target),
        "images": len(images),
        "resourceDir": str(resource_dir) if images else None,
        "warnings": messages,
    }


def _fix_docx_markdown(markdown: str) -> str:
    """修 mammoth 输出的两个常见毛病。

    1. **Word 的 Title 样式变成裸文字**。mammoth 把 `<h1>` 之外的标题样式映射成
       `<p>`，于是文档大标题没有 `#`，转出来第一行是普通段落。这里把"全文第一个
       非空行且后面紧跟 `#` 标题"的情况下补成 `# `。

    2. **空表头行**。mammoth 会把表格第一行放进 `<thead>`，若原表没有表头（首行
       就是数据），markdownify 会生成一行空的 `|  |  |`。这里把全空表头行删掉，
       并把原首行数据提为表头。
    """
    lines = markdown.split("\n")

    # ── 1. 裸标题补 `#` ──
    first_content = next((i for i, l in enumerate(lines) if l.strip() != ""), None)
    if first_content is not None and first_content == 0:
        line = lines[0].strip()
        looks_like_text = (
            line != ""
            and not line.startswith(("#", "-", "*", "|", ">", "```"))
            and len(line) <= 40
            and not re.search(r"[。！？.!?]$", line)
        )
        has_heading_after = any(l.strip().startswith("#") for l in lines[1:6])
        if looks_like_text and has_heading_after:
            lines[0] = "# " + line

    # ── 2. 空表头 ──
    out: list[str] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        header = re.fullmatch(r"\s*\|(\s*\|)+\s*", line)
        separator = i + 1 < len(lines) and re.fullmatch(r"\s*\|(\s*-+\s*\|)+\s*", lines[i + 1])
        if header and separator:
            cand = i + 2
            if cand < len(lines) and lines[cand].strip().startswith("|"):
                data_line = lines[cand]
                cells = [c.strip() for c in data_line.strip().strip("|").split("|")]
                cols = len(cells)
                out.append("| " + " | ".join(cells) + " |")
                out.append("| " + " | ".join("---" for _ in range(cols)) + " |")
                i = cand + 1
                continue
            else:
                i += 2
                continue
        out.append(line)
        i += 1

    return "\n".join(out)


def _guess_ext(content_type: str | None) -> str:
    """按 MIME 猜扩展名。"""
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/svg+xml": ".svg",
        "image/webp": ".webp",
        "image/tiff": ".tif",
        "image/bmp": ".bmp",
    }
    return mapping.get((content_type or "").split(";")[0].strip(), ".png")
    """按 MIME 猜扩展名。"""
    mapping = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/svg+xml": ".svg",
        "image/webp": ".webp",
        "image/tiff": ".tif",
        "image/bmp": ".bmp",
    }
    return mapping.get((content_type or "").split(";")[0].strip(), ".png")


# ── XLSX ──────────────────────────────────────────────────────────────────


def convert_xlsx(source: Path, out_dir: Path, stem: str | None = None) -> dict:
    """Excel → Markdown 表格。

    每个 sheet 一节，输出为管道表格。第一行当表头（Excel 数据表几乎都这样）。
    """
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {"ok": False, "reason": "missing-dep", "message": "缺少 openpyxl"}

    wb = load_workbook(str(source), data_only=True, read_only=True)
    parts: list[str] = []
    sheet_info: list[dict] = []

    for sheet in wb.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            if row is None:
                continue
            cells = ["" if v is None else str(v).strip() for v in row]
            if all(c == "" for c in cells):
                continue
            rows.append(cells)
        if not rows:
            sheet_info.append({"sheet": sheet.title, "rows": 0, "skipped": True})
            continue

        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        header, body = rows[0], rows[1:]

        lines = [f"## {sheet.title}", ""]
        lines.append("| " + " | ".join(_cell(c) for c in header) + " |")
        lines.append("| " + " | ".join("---" for _ in header) + " |")
        for row in body:
            lines.append("| " + " | ".join(_cell(c) for c in row) + " |")
        lines.append("")
        parts.append("\n".join(lines))
        sheet_info.append({"sheet": sheet.title, "rows": len(body), "columns": width})

    wb.close()

    if not parts:
        return {"ok": False, "reason": "empty-workbook", "message": "工作簿里没有数据"}

    markdown = "\n".join(parts)
    target = _write(out_dir, stem or _safe_stem(source), markdown)
    return {"ok": True, "kind": "xlsx", "markdown": str(target), "sheets": sheet_info}


def _cell(value: str) -> str:
    """单元格文本转义：管道符会破坏表格结构。"""
    return value.replace("|", "\\|").replace("\n", " ")


# ── PPTX ──────────────────────────────────────────────────────────────────


def convert_pptx(source: Path, out_dir: Path, stem: str | None = None) -> dict:
    """PowerPoint → Markdown。

    每页一节。标题占位符当 `##`，其余文本框按段落输出，列表保留层级。
    """
    try:
        from pptx import Presentation
    except ImportError:
        return {"ok": False, "reason": "missing-dep", "message": "缺少 python-pptx"}

    prs = Presentation(str(source))
    parts: list[str] = []
    slide_info: list[dict] = []

    for index, slide in enumerate(prs.slides, start=1):
        title = ""
        body_blocks: list[str] = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            paragraphs = []
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs).strip()
                if text == "":
                    continue
                # 缩进层级 → 列表 / 子列表
                level = getattr(para, "level", 0) or 0
                paragraphs.append(("  " * level + "- " + text) if level > 0 else text)
            if not paragraphs:
                continue
            if not title and shape == slide.shapes.title:
                title = paragraphs[0].lstrip("- ")
            else:
                body_blocks.extend(paragraphs)

        heading = title if title else f"第 {index} 页"
        parts.append(f"## {heading}")
        parts.append("")
        if body_blocks:
            parts.extend(body_blocks)
            parts.append("")
        else:
            parts.append("")
        slide_info.append({"slide": index, "title": heading, "blocks": len(body_blocks)})

    if not parts:
        return {"ok": False, "reason": "empty-deck", "message": "演示文稿里没有文本"}

    markdown = "\n".join(parts)
    target = _write(out_dir, stem or _safe_stem(source), markdown)
    return {"ok": True, "kind": "pptx", "markdown": str(target), "slides": slide_info}


# ── WEB ───────────────────────────────────────────────────────────────────


def convert_web(source: str, out_dir: Path, stem: str = "web") -> dict:
    """网页 → Markdown。

    先删掉脚本、样式、导航、页脚、侧栏这些与正文无关的容器，再交给 markdownify。
    不这么做的话，转出来的 Markdown 会塞满菜单项和相关链接。
    """
    try:
        import requests
        from bs4 import BeautifulSoup
        from markdownify import markdownify as md
    except ImportError as exc:
        return {"ok": False, "reason": "missing-dep", "message": f"缺少依赖：{exc}"}

    try:
        response = requests.get(
            source,
            timeout=30,
            headers={"User-Agent": "Mozilla/5.0 (compatible; WordMaster/1.0)"},
        )
        response.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "reason": "fetch-failed", "message": str(exc)[:300]}

    # 让 requests 自己按 header/meta 猜编码
    if response.encoding is None or response.encoding.lower() == "iso-8859-1":
        response.encoding = response.apparent_encoding or "utf-8"

    soup = BeautifulSoup(response.text, "html.parser")

    for tag in soup(["script", "style", "noscript", "iframe", "nav", "footer", "aside", "form"]):
        tag.decompose()
    for selector in ("header.site-header", "div.sidebar", "div.nav", "nav", ".menu", ".related"):
        for node in soup.select(selector):
            node.decompose()

    title = soup.title.get_text(strip=True) if soup.title else ""
    # 优先取语义化的正文容器
    main = soup.find("article") or soup.find("main") or soup.body or soup

    markdown = md(str(main), heading_style="ATX", strip=["script", "style"])
    if title:
        markdown = f"# {title}\n\n{markdown}"

    target = _write(out_dir, stem, markdown)
    return {
        "ok": True,
        "kind": "web",
        "markdown": str(target),
        "title": title,
        "url": source,
        "chars": len(markdown),
    }


# ── 分派 ──────────────────────────────────────────────────────────────────


KIND_BY_EXT = {
    ".pdf": "pdf",
    ".docx": "docx",
    ".doc": "docx-legacy",
    ".xlsx": "xlsx",
    ".xlsm": "xlsx",
    ".pptx": "pptx",
    ".html": "html",
    ".htm": "html",
}


def detect_kind(source: str) -> str:
    """按扩展名或 URL 判断类型。"""
    if re.match(r"^https?://", source, re.IGNORECASE):
        return "web"
    return KIND_BY_EXT.get(Path(source).suffix.lower(), "unknown")


def convert(kind: str, source: str, out_dir: Path, stem: str | None = None) -> dict:
    """统一入口。"""
    if kind == "auto":
        kind = detect_kind(source)

    if kind == "pdf":
        return convert_pdf(Path(source), out_dir, stem)

    if kind == "docx":
        return convert_docx(Path(source), out_dir, stem)

    if kind == "docx-legacy":
        return {
            "ok": False,
            "reason": "legacy-doc",
            "message": "旧版 .doc 不支持直接转换。用 Word / LibreOffice 另存为 .docx 后再转，"
            "或者用 word_env 检查 LibreOffice 是否可用（可用时可先转 docx）。",
        }

    if kind == "xlsx":
        return convert_xlsx(Path(source), out_dir, stem)

    if kind == "pptx":
        return convert_pptx(Path(source), out_dir, stem)

    if kind == "web":
        return convert_web(source, out_dir, stem or "web")

    if kind == "html":
        path = Path(source)
        if not path.is_file():
            return {"ok": False, "reason": "missing-file", "message": f"找不到 {source}"}
        try:
            from bs4 import BeautifulSoup
            from markdownify import markdownify as md
        except ImportError as exc:
            return {"ok": False, "reason": "missing-dep", "message": f"缺少依赖：{exc}"}
        soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "html.parser")
        for tag in soup(["script", "style", "noscript"]):
            tag.decompose()
        target = _write(out_dir, _safe_stem(path), md(str(soup), heading_style="ATX"))
        return {"ok": True, "kind": "html", "markdown": str(target)}

    return {
        "ok": False,
        "reason": "unsupported-kind",
        "message": f"不支持的类型：{kind}。支持 pdf / docx / xlsx / pptx / html / 网页 URL。",
    }


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 素材转换")
    parser.add_argument("--kind", default="auto", choices=["auto", "pdf", "docx", "xlsx", "pptx", "web", "html"])
    parser.add_argument("--input", required=True, help="源文件路径或网页 URL")
    parser.add_argument("--out-dir", required=True, help="输出目录")
    parser.add_argument("--stem", default="", help="输出文件名（不含扩展名）")
    args = parser.parse_args(argv)

    result = convert(args.kind, args.input, Path(args.out_dir).expanduser(), args.stem or None)
    emit(result)
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
