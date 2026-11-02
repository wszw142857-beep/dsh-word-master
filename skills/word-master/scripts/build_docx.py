#!/usr/bin/env python3
"""Word Master — 从 Markdown 稿构建 DOCX。

输入是模型写的 Markdown（`draft.md`），输出是排版规范的 `.docx`。为什么要经过
Markdown 这一层：模型写正文的能力强于直接吐 OOXML，而排版是确定性工作，应该
交给代码。分工清楚，两边都稳。

支持的 Markdown 子集
--------------------
- `#` / `##` / `###` 三级标题（映射到 Word 内置标题样式，保证导航与目录可用）
- 普通段落（首行缩进、两端对齐、按偏好的行距）
- `-` / `*` 无序列表，`1.` 有序列表
- `> ` 引用块
- `| a | b |` 表格（第一行为表头，自动套三线表）
- `**粗体**`、`*斜体*`
- 空行分段

刻意不支持：嵌套超过两层的列表、脚注、嵌入图片（图片走 `word_images` 另配）。

版式全部来自规格快照里的 `formattingSnapshot`（冻结时从偏好库取的），脚本里
不写死任何字体字号。

用法
----
    python build_docx.py --project DIR [--input draft.md] [--output out.docx]
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import (  # noqa: E402
    emit,
    force_utf8_stdio,
    formatting_of,
    load_preferences,
    read_json,
)
from omml import extract_display, latex_to_omml  # noqa: E402

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ImportError:  # pragma: no cover - 环境缺失时给明确指令
    print("[ERR] 缺少 python-docx，先执行：pip install python-docx", file=sys.stderr)
    raise SystemExit(127)


INLINE_BOLD = re.compile(r"(\*\*.+?\*\*)")
INLINE_ITALIC = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")
LIST_UNORDERED = re.compile(r"^\s*[-*]\s+(.*)$")
LIST_ORDERED = re.compile(r"^\s*(\d+)[.)]\s+(.*)$")
TABLE_ROW = re.compile(r"^\s*\|(.+)\|\s*$")
TABLE_DIVIDER = re.compile(r"^\s*\|?[\s:|-]+\|?\s*$")


def set_run_font(run, latin: str, cjk: str, size_pt: float | None = None, bold: bool | None = None) -> None:
    """显式设置中西文字体。

    python-docx 的 `font.name` 只写 w:rFonts/@ascii，中文要另设 `w:eastAsia`，
    否则 Word 会用主题字体回退，看起来"设置没生效"。
    """
    run.font.name = latin
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)
    rfonts.set(qn("w:eastAsia"), cjk)


def apply_body_style(document, cfg: dict) -> None:
    """把正文样式绑到偏好版式上，让后加的段落自动继承。"""
    style = document.styles["Normal"]
    style.font.name = cfg["latinFont"]
    style.font.size = Pt(cfg["bodySizePt"])
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), cfg["latinFont"])
    rfonts.set(qn("w:hAnsi"), cfg["latinFont"])
    rfonts.set(qn("w:eastAsia"), cfg["bodyFont"])

    paragraph_format = style.paragraph_format
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    paragraph_format.line_spacing = cfg["lineSpacing"]
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)


def apply_page_setup(document, cfg: dict) -> None:
    """页面尺寸与页边距。"""
    for section in document.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(cfg["marginCm"]["top"])
        section.bottom_margin = Cm(cfg["marginCm"]["bottom"])
        section.left_margin = Cm(cfg["marginCm"]["left"])
        section.right_margin = Cm(cfg["marginCm"]["right"])


def set_paragraph_format(paragraph, cfg: dict, *, indent: bool) -> None:
    """正文段落格式：两端对齐、首行缩进、行距、段前段后 0。"""
    fmt = paragraph.paragraph_format
    fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.line_spacing = cfg["lineSpacing"]
    if indent:
        fmt.first_line_indent = Pt(cfg["bodySizePt"] * cfg["firstLineIndentChars"])


def add_inline(paragraph, text: str, cfg: dict, *, size_pt: float | None = None) -> None:
    """把一行里的粗体/斜体拆成多个 run 写进去。"""
    for bold_chunk in INLINE_BOLD.split(text):
        if bold_chunk == "":
            continue
        is_bold = bold_chunk.startswith("**") and bold_chunk.endswith("**") and len(bold_chunk) > 4
        content = bold_chunk[2:-2] if is_bold else bold_chunk
        for italic_chunk in INLINE_ITALIC.split(content):
            if italic_chunk == "":
                continue
            is_italic = (
                italic_chunk.startswith("*")
                and italic_chunk.endswith("*")
                and len(italic_chunk) > 2
                and not italic_chunk.startswith("**")
            )
            inner = italic_chunk[1:-1] if is_italic else italic_chunk
            run = paragraph.add_run(inner)
            set_run_font(run, cfg["latinFont"], cfg["bodyFont"], size_pt, is_bold or None)
            if is_italic:
                run.font.italic = True


def set_three_line_borders(table) -> None:
    """三线表：表格上下框线 + 表头下框线，其余无线。"""
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, size, val in (
        ("top", 12, "single"),
        ("bottom", 12, "single"),
        ("left", 0, "none"),
        ("right", 0, "none"),
        ("insideH", 0, "none"),
        ("insideV", 0, "none"),
    ):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), val)
        element.set(qn("w:sz"), str(size))
        element.set(qn("w:space"), "0")
        borders.append(element)
    tbl_pr.append(borders)

    for cell in table.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_borders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        tc_borders.append(bottom)
        tc_pr.append(tc_borders)


def add_table(document, rows: list[list[str]], cfg: dict) -> None:
    """按 Markdown 表格数据建三线表。"""
    if not rows:
        return
    columns = max(len(r) for r in rows)
    table = document.add_table(rows=len(rows), cols=columns)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for r_index, row in enumerate(rows):
        for c_index in range(columns):
            cell = table.cell(r_index, c_index)
            cell.text = ""
            paragraph = cell.paragraphs[0]
            text = row[c_index] if c_index < len(row) else ""
            add_inline(paragraph, text.strip(), cfg, size_pt=cfg["bodySizePt"] - 1)
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            if r_index == 0:
                for run in paragraph.runs:
                    run.font.bold = True
    set_three_line_borders(table)


def parse_markdown(text: str) -> list[tuple[str, object]]:
    """把 Markdown 切成块：(类型, 载荷)。"""
    blocks: list[tuple[str, object]] = []
    lines = text.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped == "":
            i += 1
            continue

        if stripped.startswith("#"):
            match = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            if match:
                blocks.append(("heading", (len(match.group(1)), match.group(2).strip())))
                i += 1
                continue

        if stripped.startswith("|") and TABLE_ROW.match(stripped):
            table_rows: list[list[str]] = []
            while i < len(lines):
                candidate = lines[i].strip()
                if not candidate.startswith("|"):
                    break
                if TABLE_DIVIDER.match(candidate):
                    i += 1
                    continue
                match = TABLE_ROW.match(candidate)
                if match:
                    table_rows.append([c.strip() for c in match.group(1).split("|")])
                i += 1
            if table_rows:
                blocks.append(("table", table_rows))
            continue

        # 块级公式：整行 $$...$$，也可能跨多行
        display = extract_display(stripped)
        if display is not None:
            blocks.append(("math", display))
            i += 1
            continue
        if stripped.startswith("$$"):
            # 跨行累加到闭合的 $$
            buffer = stripped[2:]
            i += 1
            while i < len(lines) and "$$" not in lines[i]:
                buffer += "\n" + lines[i]
                i += 1
            if i < len(lines):
                buffer += "\n" + lines[i].split("$$")[0]
                i += 1
            content = buffer.replace("$$", "").strip()
            if content != "":
                blocks.append(("math", content))
            continue

        unordered = LIST_UNORDERED.match(stripped)
        if unordered:
            blocks.append(("bullet", unordered.group(1).strip()))
            i += 1
            continue

        ordered = LIST_ORDERED.match(stripped)
        if ordered:
            blocks.append(("ordered", (ordered.group(1), ordered.group(2).strip())))
            i += 1
            continue

        if stripped.startswith(">"):
            blocks.append(("quote", stripped.lstrip(">").strip()))
            i += 1
            continue

        blocks.append(("paragraph", stripped))
        i += 1
    return blocks


def apply_heading_color(document, cfg: dict) -> None:
    """把标题样式改成黑色。

    python-docx 默认模板的 Heading 样式带 Word 主题蓝（#2F5496）。技术文档、
    论文、公文都用黑色标题；主题蓝在打印稿里像链接，在正式文档里是硬伤。
    这里直接改样式定义，让所有标题继承，不必逐段覆盖。
    """
    from docx.shared import RGBColor

    for level in range(0, 5):
        style_name = "Title" if level == 0 else f"Heading {level}"
        try:
            style = document.styles[style_name]
        except KeyError:
            continue
        style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        style.font.name = cfg["latinFont"]
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.append(rfonts)
        rfonts.set(qn("w:ascii"), cfg["latinFont"])
        rfonts.set(qn("w:hAnsi"), cfg["latinFont"])
        rfonts.set(qn("w:eastAsia"), cfg["bodyFont"])
        # 标题与下段同页，避免孤立在页尾
        style.paragraph_format.keep_with_next = True
        style.paragraph_format.space_before = Pt(cfg["bodySizePt"] * 0.75)
        style.paragraph_format.space_after = Pt(cfg["bodySizePt"] * 0.35)


def add_display_math(document, latex: str, cfg: dict) -> None:
    """插入一个居中的块级公式（Word 原生 OMML，可双击编辑）。

    两个坑都在这里踩过：
    1. 段落里不能留空 run —— 会在公式后多出一个行高。
    2. **必须把行距改成单倍**。正文是 1.5 倍行距，公式段若继承它，一个带分式的
       高行盒会被乘以 1.5，视觉上就是公式上下各多一大块空白。公式自己决定高度，
       行距必须放开。
    """
    paragraph = document.add_paragraph()
    fmt = paragraph.paragraph_format
    fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fmt.line_spacing_rule = WD_LINE_SPACING.SINGLE
    fmt.space_before = Pt(cfg["bodySizePt"] * 0.35)
    fmt.space_after = Pt(cfg["bodySizePt"] * 0.35)
    # 清掉可能存在的空 run
    for run_element in list(paragraph._p.findall(qn("w:r"))):
        paragraph._p.remove(run_element)
    try:
        paragraph._p.append(latex_to_omml_element(latex))
    except Exception as exc:  # noqa: BLE001 - 转不出来时降级为文本，别让构建失败
        print(f"[WARN] 公式转换失败，按文本插入：{exc}", file=sys.stderr)
        run = paragraph.add_run(latex)
        set_run_font(run, cfg["latinFont"], cfg["bodyFont"])


def latex_to_omml_element(latex: str):
    """把 LaTeX 转成 lxml 元素。"""
    from docx.oxml import parse_xml

    return parse_xml(latex_to_omml(latex))


def build(project: Path, input_name: str, output_name: str) -> dict:
    """主构建流程。"""
    spec = read_json(project / "spec.json", default=None)
    source = project / input_name
    if not source.is_file():
        return {"ok": False, "reason": "missing-draft", "message": f"找不到草稿 {source}，先写出 draft.md。"}

    preferences = load_preferences()
    cfg = formatting_of(preferences)
    if isinstance(spec, dict) and isinstance(spec.get("formattingSnapshot"), dict):
        snapshot = spec["formattingSnapshot"]
        cfg.update({k: v for k, v in snapshot.items() if k in cfg})

    text = source.read_text(encoding="utf-8-sig")
    document = Document()
    apply_body_style(document, cfg)
    apply_page_setup(document, cfg)
    apply_heading_color(document, cfg)

    title = ""
    if isinstance(spec, dict) and isinstance(spec.get("title"), str):
        title = spec["title"]

    blocks = parse_markdown(text)
    heading_count = 0
    paragraph_count = 0
    table_count = 0
    math_count = 0

    if title != "":
        first_is_title = bool(blocks) and blocks[0][0] == "heading" and blocks[0][1][0] == 1
        if not first_is_title:
            heading = document.add_heading(title, level=0)
            for run in heading.runs:
                set_run_font(run, cfg["latinFont"], cfg["bodyFont"])
            heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for kind, payload in blocks:
        if kind == "heading":
            level, heading_text = payload
            # 层级映射必须与 convert.py 的 style_map 对称，否则往返会丢一级：
            #   convert 把 Word 的 Title → `#`，Heading 1 → `##`
            #   这里就必须把 `#` → Title，`##` → Heading 1
            # 早期直接把 `#` 写成 Heading 1，于是 build→convert 往返把文档
            # 主标题降成了章节级（CI 断言"缺一级标题"抓到的就是这个）。
            word_level = 0 if level == 1 else min(level - 1, 9)
            heading = document.add_heading(heading_text, level=word_level)
            for run in heading.runs:
                set_run_font(run, cfg["latinFont"], cfg["bodyFont"])
            heading_count += 1
        elif kind == "paragraph":
            paragraph = document.add_paragraph()
            set_paragraph_format(paragraph, cfg, indent=True)
            add_inline(paragraph, payload, cfg)
            paragraph_count += 1
        elif kind == "bullet":
            paragraph = document.add_paragraph(style="List Bullet")
            set_paragraph_format(paragraph, cfg, indent=False)
            add_inline(paragraph, payload, cfg)
            segment = paragraph.paragraph_format
            segment.left_indent = Pt(cfg["bodySizePt"] * 2)
        elif kind == "ordered":
            _, ordered_text = payload
            paragraph = document.add_paragraph(style="List Number")
            set_paragraph_format(paragraph, cfg, indent=False)
            add_inline(paragraph, ordered_text, cfg)
            paragraph.paragraph_format.left_indent = Pt(cfg["bodySizePt"] * 2)
        elif kind == "quote":
            paragraph = document.add_paragraph()
            set_paragraph_format(paragraph, cfg, indent=False)
            paragraph.paragraph_format.left_indent = Pt(cfg["bodySizePt"] * 2)
            add_inline(paragraph, payload, cfg)
        elif kind == "table":
            add_table(document, payload, cfg)
            table_count += 1
        elif kind == "math":
            add_display_math(document, payload, cfg)
            math_count += 1

    target = project / output_name
    target.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(target))

    return {
        "ok": True,
        "path": str(target),
        "headings": heading_count,
        "paragraphs": paragraph_count,
        "tables": table_count,
        "formulas": math_count,
        "formatting": cfg,
        "title": title,
    }


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 构建 DOCX")
    parser.add_argument("--project", required=True)
    parser.add_argument("--input", default="draft.md")
    parser.add_argument("--output", default="output.docx")
    args = parser.parse_args(argv)

    result = build(Path(args.project).expanduser(), args.input, args.output)
    emit(result)
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
