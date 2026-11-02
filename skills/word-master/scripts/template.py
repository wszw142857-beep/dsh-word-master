#!/usr/bin/env python3
"""Word Master — 模板：从现成 .docx 提取版式，存成可复用配置。

为什么 Word 的「模板」和 PPT 不一样
----------------------------------
PPT 模板是一套**版式骨架**（占位符摆位、配色、字体组合），页面从零拼装；Word
模板是**样式集合** —— 正文用宋体小四还是仿宋三号、标题层级怎么排、页边距多少、
页眉页脚写什么。内容流是线性的，没有"版式"可套。

所以真实需求是：**用户有一份单位/学校的规范文档，希望新文档长得一样。**
本模块就是做这件事：从那份文档里把样式读出来，变成一份可复用的配置。

四个动作
--------
- `inspect --template T.docx`  读出一份 .docx 的版式画像（不落盘）
- `save --template T.docx --id name`  提取并存进模板库
- `list`  列出已存的模板
- `apply --id name --project DIR`  把模板版式写进项目规格的 formattingSnapshot

存储位置：`$DSH_HOME/word-master/templates/<id>.json`

用法
----
    python template.py inspect --template 规范.docx
    python template.py save    --template 规范.docx --id 单位公文
    python template.py list
    python template.py apply   --id 单位公文 --project ./我的报告
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import (  # noqa: E402
    dsh_home,
    emit,
    force_utf8_stdio,
    formatting_of,
    load_preferences,
    read_json,
    write_json,
)

try:
    from docx import Document
    from docx.shared import Emu
except ImportError:  # pragma: no cover
    print("[ERR] 缺少 python-docx，先执行：pip install python-docx", file=sys.stderr)
    raise SystemExit(127)


HEADING_NAMES = ("Title", "标题", "Heading 1", "标题 1", "Heading 2", "标题 2", "Heading 3", "标题 3")


def templates_dir() -> Path:
    """模板库目录。"""
    return dsh_home() / "word-master" / "templates"


def emu_to_cm(value) -> float | None:
    """EMU → 厘米，保留两位。"""
    if value is None:
        return None
    try:
        return round(Emu(int(value)).cm, 2)
    except (TypeError, ValueError):
        return None


def _style_font(style) -> dict:
    """读一个样式的字体信息（中西文分别读）。"""
    font = getattr(style, "font", None)
    if font is None:
        return {}
    info: dict = {}
    if font.name:
        info["latin"] = font.name
    if font.size is not None:
        info["sizePt"] = round(font.size.pt, 1)

    # 中文字体在 w:eastAsia，python-docx 不给直接属性，读 XML
    try:
        rpr = style.element.get_or_add_rPr()
        rfonts = rpr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts")
        if rfonts is not None:
            east = rfonts.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia")
            if east:
                info["eastAsia"] = east
            ascii_font = rfonts.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii")
            if ascii_font and "latin" not in info:
                info["latin"] = ascii_font
    except Exception:  # noqa: BLE001 - 读不到就算了，不阻断
        pass

    try:
        rpr = style.element.get_or_add_rPr()
        color = rpr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color")
        if color is not None:
            val = color.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
            if val and val not in ("auto",):
                info["color"] = f"#{val}" if re.fullmatch(r"[0-9A-Fa-f]{6}", val) else val
    except Exception:  # noqa: BLE001
        pass

    para = getattr(style, "paragraph_format", None)
    if para is not None:
        if para.line_spacing is not None:
            try:
                info["lineSpacing"] = round(float(para.line_spacing), 2)
            except (TypeError, ValueError):
                pass
        if para.first_line_indent is not None:
            try:
                info["firstLineIndentPt"] = round(para.first_line_indent.pt, 1)
            except (TypeError, ValueError):
                pass
        if para.space_before is not None:
            try:
                info["spaceBeforePt"] = round(para.space_before.pt, 1)
            except (TypeError, ValueError):
                pass
        if para.space_after is not None:
            try:
                info["spaceAfterPt"] = round(para.space_after.pt, 1)
            except (TypeError, ValueError):
                pass
        if para.alignment is not None:
            info["alignment"] = str(para.alignment).split(".")[-1].split(" ")[0]
    return info


def _read_doc_defaults(doc) -> dict:
    """读 styles.xml 的 `docDefaults`。

    Word 的「真默认」不在 Normal 样式里，而在 `w:docDefaults/w:rPrDefault`。
    Normal 样式常常是空的（继承默认），此时只看 Normal 会得到一片 None，
    误判成"这份模板没有版式"。
    """
    info: dict = {}
    try:
        styles_element = doc.styles.element
        ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        defaults = styles_element.find(f"{ns}docDefaults")
        if defaults is None:
            return info
        rpr_default = defaults.find(f"{ns}rPrDefault")
        if rpr_default is not None:
            rpr = rpr_default.find(f"{ns}rPr")
            if rpr is not None:
                rfonts = rpr.find(f"{ns}rFonts")
                if rfonts is not None:
                    east = rfonts.get(f"{ns}eastAsia")
                    if east:
                        info["eastAsia"] = east
                    ascii_font = rfonts.get(f"{ns}ascii")
                    if ascii_font:
                        info["latin"] = ascii_font
                sz = rpr.find(f"{ns}sz")
                if sz is not None:
                    try:
                        info["sizePt"] = round(int(sz.get(f"{ns}val")) / 2, 1)
                    except (TypeError, ValueError):
                        pass
        ppr_default = defaults.find(f"{ns}pPrDefault")
        if ppr_default is not None:
            ppr = ppr_default.find(f"{ns}pPr")
            if ppr is not None:
                spacing = ppr.find(f"{ns}spacing")
                if spacing is not None:
                    line = spacing.get(f"{ns}line")
                    if line:
                        try:
                            # w:line 是二十分之一磅；240 = 单倍
                            info["lineSpacing"] = round(int(line) / 240, 2)
                        except (TypeError, ValueError):
                            pass
    except Exception:  # noqa: BLE001
        pass
    return info


def inspect_template(path: Path) -> dict:
    """读出一份 .docx 的版式画像。"""
    doc = Document(str(path))

    sections = []
    for section in doc.sections:
        sections.append(
            {
                "pageWidthCm": emu_to_cm(section.page_width),
                "pageHeightCm": emu_to_cm(section.page_height),
                "marginCm": {
                    "top": emu_to_cm(section.top_margin),
                    "bottom": emu_to_cm(section.bottom_margin),
                    "left": emu_to_cm(section.left_margin),
                    "right": emu_to_cm(section.right_margin),
                },
                "headerDistanceCm": emu_to_cm(section.header_distance),
                "footerDistanceCm": emu_to_cm(section.footer_distance),
            }
        )

    styles: dict = {}
    for name in HEADING_NAMES + ("Normal", "正文", "List Bullet", "List Number"):
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        info = _style_font(style)
        if info:
            styles[name] = info

    # Normal 常为空（继承 docDefaults）；空的时候用 docDefaults 兜底，
    # 否则会误判成"这份模板没有版式信息"。
    defaults = _read_doc_defaults(doc)
    if defaults:
        normal_key = "Normal" if "Normal" in styles else ("正文" if "正文" in styles else None)
        if normal_key is None:
            styles["Normal"] = dict(defaults)
        else:
            merged = dict(defaults)
            merged.update(styles[normal_key])  # 显式设置优先于默认
            styles[normal_key] = merged

    # 正文样板：取第一个非标题、非空的段落，作为"实际用的正文样式"参照
    sample_body = None
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text == "":
            continue
        style_name = (paragraph.style.name or "").strip()
        if re.match(r"^(Heading|标题)\s*\d+$", style_name, re.IGNORECASE):
            continue
        sample_body = {"styleName": style_name, "text": text[:60]}
        break

    headers = []
    for section in doc.sections:
        for part, label in ((section.header, "header"), (section.footer, "footer")):
            try:
                text = " | ".join(p.text.strip() for p in part.paragraphs if p.text.strip())
            except Exception:  # noqa: BLE001
                text = ""
            if text:
                headers.append({"where": label, "text": text[:120]})

    return {
        "source": str(path),
        "sourceName": path.name,
        "sections": sections,
        "styles": styles,
        "sampleBody": sample_body,
        "headerFooter": headers,
        "inspectedAt": datetime.now(timezone.utc).isoformat(),
    }


def profile_to_formatting(profile: dict, base: dict) -> dict:
    """把模板画像映射成我们 build 用的 formatting 配置。"""
    fmt = dict(base)
    sections = profile.get("sections") or []
    if sections:
        first = sections[0]
        margins = first.get("marginCm") or {}
        fmt["marginCm"] = {
            "top": margins.get("top", fmt["marginCm"]["top"]),
            "bottom": margins.get("bottom", fmt["marginCm"]["bottom"]),
            "left": margins.get("left", fmt["marginCm"]["left"]),
            "right": margins.get("right", fmt["marginCm"]["right"]),
        }
        if first.get("pageWidthCm") and first.get("pageHeightCm"):
            fmt["pageSizeCm"] = {"width": first["pageWidthCm"], "height": first["pageHeightCm"]}

    styles = profile.get("styles") or {}
    normal = styles.get("Normal") or styles.get("正文") or {}
    if normal.get("eastAsia"):
        fmt["bodyFont"] = normal["eastAsia"]
    if normal.get("latin"):
        fmt["latinFont"] = normal["latin"]
    if normal.get("sizePt"):
        fmt["bodySizePt"] = normal["sizePt"]
    if normal.get("lineSpacing"):
        fmt["lineSpacing"] = normal["lineSpacing"]
    if normal.get("firstLineIndentPt") and normal.get("sizePt"):
        fmt["firstLineIndentChars"] = round(normal["firstLineIndentPt"] / normal["sizePt"], 2)

    fmt["headingStyles"] = {
        name: styles[name] for name in ("Title", "标题", "Heading 1", "标题 1", "Heading 2", "标题 2", "Heading 3", "标题 3")
        if name in styles
    }
    return fmt


def save_template(profile: dict, template_id: str) -> dict:
    """存进模板库。"""
    safe = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", template_id).strip("_")
    if safe == "":
        return {"ok": False, "error": "模板 id 不能为空"}
    target = templates_dir() / f"{safe}.json"
    payload = {**profile, "id": safe, "savedAt": datetime.now(timezone.utc).isoformat()}
    write_json(target, payload)
    return {"ok": True, "id": safe, "path": str(target), "profile": payload}


def list_templates() -> list[dict]:
    """列出模板库。"""
    directory = templates_dir()
    if not directory.is_dir():
        return []
    rows = []
    for path in sorted(directory.glob("*.json")):
        data = read_json(path, default=None)
        if not isinstance(data, dict):
            continue
        styles = data.get("styles") or {}
        normal = styles.get("Normal") or styles.get("正文") or {}
        rows.append(
            {
                "id": data.get("id") or path.stem,
                "sourceName": data.get("sourceName"),
                "bodyFont": normal.get("eastAsia"),
                "bodySizePt": normal.get("sizePt"),
                "path": str(path),
            }
        )
    return rows


def apply_template(template_id: str, project: Path) -> dict:
    """把模板版式写进项目规格。"""
    safe = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", template_id).strip("_")
    path = templates_dir() / f"{safe}.json"
    profile = read_json(path, default=None)
    if not isinstance(profile, dict):
        return {"ok": False, "reason": "missing-template", "message": f"没有模板 {safe}，先用 save 存一份"}

    spec_path = project / "spec.json"
    spec = read_json(spec_path, default=None)
    if not isinstance(spec, dict):
        return {"ok": False, "reason": "missing-spec", "message": f"{project} 下没有 spec.json，先 word_spec init"}

    base = formatting_of(load_preferences())
    if isinstance(spec.get("formattingSnapshot"), dict):
        base.update({k: v for k, v in spec["formattingSnapshot"].items() if k in base or k == "headingStyles"})
    spec["formattingSnapshot"] = profile_to_formatting(profile, base)
    spec["templateId"] = safe
    write_json(spec_path, spec)
    return {
        "ok": True,
        "templateId": safe,
        "formatting": spec["formattingSnapshot"],
        "note": "版式已写进规格；之后 word_build 会用这套版式",
    }


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 模板")
    sub = parser.add_subparsers(dest="action", required=True)

    inspect_p = sub.add_parser("inspect")
    inspect_p.add_argument("--template", required=True)

    save_p = sub.add_parser("save")
    save_p.add_argument("--template", required=True)
    save_p.add_argument("--id", required=True)

    sub.add_parser("list")

    apply_p = sub.add_parser("apply")
    apply_p.add_argument("--id", required=True)
    apply_p.add_argument("--project", required=True)

    args = parser.parse_args(argv)

    if args.action == "inspect":
        source = Path(args.template).expanduser()
        if not source.is_file():
            emit({"ok": False, "reason": "missing-file", "message": f"找不到 {source}"})
            return 2
        emit({"ok": True, "profile": inspect_template(source)})
        return 0

    if args.action == "save":
        source = Path(args.template).expanduser()
        if not source.is_file():
            emit({"ok": False, "reason": "missing-file", "message": f"找不到 {source}"})
            return 2
        profile = inspect_template(source)
        result = save_template(profile, args.id)
        emit(result)
        return 0 if result.get("ok") else 2

    if args.action == "list":
        rows = list_templates()
        emit({"ok": True, "templates": rows, "count": len(rows)})
        return 0

    if args.action == "apply":
        result = apply_template(args.id, Path(args.project).expanduser())
        emit(result)
        return 0 if result.get("ok") else 2

    emit({"ok": False, "reason": "unknown-action"})
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
