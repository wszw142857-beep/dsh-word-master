#!/usr/bin/env python3
"""Word Master — 把 draft.md / docx 统一成块序列，供检查引擎使用。

两个提取器输出同一种形状（见 wm_checks 模块文档），规则层只认那个形状。这样
"审阅界面查 draft.md" 与 "质检门查 docx" 用的是同一套判据，不会漂移。

关键约定：`line` 字段
--------------------
- draft.md：**1-based 行号**，审阅界面据此定位并支持回写。
- docx：段落序号 + 1（同样 1-based，语义是"第几段"）。

审阅界面的回写只对 draft.md 生效 —— 回写 docx 会破坏样式。所以界面拿到的源
永远是 draft.md；docx 只用来做最终质检。
"""

from __future__ import annotations

import re
from pathlib import Path

LIST_UNORDERED = re.compile(r"^\s*[-*]\s+(.*)$")
LIST_ORDERED = re.compile(r"^\s*(\d+)[.)]\s+(.*)$")
TABLE_ROW = re.compile(r"^\s*\|(.+)\|\s*$")
TABLE_DIVIDER = re.compile(r"^\s*\|?[\s:|-]+\|?\s*$")


def _clean_inline(text: str) -> str:
    """去掉行内标记，供检查使用（判据针对"读者看到的字"）。"""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    return text


def blocks_from_markdown(path: str | Path) -> list[dict]:
    """解析 draft.md 成块序列。"""
    source = Path(path)
    raw = source.read_text(encoding="utf-8-sig")
    lines = raw.replace("\r\n", "\n").split("\n")

    blocks: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        line_no = i + 1  # 1-based

        if stripped == "":
            i += 1
            continue

        # 标题
        if stripped.startswith("#"):
            match = re.match(r"^(#{1,6})\s+(.*)$", stripped)
            if match:
                blocks.append(
                    {
                        "kind": "heading",
                        "level": len(match.group(1)),
                        "text": _clean_inline(match.group(2).strip()),
                        "raw": stripped,
                        "line": line_no,
                    }
                )
                i += 1
                continue

        # 表格：整块收成一个块，文本是各单元格拼接
        if stripped.startswith("|") and TABLE_ROW.match(stripped):
            start = line_no
            cells: list[str] = []
            while i < len(lines):
                candidate = lines[i].strip()
                if not candidate.startswith("|"):
                    break
                if TABLE_DIVIDER.match(candidate):
                    i += 1
                    continue
                match = TABLE_ROW.match(candidate)
                if match:
                    cells.extend(c.strip() for c in match.group(1).split("|"))
                i += 1
            blocks.append(
                {
                    "kind": "table",
                    "level": None,
                    "text": " | ".join(c for c in cells if c != ""),
                    "raw": "\n".join(lines[start - 1:i]),
                    "line": start,
                }
            )
            continue

        # 块级公式
        if stripped.startswith("$$"):
            start = line_no
            if stripped.endswith("$$") and len(stripped) > 4:
                i += 1
            else:
                i += 1
                while i < len(lines) and "$$" not in lines[i]:
                    i += 1
                if i < len(lines):
                    i += 1
            blocks.append(
                {
                    "kind": "math",
                    "level": None,
                    "text": "",
                    "raw": "\n".join(lines[start - 1:i]),
                    "line": start,
                }
            )
            continue

        unordered = LIST_UNORDERED.match(stripped)
        if unordered:
            blocks.append(
                {
                    "kind": "list",
                    "level": None,
                    "text": _clean_inline(unordered.group(1).strip()),
                    "raw": line,
                    "line": line_no,
                }
            )
            i += 1
            continue

        ordered = LIST_ORDERED.match(stripped)
        if ordered:
            blocks.append(
                {
                    "kind": "list",
                    "level": None,
                    "text": _clean_inline(ordered.group(2).strip()),
                    "raw": line,
                    "line": line_no,
                }
            )
            i += 1
            continue

        if stripped.startswith(">"):
            blocks.append(
                {
                    "kind": "quote",
                    "level": None,
                    "text": _clean_inline(stripped.lstrip(">").strip()),
                    "raw": line,
                    "line": line_no,
                }
            )
            i += 1
            continue

        blocks.append(
            {
                "kind": "paragraph",
                "level": None,
                "text": _clean_inline(stripped),
                "raw": line,
                "line": line_no,
            }
        )
        i += 1

    return blocks


def blocks_from_docx(path: str | Path) -> list[dict]:
    """解析 docx 成块序列（供最终质检用）。"""
    from docx import Document

    document = Document(str(path))
    blocks: list[dict] = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if text == "":
            continue
        style = (paragraph.style.name or "").strip()
        match = re.match(r"^(?:Heading|标题)\s*(\d+)$", style, re.IGNORECASE)
        level = None
        kind = "paragraph"
        if match:
            level = int(match.group(1))
            kind = "heading"
        elif style.lower() in ("title", "标题"):
            level = 0
            kind = "heading"
        blocks.append(
            {
                "kind": kind,
                "level": level,
                "text": text,
                "raw": text,
                "line": index + 1,
            }
        )
    for table in document.tables:
        cells = []
        for row in table.rows:
            for cell in row.cells:
                value = cell.text.strip()
                if value != "":
                    cells.append(value)
        if cells:
            blocks.append(
                {
                    "kind": "table",
                    "level": None,
                    "text": " | ".join(cells),
                    "raw": "",
                    "line": blocks[-1]["line"] + 1 if blocks else 1,
                }
            )
    return blocks


def source_text(path: str | Path) -> str:
    """读源文件原文。"""
    return Path(path).read_text(encoding="utf-8-sig")


def write_back(path: str | Path, original: str, replacements: list[dict]) -> dict:
    """把若干块级替换写回 draft.md。

    `replacements` 每项 `{line, kind, before, after}`：
    - 只替换**该行**（多行块用 `before` 全文匹配）；
    - 替换前校验 `before` 仍在文件里，否则拒绝（并发编辑保护）。

    刻意只做「整块替换」而不做行内定位：行内定位需要精确的字节偏移，一旦用户
    在浏览器里同时改了别的段落就会错位。整块匹配最稳。
    """
    target = Path(path)
    applied = []
    rejected = []

    for item in replacements:
        before = item.get("before")
        after = item.get("after")
        if not isinstance(before, str) or not isinstance(after, str):
            rejected.append({"line": item.get("line"), "reason": "invalid-payload"})
            continue
        if before.strip() == "":
            rejected.append({"line": item.get("line"), "reason": "empty-before"})
            continue
        if before not in original:
            rejected.append({"line": item.get("line"), "reason": "not-found"})
            continue
        count = original.count(before)
        if count > 1:
            rejected.append({"line": item.get("line"), "reason": f"ambiguous({count})"})
            continue
        original = original.replace(before, after, 1)
        applied.append({"line": item.get("line")})

    if applied:
        tmp = target.with_suffix(target.suffix + ".tmp")
        tmp.write_text(original, encoding="utf-8")
        import os

        os.replace(tmp, target)

    return {"ok": len(applied) > 0 or len(replacements) == 0, "applied": applied, "rejected": rejected}
