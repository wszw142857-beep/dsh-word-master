#!/usr/bin/env python3
"""Word Master — 检查规则引擎（纯函数，不碰文件）。

为什么单独成模块
----------------
同一个判据要在两个地方生效：质检门查最终 docx，审阅界面查 draft.md。如果各写
一套，迟早漂移——改了一边忘了另一边，用户看到"界面说没问题、交付却被打回"。

所以规则只写一次，作用在统一的**块序列**上：

    [{"line": 12, "text": "...", "kind": "paragraph", "level": None}, ...]

- `line`：该块在源文件里的起始行（draft.md 是 1-based 行号；docx 是段落序号）。
- `kind`：`heading` / `paragraph` / `table` / `list` / `math` / `quote`
- `level`：标题级别（1 起），非标题为 None

两个提取器（`draft_model.py`）负责把 markdown 与 docx 都转成这个形状，规则层只
认这个形状。

每条发现（finding）
-------------------
    {
      "id": "fullwidth-brackets",
      "severity": "P0" | "P1",
      "line": 23,              # 定位到源文件行
      "message": "…",
      "snippet": "…",          # 出问题的原文片段
      "suggestion": "…",       # 怎么改（可空）
      "autofix": "…" | None,   # 能自动改的新文本（作用于该块）
    }
"""

from __future__ import annotations

import re

# 全角括号
FULLWIDTH_OPEN = "（"
FULLWIDTH_CLOSE = "）"

# 节首句句式："为了……，需要……"
# 逗号后允许 0 个字符直接接"需要"——"为了判断风险，需要先…"是最常见的合法写法。
OPENER_PATTERN = re.compile(r"^为了.{2,}?[，,]\s*(?:需要|需|应当|应|要|须)")

NUMBER_PATTERN = re.compile(r"\d+(?:\.\d+)?")

# 方法章标题识别词
METHOD_HEADINGS = ("方法", "材料与方法", "研究方法", "method", "materials and methods")
RESULTS_HEADINGS = ("结果", "结果与分析", "results")

# 方法章允许出现的数字
ALLOWED_NUMBER_PATTERNS = (
    re.compile(r"\b(?:HC|EC|LC|IC|NOEC|LOEC)\d*\b", re.IGNORECASE),
    re.compile(r"\b(?:表|图|Table|Figure)\s*[A-Z]?\d+\b", re.IGNORECASE),
    re.compile(r"\b(?:19|20)\d{2}\b"),
    re.compile(r"\br\s*/\s*\(\s*Kd\s*\+\s*r\s*\)"),
    re.compile(r"\b(?:AICc|AIC|BIC|R2|R²|p)\s*[=<>]\s*0?\.\d+"),
    re.compile(r"\b10\s*年\b|\b\d+\s*次抽样\b|\b5000\s*次\b"),
    re.compile(r"\b\d+(?:\.\d+)?\s*(?:%|‰|mg\/L|μg\/L|ng\/L|m|km|cm|mm|kg|g|h|min|s|℃)\b"),
)

BANNED_LABELS = {
    "defensive": "防御性表述",
    "delivery": "交付性用语",
    "meta": "元表述",
    "process": "过程性叙述",
}


def _snippet(text: str, index: int, width: int = 18) -> str:
    """截取命中位置附近的片段。"""
    start = max(0, index - width)
    end = min(len(text), index + width)
    prefix = "…" if start > 0 else ""
    suffix = "…" if end < len(text) else ""
    return f"{prefix}{text[start:end]}{suffix}"


def _is_method_heading(title: str) -> bool:
    """标题是否属于方法章节。"""
    lowered = title.lower()
    return any(word in lowered for word in METHOD_HEADINGS)


def _is_results_heading(title: str) -> bool:
    """标题是否属于结果章节。"""
    lowered = title.lower()
    return any(word in lowered for word in RESULTS_HEADINGS)


def strip_allowed_numbers(text: str) -> str:
    """抹掉白名单数字，剩下的才是违规数字。"""
    cleaned = text
    for pattern in ALLOWED_NUMBER_PATTERNS:
        cleaned = pattern.sub("", cleaned)
    return cleaned


# ── 各项检查 ──────────────────────────────────────────────────────────────


def check_fullwidth_brackets(blocks, limit: int = 0, scope: str = "body") -> list[dict]:
    """全角括号。可自动修复。"""
    findings = []
    for block in blocks:
        if block["kind"] in ("math",):
            continue
        if scope == "body" and block["kind"] == "table":
            pass  # 表头也要查，用户明确说过"表头也不例外"
        text = block["text"]
        total = text.count(FULLWIDTH_OPEN) + text.count(FULLWIDTH_CLOSE)
        if total == 0:
            continue
        findings.append(
            {
                "id": "fullwidth-brackets",
                "severity": "P0",
                "line": block["line"],
                "message": f"全角括号 {total} 处，应为半角 ()",
                "snippet": _snippet(text, text.find(FULLWIDTH_OPEN if FULLWIDTH_OPEN in text else FULLWIDTH_CLOSE)),
                "suggestion": "把 （） 换成 ()",
                "autofix": text.replace(FULLWIDTH_OPEN, "(").replace(FULLWIDTH_CLOSE, ")"),
            }
        )
    if limit > 0:
        return findings[:limit]
    return findings


def check_banned_phrases(blocks, tables: dict[str, list[str]]) -> list[dict]:
    """禁用词。不自动改——替换词依赖语义，让模型/用户判断。"""
    findings = []
    for block in blocks:
        if block["kind"] in ("math",):
            continue
        text = block["text"]
        for key, words in tables.items():
            for word in words:
                index = text.find(word)
                if index < 0:
                    continue
                findings.append(
                    {
                        "id": f"banned-{key}",
                        "severity": "P0",
                        "line": block["line"],
                        "message": f"命中{BANNED_LABELS.get(key, key)}「{word}」",
                        "snippet": _snippet(text, index),
                        "suggestion": "删除该表述，或改为正面陈述具体结论",
                        "autofix": None,
                    }
                )
                break  # 同一段落同一类只报一次，避免刷屏
    return findings


def check_heading_jumps(blocks) -> list[dict]:
    """标题跳级。"""
    findings = []
    previous = None
    for block in blocks:
        if block["kind"] != "heading" or block.get("level") is None:
            continue
        level = block["level"]
        if level > 0 and previous is not None and level - previous > 1:
            findings.append(
                {
                    "id": "heading-jump",
                    "severity": "P0",
                    "line": block["line"],
                    "message": f"标题从 {previous} 级跳到 {level} 级",
                    "snippet": block["text"][:40],
                    "suggestion": f"补一个 {previous + 1} 级标题，或把该标题降为 {previous + 1} 级",
                    "autofix": None,
                }
            )
        previous = level
    return findings


def sections_of(blocks) -> list[dict]:
    """按 1/2 级标题切节，返回 [{level, title, line, body: [block]}]。

    3 级及以下标题**不进 body**：它们是结构，不是句子。早期把 h3 收进 body，
    导致"结果与分析"的首句被取成子标题、进而误报节首句。
    """
    sections = []
    current = None
    for block in blocks:
        if block["kind"] == "heading" and block.get("level") in (1, 2):
            if current is not None:
                sections.append(current)
            current = {
                "level": block["level"],
                "title": block["text"],
                "line": block["line"],
                "body": [],
            }
            continue
        if block["kind"] == "heading":
            continue  # h3+
        if current is not None and block["kind"] in ("paragraph", "list", "quote", "table"):
            current["body"].append(block)
    if current is not None:
        sections.append(current)
    return sections


def check_section_openers(blocks, enabled: bool, method_exempt: bool) -> list[dict]:
    """每节首句是否符合"为了……"句式。"""
    if not enabled:
        return []
    findings = []
    for section in sections_of(blocks):
        if not section["body"]:
            continue
        if method_exempt and _is_method_heading(section["title"]):
            continue
        first = section["body"][0]
        if first["kind"] != "paragraph":
            continue
        text = first["text"]
        if not OPENER_PATTERN.match(text):
            findings.append(
                {
                    "id": "section-opener",
                    "severity": "P1",
                    "line": first["line"],
                    "message": f"「{section['title']}」首句没有用「为了……」总起",
                    "snippet": text[:40],
                    "suggestion": "改为：为了 ……，需要 ……。说明这一节为何存在，不复述方法",
                    "autofix": None,
                }
            )
    return findings


def check_method_numbers(blocks) -> list[dict]:
    """方法章节不出现具体数值。"""
    findings = []
    for section in sections_of(blocks):
        if not _is_method_heading(section["title"]):
            continue
        for block in section["body"]:
            if block["kind"] in ("math", "table"):
                continue
            cleaned = strip_allowed_numbers(block["text"])
            numbers = NUMBER_PATTERN.findall(cleaned)
            if numbers:
                findings.append(
                    {
                        "id": "method-numbers",
                        "severity": "P1",
                        "line": block["line"],
                        "message": f"方法章节出现具体数值：{', '.join(numbers[:4])}",
                        "snippet": block["text"][:40],
                        "suggestion": "数值移到结果章节；方法章只保留公式常数、术语内数字、引用年份、表图号",
                        "autofix": None,
                    }
                )
                break
    return findings


def check_sections_present(blocks, wanted: list) -> list[dict]:
    """规格列出的章节是否都出现。"""
    if not wanted:
        return []
    titles = {re.sub(r"\s+", "", s["title"]) for s in sections_of(blocks)}
    findings = []
    for item in wanted:
        name = item if isinstance(item, str) else (item.get("title") if isinstance(item, dict) else None)
        if not isinstance(name, str) or name.strip() == "":
            continue
        if re.sub(r"\s+", "", name) not in titles:
            findings.append(
                {
                    "id": "missing-sections",
                    "severity": "P0",
                    "line": None,
                    "message": f"规格要求的章节缺失：「{name}」",
                    "snippet": "",
                    "suggestion": f"补一节 `## {name}`",
                    "autofix": None,
                }
            )
    return findings


def check_title(blocks, expected: str | None) -> list[dict]:
    """文档是否有标题。

    判据只看**第一个非空块**：标题必须是最开头的那一行。早期版本拼了前 3 块去
    匹配，导致"正文里恰好提到标题"也算通过，而且 snippet 显示混乱。
    """
    if not isinstance(expected, str) or expected.strip() == "":
        if not blocks:
            return [
                {
                    "id": "empty-document",
                    "severity": "P0",
                    "line": None,
                    "message": "文档没有内容",
                    "snippet": "",
                    "suggestion": "",
                    "autofix": None,
                }
            ]
        return []
    if not blocks:
        return [
            {
                "id": "empty-document",
                "severity": "P0",
                "line": None,
                "message": "文档没有内容",
                "snippet": "",
                "suggestion": "",
                "autofix": None,
            }
        ]
    first = blocks[0]
    if expected.strip() in first["text"]:
        return []
    return [
        {
            "id": "missing-title",
            "severity": "P0",
            "line": first["line"],
            "message": f"开头没有出现规格标题「{expected}」",
            "snippet": first["text"][:40],
            "suggestion": f"在正文最前面加一行 `# {expected}`",
            "autofix": None,
        }
    ]


# ── 统一入口 ──────────────────────────────────────────────────────────────


def run_all(
    blocks,
    *,
    preferences: dict,
    spec: dict,
    gate_cfg: dict,
) -> dict:
    """跑全部检查，返回 {p0, p1, verdict, counts}。"""
    from wm_common import banned_words, formatting_of  # 延迟导入，避免循环

    banned = banned_words(preferences)
    opener = preferences.get("sectionOpener") if isinstance(preferences.get("sectionOpener"), dict) else {}
    method_cfg = preferences.get("methodChapter") if isinstance(preferences.get("methodChapter"), dict) else {}

    findings: list[dict] = []
    findings.extend(check_title(blocks, spec.get("title") if isinstance(spec, dict) else None))
    findings.extend(check_fullwidth_brackets(blocks))
    findings.extend(check_banned_phrases(blocks, banned))
    findings.extend(check_heading_jumps(blocks))
    findings.extend(check_sections_present(blocks, spec.get("sections") if isinstance(spec, dict) else []))
    findings.extend(
        check_section_openers(
            blocks,
            enabled=opener.get("enabled") is True and gate_cfg.get("requireSectionOpeners") is not False,
            method_exempt=opener.get("methodChapterExempt") is True,
        )
    )
    if method_cfg.get("noConcreteNumbers") is True:
        findings.extend(check_method_numbers(blocks))

    p0 = [f for f in findings if f["severity"] == "P0"]
    p1 = [f for f in findings if f["severity"] == "P1"]
    return {
        "p0": p0,
        "p1": p1,
        "verdict": "PASS" if not p0 and not p1 else ("BLOCK" if p0 else "WARN"),
        "counts": {"p0": len(p0), "p1": len(p1), "total": len(findings)},
        "formatting": formatting_of(preferences),
    }
