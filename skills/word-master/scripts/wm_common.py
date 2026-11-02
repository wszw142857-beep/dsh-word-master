#!/usr/bin/env python3
"""Word Master — 共享工具层。

设计要点
--------
1. **规格即契约**。`spec.json` 冻结一次写作的全部可判定项（标题、章节、目标
   读者、语气、版式快照、闸门阈值）。构建与质检都读同一份，避免"写的时候按
   一套、查的时候按另一套"。

2. **偏好从偏好库读**。版式（字体/字号/边距/缩进）来自 `preferences.json`，
   脚本不写死。用户改偏好，下次构建自动生效。

3. **中文路径与编码**。Windows 控制台默认 GBK，所有读写显式 `encoding='utf-8'`，
   stdout 在入口处重设为 utf-8，避免 emoji / 中文在管道里炸掉。
"""

from __future__ import annotations

import copy
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

# 内置默认偏好。必须与 lib/preferences.js 的 DEFAULT_PREFERENCES 保持一致 ——
# 两处不一致会导致"注入给模型看的"和"质检门实际查的"不是同一套规则。
DEFAULT_PREFERENCES: dict[str, Any] = {
    "punctuation": {"halfWidthBracketsOnly": True},
    "bannedPhrases": {
        "defensive": ["不能", "不宜", "不应", "并非", "无法", "缺乏", "不代表", "不等于"],
        "delivery": ["待补", "待核", "待定", "待授权", "台账", "核对状态", "证据缺口", "占位", "工作稿", "候选版"],
        "meta": ["需要说明的是", "上述份额分别对应", "以便读者", "本文不再重复"],
        "process": ["经核实", "逐条记录", "纳入N条记录"],
    },
    "paragraph": {
        "oneTopicPerParagraph": True,
        "mergeAdjacentShortParagraphs": True,
        "progression": "给结果 → 说不确定性 → 给判断",
    },
    "sectionOpener": {
        "enabled": True,
        "pattern": "为了 + 本节要解决的问题,需要 + 相应动作",
        "methodChapterExempt": True,
    },
    "methodChapter": {
        "noConcreteNumbers": True,
        "keepOnly": ["公式常数", "术语内数字（HC5/EC50）", "引用年份", "表图号"],
        "allowSimulationSettings": True,
        "noTablePointers": True,
        "keepFormulaExplanations": True,
        "noOpenerSentence": True,
    },
    "resultsChapter": {
        "noForwardReference": True,
        "noRepeatFromMethod": True,
        "oneNumberPerExperiment": True,
    },
    "terminology": {
        "plainLanguage": True,
        "avoidLiteraryWords": ["时机", "历程", "特征", "相对"],
        "significanceNeedsTest": True,
        "explainFirstOccurrence": True,
    },
    "formatting": {
        "bodyFont": "宋体",
        "latinFont": "Times New Roman",
        "bodySizePt": 12,
        "lineSpacing": 1.5,
        "firstLineIndentChars": 2,
        "marginCm": {"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
        "headingStyleBuiltin": True,
    },
    "citation": {"style": "GB/T 7714", "numbering": "顺序编码制"},
    "autoInjection": {"enabled": True, "maxChars": 4000},
    "notes": [],
}

# 全角标点（括号为主，另含几个高频错用）
FULLWIDTH_BRACKETS = "（）"
FULLWIDTH_OTHER = "，。；：！？"

# 中文正文里允许出现的全角标点（正文用全角是正常的，只卡括号）
ALLOWED_FULLWIDTH = set("，。；：！？、“”‘’《》〈〉【】…—～·")


def force_utf8_stdio() -> None:
    """把 stdout/stderr 切成 utf-8，避免 Windows 控制台 GBK 编码错误。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def read_json(path: str | Path, default: Any = None) -> Any:
    """读一个 JSON 文件；不存在或损坏时返回 default 并打印到 stderr。"""
    target = Path(path)
    if not target.is_file():
        return default
    try:
        raw = target.read_text(encoding="utf-8-sig")
        return json.loads(raw)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"[WARN] 读取 {target} 失败：{exc}", file=sys.stderr)
        return default


def write_json(path: str | Path, payload: Any) -> None:
    """写 JSON（原子替换：临时文件 + rename）。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(tmp, target)


def emit(payload: Any) -> None:
    """给模型看的输出：单行 JSON。"""
    print(json.dumps(payload, ensure_ascii=False))


def dsh_home() -> Path:
    """解析 DSH 主目录。"""
    env = os.environ.get("DSH_HOME", "").strip()
    return Path(env) if env else Path.home() / ".dsh"


def preferences_path() -> Path:
    """偏好库绝对路径。"""
    return dsh_home() / "word-master" / "preferences.json"


def deep_merge(target: dict, patch: dict) -> dict:
    """深合并；`None` 表示显式关闭，保留 None。"""
    for key, value in patch.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            deep_merge(target[key], value)
        else:
            target[key] = value
    return target


def load_preferences() -> dict[str, Any]:
    """读偏好库，并**合并内置默认值**。

    关键：文件不存在时必须回落到默认值，不能返回空 dict。否则质检门的禁用词
    检查会因为词表为空而静默失效 —— 看起来"全部通过"，实际一条都没查。
    这与 `lib/preferences.js` 的行为必须一致。
    """
    defaults = {k: __import__("copy").deepcopy(v) for k, v in DEFAULT_PREFERENCES.items()}
    payload = read_json(preferences_path(), default=None)
    if isinstance(payload, dict) and isinstance(payload.get("values"), dict):
        return deep_merge(defaults, payload["values"])
    if isinstance(payload, dict) and payload:
        return deep_merge(defaults, payload)
    return defaults


def formatting_of(prefs: dict[str, Any]) -> dict[str, Any]:
    """从偏好里取版式，并给出安全默认值。"""
    fmt = prefs.get("formatting") if isinstance(prefs.get("formatting"), dict) else {}
    margins = fmt.get("marginCm") if isinstance(fmt.get("marginCm"), dict) else {}
    return {
        "bodyFont": fmt.get("bodyFont") or "宋体",
        "latinFont": fmt.get("latinFont") or "Times New Roman",
        "bodySizePt": fmt.get("bodySizePt") if isinstance(fmt.get("bodySizePt"), (int, float)) else 12,
        "lineSpacing": fmt.get("lineSpacing") if isinstance(fmt.get("lineSpacing"), (int, float)) else 1.5,
        "firstLineIndentChars": fmt.get("firstLineIndentChars")
        if isinstance(fmt.get("firstLineIndentChars"), (int, float))
        else 2,
        "marginCm": {
            "top": margins.get("top", 2.54),
            "bottom": margins.get("bottom", 2.54),
            "left": margins.get("left", 3.17),
            "right": margins.get("right", 3.17),
        },
    }


def banned_words(prefs: dict[str, Any]) -> dict[str, list[str]]:
    """取出各类禁用词表。"""
    banned = prefs.get("bannedPhrases") if isinstance(prefs.get("bannedPhrases"), dict) else {}
    out: dict[str, list[str]] = {}
    for key in ("defensive", "delivery", "meta", "process"):
        words = banned.get(key)
        if isinstance(words, list):
            out[key] = [w for w in words if isinstance(w, str) and w.strip() != ""]
    return out


def iter_paragraph_texts(document) -> list[tuple[int, str]]:
    """遍历 docx 的段落，返回 (段落序号, 文本)；跳过空段。"""
    rows: list[tuple[int, str]] = []
    for index, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()
        if text != "":
            rows.append((index, text))
    return rows


def heading_level(paragraph) -> int | None:
    """从样式名推断标题级别；不是标题返回 None。"""
    name = (paragraph.style.name or "").strip()
    match = re.match(r"^(?:Heading|标题)\s*(\d+)$", name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    if name.lower() in {"title", "标题"}:
        return 0
    return None


def count_fullwidth_brackets(text: str) -> int:
    """统计全角括号数量。"""
    return sum(text.count(ch) for ch in FULLWIDTH_BRACKETS)


def normalize(text: str) -> str:
    """压掉空白，便于短语匹配（避免跨行匹配失败）。"""
    return re.sub(r"\s+", "", text)


def find_all(haystack: str, needle: str) -> list[int]:
    """返回 needle 在 haystack 里全部出现位置。"""
    if needle == "":
        return []
    hits: list[int] = []
    start = 0
    while True:
        index = haystack.find(needle, start)
        if index < 0:
            return hits
        hits.append(index)
        start = index + len(needle)


def relpath(path: Path, base: Path) -> str:
    """尽量给相对路径，便于报告阅读。"""
    try:
        return str(path.resolve().relative_to(base.resolve()))
    except (ValueError, OSError):
        return str(path)
