#!/usr/bin/env python3
"""Word Master — 视觉判据（rubric）检查。

与 ppt-master `visual-review.md` 的关系
--------------------------------------
它那份 rubric 是给 SVG 用的（bbox 越界、元素重叠、对比度），判据都建立在坐标
几何上。Word 没有 bbox 可算，但**检查类别可以映射**：

| 它（SVG）        | 我们（Word 渲染后的 PDF）            |
|------------------|-------------------------------------|
| H1 元素越界      | 文字超出页边距                       |
| H2 文字溢出容器  | 行内文字越过右边距                   |
| H3 文字重叠      | 同一位置多段重叠（表格/文本框常见）   |
| H7 锚定元素错位  | 页眉页脚缺失或错位                   |
| S1 行距过紧      | 行距小于字号 1.05 倍                 |
| S2 行距过松      | 段内出现大块空白                     |
| S4 对齐漂移      | 同列左边界不一致                     |

判据从**渲染后的 PDF**上量，不是从 docx XML 上算 —— 因为只有渲染后才知道真实
排版结果。字体回退、分页位置这些，XML 里根本看不出来。

与 gate.py 的分工
-----------------
- `gate.py`：查**内容与规则**（括号、禁用词、章节、节首句），能自动修。
- 本模块：查**排版**（溢出、行距、重叠、孤行），只报告不自动修 —— 排版修正
  需要判断意图，机器代劳容易越改越糟。

退出码
------
- `0`：无硬性问题
- `3`：有 P0（排版破损，必须修）
- `4`：只有 P1（观感问题，建议修）

用法
----
    python visual_check.py --project DIR [--pdf render/output.pdf] [--json]
"""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, read_json, write_json  # noqa: E402

# 判据阈值（集中在这里，便于按需调）
THRESHOLDS = {
    "margin_tolerance_pt": 2.0,      # 超出页边距多少算越界
    "min_line_spacing_ratio": 1.05,  # 行距 / 字号 下限
    "max_block_gap_ratio": 2.6,      # 段内空白 / 行高 上限（超过算空洞）
    "alignment_tolerance_pt": 4.0,   # 同列左边界容差
    "orphan_heading_bottom_ratio": 0.14,  # 标题落在页面底部多少比例内算孤行
}


def _load_pdf(pdf: Path):
    """打开 PDF，返回 pymupdf 文档对象或错误。"""
    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz  # 旧名
        except ImportError:
            return None, "缺少 PyMuPDF（pip install pymupdf）"
    try:
        return fitz.open(str(pdf)), None
    except Exception as exc:  # noqa: BLE001
        return None, f"PDF 打不开：{exc}"


def _blocks_of(page) -> list[dict]:
    """把一页拆成文本块，带 bbox 与字号。"""
    out = []
    for block in page.get_text("dict")["blocks"]:
        if block.get("type") != 0:
            continue
        lines = []
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            text = "".join(s.get("text", "") for s in spans)
            if text.strip() == "":
                continue
            sizes = [s.get("size", 0) for s in spans if s.get("size")]
            lines.append(
                {
                    "text": text,
                    "bbox": line.get("bbox"),
                    "size": max(sizes) if sizes else 0,
                    "color": spans[0].get("color", 0),
                    "font": spans[0].get("font", ""),
                }
            )
        if lines:
            out.append({"bbox": block.get("bbox"), "lines": lines})
    return out


def check_page_bounds(page, page_index: int, tolerance: float) -> list[dict]:
    """H1/H2：文字是否越出页边距。"""
    findings = []
    width = page.rect.width
    height = page.rect.height
    # PDF 页里内容区的右/下边界；留一点容差给标点悬挂
    for block in _blocks_of(page):
        for line in block["lines"]:
            bbox = line["bbox"]
            if bbox is None:
                continue
            x0, y0, x1, y1 = bbox
            if x1 > width - tolerance or x0 < tolerance - 40:
                findings.append(
                    {
                        "id": "text-out-of-margin",
                        "severity": "P0",
                        "page": page_index + 1,
                        "message": "文字越过页面边界",
                        "snippet": line["text"][:40],
                        "detail": f"x0={x0:.1f} x1={x1:.1f} 页宽={width:.1f}",
                    }
                )
                break
            if y1 > height + tolerance or y0 < -tolerance:
                findings.append(
                    {
                        "id": "text-out-of-page",
                        "severity": "P0",
                        "page": page_index + 1,
                        "message": "文字超出页面上下边界",
                        "snippet": line["text"][:40],
                        "detail": f"y0={y0:.1f} y1={y1:.1f} 页高={height:.1f}",
                    }
                )
                break
        if len(findings) >= 6:
            break
    return findings


def check_line_spacing(page, page_index: int, min_ratio: float) -> list[dict]:
    """S1：行距是否过紧。

    判据用**基线间距 / 字号**，而不是「行间空白 / 字号」。

    为什么：PDF 里一行的 bbox 高度已经包含字体的上下缘（实测 12pt 字 ≈ 15.8pt
    行盒），相邻两行的 bbox 间距（gap）在正常单倍行距下也只有 1~2pt。用 gap 做
    分母会得到"永远不触发"的阈值 —— 这正是第一版的 bug。

    正确的量是**基线到基线**（baseline pitch）：相邻行 bbox 顶部的差。单倍行距约
    1.2×字号，1.5 倍行距约 1.8×。小于 1.05× 才算过紧。
    """
    findings = []
    for block in _blocks_of(page):
        lines = block["lines"]
        if len(lines) < 2:
            continue
        for prev, cur in zip(lines, lines[1:]):
            if prev["bbox"] is None or cur["bbox"] is None:
                continue
            pitch = cur["bbox"][1] - prev["bbox"][1]
            size = max(prev["size"], cur["size"])
            if size <= 0 or pitch <= 0:
                continue
            ratio = pitch / size
            if ratio < min_ratio:
                findings.append(
                    {
                        "id": "line-spacing-tight",
                        "severity": "P1",
                        "page": page_index + 1,
                        "message": f"行距过紧（基线间距 {pitch:.1f}pt / 字号 {size:.1f}pt = {ratio:.2f}×）",
                        "snippet": (prev["text"] + cur["text"])[:40],
                        "detail": f"pitch={pitch:.1f} size={size:.1f} ratio={ratio:.2f}（下限 {min_ratio}）",
                    }
                )
                break
        if len(findings) >= 4:
            break
    return findings


def check_overlap(page, page_index: int) -> list[dict]:
    """H3：两个文本块 bbox 相交（排除同块内行）。"""
    findings = []
    blocks = [b for b in _blocks_of(page) if b["bbox"]]
    for i, a in enumerate(blocks):
        for b in blocks[i + 1 :]:
            ax0, ay0, ax1, ay1 = a["bbox"]
            bx0, by0, bx1, by1 = b["bbox"]
            ox = min(ax1, bx1) - max(ax0, bx0)
            oy = min(ay1, by1) - max(ay0, by0)
            if ox > 6 and oy > 6:
                findings.append(
                    {
                        "id": "text-overlap",
                        "severity": "P0",
                        "page": page_index + 1,
                        "message": "两段文字在渲染后重叠",
                        "snippet": (a["lines"][0]["text"][:20] + " ‖ " + b["lines"][0]["text"][:20]),
                        "detail": f"重叠 {ox:.1f}×{oy:.1f}pt",
                    }
                )
                return findings
    return findings


def check_orphan_heading(page, page_index: int, bottom_ratio: float) -> list[dict]:
    """标题孤行：标题落在页面底部、正文跑到下一页。"""
    height = page.rect.height
    blocks = _blocks_of(page)
    if not blocks:
        return []
    last = blocks[-1]
    if last["bbox"] is None:
        return []
    last_line = last["lines"][-1]
    # 粗体或明显大于正文 = 疑似标题
    is_bold = "Bold" in (last_line.get("font") or "") or "Heavy" in (last_line.get("font") or "")
    body_size = _body_size(blocks)
    is_big = last_line["size"] > body_size * 1.12
    if (is_bold or is_big) and last["bbox"][3] > height * (1 - bottom_ratio):
        return [
            {
                "id": "orphan-heading",
                "severity": "P1",
                "page": page_index + 1,
                "message": "标题落在页面底部，正文被推到下页",
                "snippet": last_line["text"][:40],
                "detail": f"y1={last['bbox'][3]:.1f} 页高={height:.1f}",
            }
        ]
    return []


def _body_size(blocks: list[dict]) -> float:
    """估算正文字号：出现字符最多的那个字号。"""
    tally: dict[float, int] = {}
    for block in blocks:
        for line in block["lines"]:
            size = round(line["size"], 1)
            tally[size] = tally.get(size, 0) + len(line["text"])
    return max(tally, key=tally.get) if tally else 10.0


def check_font_fallback(doc, tolerance_pages: int = 3) -> list[dict]:
    """字体回退检测：同一逻辑角色的字号在页间剧烈跳变。"""
    per_page: list[float] = []
    for page in doc:
        blocks = _blocks_of(page)
        if blocks:
            per_page.append(_body_size(blocks))
    if len(per_page) < 2:
        return []
    lo, hi = min(per_page), max(per_page)
    if hi - lo > 2.5:
        return [
            {
                "id": "font-size-drift",
                "severity": "P1",
                "page": None,
                "message": f"各页正文字号不一致（{lo:.1f}–{hi:.1f}pt）",
                "snippet": " ".join(f"{s:.1f}" for s in per_page[:8]),
                "detail": "可能是字体回退或样式未统一",
            }
        ]
    return []


def run_checks(project: Path, pdf_name: str) -> dict:
    """跑全部视觉判据。"""
    pdf_path = project / pdf_name
    if not pdf_path.is_file():
        # 回落到 render/ 下找
        candidates = sorted(project.glob("render/*.pdf"))
        if candidates:
            pdf_path = candidates[0]
        else:
            return {
                "ok": False,
                "reason": "missing-pdf",
                "message": f"没有 {pdf_name}，先跑 word_render 生成 PDF，视觉判据只能从渲染结果上量。",
            }

    doc, error = _load_pdf(pdf_path)
    if doc is None:
        return {"ok": False, "reason": "pdf-unreadable", "message": error}

    findings: list[dict] = []
    for index, page in enumerate(doc):
        findings.extend(check_page_bounds(page, index, THRESHOLDS["margin_tolerance_pt"]))
        findings.extend(check_overlap(page, index))
        findings.extend(check_line_spacing(page, index, THRESHOLDS["min_line_spacing_ratio"]))
        findings.extend(check_orphan_heading(page, index, THRESHOLDS["orphan_heading_bottom_ratio"]))

    findings.extend(check_font_fallback(doc))
    page_count = doc.page_count
    doc.close()

    p0 = [f for f in findings if f["severity"] == "P0"]
    p1 = [f for f in findings if f["severity"] == "P1"]

    report = {
        "ok": not p0,
        "checkedAt": datetime.now(timezone.utc).isoformat(),
        "pdf": str(pdf_path),
        "pages": page_count,
        "p0": p0,
        "p1": p1,
        "counts": {"p0": len(p0), "p1": len(p1)},
        "verdict": "PASS" if not p0 and not p1 else ("BLOCK" if p0 else "WARN"),
        "thresholds": THRESHOLDS,
    }
    write_json(project / "visual-report.json", report)
    return report


def format_report(report: dict) -> str:
    """渲染成给模型读的文本。"""
    if report.get("verdict") is None:
        return f"视觉判据无法运行：{report.get('message') or report.get('reason')}"

    lines = [
        f"# 视觉判据：{report['verdict']}",
        "",
        f"- 对象：{report['pdf']}（{report['pages']} 页）",
    ]
    for key, label in (("p0", "P0 排版破损"), ("p1", "P1 观感问题")):
        items = report.get(key) or []
        if not items:
            lines.append(f"- {label}：无")
            continue
        lines.append(f"- {label}：{len(items)} 条")
        for item in items[:10]:
            where = f"第{item['page']}页 " if item.get("page") else ""
            lines.append(f"  - {where}{item['message']}")
            if item.get("snippet"):
                lines.append(f"    - {item['snippet']}")
            if item.get("detail"):
                lines.append(f"    - {item['detail']}")
        if len(items) > 10:
            lines.append(f"  - …… 还有 {len(items) - 10} 条，见 visual-report.json")

    lines.append("")
    if report["verdict"] == "BLOCK":
        lines.append("结论：**排版有破损**，改版式或内容后重新构建并重跑。")
    elif report["verdict"] == "WARN":
        lines.append("结论：可以交付，观感问题建议处理。")
    else:
        lines.append("结论：**排版干净**。")
    lines.append("")
    lines.append("⚠️ 机器判据只覆盖可量化的部分。**仍须自己逐页看图** —— 字体回退、")
    lines.append("   表格观感、图注对齐这些，量不出来但一眼能看出。")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 视觉判据")
    parser.add_argument("--project", required=True)
    parser.add_argument("--pdf", default="render/output.pdf")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    report = run_checks(Path(args.project).expanduser(), args.pdf)
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
