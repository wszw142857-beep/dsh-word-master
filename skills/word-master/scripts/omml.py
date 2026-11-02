#!/usr/bin/env python3
"""LaTeX → OMML（Word 原生公式）转换器。

为什么要做这一步
----------------
Markdown 里的 `$$...$$` 如果原样塞进 DOCX，Word 会把它当普通文字显示 —— 渲染
闸门里能直接看到 `$$HQ = \\frac{CDI}{RfD}$$` 这种字样。用户要的是**可继续编辑的
Word 公式**，不是一段 LaTeX 源码。

方案选择
--------
- 装 MathType / 调 pandoc：多一个外部依赖，交付出去还得让用户也装。
- 直接写 OMML：Word 原生，零依赖，打开即可双击编辑。

OMML 是 OOXML 的公式命名空间（`m:` 前缀），这里只覆盖中文技术文档里高频出现的
子集：分式、根式、上下标、求和/积分、希腊字母、常用运算符、文本块。

刻意不支持：矩阵、多行对齐、化学式。遇到不认识的构造时**降级为普通文本**并在
stderr 提示，而不是生成坏结构 —— 一份能看的降级文本比一个打不开的公式好。

用法
----
    omml = latex_to_omml(r"\\frac{a}{b}")
    element = omml_element(omml)   # 可直接 append 进段落
"""

from __future__ import annotations

import re
import sys

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 希腊字母与常用符号
SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ",
    "sigma": "σ", "tau": "τ", "upsilon": "υ", "phi": "φ", "chi": "χ",
    "psi": "ψ", "omega": "ω",
    "Alpha": "Α", "Beta": "Β", "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ",
    "Lambda": "Λ", "Pi": "Π", "Sigma": "Σ", "Phi": "Φ", "Psi": "Ψ",
    "Omega": "Ω",
    "times": "×", "div": "÷", "pm": "±", "mp": "∓", "leq": "≤", "geq": "≥",
    "neq": "≠", "approx": "≈", "equiv": "≡", "propto": "∝", "cdot": "·",
    "to": "→", "rightarrow": "→", "leftarrow": "←", "infty": "∞",
    "partial": "∂", "nabla": "∇", "sum": "∑", "prod": "∏", "int": "∫",
    "sqrt": "√", "angle": "∠", "degree": "°", "circ": "∘",
}

# 需要转义的 XML 字符
XML_ESCAPE = {"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;"}


def _esc(text: str) -> str:
    """XML 转义。"""
    return "".join(XML_ESCAPE.get(ch, ch) for ch in text)


def _run(text: str, *, plain: bool = False) -> str:
    """生成一个 OMML run。`plain` 用于函数名/文本块（正体）。"""
    style = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else ""
    return (
        "<m:r>"
        f"{style}"
        f'<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>'
        f"<m:t>{_esc(text)}</m:t>"
        "</m:r>"
    )


def _frac(numerator: str, denominator: str) -> str:
    """分式。"""
    return (
        "<m:f>"
        f"<m:num>{numerator}</m:num>"
        f"<m:den>{denominator}</m:den>"
        "</m:f>"
    )


def _sqrt(body: str) -> str:
    """平方根。"""
    return (
        "<m:rad>"
        '<m:radPr><m:degHide m:val="1"/></m:radPr>'
        "<m:deg/>"
        f"<m:e>{body}</m:e>"
        "</m:rad>"
    )


def _sup(base: str, sup: str) -> str:
    """上标。"""
    return f"<m:sSup><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>"


def _sub(base: str, sub: str) -> str:
    """下标。"""
    return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"


def _subsup(base: str, sub: str, sup: str) -> str:
    """同时上下标。"""
    return f"<m:sSubSup><m:e>{base}</m:e><m:sub>{sub}</m:sub><m:sup>{sup}</m:sup></m:sSubSup>"


def _delim(body: str) -> str:
    """圆括号包裹（自动伸缩）。"""
    return (
        '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr>'
        f"<m:e>{body}</m:e></m:d>"
    )


def _nary(char: str, lower: str, upper: str, body: str) -> str:
    """求和 / 积分等大型运算符。"""
    props = '<m:naryPr><m:chr m:val="' + char + '"/>'
    props += '<m:limLoc m:val="undOvr"/>'
    if lower == "":
        props += '<m:subHide m:val="1"/>'
    if upper == "":
        props += '<m:supHide m:val="1"/>'
    props += "</m:naryPr>"
    return f"<m:nary>{props}<m:sub>{lower}</m:sub><m:sup>{upper}</m:sup><m:e>{body}</m:e></m:nary>"


def _read_group(text: str, start: int) -> tuple[str, int]:
    """从 `start` 处的 `{` 读取配平的组内容，返回 (内容, 结束位置)。"""
    if start >= len(text) or text[start] != "{":
        # 无花括号：只取一个字符（LaTeX 允许 \\frac12 这种写法）
        if start < len(text):
            return text[start], start + 1
        return "", start
    depth = 0
    for index in range(start, len(text)):
        if text[index] == "{":
            depth += 1
        elif text[index] == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1:index], index + 1
    return text[start + 1:], len(text)


def _parse(text: str) -> str:
    """递归下降解析 LaTeX 子集，产出 OMML 片段。"""
    out: list[str] = []
    i = 0
    while i < len(text):
        ch = text[i]

        if ch == "\\":
            match = re.match(r"\\([a-zA-Z]+)", text[i:])
            if match is None:
                # \( \) \[ \] 等定界符，直接跳过
                i += 2
                continue
            name = match.group(1)
            i += len(match.group(0))

            if name == "frac":
                num, i = _read_group(text, i)
                den, i = _read_group(text, i)
                out.append(_frac(_parse(num), _parse(den)))
                continue

            if name == "sqrt":
                body, i = _read_group(text, i)
                out.append(_sqrt(_parse(body)))
                continue

            if name in ("sum", "prod", "int", "iint"):
                char = SYMBOLS.get(name, "∑")
                lower, upper = "", ""
                # 读紧随其后的 _{...} ^{...}
                while i < len(text) and text[i] in "_^":
                    marker = text[i]
                    i += 1
                    group, i = _read_group(text, i)
                    if marker == "_":
                        lower = _parse(group)
                    else:
                        upper = _parse(group)
                body, i = _read_group(text, i)
                out.append(_nary(char, lower, upper, _parse(body)))
                continue

            if name in ("text", "mathrm", "mathbf", "operatorname"):
                body, i = _read_group(text, i)
                out.append(_run(body, plain=True))
                continue

            if name in ("left", "right"):
                continue

            if name in SYMBOLS:
                out.append(_run(SYMBOLS[name]))
                continue

            # 未知命令：降级为文本，stderr 提示一次
            print(f"[WARN] 未支持的 LaTeX 命令 \\{name}，按文本处理", file=sys.stderr)
            out.append(_run("\\" + name))
            continue

        if ch in "_^":
            # 附着在前一个 run 上；简化处理：单独成上/下标
            marker = ch
            i += 1
            group, i = _read_group(text, i)
            base = out.pop() if out else _run("")
            if marker == "_":
                out.append(_sub(base, _parse(group)))
            else:
                out.append(_sup(base, _parse(group)))
            continue

        if ch == "{":
            body, i = _read_group(text, i)
            out.append(_parse(body))
            continue

        if ch in "()":
            # 成对圆括号 → 伸缩定界符
            if ch == "(":
                depth = 0
                j = i
                while j < len(text):
                    if text[j] == "(":
                        depth += 1
                    elif text[j] == ")":
                        depth -= 1
                        if depth == 0:
                            break
                    j += 1
                if j < len(text):
                    inner = text[i + 1:j]
                    out.append(_delim(_parse(inner)))
                    i = j + 1
                    continue
            out.append(_run(ch))
            i += 1
            continue

        # 普通字符：合并连续的非特殊字符成一个 run
        match = re.match(r"[^\\_^{}()]+", text[i:])
        if match is not None:
            chunk = match.group(0)
            out.append(_run(chunk))
            i += len(chunk)
            continue

        out.append(_run(ch))
        i += 1

    return "".join(out)


def latex_to_omml(latex: str) -> str:
    """把一段 LaTeX 转成 `<m:oMath>` 元素字符串。"""
    body = _parse(latex.strip())
    return f'<m:oMath xmlns:m="{M_NS}" xmlns:w="{W_NS}">{body}</m:oMath>'


def omml_element(latex: str):
    """把 LaTeX 转成可直接 append 到段落的 lxml 元素。"""
    from docx.oxml import parse_xml  # 延迟导入，便于纯逻辑测试

    return parse_xml(latex_to_omml(latex))


INLINE_MATH = re.compile(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)")
DISPLAY_MATH = re.compile(r"\$\$(.+?)\$\$", re.DOTALL)


def is_display_math(line: str) -> bool:
    """判断整行是否是一个块级公式。"""
    stripped = line.strip()
    return stripped.startswith("$$") and stripped.endswith("$$") and len(stripped) > 4


def extract_display(line: str) -> str | None:
    """取出块级公式内容；不是块级公式返回 None。"""
    if not is_display_math(line):
        return None
    return line.strip()[2:-2].strip()
