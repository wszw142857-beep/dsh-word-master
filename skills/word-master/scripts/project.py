#!/usr/bin/env python3
"""Word Master — 项目生命周期：init / import-sources / validate / info。

为什么需要这一层
----------------
写作不是一次线性冲刺：素材会陆续进来、稿子要跨会话接着改、交付前要确认没有
漏步骤。散落在目录里的文件（spec.json / draft.md / review-notes.json / 各种
报告）没有统一的"项目状态"概念时，模型每次都要重新勘察一遍，人也说不清
"这个项目到底做到哪一步了"。

本模块给每个项目一个**明确的状态视图**：

    init            建目录骨架 + 项目清单
    import-sources  把素材收进 sources/ 并自动转 Markdown
    validate        按阶段判据检查"该有的东西有没有、该做的步骤做了没"
    info            一行看清项目进度

目录骨架
--------
    <project>/
    ├── project.json        项目清单（phase / 创建时间 / 素材登记）
    ├── spec.json           规格（word_spec 管）
    ├── draft.md            正文（模型写，审阅界面可改）
    ├── output.docx         构建产物
    ├── final.docx          交付定稿
    ├── sources/            原始素材（PDF/DOCX/...）
    ├── converted/          素材转换出的 Markdown
    ├── gate-report.json    质检报告
    ├── visual-report.json  视觉判据报告
    ├── review-notes.json   审阅界面收集的决定
    ├── review-result.json  交接记录
    └── render/             渲染产物（PDF + 逐页 PNG）

阶段（phase）
------------
`draft` → `frozen` → `built` → `reviewed` → `delivered`

不是强制的状态机（用户可以跳着做），而是**推断出来的进度快照**：
validate 会告诉你"现在处于哪个阶段、下一步该做什么"。

用法
----
    python project.py init <dir> --title "标题"
    python project.py import-sources <dir> a.pdf b.docx https://...
    python project.py validate <dir>
    python project.py info <dir>
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import emit, force_utf8_stdio, read_json, write_json  # noqa: E402

MANIFEST_NAME = "project.json"
SOURCES_DIR = "sources"
CONVERTED_DIR = "converted"

# 可转换的素材类型（与 convert.py 的 KIND_BY_EXT 保持一致）
CONVERTIBLE = {".pdf", ".docx", ".xlsx", ".xlsm", ".pptx", ".html", ".htm"}
# 直接当文本用的
TEXT_SUFFIXES = {".md", ".markdown", ".txt"}


# ── 清单 ──────────────────────────────────────────────────────────────────


def manifest_path(project: Path) -> Path:
    """项目清单路径。"""
    return project / MANIFEST_NAME


def load_manifest(project: Path) -> dict:
    """读清单；不存在返回空结构。"""
    data = read_json(manifest_path(project), default=None)
    if isinstance(data, dict):
        return data
    return {
        "schemaVersion": 1,
        "createdAt": None,
        "title": None,
        "sources": [],
    }


def save_manifest(project: Path, data: dict) -> Path:
    """写清单。"""
    data["updatedAt"] = datetime.now(timezone.utc).isoformat()
    target = manifest_path(project)
    write_json(target, data)
    return target


# ── init ──────────────────────────────────────────────────────────────────


def cmd_init(project: Path, title: str | None, force: bool) -> dict:
    """建项目骨架。

    刻意**不覆盖**已有的 spec.json / draft.md —— init 在已有项目上重复执行时
    应该是幂等的（补齐缺的目录），而不是把用户的稿子清空。要重建得显式 --force。
    """
    existed = project.is_dir() and any(project.iterdir())
    if existed and not force:
        # 幂等模式：只补缺的目录
        created = []
        for sub in (SOURCES_DIR, CONVERTED_DIR):
            target = project / sub
            if not target.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                created.append(sub)
        manifest = load_manifest(project)
        if title and not manifest.get("title"):
            manifest["title"] = title
        save_manifest(project, manifest)
        return {
            "ok": True,
            "mode": "idempotent",
            "path": str(project),
            "createdDirs": created,
            "manifest": str(manifest_path(project)),
            "note": "项目已存在，只补齐缺失的目录；已有的 spec/draft 未被改动",
        }

    project.mkdir(parents=True, exist_ok=True)
    for sub in (SOURCES_DIR, CONVERTED_DIR):
        (project / sub).mkdir(parents=True, exist_ok=True)

    manifest = load_manifest(project)
    manifest["schemaVersion"] = 1
    manifest["createdAt"] = datetime.now(timezone.utc).isoformat()
    manifest["title"] = title or project.name
    manifest.setdefault("sources", [])
    save_manifest(project, manifest)

    return {
        "ok": True,
        "mode": "created",
        "path": str(project),
        "manifest": str(manifest_path(project)),
        "next": "用 import-sources 收素材，或直接 word_spec init 建规格",
    }


# ── import-sources ────────────────────────────────────────────────────────


def _is_url(value: str) -> bool:
    """是否网页 URL。"""
    return bool(re.match(r"^https?://", value, re.IGNORECASE))


def _run_convert(kind: str, source: str, out_dir: Path, stem: str) -> dict:
    """调用 convert.py，返回转换结果。

    用子进程而不是 import：convert.py 依赖一堆可选库（pymupdf/mammoth/
    openpyxl/python-pptx/bs4），import 进来会让 project.py 也背上这些依赖。
    子进程失败只是"这次转换不成功"，不影响 import-sources 本身。
    """
    script = Path(__file__).resolve().parent / "convert.py"
    if not script.is_file():
        return {"ok": False, "reason": "missing-converter", "message": "找不到 convert.py"}
    try:
        completed = subprocess.run(
            [sys.executable, str(script), "--kind", kind, "--input", source, "--out-dir", str(out_dir), "--stem", stem],
            capture_output=True,
            timeout=300,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1", "PYTHONIOENCODING": "utf-8"},
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"ok": False, "reason": "convert-failed", "message": str(exc)[:200]}

    text = (completed.stdout or b"").decode("utf-8", errors="replace").strip()
    if completed.returncode != 0:
        return {"ok": False, "reason": "convert-failed", "message": text[:300] or f"退出码 {completed.returncode}"}
    try:
        return json.loads(text.splitlines()[-1])
    except (json.JSONDecodeError, IndexError):
        return {"ok": False, "reason": "convert-bad-output", "message": text[:200]}


def _promote_to_draft(converted_path: Path, project: Path) -> str | None:
    """把转换结果提升为项目的 draft.md。

    只在**项目还没有 draft.md** 时做，且只做一次（第一份文档类素材）。
    理由：用户丢一份成品文档过来，最想做的第一件事就是"打开它改"，而不是先
    自己把 md 挪成 draft.md。但这不能覆盖已有稿子——那等于把用户的工作抹掉。
    """
    draft = project / "draft.md"
    if draft.is_file():
        return None
    try:
        shutil.copy2(converted_path, draft)
    except OSError:
        return None
    return str(draft)


def cmd_import(project: Path, sources: list[str], move: bool, convert: bool = True) -> dict:
    """把素材收进项目，可选地就地转换。

    本地文件默认**复制**（不 move）—— 用户的原始文件不该因为我们收素材就消失。
    只有在显式 --move 时才移动。

    `convert=True` 时，文档类素材会立刻转成 Markdown 放进 `converted/`，
    并且**如果项目还没有 draft.md**，第一份转换结果会被提升为 draft.md，
    这样"丢一份成品文档进来 → 直接开审阅台改"就是一条命令的事。
    """
    if not project.is_dir():
        return {"ok": False, "reason": "missing-project", "message": f"{project} 不存在，先 init"}

    sources_dir = project / SOURCES_DIR
    converted_dir = project / CONVERTED_DIR
    sources_dir.mkdir(parents=True, exist_ok=True)
    converted_dir.mkdir(parents=True, exist_ok=True)

    manifest = load_manifest(project)
    registered = manifest.setdefault("sources", [])

    imported: list[dict] = []
    skipped: list[dict] = []
    conversions: list[dict] = []
    promoted: str | None = None

    for raw in sources:
        if _is_url(raw):
            stem = f"web_{len(conversions) + 1}"
            entry = {"kind": "url", "source": raw, "dest": None, "convert": "web"}
            imported.append(entry)
            if convert:
                result = _run_convert("web", raw, converted_dir, stem)
                conversions.append({"source": raw, "ok": result.get("ok"), "markdown": result.get("markdown"),
                                    "error": result.get("message")})
                if result.get("ok") and promoted is None:
                    promoted = _promote_to_draft(Path(result["markdown"]), project)
            continue

        src = Path(raw).expanduser()
        if not src.is_file():
            skipped.append({"source": raw, "reason": "not-found"})
            continue

        suffix = src.suffix.lower()
        if suffix in TEXT_SUFFIXES:
            # 已经是 Markdown/文本，直接放进 converted/
            dest = converted_dir / src.name
            if dest.resolve() != src.resolve():
                shutil.copy2(src, dest)
            imported.append({"kind": "text", "source": str(src), "dest": str(dest), "convert": None})
            if promoted is None:
                promoted = _promote_to_draft(dest, project)
            continue

        if suffix in CONVERTIBLE:
            dest = sources_dir / src.name
            if dest.exists() and dest.resolve() == src.resolve():
                pass  # 已经在 sources 里了
            elif move:
                shutil.move(str(src), str(dest))
            else:
                shutil.copy2(src, dest)
            imported.append({"kind": "document", "source": str(src), "dest": str(dest), "convert": suffix.lstrip(".")})

            if convert:
                stem = re.sub(r"[^\w\u4e00-\u9fff-]+", "_", dest.stem).strip("_") or "doc"
                result = _run_convert(dest.suffix.lstrip("."), str(dest), converted_dir, stem)
                conversions.append({"source": str(dest), "ok": result.get("ok"),
                                    "markdown": result.get("markdown"), "error": result.get("message")})
                if result.get("ok") and promoted is None:
                    promoted = _promote_to_draft(Path(result["markdown"]), project)
            continue

        skipped.append({"source": raw, "reason": f"unsupported-suffix:{suffix}"})

    # 登记到清单（去重）
    known = {(entry.get("source"), entry.get("dest")) for entry in registered}
    for entry in imported:
        key = (entry.get("source"), entry.get("dest"))
        if key not in known:
            registered.append({**entry, "importedAt": datetime.now(timezone.utc).isoformat()})
            known.add(key)
    save_manifest(project, manifest)

    next_step = None
    if promoted:
        next_step = "直接开审阅台：word_review action=open"
    elif any(not c.get("ok") for c in conversions):
        next_step = "部分素材转换失败，见 conversions[].error；可用 word_convert 单独重试"
    elif imported:
        next_step = "写 draft.md，或 word_spec init 建规格"

    return {
        "ok": True,
        "imported": imported,
        "skipped": skipped,
        "conversions": conversions,
        "promotedToDraft": promoted,
        "draftCreated": promoted is not None,
        "registered": len(registered),
        "next": next_step,
    }


# ── validate ──────────────────────────────────────────────────────────────


def detect_phase(project: Path) -> str:
    """推断当前阶段。"""
    if (project / "final.docx").is_file():
        return "delivered"
    if (project / "review-result.json").is_file():
        return "reviewed"
    if (project / "output.docx").is_file():
        return "built"
    spec = read_json(project / "spec.json", default=None)
    if isinstance(spec, dict) and spec.get("frozen") is True:
        return "frozen"
    if (project / "draft.md").is_file():
        return "draft"
    return "empty"


def cmd_validate(project: Path) -> dict:
    """按阶段判据检查。

    判据分三级：
      required  缺了就无法继续（P0）
      expected  该阶段应该有（P1）
      advisory  建议有（P2）
    """
    if not project.is_dir():
        return {"ok": False, "reason": "missing-project", "message": f"{project} 不存在"}

    phase = detect_phase(project)
    errors: list[dict] = []
    warnings: list[dict] = []

    manifest = load_manifest(project)
    spec = read_json(project / "spec.json", default=None)
    has_spec = isinstance(spec, dict) and bool(spec)
    spec_frozen = has_spec and spec.get("frozen") is True
    has_draft = (project / "draft.md").is_file()
    has_output = (project / "output.docx").is_file()
    has_final = (project / "final.docx").is_file()

    if not manifest_path(project).is_file():
        warnings.append({"id": "no-manifest", "message": "没有 project.json，建议先跑 project init"})

    if not has_spec:
        errors.append({"id": "missing-spec", "message": "没有 spec.json，先 word_spec init"})
    elif not spec_frozen and phase in ("built", "reviewed", "delivered"):
        errors.append({"id": "spec-not-frozen", "message": "已有产物但规格未冻结，产物可能不对应任何已确认规格"})

    if has_spec and spec_frozen:
        title = spec.get("title")
        if not isinstance(title, str) or title.strip() == "":
            errors.append({"id": "empty-title", "message": "规格缺标题"})
        sections = spec.get("sections")
        if not isinstance(sections, list) or len(sections) == 0:
            warnings.append({"id": "no-sections", "message": "规格没有列章节，质检门的缺章节检查无法生效"})

    if phase in ("built", "reviewed", "delivered"):
        if not has_draft:
            errors.append({"id": "missing-draft", "message": "有产物但没 draft.md，无法再改"})
        gate = read_json(project / "gate-report.json", default=None)
        if not isinstance(gate, dict):
            warnings.append({"id": "no-gate-report", "message": "没跑过 word_gate"})
        else:
            p0 = (gate.get("p0") or [])
            if p0:
                errors.append(
                    {
                        "id": "gate-p0",
                        "message": f"质检还有 {len(p0)} 条 P0 未清",
                        "detail": [f.get("message") for f in p0[:5]],
                    }
                )
            if gate.get("checkedAt"):
                # 报告比 draft 旧 = 报告过期
                try:
                    gate_ts = datetime.fromisoformat(gate["checkedAt"].replace("Z", "+00:00"))
                    draft_ts = datetime.fromtimestamp(
                        (project / "draft.md").stat().st_mtime, tz=timezone.utc
                    )
                    if draft_ts > gate_ts:
                        warnings.append(
                            {"id": "stale-gate", "message": "draft.md 比质检报告新，报告已过期，重跑 word_gate"}
                        )
                except (ValueError, OSError):
                    pass

    if phase == "reviewed":
        notes = read_json(project / "review-notes.json", default=None)
        if isinstance(notes, dict):
            decisions = notes.get("decisions") or []
            result = read_json(project / "review-result.json", default=None)
            if isinstance(result, dict) and result.get("p0", 0) > 0:
                warnings.append(
                    {
                        "id": "review-with-p0",
                        "message": f"交接时还有 {result['p0']} 条 P0，AI 会先处理这些",
                    }
                )

    if phase == "delivered":
        if not (project / "visual-report.json").is_file():
            warnings.append({"id": "no-visual-report", "message": "已交付但没跑 word_visual"})
        render_dir = project / "render"
        if not render_dir.is_dir() or not any(render_dir.glob("*.pdf")):
            warnings.append({"id": "no-render", "message": "已交付但没有渲染产物，视觉验收可能没做"})

    sources_dir = project / SOURCES_DIR
    if sources_dir.is_dir():
        raw = [p for p in sources_dir.iterdir() if p.is_file()]
        converted_dir = project / CONVERTED_DIR
        converted_stems = (
            {p.stem for p in converted_dir.iterdir() if p.is_file()} if converted_dir.is_dir() else set()
        )
        unconverted = [
            p.name for p in raw if not any(p.stem.startswith(s) or s.startswith(p.stem) for s in converted_stems)
        ]
        if unconverted:
            warnings.append(
                {
                    "id": "unconverted-sources",
                    "message": f"{len(unconverted)} 份素材还没转成 Markdown",
                    "detail": unconverted[:5],
                }
            )

    next_step = {
        "empty": "project init → word_spec init",
        "draft": "写完 draft.md 后 word_build",
        "frozen": "写 draft.md",
        "built": "word_gate（P0 清零）→ word_review",
        "reviewed": "按 review-notes.json 改稿 → 重建 → word_render + word_visual",
        "delivered": "已完成",
    }.get(phase, "")

    return {
        "ok": len(errors) == 0,
        "phase": phase,
        "errors": errors,
        "warnings": warnings,
        "nextStep": next_step,
        "checkedAt": datetime.now(timezone.utc).isoformat(),
    }


# ── info ──────────────────────────────────────────────────────────────────


def cmd_info(project: Path) -> dict:
    """项目状态快照。"""
    if not project.is_dir():
        return {"ok": False, "reason": "missing-project", "message": f"{project} 不存在"}

    manifest = load_manifest(project)
    spec = read_json(project / "spec.json", default=None)
    spec = spec if isinstance(spec, dict) else {}
    gate = read_json(project / "gate-report.json", default=None)
    gate = gate if isinstance(gate, dict) else {}
    notes = read_json(project / "review-notes.json", default=None)
    notes = notes if isinstance(notes, dict) else {}

    shots = sorted((project / "render" / "pages").glob("*.png")) if (project / "render" / "pages").is_dir() else []

    def size_of(name: str) -> int | None:
        path = project / name
        return path.stat().st_size if path.is_file() else None

    return {
        "ok": True,
        "phase": detect_phase(project),
        "title": spec.get("title") or manifest.get("title"),
        "path": str(project),
        "specFrozen": spec.get("frozen") is True,
        "sections": len(spec.get("sections") or []) if isinstance(spec.get("sections"), list) else 0,
        "sources": len(manifest.get("sources") or []),
        "converted": len(list((project / CONVERTED_DIR).glob("*.md")))
        if (project / CONVERTED_DIR).is_dir()
        else 0,
        "files": {
            "draftMd": size_of("draft.md"),
            "outputDocx": size_of("output.docx"),
            "finalDocx": size_of("final.docx"),
            "gateReport": bool(gate),
            "visualReport": (project / "visual-report.json").is_file(),
            "renderPages": len(shots),
        },
        "gate": {
            "verdict": gate.get("verdict"),
            "p0": len(gate.get("p0") or []),
            "p1": len(gate.get("p1") or []),
            "checkedAt": gate.get("checkedAt"),
        },
        "review": {
            "decisions": len(notes.get("decisions") or []),
            "globalIntents": len(notes.get("globalIntents") or []),
            "confirmed": (project / "review-result.json").is_file(),
        },
    }


def format_info(info: dict) -> str:
    """一行行说明，给人看。"""
    if not info.get("ok"):
        return info.get("message", "项目不可用")
    phase_label = {
        "empty": "空项目",
        "draft": "草稿中",
        "frozen": "规格已冻结",
        "built": "已构建",
        "reviewed": "已审阅",
        "delivered": "已交付",
    }.get(info["phase"], info["phase"])

    lines = [
        f"# {info.get('title') or '(无标题)'}",
        "",
        f"- 阶段：{phase_label}",
        f"- 规格：{'已冻结' if info['specFrozen'] else '未冻结'}（{info['sections']} 章节）",
        f"- 素材：{info['sources']} 份，已转 {info['converted']} 份",
    ]
    files = info["files"]
    lines.append(
        "- 文件："
        + " ".join(
            [
                f"draft {'✓' if files['draftMd'] else '✗'}",
                f"output {'✓' if files['outputDocx'] else '✗'}",
                f"final {'✓' if files['finalDocx'] else '✗'}",
                f"渲染图 {files['renderPages']} 页",
            ]
        )
    )
    gate = info["gate"]
    if gate["verdict"]:
        lines.append(f"- 质检：{gate['verdict']}（P0 {gate['p0']}，P1 {gate['p1']}）")
    else:
        lines.append("- 质检：未跑过")
    review = info["review"]
    if review["decisions"] or review["confirmed"]:
        lines.append(
            f"- 审阅：{review['decisions']} 条决定"
            + (f"，{review['globalIntents']} 条全局要求" if review["globalIntents"] else "")
            + ("，已交接" if review["confirmed"] else "")
        )
    return "\n".join(lines)


# ── 入口 ──────────────────────────────────────────────────────────────────


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 项目管理")
    sub = parser.add_subparsers(dest="action", required=True)

    init_p = sub.add_parser("init")
    init_p.add_argument("project")
    init_p.add_argument("--title", default="")
    init_p.add_argument("--force", action="store_true", help="已存在也重建（会保留 spec/draft）")

    imp_p = sub.add_parser("import-sources")
    imp_p.add_argument("project")
    imp_p.add_argument("sources", nargs="+")
    imp_p.add_argument("--move", action="store_true", help="移动而非复制本地文件")
    imp_p.add_argument(
        "--no-convert",
        action="store_true",
        help="只收素材不转换（默认会转换，并把第一份转换结果提升为 draft.md）",
    )

    val_p = sub.add_parser("validate")
    val_p.add_argument("project")
    val_p.add_argument("--json", action="store_true")

    info_p = sub.add_parser("info")
    info_p.add_argument("project")
    info_p.add_argument("--json", action="store_true")

    args = parser.parse_args(argv)
    project = Path(args.project).expanduser()

    if args.action == "init":
        emit(cmd_init(project, args.title or None, args.force))
        return 0

    if args.action == "import-sources":
        result = cmd_import(project, args.sources, args.move, convert=not args.no_convert)
        emit(result)
        return 0 if result.get("ok") else 2

    if args.action == "validate":
        result = cmd_validate(project)
        emit(result)
        if not result.get("ok"):
            return 3
        return 0 if not result.get("warnings") else 4

    if args.action == "info":
        info = cmd_info(project)
        emit(info)
        if not args.json and info.get("ok"):
            print(format_info(info), file=sys.stderr)
        return 0 if info.get("ok") else 2

    emit({"ok": False, "reason": "unknown-action"})
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
