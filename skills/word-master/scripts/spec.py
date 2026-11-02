#!/usr/bin/env python3
"""Word Master — 规格冻结（spec freeze）。

流水线的第一道闸门：把「要写什么」变成一份可判定、可复核的 JSON。之后的构建、
质检、渲染全部读它，任何一步都不再临时决定口径。

用法
----
    python spec.py init   --project DIR --title T [--audience A] [--length N] [--tone tone]
    python spec.py read   --project DIR
    python spec.py freeze --project DIR [--confirm]

`init` 只写草稿；`freeze` 才把它标记为已冻结（`frozen: true`），并在冻结时快照
当前偏好与版式 —— 之后用户改偏好不会倒过来改变已完成的项目。
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import (  # noqa: E402
    emit,
    force_utf8_stdio,
    formatting_of,
    load_preferences,
    preferences_path,
    read_json,
    write_json,
)

SPEC_NAME = "spec.json"

DEFAULT_GATE = {
    "maxFullwidthBrackets": 0,
    "maxBannedHits": 0,
    "maxHeadingJumps": 0,
    "requireTitle": True,
    "requireSectionOpeners": True,
    "minSections": 1,
}


def spec_path(project: Path) -> Path:
    """项目内 spec 文件路径。"""
    return project / SPEC_NAME


def build_draft(title: str, audience: str, length: int, tone: str) -> dict:
    """生成一份未冻结的规格草稿。"""
    return {
        "schemaVersion": 1,
        "title": title,
        "audience": audience,
        "targetLength": length,
        "tone": tone,
        "genre": "",
        "sections": [],
        "frozen": False,
        "createdAt": datetime.now(timezone.utc).isoformat(),
        "gate": dict(DEFAULT_GATE),
    }


def freeze(payload: dict, preferences: dict) -> dict:
    """把草稿标记为已冻结，并快照偏好与版式。

    ⚠ 已有 `formattingSnapshot` 时**不覆盖** —— 那是 `template apply` 写进来的
    模板版式，优先级高于偏好。早期版本无条件用偏好覆盖，导致"套了模板却没用上"
    （构建仍按偏好的宋体小四走）。规则：模板 > 偏好 > 内置默认。
    """
    frozen = dict(payload)
    frozen["frozen"] = True
    frozen["frozenAt"] = datetime.now(timezone.utc).isoformat()
    frozen["preferencesSnapshot"] = preferences

    existing = frozen.get("formattingSnapshot")
    if isinstance(existing, dict) and existing:
        # 模板版式已在，保留它；只补它没覆盖到的键
        merged_fmt = formatting_of(preferences)
        merged_fmt.update(existing)
        frozen["formattingSnapshot"] = merged_fmt
        frozen["formattingSource"] = frozen.get("templateId") or "template"
    else:
        frozen["formattingSnapshot"] = formatting_of(preferences)
        frozen["formattingSource"] = "preferences"

    if not isinstance(frozen.get("gate"), dict):
        frozen["gate"] = dict(DEFAULT_GATE)
    else:
        merged = dict(DEFAULT_GATE)
        merged.update({k: v for k, v in frozen["gate"].items()})
        frozen["gate"] = merged
    return frozen


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 规格冻结")
    sub = parser.add_subparsers(dest="action", required=True)

    init_p = sub.add_parser("init", help="创建规格草稿")
    init_p.add_argument("--project", required=True)
    init_p.add_argument("--title", required=True)
    init_p.add_argument("--audience", default="")
    init_p.add_argument("--length", type=int, default=3000)
    init_p.add_argument("--tone", default="正式、克制")
    init_p.add_argument("--genre", default="")

    read_p = sub.add_parser("read", help="读取规格")
    read_p.add_argument("--project", required=True)

    freeze_p = sub.add_parser("freeze", help="冻结规格（此后构建与质检都以它为准）")
    freeze_p.add_argument("--project", required=True)

    args = parser.parse_args(argv)
    project = Path(args.project).expanduser()

    if args.action == "init":
        project.mkdir(parents=True, exist_ok=True)
        target = spec_path(project)
        if target.is_file():
            existing = read_json(target, default={})
            if isinstance(existing, dict) and existing.get("frozen") is True:
                emit(
                    {
                        "ok": False,
                        "reason": "already-frozen",
                        "message": "该项目的规格已冻结。要改规格先人工确认后重新 init（会覆盖草稿，不会覆盖已冻结版本）。",
                        "spec": existing,
                    }
                )
                return 2
        draft = build_draft(args.title, args.audience, args.length, args.tone)
        if args.genre:
            draft["genre"] = args.genre
        write_json(target, draft)
        emit({"ok": True, "path": str(target), "spec": draft, "next": "补全 sections 后调 word_spec freeze"})
        return 0

    if args.action == "read":
        payload = read_json(spec_path(project), default=None)
        if payload is None:
            emit({"ok": False, "reason": "missing-spec", "message": "该项目还没有规格，先调 word_spec init。"})
            return 2
        emit({"ok": True, "spec": payload})
        return 0

    if args.action == "freeze":
        payload = read_json(spec_path(project), default=None)
        if payload is None:
            emit({"ok": False, "reason": "missing-spec", "message": "没有可冻结的规格，先调 word_spec init。"})
            return 2
        if not isinstance(payload.get("title"), str) or payload["title"].strip() == "":
            emit({"ok": False, "reason": "missing-title", "message": "规格缺标题，冻结前必须填。"})
            return 2
        preferences = load_preferences()
        frozen = freeze(payload, preferences)
        write_json(spec_path(project), frozen)
        emit(
            {
                "ok": True,
                "path": str(spec_path(project)),
                "frozenAt": frozen["frozenAt"],
                "preferencesPath": str(preferences_path()),
                "gate": frozen["gate"],
                "formatting": frozen["formattingSnapshot"],
            }
        )
        return 0

    emit({"ok": False, "reason": "unknown-action"})
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
