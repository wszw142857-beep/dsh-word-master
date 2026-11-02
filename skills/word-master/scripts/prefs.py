#!/usr/bin/env python3
"""Word Master — 偏好读写与体检。

三个动作：
- `read`：把当前偏好打印成 JSON（模型据此知道哪些约束已生效）；
- `set`：把一段补丁深合并进偏好库（用户说"以后都用宋体小四"就调它）；
- `env`：环境体检（解释器、python-docx、LibreOffice）。

`set` 的补丁格式与偏好文件同构，只写要改的键，其余保持不动。设 `null` 表示
显式关闭某一项。

用法
----
    python prefs.py read
    python prefs.py set --patch '{"formatting": {"bodySizePt": 10.5}}'
    python prefs.py set --patch-file patch.json
    python prefs.py env
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import (  # noqa: E402
    emit,
    force_utf8_stdio,
    load_preferences,
    preferences_path,
    write_json,
)


def read_action() -> dict:
    """读偏好。"""
    values = load_preferences()
    return {"ok": True, "path": str(preferences_path()), "preferences": values}


def merge(target: dict, patch: dict) -> dict:
    """深合并；`null` 保留表示显式关闭。"""
    for key, value in patch.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            merge(target[key], value)
        else:
            target[key] = value
    return target


def set_action(patch: dict) -> dict:
    """写偏好。"""
    path = preferences_path()
    current = load_preferences()
    merged = merge(current, patch)
    payload = {"schemaVersion": 1, "updatedAt": __import__("datetime").datetime.now(
        __import__("datetime").timezone.utc).isoformat(), "values": merged}
    write_json(path, payload)
    return {"ok": True, "path": str(path), "preferences": merged}


def reset_action() -> dict:
    """恢复内置默认：直接删文件，load 时会自动回落到默认值。

    刻意不写入一份"默认值副本"——那样用户改过的痕迹还在文件里，下次分不清
    哪些是他要的、哪些是默认的。删掉最干净。
    """
    path = preferences_path()
    removed = False
    try:
        if path.is_file():
            path.unlink()
            removed = True
    except OSError as exc:
        return {"ok": False, "error": str(exc)}
    return {"ok": True, "path": str(path), "removed": removed, "note": "已恢复内置默认"}


def env_action() -> dict:
    """体检。"""
    from importlib.util import find_spec

    modules = {"docx": "python-docx", "lxml": "lxml", "fitz": "PyMuPDF"}
    present = {name: find_spec(name) is not None for name in modules}
    soffice = None
    for candidate in (
        "C:/Program Files/LibreOffice/program/soffice.exe",
        "C:/Program Files (x86)/LibreOffice/program/soffice.exe",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(candidate).is_file():
            soffice = candidate
            break
    return {
        "ok": present["docx"] and present["lxml"],
        "python": sys.executable,
        "pythonVersion": sys.version.split()[0],
        "modules": {name: {"present": present[name], "install": pkg} for name, pkg in modules.items()},
        "libreoffice": soffice,
        "preferencesPath": str(preferences_path()),
        "preferencesExists": preferences_path().is_file(),
    }


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 偏好与体检")
    sub = parser.add_subparsers(dest="action", required=True)
    sub.add_parser("read")
    set_p = sub.add_parser("set")
    set_p.add_argument("--patch", default="")
    set_p.add_argument("--patch-file", default="")
    sub.add_parser("reset")
    sub.add_parser("env")
    args = parser.parse_args(argv)

    if args.action == "read":
        emit(read_action())
        return 0

    if args.action == "reset":
        emit(reset_action())
        return 0

    if args.action == "set":
        raw = args.patch
        if args.patch_file != "":
            raw = Path(args.patch_file).expanduser().read_text(encoding="utf-8-sig")
        if raw.strip() == "":
            emit({"ok": False, "message": "需要 --patch 或 --patch-file"})
            return 2
        try:
            patch = json.loads(raw)
        except json.JSONDecodeError as exc:
            emit({"ok": False, "message": f"补丁不是合法 JSON：{exc}"})
            return 2
        if not isinstance(patch, dict):
            emit({"ok": False, "message": "补丁必须是对象"})
            return 2
        emit(set_action(patch))
        return 0

    if args.action == "env":
        result = env_action()
        emit(result)
        return 0 if result["ok"] else 3

    emit({"ok": False, "message": "未知动作"})
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
