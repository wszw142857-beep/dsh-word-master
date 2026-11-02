#!/usr/bin/env python3
"""Word Master — 标注式 HTML 审阅界面（后端）。

与 ppt-master `confirm_ui` 的关系
--------------------------------
结构完全对齐：Flask 本地服务 + `static/` 静态资源 + `result.json` 交接文件 +
文件锁防重复启动 + 空闲看护 + `/api/shutdown`。视觉语言也沿用同一套色板与布局。

差别在**业务形态**：ppt-master 是「选参数」，一次提交就结束；Word 的审阅需要
逐条处理问题并反复改稿，所以多了三个东西：

1. **双向定位**：发现带行号 → 点它跳到正文对应块并高亮。
2. **回写**：正文块可编辑，保存时整块替换写回 `draft.md`。
3. **意见通道**：段落级 💬 与整篇级文本框，全部写进 `review-notes.json`。
   这是给 AI 的指令，不是给自己看的便签 —— 所以带原文引用，防止改稿后错位。

三个硬约束
----------
1. **只改 draft.md，不改 docx**。docx 由构建脚本生成，手改会破坏三线表与样式继承。
2. **整块替换而非行内定位**。行内字节偏移一旦用户在别处同时编辑就错位；整块
   匹配最稳，不匹配时拒绝写入而不是写坏。
3. **服务器必须能自己退出**。占着端口的僵尸进程比没有界面更糟。

用法
----
    python review_server.py <project> --port 5070 --timeout 1800
    python review_server.py <project> --daemon      # 后台起并打印 URL
"""

from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wm_common import force_utf8_stdio, load_preferences, read_json, write_json  # noqa: E402
import wm_checks  # noqa: E402
import draft_model  # noqa: E402

DEFAULT_PORT = 5070
LOCK_NAME = ".review-lock.json"
RESULT_NAME = "review-result.json"
NOTES_NAME = "review-notes.json"
UI_DIR = Path(__file__).resolve().parent / "review_ui"
CATALOGS_NAME = "catalogs.json"


def load_catalogs() -> dict:
    """读结构化选项清单（枚举完整清单，不经 AI 现场生成）。"""
    payload = read_json(UI_DIR / CATALOGS_NAME, default=None)
    return payload if isinstance(payload, dict) else {}


def derive_recommendations(issues: list[dict], catalogs: dict) -> dict:
    """按当前质检发现标出推荐的处理方式。

    这是把 ppt-master「AI 标注 recommended」的机制移植到审阅场景：不由模型现场
    生成推荐（那样每次要多一轮调用、且同一问题可能给不同答案），而是按发现类型
    做确定性映射 —— 同样的问题总是推荐同样的处理方式，用户能预期。
    """
    verdict_by_issue = {
        "fullwidth-brackets": "fix_language",
        "banned-defensive": "rewrite",
        "banned-delivery": "rewrite",
        "banned-meta": "rewrite",
        "banned-process": "rewrite",
        "section-opener": "fix_logic",
        "method-numbers": "move_to_method",
        "heading-jump": "fix_structure",
        "missing-title": "fix_language",
        "missing-sections": "fix_structure",
    }
    global_by_issue = {
        "section-opener": "strengthen_opener",
        "method-numbers": "fix_structure",
        "banned-meta": "unify_tone",
        "banned-delivery": "tighten",
        "missing-sections": "fix_structure",
    }

    verdict_counts: dict[str, int] = {}
    global_counts: dict[str, int] = {}
    for issue in issues:
        verdict = verdict_by_issue.get(issue.get("id"))
        if verdict:
            verdict_counts[verdict] = verdict_counts.get(verdict, 0) + 1
        intent = global_by_issue.get(issue.get("id"))
        if intent:
            global_counts[intent] = global_counts.get(intent, 0) + 1

    top_verdict = max(verdict_counts, key=verdict_counts.get) if verdict_counts else None
    top_global = max(global_counts, key=global_counts.get) if global_counts else None

    valid_verdicts = {v["id"] for v in catalogs.get("verdicts", []) if isinstance(v, dict) and "id" in v}
    valid_globals = {v["id"] for v in catalogs.get("global_intents", []) if isinstance(v, dict) and "id" in v}

    return {
        "verdict": top_verdict if top_verdict in valid_verdicts else None,
        "global": top_global if top_global in valid_globals else None,
    }


DONE_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Word Master — 已提交</title>
<style>
  html,body{height:100%;margin:0;background:#1a1a2e;color:#e0e0e0;
    font-family:-apple-system,"Segoe UI","Microsoft YaHei",sans-serif;
    display:grid;place-items:center;}
  .card{background:#1e1e36;border:1px solid #2a2a4a;border-radius:14px;
    padding:40px 48px;text-align:center;max-width:480px;
    box-shadow:0 16px 50px rgba(0,0,0,.6);}
  .t{font-size:24px;font-weight:700;color:#3fb950;margin-bottom:14px;}
  .h{font-size:14px;line-height:1.7;color:#a8a8c0;}
  code{background:#12121f;padding:2px 7px;border-radius:4px;font-size:13px;color:#8fb0ff;}
</style>
</head>
<body>
  <div class="card">
    <div class="t">✓ 已提交</div>
    <div class="h">
      标注已写入 <code>review-notes.json</code>，正文改动已写回 <code>draft.md</code>。<br><br>
      现在回到对话里说「继续」，AI 会读这两份文件按你的决定改稿。
    </div>
  </div>
</body>
</html>
"""


def find_free_port(preferred: int, host: str = "127.0.0.1", span: int = 50) -> int:
    """从 preferred 起找一个能绑的端口。"""
    for port in range(preferred, preferred + span):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
            probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                probe.bind((host, port))
                return port
            except OSError:
                continue
    return preferred


def process_alive(pid: int) -> bool:
    """进程是否还活着。"""
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True


def read_lock(path: Path) -> dict | None:
    """读锁文件。"""
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def claim_lock(path: Path, port: int) -> bool:
    """占锁；已有活进程持锁则失败。"""
    existing = read_lock(path)
    if existing and process_alive(int(existing.get("pid", 0))):
        return False
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"pid": os.getpid(), "port": port, "startedAt": time.time()}),
        encoding="utf-8",
    )
    return True


def release_lock(path: Path) -> None:
    """释放锁。"""
    try:
        path.unlink()
    except OSError:
        pass


def load_notes(project: Path) -> dict:
    """读决定文件；不存在时返回空结构。

    结构（与 ppt-master result.json 的精神一致：结构化字段，不是自由文本）：
        {
          "schemaVersion": 2,
          "decisions": [{"line", "quote", "verdicts": [id], "note"}],
          "globalIntents": [id],
          "globalNote": "..."
        }
    """
    payload = read_json(project / NOTES_NAME, default=None)
    if isinstance(payload, dict):
        return {
            "schemaVersion": 2,
            "decisions": payload.get("decisions") if isinstance(payload.get("decisions"), list) else [],
            "globalIntents": payload.get("globalIntents") if isinstance(payload.get("globalIntents"), list) else [],
            "globalNote": payload.get("globalNote") if isinstance(payload.get("globalNote"), str) else "",
        }
    return {"schemaVersion": 2, "decisions": [], "globalIntents": [], "globalNote": ""}


def save_notes(project: Path, payload: dict) -> None:
    """写意见文件。"""
    payload["updatedAt"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    write_json(project / NOTES_NAME, payload)


def build_app(project: Path, state: dict):
    """构造 Flask 应用。"""
    from flask import Flask, jsonify, request, send_from_directory

    app = Flask(__name__, static_folder=None)

    def current_blocks():
        draft = project / "draft.md"
        if not draft.is_file():
            return []
        return draft_model.blocks_from_markdown(draft)

    def draft_mtime() -> float | None:
        """draft.md 的修改时间；文件不存在返回 None。

        这是前端检测"磁盘被改过"的唯一依据 —— 用 mtime 而不是内容哈希，因为
        轮询每 2 秒一次，哈希整个文件太贵；mtime 变化的粒度足够（用户/AI 的写
        操作都在秒级）。
        """
        draft = project / "draft.md"
        try:
            return draft.stat().st_mtime
        except OSError:
            return None

    def current_spec() -> dict:
        spec = read_json(project / "spec.json", default={})
        return spec if isinstance(spec, dict) else {}

    def current_issues(blocks):
        preferences = load_preferences()
        spec = current_spec()
        gate_cfg = spec.get("gate") if isinstance(spec.get("gate"), dict) else {}
        return wm_checks.run_all(blocks, preferences=preferences, spec=spec, gate_cfg=gate_cfg)

    # ── 静态资源 ──────────────────────────────────────────────────────

    @app.route("/")
    def index():
        response = send_from_directory(str(UI_DIR), "index.html")
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.route("/static/<path:name>")
    def static_file(name):
        response = send_from_directory(str(UI_DIR), name)
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.route("/done")
    def done():
        """提交后的落地页（等价于 svg_editor 提交后关页面）。"""
        from flask import Response

        return Response(DONE_HTML, mimetype="text/html; charset=utf-8")

    # ── 查询 ──────────────────────────────────────────────────────────

    @app.route("/api/meta")
    def meta():
        spec = current_spec()
        return jsonify(
            {
                "title": spec.get("title"),
                "project": str(project),
                "genre": spec.get("genre"),
                "frozen": spec.get("frozen"),
            }
        )

    @app.route("/api/check")
    def check():
        blocks = current_blocks()
        result = current_issues(blocks)
        state["lastCheck"] = time.time()

        payload_blocks = [
            {
                "line": b["line"],
                "kind": b["kind"],
                "level": b.get("level"),
                "text": b["text"],
                "raw": b["raw"],
            }
            for b in blocks
        ]

        issues = result["p0"] + result["p1"]
        title = current_spec().get("title")
        # 缺标题可以一键补：把标题行插到文件最前面。
        for issue in issues:
            if issue["id"] == "missing-title" and isinstance(title, str) and title.strip() != "":
                issue["insertTitle"] = f"# {title.strip()}"

        return jsonify(
            {
                "ok": True,
                "blocks": payload_blocks,
                "issues": issues,
                "verdict": result["verdict"],
                "counts": result["counts"],
                "mtime": draft_mtime(),
            }
        )

    @app.route("/api/mtime")
    def mtime_only():
        """轻量轮询端点：只回 draft.md 的 mtime，不重跑检查。

        前端每 2 秒打一次（对齐 ppt-master svg_editor 的轮询节奏），发现 mtime
        与上次渲染时记录的不一致，就弹「文件已在磁盘更新」的横幅 —— 让用户知道
        AI 那边改过稿了，而不是看着旧内容继续做决定。
        """
        return jsonify({"ok": True, "mtime": draft_mtime()})

    @app.route("/api/catalogs")
    def catalogs():
        """完整选项清单 + 推荐标记（前端渲染 chips 用）。"""
        catalog = load_catalogs()
        blocks = current_blocks()
        result = current_issues(blocks)
        recommendations = derive_recommendations(result["p0"] + result["p1"], catalog)
        return jsonify({**catalog, "recommendations": recommendations})

    @app.route("/api/decisions")
    def decisions():
        return jsonify({"ok": True, **load_notes(project)})

    @app.route("/api/opinions")
    def opinions_legacy():
        """兼容旧字段名；新前端用 /api/decisions。"""
        return jsonify({"ok": True, **load_notes(project)})

    # ── 写操作 ────────────────────────────────────────────────────────

    @app.route("/api/stage", methods=["POST"])
    def stage_edit():
        """暂存一次正文改动（不落盘）。

        与 ppt-master svg_editor 的 `/api/slide/<name>/edit` 同义：编辑立即在预览里
        可见，但磁盘写入延后到「提交标注」时一次性完成。

        为什么要暂存而不是直接写：
        1. 用户可能改到一半反悔 —— 暂存 + 撤销可以完全回退，直接写就要靠备份。
        2. 提交前应该是一个**原子动作**：要么全写、要么全不写，避免"改了一半
           服务器崩了"留下半成品。
        3. 配合 mtime 漂移检测：暂存改动不碰磁盘，所以不会触发自己引起的漂移。

        暂存队列只在内存（`state["staged"]`），服务器一关就没了 —— 这是刻意的，
        没提交的东西本来就该丢。
        """
        payload = request.get_json(silent=True) or {}
        line = payload.get("line")
        after = payload.get("after")
        if not isinstance(line, int):
            return jsonify({"ok": False, "error": "line 必须是整数"}), 400
        if not isinstance(after, str):
            return jsonify({"ok": False, "error": "after 必须是字符串"}), 400

        blocks = current_blocks()
        block = next((b for b in blocks if b["line"] == line), None)
        if block is None:
            return jsonify({"ok": False, "error": f"没有第 {line} 行对应的块"}), 404

        staged: list = state.setdefault("staged", [])
        # 同一行只保留最后一次编辑（栈里存原始值，用于撤销）
        original = block["raw"]
        staged[:] = [s for s in staged if s.get("line") != line]
        staged.append({"line": line, "kind": block["kind"], "before": original, "after": after})

        return jsonify({"ok": True, "staged": len(staged), "line": line})

    @app.route("/api/stage/undo", methods=["POST"])
    def undo_stage():
        """撤销上一次暂存（不是撤销已落盘的）。"""
        staged: list = state.setdefault("staged", [])
        if not staged:
            return jsonify({"ok": True, "removed": None, "staged": 0, "note": "没有可撤销的暂存"})
        removed = staged.pop()
        return jsonify({"ok": True, "removed": removed, "staged": len(staged)})

    @app.route("/api/staged")
    def list_staged():
        """列出当前暂存队列（供页面刷新后恢复）。"""
        staged: list = state.setdefault("staged", [])
        return jsonify({"ok": True, "staged": staged})

    @app.route("/api/save", methods=["POST"])
    def save():
        """把暂存队列一次性落盘（原子动作）。"""
        draft = project / "draft.md"
        if not draft.is_file():
            return jsonify({"ok": False, "error": "draft.md 不存在"}), 400

        staged: list = state.setdefault("staged", [])
        payload = request.get_json(silent=True) or {}
        # 也允许直接传 edits（兼容旧调用与 curl 调试）
        edits = payload.get("edits") if isinstance(payload.get("edits"), list) else staged
        if not edits:
            return jsonify({"ok": True, "applied": [], "rejected": [], "mtime": draft_mtime()})

        original = draft_model.source_text(draft)
        result = draft_model.write_back(draft, original, edits)
        if result.get("applied"):
            state["staged"] = []
        state["lastCheck"] = time.time()
        result["mtime"] = draft_mtime()
        return jsonify(result)

    @app.route("/api/insert-title", methods=["POST"])
    def insert_title():
        """在文件最前面插入标题行。"""
        title = current_spec().get("title")
        if not isinstance(title, str) or title.strip() == "":
            return jsonify({"ok": False, "error": "spec 里没有标题"}), 400
        draft = project / "draft.md"
        if not draft.is_file():
            return jsonify({"ok": False, "error": "draft.md 不存在"}), 400
        original = draft_model.source_text(draft)
        if original.lstrip().startswith(f"# {title.strip()}"):
            return jsonify({"ok": True, "applied": [], "note": "已有标题"})
        draft.write_text(f"# {title.strip()}\n\n{original.lstrip()}", encoding="utf-8")
        state["lastCheck"] = time.time()
        return jsonify({"ok": True, "applied": [{"line": 1}]})

    @app.route("/api/decision", methods=["POST"])
    def add_decision():
        """记一段的处理决定。

        两个通道都成立，且可以同时用：
        - `verdicts`：结构化选择（来自 catalogs.json 的枚举）
        - `note`：自由文本补充

        **只要有一个非空就接受**。早期版本强制要求 verdicts 非空，等于把自由
        文本降级成"补充说明"——但用户常常只有一句话要说、并不想归类，那种时候
        逼他选 chips 是多余摩擦。
        """
        payload = request.get_json(silent=True) or {}
        line = payload.get("line")
        if not isinstance(line, int):
            return jsonify({"ok": False, "error": "line 必须是整数"}), 400

        verdicts = payload.get("verdicts")
        if verdicts is None:
            verdicts = []
        if not isinstance(verdicts, list):
            return jsonify({"ok": False, "error": "verdicts 必须是数组"}), 400

        note = payload.get("note") if isinstance(payload.get("note"), str) else ""
        note = note.strip()

        if not verdicts and note == "":
            return jsonify({"ok": False, "error": "至少选一个处理方式，或写一句具体要求"}), 400

        catalog = load_catalogs()
        valid = {v["id"] for v in catalog.get("verdicts", []) if isinstance(v, dict) and "id" in v}
        unknown = [v for v in verdicts if v not in valid]
        if unknown:
            return jsonify({"ok": False, "error": f"未知的处理方式：{unknown}"}), 400

        quote = payload.get("quote") if isinstance(payload.get("quote"), str) else ""

        notes = load_notes(project)
        entry = {
            "line": line,
            "quote": quote,
            "verdicts": verdicts,
            "note": note,
            "updatedAt": time.strftime("%Y-%m-%dT%H:%M:%S"),
        }
        notes["decisions"] = [d for d in notes["decisions"] if d.get("line") != line]
        notes["decisions"].append(entry)
        notes["decisions"].sort(key=lambda d: d.get("line") or 0)
        save_notes(project, notes)
        state["lastCheck"] = time.time()
        return jsonify({"ok": True, "decision": entry})

    @app.route("/api/decision/delete", methods=["POST"])
    def delete_decision():
        payload = request.get_json(silent=True) or {}
        line = payload.get("line")
        notes = load_notes(project)
        before = len(notes["decisions"])
        notes["decisions"] = [d for d in notes["decisions"] if d.get("line") != line]
        save_notes(project, notes)
        state["lastCheck"] = time.time()
        return jsonify({"ok": True, "removed": before - len(notes["decisions"])})

    @app.route("/api/global", methods=["POST"])
    def set_global():
        """整篇要求：结构化 intents + 自由补充说明。"""
        payload = request.get_json(silent=True) or {}
        notes = load_notes(project)

        if "intents" in payload:
            intents = payload["intents"]
            if not isinstance(intents, list):
                return jsonify({"ok": False, "error": "intents 必须是数组"}), 400
            catalog = load_catalogs()
            valid = {v["id"] for v in catalog.get("global_intents", []) if isinstance(v, dict) and "id" in v}
            unknown = [i for i in intents if i not in valid]
            if unknown:
                return jsonify({"ok": False, "error": f"未知的整篇要求：{unknown}"}), 400
            notes["globalIntents"] = intents

        if isinstance(payload.get("note"), str):
            notes["globalNote"] = payload["note"].strip()

        save_notes(project, notes)
        state["lastCheck"] = time.time()
        return jsonify({"ok": True, "globalIntents": notes["globalIntents"], "globalNote": notes["globalNote"]})


    @app.route("/api/confirm", methods=["POST"])
    def confirm():
        """交接：写 review-result.json，供 AI 读取。"""
        blocks = current_blocks()
        issues = current_issues(blocks)
        notes = load_notes(project)
        record = {
            "status": "confirmed" if issues["counts"]["p0"] == 0 else "blocked",
            "confirmedAt": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "p0": issues["counts"]["p0"],
            "p1": issues["counts"]["p1"],
            "decisions": len(notes["decisions"]),
            "globalIntents": len(notes["globalIntents"]),
            "notesFile": NOTES_NAME,
            "draftFile": "draft.md",
        }
        write_json(project / RESULT_NAME, record)
        return jsonify({"ok": True, "record": record})

    @app.route("/api/shutdown", methods=["POST"])
    def shutdown():
        def stop():
            time.sleep(0.3)
            os._exit(0)

        threading.Thread(target=stop, daemon=True).start()
        return jsonify({"ok": True})

    @app.route("/api/health")
    def health():
        return jsonify({"ok": True, "pid": os.getpid(), "lastCheck": state.get("lastCheck")})

    return app


def main(argv: list[str] | None = None) -> int:
    """入口。"""
    force_utf8_stdio()
    parser = argparse.ArgumentParser(description="Word Master 审阅界面")
    parser.add_argument("project")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--timeout", type=int, default=1800, help="空闲超时秒数，0 表示不超时")
    parser.add_argument("--no-browser", action="store_true")
    parser.add_argument("--daemon", action="store_true", help="后台启动并打印 URL")
    args = parser.parse_args(argv)

    project = Path(args.project).expanduser().resolve()
    if not (project / "draft.md").is_file():
        print(f"[ERR] {project} 下没有 draft.md，先写源稿再开审阅界面。", file=sys.stderr)
        return 2
    if not (UI_DIR / "index.html").is_file():
        print(f"[ERR] 界面资源缺失：{UI_DIR / 'index.html'}", file=sys.stderr)
        return 2

    lock_path = project / LOCK_NAME

    if args.daemon:
        existing = read_lock(lock_path)
        if existing and process_alive(int(existing.get("pid", 0))):
            print(
                json.dumps(
                    {
                        "ok": True,
                        "already": True,
                        "url": f"http://127.0.0.1:{existing.get('port')}/",
                        "pid": existing.get("pid"),
                    },
                    ensure_ascii=False,
                )
            )
            return 0
        port = find_free_port(args.port)
        cmd = [
            sys.executable,
            str(Path(__file__).resolve()),
            str(project),
            "--port",
            str(port),
            "--timeout",
            str(args.timeout),
        ]
        creationflags = 0
        if os.name == "nt":
            creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
        log_dir = project / ".review"
        log_dir.mkdir(parents=True, exist_ok=True)
        log = open(log_dir / "server.log", "ab")
        child = subprocess.Popen(cmd, stdout=log, stderr=log, creationflags=creationflags, close_fds=True)

        for _ in range(120):
            try:
                with socket.create_connection(("127.0.0.1", port), timeout=0.4):
                    print(
                        json.dumps(
                            {
                                "ok": True,
                                "url": f"http://127.0.0.1:{port}/",
                                "pid": child.pid,
                                "project": str(project),
                            },
                            ensure_ascii=False,
                        )
                    )
                    return 0
            except OSError:
                time.sleep(0.1)
        print(json.dumps({"ok": False, "error": "server did not start", "pid": child.pid}, ensure_ascii=False))
        return 3

    port = find_free_port(args.port)
    if not claim_lock(lock_path, port):
        existing = read_lock(lock_path) or {}
        print(
            f"[ERR] 该项目已有审阅界面在运行（pid={existing.get('pid')}, "
            f"port={existing.get('port')}）。打开 http://127.0.0.1:{existing.get('port')}/",
            file=sys.stderr,
        )
        return 1

    state: dict = {"lastCheck": time.time()}
    app = build_app(project, state)

    if args.timeout > 0:

        def watchdog():
            while True:
                time.sleep(5)
                if time.time() - state.get("lastCheck", time.time()) > args.timeout:
                    release_lock(lock_path)
                    os._exit(0)

        threading.Thread(target=watchdog, daemon=True).start()

    url = f"http://127.0.0.1:{port}/"
    print(f"Word Master 审阅界面：{url}", file=sys.stderr)
    print(json.dumps({"ok": True, "url": url, "pid": os.getpid()}, ensure_ascii=False))
    sys.stderr.flush()

    if not args.no_browser:
        threading.Thread(target=lambda: (time.sleep(0.8), webbrowser.open(url)), daemon=True).start()

    try:
        app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False)
    finally:
        release_lock(lock_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
