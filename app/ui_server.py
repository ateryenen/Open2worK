"""Flask UI server for Open2work desktop automation agent."""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

import ui_tree2

from .executor import DeterministicExecutor
from .monitor import LocalLLMScreenMonitor
from .planner import RuleBasedPlanner
from .schemas import ExecutionPlan

app = Flask(__name__, template_folder="templates", static_folder="static")

_execution_log: list[dict] = []
_log_lock = threading.Lock()
_TEMPLATE_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}


def _append_log(entry: dict) -> None:
    with _log_lock:
        _execution_log.append(entry)
        if len(_execution_log) > 200:
            _execution_log.pop(0)


def _format_tree_text(tree: dict) -> str:
    if not isinstance(tree, dict):
        return ""

    root_name = str(tree.get("name") or "Unknown")
    lines = [root_name]

    def walk(node: dict, prefix: str, is_last: bool) -> None:
        branch = "└─ " if is_last else "├─ "
        lines.append(prefix + branch + str(node.get("name") or "Unknown"))
        children = node.get("children") if isinstance(node.get("children"), list) else []
        next_prefix = prefix + ("   " if is_last else "│  ")
        for idx, child in enumerate(children):
            if isinstance(child, dict):
                walk(child, next_prefix, idx == len(children) - 1)

    children = tree.get("children") if isinstance(tree.get("children"), list) else []
    for idx, child in enumerate(children):
        if isinstance(child, dict):
            walk(child, "", idx == len(children) - 1)

    return "\n".join(lines)


def _safe_folder_name(value: str) -> str:
    cleaned = re.sub(r"[^\w\-\u4e00-\u9fff]+", "_", value, flags=re.UNICODE).strip("_")
    return cleaned[:64] or "unknown_app"


def _persist_ui_tree(repo_root: Path, payload: dict) -> list[str]:
    app_name = str(payload.get("name") or "unknown_app")
    app_folder = _safe_folder_name(app_name)
    output_dir = repo_root / "assets" / "ui_trees" / app_folder
    output_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_path = output_dir / f"{ts}.ui_tree.json"
    txt_path = output_dir / f"{ts}.ui_tree.txt"

    with json_path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    txt_path.write_text(_format_tree_text(payload), encoding="utf-8")

    return [
        json_path.relative_to(repo_root).as_posix(),
        txt_path.relative_to(repo_root).as_posix(),
    ]


def _is_rect_overlap(a: dict, b: dict) -> bool:
    return not (a["right"] < b["left"] or a["left"] > b["right"] or a["bottom"] < b["top"] or a["top"] > b["bottom"])


def _rect_area(r: dict) -> int:
    width = max(0, int(r.get("right", 0)) - int(r.get("left", 0)))
    height = max(0, int(r.get("bottom", 0)) - int(r.get("top", 0)))
    return width * height


def _intersection_area(a: dict, b: dict) -> int:
    left = max(int(a["left"]), int(b["left"]))
    top = max(int(a["top"]), int(b["top"]))
    right = min(int(a["right"]), int(b["right"]))
    bottom = min(int(a["bottom"]), int(b["bottom"]))
    if right <= left or bottom <= top:
        return 0
    return (right - left) * (bottom - top)


def _roi_match(rect: dict, roi: dict, min_candidate_cover: float = 0.35) -> bool:
    inter = _intersection_area(rect, roi)
    if inter <= 0:
        return False
    candidate_area = _rect_area(rect)
    if candidate_area <= 0:
        return False
    candidate_cover = inter / candidate_area
    return candidate_cover >= min_candidate_cover


def _score_candidate(node: dict) -> int:
    score = 0
    control_type = str(node.get("control_type") or "")
    class_name = str(node.get("class_name") or "").lower()
    interactive_types = {"Button", "MenuItem", "Edit", "ComboBox", "TabItem", "TreeItem", "ListItem", "Hyperlink"}
    if control_type in interactive_types:
        score += 50
    if any(key in class_name for key in ["button", "edit", "combo", "list", "tree", "tab", "menu"]):
        score += 35
    if node.get("is_enabled"):
        score += 20
    if node.get("is_visible"):
        score += 20
    if node.get("name"):
        score += 5
    if node.get("automation_id"):
        score += 5
    return score


def _collect_roi_candidates(raw_tree: dict, roi: dict) -> list[dict]:
    candidates: list[dict] = []
    seen: set[tuple] = set()

    def walk(node: dict, depth: int = 0) -> None:
        if not isinstance(node, dict):
            return
        rect = node.get("rectangle")
        if isinstance(rect, dict) and all(k in rect for k in ["left", "top", "right", "bottom"]):
            if _roi_match(rect, roi):
                item = {
                    "name": node.get("name") or "",
                    "control_type": node.get("control_type") or "",
                    "class_name": node.get("class_name"),
                    "automation_id": node.get("automation_id"),
                    "rectangle": rect,
                    "is_visible": bool(node.get("is_visible")),
                    "is_enabled": bool(node.get("is_enabled")),
                }
                item["rank_score"] = _score_candidate(item)
                # Skip depth-0 root window to avoid container-only false positives
                if depth > 0:
                    dedupe_key = (
                        item.get("name"),
                        item.get("control_type"),
                        item.get("class_name"),
                        item.get("rectangle", {}).get("left"),
                        item.get("rectangle", {}).get("top"),
                        item.get("rectangle", {}).get("right"),
                        item.get("rectangle", {}).get("bottom"),
                    )
                    if dedupe_key not in seen:
                        seen.add(dedupe_key)
                        candidates.append(item)

        children = node.get("children") if isinstance(node.get("children"), list) else []
        for child in children:
            walk(child, depth + 1)

    walk(raw_tree, depth=0)
    candidates.sort(key=lambda x: x.get("rank_score", 0), reverse=True)
    return candidates


def _container_only(candidates: list[dict]) -> bool:
    if not candidates:
        return True
    kinds = {str(item.get("control_type") or "") for item in candidates}
    containers = {"Window", "Pane", "TitleBar", "Group"}
    has_interactive = any(item.get("rank_score", 0) >= 50 for item in candidates)
    return kinds.issubset(containers) and not has_interactive


def _filter_tree_to_roi(node: dict, roi: dict) -> dict | None:
    if not isinstance(node, dict):
        return None

    rect = node.get("rectangle")
    children = node.get("children") if isinstance(node.get("children"), list) else []
    filtered_children: list[dict] = []
    for child in children:
        child_filtered = _filter_tree_to_roi(child, roi)
        if child_filtered is not None:
            filtered_children.append(child_filtered)

    overlap = isinstance(rect, dict) and all(k in rect for k in ["left", "top", "right", "bottom"]) and _roi_match(rect, roi)
    if overlap or filtered_children:
        kept = dict(node)
        kept["children"] = filtered_children
        return kept
    return None


def _candidate_quality(candidates: list[dict]) -> int:
    total = 0
    for item in candidates:
        score = int(item.get("rank_score", 0))
        if str(item.get("control_type") or "") == "Unknown":
            score -= 15
        if not item.get("name"):
            score -= 10
        total += score
    return total


def _capture_roi_image(roi: dict):
    from PIL import ImageGrab

    return ImageGrab.grab(
        bbox=(
            int(roi["left"]),
            int(roi["top"]),
            int(roi["right"]),
            int(roi["bottom"]),
        )
    )


def _persist_roi_image(repo_root: Path, app_name: str, image) -> str:
    output_dir = repo_root / "assets" / "ui_trees" / _safe_folder_name(app_name)
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    image_path = output_dir / f"{ts}.roi.png"
    image.save(image_path)
    return image_path.relative_to(repo_root).as_posix()


def _collect_ocr_candidates(
    roi: dict,
    image=None,
    source_image_path: str | None = None,
    min_confidence: float = 45.0,
) -> tuple[list[dict], dict]:
    meta: dict = {
        "enabled": False,
        "engine": "tesseract",
        "error": None,
        "version": None,
        "kept": 0,
        "source_image_path": source_image_path,
    }

    try:
        import pytesseract
        from pytesseract import Output
    except Exception as exc:
        meta["error"] = f"python ocr import failed: {exc}"
        return [], meta

    tesseract_cmd = os.getenv("TESSERACT_CMD", "").strip()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    try:
        meta["version"] = str(pytesseract.get_tesseract_version())
    except Exception as exc:
        meta["error"] = f"tesseract unavailable: {exc}"
        return [], meta

    if image is None:
        try:
            image = _capture_roi_image(roi)
        except Exception as exc:
            meta["error"] = f"roi screenshot failed: {exc}"
            return [], meta

    try:
        data = pytesseract.image_to_data(image, output_type=Output.DICT, config="--oem 3 --psm 6")
    except Exception as exc:
        meta["error"] = f"ocr execution failed: {exc}"
        return [], meta

    candidates: list[dict] = []
    seen: set[tuple] = set()
    count = max(
        len(data.get("text", [])),
        len(data.get("conf", [])),
        len(data.get("left", [])),
        len(data.get("top", [])),
    )

    for idx in range(count):
        raw_text = str(data.get("text", [""])[idx] if idx < len(data.get("text", [])) else "").strip()
        if not raw_text:
            continue

        raw_conf = data.get("conf", ["-1"])[idx] if idx < len(data.get("conf", [])) else "-1"
        try:
            conf = float(raw_conf)
        except Exception:
            conf = -1.0
        if conf < min_confidence:
            continue

        left = int(data.get("left", [0])[idx] if idx < len(data.get("left", [])) else 0) + int(roi["left"])
        top = int(data.get("top", [0])[idx] if idx < len(data.get("top", [])) else 0) + int(roi["top"])
        width = int(data.get("width", [0])[idx] if idx < len(data.get("width", [])) else 0)
        height = int(data.get("height", [0])[idx] if idx < len(data.get("height", [])) else 0)
        if width <= 1 or height <= 1:
            continue

        rectangle = {
            "left": left,
            "top": top,
            "right": left + width,
            "bottom": top + height,
            "width": width,
            "height": height,
        }

        dedupe_key = (raw_text.lower(), left, top, width, height)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)

        rank = int(min(100.0, conf)) + 45
        if len(raw_text) <= 2:
            rank -= 8

        candidates.append(
            {
                "name": raw_text,
                "control_type": "OCRText",
                "class_name": "ocr",
                "automation_id": None,
                "rectangle": rectangle,
                "is_visible": True,
                "is_enabled": True,
                "rank_score": max(0, rank),
                "confidence": conf,
                "source": "ocr",
            }
        )

    candidates.sort(key=lambda x: x.get("rank_score", 0), reverse=True)
    meta["enabled"] = True
    meta["kept"] = len(candidates)
    return candidates, meta


def _merge_candidates(ui_candidates: list[dict], ocr_candidates: list[dict]) -> list[dict]:
    merged: list[dict] = []
    seen: set[tuple] = set()

    def push(items: list[dict], source: str) -> None:
        for item in items:
            rect = item.get("rectangle") if isinstance(item.get("rectangle"), dict) else {}
            key = (
                str(item.get("name") or "").strip().lower(),
                str(item.get("control_type") or ""),
                int(rect.get("left", 0)),
                int(rect.get("top", 0)),
                int(rect.get("right", 0)),
                int(rect.get("bottom", 0)),
            )
            if key in seen:
                continue
            seen.add(key)
            if "source" not in item:
                item["source"] = source
            merged.append(item)

    push(ui_candidates, "ui")
    push(ocr_candidates, "ocr")
    merged.sort(key=lambda x: x.get("rank_score", 0), reverse=True)
    return merged


def _persist_roi_artifacts(
    repo_root: Path,
    app_name: str,
    roi: dict,
    ui_candidates: list[dict],
    ocr_candidates: list[dict],
    merged_candidates: list[dict],
    ocr_meta: dict,
) -> list[str]:
    output_dir = repo_root / "assets" / "ui_trees" / _safe_folder_name(app_name)
    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    roi_file = output_dir / f"{ts}.roi_candidates.json"
    with roi_file.open("w", encoding="utf-8") as f:
        json.dump(
            {
                "roi": roi,
                "count": len(merged_candidates),
                "ui_candidates_count": len(ui_candidates),
                "ocr_candidates_count": len(ocr_candidates),
                "ocr": ocr_meta,
                "candidates": merged_candidates,
                "ui_candidates": ui_candidates,
                "ocr_candidates": ocr_candidates,
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    return [roi_file.relative_to(repo_root).as_posix()]


def _resolve_repo_path(repo_root: Path, value: str) -> Path | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    candidate = Path(os.path.expandvars(raw)).expanduser()
    if not candidate.is_absolute():
        candidate = repo_root / candidate
    try:
        resolved = candidate.resolve()
        resolved.relative_to(repo_root.resolve())
        return resolved
    except Exception:
        return None


def _list_template_images(repo_root: Path) -> list[dict]:
    images: list[dict] = []
    for root in [repo_root / "assets"]:
        if not root.exists():
            continue
        for path in root.rglob("*"):
            if not path.is_file() or path.suffix.lower() not in _TEMPLATE_IMAGE_EXTS:
                continue
            try:
                rel = path.relative_to(repo_root).as_posix()
                stat = path.stat()
                images.append(
                    {
                        "path": rel,
                        "name": path.name,
                        "size": int(stat.st_size),
                        "updated_at": int(stat.st_mtime),
                        "preview_url": f"/api/template-image?path={rel}",
                    }
                )
            except Exception:
                continue
    images.sort(key=lambda x: (x.get("updated_at", 0), x.get("path", "")), reverse=True)
    return images[:500]


def _plan_from_steps_payload(steps_payload: list, goal: str = "UI editor flow") -> ExecutionPlan:
    if not isinstance(steps_payload, list) or not steps_payload:
        raise ValueError("steps is required")
    plan_dict = {
        "goal": str(goal or "UI editor flow"),
        "steps": steps_payload,
    }
    return ExecutionPlan.model_validate(plan_dict)


def _resolve_template_paths(steps_payload: list, repo_root: Path) -> list:
    """Return a copy of steps_payload with relative template_path values resolved to absolute."""
    resolved: list = []
    for step in steps_payload:
        if not isinstance(step, dict):
            resolved.append(step)
            continue
        if step.get("action") == "click_image":
            raw = str(step.get("template_path", "")).strip()
            if raw:
                p = Path(os.path.expandvars(raw)).expanduser()
                if not p.is_absolute():
                    p = (repo_root / p).resolve()
                step = {**step, "template_path": str(p)}
        resolved.append(step)
    return resolved


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/")
@app.route("/dashboard")
def dashboard():
    return render_template("dashboard.html", active="dashboard")


@app.route("/editor")
@app.route("/editor/<tree_id>")
def editor(tree_id: str = "new"):
    return render_template("editor.html", active="editor", tree_id=tree_id)


@app.route("/library")
def library():
    return render_template("library.html", active="library")


@app.route("/workspace")
@app.route("/settings")
def workspace():
    return render_template("workspace.html", active="workspace")


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------

@app.route("/api/run", methods=["POST"])
def api_run():
    """Execute an automation task from user input."""
    data = request.get_json(force=True, silent=True) or {}
    user_input: str = str(data.get("user_input", "")).strip()
    dry_run: bool = bool(data.get("dry_run", False))

    if not user_input:
        return jsonify({"ok": False, "error": "user_input is required"}), 400

    _append_log({"type": "input", "message": f"Received: {user_input}"})

    try:
        planner = RuleBasedPlanner()
        plan = planner.build_plan(user_input)
        _append_log({"type": "plan", "message": plan.model_dump_json()})

        if dry_run:
            _append_log({"type": "result", "message": "dry-run: skipped execution"})
            return jsonify({"ok": True, "dry_run": True, "plan": plan.model_dump()})

        monitor = LocalLLMScreenMonitor(enabled=False)
        executor = DeterministicExecutor(monitor=monitor)
        executor.execute(plan)
        _append_log({"type": "result", "message": "execution succeeded"})
        return jsonify({"ok": True, "plan": plan.model_dump()})

    except Exception as exc:
        _append_log({"type": "error", "message": str(exc)})
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/logs")
def api_logs():
    with _log_lock:
        return jsonify(list(_execution_log))


@app.route("/api/plan", methods=["POST"])
def api_plan():
    """Preview plan without executing."""
    data = request.get_json(force=True, silent=True) or {}
    user_input: str = str(data.get("user_input", "")).strip()
    if not user_input:
        return jsonify({"ok": False, "error": "user_input is required"}), 400
    try:
        planner = RuleBasedPlanner()
        plan = planner.build_plan(user_input)
        return jsonify({"ok": True, "plan": plan.model_dump()})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/plan-direct", methods=["POST"])
def api_plan_direct():
    """Preview plan from explicit steps payload (bypass text planner)."""
    data = request.get_json(force=True, silent=True) or {}
    steps_payload = data.get("steps", [])
    goal = str(data.get("goal", "UI editor flow")).strip() or "UI editor flow"
    try:
        repo_root = Path(__file__).resolve().parents[1]
        steps_payload = _resolve_template_paths(steps_payload, repo_root)
        plan = _plan_from_steps_payload(steps_payload, goal=goal)
        return jsonify({"ok": True, "plan": plan.model_dump()})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/run-direct", methods=["POST"])
def api_run_direct():
    """Execute explicit steps payload directly (bypass text planner)."""
    data = request.get_json(force=True, silent=True) or {}
    steps_payload = data.get("steps", [])
    goal = str(data.get("goal", "UI editor flow")).strip() or "UI editor flow"
    dry_run: bool = bool(data.get("dry_run", False))

    try:
        repo_root = Path(__file__).resolve().parents[1]
        steps_payload = _resolve_template_paths(steps_payload, repo_root)
        plan = _plan_from_steps_payload(steps_payload, goal=goal)
        _append_log({"type": "plan", "message": plan.model_dump_json()})

        if dry_run:
            _append_log({"type": "result", "message": "dry-run: skipped execution"})
            return jsonify({"ok": True, "dry_run": True, "plan": plan.model_dump()})

        monitor = LocalLLMScreenMonitor(enabled=False)
        executor = DeterministicExecutor(monitor=monitor)
        executor.execute(plan)
        _append_log({"type": "result", "message": "execution succeeded"})
        return jsonify({"ok": True, "plan": plan.model_dump()})
    except Exception as exc:
        _append_log({"type": "error", "message": str(exc)})
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.route("/api/ui-tree", methods=["GET"])
def api_ui_tree():
    """Collect UI tree. mode=notepad (default) or mode=foreground (active window)."""
    try:
        repo_root = Path(__file__).resolve().parents[1]
        ui_tree_script = repo_root / "ui_tree.py"
        if not ui_tree_script.exists():
            return jsonify({"ok": False, "error": "ui_tree.py not found"}), 404

        mode = request.args.get("mode", "notepad")
        persist_raw = str(request.args.get("persist", "1")).lower()
        should_persist = persist_raw not in {"0", "false", "no"}

        if mode == "roi":
            roi = ui_tree2.select_screen_roi()
            if roi is None:
                return jsonify({"ok": False, "error": "roi selection cancelled"}), 400

            raw_tree = ui_tree2.collect_window_tree_by_screen_point(roi.center_x, roi.center_y, max_depth=6, backend="uia")
            roi_dict = roi.to_dict()
            ui_candidates = _collect_roi_candidates(raw_tree, roi_dict)
            backend_used = "uia"

            if _container_only(ui_candidates):
                raw_tree_win32 = ui_tree2.collect_window_tree_by_screen_point(roi.center_x, roi.center_y, max_depth=6, backend="win32")
                candidates_win32 = _collect_roi_candidates(raw_tree_win32, roi_dict)
                if _candidate_quality(candidates_win32) > _candidate_quality(ui_candidates):
                    raw_tree = raw_tree_win32
                    ui_candidates = candidates_win32
                    backend_used = "win32"

            roi_tree = _filter_tree_to_roi(raw_tree, roi_dict)
            payload_source = roi_tree if roi_tree is not None else raw_tree
            payload = ui_tree2._simplify_tree_dict(payload_source)

            roi_image = None
            roi_image_path: str | None = None
            try:
                roi_image = _capture_roi_image(roi_dict)
            except Exception:
                roi_image = None

            if should_persist and roi_image is not None:
                try:
                    roi_image_path = _persist_roi_image(repo_root, str(payload.get("name") or "unknown_app"), roi_image)
                except Exception:
                    roi_image_path = None

            ocr_candidates, ocr_meta = _collect_ocr_candidates(
                roi_dict,
                image=roi_image,
                source_image_path=roi_image_path,
            )
            merged_candidates = _merge_candidates(ui_candidates, ocr_candidates)

            hint_lines: list[str] = []
            if not merged_candidates:
                hint_lines.append("ROI 區域內找不到可用控制項，請框選更大的內容區。")
            if backend_used == "win32":
                hint_lines.append("UIA 控制項不足，已自動改用 win32 後端擷取。")
            if ocr_meta.get("enabled"):
                hint_lines.append(f"OCR 已啟用，偵測到 {ocr_meta.get('kept', 0)} 筆文字候選。")
                if ocr_meta.get("source_image_path"):
                    hint_lines.append(f"OCR 來源圖片：{ocr_meta.get('source_image_path')}")
            elif ocr_meta.get("error"):
                hint_lines.append(f"OCR 未啟用：{ocr_meta.get('error')}")
            hint = "\n".join(hint_lines)

            artifacts: list[str] = []
            if roi_image_path:
                artifacts.append(roi_image_path)
            if should_persist:
                artifacts.extend(_persist_ui_tree(repo_root, payload))
                artifacts.extend(
                    _persist_roi_artifacts(
                        repo_root,
                        str(payload.get("name") or "unknown_app"),
                        roi_dict,
                        ui_candidates,
                        ocr_candidates,
                        merged_candidates,
                        ocr_meta,
                    )
                )

            return jsonify({
                "ok": True,
                "mode": "roi",
                "roi": roi_dict,
                "backend": backend_used,
                "tree": payload,
                "ocr": ocr_meta,
                "candidates": merged_candidates[:80],
                "ui_candidates": ui_candidates[:50],
                "ocr_candidates": ocr_candidates[:50],
                "hint": hint,
                "artifacts": artifacts,
            })

        flag = "--foreground" if mode == "foreground" else "--json"

        proc = subprocess.run(
            [sys.executable, str(ui_tree_script), flag],
            capture_output=True,
            text=True,
            timeout=30,
            cwd=str(repo_root),
        )

        if proc.returncode != 0:
            return jsonify(
                {
                    "ok": False,
                    "error": "ui_tree command failed",
                    "details": (proc.stderr or proc.stdout).strip(),
                }
            ), 500

        payload = json.loads(proc.stdout)
        artifacts = _persist_ui_tree(repo_root, payload) if should_persist else []
        return jsonify({"ok": True, "tree": payload, "artifacts": artifacts})
    except json.JSONDecodeError:
        return jsonify({"ok": False, "error": "invalid ui_tree json output"}), 500
    except subprocess.TimeoutExpired:
        return jsonify({"ok": False, "error": "ui_tree command timeout"}), 504
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/template-images", methods=["GET"])
def api_template_images():
    try:
        repo_root = Path(__file__).resolve().parents[1]
        images = _list_template_images(repo_root)
        return jsonify({"ok": True, "images": images})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/template-image", methods=["GET"])
def api_template_image():
    try:
        repo_root = Path(__file__).resolve().parents[1]
        rel_path = str(request.args.get("path", "")).strip()
        resolved = _resolve_repo_path(repo_root, rel_path)
        if resolved is None:
            return jsonify({"ok": False, "error": "invalid template image path"}), 400
        if not resolved.exists() or not resolved.is_file():
            return jsonify({"ok": False, "error": "template image not found"}), 404
        if resolved.suffix.lower() not in _TEMPLATE_IMAGE_EXTS:
            return jsonify({"ok": False, "error": "unsupported image type"}), 400
        return send_file(resolved)
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/api/capture-image-roi", methods=["POST"])
def api_capture_image_roi():
    try:
        repo_root = Path(__file__).resolve().parents[1]
        roi = ui_tree2.select_screen_roi()
        if roi is None:
            return jsonify({"ok": False, "error": "roi selection cancelled"}), 400

        roi_dict = roi.to_dict()
        image = _capture_roi_image(roi_dict)

        out_dir = repo_root / "assets" / "templates"
        out_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_file = out_dir / f"{ts}.png"
        image.save(out_file)

        rel = out_file.relative_to(repo_root).as_posix()
        return jsonify({
            "ok": True,
            "path": rel,
            "roi": roi_dict,
            "preview_url": f"/api/template-image?path={rel}",
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


def run(host: str = "127.0.0.1", port: int = 5000, debug: bool = False) -> None:
    app.run(host=host, port=port, debug=debug)
