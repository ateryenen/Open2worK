import os
import subprocess
import time
from pathlib import Path
from typing import Any, Optional

import pyautogui
import pyscreeze
from pywinauto import Application, Desktop

from .execution_tree import ExecutionTreeCache
from .schemas import ExecutionPlan
from .utils import log


class DeterministicExecutor:
    def __init__(self, monitor: Any = None) -> None:
        self.app: Optional[Application] = None
        self.window = None
        self.monitor = monitor
        self._last_typed_text = ""
        self.execution_tree = ExecutionTreeCache()

    def execute(self, plan: ExecutionPlan) -> None:
        for idx, step in enumerate(plan.steps, start=1):
            log(f"executing step {idx}: {step.action}")
            self._emit_monitor_event("before_step", idx, step)
            try:
                if step.action == "open_app":
                    self._open_app(step.target)
                elif step.action == "wait":
                    time.sleep(step.seconds)
                elif step.action == "type_text":
                    self._type_text(step.text)
                elif step.action == "save_file":
                    self._save_file(step.path)
                elif step.action == "click_image":
                    self._click_image(
                        template_path=step.template_path,
                        confidence=step.confidence,
                        timeout_seconds=step.timeout_seconds,
                        click_count=step.click_count,
                    )
            except Exception as exc:
                self._emit_monitor_event("error_step", idx, step, {"error": str(exc)})
                raise
            self._emit_monitor_event("after_step", idx, step)
            log(f"step {idx} completed")

    def _emit_monitor_event(self, event: str, idx: int, step: Any, extra: Optional[dict] = None) -> None:
        if self.monitor is None or not hasattr(self.monitor, "observe_step"):
            return
        payload = {
            "index": idx,
            "action": getattr(step, "action", ""),
        }
        if hasattr(step, "target"):
            payload["target"] = step.target
        if hasattr(step, "text"):
            payload["text_preview"] = str(step.text)[:80]
        if hasattr(step, "path"):
            payload["path"] = step.path
        if hasattr(step, "seconds"):
            payload["seconds"] = step.seconds
        if extra:
            payload.update(extra)
        self.monitor.observe_step(event, payload)

    def is_notepad_running(self) -> bool:
        return self.window is not None

    def _open_app(self, target: str) -> None:
        if target.lower() != "notepad":
            raise ValueError("Only notepad is supported in this POC")

        desktop = Desktop(backend="uia")
        before_handles = {w.handle for w in desktop.windows(control_type="Window")}

        subprocess.Popen(["notepad.exe"])
        time.sleep(1.0)

        selected_window = None
        for _ in range(30):
            windows = desktop.windows(control_type="Window")
            for window in windows:
                try:
                    title = window.window_text().strip().lower()
                    cls = window.class_name()
                    is_new = window.handle not in before_handles
                    is_notepad_like = cls == "Notepad" or "notepad" in title or "記事本" in title
                    if is_new and is_notepad_like:
                        selected_window = window
                        break
                except Exception:
                    continue
            if selected_window is not None:
                break
            time.sleep(0.2)

        if selected_window is None:
            for window in desktop.windows(control_type="Window"):
                try:
                    title = window.window_text().strip().lower()
                    cls = window.class_name()
                    if cls == "Notepad" or "notepad" in title or "記事本" in title:
                        selected_window = window
                        break
                except Exception:
                    continue

        if selected_window is None:
            raise RuntimeError("Failed to locate Notepad window")

        try:
            self.app = Application(backend="uia").connect(process=selected_window.process_id(), timeout=3)
            self.window = self.app.window(handle=selected_window.handle)
            self.window.wait("visible", timeout=6)
            self.window.set_focus()
            return
        except Exception as e:
            raise RuntimeError(f"Failed to bind Notepad window: {e}")

    def _type_text(self, text: str) -> None:
        if self.window is None:
            raise RuntimeError("Notepad window is not ready")
        self._last_typed_text = text
        self.window.set_focus()
        self.window.type_keys("^a{BACKSPACE}", pause=0.01)
        time.sleep(0.1)
        self.window.type_keys(text, with_spaces=True, pause=0.01)

    def _save_file(self, target_path: str) -> None:
        if self.window is None:
            raise RuntimeError("Notepad window is not ready")

        expanded = os.path.expandvars(target_path)
        normalized = str(Path(expanded))

        for attempt in range(1, 3):
            dialog = self._find_save_dialog(timeout_seconds=0.6)
            if dialog is None:
                try:
                    self.window.set_focus()
                    self.window.type_keys("{F12}")
                    time.sleep(0.8)
                except Exception:
                    pass

            dialog = self._find_save_dialog(timeout_seconds=3.0)
            if dialog is None:
                try:
                    self.window.set_focus()
                    self.window.type_keys("^s")
                    time.sleep(0.8)
                except Exception:
                    pass
                dialog = self._find_save_dialog(timeout_seconds=2.0)

            if dialog is None:
                log(f"save dialog not found ({attempt}/2)")
                continue

            if not self._fill_save_filename(dialog, normalized):
                log(f"save filename fill failed ({attempt}/2)")
                continue

            time.sleep(0.8)
            self._handle_overwrite_if_present()
            time.sleep(0.3)

            if Path(normalized).exists():
                return

            log(f"save retry triggered ({attempt}/2)")

        try:
            log("save dialog automation failed, fallback to direct file write")
            Path(normalized).write_text(self._last_typed_text, encoding="utf-8")
            return
        except Exception as e:
            raise RuntimeError(f"Save file failed after retries: {normalized}; fallback failed: {e}")

    def _fill_save_filename(self, dialog, normalized: str) -> bool:
        try:
            dialog.set_focus()
        except Exception:
            pass

        edit_candidates = []
        for spec in [
            {"auto_id": "1001", "control_type": "Edit"},
            {"title_re": ".*(File name|檔案名稱).*", "control_type": "Edit"},
            {"control_type": "Edit"},
        ]:
            try:
                edit = dialog.child_window(**spec)
                if edit.exists(timeout=0.2):
                    edit_candidates.append(edit)
            except Exception:
                continue

        for edit in edit_candidates:
            try:
                edit.set_focus()
                edit.set_edit_text(normalized)
                dialog.type_keys("{ENTER}")
                return True
            except Exception:
                continue

        return False

    def _handle_overwrite_if_present(self) -> None:
        dialog = self._find_save_dialog(timeout_seconds=1.2)
        if dialog is None:
            return
        try:
            title = dialog.window_text().lower()
            if "confirm" in title or "replace" in title or "確認" in title or "取代" in title:
                dialog.type_keys("{LEFT}{ENTER}")
        except Exception:
            pass

    def _find_save_dialog(self, timeout_seconds: float):
        if self.app is None:
            return None

        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            try:
                active = Desktop(backend="uia").get_active()
                if active is not None:
                    if self.window is not None and active.handle == self.window.handle:
                        pass
                    else:
                        try:
                            title = active.window_text().lower()
                        except Exception:
                            title = ""
                        try:
                            has_edit = active.child_window(control_type="Edit").exists(timeout=0.1)
                        except Exception:
                            has_edit = False
                        if has_edit or "save" in title or "另存" in title or "儲存" in title:
                            return active
            except Exception:
                pass

            for dlg in Desktop(backend="uia").windows():
                if self.window is not None and dlg.handle == self.window.handle:
                    continue
                try:
                    title = dlg.window_text().lower()
                    has_save_keyword = (
                        "save" in title or "另存" in title or "儲存" in title or "confirm" in title
                    )
                    if has_save_keyword:
                        return dlg
                    edit = dlg.child_window(control_type="Edit")
                    if edit.exists(timeout=0.1):
                        return dlg
                except Exception:
                    continue
            time.sleep(0.15)

        return None

    def _click_image(
        self,
        template_path: str,
        confidence: float,
        timeout_seconds: float,
        click_count: int,
    ) -> None:
        template = Path(os.path.expandvars(template_path)).expanduser()
        if not template.exists():
            raise FileNotFoundError(f"Template image not found: {template}")

        object_name = template.stem

        cached_point = self.execution_tree.get_cached_point(str(template), object_name)
        if cached_point is not None:
            region = self._region_around(cached_point[0], cached_point[1], width=260, height=260)
            point = self._locate_image_center(template, confidence, region=region)
            if point is not None:
                log(f"execution-tree cache hit: {object_name} @({point.x},{point.y})")
                self._click_point(point.x, point.y, click_count)
                self.execution_tree.update_hit(
                    template_path=str(template),
                    object_name=object_name,
                    x=point.x,
                    y=point.y,
                    confidence=confidence,
                    source="cache-region",
                )
                return
            log(f"execution-tree cache miss: {object_name}, fallback to full scan")

        deadline = time.time() + timeout_seconds
        last_error = None

        while time.time() < deadline:
            try:
                point = self._locate_image_center(template, confidence)
            except Exception as exc:
                last_error = exc
                point = None

            if point is not None:
                self._click_point(point.x, point.y, click_count)
                self.execution_tree.update_hit(
                    template_path=str(template),
                    object_name=object_name,
                    x=point.x,
                    y=point.y,
                    confidence=confidence,
                    source="full-scan",
                )
                return

            time.sleep(0.3)

        self.execution_tree.record_miss(str(template), object_name)

        if last_error is not None:
            raise RuntimeError(
                f"Image locate timeout: {template} (last matcher error: {last_error})"
            )
        raise RuntimeError(f"Image locate timeout: {template}")

    def _locate_image_center(self, template: Path, confidence: float, region: tuple[int, int, int, int] | None = None):
        try:
            kwargs = {"confidence": confidence}
            if region is not None:
                kwargs["region"] = region
            return pyautogui.locateCenterOnScreen(str(template), **kwargs)
        except NotImplementedError:
            try:
                if region is not None:
                    return pyautogui.locateCenterOnScreen(str(template), region=region)
                return pyautogui.locateCenterOnScreen(str(template))
            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                return None
        except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
            return None

    def _click_point(self, x: int, y: int, click_count: int) -> None:
        pyautogui.moveTo(x, y, duration=0.1)
        for _ in range(click_count):
            pyautogui.click(x, y)
            time.sleep(0.08)

    def _region_around(self, x: int, y: int, width: int, height: int) -> tuple[int, int, int, int]:
        screen_w, screen_h = pyautogui.size()
        left = max(0, int(x - width // 2))
        top = max(0, int(y - height // 2))
        max_w = max(1, int(screen_w - left))
        max_h = max(1, int(screen_h - top))
        return (left, top, min(width, max_w), min(height, max_h))
