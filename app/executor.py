import os
import subprocess
import time
import ctypes
from pathlib import Path
from typing import Any, Optional

import pyautogui
import pyscreeze
from pywinauto import Application, Desktop

from .execution_tree import ExecutionTreeCache
from .schemas import ExecutionPlan
from .utils import log


_MOUSEEVENTF_LEFTDOWN = 0x0002
_MOUSEEVENTF_LEFTUP = 0x0004
_MOUSEEVENTF_RIGHTDOWN = 0x0008
_MOUSEEVENTF_RIGHTUP = 0x0010

# SendInput 結構（繼承者 API，對 UIPI 境界相容性更佳）
_INPUT_MOUSE = 0


class _MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx",          ctypes.c_long),
        ("dy",          ctypes.c_long),
        ("mouseData",   ctypes.c_ulong),
        ("dwFlags",     ctypes.c_ulong),
        ("time",        ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_size_t),  # ULONG_PTR
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", _MOUSEINPUT)]


class _INPUT(ctypes.Structure):
    _fields_ = [
        ("type",   ctypes.c_ulong),
        ("_input", _INPUT_UNION),
    ]


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
                elif step.action == "key_press":
                    self._key_press(
                        key=step.key,
                        repeat=step.repeat,
                        interval=step.interval,
                    )
                elif step.action == "save_file":
                    self._save_file(step.path)
                elif step.action == "click_image":
                    self._click_image(
                        template_path=step.template_path,
                        confidence=step.confidence,
                        timeout_seconds=step.timeout_seconds,
                        click_count=step.click_count,
                        press_duration=step.press_duration,
                        click_mode=step.click_mode,
                    )
                elif step.action == "move_mouse_horizontal":
                    self._move_mouse_horizontal(
                        direction=step.direction,
                        pixels=step.pixels,
                        duration=step.duration,
                    )
                elif step.action == "mouse_click":
                    self._mouse_click(
                        button=step.button,
                        click_count=step.click_count,
                        interval=step.interval,
                        press_duration=step.press_duration,
                        click_mode=step.click_mode,
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
        if hasattr(step, "key"):
            payload["key"] = step.key
        if hasattr(step, "repeat"):
            payload["repeat"] = step.repeat
        if hasattr(step, "path"):
            payload["path"] = step.path
        if hasattr(step, "seconds"):
            payload["seconds"] = step.seconds
        if hasattr(step, "direction"):
            payload["direction"] = step.direction
        if hasattr(step, "pixels"):
            payload["pixels"] = step.pixels
        if hasattr(step, "duration"):
            payload["duration"] = step.duration
        if hasattr(step, "button"):
            payload["button"] = step.button
        if hasattr(step, "click_count"):
            payload["click_count"] = step.click_count
        if hasattr(step, "interval"):
            payload["interval"] = step.interval
        if hasattr(step, "press_duration"):
            payload["press_duration"] = step.press_duration
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

    def _key_press(self, key: str, repeat: int, interval: float) -> None:
        normalized = str(key or "").strip().lower()
        if not normalized:
            raise ValueError("key_press requires non-empty key")

        count = max(1, int(repeat))
        gap = max(0.0, float(interval))

        if self.window is not None:
            try:
                self.window.set_focus()
            except Exception:
                pass

        for idx in range(count):
            if "+" in normalized:
                combo = [part.strip() for part in normalized.split("+") if part.strip()]
                if not combo:
                    raise ValueError(f"Invalid key combination: {key}")
                pyautogui.hotkey(*combo)
            else:
                pyautogui.press(normalized)

            if idx < count - 1:
                time.sleep(gap)

        log(f"key pressed: {normalized}, repeat={count}, interval={gap}")

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
        press_duration: float,
        click_mode: str = "mouse_event",
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
                self._click_point(point.x, point.y, click_count, press_duration=press_duration, click_mode=click_mode)
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
                self._click_point(point.x, point.y, click_count, press_duration=press_duration, click_mode=click_mode)
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
        confidence_levels = [float(confidence)]
        for fallback in (confidence - 0.05, confidence - 0.1, confidence - 0.15):
            if fallback >= 0.6:
                confidence_levels.append(round(float(fallback), 2))

        seen: set[tuple[float, bool]] = set()
        attempts: list[tuple[float, bool]] = []
        for level in confidence_levels:
            for grayscale in (False, True):
                key = (level, grayscale)
                if key not in seen:
                    seen.add(key)
                    attempts.append(key)

        for level, grayscale in attempts:
            try:
                kwargs = {"confidence": level, "grayscale": grayscale}
                if region is not None:
                    kwargs["region"] = region
                point = pyautogui.locateCenterOnScreen(str(template), **kwargs)
                if point is not None:
                    return point
            except NotImplementedError:
                try:
                    if region is not None:
                        point = pyautogui.locateCenterOnScreen(str(template), region=region, grayscale=grayscale)
                    else:
                        point = pyautogui.locateCenterOnScreen(str(template), grayscale=grayscale)
                    if point is not None:
                        return point
                except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                    continue
            except (pyautogui.ImageNotFoundException, pyscreeze.ImageNotFoundException):
                continue

        return None

    def _click_point(self, x: int, y: int, click_count: int, press_duration: float = 0.03, click_mode: str = "mouse_event") -> None:
        pyautogui.moveTo(x, y, duration=0.12)
        time.sleep(0.05)
        for idx in range(click_count):
            self._dispatch_click("left", press_duration=press_duration, mode=click_mode)
            if idx < click_count - 1:
                time.sleep(0.1)
        pos = pyautogui.position()
        log(f"click dispatched @({pos.x},{pos.y}), count={click_count}, press={press_duration}s, mode={click_mode}")

    def _move_mouse_horizontal(self, direction: str, pixels: int, duration: float) -> None:
        if direction not in {"left", "right"}:
            raise ValueError(f"Unsupported direction: {direction}")
        dx = -int(pixels) if direction == "left" else int(pixels)
        pyautogui.moveRel(dx, 0, duration=max(0.0, float(duration)))
        pos = pyautogui.position()
        log(f"mouse moved {direction} by {pixels}px -> ({pos.x},{pos.y})")

    def _mouse_click(self, button: str, click_count: int, interval: float, press_duration: float, click_mode: str = "mouse_event") -> None:
        if button not in {"left", "right"}:
            raise ValueError(f"Unsupported mouse button: {button}")
        pos = pyautogui.position()
        for idx in range(max(1, int(click_count))):
            self._dispatch_click(button, press_duration=press_duration, mode=click_mode)
            if idx < click_count - 1:
                time.sleep(max(0.0, float(interval)))
        log(f"mouse click dispatched button={button}, count={click_count}, press={press_duration}s, mode={click_mode} @({pos.x},{pos.y})")

    def _dispatch_click(self, button: str, press_duration: float = 0.03, mode: str = "mouse_event") -> None:
        hold = max(0.0, float(press_duration))
        if os.name == "nt":
            down_flag = _MOUSEEVENTF_LEFTDOWN if button == "left" else _MOUSEEVENTF_RIGHTDOWN
            up_flag   = _MOUSEEVENTF_LEFTUP   if button == "left" else _MOUSEEVENTF_RIGHTUP
            if mode == "send_input":
                inp = _INPUT()
                inp.type = _INPUT_MOUSE
                inp._input.mi.dx = 0
                inp._input.mi.dy = 0
                inp._input.mi.mouseData = 0
                inp._input.mi.time = 0
                inp._input.mi.dwExtraInfo = 0
                inp._input.mi.dwFlags = down_flag
                ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))
                time.sleep(hold)
                inp._input.mi.dwFlags = up_flag
                ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(_INPUT))
                return
            else:  # mouse_event（原版保留）
                ctypes.windll.user32.mouse_event(down_flag, 0, 0, 0, 0)
                time.sleep(hold)
                ctypes.windll.user32.mouse_event(up_flag, 0, 0, 0, 0)
                return
        pyautogui.mouseDown(button=button)
        time.sleep(hold)
        pyautogui.mouseUp(button=button)

    def _region_around(self, x: int, y: int, width: int, height: int) -> tuple[int, int, int, int]:
        screen_w, screen_h = pyautogui.size()
        left = max(0, int(x - width // 2))
        top = max(0, int(y - height // 2))
        max_w = max(1, int(screen_w - left))
        max_h = max(1, int(screen_h - top))
        return (left, top, min(width, max_w), min(height, max_h))
