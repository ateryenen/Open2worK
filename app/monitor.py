import ctypes
import json
import threading
import time
import urllib.error
import urllib.request

from pywinauto import Desktop

from .utils import log


class LocalLLMScreenMonitor:
    def __init__(
        self,
        enabled: bool = False,
        interval_seconds: float = 2.0,
        model: str = "phi:latest",
        endpoint: str = "http://127.0.0.1:11434/api/generate",
        prompt: str = "You are a desktop automation monitor. Analyze JSON state and return one concise line: status + next-risk."
    ) -> None:
        self.enabled = enabled
        self.interval_seconds = max(1.0, interval_seconds)
        self.model = model
        self.endpoint = endpoint
        self.prompt = prompt
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._last_warning = ""
        self._last_warning_at = 0.0

    def start(self) -> None:
        if not self.enabled:
            return
        if self._thread is not None and self._thread.is_alive():
            return

        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        log(f"monitor started (model={self.model}, interval={self.interval_seconds}s)")

    def stop(self) -> None:
        if not self.enabled:
            return
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
        log("monitor stopped")

    def _run(self) -> None:
        while not self._stop_event.is_set():
            try:
                state = self._collect_screen_state()
                observation = self._query_ollama(state)
                if observation:
                    log(f"monitor observation: {observation[:180]}")
            except Exception as e:
                self._warn(str(e))
            self._stop_event.wait(self.interval_seconds)

    def observe_step(self, event: str, payload: dict) -> None:
        if not self.enabled:
            return
        try:
            state = self._collect_screen_state()
            state["step_event"] = event
            state["step_payload"] = payload
            observation = self._query_ollama(state)
            if observation:
                log(f"monitor step {event}: {observation[:180]}")
        except Exception as e:
            self._warn(f"step hook failed: {e}")

    def _warn(self, message: str) -> None:
        now = time.time()
        if message != self._last_warning or (now - self._last_warning_at) >= 10:
            log(f"monitor warning: {message}")
            self._last_warning = message
            self._last_warning_at = now

    def _collect_screen_state(self) -> dict:
        windows = []
        for win in Desktop(backend="uia").windows():
            try:
                title = win.window_text().strip()
                if title:
                    windows.append(title)
            except Exception:
                continue

        foreground_title = self._get_foreground_window_title()
        lowered = [title.lower() for title in windows]

        notepad_detected = any("notepad" in title or "記事本" in title for title in lowered)
        save_dialog_detected = any(
            "save as" in title or "另存" in title or "儲存" in title for title in lowered
        )

        return {
            "timestamp": time.time(),
            "foreground_window": foreground_title,
            "window_count": len(windows),
            "top_window_samples": windows[:8],
            "notepad_detected": notepad_detected,
            "save_dialog_detected": save_dialog_detected,
        }

    def _get_foreground_window_title(self) -> str:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return ""
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value

    def _query_ollama(self, state: dict) -> str:
        state_json = json.dumps(state, ensure_ascii=False)
        composed_prompt = (
            f"{self.prompt}\n"
            "Return format: <status>; <risk>; <next_hint>.\n"
            f"State JSON:\n{state_json}"
        )

        payload = {
            "model": self.model,
            "prompt": composed_prompt,
            "stream": False,
        }
        req = urllib.request.Request(
            url=self.endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as response:
                body = response.read().decode("utf-8")
            data = json.loads(body)
            return str(data.get("response", "")).strip()
        except urllib.error.HTTPError as e:
            if e.code == 404 and self.endpoint.endswith("/api/generate"):
                return self._query_ollama_chat(composed_prompt)
            raise RuntimeError(f"ollama endpoint error ({e.code}) at {self.endpoint}")
        except urllib.error.URLError as e:
            raise RuntimeError(f"cannot reach ollama at {self.endpoint}: {e}")

    def _query_ollama_chat(self, composed_prompt: str) -> str:
        chat_endpoint = self.endpoint.replace("/api/generate", "/api/chat")
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "user",
                    "content": composed_prompt,
                }
            ],
            "stream": False,
        }
        req = urllib.request.Request(
            url=chat_endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as response:
                body = response.read().decode("utf-8")
            data = json.loads(body)
            message = data.get("message", {})
            return str(message.get("content", "")).strip()
        except urllib.error.HTTPError as e:
            raise RuntimeError(
                f"ollama endpoint not supported: tried {self.endpoint} and {chat_endpoint} ({e.code})"
            )
