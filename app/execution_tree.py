import json
import os
import time
from pathlib import Path

import pyautogui


class ExecutionTreeCache:
    def __init__(self, cache_path: str = "C:/GITHUB/Open2worK/.cache/execution_tree.json") -> None:
        self.cache_path = Path(os.path.expandvars(cache_path)).expanduser()
        self.cache_path.parent.mkdir(parents=True, exist_ok=True)
        self._data = self._load()

    def _load(self) -> dict:
        if not self.cache_path.exists():
            return {"version": 1, "objects": {}}
        try:
            content = json.loads(self.cache_path.read_text(encoding="utf-8"))
            if not isinstance(content, dict):
                return {"version": 1, "objects": {}}
            if "objects" not in content or not isinstance(content.get("objects"), dict):
                content["objects"] = {}
            return content
        except Exception:
            return {"version": 1, "objects": {}}

    def _save(self) -> None:
        self.cache_path.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _screen_key(self) -> str:
        size = pyautogui.size()
        return f"{size.width}x{size.height}"

    def _object_key(self, template_path: str, object_name: str) -> str:
        normalized_template = str(Path(template_path)).replace("\\", "/").lower()
        normalized_name = object_name.strip().lower()
        return f"{normalized_name}|{normalized_template}"

    def get_cached_point(self, template_path: str, object_name: str) -> tuple[int, int] | None:
        key = self._object_key(template_path, object_name)
        obj = self._data.get("objects", {}).get(key)
        if not isinstance(obj, dict):
            return None
        if obj.get("screen") != self._screen_key():
            return None
        x = obj.get("x")
        y = obj.get("y")
        if isinstance(x, int) and isinstance(y, int):
            return (x, y)
        return None

    def update_hit(
        self,
        template_path: str,
        object_name: str,
        x: int,
        y: int,
        confidence: float,
        source: str,
    ) -> None:
        key = self._object_key(template_path, object_name)
        self._data.setdefault("objects", {})[key] = {
            "name": object_name,
            "template_path": str(Path(template_path)),
            "x": int(x),
            "y": int(y),
            "screen": self._screen_key(),
            "confidence": float(confidence),
            "source": source,
            "miss_count": 0,
            "last_seen": time.time(),
        }
        self._save()

    def record_miss(self, template_path: str, object_name: str) -> None:
        key = self._object_key(template_path, object_name)
        obj = self._data.setdefault("objects", {}).get(key)
        if not isinstance(obj, dict):
            return
        misses = int(obj.get("miss_count", 0)) + 1
        obj["miss_count"] = misses
        obj["last_miss"] = time.time()
        if misses >= 3:
            self._data["objects"].pop(key, None)
        self._save()
