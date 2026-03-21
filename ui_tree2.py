import argparse
import json
import subprocess
import sys
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

from pywinauto import Application, Desktop
from pywinauto.base_wrapper import BaseWrapper


NOTEPAD_MENU_ALIASES = {
    "檔案": "File",
    "文件": "File",
    "編輯": "Edit",
    "格式": "Format",
    "檢視": "View",
    "檢視(&V)": "View",
    "說明": "Help",
    "File": "File",
    "Edit": "Edit",
    "Format": "Format",
    "View": "View",
    "Help": "Help",
}

PREFERRED_MENUS = ["File", "Edit", "Format", "View"]


@dataclass
class ScreenROI:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def center_x(self) -> int:
        return int((self.left + self.right) / 2)

    @property
    def center_y(self) -> int:
        return int((self.top + self.bottom) / 2)

    def to_dict(self) -> Dict[str, int]:
        return {
            "left": self.left,
            "top": self.top,
            "right": self.right,
            "bottom": self.bottom,
            "width": self.right - self.left,
            "height": self.bottom - self.top,
        }


def safe_get(func, default=None):
    try:
        return func()
    except Exception:
        return default


def rect_to_dict(wrapper: BaseWrapper) -> Optional[Dict[str, int]]:
    rect = safe_get(wrapper.rectangle)
    if not rect:
        return None
    return {
        "left": rect.left,
        "top": rect.top,
        "right": rect.right,
        "bottom": rect.bottom,
        "width": rect.width(),
        "height": rect.height(),
    }


def get_element_info(wrapper: BaseWrapper) -> Dict[str, Any]:
    element_info = safe_get(lambda: wrapper.element_info, None)

    info = {
        "name": safe_get(wrapper.window_text, ""),
        "control_type": safe_get(lambda: element_info.control_type if element_info else None, None),
        "class_name": safe_get(wrapper.class_name, None),
        "automation_id": safe_get(lambda: element_info.automation_id if element_info else None, None),
        "handle": safe_get(wrapper.handle, None),
        "rectangle": rect_to_dict(wrapper),
        "is_visible": safe_get(wrapper.is_visible, None),
        "is_enabled": safe_get(wrapper.is_enabled, None),
        "children": [],
    }
    return info


def print_tree(
    wrapper: BaseWrapper,
    depth: int = 0,
    max_depth: int = 6,
    show_rect: bool = False,
) -> None:
    if depth > max_depth:
        return

    info = get_element_info(wrapper)
    indent = "  " * depth

    name = info["name"] or ""
    control_type = info["control_type"] or "Unknown"
    class_name = info["class_name"] or ""
    automation_id = info["automation_id"] or ""

    line = f"{indent}- {control_type}"
    if name:
        line += f' | name="{name}"'
    if class_name:
        line += f' | class="{class_name}"'
    if automation_id:
        line += f' | auto_id="{automation_id}"'

    if show_rect and info["rectangle"]:
        r = info["rectangle"]
        line += f' | rect=({r["left"]},{r["top"]},{r["right"]},{r["bottom"]})'

    print(line)

    children = safe_get(wrapper.children, [])
    if not children:
        return

    for child in children:
        print_tree(child, depth + 1, max_depth=max_depth, show_rect=show_rect)


def build_tree_dict(
    wrapper: BaseWrapper,
    depth: int = 0,
    max_depth: int = 6,
) -> Dict[str, Any]:
    info = get_element_info(wrapper)

    if depth >= max_depth:
        return info

    children = safe_get(wrapper.children, [])
    result_children: List[Dict[str, Any]] = []

    for child in children:
        result_children.append(build_tree_dict(child, depth + 1, max_depth=max_depth))

    info["children"] = result_children
    return info


def find_window_by_title(title_re: str, timeout: float = 10.0) -> BaseWrapper:
    deadline = time.time() + timeout
    last_error = None

    while time.time() < deadline:
        try:
            win = Desktop(backend="uia").window(title_re=title_re)
            win.wait("exists ready", timeout=1)
            return win
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)

    raise RuntimeError(f"找不到符合 title_re 的視窗: {title_re}\n最後錯誤: {last_error}")


def launch_app(command: str, wait_time: float = 2.0) -> None:
    subprocess.Popen(command, shell=True)
    time.sleep(wait_time)


def connect_by_exe(exe_name: str, timeout: float = 10.0) -> BaseWrapper:
    deadline = time.time() + timeout
    last_error = None

    while time.time() < deadline:
        try:
            app = Application(backend="uia").connect(path=exe_name)
            win = app.top_window()
            win.wait("exists ready", timeout=1)
            return win
        except Exception as exc:
            last_error = exc
            time.sleep(0.5)

    raise RuntimeError(f"無法附著到程式: {exe_name}\n最後錯誤: {last_error}")


def save_json(data: Dict[str, Any], output_path: str) -> None:
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _find_notepad_window(timeout_seconds: float = 5.0) -> Optional[BaseWrapper]:
    deadline = time.time() + timeout_seconds
    desktop = Desktop(backend="uia")

    while time.time() < deadline:
        windows = desktop.windows(control_type="Window")
        for win in windows:
            title = safe_get(lambda: (win.window_text() or "").lower(), "")
            class_name = safe_get(lambda: (win.class_name() or "").lower(), "")
            if "notepad" in title or "記事本" in title or class_name == "notepad":
                return win
        time.sleep(0.2)

    return None


def ensure_notepad_window() -> BaseWrapper:
    win = _find_notepad_window(timeout_seconds=1.5)
    if win is not None:
        return win

    subprocess.Popen(["notepad.exe"])
    win = _find_notepad_window(timeout_seconds=5.0)
    if win is None:
        raise RuntimeError("Cannot find Notepad window. Please open Notepad and retry.")
    return win


def extract_notepad_menu_items(notepad_win: BaseWrapper) -> List[str]:
    try:
        menubar = notepad_win.child_window(control_type="MenuBar").wrapper_object()
        raw_items = [item.window_text().strip() for item in menubar.children() if item.window_text().strip()]
    except Exception:
        raw_items = []

    normalized: List[str] = []
    for item in raw_items:
        if item in NOTEPAD_MENU_ALIASES:
            name = NOTEPAD_MENU_ALIASES[item]
            if name not in normalized:
                normalized.append(name)

    for preferred in PREFERRED_MENUS:
        if preferred not in normalized:
            normalized.append(preferred)

    return normalized


def has_notepad_edit_area(notepad_win: BaseWrapper) -> bool:
    try:
        if notepad_win.child_window(control_type="Edit").exists(timeout=0.2):
            return True
    except Exception:
        pass

    try:
        if notepad_win.child_window(control_type="Document").exists(timeout=0.2):
            return True
    except Exception:
        pass

    return True


def collect_notepad_simple_tree() -> Dict[str, Any]:
    notepad_win = ensure_notepad_window()
    menus = extract_notepad_menu_items(notepad_win)
    has_edit_area = has_notepad_edit_area(notepad_win)

    tree: Dict[str, Any] = {
        "name": "Notepad",
        "children": [
            {
                "name": "MenuBar",
                "children": [{"name": menu} for menu in menus],
            }
        ],
    }

    if has_edit_area:
        tree["children"].append({"name": "EditArea"})

    return tree


def print_notepad_simple_tree() -> None:
    tree = collect_notepad_simple_tree()
    menu_children = tree["children"][0]["children"]

    print("Notepad")
    print("├─ MenuBar")
    for menu_node in menu_children:
        print(f"│ ├─ {menu_node['name']}")

    if len(tree["children"]) > 1 and tree["children"][1]["name"] == "EditArea":
        print("└─ EditArea")


def print_notepad_simple_tree_json() -> None:
    print(json.dumps(collect_notepad_simple_tree(), ensure_ascii=False, indent=2))


def _simplify_tree_dict(node: Dict[str, Any]) -> Dict[str, Any]:
    """Keep only name and children for frontend display."""
    name = (node.get("name") or "").strip() or (node.get("control_type") or "")
    result: Dict[str, Any] = {"name": name}
    kids = [_simplify_tree_dict(c) for c in node.get("children", []) if c]
    if kids:
        result["children"] = kids
    return result


def select_screen_roi() -> Optional[ScreenROI]:
    import tkinter as tk

    root = tk.Tk()
    root.attributes("-fullscreen", True)
    root.attributes("-topmost", True)
    root.attributes("-alpha", 0.2)
    root.configure(bg="black")
    root.title("Select ROI")

    canvas = tk.Canvas(root, cursor="cross", bg="black", highlightthickness=0)
    canvas.pack(fill="both", expand=True)

    hint = canvas.create_text(
        40,
        30,
        anchor="w",
        text="拖曳框選要擷取的區域（放開滑鼠即確認），Esc 取消",
        fill="white",
        font=("Segoe UI", 16, "bold"),
    )

    state: Dict[str, Any] = {"start": None, "rect": None, "roi": None, "cancelled": False}

    def normalize(x1: int, y1: int, x2: int, y2: int) -> ScreenROI:
        left = min(x1, x2)
        right = max(x1, x2)
        top = min(y1, y2)
        bottom = max(y1, y2)
        return ScreenROI(left=left, top=top, right=right, bottom=bottom)

    def on_down(event):
        state["start"] = (event.x, event.y)
        if state["rect"] is not None:
            canvas.delete(state["rect"])
            state["rect"] = None

    def on_move(event):
        if not state["start"]:
            return
        sx, sy = state["start"]
        if state["rect"] is None:
            state["rect"] = canvas.create_rectangle(sx, sy, event.x, event.y, outline="#22d3ee", width=3)
        else:
            canvas.coords(state["rect"], sx, sy, event.x, event.y)

    def on_up(event):
        if not state["start"]:
            return
        sx, sy = state["start"]
        roi = normalize(sx, sy, event.x, event.y)
        if roi.right - roi.left < 6 or roi.bottom - roi.top < 6:
            state["roi"] = None
            return
        state["roi"] = roi
        root.quit()

    def on_enter(_event):
        root.quit()

    def on_esc(_event):
        state["cancelled"] = True
        state["roi"] = None
        root.quit()

    canvas.bind("<ButtonPress-1>", on_down)
    canvas.bind("<B1-Motion>", on_move)
    canvas.bind("<ButtonRelease-1>", on_up)
    root.bind("<Return>", on_enter)
    root.bind("<Escape>", on_esc)
    canvas.tag_raise(hint)

    root.mainloop()
    roi = state["roi"]
    root.destroy()

    if state["cancelled"]:
        return None
    return roi


def _root_hwnd_from_screen_point(screen_x: int, screen_y: int) -> int:
    import ctypes

    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    user32 = ctypes.windll.user32  # type: ignore[attr-defined]
    pt = POINT(screen_x, screen_y)
    hwnd = user32.WindowFromPoint(pt)
    if not hwnd:
        raise RuntimeError("無法從框選區域取得目標視窗")

    GA_ROOT = 2
    root_hwnd = user32.GetAncestor(hwnd, GA_ROOT)
    if root_hwnd:
        hwnd = root_hwnd

    return int(hwnd)


def collect_window_tree_by_screen_point(
    screen_x: int,
    screen_y: int,
    max_depth: int = 6,
    backend: str = "uia",
) -> Dict[str, Any]:
    hwnd = _root_hwnd_from_screen_point(screen_x, screen_y)

    target_window = Desktop(backend=backend).window(handle=hwnd)
    target_window.wait("exists", timeout=2)
    return build_tree_dict(target_window, depth=0, max_depth=max_depth)


def collect_foreground_window_tree(max_depth: int = 3) -> Dict[str, Any]:
    """Capture the currently active foreground window's UI tree."""
    import ctypes
    hwnd = ctypes.windll.user32.GetForegroundWindow()  # type: ignore[attr-defined]
    if not hwnd:
        raise RuntimeError("無法取得前景視窗 HWND")
    win = Desktop(backend="uia").window(handle=hwnd)
    win.wait("exists", timeout=2)
    raw = build_tree_dict(win, depth=0, max_depth=max_depth)
    return _simplify_tree_dict(raw)


def print_foreground_window_tree_json(max_depth: int = 3) -> None:
    print(json.dumps(collect_foreground_window_tree(max_depth=max_depth), ensure_ascii=False))


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="列出 Windows App 的 UI Tree")
    parser.add_argument("--launch", type=str, help="啟動指令，例如 notepad.exe")
    parser.add_argument("--exe", type=str, help="附著到已開啟程式，例如 notepad.exe")
    parser.add_argument("--title", type=str, help="用視窗標題 regex 尋找，例如 .*記事本.*|.*Notepad.*")
    parser.add_argument("--max-depth", type=int, default=6, help="最大遞迴層數")
    parser.add_argument("--show-rect", action="store_true", help="顯示元素座標範圍")
    parser.add_argument("--json-out", type=str, help="輸出 JSON 檔案路徑")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)

    if not any([args.launch, args.exe, args.title]):
        print("請至少提供一個參數：--launch 或 --exe 或 --title")
        print(r'範例: python ui_tree.py --launch notepad.exe --title ".*記事本.*|.*Notepad.*"')
        sys.exit(1)

    if args.launch:
        print(f"[INFO] 啟動程式: {args.launch}")
        launch_app(args.launch)

    target_window: Optional[BaseWrapper] = None

    if args.exe:
        print(f"[INFO] 附著到程式: {args.exe}")
        target_window = connect_by_exe(args.exe)

    if args.title:
        print(f"[INFO] 依標題尋找視窗: {args.title}")
        target_window = find_window_by_title(args.title)

    if target_window is None:
        raise RuntimeError("無法取得目標視窗")

    print("\n[UI TREE]")
    print_tree(target_window, depth=0, max_depth=args.max_depth, show_rect=args.show_rect)

    if args.json_out:
        tree_data = build_tree_dict(target_window, depth=0, max_depth=args.max_depth)
        save_json(tree_data, args.json_out)
        print(f"\n[INFO] UI Tree JSON 已輸出: {args.json_out}")


if __name__ == "__main__":
    main()