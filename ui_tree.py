from __future__ import annotations

import argparse
from typing import List, Optional

import ui_tree2


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Notepad UI Tree Viewer (compatibility wrapper)",
        add_help=True,
    )
    parser.add_argument("--json", action="store_true", dest="as_json", help="Output UI tree as JSON")
    parser.add_argument("--foreground", action="store_true", help="擷取目前前景視窗 UI Tree（JSON 輸出）")
    parser.add_argument(
        "--advanced",
        action="store_true",
        help="Forward remaining arguments to ui_tree2 advanced explorer",
    )
    args, unknown = parser.parse_known_args(argv)
    args.forward_args = unknown
    return args


def main(argv: Optional[List[str]] = None) -> None:
    args = parse_args(argv)

    if args.advanced:
        ui_tree2.main(args.forward_args)
        return

    if args.foreground:
        ui_tree2.print_foreground_window_tree_json()
        return

    if args.forward_args:
        raise SystemExit(
            "Unsupported arguments for compatibility mode. "
            "Use --advanced to forward options to ui_tree2."
        )

    if args.as_json:
        ui_tree2.print_notepad_simple_tree_json()
        return

    ui_tree2.print_notepad_simple_tree()


if __name__ == "__main__":
    main()
