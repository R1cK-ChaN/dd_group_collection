"""Dump DingTalk's UI control tree for debugging UI selector changes.

Usage:
    python tools/inspect_dingtalk.py
    python tools/inspect_dingtalk.py --depth 5
    python tools/inspect_dingtalk.py --search "文件"
    python tools/inspect_dingtalk.py --class StandardFrame_DingTalk
"""

from __future__ import annotations

import argparse
import sys

try:
    import uiautomation as auto
except ImportError:
    print("ERROR: uiautomation not installed. Run: pip install uiautomation", file=sys.stderr)
    sys.exit(1)


def dump_tree(control: auto.Control, depth: int, max_depth: int, search: str, indent: int = 0) -> None:
    """Recursively print the control tree."""
    if indent > max_depth:
        return

    name = control.Name or ""
    ctrl_type = control.ControlTypeName
    class_name = control.ClassName or ""
    auto_id = control.AutomationId or ""

    # Apply search filter
    if search:
        match = (
            search.lower() in name.lower()
            or search.lower() in ctrl_type.lower()
            or search.lower() in class_name.lower()
            or search.lower() in auto_id.lower()
        )
    else:
        match = True

    prefix = "  " * indent
    line = (
        f"{prefix}[{ctrl_type}] "
        f"Name={name!r}  "
        f"Class={class_name!r}  "
        f"AutoId={auto_id!r}"
    )

    if match or not search:
        print(line)

    try:
        children = control.GetChildren()
    except Exception:
        children = []

    for child in children:
        dump_tree(child, depth, max_depth, search, indent + 1)


def main():
    parser = argparse.ArgumentParser(description="Inspect DingTalk UI control tree")
    parser.add_argument(
        "--depth", type=int, default=6,
        help="Max depth to traverse (default: 6)",
    )
    parser.add_argument(
        "--search", type=str, default="",
        help="Filter: only show controls matching this text (in Name, Type, Class, or AutoId)",
    )
    parser.add_argument(
        "--class", dest="window_class", type=str, default="StandardFrame_DingTalk",
        help="DingTalk window class name (default: StandardFrame_DingTalk)",
    )
    args = parser.parse_args()

    print(f"Looking for DingTalk window (class={args.window_class})...")
    window = auto.WindowControl(ClassName=args.window_class, searchDepth=1)

    if not window.Exists(maxSearchSeconds=5):
        print("ERROR: DingTalk window not found. Is DingTalk running?", file=sys.stderr)
        print("Tip: Try --class with a different class name.", file=sys.stderr)
        print("\nListing all top-level windows:", file=sys.stderr)
        for w in auto.GetRootControl().GetChildren():
            print(f"  [{w.ControlTypeName}] Name={w.Name!r} Class={w.ClassName!r}", file=sys.stderr)
        sys.exit(1)

    print(f"Found DingTalk window: {window.Name!r}")
    print(f"Dumping control tree (depth={args.depth}, search={args.search!r})...")
    print("=" * 80)

    dump_tree(window, args.depth, args.depth, args.search)


if __name__ == "__main__":
    main()
