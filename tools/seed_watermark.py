"""One-time tool: seed the watermark for each group.

- Degg: navigates to group, reads the newest file timestamp, sets watermark.
- 资料分享群: sets watermark to current time (only future files).

Usage:
    cd C:/Users/vm/dd_group_collection
    python -X utf8 tools/seed_watermark.py
"""
import io
import sys
import time

# Ensure console can print Chinese characters
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

from datetime import datetime

from dd_collector.config import load_config
from dd_collector.dedup import DedupTracker
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
dedup = DedupTracker(cfg.dedup.path)
ctrl = DingTalkController(cfg)


def seed_from_file_list(group_name: str) -> None:
    """Navigate to a group, list files, and set watermark to the newest timestamp."""
    print(f"\n{'='*50}")
    print(f"Seeding watermark for: {group_name}")

    existing = dedup.get_watermark(group_name)
    if existing:
        print(f"  Watermark already set: {existing}")
        resp = input("  Overwrite? [y/N] ").strip().lower()
        if resp != "y":
            print("  Skipped.")
            return

    if not ctrl.navigate_to_group(group_name):
        print(f"  FAIL: cannot navigate to {group_name}")
        return

    if not ctrl.open_files_tab():
        print(f"  FAIL: cannot open files tab for {group_name}")
        return

    # Minimal scroll — newest files are at top
    files = ctrl.list_files(max_scrolls=2)
    if not files:
        print(f"  No files found in {group_name}")
        return

    # Files are sorted newest-first; take the first file's timestamp
    newest = files[0]
    if not newest.timestamp:
        print(f"  WARNING: newest file has no timestamp: {newest.name}")
        return

    print(f"  Newest file: {newest.name}")
    print(f"  Timestamp:   {newest.timestamp}")
    dedup.set_watermark(group_name, newest.timestamp)
    print(f"  ✓ Watermark set to {newest.timestamp}")


def seed_from_now(group_name: str) -> None:
    """Set watermark to the current time (only pick up future uploads)."""
    print(f"\n{'='*50}")
    print(f"Seeding watermark for: {group_name} (from now)")

    existing = dedup.get_watermark(group_name)
    if existing:
        print(f"  Watermark already set: {existing}")
        resp = input("  Overwrite? [y/N] ").strip().lower()
        if resp != "y":
            print("  Skipped.")
            return

    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    dedup.set_watermark(group_name, now_str)
    print(f"  ✓ Watermark set to {now_str}")


def main() -> None:
    if not ctrl.connect():
        print("FAIL: Cannot connect to DingTalk. Is it running?")
        sys.exit(1)

    # Degg: read newest file timestamp from the file list
    seed_from_file_list("Degg")

    time.sleep(1)

    # 资料分享群: set to current time
    seed_from_now("资料分享群")

    print(f"\n{'='*50}")
    print("Done. Current watermarks:")
    for group in cfg.groups:
        wm = dedup.get_watermark(group.name)
        print(f"  {group.name}: {wm or '(not set)'}")


if __name__ == "__main__":
    main()
