"""Test the incremental watermark logic.

Runs a single cycle for the Degg group and prints how many new files
the watermark filter detects. With watermark seeded, should report 0.

Usage:
    cd C:/Users/vm/dd_group_collection
    python -X utf8 tools/test_incremental.py
"""
import io
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

from dd_collector.config import load_config
from dd_collector.dedup import DedupTracker
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
dedup = DedupTracker(cfg.dedup.path)
ctrl = DingTalkController(cfg)

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

group = cfg.groups[0]  # Degg
print(f"Group: {group.name}")

watermark = dedup.get_watermark(group.name)
print(f"Current watermark: {watermark}")

if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate to group")
    sys.exit(1)

if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

# Incremental scan — minimal scroll
files = ctrl.list_files(max_scrolls=3)
print(f"Visible files: {len(files)}")

if files:
    print(f"  Newest: {files[0].name}  ts={files[0].timestamp}")
    print(f"  Oldest visible: {files[-1].name}  ts={files[-1].timestamp}")

# Apply watermark filter
new_files = []
for f in files:
    if watermark and f.timestamp and f.timestamp <= watermark:
        print(f"  BREAK at: {f.name}  ts={f.timestamp} <= watermark")
        break
    if dedup.is_downloaded(group.name, f.name):
        print(f"  SKIP (dedup): {f.name}")
        continue
    new_files.append(f)

print(f"\nNew files after filter: {len(new_files)}")
for f in new_files:
    print(f"  NEW: {f.name}  ts={f.timestamp}")

if not new_files:
    print("\n✓ Watermark filter working correctly — no new files detected.")
else:
    print(f"\n⚠ Found {len(new_files)} new file(s) above watermark.")
