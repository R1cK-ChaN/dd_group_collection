"""Test VLM-based download: right-click a file → VLM finds Download → click.

Usage:
    cd C:/Users/vm/dd_group_collection
    python -X utf8 tools/test_vlm_download.py
"""
import io
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

import logging
logging.basicConfig(level=logging.DEBUG, format="%(levelname)s %(name)s: %(message)s")

from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
ctrl = DingTalkController(cfg)

print(f"VLM model: {cfg.vlm.model}")
print(f"VLM key set: {bool(cfg.vlm.api_key)}")

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

group = cfg.groups[0]  # Degg
print(f"\nNavigating to: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

# max_scrolls=0: don't scroll — newest files are already visible at top
files = ctrl.list_files(max_scrolls=0)
if not files:
    print("FAIL: No files found")
    sys.exit(1)

print(f"\nFound {len(files)} files.")
for i, f in enumerate(files[:5]):
    rect = f.control.BoundingRectangle
    print(f"  [{i}] {f.name}  ts={f.timestamp}  rect=({rect.left},{rect.top},{rect.right},{rect.bottom})")

target = files[0]
print(f"\nTesting download on: {target.name}")

print("\nAttempting VLM-guided download...")
result = ctrl.download_file(target)
print(f"\nResult: {'SUCCESS' if result else 'FAILED'}")
