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

group = cfg.groups[1]  # 资料分享群 (higher update frequency)
print(f"\nNavigating to: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

# Wait extra time for CefBrowserWindow to render file list
print("Waiting 5s for file list to load...")
time.sleep(5)

# max_scrolls=0: don't scroll — newest files are already visible at top
files = ctrl.list_files(max_scrolls=0)
if not files:
    print("No files on first try. Waiting 5s and retrying...")
    time.sleep(5)
    files = ctrl.list_files(max_scrolls=0)

if not files:
    # Debug: dump the grid area
    import uiautomation as auto
    print("\nDEBUG: Dumping file grid area...")
    window = auto.WindowControl(ClassName="DtMainFrameView", searchDepth=1)
    grid = None
    try:
        grid = auto.GroupControl(searchFromControl=window, Name="grid", searchDepth=10)
        if grid.Exists(maxSearchSeconds=3):
            print(f"  Grid found: Name={grid.Name!r}, rect={grid.BoundingRectangle}")
            for child in grid.GetChildren()[:5]:
                print(f"    Child: type={child.ControlTypeName}, Name={child.Name!r:.60s}")
                for sub in child.GetChildren()[:3]:
                    print(f"      Sub: type={sub.ControlTypeName}, Name={sub.Name!r:.60s}")
        else:
            print("  Grid NOT found.")
    except Exception as e:
        print(f"  Grid search error: {e}")
    print("FAIL: No files found")
    sys.exit(1)

print(f"\nFound {len(files)} files.")
for i, f in enumerate(files[:8]):
    rect = f.control.BoundingRectangle
    print(f"  [{i}] {f.name}  ts={f.timestamp}  rect=({rect.left},{rect.top},{rect.right},{rect.bottom})")

# Pick the newest file to test download
target = files[0]
print(f"\nTesting download on: {target.name}  (ts={target.timestamp})")

# Snapshot download dir before
from pathlib import Path
dl_dir = Path(cfg.dingtalk.download_dir)
before = {f.name for f in dl_dir.iterdir() if f.is_file()} if dl_dir.is_dir() else set()
print(f"Download dir files before: {len(before)}")

print("\nAttempting VLM-guided download...")
result = ctrl.download_file(target)
print(f"\nDownload triggered: {'YES' if result else 'NO'}")

# Check what appeared in download dir
time.sleep(3)
if dl_dir.is_dir():
    after_files = {f.name for f in dl_dir.iterdir() if f.is_file()}
    after_dirs = {f.name for f in dl_dir.iterdir() if f.is_dir()}
    new_files = after_files - before
    print(f"New files in download dir: {new_files or '(none)'}")
    print(f"Directories in download dir: {after_dirs or '(none)'}")
