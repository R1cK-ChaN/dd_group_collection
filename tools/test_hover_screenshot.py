"""Hover over the first file row and save a screenshot to inspect hover icons.

Usage:
    cd C:/Users/vm/dd_group_collection
    python -X utf8 tools/test_hover_screenshot.py
"""
import io
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

import logging
logging.basicConfig(level=logging.DEBUG, format="%(levelname)s %(name)s: %(message)s")

import pyautogui
from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController
from dd_collector.vlm import grab_screenshot_base64
import base64
from pathlib import Path

cfg = load_config()
ctrl = DingTalkController(cfg)

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

group = cfg.groups[1]  # 资料分享群
print(f"\nNavigating to: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

print("Waiting 5s for file list to load...")
time.sleep(5)

files = ctrl.list_files(max_scrolls=0)
if not files:
    print("No files found. Retrying...")
    time.sleep(5)
    files = ctrl.list_files(max_scrolls=0)

if not files:
    print("FAIL: No files found")
    sys.exit(1)

target = files[0]
rect = target.control.BoundingRectangle
print(f"\nTarget file: {target.name}")
print(f"  rect = ({rect.left}, {rect.top}, {rect.right}, {rect.bottom})")
print(f"  size = {rect.width()}x{rect.height()}")

# Ensure DingTalk is focused
try:
    ctrl._window.SetActive()
    ctrl._window.SetFocus()
    time.sleep(0.3)
except Exception:
    pass

# Step 1: Take screenshot BEFORE hover (baseline)
print("\n1) Capturing screenshot BEFORE hover...")
cap_left = max(0, rect.left - 50)
cap_top = max(0, rect.top - 20)
cap_right = rect.right + 200
cap_bottom = rect.bottom + 20
region = (cap_left, cap_top, cap_right, cap_bottom)

b64_before = grab_screenshot_base64(region)
before_path = Path("data/hover_before.png")
before_path.parent.mkdir(parents=True, exist_ok=True)
before_path.write_bytes(base64.b64decode(b64_before))
print(f"   Saved to {before_path.resolve()}")

# Step 2: Hover over row center
row_cx = (rect.left + rect.right) // 2
row_cy = (rect.top + rect.bottom) // 2
print(f"\n2) Hovering at row center ({row_cx}, {row_cy})...")
pyautogui.moveTo(row_cx, row_cy)
time.sleep(2.0)  # generous wait

# Step 3: Take screenshot AFTER hover (should show action icons)
print("\n3) Capturing screenshot AFTER hover...")
b64_after = grab_screenshot_base64(region)
after_path = Path("data/hover_after.png")
after_path.write_bytes(base64.b64decode(b64_after))
print(f"   Saved to {after_path.resolve()}")

# Step 4: Also save a wider screenshot to see the full row area
print("\n4) Capturing wide screenshot (full DingTalk file area)...")
wide_region = (rect.left - 100, rect.top - 50, rect.right + 400, rect.bottom + 200)
b64_wide = grab_screenshot_base64(wide_region)
wide_path = Path("data/hover_wide.png")
wide_path.write_bytes(base64.b64decode(b64_wide))
print(f"   Saved to {wide_path.resolve()}")

print(f"\n5) Download icon offset from config: {cfg.dingtalk.download_icon_offset}")
print(f"   Offset click would land at x={rect.right - cfg.dingtalk.download_icon_offset}, y={row_cy}")

print("\nDone! Check the screenshots in data/ to see if hover icons appeared.")
print("Compare hover_before.png and hover_after.png to spot the icons.")
