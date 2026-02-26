"""Navigate to a group, open files tab, and inspect the file grid area."""
import sys
import time
sys.path.insert(0, ".")

import uiautomation as auto
from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
ctrl = DingTalkController(cfg)

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

group = cfg.groups[0]  # "Degg"
print(f"Navigating to group: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate to group")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

# Wait extra time for the CefBrowserWindow to load
print("Waiting 3s for web view to load...")
time.sleep(3)

# Now inspect the tree deeply
def dump(control, depth=0, max_depth=12):
    if depth > max_depth:
        return
    name = control.Name or ""
    ct = control.ControlTypeName
    cn = control.ClassName or ""
    aid = control.AutomationId or ""
    rect = control.BoundingRectangle
    r = f"({rect.left},{rect.top},{rect.right},{rect.bottom})" if rect else "no-rect"
    prefix = "  " * depth
    print(f"{prefix}[{ct}] Name={name!r:.60s} Class={cn!r} AutoId={aid!r} Rect={r}")
    try:
        for child in control.GetChildren():
            dump(child, depth + 1, max_depth)
    except Exception:
        pass

# Find the CefBrowserWindow
window = auto.WindowControl(ClassName="StandardFrame_DingTalk", searchDepth=1)
print("\n=== Searching for CefBrowserWindow ===")

# Search for any PaneControl with CefBrowserWindow class
cef = None
for child in window.GetChildren():
    if child.ClassName == "CefBrowserWindow":
        cef = child
        break
    for grandchild in child.GetChildren():
        if grandchild.ClassName == "CefBrowserWindow":
            cef = grandchild
            break
        for gg in grandchild.GetChildren():
            if gg.ClassName == "CefBrowserWindow":
                cef = gg
                break
    if cef:
        break

if cef:
    print(f"Found CefBrowserWindow: Name={cef.Name!r}")
    print("\n=== CefBrowserWindow subtree ===")
    dump(cef, max_depth=8)
else:
    print("CefBrowserWindow not found. Dumping full window tree:")
    dump(window, max_depth=10)
