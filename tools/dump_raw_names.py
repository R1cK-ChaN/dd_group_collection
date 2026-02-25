"""Dump raw Name attributes of file rows in the second group."""
import io, sys, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

import uiautomation as auto
from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
ctrl = DingTalkController(cfg)
if not ctrl.connect():
    print("FAIL: connect"); sys.exit(1)
print("Connected")

group = cfg.groups[1]  # 资料分享群
print(f"Group: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: navigate"); sys.exit(1)
print("Navigated")

if not ctrl.open_files_tab():
    print("FAIL: files tab"); sys.exit(1)
print("Files tab opened, waiting 5s for content...")
time.sleep(5)

# Try list_files directly (it has all the grid-finding logic)
files = ctrl.list_files(max_scrolls=0)
print(f"\nlist_files returned: {len(files)} files")
if files:
    for i, f in enumerate(files[:17]):
        print(f"  [{i}] raw Name: {f.control.Name!r}")
        print(f"       parsed: name={f.name}, ts={f.timestamp}")
else:
    # Fallback: dump accessibility tree around CefBrowserWindow
    print("\nFallback: searching for any GroupControl with Name='grid'...")
    window = auto.WindowControl(ClassName="DtMainFrameView", searchDepth=1)

    # Search for DocumentControl (Chrome_RenderWidgetHostHWND)
    doc = None
    for ctrl_type in ("DocumentControl", "PaneControl"):
        try:
            c = getattr(auto, ctrl_type)(searchFromControl=window,
                                          ClassName="Chrome_RenderWidgetHostHWND",
                                          searchDepth=10)
            if c.Exists(maxSearchSeconds=3):
                doc = c
                print(f"  Found {ctrl_type} Chrome_RenderWidgetHostHWND")
                break
        except Exception:
            pass

    if doc:
        print(f"  Children of doc:")
        for child in doc.GetChildren()[:10]:
            print(f"    {child.ControlTypeName} Name={child.Name!r:.60s} Class={child.ClassName!r}")
    else:
        print("  No Chrome_RenderWidgetHostHWND found. Dumping top-level children:")
        for child in window.GetChildren()[:8]:
            print(f"    {child.ControlTypeName} Name={child.Name!r:.40s} Class={child.ClassName!r}")
