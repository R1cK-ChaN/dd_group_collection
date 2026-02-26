"""Navigate to group files tab, screenshot, and dump control info."""
import sys
import time
import subprocess
sys.path.insert(0, ".")

import uiautomation as auto
from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController

cfg = load_config()
ctrl = DingTalkController(cfg)

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

# Maximize DingTalk
ctrl._window.Maximize()
time.sleep(0.5)

group = cfg.groups[0]  # "Degg"
print(f"Navigating to group: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Cannot navigate to group")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Cannot open files tab")
    sys.exit(1)

# Wait for web view to load
print("Waiting 5s for web view to load...")
time.sleep(5)

# Take screenshot
print("Taking screenshot...")
subprocess.run([
    "powershell", "-ExecutionPolicy", "Bypass", "-Command",
    r"""
    Add-Type -AssemblyName System.Windows.Forms
    $bmp = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen(0, 0, 0, 0, $bmp.Size)
    $bmp.Save("tools\screenshot_files_tab.png")
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Screenshot saved"
    """
], check=True)

# Also dump what controls we can find
print("\n=== Controls near file area ===")
window = ctrl._window

# Search for any controls with 'File' or 'file' in name
for name_search in ["File", "grid", "Upload", "Download", "pdf", "xlsx"]:
    results = []
    def search(c, depth=0, max_d=12):
        if depth > max_d: return
        n = c.Name or ""
        if name_search.lower() in n.lower():
            r = c.BoundingRectangle
            results.append(f"  [{c.ControlTypeName}] Name={n!r:.60s} Rect=({r.left},{r.top},{r.right},{r.bottom})")
        try:
            for ch in c.GetChildren():
                search(ch, depth+1, max_d)
        except: pass
    search(window)
    if results:
        print(f"\nSearching '{name_search}':")
        for r in results:
            print(r)

print("\nDone.")
