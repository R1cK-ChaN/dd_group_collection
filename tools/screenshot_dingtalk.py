"""Navigate to group files tab, minimize others, screenshot DingTalk."""
import ctypes
import subprocess
import sys
import time
sys.path.insert(0, ".")

import uiautomation as auto
from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController

# Kill any leftover Narrator
subprocess.run(['taskkill', '/f', '/im', 'Narrator.exe'], capture_output=True)
subprocess.run(['taskkill', '/f', '/im', 'NarratorQuickStart.exe'], capture_output=True)

# Close any "Narrator updates" or "Narrator Home" windows
for name in ["Narrator updates", "Narrator Home", "Welcome to Narrator"]:
    w = auto.WindowControl(Name=name, searchDepth=3)
    if w.Exists(maxSearchSeconds=1):
        try:
            w.Close()
        except:
            pass

time.sleep(0.5)

cfg = load_config()
ctrl = DingTalkController(cfg)

if not ctrl.connect():
    print("FAIL: Cannot connect to DingTalk")
    sys.exit(1)

# Navigate and open files tab
group = cfg.groups[0]
print(f"Navigating to: {group.name}")
if not ctrl.navigate_to_group(group.name):
    print("FAIL: Navigate failed")
    sys.exit(1)

print("Opening files tab...")
if not ctrl.open_files_tab():
    print("FAIL: Files tab failed")
    sys.exit(1)

print("Waiting 3s for content to load...")
time.sleep(3)

# Minimize console/terminal
user32 = ctypes.windll.user32
console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()
if console_hwnd:
    user32.ShowWindow(console_hwnd, 6)

# Minimize OpenClaw windows
def minimize_windows(partial_title):
    def callback(hwnd, _):
        length = user32.GetWindowTextLengthW(hwnd)
        if length > 0:
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            if partial_title.lower() in buf.value.lower():
                user32.ShowWindow(hwnd, 6)
        return True
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    user32.EnumWindows(WNDENUMPROC(callback), 0)

minimize_windows("OpenClaw")
minimize_windows("Select")
minimize_windows("Narrator")
time.sleep(0.3)

# Activate DingTalk
ctrl._window.Maximize()
ctrl._window.SetActive()
ctrl._window.SetFocus()
time.sleep(1)

# Take screenshot
subprocess.run([
    "powershell", "-ExecutionPolicy", "Bypass", "-Command",
    r"""
    Add-Type -AssemblyName System.Windows.Forms
    $bmp = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen(0, 0, 0, 0, $bmp.Size)
    $bmp.Save("tools\screenshot_dingtalk.png")
    $g.Dispose()
    $bmp.Dispose()
    """
], check=True)
print("Screenshot saved")
