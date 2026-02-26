"""Dump all ButtonControls visible when the DingTalk search overlay is open."""
import sys, io, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import uiautomation as auto
import pyautogui

win = auto.WindowControl(ClassName='StandardFrame_DingTalk', searchDepth=1)
win.Exists(maxSearchSeconds=3)
win.SetActive(); win.SetFocus()
time.sleep(0.3)

box = auto.EditControl(searchFromControl=win, ClassName='QLineEdit', searchDepth=10)
box.Exists(3)
box.Click(simulateMove=False)
time.sleep(0.5)
box.SendKeys('Degg', interval=0.05)
time.sleep(2.0)

print('=== ButtonControls in DingTalk window ===')

def dump_buttons(ctrl, depth=0, max_depth=8):
    if depth > max_depth:
        return
    if ctrl.ControlTypeName == 'ButtonControl':
        r = ctrl.BoundingRectangle
        name = (ctrl.Name or '')[:60]
        print('  ' * depth + '[BTN] Name=' + repr(name) + ' rect=(' + str(r.left) + ',' + str(r.top) + ',' + str(r.right) + ',' + str(r.bottom) + ')')
    try:
        for child in ctrl.GetChildren():
            dump_buttons(child, depth+1, max_depth)
    except Exception:
        pass

dump_buttons(win)
print('=== ALL DONE ===')

pyautogui.press('escape')
time.sleep(0.3)
pyautogui.press('escape')
