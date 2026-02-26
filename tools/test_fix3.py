"""Test navigation fix3 (Ctrl+Shift+F overlay close) + chat scan."""
import sys, io, time, base64
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, '.')
import logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logging.getLogger('openai').setLevel(logging.WARNING)
logging.getLogger('httpx').setLevel(logging.WARNING)

from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController
from dd_collector.vlm import grab_screenshot_base64

cfg = load_config()
ctrl = DingTalkController(cfg)
ctrl.connect()

print('--- Navigating to 资料分享群 ---')
r = ctrl.navigate_to_group(cfg.groups[1].name)
print('Result: ' + str(r))
ss = grab_screenshot_base64()
png = base64.b64decode(ss)
with open('tools/fix3_ziliao.png', 'wb') as f:
    f.write(png)
print('Saved fix3_ziliao.png')

print()
print('--- Scanning chat (max_scrolls=1) ---')
attachments = ctrl.scan_chat_attachments(max_scrolls=1)
print('Found ' + str(len(attachments)) + ' attachments:')
for a in attachments[:10]:
    print('  [' + a.msg_type + '] ' + repr(a.filename) + '  click=' + str(a.download_click))
