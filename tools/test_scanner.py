"""Diagnostic: navigate to each group, save chat panel screenshot, run scanner."""
import sys, io, base64, time, logging
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, '.')
logging.basicConfig(level=logging.DEBUG, format='%(levelname)s: %(message)s')
logging.getLogger('anthropic').setLevel(logging.WARNING)
logging.getLogger('httpx').setLevel(logging.WARNING)

from dd_collector.config import load_config
from dd_collector.dingtalk_ui import DingTalkController
from dd_collector.chat_scanner import ChatScanner, _PROMPT

cfg = load_config()
ctrl = DingTalkController(cfg)

print('=== ensure_running ===')
ctrl.ensure_running()
ctrl.connect()
ctrl.wait_for_ready()

scanner = ChatScanner(cfg.claude.oauth_token, cfg.claude.model)

# Patch to show raw Claude response
orig = scanner._client.messages.create
def patched(**kw):
    resp = orig(**kw)
    raw = ''.join(b.text for b in resp.content if hasattr(b, 'text'))
    print('  [RAW CLAUDE RESPONSE]:', repr(raw[:500]))
    return resp
scanner._client.messages.create = patched

for i, group in enumerate(cfg.groups):
    print(f'\n=== Group {i+1}: {group.name} ===')
    ok = ctrl.navigate_to_group(group.name)
    print(f'  navigate_to_group: {ok}')
    if not ok:
        continue
    time.sleep(1.5)

    ss_b64, ox, oy = ctrl.get_chat_panel_screenshot()
    save_path = f'tools/debug_group{i+1}.png'
    with open(save_path, 'wb') as f:
        f.write(base64.b64decode(ss_b64))
    print(f'  Screenshot saved: {save_path}  offset=({ox},{oy})  b64_len={len(ss_b64)}')

    results = scanner.find_downloads(ss_b64, region_offset=(ox, oy))
    print(f'  Results: {results}')
