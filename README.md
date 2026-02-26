# DingTalk Group File Collector

Automated file collection from DingTalk group chats, synced to Google Drive.

Monitors two DingTalk groups for new file attachments (PDFs, spreadsheets, images, etc.), downloads them automatically, and moves them into an organised Google Drive folder structure.

---

## Architecture — Hybrid Programmatic + Claude Vision

DingTalk renders its chat content inside a **CefBrowserWindow** (Chromium Embedded Framework) that is completely invisible to Windows UI Automation. The solution splits every task into two categories:

| Task | Approach | Cost |
|------|----------|------|
| Launch DingTalk if not running | `subprocess.Popen` | Free |
| Find & activate DingTalk window | `uiautomation` | Free |
| Wait for UI to be ready, dismiss dialogs | `uiautomation` | Free |
| Navigate to a group | `uiautomation` + `pyautogui` | Free |
| Scroll chat panel up/down | `pyautogui.scroll` | Free |
| **Find file cards + Download button coords** | **Claude Haiku vision (1 API call/scroll)** | ~$0.001/call |
| Click Download button | `pyautogui.click` | Free |
| Detect & move new files to GDrive | `file_mover.py` | Free |
| Dedup tracking | `dedup.py` (JSON) | Free |

**Token cost per group run:** 3–8 Claude API calls × ~880 tokens = **~$0.003–0.007/group**
vs. old autonomous agent approach: 30–60 calls × ~8000 tokens = **~$0.45–0.90/group** (~100× cheaper)

```
run_claude.py
 │
 ├─ DingTalkController.ensure_running()    # auto-launch if needed
 ├─ DingTalkController.connect()           # find + focus window
 ├─ DingTalkController.wait_for_ready()    # dismiss startup dialogs until search box appears
 │
 └─ for each group:
     ├─ navigate_to_group()               # uiautomation: search → Down+Enter → Esc overlay
     │
     └─ for scroll_pass in range(max_scrolls):
         ├─ get_chat_panel_screenshot()   # mss DirectX capture (captures CefBrowserWindow)
         ├─ ChatScanner.find_downloads()  # 1× Claude Haiku API call → JSON coordinates
         ├─ click_download_at(x, y)       # pyautogui click + save-dialog handler
         └─ scroll_chat_up()             # pyautogui scroll to reveal older messages
     │
     ├─ get_new_files()                  # detect settled files in download dir
     ├─ move_file_to_gdrive()            # organise into {group}/{YYYY-MM}/{filename}
     └─ dedup.mark_downloaded()         # prevent re-downloads
```

---

## Deployment Status (2026-02-26)

| Component | Status | Notes |
|-----------|--------|-------|
| Python environment | ✅ Working | Python 3.14.3, venv at `.venv/` |
| DingTalk auto-launch | ✅ Working | Detects if not running, launches exe, waits up to 30s |
| DingTalk window connect | ✅ Working | Class `StandardFrame_DingTalk` |
| Startup dialog dismissal | ✅ Working | `wait_for_ready()` loops until search box appears |
| Navigate to group | ✅ Working | Search overlay handled (Down+Enter → Escape) |
| Chat panel screenshot | ✅ Fixed | Uses `mss` (DirectX) — `pyautogui` returned blank for CEF content |
| Claude Haiku API | ✅ Working | Model `claude-haiku-4-5-20251001`, plain messages API (no beta) |
| Download button detection | 🔧 Testing | End-to-end verification in progress |
| File move to GDrive | ✅ Implemented | `G:\My Drive\DingTalk Files\{group}\{YYYY-MM}\` |
| Dedup tracking | ✅ Working | `data/downloaded.json` |
| Auto-start on boot | ✅ Deployed | `dd_group_collection.bat` in `shell:startup` |

---

## Quick Start

```bash
git clone https://github.com/R1cK-ChaN/dd_group_collection.git
cd dd_group_collection
pip install -r requirements.txt
```

Edit `config.yaml`:

```yaml
dingtalk:
  exe_path: "C:\\Program Files (x86)\\DingDing\\main\\current\\DingTalk.exe"
  download_dir: "C:\\Users\\YOU\\Documents\\DingDing\\download"

groups:
  - name: "Your Group Name"
    alias: "FolderName"

gdrive:
  base_path: "G:\\My Drive\\DingTalk Files"

claude:
  oauth_token: "sk-ant-api03-..."   # Anthropic API key from console.anthropic.com
  model: "claude-haiku-4-5-20251001"
  max_scrolls: 5
```

Run once (processes all groups then exits):

```bash
python run_claude.py
```

Run in polling loop (every 30 minutes by default):

```bash
python run_claude.py --loop
```

Process one group only:

```bash
python run_claude.py --group "Degg"
```

---

## Requirements

- Windows (physical or VM)
- Python 3.8+
- DingTalk PC client installed (`exe_path` in config handles auto-launch)
- Google Drive for Desktop — signed in, target folder syncing as a mapped drive
- Anthropic API key (`sk-ant-api03-...`) from [console.anthropic.com](https://console.anthropic.com)

```
pip install -r requirements.txt
# installs: uiautomation, PyYAML, pyautogui, anthropic, mss, Pillow
```

---

## Key Technical Notes

### Why `mss` not `pyautogui` for screenshots

DingTalk's chat messages render inside a **CefBrowserWindow** (Chromium). This uses hardware acceleration (DirectX/DXGI), which makes the content invisible to GDI-based tools like `pyautogui.screenshot()` — it returns a blank white panel. `mss` uses DXGI screen capture and correctly captures CEF content.

### Why Claude Haiku not a full computer-use agent

The original implementation used Claude's computer-use API (autonomous agent loop: screenshot → decide → act → repeat). This worked but was expensive (~50 API calls/run) and slow.

The current approach asks Claude ONE targeted question per scroll position:
> *"Here is a screenshot of the DingTalk chat panel. Return JSON coordinates of every Download button you can see."*

Everything else (navigation, scrolling, clicking, file moving) is plain Python. This reduces Claude usage by ~100× and makes the code deterministic and debuggable.

### Why `claude-haiku-4-5-20251001` not Opus/Sonnet

Coordinate detection from a screenshot is a simple vision task. Haiku is ~20× cheaper than Sonnet and ~80× cheaper than Opus, with sufficient accuracy for finding labelled buttons.

### Search overlay workaround

DingTalk's newer UI opens a full-screen search overlay when the search box is clicked. The overlay covers the chat content but leaves group header buttons (Files, Group Settings) visible — causing naive verification to return false positives. The fix:

1. Type group name → press Down → Enter (selects result, navigates background chat)
2. Press Escape ×3 to dismiss the overlay
3. Verify by checking for group-specific header buttons

---

## Module Reference

| File | Purpose |
|------|---------|
| `run_claude.py` | **Main entry point** — hybrid programmatic + Claude vision loop |
| `dd_collector/chat_scanner.py` | Claude Haiku vision: screenshot → Download button coordinates |
| `dd_collector/dingtalk_ui.py` | All DingTalk UI interaction (UIA + pyautogui) |
| `dd_collector/file_mover.py` | Detect new downloads, move to GDrive with collision handling |
| `dd_collector/dedup.py` | JSON-based tracker to prevent re-downloading |
| `dd_collector/config.py` | Typed dataclasses for config.yaml |
| `dd_collector/vlm.py` | Legacy Qwen VLM helpers (OpenRouter) — still used by old `run.py` |
| `run.py` | Legacy entry point (Qwen VLM approach) |

---

## Google Drive File Organisation

```
G:\My Drive\DingTalk Files\       ← gdrive.base_path
 ├── Degg\
 │    └── 2026-02\
 │         ├── report.pdf
 │         └── data.xlsx
 └── 资料分享群\
      └── 2026-02\
           ├── 高盛亚洲交易台-260226.PDF
           └── 高盛亚洲交易台-260226_1.PDF   ← auto-increment on name collision
```

---

## Troubleshooting

**DingTalk window not found**
→ Set `dingtalk.exe_path` in `config.yaml`. The collector auto-launches DingTalk and waits up to 30s for the window.

**Navigation fails / "search box not found"**
→ DingTalk may still be showing startup dialogs. `wait_for_ready()` handles this automatically; if it times out (45s), start DingTalk manually and re-run.

**0 download buttons found**
→ Ensure `mss` is installed (`pip install mss`). The screenshot must use DirectX capture; GDI returns a blank panel for CEF content.

**Claude API key rejected**
→ Use a regular API key (`sk-ant-api03-...`) from [console.anthropic.com](https://console.anthropic.com). Claude Code OAuth tokens (`sk-ant-oat01-...`) are not accepted by the REST API.

**DingTalk UI changes break navigation**
→ Run `python tools/inspect_dingtalk.py` to explore the current control tree. Update `ui_selectors` in `config.yaml` as needed.
