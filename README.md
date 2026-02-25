# DingTalk Group File Collector

Automated file collection from DingTalk group chats, synced to Google Drive.

Uses RPA-style UI automation (`uiautomation`) to drive the DingTalk PC client — no API keys or admin permissions required.

## Deployment Status (2026-02-25)

**Environment:** Windows 11 VM, Python 3.14.3, DingTalk English version, Google Drive for Desktop (G: drive)

| Step | Status | Notes |
|------|--------|-------|
| Clone & install deps | Done | `uiautomation 2.0.29`, `PyYAML 6.0.3` |
| Config paths | Done | download_dir, gdrive base_path set for user `vm` |
| Google Drive sync | Verified | Files placed in `G:\My Drive\DingTalk Files\` sync to cloud |
| Auto-start on boot | Done | `dd_group_collection.bat` in `shell:startup` |
| Connect to DingTalk | Working | Window class `DtMainFrameView` (was `StandardFrame_DingTalk`) |
| Navigate to group | Working | Search box `QLineEdit` → `SendKeys` → `Enter` selects first result |
| Open Files tab | Working | `ButtonControl Name='Files'` (was `'File'`) in group header |
| List files | **Blocked** | CefBrowserWindow accessibility tree still empty after registry fix |
| Download files | **Blocked** | Depends on list_files; pyautogui hover approach ready but untested |
| Move to GDrive | Not yet tested | Depends on download working |

## Known Issues & Challenges

### 1. CefBrowserWindow accessibility tree not exposed (primary blocker)

The Files tab renders inside a Chromium Embedded Framework web view (`CefBrowserWindow`). The file grid (`DocumentControl` → `GroupControl Name='grid'` → `CustomControl` rows) was visible in an earlier session but is **no longer accessible** — the CefBrowserWindow now returns only an empty `CustomControl` with no children.

**What was tried:**
- Sending `WM_GETOBJECT` to `Chrome_WidgetWin_1` — returned non-zero but tree stayed empty
- Starting Windows Narrator to trigger Chromium accessibility — no effect
- Searching with `searchDepth=15` — no deeper controls found
- Checking for Chrome DevTools Protocol (CDP) ports — none open
- **Registry fix (2026-02-25):** Set `HKCU\Software\Chromium\Accessibility\AXMode=1` and `HKCU\Software\Google\Chrome\Accessibility\AXMode=1`, restarted DingTalk — **tree still empty**. DingTalk's CEF build appears to ignore these standard Chromium accessibility flags.

**Current CefBrowserWindow tree (still empty after registry fix):**
```
PaneControl (Class='CefBrowserWindow')
 └─ PaneControl (Class='Chrome_WidgetWin_1')
     └─ PaneControl
         └─ CustomControl (empty — no Name, no children)
```

**Native controls that ARE visible** (proving the files tab is open):
- `ButtonControl Name='Upload File'` at (594,626)
- `ButtonControl Name='Files'` at (986,106)

**Remaining solutions to investigate:**
1. ~~Set Chromium accessibility via registry~~ — tried, did not work
2. ~~Launch DingTalk with `--force-renderer-accessibility` flag~~ — DingTalk launcher does not pass flags to CEF
3. **Use screenshot + OCR to identify file rows visually and click by coordinates** — next approach
4. Inject CDP access by modifying DingTalk's CEF launch flags
5. Use keyboard navigation (Tab/Arrow) to select files within the web view

### 2. Download mechanism — pyautogui hover approach implemented but blocked by #1

`_download_via_hover_pyautogui()` was added to use `pyautogui` real mouse events:
1. Gets the file row's `BoundingRectangle` (requires accessibility tree from #1)
2. Moves the real mouse cursor to the row center (triggers Chromium hover state)
3. Checks for any newly accessible download button children
4. If none found, clicks at `(row_right - offset, row_center_y)` where the download icon appears

The offset is configurable via `dingtalk.download_icon_offset` in `config.yaml` (default: 95px). Cannot be tested until the file listing (Issue #1) is resolved.

### 3. Dialog dismissal was closing DingTalk

The `dismiss_buttons` config originally included `"Close"` and `"关闭"`, which matched the **window title bar close button** (`ButtonControl Name='Close'`). This caused the collector to close DingTalk entirely. Fixed by removing these from `config.yaml`.

### 4. Second group navigation can fail

Searching for `资料分享群` returned no results in one test. Root cause identified: leftover text in the search box from the previous search. **Fixed** by clearing the search box with `Ctrl+A` → `Delete` before typing the new group name.

### 5. Python 3.14 — Pillow / pyscreeze not supported

`pyautogui.screenshot()` fails because Pillow has not released a build compatible with Python 3.14. This blocks any screenshot-based approach using pyautogui. Alternatives:
- Use `mss` (pure-Python screenshot library) instead of Pillow
- Use Windows native screenshot via `ctypes` / `win32api`
- Downgrade to Python 3.12/3.13

## Requirements

- Windows (any — physical or VM)
- Python 3.8+
- DingTalk PC client — logged in and running
- Google Drive for Desktop — signed in, with the target folder syncing

## Quick Start

```bash
git clone https://github.com/R1cK-ChaN/dd_group_collection.git
cd dd_group_collection
pip install -r requirements.txt
```

Edit `config.yaml`:

1. Set `dingtalk.download_dir` to your DingTalk download path
2. Set `gdrive.base_path` to your Google Drive sync folder
3. Add your groups under `groups:`

Verify DingTalk UI access:

```bash
python tools/inspect_dingtalk.py
```

Run the collector:

```bash
python run.py
```

## Auto-start on Windows Boot

Copy `start.bat` to `shell:startup` (press `Win+R`, type `shell:startup`). The batch file auto-restarts the collector with a 30-second delay on exit.

Note: the startup copy must `cd /d` to the project directory. The deployed version at `shell:startup\dd_group_collection.bat` already does this.

## Architecture

```
┌─────────────┐     ┌──────────────────────────────────────────────────────────┐
│  start.bat  │────>│  run.py                                                  │
│ (auto-restart)    │    └─> main.py  (polling loop)                           │
└─────────────┘     │          │                                               │
                    │          v                                               │
                    │    ┌─── per cycle ──────────────────────────────────┐    │
                    │    │  for each group in config.yaml:               │    │
                    │    │    1. connect()        ─── dingtalk_ui.py ──┐ │    │
                    │    │    2. navigate_to_group()                   │ │    │
                    │    │    3. open_files_tab()                      │ │    │
                    │    │    4. list_files()          ** web view **  │ │    │
                    │    │    5. filter via dedup  ─── dedup.py        │ │    │
                    │    │    6. download_file()   ─── dingtalk_ui.py   │ │    │
                    │    │    7. move to GDrive    ─── file_mover.py   │ │    │
                    │    │    8. mark downloaded   ─── dedup.py        │ │    │
                    │    └────────────────────────────────────────────┘ │    │
                    │          │                                               │
                    │          v                                               │
                    │    sleep(interval_minutes)                                │
                    └──────────────────────────────────────────────────────────┘
```

## DingTalk UI Control Tree

Discovered via `tools/inspect_dingtalk.py` and manual inspection (updated 2026-02-25).

### Current layout (DtMainFrameView)

```
WindowControl Name='DingTalk' Class='DtMainFrameView'
 ├─ WindowControl Class='DingChatWnd'
 │   ├─ WindowControl Name='ConvTabListView' Class='ConvListView'
 │   │   └─ GroupControl > QStackedWidget > GroupControl
 │   │       └─ (conversation items — Names are EMPTY, not accessible)
 │   ├─ WindowControl Name='ConvTabTopBar' Class='ConvTabTopBarV2'
 │   │   └─ QStackedWidget > GroupControl
 │   └─ WindowControl (right panel — group chat content)
 │       ├─ WindowControl Name='DTIMContentModule'              ← chat messages
 │       ├─ ButtonControl Name='Group Notice'                   ← group header
 │       ├─ ButtonControl Name='Files'                          ← ** FILES TAB **
 │       ├─ ButtonControl Name='Chat History'
 │       ├─ ButtonControl Name='More'
 │       └─ ButtonControl Name='Group Settings'
 ├─ GroupControl Class='QWidget'
 │   ├─ GroupControl Class='client_ding::TitlebarView'
 │   │   └─ EditControl Class='QLineEdit'                      ← ** SEARCH BOX ** (no Name!)
 │   ├─ GroupControl Class='client_ding::NavigatorView'
 │   │   ├─ ButtonControl Name='Standard Edition'
 │   │   ├─ ButtonControl Name='Messages'
 │   │   └─ ButtonControl Name='More'
 │   └─ GroupControl Class='main_frame::DtContentAreaView'
 └─ GroupControl Class='ddesign::TopWindowToolBar'
     ├─ ButtonControl Name='Minimize'
     ├─ ButtonControl Name='Maximize'
     └─ ButtonControl Name='Close'
```

### Files tab (CefBrowserWindow — accessibility tree EMPTY)

```
PaneControl Class='CefBrowserWindow'  (inside right panel after clicking Files)
 └─ PaneControl Class='Chrome_WidgetWin_1'
     └─ PaneControl
         └─ CustomControl  (no Name, no children — NOT ACCESSIBLE)
```

### Legacy layout (StandardFrame_DingTalk) — no longer seen

```
WindowControl (Class='StandardFrame_DingTalk')
 ├─ WindowControl (Class='ChatFileWnd')
 │   └─ PaneControl (Class='CefBrowserWindow')
 │       └─ DocumentControl (Class='Chrome_RenderWidgetHostHWND', Name='群文件-Online')
 │           └─ GroupControl Name='grid'                        ← file list (was accessible)
 └─ EditControl Name='Search'                                  ← search box (had a Name)
```

### Key behavioral notes

- **Search box** has no Name — must be found by `ClassName='QLineEdit'`.
- **Conversation list items** all have empty Name attributes — cannot be enumerated.
- **Navigation** works by typing in the search box via `SendKeys`, then pressing `Enter`
  (via pyautogui) to select the first search result.
- After clicking the search box, the foreground window changes to `DtQtWebView`
  (a search overlay), so pyautogui keyboard events land there correctly.
- DingTalk must have **both** `SetActive()` and `SetFocus()` before pyautogui
  mouse/keyboard events will reach it.

## Module Dependency Graph

```
run.py
 └─> dd_collector/main.py          # orchestration loop
      ├─> config.py                 # load config.yaml → dataclasses
      ├─> logger.py                 # rotating file + console logging
      ├─> dedup.py                  # JSON-based download tracker
      ├─> file_mover.py            # detect new files, move to GDrive
      ├─> dingtalk_ui.py           # all DingTalk UI interaction
      │    └─> ui_helpers.py       # reusable uiautomation wrappers
      └─> ui_helpers.py            # send_escape() for error recovery

tools/inspect_dingtalk.py           # standalone — no internal deps
```

## Configuration Reference

All settings live in `config.yaml`. Key sections:

| Section | Purpose |
|---------|---------|
| `dingtalk` | Window class, download directory, timeouts, download icon offset |
| `groups` | List of group names and folder aliases |
| `gdrive` | Google Drive base path for organized file storage |
| `polling` | Interval and per-group download cap |
| `dedup` | Path to the JSON download tracker |
| `logging` | Log directory, level, rotation settings |
| `ui_selectors` | DingTalk UI control mappings (update after DingTalk UI changes) |

## File Organization on Google Drive

```
G:\My Drive\DingTalk Files\         <- gdrive.base_path
 ├── Degg/                          <- group alias
 │    └── 2026-02/
 │         ├── report.pdf
 │         └── data.xlsx
 └── 资料分享群/
      └── 2026-02/
           ├── notes.pdf
           └── notes_1.pdf          <- auto-increment on collision
```

## Troubleshooting UI Selectors

When DingTalk updates its UI, selectors may break. To find new values:

```bash
python tools/inspect_dingtalk.py
python tools/inspect_dingtalk.py --search "File"
python tools/inspect_dingtalk.py --depth 8
```

Use `-X utf8` flag if you see `UnicodeEncodeError` with Chinese characters:

```bash
python -X utf8 tools/inspect_dingtalk.py --depth 8 --search "File"
```

Update the `ui_selectors` section in `config.yaml` with the new control types and names.
