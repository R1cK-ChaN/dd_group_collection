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
| Connect to DingTalk | Working | Window class `StandardFrame_DingTalk` found |
| Navigate to group | Working | Search box (`EditControl Name='Search'`) works; `set_text` falls back to `SendKeys` |
| Open Files tab | Working | `ButtonControl Name='File'` clicks successfully |
| List files | Partially working | See **Known Issues #1** below |
| Download files | Working | Uses `pyautogui` hover + click (see **Known Issues #2**) |
| Move to GDrive | Not yet tested | Depends on download working |

## Known Issues & Challenges

### 1. File list is inside a CefBrowserWindow (web view)

The Files tab renders its content in a Chromium Embedded Framework web view, not native Windows controls.

**Original assumption:** File list uses native `ListControl` / `ListItemControl`.
**Reality:** The actual file entries are:

```
DocumentControl (ClassName='Chrome_RenderWidgetHostHWND', Name='群文件-Online')
 └─ GroupControl (AutoId='root')
     └─ ... → TableControl
              └─ GroupControl (Name='grid')
                  └─ GroupControl (container)
                      ├─ CustomControl (Name=' filename.pdf 125.1 KB  ·2026/02/21 10:50Author')
                      ├─ CustomControl (Name=' filename2.pdf ...')
                      └─ ...
```

The original `list_files()` was matching the **chat sidebar** (`ListItemControl` from the conversation list) instead of actual files. A fix has been applied to `dingtalk_ui.py` to navigate into the web view and read `CustomControl` items from the `grid` `GroupControl`. File names are parsed from the `Name` property using regex.

### 2. Download mechanism — fixed with pyautogui hover

The file list's hover-revealed download icons are web-rendered inside the CefBrowserWindow and invisible to Windows UI Automation. The original context menu and hover button strategies both failed.

**Fix:** `_download_via_hover_pyautogui()` uses `pyautogui` to generate real mouse events:
1. Gets the file row's `BoundingRectangle` (visible to uiautomation)
2. Moves the real mouse cursor to the row center (triggers Chromium hover state)
3. Checks for any newly accessible download button children
4. If none found, clicks at `(row_right - offset, row_center_y)` where the download icon appears

The offset is configurable via `dingtalk.download_icon_offset` in `config.yaml` (default: 95px). If DingTalk changes its hover icon layout, adjust this value. The legacy context menu and hover button strategies are kept as fallbacks.

### 3. Dialog dismissal was closing DingTalk

The `dismiss_buttons` config originally included `"Close"` and `"关闭"`, which matched the **window title bar close button** (`ButtonControl Name='Close'`). This caused the collector to close DingTalk entirely. Fixed by removing these from `config.yaml`.

### 4. Second group navigation can fail

Searching for `资料分享群` returned no results in one test. May be caused by leftover text in the search box or timing issues. Needs investigation.

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

## DingTalk UI Control Tree (Files Tab)

Discovered via `tools/inspect_dingtalk.py` and manual inspection:

```
WindowControl (Class='StandardFrame_DingTalk')
 ├─ WindowControl (Class='ChatFileWnd')          ← Files tab window
 │   └─ PaneControl (Class='CefBrowserWindow')   ← Chromium web view
 │       └─ DocumentControl (Class='Chrome_RenderWidgetHostHWND', Name='群文件-Online')
 │           ├─ GroupControl (AutoId='root')
 │           │   ├─ TextControl Name='File' / 'Media' / 'Link'   ← tab filters
 │           │   ├─ ButtonControl Name='Upload/Create'
 │           │   ├─ CheckBoxControl / HeaderControl               ← select-all / sort
 │           │   ├─ GroupControl Name='grid'                      ← ** file list **
 │           │   │   └─ GroupControl (container)
 │           │   │       ├─ CustomControl Name='  filename.pdf 125.1 KB  ·date author'
 │           │   │       └─ ...
 │           │   ├─ EditControl Name='Search files'
 │           │   └─ TextControl Name='Recycle Bin'
 │           └─ GroupControl (AutoId='transfer-file-task-container')
 └─ WindowControl (title bar)
     ├─ EditControl Name='Search'                 ← main search box
     ├─ ButtonControl Name='Minimize' / 'Maximize' / 'Close'  ← title bar buttons
     └─ ...
```

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
