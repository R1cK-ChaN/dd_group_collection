# DingTalk Group File Collector

Automated file collection from DingTalk group chats, synced to Google Drive.

Uses RPA-style UI automation (`uiautomation`) to drive the DingTalk PC client — no API keys or admin permissions required.

## Requirements

- Windows (any — physical or VM)
- Python 3.8+
- DingTalk PC client — logged in and running
- Google Drive for Desktop — signed in, with the target folder syncing

## Quick Start

```bash
git clone https://github.com/user/dd_group_collection.git
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
                    │    │    4. list_files()                          │ │    │
                    │    │    5. filter via dedup  ─── dedup.py        │ │    │
                    │    │    6. download_file()   ─── dingtalk_ui.py  │ │    │
                    │    │    7. move to GDrive    ─── file_mover.py   │ │    │
                    │    │    8. mark downloaded   ─── dedup.py        │ │    │
                    │    └────────────────────────────────────────────┘ │    │
                    │          │                                               │
                    │          v                                               │
                    │    sleep(interval_minutes)                                │
                    └──────────────────────────────────────────────────────────┘
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
| `dingtalk` | Window class, download directory, timeouts |
| `groups` | List of group names and folder aliases |
| `gdrive` | Google Drive base path for organized file storage |
| `polling` | Interval and per-group download cap |
| `dedup` | Path to the JSON download tracker |
| `logging` | Log directory, level, rotation settings |
| `ui_selectors` | DingTalk UI control mappings (update after DingTalk UI changes) |

## File Organization on Google Drive

```
G:\My Drive\DingTalk Files\         <- gdrive.base_path
 ├── ExampleGroup1/                 <- group alias
 │    ├── 2026-01/
 │    │    ├── report.pdf
 │    │    └── data.xlsx
 │    └── 2026-02/
 │         └── summary.docx
 └── ExampleGroup2/
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

Update the `ui_selectors` section in `config.yaml` with the new control types and names.
