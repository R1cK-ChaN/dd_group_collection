# DingTalk Group File Collector

Automated file collection from DingTalk (钉钉) group chats, synced to Google Drive.

Uses RPA-style UI automation (`uiautomation`) to drive the DingTalk PC client — no API keys or admin permissions required.

## Architecture

```
┌─────────────┐     ┌──────────────────────────────────────────────────────────┐
│  start.bat  │────▶│  run.py                                                  │
│ (auto-restart)    │    └─▶ main.py  (polling loop)                           │
└─────────────┘     │          │                                               │
                    │          ▼                                               │
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
                    │          ▼                                               │
                    │    sleep(interval_minutes)                                │
                    └──────────────────────────────────────────────────────────┘

                    ┌────────────────────────┐       ┌─────────────────────┐
                    │  DingTalk PC Client     │       │  Google Drive for   │
                    │  (UI driven via         │       │  Desktop            │
                    │   uiautomation)         │       │  (file sync)        │
                    └────────┬───────────────┘       └──────────▲──────────┘
                             │ downloads to                     │ moved to
                             ▼                                  │
                    ┌────────────────────────────────────────────┘
                    │  DingTalk download dir ──▶ GDrive/{Alias}/{YYYY-MM}/
                    └─────────────────────────────────────────────
```

## Module Dependency Graph

```
run.py
 └─▶ dd_collector/main.py          # orchestration loop
      ├─▶ config.py                 # load config.yaml → dataclasses
      ├─▶ logger.py                 # rotating file + console logging
      ├─▶ dedup.py                  # JSON-based download tracker
      ├─▶ file_mover.py            # detect new files, move to GDrive
      ├─▶ dingtalk_ui.py           # all DingTalk UI interaction
      │    └─▶ ui_helpers.py       # reusable uiautomation wrappers
      └─▶ ui_helpers.py            # send_escape() for error recovery

tools/inspect_dingtalk.py           # standalone — no internal deps
```

### Module Responsibilities

| Module | Purpose |
|--------|---------|
| `config.py` | Loads `config.yaml` into typed dataclasses (`AppConfig`, `GroupConfig`, etc.). All user-editable settings live in YAML, not in code. |
| `logger.py` | Sets up a `RotatingFileHandler` (default 10 MB, 5 backups) + `StreamHandler` for console output. |
| `ui_helpers.py` | Low-level `uiautomation` wrappers: `find_control()`, `safe_click()`, `safe_right_click()`, `set_text()`, `scroll_to_bottom()`, `send_escape()`. Resolution-independent (control-based, not coordinate-based). |
| `dingtalk_ui.py` | `DingTalkController` class — all DingTalk-specific UI interaction. When DingTalk updates its UI, this is the only Python file that needs editing. Uses selectors from `config.yaml` so many changes require only a config update. |
| `file_mover.py` | Scans the DingTalk download directory for new files (skipping `.tmp`/`.partial`, waiting for writes to settle), moves them to `{GDrive}/{GroupAlias}/{YYYY-MM}/{filename}` with auto-increment on collisions. |
| `dedup.py` | `DedupTracker` — JSON file keyed by `group::filename`. Prevents re-downloading files across cycles. Recovers gracefully from corrupt JSON. |
| `main.py` | Polling loop with 5-layer error handling: action → file → group → cycle → process (`.bat` restart). Caps downloads per group per cycle. |
| `tools/inspect_dingtalk.py` | Standalone diagnostic script. Dumps DingTalk's control tree with `--depth` and `--search` filters. Essential for finding selectors after DingTalk UI updates. |

## Error Handling Layers

```
start.bat                       ← Layer 5: process restart (30s delay)
 └─ main._polling_loop          ← Layer 4: cycle-level catch (log & continue)
     └─ main._run_cycle         ← Layer 3: group-level catch (Esc recovery)
         └─ main._process_group ← Layer 2: file-level catch (skip file)
             └─ dingtalk_ui.*   ← Layer 1: action-level (return False)
```

## File Organization on Google Drive

```
G:\My Drive\DingTalk Files\         ← gdrive.base_path
 ├── ExampleGroup1/                 ← group alias
 │    ├── 2026-01/
 │    │    ├── report.pdf
 │    │    └── data.xlsx
 │    └── 2026-02/
 │         └── summary.docx
 └── ExampleGroup2/
      └── 2026-02/
           ├── notes.pdf
           └── notes_1.pdf          ← auto-increment on collision
```

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| [uiautomation](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows) | ≥ 2.0.20 | Windows UI Automation (UIA) wrapper. Finds and interacts with controls by type, name, and automation ID — resolution-independent unlike coordinate-based tools (PyAutoGUI). Windows-only. |
| [PyYAML](https://pyyaml.org/) | ≥ 6.0 | Parse `config.yaml`. Chosen over TOML/INI for nested structure support and readability. |

**Standard library only** (no extra install): `json` (dedup storage), `shutil` (file moves), `logging` (rotating logs), `dataclasses` (typed config), `pathlib`, `argparse`.

**Runtime requirements:**
- Windows (the `uiautomation` package wraps the Windows UIA COM interface)
- Python 3.8+
- DingTalk PC client — logged in and running
- Google Drive for Desktop — signed in, with the target folder syncing

## Setup

```bash
pip install -r requirements.txt
```

## Configuration

Edit `config.yaml`:

1. Set `dingtalk.download_dir` to your DingTalk download path (check DingTalk Settings → File Management)
2. Set `gdrive.base_path` to your Google Drive sync folder
3. Add your groups under `groups:` — `name` must match the exact DingTalk display name, `alias` is used as the folder name
4. Adjust `polling.interval_minutes` and `polling.max_downloads_per_group` as needed
5. If DingTalk updates its UI, run `inspect_dingtalk.py` and update `ui_selectors`

## Usage

**First run — verify DingTalk UI access:**

```bash
python tools/inspect_dingtalk.py
python tools/inspect_dingtalk.py --search "文件"
python tools/inspect_dingtalk.py --depth 8
```

**Run the collector:**

```bash
python run.py
```

**Auto-start on Windows boot:**

Copy `start.bat` to `shell:startup` (press `Win+R`, type `shell:startup`). The batch file auto-restarts the collector with a 30-second delay on exit.

## Design Decisions

| Decision | Rationale |
|----------|-----------|
| UI selectors in `config.yaml` | DingTalk updates frequently; fix breakages without editing Python |
| Single `dingtalk_ui.py` | One file to review/diff when DingTalk UI changes |
| JSON dedup (not SQLite) | Human-readable, no extra deps, sufficient for hundreds/thousands of files |
| Polling (not filesystem watcher) | The script must actively trigger downloads; filesystem watching only helps for the move step |
| `uiautomation` (not PyAutoGUI) | Control-based, not coordinate-based — works across resolutions and DPI settings |
| `shutil.move` (not copy) | Avoids doubling disk usage; DingTalk download dir is a transient staging area |
