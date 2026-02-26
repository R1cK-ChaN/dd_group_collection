"""Optimized DingTalk file collector — hybrid programmatic + Claude vision.

Architecture
============
Every step that does NOT require understanding a screenshot is done in plain
Python (uiautomation + pyautogui).  Claude is called ONLY for the one thing
Python can't do: reading the Chrome-Embedded-Framework chat panel to find
Download button coordinates.

Per-group cost:
  Programmatic (free):  navigate, focus, scroll, click, wait, move files
  Claude API (paid):    1 call per scroll position to identify Download buttons
                        ≈ 3-8 calls/group @ ~880 tokens each  (~$0.003-0.007)
  vs old autonomous agent: 30-60 calls @ ~8000 tokens each  (~$0.45-0.90)

Usage
=====
    python run_claude.py                      # process all groups once (Haiku scanner)
    python run_claude.py --loop               # poll on configured interval
    python run_claude.py --group Degg         # one group only
    python run_claude.py --trigger            # watch for unread, fire on detection
    python run_claude.py --trigger --group Degg
    python run_claude.py --autonomous         # full Claude Opus computer-use agent
    python run_claude.py --autonomous --group Degg
"""

from __future__ import annotations

import io
import logging
import sys
import time
from pathlib import Path
from typing import Optional, Set

# Windows console UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, str(Path(__file__).parent))

from dd_collector.chat_scanner import ChatScanner
from dd_collector.claude_agent import ClaudeAgent
from dd_collector.config import load_config
from dd_collector.dedup import DedupTracker
from dd_collector.dingtalk_ui import DingTalkController
from dd_collector.file_mover import get_new_files, move_file_to_gdrive
from dd_collector.trigger_watcher import TriggerWatcher


# ── Logging ───────────────────────────────────────────────────

def _setup_logging(cfg) -> None:
    log_dir = Path(cfg.logging.dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    level = getattr(logging, cfg.logging.level.upper(), logging.INFO)
    handlers = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(log_dir / "run_claude.log", encoding="utf-8"),
    ]
    logging.basicConfig(
        level=level,
        format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
        handlers=handlers,
        force=True,
    )
    for lib in ("httpx", "httpcore", "anthropic._base_client"):
        logging.getLogger(lib).setLevel(logging.WARNING)


log = logging.getLogger("dd_collector")


# ── Helpers ───────────────────────────────────────────────────

def _snapshot_dir(download_dir: str) -> Set[str]:
    p = Path(download_dir)
    if not p.is_dir():
        return set()
    return {f.name for f in p.iterdir() if f.is_file()}


# ── Per-group collection ──────────────────────────────────────

def _process_group(
    ctrl: DingTalkController,
    scanner: ChatScanner,
    dedup: DedupTracker,
    group_name: str,
    group_alias: str,
    download_dir: str,
    gdrive_base: str,
    max_scrolls: int,
    download_wait: int,
    max_downloads: int,
) -> int:
    """Navigate to a group, scan for file cards, download new ones.

    Returns the number of files successfully moved to GDrive.
    """
    # ── Step 1: Navigate (programmatic) ──────────────────────
    log.info("Navigating to group: %s", group_name)
    if not ctrl.navigate_to_group(group_name):
        log.error("Navigation failed for group: %s — skipping.", group_name)
        return 0

    time.sleep(3.0)  # let CefBrowserWindow fully render chat content

    # Snapshot download dir before we start clicking
    before = _snapshot_dir(download_dir)

    # Files already in dedup for this group (skip re-downloading)
    already = set(dedup.get_downloaded_for_group(group_name))
    log.info("Already downloaded for '%s': %d files", group_name, len(already))

    # Track coords seen this run to avoid double-clicking same button
    seen_coords: Set[tuple] = set()
    total_clicks = 0

    # ── Step 2: Scroll + Scan + Click loop ───────────────────
    for scroll_i in range(max_scrolls + 1):
        log.info(
            "Scan pass %d/%d for group '%s'",
            scroll_i + 1, max_scrolls + 1, group_name,
        )

        # Take chat panel screenshot (programmatic — free)
        screenshot_b64, offset_x, offset_y = ctrl.get_chat_panel_screenshot()

        # Ask Claude Haiku to find Download buttons (1 API call)
        downloads = scanner.find_downloads(
            screenshot_b64, region_offset=(offset_x, offset_y)
        )

        for item in downloads:
            fname = item["filename"]
            x, y = item["x"], item["y"]

            # Skip if already downloaded
            if fname in already:
                log.info("  skip (dedup): %s", fname)
                continue

            # Skip if we already clicked this exact pixel this run
            coord_key = (x, y)
            if coord_key in seen_coords:
                log.debug("  skip (already clicked this run): (%d,%d)", x, y)
                continue
            seen_coords.add(coord_key)

            # Safety cap
            if total_clicks >= max_downloads:
                log.warning(
                    "  reached max_downloads=%d — stopping.", max_downloads
                )
                break

            # Click Download button (programmatic — free)
            log.info("  downloading: %s  at (%d, %d)", fname, x, y)
            ctrl.click_download_at(x, y)
            total_clicks += 1
            time.sleep(0.3)

        # Scroll up to reveal older messages (programmatic — free)
        if scroll_i < max_scrolls:
            ctrl.scroll_chat_up()
            time.sleep(1.2)

    log.info(
        "Scan complete for '%s': %d download click(s) across %d passes.",
        group_name, total_clicks, max_scrolls + 1,
    )

    # ── Step 3: Wait for downloads to settle ─────────────────
    if total_clicks > 0:
        log.info("Waiting %ds for downloads to finish writing…", download_wait)
        time.sleep(download_wait)

    # ── Step 4: Detect new files (programmatic) ──────────────
    new_files = get_new_files(download_dir, before)
    log.info("New files in download dir: %d", len(new_files))

    # ── Step 5: Move to GDrive + record dedup (programmatic) ─
    moved = 0
    for fpath in new_files:
        if dedup.is_downloaded(group_name, fpath.name):
            log.info("  skip (dedup, already moved): %s", fpath.name)
            continue
        try:
            dest = move_file_to_gdrive(
                file_path=fpath,
                group_folder_name=group_alias,
                gdrive_base_path=gdrive_base,
            )
            dedup.mark_downloaded(group_name, fpath.name, dest)
            log.info("  moved: %s → %s", fpath.name, dest)
            moved += 1
        except Exception:
            log.exception("  failed to move: %s", fpath.name)

    return moved


# ── One full cycle ────────────────────────────────────────────

def run_once(
    cfg,
    ctrl: DingTalkController,
    scanner: ChatScanner,
    dedup: DedupTracker,
    only_group: Optional[str] = None,
) -> None:
    """Process all configured groups (or just *only_group*)."""
    for group in cfg.groups:
        if only_group and group.name != only_group:
            continue

        log.info("=" * 60)
        log.info("Group: %s  (alias: %s)", group.name, group.alias)
        log.info("=" * 60)

        moved = _process_group(
            ctrl=ctrl,
            scanner=scanner,
            dedup=dedup,
            group_name=group.name,
            group_alias=group.alias or group.name,
            download_dir=cfg.dingtalk.download_dir,
            gdrive_base=cfg.gdrive.base_path,
            max_scrolls=cfg.claude.max_scrolls,
            download_wait=cfg.dingtalk.download_wait,
            max_downloads=cfg.polling.max_downloads_per_group,
        )

        log.info("Group '%s' done — %d file(s) moved to GDrive.", group.name, moved)


# ── Entry point ───────────────────────────────────────────────

def main() -> None:
    cfg = load_config()
    _setup_logging(cfg)

    log.info("DingTalk Collector (optimized hybrid) starting")

    if not cfg.claude.oauth_token:
        log.error(
            "claude.oauth_token missing from config.yaml. "
            "Set it or export ANTHROPIC_AUTH_TOKEN."
        )
        sys.exit(1)

    # Ensure DingTalk is running — auto-launch if needed, then connect
    ctrl = DingTalkController(cfg)
    if not ctrl.ensure_running():
        log.error(
            "DingTalk is not running and could not be launched. "
            "Start it manually or set dingtalk.exe_path in config.yaml."
        )
        sys.exit(1)
    if not ctrl.connect():
        log.error("Cannot connect to DingTalk window.")
        sys.exit(1)
    if not ctrl.wait_for_ready():
        log.error("DingTalk UI did not become ready in time.")
        sys.exit(1)

    scanner = ChatScanner(
        api_key=cfg.claude.oauth_token,
        model=cfg.claude.model,
    )
    log.info("Scanner model: %s", cfg.claude.model)

    dedup = DedupTracker(cfg.dedup.path)

    # Parse CLI flags
    only_group: Optional[str] = None
    loop_mode = False
    trigger_mode = False
    autonomous_mode = False
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--loop":
            loop_mode = True
        elif args[i] == "--trigger":
            trigger_mode = True
        elif args[i] == "--autonomous":
            autonomous_mode = True
        elif args[i] == "--group" and i + 1 < len(args):
            only_group = args[i + 1]
            i += 1
        i += 1

    if loop_mode:
        interval = cfg.polling.interval_minutes * 60
        log.info(
            "Loop mode: running every %d minutes. Ctrl+C to stop.",
            cfg.polling.interval_minutes,
        )
        while True:
            try:
                run_once(cfg, ctrl, scanner, dedup, only_group=only_group)
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
            except Exception:
                log.exception("Cycle error — will retry next interval.")
            log.info("Sleeping %ds…", interval)
            try:
                time.sleep(interval)
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
    elif trigger_mode:
        watcher = TriggerWatcher(cfg)
        log.info("Trigger mode: watching for unread messages every %ds. Ctrl+C to stop.",
                 cfg.trigger.check_interval_seconds)
        while True:
            try:
                reason = watcher.wait_for_trigger()
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
            log.info("Trigger fired (%s) — starting download cycle.", reason)
            try:
                run_once(cfg, ctrl, scanner, dedup, only_group=only_group)
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
            except Exception:
                log.exception("Cycle error after trigger — resuming watch.")
            log.info("Cooling down %ds before resuming watch…", cfg.trigger.cooldown_seconds)
            try:
                time.sleep(cfg.trigger.cooldown_seconds)
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
    elif autonomous_mode:
        # Prefer OpenRouter key (vlm.api_key) if present — avoids Anthropic credit dependency.
        # Fall back to claude.oauth_token for direct Anthropic API access.
        if cfg.vlm.api_key and cfg.vlm.api_key.startswith("sk-or-"):
            agent_key = cfg.vlm.api_key
            # Anthropic SDK appends /v1/messages itself, so base_url must NOT end in /v1
            agent_base_url = cfg.vlm.base_url.rstrip("/")
            if agent_base_url.endswith("/v1"):
                agent_base_url = agent_base_url[:-3]   # "https://openrouter.ai/api"
            agent_model = "anthropic/claude-opus-4-5-20251101"
            log.info("Autonomous mode: using OpenRouter key → %s", agent_model)
        else:
            agent_key = cfg.claude.oauth_token
            agent_base_url = None
            agent_model = "claude-opus-4-5-20251101"
            log.info("Autonomous mode: using Anthropic key → %s", agent_model)
        agent = ClaudeAgent(
            oauth_token=agent_key,
            model=agent_model,
            log_dir=cfg.logging.dir,
            base_url=agent_base_url,
        )
        for group in cfg.groups:
            if only_group and group.name != only_group:
                continue
            log.info("=" * 60)
            log.info("Autonomous: Group '%s'  (alias: %s)", group.name, group.alias)
            log.info("=" * 60)

            # Activate DingTalk window before handing off to Claude
            try:
                ctrl._window.SetActive()
                ctrl._window.SetFocus()
                time.sleep(0.5)
            except Exception:
                pass

            before = _snapshot_dir(cfg.dingtalk.download_dir)
            already = dedup.get_downloaded_for_group(group.name)
            log.info("Already downloaded for '%s': %d file(s)", group.name, len(already))

            try:
                agent.run_download_task(
                    group_name=group.name,
                    download_dir=cfg.dingtalk.download_dir,
                    already_downloaded=already,
                    max_scrolls=cfg.claude.max_scrolls,
                )
            except KeyboardInterrupt:
                log.info("Stopped by user.")
                return
            except Exception:
                log.exception("Autonomous agent error for group '%s' — continuing.", group.name)

            log.info("Waiting %ds for downloads to settle…", cfg.dingtalk.download_wait)
            time.sleep(cfg.dingtalk.download_wait)

            new_files = get_new_files(cfg.dingtalk.download_dir, before)
            log.info("New files detected: %d", len(new_files))
            for fpath in new_files:
                if dedup.is_downloaded(group.name, fpath.name):
                    log.info("  skip (dedup): %s", fpath.name)
                    continue
                try:
                    dest = move_file_to_gdrive(
                        file_path=fpath,
                        group_folder_name=group.alias or group.name,
                        gdrive_base_path=cfg.gdrive.base_path,
                    )
                    dedup.mark_downloaded(group.name, fpath.name, dest)
                    log.info("  moved: %s → %s", fpath.name, dest)
                except Exception:
                    log.exception("  failed to move: %s", fpath.name)

        log.info("Autonomous mode complete.")
    else:
        run_once(cfg, ctrl, scanner, dedup, only_group=only_group)
        log.info("All done.")


if __name__ == "__main__":
    main()
