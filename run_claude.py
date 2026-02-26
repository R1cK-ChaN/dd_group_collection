"""Entry point: Claude computer-use agent for DingTalk file download.

Usage:
    python run_claude.py            # Run once then exit
    python run_claude.py --loop     # Run, then poll on interval
    python run_claude.py --group Degg   # Process only one specific group
"""

from __future__ import annotations

import io
import logging
import os
import sys
import time
from pathlib import Path
from typing import Optional, Set

# Fix stdout encoding on Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# Project root on path
sys.path.insert(0, str(Path(__file__).parent))

from dd_collector.config import load_config
from dd_collector.dedup import DedupTracker
from dd_collector.file_mover import get_new_files, move_file_to_gdrive


# ── Logging setup ────────────────────────────────────────────

def _setup_logging(cfg) -> None:
    log_dir = Path(cfg.logging.dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    level = getattr(logging, cfg.logging.level.upper(), logging.INFO)
    fmt = "%(asctime)s %(levelname)s [%(name)s] %(message)s"
    handlers = [
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(log_dir / "run_claude.log", encoding="utf-8"),
    ]
    logging.basicConfig(level=level, format=fmt, handlers=handlers, force=True)
    # Suppress noisy libraries
    for noisy in ("httpx", "httpcore", "anthropic._base_client"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


log = logging.getLogger("dd_collector")


# ── Helpers ───────────────────────────────────────────────────

def _snapshot_dir(download_dir: str) -> Set[str]:
    """Return set of filenames currently present in download_dir."""
    p = Path(download_dir)
    if not p.is_dir():
        return set()
    return {f.name for f in p.iterdir() if f.is_file()}


# ── One collection cycle ──────────────────────────────────────

def run_once(cfg, agent, dedup: DedupTracker, only_group: Optional[str] = None) -> None:
    """Run the Claude agent for each configured group and move new downloads."""
    for group in cfg.groups:
        if only_group and group.name != only_group:
            continue

        log.info("=" * 60)
        log.info("Processing group: %s", group.name)
        log.info("=" * 60)

        # Snapshot the download dir before agent runs
        before = _snapshot_dir(cfg.dingtalk.download_dir)

        # Files already recorded in dedup for this group
        already = dedup.get_downloaded_for_group(group.name)
        log.info("Already downloaded for '%s': %d files", group.name, len(already))

        # Run autonomous Claude agent
        try:
            agent.run_download_task(
                group_name=group.name,
                download_dir=cfg.dingtalk.download_dir,
                already_downloaded=already,
                max_scrolls=cfg.claude.max_scrolls,
            )
        except Exception:
            log.exception("Claude agent failed for group: %s", group.name)
            continue

        # Wait for files to finish writing
        log.info("Waiting %ds for downloads to settle…", cfg.dingtalk.download_wait)
        time.sleep(cfg.dingtalk.download_wait)

        # Detect new files
        new_files = get_new_files(cfg.dingtalk.download_dir, before)
        log.info("New files detected: %d", len(new_files))

        if not new_files:
            log.info("No new files for group '%s'.", group.name)
            continue

        # Move to GDrive and record in dedup
        group_alias = group.alias or group.name
        for fpath in new_files:
            if dedup.is_downloaded(group.name, fpath.name):
                log.info("  skip (already in dedup): %s", fpath.name)
                continue
            try:
                dest = move_file_to_gdrive(
                    file_path=fpath,
                    group_folder_name=group_alias,
                    gdrive_base_path=cfg.gdrive.base_path,
                )
                dedup.mark_downloaded(group.name, fpath.name, dest)
                log.info("  moved: %s → %s", fpath.name, dest)
            except Exception:
                log.exception("  failed to move: %s", fpath.name)


# ── Entry point ───────────────────────────────────────────────

def main() -> None:
    cfg = load_config()
    _setup_logging(cfg)

    log.info("DingTalk Claude Agent — starting")

    if not cfg.claude.oauth_token:
        log.error(
            "claude.oauth_token is missing from config.yaml. "
            "Set it or export ANTHROPIC_AUTH_TOKEN."
        )
        sys.exit(1)

    from dd_collector.claude_agent import ClaudeAgent

    agent = ClaudeAgent(
        oauth_token=cfg.claude.oauth_token,
        model=cfg.claude.model,
    )
    log.info("Claude model: %s", cfg.claude.model)

    dedup = DedupTracker(cfg.dedup.path)

    # Parse CLI args
    only_group: Optional[str] = None
    loop_mode = False
    args = sys.argv[1:]
    i = 0
    while i < len(args):
        if args[i] == "--loop":
            loop_mode = True
        elif args[i] == "--group" and i + 1 < len(args):
            only_group = args[i + 1]
            i += 1
        i += 1

    if loop_mode:
        interval = cfg.polling.interval_minutes * 60
        log.info(
            "Loop mode: polling every %d minutes. Ctrl+C to stop.",
            cfg.polling.interval_minutes,
        )
        while True:
            try:
                run_once(cfg, agent, dedup, only_group=only_group)
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
    else:
        run_once(cfg, agent, dedup, only_group=only_group)
        log.info("Done.")


if __name__ == "__main__":
    main()
