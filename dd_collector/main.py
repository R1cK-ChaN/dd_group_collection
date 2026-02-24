"""Main orchestration loop — polling, downloading, moving, dedup."""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path

from .config import AppConfig, GroupConfig, load_config
from .dedup import DedupTracker
from .dingtalk_ui import DingTalkController, FileInfo
from .file_mover import get_new_files, move_file_to_gdrive
from .logger import setup_logging
from .ui_helpers import send_escape

log: logging.Logger


def run() -> int:
    """Entry point called from run.py. Returns exit code."""
    global log

    try:
        cfg = load_config()
    except Exception as exc:
        print(f"Failed to load config: {exc}", file=sys.stderr)
        return 1

    log = setup_logging(
        log_dir=cfg.logging.dir,
        level=cfg.logging.level,
        max_bytes=cfg.logging.max_bytes,
        backup_count=cfg.logging.backup_count,
    )

    log.info("=" * 60)
    log.info("DingTalk Group File Collector starting")
    log.info("Groups: %s", [g.name for g in cfg.groups])
    log.info("Polling interval: %d minutes", cfg.polling.interval_minutes)
    log.info("=" * 60)

    dedup = DedupTracker(cfg.dedup.path)
    controller = DingTalkController(cfg)

    try:
        _polling_loop(cfg, controller, dedup)
    except KeyboardInterrupt:
        log.info("Interrupted by user. Exiting.")
    except Exception as exc:
        log.critical("Unhandled error in polling loop: %s", exc, exc_info=True)
        return 2

    return 0


def _polling_loop(
    cfg: AppConfig,
    controller: DingTalkController,
    dedup: DedupTracker,
) -> None:
    """Run collection cycles forever."""
    cycle = 0
    while True:
        cycle += 1
        log.info("── Cycle %d starting ──", cycle)
        try:
            _run_cycle(cfg, controller, dedup)
        except Exception as exc:
            log.error("Cycle %d failed: %s", cycle, exc, exc_info=True)

        log.info(
            "── Cycle %d done. Sleeping %d minutes ──",
            cycle, cfg.polling.interval_minutes,
        )
        time.sleep(cfg.polling.interval_minutes * 60)


def _run_cycle(
    cfg: AppConfig,
    controller: DingTalkController,
    dedup: DedupTracker,
) -> None:
    """Process all groups once."""
    if not controller.connect():
        log.error("Cannot connect to DingTalk. Skipping cycle.")
        return

    for group in cfg.groups:
        try:
            _process_group(cfg, controller, dedup, group)
        except Exception as exc:
            log.error(
                "Error processing group '%s': %s", group.name, exc, exc_info=True,
            )
            # Try to recover for the next group
            send_escape()
            time.sleep(1)


def _process_group(
    cfg: AppConfig,
    controller: DingTalkController,
    dedup: DedupTracker,
    group: GroupConfig,
) -> None:
    """Navigate to a group, list files, download new ones, move to GDrive."""
    log.info("Processing group: %s (alias: %s)", group.name, group.alias)

    controller.dismiss_dialogs()

    if not controller.navigate_to_group(group.name):
        log.warning("Skipping group (navigation failed): %s", group.name)
        return

    if not controller.open_files_tab():
        log.warning("Skipping group (files tab failed): %s", group.name)
        return

    files = controller.list_files()
    if not files:
        log.info("No files found in group: %s", group.name)
        return

    # Filter out already-downloaded files
    new_files = [
        f for f in files
        if not dedup.is_downloaded(group.name, f.name)
    ]
    log.info(
        "Group '%s': %d total files, %d new.",
        group.name, len(files), len(new_files),
    )

    if not new_files:
        return

    # Apply safety cap
    cap = cfg.polling.max_downloads_per_group
    if len(new_files) > cap:
        log.warning(
            "Capping downloads for '%s' from %d to %d.",
            group.name, len(new_files), cap,
        )
        new_files = new_files[:cap]

    # Take a snapshot of existing files in the download dir before downloading
    dl_dir = cfg.dingtalk.download_dir
    existing_before = _snapshot_download_dir(dl_dir)

    downloaded_count = 0
    for file_info in new_files:
        try:
            success = _download_and_move(
                cfg, controller, dedup, group, file_info,
                dl_dir, existing_before,
            )
            if success:
                downloaded_count += 1
                # Update snapshot after each successful download
                existing_before = _snapshot_download_dir(dl_dir)
        except Exception as exc:
            log.error(
                "Error downloading '%s' from '%s': %s",
                file_info.name, group.name, exc, exc_info=True,
            )

    log.info(
        "Group '%s': downloaded %d / %d files.",
        group.name, downloaded_count, len(new_files),
    )


def _download_and_move(
    cfg: AppConfig,
    controller: DingTalkController,
    dedup: DedupTracker,
    group: GroupConfig,
    file_info: FileInfo,
    dl_dir: str,
    existing_before: set,
) -> bool:
    """Download one file and move it to GDrive.

    Returns True if the file was downloaded and moved successfully.
    """
    log.info("Downloading: %s", file_info.name)

    if not controller.download_file(file_info):
        return False

    # Detect newly appeared files in the download directory
    new_files = get_new_files(dl_dir, existing_before)
    if not new_files:
        log.warning(
            "Download triggered but no new file appeared for: %s", file_info.name,
        )
        return False

    # Take the most recently modified file as the downloaded one
    downloaded_path = new_files[-1]

    # Move to GDrive
    dest = move_file_to_gdrive(
        downloaded_path,
        group.alias,
        cfg.gdrive.base_path,
    )

    # Mark as downloaded
    dedup.mark_downloaded(group.name, file_info.name, dest)
    log.info("Completed: %s → %s", file_info.name, dest)
    return True


def _snapshot_download_dir(dl_dir: str) -> set:
    """Return a set of filenames currently in the download directory."""
    dl_path = Path(dl_dir)
    if not dl_path.is_dir():
        return set()
    return {f.name for f in dl_path.iterdir() if f.is_file()}
