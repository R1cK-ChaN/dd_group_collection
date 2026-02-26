"""Main orchestration loop — polling, downloading, moving, dedup."""

from __future__ import annotations

import logging
import os
import sys
import time
from pathlib import Path

from .config import AppConfig, GroupConfig, load_config
from .dedup import DedupTracker
from .dingtalk_ui import ChatAttachment, DingTalkController
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
    """Navigate to a group, scan chat for attachments, download, move to GDrive.

    Chat-based two-path processing (replaces the old Files-tab approach):
    - **File cards** (PDF, XLSX, etc.): click the native Download button.
    - **Images**: right-click → context menu → Download.

    Uses a two-layer filter:
    1. **Watermark** — timestamp high-water mark; attachments at or before
       this are skipped.
    2. **Composite dedup** — ``group::timestamp::filename`` prevents
       re-downloading even if the same filename is re-shared in chat.

    New attachments are downloaded **oldest-first** so the watermark advances
    linearly.  If a download fails mid-batch the watermark is only advanced
    up to the last success.
    """
    log.info("Processing group: %s (alias: %s)", group.name, group.alias)

    controller.dismiss_dialogs()

    if not controller.navigate_to_group(group.name):
        log.warning("Skipping group (navigation failed): %s", group.name)
        return

    # Scan the chat for file cards and images (no Files tab needed)
    attachments = controller.scan_chat_attachments(
        max_scrolls=cfg.polling.chat_scroll_pages,
    )
    if not attachments:
        log.info("No attachments found in group chat: %s", group.name)
        return

    # ── Watermark + dedup filtering ───────────────────────────
    watermark = dedup.get_watermark(group.name)
    log.info(
        "Group '%s': watermark=%s, scanned attachments=%d",
        group.name, watermark, len(attachments),
    )

    new_attachments: list[ChatAttachment] = []
    for att in attachments:
        # Skip if at or before watermark
        if watermark and att.timestamp and att.timestamp <= watermark:
            continue
        # Skip if already downloaded (composite key)
        ts = att.timestamp or ""
        if dedup.is_downloaded_chat(group.name, att.filename, ts):
            continue
        new_attachments.append(att)

    log.info(
        "Group '%s': %d new attachments after watermark+dedup filter.",
        group.name, len(new_attachments),
    )

    if not new_attachments:
        return

    # Apply safety cap
    cap = cfg.polling.max_downloads_per_group
    if len(new_attachments) > cap:
        log.warning(
            "Capping downloads for '%s' from %d to %d.",
            group.name, len(new_attachments), cap,
        )
        new_attachments = new_attachments[:cap]

    # Download oldest-first so watermark advances linearly
    new_attachments.reverse()

    # Snapshot of existing files in the download dir
    dl_dir = cfg.dingtalk.download_dir
    existing_before = _snapshot_download_dir(dl_dir)

    downloaded_count = 0
    newest_downloaded_ts: str | None = None

    for att in new_attachments:
        try:
            success = _download_and_move_chat(
                cfg, controller, dedup, group, att,
                dl_dir, existing_before,
            )
            if success:
                downloaded_count += 1
                if att.timestamp:
                    if (
                        newest_downloaded_ts is None
                        or att.timestamp > newest_downloaded_ts
                    ):
                        newest_downloaded_ts = att.timestamp
                # Update snapshot after each successful download
                existing_before = _snapshot_download_dir(dl_dir)
        except Exception as exc:
            log.error(
                "Error downloading '%s' from '%s': %s",
                att.filename, group.name, exc, exc_info=True,
            )

    # Advance watermark to the newest successfully-downloaded timestamp
    if newest_downloaded_ts:
        dedup.set_watermark(group.name, newest_downloaded_ts)

    log.info(
        "Group '%s': downloaded %d / %d attachments.",
        group.name, downloaded_count, len(new_attachments),
    )


def _download_and_move_chat(
    cfg: AppConfig,
    controller: DingTalkController,
    dedup: DedupTracker,
    group: GroupConfig,
    att: ChatAttachment,
    dl_dir: str,
    existing_before: set,
) -> bool:
    """Download one chat attachment and move it to GDrive.

    Returns True if the file was downloaded and moved successfully.
    """
    log.info("Downloading (%s): %s", att.msg_type, att.filename)

    if not controller.download_chat_attachment(att):
        return False

    # Detect newly appeared files in the download directory
    new_files = get_new_files(dl_dir, existing_before)
    if not new_files:
        log.warning(
            "Download triggered but no new file appeared for: %s",
            att.filename,
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

    # Mark as downloaded with composite key
    ts = att.timestamp or ""
    dedup.mark_downloaded_chat(group.name, att.filename, ts, dest)
    log.info("Completed: %s → %s", att.filename, dest)
    return True


def _snapshot_download_dir(dl_dir: str) -> set:
    """Return a set of filenames currently in the download directory."""
    dl_path = Path(dl_dir)
    if not dl_path.is_dir():
        return set()
    return {f.name for f in dl_path.iterdir() if f.is_file()}
