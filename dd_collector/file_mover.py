"""Move downloaded files from DingTalk download dir to Google Drive folders."""

from __future__ import annotations

import logging
import os
import shutil
import time
from datetime import datetime
from pathlib import Path
from typing import List, Set

log = logging.getLogger("dd_collector")

# Files with these suffixes are considered incomplete downloads
_INCOMPLETE_SUFFIXES = {".tmp", ".partial", ".crdownload", ".downloading"}

# Seconds since last modification before a file is considered "settled"
_SETTLE_SECONDS = 3


def get_new_files(
    download_dir: str,
    known_files: Set[str],
) -> List[Path]:
    """Scan download_dir for new, fully-written files.

    Skips:
    - Files whose names are in known_files
    - Files with incomplete-download suffixes
    - Files modified less than _SETTLE_SECONDS ago (still being written)

    Returns:
        Sorted list of Path objects for new files.
    """
    result: List[Path] = []
    dl_path = Path(download_dir)
    if not dl_path.is_dir():
        log.warning("Download dir does not exist: %s", download_dir)
        return result

    now = time.time()
    for entry in dl_path.iterdir():
        if not entry.is_file():
            continue
        if entry.name in known_files:
            continue
        if entry.suffix.lower() in _INCOMPLETE_SUFFIXES:
            continue
        try:
            mtime = entry.stat().st_mtime
        except OSError:
            continue
        if now - mtime < _SETTLE_SECONDS:
            continue
        result.append(entry)

    result.sort(key=lambda p: p.stat().st_mtime)
    return result


def move_file_to_gdrive(
    file_path: Path,
    group_folder_name: str,
    gdrive_base_path: str,
) -> str:
    """Move a file into the Google Drive folder structure.

    Target: {gdrive_base_path}/{group_folder_name}/{YYYY-MM}/{filename}
    On collision: appends _1, _2, ... before the extension.

    Returns:
        The final destination path as a string.
    """
    now = datetime.now()
    month_folder = now.strftime("%Y-%m")
    dest_dir = Path(gdrive_base_path) / group_folder_name / month_folder
    dest_dir.mkdir(parents=True, exist_ok=True)

    dest = dest_dir / file_path.name
    dest = _resolve_collision(dest)

    log.info("Moving %s → %s", file_path.name, dest)
    shutil.move(str(file_path), str(dest))
    return str(dest)


def _resolve_collision(dest: Path) -> Path:
    """If dest exists, append _1, _2, ... before the extension."""
    if not dest.exists():
        return dest

    stem = dest.stem
    suffix = dest.suffix
    parent = dest.parent
    counter = 1
    while True:
        candidate = parent / f"{stem}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1
