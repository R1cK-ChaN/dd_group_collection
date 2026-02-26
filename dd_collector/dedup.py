"""JSON-based dedup tracker keyed by group::filename.

Also stores a per-group timestamp *watermark* so the incremental polling loop
can skip files older than the last successful download.
"""

from __future__ import annotations

import json
import logging
import time
from pathlib import Path
from typing import Dict, List, Optional

log = logging.getLogger("dd_collector")


class DedupTracker:
    """Track which files have been downloaded per group.

    Storage format (data/downloaded.json):
    {
        "GroupA::report.pdf": {
            "timestamp": 1700000000.0,
            "dest": "G:/My Drive/DingTalk Files/GroupA/2024-01/report.pdf"
        },
        ...
    }
    """

    def __init__(self, path: str = "data/downloaded.json"):
        self._path = Path(path)
        self._data: Dict[str, dict] = {}
        self._load()

    # ── Public API ───────────────────────────────────────────

    def is_downloaded(self, group_name: str, file_name: str) -> bool:
        key = self._key(group_name, file_name)
        return key in self._data

    def mark_downloaded(
        self, group_name: str, file_name: str, dest_path: str
    ) -> None:
        key = self._key(group_name, file_name)
        self._data[key] = {
            "timestamp": time.time(),
            "dest": dest_path,
        }
        self._save()

    def get_downloaded_for_group(self, group_name: str) -> List[str]:
        """Return list of filenames already downloaded for a group.

        Normalises composite keys written by the old main.py path
        (``group::timestamp::filename``) to plain filenames so the
        dedup check works regardless of which system created the entry.
        """
        prefix = f"{group_name}::"
        results = []
        for k in self._data:
            if not k.startswith(prefix) or k.startswith("__watermark__::"):
                continue
            suffix = k[len(prefix):]
            # Composite key "timestamp::filename" → extract just the filename
            if "::" in suffix:
                suffix = suffix.rsplit("::", 1)[-1]
            results.append(suffix)
        return results

    # ── Watermark (incremental high-water mark) ──────────────

    _WATERMARK_PREFIX = "__watermark__::"

    def get_watermark(self, group_name: str) -> Optional[str]:
        """Return the last-downloaded timestamp string for *group_name*.

        Format matches DingTalk file row timestamps: ``'2025/07/02 13:52'``.
        Returns ``None`` if no watermark has been set yet.
        """
        key = f"{self._WATERMARK_PREFIX}{group_name}"
        entry = self._data.get(key)
        if isinstance(entry, dict):
            return entry.get("timestamp_str")
        return None

    def set_watermark(self, group_name: str, timestamp_str: str) -> None:
        """Persist the high-water mark for *group_name*."""
        key = f"{self._WATERMARK_PREFIX}{group_name}"
        self._data[key] = {
            "timestamp_str": timestamp_str,
            "updated": time.time(),
        }
        self._save()
        log.info(
            "Watermark updated for '%s': %s", group_name, timestamp_str,
        )

    # ── Chat-based dedup (composite key: group::timestamp::filename) ──

    @staticmethod
    def _chat_key(
        group_name: str, file_name: str, msg_timestamp: str,
    ) -> str:
        """Build a composite dedup key for chat-based downloads.

        Using ``group::timestamp::filename`` instead of ``group::filename``
        because the same filename can be re-shared in chat.
        """
        return f"{group_name}::{msg_timestamp}::{file_name}"

    def is_downloaded_chat(
        self, group_name: str, file_name: str, msg_timestamp: str,
    ) -> bool:
        """Check if a chat attachment has been downloaded (composite key).

        Also checks the legacy ``group::filename`` key for backward
        compatibility with files downloaded before the chat-based switch.
        """
        key = self._chat_key(group_name, file_name, msg_timestamp)
        if key in self._data:
            return True
        # Backward compat: check legacy key
        legacy_key = self._key(group_name, file_name)
        return legacy_key in self._data

    def mark_downloaded_chat(
        self,
        group_name: str,
        file_name: str,
        msg_timestamp: str,
        dest_path: str,
    ) -> None:
        """Mark a chat attachment as downloaded using the composite key."""
        key = self._chat_key(group_name, file_name, msg_timestamp)
        self._data[key] = {
            "timestamp": time.time(),
            "msg_timestamp": msg_timestamp,
            "dest": dest_path,
        }
        self._save()

    # ── Internal ─────────────────────────────────────────────

    @staticmethod
    def _key(group_name: str, file_name: str) -> str:
        return f"{group_name}::{file_name}"

    def _load(self) -> None:
        if not self._path.exists():
            self._data = {}
            return
        try:
            with open(self._path, "r", encoding="utf-8") as f:
                self._data = json.load(f)
            if not isinstance(self._data, dict):
                raise ValueError("Root is not a dict")
        except Exception as exc:
            log.warning(
                "Dedup file corrupt or unreadable (%s), starting fresh: %s",
                self._path, exc,
            )
            self._data = {}

    def _save(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self._path.with_suffix(".tmp")
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._data, f, ensure_ascii=False, indent=2)
            tmp.replace(self._path)
        except Exception as exc:
            log.error("Failed to save dedup file: %s", exc)
            if tmp.exists():
                tmp.unlink()
