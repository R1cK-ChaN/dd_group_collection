"""JSON-based dedup tracker keyed by group::filename."""

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
        """Return list of filenames already downloaded for a group."""
        prefix = f"{group_name}::"
        return [
            k[len(prefix):] for k in self._data if k.startswith(prefix)
        ]

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
