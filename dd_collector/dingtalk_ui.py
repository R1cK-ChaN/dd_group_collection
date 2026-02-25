"""All DingTalk UI interaction in one file.

When DingTalk updates its UI, this is the only Python file that needs editing.
UI selectors are loaded from config.yaml so many changes only require a config edit.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import List, Optional

import uiautomation as auto

from .config import AppConfig, SelectorConfig
from .ui_helpers import (
    find_control,
    safe_click,
    safe_right_click,
    scroll_to_bottom,
    send_escape,
    set_text,
)

log = logging.getLogger("dd_collector")


@dataclass
class FileInfo:
    """Represents one file entry in a DingTalk group's file list."""
    name: str
    control: auto.Control  # reference to the ListItem for clicking


class DingTalkController:
    """High-level controller for DingTalk PC client UI."""

    def __init__(self, config: AppConfig):
        self.cfg = config
        self.dt = config.dingtalk
        self.sel = config.ui_selectors
        self._window: Optional[auto.WindowControl] = None

    # ── Connection ───────────────────────────────────────────

    def connect(self) -> bool:
        """Find and activate the DingTalk main window.

        Returns True if the window was found and activated.
        """
        try:
            self._window = auto.WindowControl(
                ClassName=self.dt.window_class,
                searchDepth=1,
            )
            if not self._window.Exists(maxSearchSeconds=self.dt.timeout):
                log.error(
                    "DingTalk window not found (class=%s). Is DingTalk running?",
                    self.dt.window_class,
                )
                return False
            self._window.SetActive()
            self._window.SetFocus()
            time.sleep(0.5)
            log.info("Connected to DingTalk window.")
            return True
        except Exception as exc:
            log.error("Failed to connect to DingTalk: %s", exc)
            return False

    # ── Navigation ───────────────────────────────────────────

    def navigate_to_group(self, group_name: str) -> bool:
        """Search for a group by name and click on the result.

        Returns True if the group was successfully opened.
        """
        if not self._window:
            return False

        self.dismiss_dialogs()

        # Click search box
        sel = self.sel.search_box
        search_box = find_control(
            self._window,
            sel.control_type,
            timeout=self.dt.timeout,
            Name=sel.name,
        )
        if not search_box:
            log.error("Search box not found.")
            return False

        if not safe_click(search_box, delay_after=0.5):
            return False

        # Type group name
        if not set_text(search_box, group_name, delay_after=1.0):
            log.error("Failed to type group name: %s", group_name)
            send_escape()
            return False

        # Wait for results and click first match
        time.sleep(1.5)
        sel_item = self.sel.search_result_item
        result = find_control(
            self._window,
            sel_item.control_type,
            timeout=self.dt.timeout,
        )
        if not result:
            log.error("No search results found for group: %s", group_name)
            send_escape()
            return False

        if not safe_click(result, delay_after=1.0):
            send_escape()
            return False

        # Press Escape to close the search overlay
        send_escape()
        time.sleep(0.5)

        log.info("Navigated to group: %s", group_name)
        return True

    # ── Files Tab ────────────────────────────────────────────

    def open_files_tab(self) -> bool:
        """Click the '文件' (Files) tab in the current group.

        Tries the primary control type first, then the fallback.
        """
        if not self._window:
            return False

        sel = self.sel.files_tab
        tab = find_control(
            self._window,
            sel.control_type,
            timeout=self.dt.timeout,
            Name=sel.name,
        )

        # Try fallback control type
        if not tab and sel.fallback_control_type:
            tab = find_control(
                self._window,
                sel.fallback_control_type,
                timeout=self.dt.timeout / 2,
                Name=sel.name,
            )

        if not tab:
            log.error("Files tab not found.")
            return False

        if not safe_click(tab, delay_after=1.0):
            return False

        log.info("Opened files tab.")
        return True

    # ── File Listing ─────────────────────────────────────────

    def list_files(self) -> List[FileInfo]:
        """Enumerate visible files in the files tab, scrolling to load more.

        Returns a list of FileInfo with name and control reference.
        """
        if not self._window:
            return []

        sel_list = self.sel.file_list
        file_list = find_control(
            self._window,
            sel_list.control_type,
            timeout=self.dt.timeout,
            Name=sel_list.name if sel_list.name else None,
        )
        if not file_list:
            # Try without Name filter
            file_list = find_control(
                self._window,
                sel_list.control_type,
                timeout=self.dt.timeout / 2,
            )
        if not file_list:
            log.warning("File list control not found.")
            return []

        # Scroll to load all files
        scroll_to_bottom(file_list, max_scrolls=30)
        time.sleep(0.5)

        # Enumerate items
        sel_item = self.sel.file_item
        item_class = getattr(auto, sel_item.control_type, auto.ListItemControl)
        children = file_list.GetChildren()

        files: List[FileInfo] = []
        seen_names: set = set()
        for child in children:
            if not isinstance(child, item_class) and child.ControlTypeName != sel_item.control_type.replace("Control", ""):
                continue
            name = child.Name.strip() if child.Name else ""
            if not name or name in seen_names:
                continue
            seen_names.add(name)
            files.append(FileInfo(name=name, control=child))

        log.info("Found %d files in list.", len(files))
        return files

    # ── Download ─────────────────────────────────────────────

    def download_file(self, file_info: FileInfo) -> bool:
        """Download a file. Tries right-click context menu first, then hover+button.

        Returns True if the download was triggered.
        """
        # Strategy 1: Right-click → context menu "下载"
        if self._download_via_context_menu(file_info):
            return True

        # Strategy 2: Hover → click download button
        if self._download_via_hover_button(file_info):
            return True

        log.error("Failed to download file: %s", file_info.name)
        return False

    def _download_via_context_menu(self, file_info: FileInfo) -> bool:
        """Right-click the file item, then click '下载' in the context menu."""
        if not safe_right_click(file_info.control, delay_after=0.5):
            return False

        sel = self.sel.context_menu_download
        menu_item = find_control(
            auto.GetRootControl(),
            sel.control_type,
            timeout=3,
            Name=sel.name,
        )
        if not menu_item:
            send_escape()  # dismiss context menu
            return False

        if not safe_click(menu_item, delay_after=0.5):
            send_escape()
            return False

        log.info("Download triggered (context menu): %s", file_info.name)
        time.sleep(self.dt.download_wait)
        return True

    def _download_via_hover_button(self, file_info: FileInfo) -> bool:
        """Hover over the file item to reveal the download button, then click it."""
        try:
            file_info.control.MoveCursorToMyCenter()
            time.sleep(0.5)
        except Exception as exc:
            log.debug("Hover failed: %s", exc)
            return False

        sel = self.sel.download_button
        dl_btn = find_control(
            file_info.control,
            sel.control_type,
            timeout=3,
            Name=sel.name,
        )
        if not dl_btn:
            # Try searching from window root as the button might be a popup
            dl_btn = find_control(
                self._window,
                sel.control_type,
                timeout=2,
                Name=sel.name,
            )
        if not dl_btn:
            return False

        if not safe_click(dl_btn, delay_after=0.5):
            return False

        log.info("Download triggered (hover button): %s", file_info.name)
        time.sleep(self.dt.download_wait)
        return True

    # ── Dialog Dismissal ─────────────────────────────────────

    def dismiss_dialogs(self, max_rounds: int = 3) -> None:
        """Dismiss update/notification popups by looking for known button texts.

        Repeats up to *max_rounds* times so stacked dialogs are handled,
        but cannot loop forever if a button keeps reappearing.
        """
        if not self._window:
            return

        for _round in range(max_rounds):
            dismissed = False
            for btn_text in self.sel.dismiss_buttons:
                btn = find_control(
                    self._window,
                    "ButtonControl",
                    timeout=1,
                    Name=btn_text,
                )
                if btn:
                    log.info("Dismissing dialog with button: %s", btn_text)
                    safe_click(btn, delay_after=0.3)
                    dismissed = True
            if not dismissed:
                break
