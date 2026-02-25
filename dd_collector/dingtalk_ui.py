"""All DingTalk UI interaction in one file.

When DingTalk updates its UI, this is the only Python file that needs editing.
UI selectors are loaded from config.yaml so many changes only require a config edit.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import List, Optional

import pyautogui
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

    def _find_file_grid(self) -> Optional[auto.Control]:
        """Locate the 'grid' GroupControl inside the CefBrowserWindow file view."""
        try:
            doc = find_control(
                self._window,
                "DocumentControl",
                timeout=self.dt.timeout,
                ClassName="Chrome_RenderWidgetHostHWND",
            )
            if not doc:
                log.debug("CefBrowser DocumentControl not found.")
                return None

            # Find the GroupControl named 'grid' which holds file rows
            grid = find_control(doc, "GroupControl", timeout=self.dt.timeout, Name="grid")
            return grid
        except Exception as exc:
            log.debug("_find_file_grid error: %s", exc)
            return None

    @staticmethod
    def _parse_filename(raw_name: str) -> str:
        """Extract the filename from a CustomControl Name like:
        '\\xa0 report.pdf 125.1 KB\\xa0\\xa0·2026/02/21 10:50Author'
        """
        import re
        cleaned = raw_name.replace("\xa0", " ").strip()
        # Match: filename.ext  size_number size_unit
        m = re.match(r"(.+?\.\w+)\s+[\d.]+ [KMGT]?B", cleaned)
        if m:
            return m.group(1).strip()
        return cleaned

    def list_files(self) -> List[FileInfo]:
        """Enumerate visible files in the files tab.

        The file list lives inside a CefBrowserWindow as a TableControl
        with CustomControl children inside a GroupControl named 'grid'.

        Returns a list of FileInfo with name and control reference.
        """
        if not self._window:
            return []

        grid = self._find_file_grid()
        if not grid:
            log.warning("File grid not found. Falling back to TableControl search.")
            # Direct fallback: find TableControl in the window
            table = find_control(self._window, "TableControl", timeout=self.dt.timeout)
            if table:
                grid = find_control(table, "GroupControl", timeout=3, Name="grid")
            if not grid:
                log.warning("File list control not found.")
                return []

        # Scroll the grid to load all files
        scroll_to_bottom(grid, max_scrolls=30)
        time.sleep(0.5)

        # Enumerate CustomControl children (each is a file row)
        files: List[FileInfo] = []
        seen_names: set = set()
        for child in grid.GetChildren():
            raw = child.Name.strip() if child.Name else ""
            if not raw:
                continue
            fname = self._parse_filename(raw)
            if not fname or fname in seen_names:
                continue
            # Skip items that look like folder entries (no file extension match)
            if "Last update:" in raw and "." not in fname.split()[-1]:
                continue
            seen_names.add(fname)
            files.append(FileInfo(name=fname, control=child))

        log.info("Found %d files in list.", len(files))
        return files

    # ── Download ─────────────────────────────────────────────

    def download_file(self, file_info: FileInfo) -> bool:
        """Download a file. Tries pyautogui hover first, then legacy strategies.

        Returns True if the download was triggered.
        """
        # Strategy 1 (primary): pyautogui real mouse hover + click
        if self._download_via_hover_pyautogui(file_info):
            return True

        # Strategy 2 (fallback): Right-click → context menu "下载"
        if self._download_via_context_menu(file_info):
            return True

        # Strategy 3 (fallback): uiautomation hover → click download button
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

    # ── pyautogui Download ────────────────────────────────────

    def _download_via_hover_pyautogui(self, file_info: FileInfo) -> bool:
        """Hover over the file row with a real mouse move (pyautogui) to trigger
        the web-rendered download icon, then click it.

        The CefBrowserWindow hover icons are invisible to Windows UI Automation
        but respond to real mouse events generated by pyautogui.
        """
        try:
            rect = file_info.control.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                log.debug("File row has no valid BoundingRectangle.")
                return False

            row_cx = (rect.left + rect.right) // 2
            row_cy = (rect.top + rect.bottom) // 2

            # Step 1: Move real mouse to row center to trigger hover state
            pyautogui.moveTo(row_cx, row_cy, duration=0.2)
            time.sleep(0.8)

            # Step 2: Check if hover revealed any accessible download button
            for child in file_info.control.GetChildren():
                name = (child.Name or "").lower()
                if name in ("download", "下载"):
                    log.debug("Found accessible download button after hover.")
                    safe_click(child)
                    self._handle_save_dialog()
                    time.sleep(self.dt.download_wait)
                    log.info("Download triggered (pyautogui hover, accessible button): %s", file_info.name)
                    return True

            # Step 3: Click at estimated download icon position (right side of row)
            offset = self.dt.download_icon_offset
            click_x = int(rect.right) - offset
            click_y = row_cy
            log.debug(
                "Clicking estimated download icon at (%d, %d) for %s",
                click_x, click_y, file_info.name,
            )
            pyautogui.click(click_x, click_y)
            self._handle_save_dialog()
            time.sleep(self.dt.download_wait)
            log.info("Download triggered (pyautogui hover): %s", file_info.name)
            return True

        except Exception as exc:
            log.warning("_download_via_hover_pyautogui failed: %s", exc)
            return False

    def _handle_save_dialog(self) -> None:
        """Check for and dismiss any save/download confirmation dialog."""
        time.sleep(0.5)
        for btn_name in ("Save", "保存", "OK", "确定"):
            btn = find_control(
                auto.GetRootControl(),
                "ButtonControl",
                timeout=1,
                Name=btn_name,
            )
            if btn:
                log.debug("Dismissing save dialog with button: %s", btn_name)
                safe_click(btn)
                return

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
