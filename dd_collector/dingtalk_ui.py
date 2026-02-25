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

    # Known DingTalk window classes (tried in order).
    _WINDOW_CLASSES = ("DtMainFrameView", "StandardFrame_DingTalk")

    def connect(self) -> bool:
        """Find and activate the DingTalk main window.

        Tries the configured window_class first, then falls back to known
        alternatives.  DingTalk must be both SetActive *and* SetFocus before
        pyautogui mouse/keyboard events will land on it.

        Returns True if the window was found and activated.
        """
        classes_to_try = [self.dt.window_class] + [
            c for c in self._WINDOW_CLASSES if c != self.dt.window_class
        ]
        for cls_name in classes_to_try:
            try:
                win = auto.WindowControl(ClassName=cls_name, searchDepth=1)
                if win.Exists(maxSearchSeconds=3):
                    self._window = win
                    self._window.SetActive()
                    time.sleep(0.3)
                    self._window.SetFocus()
                    time.sleep(0.3)
                    log.info(
                        "Connected to DingTalk window (class=%s).", cls_name,
                    )
                    return True
            except Exception as exc:
                log.debug("Trying window class %s: %s", cls_name, exc)

        log.error(
            "DingTalk window not found (tried %s). Is DingTalk running?",
            classes_to_try,
        )
        return False

    # ── Navigation ───────────────────────────────────────────

    def _find_search_box(self) -> Optional[auto.Control]:
        """Locate the search box, trying ClassName first then Name.

        The DingTalk search box is a QLineEdit with no Name attribute.
        Falls back to searching by Name='Search' for older UI versions.
        """
        sel = self.sel.search_box
        # Prefer ClassName match (works on current DingTalk builds)
        if sel.class_name:
            box = find_control(
                self._window,
                sel.control_type,
                timeout=self.dt.timeout,
                ClassName=sel.class_name,
            )
            if box:
                return box

        # Fall back to Name match (older builds)
        if sel.name:
            box = find_control(
                self._window,
                sel.control_type,
                timeout=self.dt.timeout / 2,
                Name=sel.name,
            )
            if box:
                return box

        return None

    def navigate_to_group(self, group_name: str) -> bool:
        """Search for a group by name and open the conversation.

        DingTalk's conversation list items have empty Name attributes, so
        we cannot enumerate or identify them via UI Automation.  Instead:

        1. Click the search box (QLineEdit) to focus it.
        2. Type the group name via uiautomation SendKeys — this populates
           the search field and DingTalk filters/searches in real time.
        3. Press Enter via pyautogui to select the first result.  Clicking
           the search box changes the foreground to a DtQtWebView search
           overlay, so pyautogui keyboard events land there correctly.
        4. Wait for the group chat to open — verified by checking that the
           welcome screen is gone or that group header buttons appear.

        Returns True if the group was successfully opened.
        """
        if not self._window:
            return False

        self.dismiss_dialogs()

        # Ensure DingTalk has real focus before pyautogui events
        try:
            self._window.SetActive()
            self._window.SetFocus()
            time.sleep(0.3)
        except Exception:
            pass

        # Click search box
        search_box = self._find_search_box()
        if not search_box:
            log.error("Search box not found.")
            return False

        if not safe_click(search_box, delay_after=0.5):
            return False

        # Clear any previous text and type group name.
        # After clicking the search box the foreground may switch to
        # DtQtWebView (a search overlay), so we use SendKeys which works
        # regardless of the foreground window.
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.press("delete")
        time.sleep(0.2)

        if not set_text(search_box, group_name, delay_after=1.5):
            log.error("Failed to type group name: %s", group_name)
            send_escape()
            return False

        # Select the first search result by pressing Enter.
        # The search overlay is focused so pyautogui's Enter goes there.
        pyautogui.press("enter")
        time.sleep(2)

        # Verify that a group chat actually opened.
        # On success, group header buttons like "Files" appear.
        if self._verify_group_opened():
            log.info("Navigated to group: %s", group_name)
            return True

        # If not opened, try pressing Escape and retry with Ctrl-click
        log.warning("Group chat may not have opened. Retrying...")
        send_escape()
        time.sleep(0.5)
        return False

    def _verify_group_opened(self) -> bool:
        """Check whether a group conversation is currently open.

        Looks for characteristic group header buttons (Files, Group Settings, etc.)
        that only appear inside a group chat.
        """
        for btn_name in ("Files", "Group Settings", "Group Notice", "Chat History"):
            btn = find_control(
                self._window, "ButtonControl", timeout=2, Name=btn_name,
            )
            if btn:
                log.debug("Group header button found: %s", btn_name)
                return True

        # Fallback: check that the welcome screen is gone
        welcome = find_control(
            self._window, "TextControl", timeout=1,
            Name="DingTalk, the way of working in the AI era",
        )
        return welcome is None

    # ── Files Tab ────────────────────────────────────────────

    def open_files_tab(self) -> bool:
        """Click the 'Files' button in the group chat header.

        The button appears alongside Group Notice, Chat History, More, and
        Group Settings in the top-right area of the chat.  Tries the
        configured name first ("Files"), then common alternatives.
        """
        if not self._window:
            return False

        sel = self.sel.files_tab
        names_to_try = [sel.name] if sel.name else []
        # Add common alternatives for English/Chinese UI
        for alt in ("Files", "File", "文件"):
            if alt not in names_to_try:
                names_to_try.append(alt)

        tab = None
        for name in names_to_try:
            tab = find_control(
                self._window,
                sel.control_type,
                timeout=self.dt.timeout / 2,
                Name=name,
            )
            if tab:
                break
            # Try fallback control type
            if sel.fallback_control_type:
                tab = find_control(
                    self._window,
                    sel.fallback_control_type,
                    timeout=2,
                    Name=name,
                )
                if tab:
                    break

        if not tab:
            log.error("Files tab not found (tried names: %s).", names_to_try)
            return False

        # Use pyautogui click to ensure the click lands on the right element,
        # since the Files button is small (28x28) and near other header buttons.
        rect = tab.BoundingRectangle
        cx = (rect.left + rect.right) // 2
        cy = (rect.top + rect.bottom) // 2
        pyautogui.click(cx, cy)
        time.sleep(1.5)

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
