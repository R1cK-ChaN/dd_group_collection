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
    scroll_to_top,
    send_escape,
    set_text,
)

log = logging.getLogger("dd_collector")


@dataclass
class FileInfo:
    """Represents one file entry in a DingTalk group's file list."""
    name: str
    control: auto.Control  # reference to the ListItem for clicking
    timestamp: Optional[str] = None  # "YYYY/MM/DD HH:MM" from the raw Name


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
        """Locate the 'grid' GroupControl inside the file view.

        The grid can appear in two layouts:
        - Current: GroupControl Name='grid' directly accessible from the window
        - Legacy: inside DocumentControl(Chrome_RenderWidgetHostHWND) → GroupControl Name='grid'
        """
        try:
            # Try direct search first (current DtMainFrameView layout)
            grid = find_control(
                self._window, "GroupControl",
                timeout=self.dt.timeout, Name="grid",
            )
            if grid:
                return grid

            # Legacy: inside DocumentControl
            doc = find_control(
                self._window, "DocumentControl",
                timeout=3, ClassName="Chrome_RenderWidgetHostHWND",
            )
            if doc:
                grid = find_control(
                    doc, "GroupControl", timeout=3, Name="grid",
                )
                if grid:
                    return grid

            log.debug("File grid not found in any known layout.")
            return None
        except Exception as exc:
            log.debug("_find_file_grid error: %s", exc)
            return None

    @staticmethod
    def _parse_filename(raw_name: str) -> str:
        """Extract the filename from a CustomControl Name like:
        '  250702GMF.PDF 1.1 MB  ·2025/07/02 13:52沧海一土狗'
        """
        import re
        cleaned = raw_name.replace("\xa0", " ").strip()
        # Match: filename.ext  size_number size_unit
        m = re.match(r"(.+?\.\w+)\s+[\d.]+ [KMGT]?B", cleaned)
        if m:
            return m.group(1).strip()
        return cleaned

    @staticmethod
    def _parse_timestamp(raw_name: str) -> Optional[str]:
        """Extract the upload timestamp from a file row Name.

        DingTalk uses relative date labels for recent files:
        - Today:     '·Today 18:51AuthorName'
        - Yesterday: '·Yesterday 14:30AuthorName'
        - Older:     '·2025/07/02 13:52AuthorName'

        All are normalised to 'YYYY/MM/DD HH:MM' for consistent comparison.
        Returns None if no timestamp can be parsed.
        """
        import re
        from datetime import datetime, timedelta

        # Pattern 1: absolute date  ·YYYY/MM/DD HH:MM
        m = re.search(r"(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2})", raw_name)
        if m:
            return m.group(1)

        # Pattern 2: ·Today HH:MM
        m = re.search(r"·Today\s+(\d{2}:\d{2})", raw_name)
        if m:
            today = datetime.now().strftime("%Y/%m/%d")
            return f"{today} {m.group(1)}"

        # Pattern 3: ·Yesterday HH:MM
        m = re.search(r"·Yesterday\s+(\d{2}:\d{2})", raw_name)
        if m:
            yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y/%m/%d")
            return f"{yesterday} {m.group(1)}"

        return None

    def list_files(self, max_scrolls: int = 30) -> List[FileInfo]:
        """Enumerate visible files in the files tab.

        The file grid has two possible structures:
        - Current: grid → GroupControl(container) → CustomControl(file rows)
        - Legacy:  grid → CustomControl(file rows) directly

        Args:
            max_scrolls: How many scroll-down iterations to perform to load
                files.  Use a small value (e.g. 3) for incremental checks
                where only the newest files at the top of the list matter.
                Use a large value (e.g. 30) for a full scan.

        Returns a list of FileInfo with name and control reference.
        """
        if not self._window:
            return []

        grid = self._find_file_grid()
        if not grid:
            log.warning("File list control not found.")
            return []

        # Scroll the grid to load files
        scroll_to_bottom(grid, max_scrolls=max_scrolls)
        time.sleep(0.5)

        # Collect all candidate controls — file rows may be direct children
        # of the grid or nested inside a container GroupControl.
        candidates: list = []
        for child in grid.GetChildren():
            if child.Name and child.Name.strip():
                candidates.append(child)
            else:
                # Unnamed container — check its children for file rows
                for sub in child.GetChildren():
                    if sub.Name and sub.Name.strip():
                        candidates.append(sub)

        # Parse filenames and timestamps from candidate controls
        files: List[FileInfo] = []
        seen_names: set = set()
        for ctrl in candidates:
            raw = ctrl.Name.strip()
            fname = self._parse_filename(raw)
            if not fname or fname in seen_names:
                continue
            # Skip items that look like folder entries (no file extension match)
            if "Last update:" in raw and "." not in fname.split()[-1]:
                continue
            seen_names.add(fname)
            ts = self._parse_timestamp(raw)
            files.append(FileInfo(name=fname, control=ctrl, timestamp=ts))

        log.info("Found %d files in list.", len(files))

        # Scroll back to top so newest files are visible for interaction
        scroll_to_top(grid, max_scrolls=max_scrolls + 5)
        time.sleep(0.3)

        return files

    # ── Download ─────────────────────────────────────────────

    def _scroll_file_into_view(self, file_info: FileInfo) -> bool:
        """Ensure the file row is visible on screen before interacting with it.

        Returns True if the BoundingRectangle is valid (non-zero).
        """
        rect = file_info.control.BoundingRectangle
        if rect.width() > 0 and rect.height() > 0:
            return True

        # Try ScrollItemPattern to bring the control into view
        try:
            sip = file_info.control.GetScrollItemPattern()
            if sip:
                sip.ScrollIntoView()
                time.sleep(0.5)
                rect = file_info.control.BoundingRectangle
                if rect.width() > 0 and rect.height() > 0:
                    return True
        except Exception:
            pass

        log.warning(
            "Cannot scroll file into view: %s (rect still zero)", file_info.name,
        )
        return False

    def download_file(self, file_info: FileInfo) -> bool:
        """Download a file.  Tries VLM-guided right-click first, then fallbacks.

        Returns True if the download was triggered.
        """
        # Ensure the file row is visible on screen
        if not self._scroll_file_into_view(file_info):
            log.error("File not visible, cannot download: %s", file_info.name)
            return False
        # Strategy 1 (primary): right-click → VLM finds "Download" in context menu
        if self._download_via_vlm(file_info):
            return True

        # Strategy 2 (fallback): uiautomation right-click → MenuItemControl
        if self._download_via_context_menu(file_info):
            return True

        log.error("Failed to download file: %s", file_info.name)
        return False

    # ── VLM-guided Download ───────────────────────────────────

    def _download_via_vlm(self, file_info: FileInfo) -> bool:
        """Right-click → screenshot context menu → VLM locates 'Download' → click.

        1. Right-click the file row center with pyautogui (real mouse event).
        2. Capture a screenshot of the region around the click.
        3. Send to Qwen-VL-Plus via OpenRouter to get the (x, y) of "Download".
        4. Click at those coordinates.
        """
        from .vlm import find_menu_item_coords, grab_screenshot_base64

        vlm_cfg = self.cfg.vlm
        if not vlm_cfg.api_key:
            log.debug("VLM API key not configured, skipping VLM strategy.")
            return False

        try:
            rect = file_info.control.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                log.debug("File row has no valid BoundingRectangle.")
                return False

            row_cx = (rect.left + rect.right) // 2
            row_cy = (rect.top + rect.bottom) // 2

            # Ensure DingTalk is focused
            if self._window:
                try:
                    self._window.SetActive()
                    self._window.SetFocus()
                    time.sleep(0.3)
                except Exception:
                    pass

            # Step 1: Right-click with real mouse
            pyautogui.rightClick(row_cx, row_cy)
            time.sleep(1.0)  # wait for context menu to render

            # Step 2: Capture region around the right-click point
            margin = vlm_cfg.capture_margin
            cap_left = max(0, row_cx - margin)
            cap_top = max(0, row_cy - margin)
            cap_right = row_cx + margin
            cap_bottom = row_cy + margin
            region = (cap_left, cap_top, cap_right, cap_bottom)

            screenshot_b64 = grab_screenshot_base64(region)
            log.debug(
                "Captured context menu screenshot (%dx%d) at (%d,%d).",
                cap_right - cap_left, cap_bottom - cap_top, cap_left, cap_top,
            )

            # Step 3: Ask VLM to find "Download"
            coords = find_menu_item_coords(
                api_key=vlm_cfg.api_key,
                screenshot_b64=screenshot_b64,
                target_label="Download",
                region_offset=(cap_left, cap_top),
                model=vlm_cfg.model,
                base_url=vlm_cfg.base_url,
            )

            if not coords:
                log.warning("VLM could not locate Download in context menu.")
                pyautogui.press("escape")
                time.sleep(0.3)
                return False

            # Step 4: Click at the VLM-identified coordinates
            click_x, click_y = coords
            log.info(
                "VLM: clicking Download at (%d, %d) for %s",
                click_x, click_y, file_info.name,
            )
            pyautogui.click(click_x, click_y)

            # Handle potential save dialog
            self._handle_save_dialog()
            time.sleep(self.dt.download_wait)
            log.info(
                "Download triggered (VLM context menu): %s", file_info.name,
            )
            return True

        except Exception as exc:
            log.warning("_download_via_vlm failed: %s", exc)
            # Dismiss any lingering context menu
            try:
                pyautogui.press("escape")
            except Exception:
                pass
            return False

    # ── Legacy Download Fallbacks ─────────────────────────────

    def _download_via_context_menu(self, file_info: FileInfo) -> bool:
        """Right-click → find MenuItemControl via uiautomation (legacy fallback)."""
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

        log.info("Download triggered (context menu uia): %s", file_info.name)
        time.sleep(self.dt.download_wait)
        return True

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
