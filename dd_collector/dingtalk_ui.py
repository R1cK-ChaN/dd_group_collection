"""All DingTalk UI interaction in one file.

When DingTalk updates its UI, this is the only Python file that needs editing.
UI selectors are loaded from config.yaml so many changes only require a config edit.
"""

from __future__ import annotations

import base64
import io
import logging
import os
import subprocess
import time
from dataclasses import dataclass
from typing import List, Optional, Tuple

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


@dataclass
class ChatAttachment:
    """A downloadable attachment found in a DingTalk chat message.

    Two types:
    - file_card: PDF/XLSX/etc rendered as a card with a Download button.
      All chat content is web-rendered in CefBrowserWindow (invisible to UIA),
      so we use VLM to find the Download button coordinates.
    - image: Inline image preview; requires right-click → context menu to download.
    """
    msg_type: str  # "file_card" or "image"
    filename: str
    timestamp: Optional[str] = None  # "YYYY/MM/DD HH:MM"
    download_click: Optional[Tuple[int, int]] = None  # screen-absolute (x, y) for file_card
    image_bounds: Optional[Tuple[int, int, int, int]] = None  # for image type


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

    def ensure_running(self) -> bool:
        """Launch DingTalk if it is not already running.

        Checks for the main window; if absent, starts the exe from
        ``dingtalk.exe_path`` in config.yaml and waits up to 30 seconds
        for the window to appear.

        Returns True if DingTalk is running and a window is visible.
        Call this before ``connect()`` so the window is guaranteed to exist.
        """
        # Fast path: window already exists
        for cls_name in self._WINDOW_CLASSES:
            try:
                win = auto.WindowControl(ClassName=cls_name, searchDepth=1)
                if win.Exists(maxSearchSeconds=1):
                    log.info("DingTalk already running (class=%s).", cls_name)
                    return True
            except Exception:
                pass

        # Need to launch — check exe_path is configured
        exe_path = self.dt.exe_path
        if not exe_path:
            log.error(
                "DingTalk is not running and dingtalk.exe_path is not set "
                "in config.yaml. Start DingTalk manually or add the exe path."
            )
            return False

        if not os.path.isfile(exe_path):
            log.error("DingTalk exe not found at configured path: %s", exe_path)
            return False

        log.info("DingTalk not running — launching: %s", exe_path)
        try:
            subprocess.Popen([exe_path])
        except Exception as exc:
            log.error("Failed to launch DingTalk: %s", exc)
            return False

        # Wait up to 30 s for the window to appear
        log.info("Waiting for DingTalk window (up to 30 s)…")
        for elapsed in range(30):
            time.sleep(1.0)
            for cls_name in self._WINDOW_CLASSES:
                try:
                    win = auto.WindowControl(ClassName=cls_name, searchDepth=1)
                    if win.Exists(maxSearchSeconds=1):
                        log.info(
                            "DingTalk window appeared after %ds (class=%s).",
                            elapsed + 1, cls_name,
                        )
                        time.sleep(3.0)  # let the app fully initialise
                        return True
                except Exception:
                    pass

        log.error("DingTalk window did not appear after 30 s.")
        return False

    def wait_for_ready(self, timeout: int = 45) -> bool:
        """Block until DingTalk's main chat UI is accessible.

        Keeps dismissing startup dialogs (update prompts, announcements) while
        waiting for the search box to appear.  This is needed when DingTalk
        was just launched — the window appears quickly but the UI takes several
        more seconds to fully initialise.

        Returns True once the search box is found; False on timeout.
        """
        log.info("Waiting for DingTalk to be ready (up to %ds)…", timeout)
        for elapsed in range(timeout):
            # Keep dismissing any dialogs that block the main UI
            self.dismiss_dialogs(max_rounds=1)
            box = self._find_search_box()
            if box:
                log.info("DingTalk ready (search box found after %ds).", elapsed)
                return True
            time.sleep(1.0)
        log.error("DingTalk not ready after %ds — search box never appeared.", timeout)
        return False

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

        DingTalk's newer UI opens a full-window search overlay when the search
        box is clicked.  The overlay shows search results in web-rendered tabs
        (Contacts, Groups, Chat Records, etc.) that are invisible to UIA.

        Navigation approach:
        1. Click the search box (QLineEdit) → search overlay opens.
        2. Type the group name via SendKeys.
        3. Press Down → Enter to select the first result in the overlay.
           This navigates the *background* group to the target even though
           the search overlay visually remains on top.
        4. Click the "Collapse (esc)" UIA button to dismiss the overlay,
           revealing the now-active group chat.
        5. Verify via group header buttons (Files, Group Settings, etc.).

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

        # Close any lingering search overlay before starting fresh
        self._collapse_search_overlay()

        # Click search box
        search_box = self._find_search_box()
        if not search_box:
            log.error("Search box not found.")
            return False

        if not safe_click(search_box, delay_after=0.5):
            return False

        # Clear any previous text and type group name.
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.press("delete")
        time.sleep(0.2)

        if not set_text(search_box, group_name, delay_after=1.5):
            log.error("Failed to type group name: %s", group_name)
            self._collapse_search_overlay()
            return False

        # Press Down then Enter to navigate to the first search result.
        # Down moves keyboard focus into the result list; Enter selects it.
        # This navigates the background group chat even though the overlay
        # visually stays open.
        pyautogui.press("down")
        time.sleep(0.4)
        pyautogui.press("enter")
        time.sleep(1.5)

        # Dismiss the search overlay.  The "Collapse (esc)" button is a UIA
        # ButtonControl that appears in the top-left of the search overlay.
        collapsed = self._collapse_search_overlay()
        if not collapsed:
            # Fallback: use Escape
            send_escape()
        time.sleep(0.5)

        # Verify that the correct group chat is now visible.
        if self._verify_group_opened():
            log.info("Navigated to group: %s", group_name)
            return True

        # Retry with a fresh search
        log.warning(
            "Group chat may not be open after first attempt; retrying: %s",
            group_name,
        )
        self._collapse_search_overlay()
        time.sleep(0.3)

        search_box2 = self._find_search_box()
        if not search_box2:
            return False
        if not safe_click(search_box2, delay_after=0.5):
            return False

        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.1)
        pyautogui.press("delete")
        time.sleep(0.2)

        if not set_text(search_box2, group_name, delay_after=2.0):
            self._collapse_search_overlay()
            return False

        pyautogui.press("down")
        time.sleep(0.4)
        pyautogui.press("enter")
        time.sleep(2.0)

        self._collapse_search_overlay()
        time.sleep(0.5)

        if self._verify_group_opened():
            log.info("Navigated to group (retry): %s", group_name)
            return True

        log.error("Navigation failed after retry: %s", group_name)
        return False

    def _collapse_search_overlay(self) -> bool:
        """Close the DingTalk search overlay.

        The search overlay shows a keyboard hint "ESC to Dismiss" at the
        bottom.  Pressing Escape once clears the query text; pressing Escape
        a second time (when the query is already empty) dismisses the overlay
        entirely.  We press Escape up to three times to handle both states.

        Returns True if the overlay was open and a close action was performed.
        """
        if not self._window:
            return False

        # Check if the overlay is open at all
        collapse_btn = find_control(
            self._window,
            "ButtonControl",
            timeout=0.5,
            Name="Collapse (esc)",
        )
        if not collapse_btn:
            return False  # Overlay not open, nothing to do

        log.debug(
            "Search overlay is open; pressing Escape twice to dismiss.",
        )
        try:
            self._window.SetActive()
            self._window.SetFocus()
        except Exception:
            pass

        # First Escape: clears the search query
        pyautogui.press("escape")
        time.sleep(0.4)
        # Second Escape: dismisses the overlay when query is already empty
        pyautogui.press("escape")
        time.sleep(0.4)
        # Third Escape: extra safety for stubborn states
        pyautogui.press("escape")
        time.sleep(0.5)

        return True

    def _verify_group_opened(self) -> bool:
        """Check whether a group conversation is currently open.

        Requires at least one positive signal — group header buttons that only
        appear inside an open group chat.  The old fallback ("welcome screen is
        gone") was too permissive: DingTalk's contacts/search pages also have
        no welcome screen, causing false positives.
        """
        positive_signals = (
            "Files", "Group Settings", "Group Notice", "Chat History",
            # Chinese UI variants
            "文件", "群设置", "群公告", "聊天记录",
        )
        for btn_name in positive_signals:
            btn = find_control(
                self._window, "ButtonControl", timeout=2, Name=btn_name,
            )
            if btn:
                log.debug("Group header button found: %s", btn_name)
                return True

        log.debug(
            "_verify_group_opened: no group header buttons found "
            "(tried: %s).",
            positive_signals,
        )
        return False

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
        """Download a file.  Tries multiple strategies in order.

        Strategy order:
        1. Hover + fixed offset click (fastest, no VLM call needed)
        2. Hover + VLM icon detection (if offset click missed)
        3. Right-click + VLM context menu (fallback)
        4. Right-click + uiautomation MenuItemControl (legacy fallback)

        Returns True if the download was triggered.
        """
        # Ensure the file row is visible on screen
        if not self._scroll_file_into_view(file_info):
            log.error("File not visible, cannot download: %s", file_info.name)
            return False

        # Strategy 1 (primary): hover → click at fixed offset from right edge
        if self._download_via_hover_offset(file_info):
            return True

        # Strategy 2: hover → VLM finds download icon → click
        if self._download_via_vlm_hover(file_info):
            return True

        # Strategy 3: right-click → VLM finds "Download" in context menu
        if self._download_via_vlm(file_info):
            return True

        # Strategy 4 (fallback): uiautomation right-click → MenuItemControl
        if self._download_via_context_menu(file_info):
            return True

        log.error("Failed to download file: %s", file_info.name)
        return False

    # ── Hover + Fixed Offset Download (primary) ─────────────

    def _download_via_hover_offset(self, file_info: FileInfo) -> bool:
        """Hover over file row → click download icon at fixed offset from right edge.

        DingTalk shows action icons (download, share, etc.) when hovering over
        a file row.  The download icon is always at a predictable position
        relative to the row's right edge.  This approach needs no VLM call.

        The offset is configured via ``dingtalk.download_icon_offset`` in
        config.yaml (pixels LEFT of the row's right edge, default 8).
        The hover icons appear near the right edge of the row's
        BoundingRectangle, with the download icon (↓) very close to it.
        """
        try:
            rect = file_info.control.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                log.debug("File row has no valid BoundingRectangle for hover offset.")
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

            # Step 1: Hover over the file row center to trigger icon appearance
            log.debug(
                "Hover-offset: moving to row center (%d, %d) for %s",
                row_cx, row_cy, file_info.name,
            )
            pyautogui.moveTo(row_cx, row_cy)
            time.sleep(1.0)  # wait for hover action icons to render

            # Step 2: Click at fixed offset to the LEFT of the row's right edge.
            # The download icon (↓) sits near the right edge of the row's
            # BoundingRectangle, typically ~8 px to the left of rect.right.
            offset = self.dt.download_icon_offset
            click_x = rect.right - offset
            click_y = row_cy
            log.info(
                "Hover-offset: clicking download icon at (%d, %d) "
                "(right_edge=%d - offset=%d) for %s",
                click_x, click_y, rect.right, offset, file_info.name,
            )
            pyautogui.click(click_x, click_y)

            # Handle potential save dialog
            self._handle_save_dialog()
            time.sleep(self.dt.download_wait)
            log.info(
                "Download triggered (hover offset): %s", file_info.name,
            )
            return True

        except Exception as exc:
            log.warning("_download_via_hover_offset failed: %s", exc)
            return False

    # ── VLM Hover-based Download ──────────────────────────────

    def _download_via_vlm_hover(self, file_info: FileInfo) -> bool:
        """Hover over file row → action icons appear → VLM locates download icon → click.

        This avoids the context-menu-fading problem: the download icon is part
        of the row itself and only appears on hover, so we keep the mouse on
        the row while screenshotting.

        Steps:
        1. Move mouse to file row center (hover).
        2. Wait for hover action icons to render (~1 s).
        3. Screenshot the row area (with some right-side margin for icons).
        4. Send screenshot to VLM → get (x, y) of download icon.
        5. Click at those coordinates (icon is near the row, so mouse moves little).
        """
        from .vlm import find_icon_coords, grab_screenshot_base64

        vlm_cfg = self.cfg.vlm
        if not vlm_cfg.api_key:
            log.debug("VLM API key not configured, skipping hover strategy.")
            return False

        try:
            rect = file_info.control.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                log.debug("File row has no valid BoundingRectangle for hover.")
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

            # Step 1: Hover over the file row center
            log.debug(
                "Hovering over file row at (%d, %d) for %s",
                row_cx, row_cy, file_info.name,
            )
            pyautogui.moveTo(row_cx, row_cy)
            time.sleep(1.5)  # wait for hover action icons to render

            # Step 2: Screenshot the row area + generous right margin for icons
            margin = vlm_cfg.capture_margin
            # Capture the full row width plus extra right margin for hover icons
            cap_left = max(0, rect.left - 20)
            cap_top = max(0, rect.top - 10)
            cap_right = rect.right + margin  # icons appear on the right
            cap_bottom = rect.bottom + 10
            region = (cap_left, cap_top, cap_right, cap_bottom)

            screenshot_b64 = grab_screenshot_base64(region)
            log.debug(
                "Captured hover row screenshot (%dx%d) at (%d,%d).",
                cap_right - cap_left, cap_bottom - cap_top, cap_left, cap_top,
            )

            # Step 3: Ask VLM to find the download icon
            img_w = cap_right - cap_left
            img_h = cap_bottom - cap_top
            coords = find_icon_coords(
                api_key=vlm_cfg.api_key,
                screenshot_b64=screenshot_b64,
                target_description="download icon (a downward arrow ↓)",
                region_offset=(cap_left, cap_top),
                image_size=(img_w, img_h),
                model=vlm_cfg.model,
                base_url=vlm_cfg.base_url,
            )

            if not coords:
                log.warning("VLM could not locate download icon in hover row.")
                return False

            # Step 4: Click at the VLM-identified coordinates
            click_x, click_y = coords
            log.info(
                "VLM hover: clicking download icon at (%d, %d) for %s",
                click_x, click_y, file_info.name,
            )
            pyautogui.click(click_x, click_y)

            # Handle potential save dialog
            self._handle_save_dialog()
            time.sleep(self.dt.download_wait)
            log.info(
                "Download triggered (VLM hover): %s", file_info.name,
            )
            return True

        except Exception as exc:
            log.warning("_download_via_vlm_hover failed: %s", exc)
            return False

    # ── VLM Right-Click Download (fallback) ───────────────────

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
            img_w = cap_right - cap_left
            img_h = cap_bottom - cap_top
            coords = find_menu_item_coords(
                api_key=vlm_cfg.api_key,
                screenshot_b64=screenshot_b64,
                target_label="Download",
                region_offset=(cap_left, cap_top),
                image_size=(img_w, img_h),
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
                # Use a short timeout (0.3 s) — we just want a quick check,
                # not a full 1-second wait per button text.  Reducing this
                # from 1 s → 0.3 s cuts dismiss_dialogs from ~8 s/round
                # to ~2.4 s/round.
                btn = find_control(
                    self._window,
                    "ButtonControl",
                    timeout=0.3,
                    Name=btn_text,
                )
                if btn:
                    log.info("Dismissing dialog with button: %s", btn_text)
                    safe_click(btn, delay_after=0.3)
                    dismissed = True
            if not dismissed:
                break

    # ── Chat-Based Attachment Scanning & Download ─────────────

    def scan_chat_attachments(
        self, max_scrolls: int = 3,
    ) -> List[ChatAttachment]:
        """Scan the chat for downloadable file cards and images.

        Processes the current visible chat area, then scrolls up to load
        older messages.  Returns all attachments found across all pages.

        Args:
            max_scrolls: Number of scroll-up iterations to load history.

        Returns:
            List of ChatAttachment (may contain both file_card and image types).
        """
        if not self._window:
            return []

        all_attachments: List[ChatAttachment] = []
        seen_keys: set = set()

        for scroll_i in range(max_scrolls + 1):  # +1 for the initial view
            # Phase 1: file cards via VLM (CefBrowserWindow blocks UIA)
            cards = self._scan_file_cards_vlm()
            for card in cards:
                key = f"{card.timestamp}::{card.filename}"
                if key not in seen_keys:
                    seen_keys.add(key)
                    all_attachments.append(card)

            # Phase 2: images via VLM (optional, requires API key)
            images = self._scan_images_vlm()
            for img in images:
                key = f"{img.timestamp}::{img.filename}"
                if key not in seen_keys:
                    seen_keys.add(key)
                    all_attachments.append(img)

            if scroll_i < max_scrolls:
                self._scroll_chat_up()
                time.sleep(1.0)

        log.info(
            "Chat scan: found %d attachments (%d file cards, %d images) "
            "across %d pages.",
            len(all_attachments),
            sum(1 for a in all_attachments if a.msg_type == "file_card"),
            sum(1 for a in all_attachments if a.msg_type == "image"),
            max_scrolls + 1,
        )
        return all_attachments

    # ── File Card Scanning (VLM) ──────────────────────────────

    def _scan_file_cards_vlm(self) -> List[ChatAttachment]:
        """Find file cards by asking the VLM to locate Download buttons.

        DingTalk renders file cards inside CefBrowserWindow (Chromium),
        making them invisible to UIA.  We screenshot the chat area and
        ask the VLM to identify file cards with their filenames and
        Download button coordinates.
        """
        from .vlm import find_file_card_downloads, grab_screenshot_base64

        vlm_cfg = self.cfg.vlm
        if not vlm_cfg.api_key:
            log.debug("No VLM API key — skipping file card scan.")
            return []

        if not self._window:
            return []

        try:
            rect = self._window.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                return []

            # Capture the chat area (right ~70% of the window, skip sidebar)
            chat_left = rect.left + int(rect.width() * 0.3)
            region = (chat_left, rect.top, rect.right, rect.bottom)

            screenshot_b64 = grab_screenshot_base64(region)
            img_w = rect.right - chat_left
            img_h = rect.bottom - rect.top

            cards = find_file_card_downloads(
                api_key=vlm_cfg.api_key,
                screenshot_b64=screenshot_b64,
                image_size=(img_w, img_h),
                model=vlm_cfg.model,
                base_url=vlm_cfg.base_url,
            )

            results: List[ChatAttachment] = []
            for card in cards:
                # Convert image-relative coords to screen-absolute
                abs_x = card["x"] + chat_left
                abs_y = card["y"] + rect.top

                results.append(ChatAttachment(
                    msg_type="file_card",
                    filename=card["filename"],
                    timestamp=None,  # VLM doesn't reliably extract timestamps
                    download_click=(abs_x, abs_y),
                ))

            return results

        except Exception as exc:
            log.warning("VLM file card scan failed: %s", exc)
            return []

    # ── Image Scanning (VLM) ──────────────────────────────────

    def _scan_images_vlm(self) -> List[ChatAttachment]:
        """Find image attachments in the chat via VLM screenshot analysis.

        Screenshots the chat area and asks the VLM to identify inline
        image attachments (not avatars, stickers, or UI icons).
        """
        from .vlm import find_image_attachments, grab_screenshot_base64

        vlm_cfg = self.cfg.vlm
        if not vlm_cfg.api_key:
            return []

        if not self._window:
            return []

        try:
            rect = self._window.BoundingRectangle
            if rect.width() <= 0 or rect.height() <= 0:
                return []

            # Capture the chat area (right ~70% of the window, skip sidebar)
            chat_left = rect.left + int(rect.width() * 0.3)
            region = (chat_left, rect.top, rect.right, rect.bottom)

            screenshot_b64 = grab_screenshot_base64(region)
            img_w = rect.right - chat_left
            img_h = rect.bottom - rect.top

            image_boxes = find_image_attachments(
                api_key=vlm_cfg.api_key,
                screenshot_b64=screenshot_b64,
                image_size=(img_w, img_h),
                model=vlm_cfg.model,
                base_url=vlm_cfg.base_url,
            )

            results: List[ChatAttachment] = []
            for i, box in enumerate(image_boxes):
                # Convert from image-relative to screen-absolute coords
                abs_left = box[0] + chat_left
                abs_top = box[1] + rect.top
                abs_right = box[2] + chat_left
                abs_bottom = box[3] + rect.top

                # Generate a timestamp-based filename for images
                from datetime import datetime
                ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"image_{ts_str}_{i + 1}.png"

                results.append(ChatAttachment(
                    msg_type="image",
                    filename=filename,
                    timestamp=None,  # VLM can't reliably extract timestamps
                    image_bounds=(abs_left, abs_top, abs_right, abs_bottom),
                ))

            return results

        except Exception as exc:
            log.warning("VLM image scan failed: %s", exc)
            return []

    # ── Chat Attachment Download ──────────────────────────────

    def download_chat_attachment(self, att: ChatAttachment) -> bool:
        """Download a chat attachment using the appropriate strategy.

        - file_card: click the native Download button.
        - image: right-click → context menu → Download.

        Returns True if the download was triggered.
        """
        if att.msg_type == "file_card":
            return self._download_file_card(att)
        elif att.msg_type == "image":
            return self._download_image_attachment(att)
        log.warning("Unknown attachment type: %s", att.msg_type)
        return False

    def _download_file_card(self, att: ChatAttachment) -> bool:
        """Download a file card by clicking its Download button coordinates.

        The Download button is web-rendered in CefBrowserWindow (invisible
        to UIA).  We use the screen-absolute coordinates found by VLM.
        """
        if not att.download_click:
            log.warning("No Download click coords for: %s", att.filename)
            return False

        self._ensure_focus()

        cx, cy = att.download_click
        try:
            pyautogui.click(cx, cy)
            time.sleep(0.5)
            self._handle_save_dialog()
            time.sleep(self.dt.download_wait)
            log.info(
                "Download triggered (file card VLM click): %s", att.filename,
            )
            return True
        except Exception as exc:
            log.error(
                "Click failed for file card '%s' at (%d,%d): %s",
                att.filename, cx, cy, exc,
            )
            return False

    def _download_image_attachment(self, att: ChatAttachment) -> bool:
        """Download an image attachment via right-click context menu.

        Strategy 1: right-click → UIA MenuItemControl Name='Download'.
        Strategy 2: right-click → VLM finds 'Download' in context menu.
        """
        if not att.image_bounds:
            log.warning("No image bounds for: %s", att.filename)
            return False

        left, top, right, bottom = att.image_bounds
        cx = (left + right) // 2
        cy = (top + bottom) // 2

        self._ensure_focus()

        # Right-click the image center
        pyautogui.rightClick(cx, cy)
        time.sleep(1.0)

        # Strategy 1: UIA MenuItemControl Name='Download'
        sel = self.sel.context_menu_download
        menu_item = find_control(
            auto.GetRootControl(),
            sel.control_type,
            timeout=3,
            Name=sel.name,
        )
        if menu_item:
            if safe_click(menu_item, delay_after=0.5):
                self._handle_save_dialog()
                time.sleep(self.dt.download_wait)
                log.info(
                    "Image download triggered (UIA menu): %s", att.filename,
                )
                return True

        # Strategy 2: VLM fallback for context menu
        from .vlm import find_menu_item_coords, grab_screenshot_base64

        vlm_cfg = self.cfg.vlm
        if vlm_cfg.api_key:
            try:
                margin = vlm_cfg.capture_margin
                cap_left = max(0, cx - margin)
                cap_top = max(0, cy - margin)
                cap_right = cx + margin
                cap_bottom = cy + margin
                region = (cap_left, cap_top, cap_right, cap_bottom)

                screenshot_b64 = grab_screenshot_base64(region)
                img_w = cap_right - cap_left
                img_h = cap_bottom - cap_top

                coords = find_menu_item_coords(
                    api_key=vlm_cfg.api_key,
                    screenshot_b64=screenshot_b64,
                    target_label="Download",
                    region_offset=(cap_left, cap_top),
                    image_size=(img_w, img_h),
                    model=vlm_cfg.model,
                    base_url=vlm_cfg.base_url,
                )

                if coords:
                    pyautogui.click(coords[0], coords[1])
                    self._handle_save_dialog()
                    time.sleep(self.dt.download_wait)
                    log.info(
                        "Image download triggered (VLM menu): %s",
                        att.filename,
                    )
                    return True
            except Exception as exc:
                log.warning("VLM context menu fallback failed: %s", exc)

        # Dismiss lingering context menu
        pyautogui.press("escape")
        time.sleep(0.3)
        log.error("Failed to download image: %s", att.filename)
        return False

    # ── Public helpers for optimized scan loop ───────────────

    def get_chat_panel_screenshot(self) -> Tuple[str, int, int]:
        """Capture the chat panel as a base64 PNG.

        Uses mss (DirectX-based) instead of pyautogui so that DingTalk's
        hardware-accelerated CefBrowserWindow (chat messages) is captured
        correctly.  pyautogui uses GDI which renders the CEF area as blank.

        Crops out the left sidebar (~30 % of window width) to reduce the
        image size sent to Claude by ~30 %.

        Returns:
            (base64_png, offset_x, offset_y) where offset_x/offset_y is the
            top-left corner of the crop in screen coordinates.  Add these to
            any image-relative coordinates returned by ChatScanner to get
            screen-absolute values for pyautogui.click().
        """
        import mss
        import mss.tools

        if not self._window:
            with mss.mss() as sct:
                monitor = sct.monitors[1]
                img = sct.grab(monitor)
                png = mss.tools.to_png(img.rgb, img.size)
            return base64.standard_b64encode(png).decode(), 0, 0

        rect = self._window.BoundingRectangle
        # Skip the left sidebar (chat/contact list, ~30 % of window width)
        chat_left = rect.left + int(rect.width() * 0.30)
        chat_top = rect.top
        width = rect.right - chat_left
        height = rect.bottom - chat_top

        monitor = {"left": chat_left, "top": chat_top, "width": width, "height": height}
        with mss.mss() as sct:
            img = sct.grab(monitor)
            png = mss.tools.to_png(img.rgb, img.size)

        b64 = base64.standard_b64encode(png).decode()
        return b64, chat_left, chat_top

    def click_download_at(self, x: int, y: int) -> None:
        """Click a Download button at screen-absolute coordinates.

        Ensures DingTalk is in focus, performs the click, and handles any
        save/confirm dialog that may appear afterwards.
        """
        self._ensure_focus()
        pyautogui.click(x, y)
        time.sleep(0.5)
        self._handle_save_dialog()

    def scroll_chat_up(self, clicks: int = 5) -> None:
        """Scroll the chat panel upward to reveal older messages."""
        self._scroll_chat_up(clicks=clicks)

    # ── Chat Helpers ──────────────────────────────────────────

    def _ensure_focus(self) -> None:
        """Bring DingTalk window to foreground and set keyboard focus."""
        if self._window:
            try:
                self._window.SetActive()
                self._window.SetFocus()
                time.sleep(0.3)
            except Exception:
                pass

    def _scroll_chat_up(self, clicks: int = 5) -> None:
        """Scroll the chat area up to load older messages.

        Uses mouse wheel scroll positioned over the chat pane (right 65%
        of the window to avoid the sidebar).
        """
        if not self._window:
            return
        try:
            rect = self._window.BoundingRectangle
            # Position in the center-right area (chat pane)
            cx = rect.left + int(rect.width() * 0.65)
            cy = (rect.top + rect.bottom) // 2
            pyautogui.scroll(clicks, x=cx, y=cy)
        except Exception as exc:
            log.debug("_scroll_chat_up failed: %s", exc)
