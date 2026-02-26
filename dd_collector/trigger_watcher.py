"""Passive trigger watcher — detects DingTalk unread messages without UI disruption."""
from __future__ import annotations
import logging
import re
import time
from typing import Optional, Tuple
import uiautomation as auto
from .config import AppConfig

log = logging.getLogger("dd_collector")
_UNREAD_RE = re.compile(r'\(\d+\)')


class TriggerWatcher:
    def __init__(self, cfg: AppConfig) -> None:
        self._cfg = cfg
        self._tc = cfg.trigger
        self._window_class = cfg.dingtalk.window_class
        self._pil_available: Optional[bool] = None

    def has_unread(self) -> bool:
        try:
            if self._window_title_has_unread():
                return True
        except Exception as exc:
            log.debug("Tier-1 check error: %s", exc)
        try:
            if self._sidebar_has_red_badge():
                return True
        except Exception as exc:
            log.debug("Tier-2 check error: %s", exc)
        return False

    def wait_for_trigger(self) -> str:
        """Block until unread signal detected. Returns reason string."""
        interval = self._tc.check_interval_seconds
        log.info("TriggerWatcher: watching for unread messages (check every %ds). Ctrl+C to stop.", interval)
        while True:
            if self._window_title_has_unread():
                return "window title showed unread count"
            if self._sidebar_has_red_badge():
                return "sidebar red badge detected"
            for _ in range(interval):   # 1-second slices → responsive Ctrl+C
                time.sleep(1.0)

    def _window_title_has_unread(self) -> bool:
        win = self._get_window()
        if win is None:
            return False
        title = win.Name or ""
        return bool(_UNREAD_RE.search(title))

    def _sidebar_has_red_badge(self) -> bool:
        region = self._get_sidebar_region()
        if region is None:
            return False
        left, top, width, height = region
        if width <= 0 or height <= 0:
            return False
        # Lazy PIL availability check (cached after first attempt)
        if self._pil_available is False:
            return False
        if self._pil_available is None:
            try:
                from PIL import Image
                self._pil_available = True
            except ImportError:
                log.warning("PIL not available — Tier-2 pixel scan disabled.")
                self._pil_available = False
                return False
        try:
            import mss
            from PIL import Image
            monitor = {"left": left, "top": top, "width": width, "height": height}
            with mss.mss() as sct:
                raw = sct.grab(monitor)
                img = Image.frombytes("RGB", raw.size, raw.rgb)
        except Exception as exc:
            log.debug("mss/PIL grab failed: %s", exc)
            return False
        pixels = img.load()
        w, h = img.size
        threshold = self._tc.red_pixel_threshold
        count = 0
        for x in range(w):
            for y in range(h):
                r, g, b = pixels[x, y]
                if r > 200 and g < 80 and b < 80:
                    count += 1
                    if count >= threshold:
                        return True
        log.debug("Sidebar scan: %d red pixels (threshold=%d) — no badge.", count, threshold)
        return False

    def _get_window(self) -> Optional[auto.WindowControl]:
        for cls in [self._window_class, "StandardFrame_DingTalk", "DtMainFrameView"]:
            try:
                win = auto.WindowControl(ClassName=cls, searchDepth=1)
                if win.Exists(maxSearchSeconds=1):
                    return win
            except Exception:
                pass
        log.debug("TriggerWatcher: DingTalk window not found.")
        return None

    def _get_sidebar_region(self) -> Optional[Tuple[int, int, int, int]]:
        win = self._get_window()
        if win is None:
            return None
        try:
            rect = win.BoundingRectangle
        except Exception as exc:
            log.debug("BoundingRectangle error: %s", exc)
            return None
        if rect.width() <= 0 or rect.height() <= 0:
            log.debug("TriggerWatcher: window minimized — skipping pixel scan.")
            return None
        return (rect.left, rect.top, min(200, rect.width()), rect.height())
