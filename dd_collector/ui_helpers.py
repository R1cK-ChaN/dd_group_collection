"""Reusable uiautomation wrappers for finding and interacting with controls."""

from __future__ import annotations

import logging
import time
from typing import Any, Optional

import uiautomation as auto

log = logging.getLogger("dd_collector")


def find_control(
    parent: auto.Control,
    control_type: str,
    timeout: float = 5,
    **kwargs: Any,
) -> Optional[auto.Control]:
    """Find a child control by type and optional properties.

    Args:
        parent: Parent control to search within.
        control_type: UIA control type name, e.g. "EditControl", "ButtonControl".
        timeout: Max seconds to wait for the control.
        **kwargs: Additional properties passed to the control constructor
                  (e.g. Name="搜索", AutomationId="searchBox").

    Returns:
        The found control, or None if not found within timeout.
    """
    ctrl_class = getattr(auto, control_type, None)
    if ctrl_class is None:
        log.error("Unknown control type: %s", control_type)
        return None

    search_props = {"searchDepth": kwargs.pop("searchDepth", 10)}
    search_props.update(kwargs)

    try:
        ctrl = parent.Control(ctrl_class, **search_props) if False else ctrl_class(
            searchFromControl=parent, **search_props
        )
        if ctrl.Exists(maxSearchSeconds=timeout):
            return ctrl
    except Exception as exc:
        log.debug("find_control(%s, %s) error: %s", control_type, kwargs, exc)

    return None


def safe_click(control: auto.Control, delay_after: float = 0.5) -> bool:
    """Click a control safely and wait.

    Returns True on success, False on failure.
    """
    try:
        control.Click(simulateMove=False)
        if delay_after > 0:
            time.sleep(delay_after)
        return True
    except Exception as exc:
        log.warning("safe_click failed: %s", exc)
        return False


def safe_right_click(control: auto.Control, delay_after: float = 0.5) -> bool:
    """Right-click a control safely and wait."""
    try:
        control.RightClick(simulateMove=False)
        if delay_after > 0:
            time.sleep(delay_after)
        return True
    except Exception as exc:
        log.warning("safe_right_click failed: %s", exc)
        return False


def set_text(control: auto.Control, text: str, delay_after: float = 0.5) -> bool:
    """Set text on an edit control using ValuePattern, falling back to SendKeys."""
    try:
        vp = control.GetValuePattern()
        if vp:
            vp.SetValue(text)
        else:
            control.SendKeys(text, interval=0.05)
        if delay_after > 0:
            time.sleep(delay_after)
        return True
    except Exception as exc:
        log.warning("set_text failed, trying SendKeys: %s", exc)
        try:
            control.SendKeys(text, interval=0.05)
            if delay_after > 0:
                time.sleep(delay_after)
            return True
        except Exception as exc2:
            log.error("set_text failed completely: %s", exc2)
            return False


def scroll_to_bottom(control: auto.Control, max_scrolls: int = 50) -> None:
    """Scroll a control to the bottom using ScrollPattern, with PageDown fallback.

    Scrolls up to max_scrolls times to prevent infinite loops.
    """
    try:
        sp = control.GetScrollPattern()
        if sp:
            for _ in range(max_scrolls):
                vert = sp.VerticalScrollPercent
                if vert >= 99.0 or vert < 0:
                    break
                sp.Scroll(auto.ScrollAmount.NoAmount, auto.ScrollAmount.LargeIncrement)
                time.sleep(0.3)
            return
    except Exception:
        pass

    # Fallback: send PageDown keys
    try:
        control.SetFocus()
        for _ in range(max_scrolls):
            control.SendKeys("{PageDown}")
            time.sleep(0.3)
    except Exception as exc:
        log.debug("scroll_to_bottom fallback failed: %s", exc)


def send_escape() -> None:
    """Press Escape key to dismiss dialogs or cancel searches."""
    try:
        auto.SendKeys("{Esc}")
        time.sleep(0.3)
    except Exception as exc:
        log.debug("send_escape failed: %s", exc)
