"""VLM (Vision Language Model) helper for locating UI elements in screenshots.

Uses OpenRouter API with Qwen-VL-Plus to find clickable elements (like
context menu items) that are invisible to Windows UI Automation because
they are web-rendered inside DingTalk's CefBrowserWindow.
"""

from __future__ import annotations

import base64
import io
import logging
import re
from typing import Optional, Tuple

import mss
import mss.tools

log = logging.getLogger("dd_collector")


def grab_screenshot_base64(
    region: Optional[Tuple[int, int, int, int]] = None,
) -> str:
    """Capture a screenshot and return it as a base64-encoded PNG string.

    Args:
        region: Optional (left, top, right, bottom) in screen pixels.
                If None, captures the full primary monitor.

    Returns:
        Base64-encoded PNG data (no data-URI prefix).
    """
    with mss.mss() as sct:
        if region:
            left, top, right, bottom = region
            monitor = {"left": left, "top": top,
                       "width": right - left, "height": bottom - top}
        else:
            monitor = sct.monitors[1]  # primary monitor

        img = sct.grab(monitor)
        # Convert to PNG bytes
        png_bytes = mss.tools.to_png(img.rgb, img.size)
        return base64.b64encode(png_bytes).decode("ascii")


def find_menu_item_coords(
    api_key: str,
    screenshot_b64: str,
    target_label: str = "Download",
    region_offset: Optional[Tuple[int, int]] = None,
    model: str = "qwen/qwen-vl-plus",
    base_url: str = "https://openrouter.ai/api/v1",
) -> Optional[Tuple[int, int]]:
    """Ask a VLM to locate a menu item in a screenshot.

    Args:
        api_key: OpenRouter API key.
        screenshot_b64: Base64-encoded PNG screenshot.
        target_label: The menu item text to find (e.g. "Download" or "下载").
        region_offset: (left, top) offset to add to the returned coordinates
                       so they map back to screen-absolute pixels.
        model: OpenRouter model identifier.
        base_url: OpenRouter API base URL.

    Returns:
        (x, y) screen coordinates of the menu item center, or None if not found.
    """
    from openai import OpenAI

    client = OpenAI(base_url=base_url, api_key=api_key)

    prompt = (
        f"Look at this screenshot of a right-click context menu. "
        f"Find the menu item labeled \"{target_label}\" (or \"下载\" in Chinese). "
        f"Return ONLY the pixel coordinates of the center of that menu item "
        f"in the format: x,y\n"
        f"For example: 320,185\n"
        f"If you cannot find it, reply: NOT_FOUND"
    )

    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{screenshot_b64}",
                            },
                        },
                    ],
                }
            ],
            max_tokens=50,
            temperature=0,
        )

        reply = resp.choices[0].message.content.strip()
        log.debug("VLM reply: %s", reply)

        if "NOT_FOUND" in reply.upper():
            log.warning("VLM could not find '%s' in screenshot.", target_label)
            return None

        # Parse "x,y" from the reply (may contain extra text)
        m = re.search(r"(\d+)\s*,\s*(\d+)", reply)
        if not m:
            log.warning("VLM reply not parseable as coordinates: %s", reply)
            return None

        x, y = int(m.group(1)), int(m.group(2))

        # Add region offset so coords map to absolute screen position
        if region_offset:
            x += region_offset[0]
            y += region_offset[1]

        log.info("VLM found '%s' at (%d, %d).", target_label, x, y)
        return (x, y)

    except Exception as exc:
        log.error("VLM API call failed: %s", exc)
        return None
