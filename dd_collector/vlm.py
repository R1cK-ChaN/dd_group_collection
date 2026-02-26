"""VLM (Vision Language Model) helper for locating UI elements in screenshots.

Uses OpenRouter API with Qwen-VL-Plus to find clickable elements (like
context menu items) that are invisible to Windows UI Automation because
they are web-rendered inside DingTalk's CefBrowserWindow.
"""

from __future__ import annotations

import base64
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


def _validate_coords(
    x: int,
    y: int,
    image_size: Optional[Tuple[int, int]],
    label: str,
) -> bool:
    """Check if coordinates fall within image bounds.

    Args:
        x, y: Coordinates to validate (relative to image, before offset).
        image_size: (width, height) of the screenshot image.
        label: Description for log messages.

    Returns:
        True if coordinates are valid (within bounds or no size given).
    """
    if image_size is None:
        return True

    w, h = image_size
    if x < 0 or x >= w or y < 0 or y >= h:
        log.warning(
            "VLM returned out-of-bounds coordinates (%d, %d) for %s "
            "(image size: %dx%d). Rejecting.",
            x, y, label, w, h,
        )
        return False
    return True


def find_menu_item_coords(
    api_key: str,
    screenshot_b64: str,
    target_label: str = "Download",
    region_offset: Optional[Tuple[int, int]] = None,
    image_size: Optional[Tuple[int, int]] = None,
    model: str = "qwen/qwen3.5-35b-a3b",
    base_url: str = "https://openrouter.ai/api/v1",
) -> Optional[Tuple[int, int]]:
    """Ask a VLM to locate a menu item in a screenshot.

    Args:
        api_key: OpenRouter API key.
        screenshot_b64: Base64-encoded PNG screenshot.
        target_label: The menu item text to find (e.g. "Download" or "下载").
        region_offset: (left, top) offset to add to the returned coordinates
                       so they map back to screen-absolute pixels.
        image_size: (width, height) of the screenshot for bounds checking.
        model: OpenRouter model identifier.
        base_url: OpenRouter API base URL.

    Returns:
        (x, y) screen coordinates of the menu item center, or None if not found.
    """
    from openai import OpenAI

    client = OpenAI(base_url=base_url, api_key=api_key)

    # Include image dimensions in the prompt so the VLM knows the coordinate space
    size_hint = ""
    if image_size:
        size_hint = f"The image is {image_size[0]}x{image_size[1]} pixels. "

    prompt = (
        f"Look at this screenshot of a right-click context menu. "
        f"{size_hint}"
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

        # Bounds check before applying offset
        if not _validate_coords(x, y, image_size, target_label):
            return None

        # Add region offset so coords map to absolute screen position
        if region_offset:
            x += region_offset[0]
            y += region_offset[1]

        log.info("VLM found '%s' at (%d, %d).", target_label, x, y)
        return (x, y)

    except Exception as exc:
        log.error("VLM API call failed: %s", exc)
        return None


def find_icon_coords(
    api_key: str,
    screenshot_b64: str,
    target_description: str = "download icon (a downward arrow ↓)",
    region_offset: Optional[Tuple[int, int]] = None,
    image_size: Optional[Tuple[int, int]] = None,
    model: str = "qwen/qwen3.5-35b-a3b",
    base_url: str = "https://openrouter.ai/api/v1",
) -> Optional[Tuple[int, int]]:
    """Ask a VLM to locate an icon in a screenshot of a file row.

    Used for hover-based download: when the mouse hovers over a file row,
    action icons (download, share, etc.) appear on the right side.  This
    function asks the VLM to find the download icon.

    Args:
        api_key: OpenRouter API key.
        screenshot_b64: Base64-encoded PNG screenshot of the file row area.
        target_description: Description of the icon to find.
        region_offset: (left, top) offset to add to returned coordinates.
        image_size: (width, height) of the screenshot for bounds checking.
        model: OpenRouter model identifier.
        base_url: OpenRouter API base URL.

    Returns:
        (x, y) screen coordinates of the icon center, or None if not found.
    """
    from openai import OpenAI

    client = OpenAI(base_url=base_url, api_key=api_key)

    # Include image dimensions in the prompt so the VLM knows the coordinate space
    size_hint = ""
    if image_size:
        size_hint = (
            f"The image is {image_size[0]}x{image_size[1]} pixels. "
            f"Your coordinates must be within 0-{image_size[0]-1} for x "
            f"and 0-{image_size[1]-1} for y. "
        )

    prompt = (
        "Look at this screenshot of a file list row from DingTalk. "
        f"{size_hint}"
        "When the mouse hovers over a file row, action icons appear on the "
        "right side of the row. "
        f"Find the {target_description}. It is typically a small clickable "
        "icon on the right portion of the row. Look for a downward-pointing "
        "arrow icon (↓) which represents download. It may also be labeled "
        "\"下载\" in Chinese.\n"
        "Return ONLY the pixel coordinates of the center of that icon "
        "in the format: x,y\n"
        "For example: 850,25\n"
        "If you cannot find any download icon or arrow, reply: NOT_FOUND"
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
        log.debug("VLM icon reply: %s", reply)

        if "NOT_FOUND" in reply.upper():
            log.warning("VLM could not find download icon in screenshot.")
            return None

        # Parse "x,y" from the reply (may contain extra text)
        m = re.search(r"(\d+)\s*,\s*(\d+)", reply)
        if not m:
            log.warning("VLM icon reply not parseable as coordinates: %s", reply)
            return None

        x, y = int(m.group(1)), int(m.group(2))

        # Bounds check before applying offset
        if not _validate_coords(x, y, image_size, "download icon"):
            return None

        # Add region offset so coords map to absolute screen position
        if region_offset:
            x += region_offset[0]
            y += region_offset[1]

        log.info("VLM found download icon at (%d, %d).", x, y)
        return (x, y)

    except Exception as exc:
        log.error("VLM icon API call failed: %s", exc)
        return None


def find_image_attachments(
    api_key: str,
    screenshot_b64: str,
    image_size: Optional[Tuple[int, int]] = None,
    model: str = "qwen/qwen3.5-35b-a3b",
    base_url: str = "https://openrouter.ai/api/v1",
) -> list:
    """Ask a VLM to identify inline image attachments in a chat screenshot.

    Returns a list of (left, top, right, bottom) bounding boxes for each
    image attachment found.  Coordinates are relative to the screenshot.

    Only returns actual shared image/photo/screenshot attachments — not
    user avatars, UI icons, stickers, or file cards.
    """
    from openai import OpenAI

    client = OpenAI(base_url=base_url, api_key=api_key)

    size_hint = ""
    if image_size:
        size_hint = (
            f"The image is {image_size[0]}x{image_size[1]} pixels. "
            f"Coordinates must be within 0-{image_size[0] - 1} for x "
            f"and 0-{image_size[1] - 1} for y. "
        )

    prompt = (
        "Look at this screenshot of a DingTalk chat conversation. "
        f"{size_hint}"
        "Identify all INLINE IMAGE attachments shared by chat participants. "
        "These are rectangular photo/screenshot/document image previews "
        "in the message area.\n"
        "Do NOT include:\n"
        "- Small circular user avatars on the left\n"
        "- UI icons, buttons, stickers, or emoji\n"
        "- File cards (rectangles with filename, size, and "
        "Download/Add buttons)\n\n"
        "For each image attachment, return its bounding box on a new line:\n"
        "image_1: left,top,right,bottom\n"
        "image_2: left,top,right,bottom\n\n"
        "Example:\n"
        "image_1: 520,100,750,300\n\n"
        "If no image attachments are found, reply: NO_IMAGES"
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
            max_tokens=200,
            temperature=0,
        )

        reply = resp.choices[0].message.content.strip()
        log.debug("VLM image scan reply: %s", reply)

        if "NO_IMAGES" in reply.upper():
            return []

        # Parse bounding boxes from "image_N: left,top,right,bottom" lines
        boxes: list = []
        for line in reply.split("\n"):
            m = re.search(
                r"(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)", line,
            )
            if m:
                left = int(m.group(1))
                top = int(m.group(2))
                right = int(m.group(3))
                bottom = int(m.group(4))

                # Clamp to image bounds
                if image_size:
                    w, h = image_size
                    if left >= w or top >= h or right <= 0 or bottom <= 0:
                        continue
                    left = max(0, min(left, w - 1))
                    top = max(0, min(top, h - 1))
                    right = max(0, min(right, w))
                    bottom = max(0, min(bottom, h))

                if right > left and bottom > top:
                    boxes.append((left, top, right, bottom))

        log.info("VLM found %d image attachments.", len(boxes))
        return boxes

    except Exception as exc:
        log.error("VLM image scan API call failed: %s", exc)
        return []
