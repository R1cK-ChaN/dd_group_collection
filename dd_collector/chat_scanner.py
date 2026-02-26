"""Claude vision scanner for DingTalk chat attachments.

Single responsibility: given a PNG screenshot of the chat panel, ask Claude
to locate every file-attachment card's Download button and return coordinates.

Uses the plain Messages API (no computer-use beta, no tool loop) so it is
compatible with any Claude model that supports vision, including Haiku which
is ~20x cheaper than Sonnet for this simple coordinate-extraction task.

Typical token cost per call:
  - Input image (chat panel crop ~896x720):  ~700 image tokens
  - Input prompt text:                        ~80 tokens
  - Output JSON (10 files):                   ~100 tokens
  Total: ~880 tokens ≈ $0.00088 at Haiku pricing (vs ~$0.018 at Sonnet)
"""

from __future__ import annotations

import json
import logging
import re
from typing import Any, Dict, List, Optional, Tuple

import anthropic

log = logging.getLogger("dd_collector")

DEFAULT_MODEL = "claude-haiku-4-5-20251001"

# ── Prompt ────────────────────────────────────────────────────
# Tight, JSON-only instruction.  No explanation requested → fewer output tokens.
_PROMPT = """\
Analyze this DingTalk (钉钉) chat panel screenshot.

Find every FILE attachment card that has a visible Download button.
Download buttons are labeled "Download" (English) or "下载" (Chinese).
They appear as buttons/links attached to file cards (PDF, Excel, Word, etc).
Do NOT include image thumbnails — only file cards.

Reply with ONLY this JSON object (no markdown, no explanation):
{"files": [{"filename": "report.pdf", "x": 712, "y": 340}, ...]}

- filename: the file name shown on the card (include extension)
- x, y: pixel coordinates of the center of the Download button in this image

If no download buttons are visible: {"files": []}"""


class ChatScanner:
    """Ask Claude to locate Download buttons in a chat panel screenshot.

    Each call to ``find_downloads()`` makes exactly ONE API request and
    returns structured results.  All navigation, scrolling, and clicking
    are handled programmatically by the caller — this class only does vision.
    """

    def __init__(self, api_key: str, model: str = DEFAULT_MODEL) -> None:
        self._client = anthropic.Anthropic(api_key=api_key)
        self._model = model

    def find_downloads(
        self,
        screenshot_b64: str,
        region_offset: Tuple[int, int] = (0, 0),
        debug_save_path: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Identify file-card Download buttons in a chat panel screenshot.

        Args:
            screenshot_b64: Base64-encoded PNG of the chat panel crop.
            region_offset:  (left, top) pixel offset of the crop's top-left
                            corner in screen coordinates.  All returned x/y
                            values are translated by this offset so callers
                            receive screen-absolute coordinates ready for
                            ``pyautogui.click()``.

        Returns:
            List of dicts: ``{"filename": str, "x": int, "y": int}``
            Coordinates are screen-absolute.  Empty list if none found or
            on any API error (logged as warning, never raises).
        """
        # Optionally save the screenshot for debugging
        if debug_save_path:
            import base64 as _b64
            with open(debug_save_path, "wb") as _f:
                _f.write(_b64.b64decode(screenshot_b64))
            log.debug("Debug screenshot saved: %s", debug_save_path)

        try:
            response = self._client.messages.create(
                model=self._model,
                max_tokens=512,   # JSON-only response; 512 is generous
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": _PROMPT},
                            {
                                "type": "image",
                                "source": {
                                    "type": "base64",
                                    "media_type": "image/png",
                                    "data": screenshot_b64,
                                },
                            },
                        ],
                    }
                ],
            )
        except anthropic.APIError as exc:
            log.error("ChatScanner API error: %s", exc)
            return []

        raw = "".join(
            block.text for block in response.content if hasattr(block, "text")
        ).strip()
        log.debug("ChatScanner raw response (%d chars): %s", len(raw), raw[:300])

        data = _extract_json(raw)
        if data is None:
            log.warning("ChatScanner: could not parse JSON from response: %s", raw[:300])
            return []

        files = data.get("files", [])
        if not isinstance(files, list):
            log.warning("ChatScanner: 'files' is not a list: %s", data)
            return []

        ox, oy = region_offset
        results: List[Dict[str, Any]] = []
        for item in files:
            try:
                results.append({
                    "filename": str(item.get("filename", "unknown")),
                    "x": int(item["x"]) + ox,
                    "y": int(item["y"]) + oy,
                })
            except (KeyError, ValueError, TypeError) as exc:
                log.debug("ChatScanner: skipping malformed entry %s: %s", item, exc)

        log.info(
            "ChatScanner (%s): %d download button(s) found", self._model, len(results)
        )
        return results


# ── JSON extraction helper ────────────────────────────────────

def _extract_json(text: str) -> Any:
    """Parse JSON from text that may contain code fences or extra prose."""
    text = text.strip()

    # Strip markdown code fences: ```json ... ``` or ``` ... ```
    text = re.sub(r"^```(?:json)?\s*\n?", "", text)
    text = re.sub(r"\n?```\s*$", "", text)
    text = text.strip()

    # Direct parse (happy path)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Find first {...} block in case there's surrounding prose
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group())
        except json.JSONDecodeError:
            pass

    return None
