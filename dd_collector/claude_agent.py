"""Autonomous DingTalk file-download agent using Claude computer use API.

Every screenshot and action is logged to logs/autonomous/<session>/ so the
full sequence can be replayed and analysed for improvement.
"""

from __future__ import annotations

import base64
import io
import json
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import anthropic
import pyautogui

log = logging.getLogger("dd_collector")

# Safety cap: agent stops after this many API round-trips per group
MAX_ITERATIONS = 60


class ClaudeAgent:
    """Autonomous agent using Claude's computer use API (computer_20250124).

    Uses a regular Anthropic API key (sk-ant-api03-…).
    Saves every screenshot to logs/autonomous/<session>/ for analysis.
    """

    def __init__(
        self,
        oauth_token: str,
        model: str = "claude-opus-4-5-20251101",
        log_dir: str = "logs",
        base_url: Optional[str] = None,
    ) -> None:
        if oauth_token.startswith("sk-ant-oat"):
            raise ValueError(
                "The provided token is a Claude Code OAuth token (sk-ant-oat01-…), "
                "which is NOT accepted by the Anthropic REST API.\n"
                "Please set claude.oauth_token in config.yaml to a regular API key "
                "(sk-ant-api03-…) from https://console.anthropic.com."
            )
        kwargs: Dict[str, Any] = {"api_key": oauth_token}
        if base_url:
            kwargs["base_url"] = base_url
            log.info("ClaudeAgent base_url: %s", base_url)
        self._client = anthropic.Anthropic(**kwargs)
        self._model = model

        # Per-session log directory: logs/autonomous/YYYYMMDD_HHMMSS/
        ts = time.strftime("%Y%m%d_%H%M%S")
        self._session_dir = Path(log_dir) / "autonomous" / ts
        self._session_dir.mkdir(parents=True, exist_ok=True)
        log.info("ClaudeAgent session dir: %s", self._session_dir)
        log.info("ClaudeAgent model: %s", self._model)

        self._screenshot_idx = 0
        self._current_group = "run"
        self._action_log: List[Dict[str, Any]] = []

    # ── Screenshot helper ────────────────────────────────────────

    def _take_screenshot(self) -> str:
        """Capture full screen, save PNG to session dir, return base64."""
        screenshot = pyautogui.screenshot()
        buf = io.BytesIO()
        screenshot.save(buf, format="PNG")
        b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")

        safe_name = self._current_group.replace("/", "_").replace("\\", "_")
        fname = f"{safe_name}_{self._screenshot_idx:04d}.png"
        path = self._session_dir / fname
        screenshot.save(str(path), "PNG")
        log.info("  [screenshot %04d] saved → %s", self._screenshot_idx, path.name)
        self._screenshot_idx += 1

        return b64

    # ── Action execution ─────────────────────────────────────────

    def _execute_action(self, action: Dict[str, Any]) -> Optional[str]:
        """Execute one computer-use action. Returns base64 screenshot or None."""
        act = action.get("action", "")
        coord = action.get("coordinate", [0, 0])

        # Log full action at INFO so it appears in run_claude.log
        log.info("  [action] %s | %s", act, json.dumps(
            {k: v for k, v in action.items() if k != "action"}, ensure_ascii=False
        ))
        self._action_log.append({"group": self._current_group, "action": action})

        if act == "screenshot":
            return self._take_screenshot()

        if act in ("left_click", "right_click", "double_click", "mouse_move", "middle_click"):
            x, y = int(coord[0]), int(coord[1])
            if act == "left_click":
                pyautogui.click(x, y, button="left")
            elif act == "right_click":
                pyautogui.click(x, y, button="right")
            elif act == "double_click":
                pyautogui.doubleClick(x, y)
            elif act == "middle_click":
                pyautogui.click(x, y, button="middle")
            elif act == "mouse_move":
                pyautogui.moveTo(x, y)
            time.sleep(0.4)
            return None

        if act == "left_click_drag":
            start = action.get("start_coordinate", [0, 0])
            end = coord
            pyautogui.mouseDown(int(start[0]), int(start[1]))
            time.sleep(0.1)
            pyautogui.dragTo(int(end[0]), int(end[1]), duration=0.3, button="left")
            pyautogui.mouseUp()
            time.sleep(0.4)
            return None

        if act == "type":
            text = action.get("text", "")
            log.info("    text=%r", text[:80])
            try:
                import pyperclip
                pyperclip.copy(text)
                pyautogui.hotkey("ctrl", "v")
            except Exception:
                pyautogui.typewrite(text, interval=0.05)
            time.sleep(0.3)
            return None

        if act == "key":
            # Both "key" and legacy "text" field names
            key_str = action.get("key", action.get("text", "")).strip()
            log.info("    key=%r", key_str)
            key_map = {
                "return": "enter", "super": "win", "ctrl": "ctrl",
                "alt": "alt", "shift": "shift", "escape": "escape",
                "esc": "escape", "backspace": "backspace", "tab": "tab",
                "delete": "delete", "del": "delete",
                "page_up": "pageup", "page_down": "pagedown",
                "home": "home", "end": "end",
                "up": "up", "down": "down", "left": "left", "right": "right",
            }
            parts = [key_map.get(p.lower(), p.lower()) for p in key_str.split("+")]
            if len(parts) == 1:
                pyautogui.press(parts[0])
            else:
                pyautogui.hotkey(*parts)
            time.sleep(0.3)
            return None

        if act == "scroll":
            x, y = int(coord[0]), int(coord[1])
            # Support both computer_20241022 (direction/amount) and computer_20250124 (scroll_direction/scroll_amount)
            direction = action.get("direction") or action.get("scroll_direction", "down")
            amount = int(action.get("amount") or action.get("scroll_amount", 3))
            if direction == "down":
                pyautogui.scroll(-amount, x=x, y=y)
            elif direction == "up":
                pyautogui.scroll(amount, x=x, y=y)
            elif direction == "right":
                pyautogui.hscroll(amount, x=x, y=y)
            elif direction == "left":
                pyautogui.hscroll(-amount, x=x, y=y)
            time.sleep(0.5)
            return None

        if act == "cursor_position":
            x, y = pyautogui.position()
            return f"Cursor at ({x}, {y})"

        if act == "hold_key":
            key_str = action.get("key", action.get("text", ""))
            duration = float(action.get("duration", 0.5))
            pyautogui.keyDown(key_str)
            time.sleep(duration)
            pyautogui.keyUp(key_str)
            time.sleep(0.2)
            return None

        log.warning("  [action] unknown action type: %r", act)
        return None

    # ── Main agent loop ──────────────────────────────────────────

    def run_download_task(
        self,
        group_name: str,
        download_dir: str,
        already_downloaded: List[str],
        max_scrolls: int = 5,
    ) -> None:
        """Run Claude autonomously to download all new files from one group.

        Claude controls the full screen: navigation, scrolling, clicking
        download buttons, and handling any dialogs.  Every screenshot is
        saved to the session log dir.
        """
        self._current_group = group_name
        self._screenshot_idx = 0

        # Dynamic screen resolution
        screen_w, screen_h = pyautogui.size()
        log.info(
            "ClaudeAgent.run_download_task: group=%r model=%s screen=%dx%d",
            group_name, self._model, screen_w, screen_h,
        )

        computer_tool: Dict[str, Any] = {
            "type": "computer_20250124",
            "name": "computer",
            "display_width_px": screen_w,
            "display_height_px": screen_h,
        }

        already_str = (
            "\n".join(f"  - {f}" for f in already_downloaded[:50])
            if already_downloaded else "  (none)"
        )

        task_prompt = (
            f"You are automating the DingTalk (钉钉) PC application on Windows "
            f"to download files from a group chat.\n\n"
            f"TASK: Download all NEW files from DingTalk group \"{group_name}\".\n\n"
            f"FILES ALREADY DOWNLOADED (skip these):\n{already_str}\n\n"
            f"STEP-BY-STEP INSTRUCTIONS:\n"
            f"1. Take a screenshot to see the current screen state.\n"
            f"2. Find the DingTalk window. If minimized, click its taskbar button.\n"
            f"3. Navigate to group \"{group_name}\":\n"
            f"   a. Look in the left sidebar for the group name and click it.\n"
            f"   b. If not visible: click the search box at the TOP of the sidebar,\n"
            f"      type \"{group_name}\", then click the matching result.\n"
            f"   c. If a full-screen search overlay appears, press Escape and retry.\n"
            f"4. Once the group chat is open, scan for TWO types of downloadable content:\n"
            f"   TYPE A — File attachment cards (PDF, XLSX, DOCX, etc.):\n"
            f"   - Cards show: file icon + filename + file size + Download (下载) button.\n"
            f"   - Hover over a card to reveal the Download button if not visible.\n"
            f"   - Click the Download button directly.\n"
            f"   TYPE B — Inline images / screenshots shared in chat:\n"
            f"   - These appear as image thumbnails or previews in the chat.\n"
            f"   - To download: RIGHT-CLICK on the image → select 'Save As' or '另存为'\n"
            f"     or '下载' from the context menu → click Save in the dialog.\n"
            f"   - The filename shown in the Save As dialog is the actual filename to check\n"
            f"     against the already-downloaded list.\n"
            f"5. For each NEW item (not in the already-downloaded list):\n"
            f"   a. File card: click its Download (下载) button.\n"
            f"   b. Inline image: right-click → Save As / 另存为 / 下载 → Save.\n"
            f"   c. FILENAME RULE — When the Save As dialog appears, rename the file\n"
            f"      to include the message's upload timestamp before saving:\n"
            f"      - Note the timestamp shown on the chat message (e.g. '02/26 17:30')\n"
            f"        BEFORE you click Download / right-click the image.\n"
            f"      - In the Save As dialog: triple-click the filename field to select all,\n"
            f"        then type the new name in this format:\n"
            f"        {{original_name}}_{{YYYY-MM-DD_HH-mm}}{{extension}}\n"
            f"        Examples:\n"
            f"          '路透晚报.pdf'  uploaded at '02/26 17:30'  →  '路透晚报_2026-02-26_17-30.pdf'\n"
            f"          'IMG_1234.jpg'  uploaded at '02/23 15:21'  →  'IMG_1234_2026-02-23_15-21.jpg'\n"
            f"        Convert DingTalk date '02/26' using current year (2026).\n"
            f"        Use hyphens in the time part (17-30 not 17:30) so it is filename-safe.\n"
            f"        If the timestamp says 'Yesterday' use today's date minus one day.\n"
            f"        If the timestamp says 'Today' use today's date.\n"
            f"      - Then click Save.\n"
            f"   d. Wait 1-2 seconds after each download.\n"
            f"6. Scroll UP in the chat to see older messages. Repeat for ~{max_scrolls} screens.\n"
            f"7. When all visible content has been processed, say: TASK COMPLETE\n\n"
            f"NOTES:\n"
            f"- '下载' means Download, '另存为' means Save As in Chinese.\n"
            f"- Do NOT open the Files tab — only scan chat messages.\n"
            f"- If an update/notification dialog appears, dismiss it first "
            f"  (click 'Later', '稍后', 'Cancel', '取消', or press Escape).\n"
            f"- Download directory: {download_dir}\n"
            f"- Start by taking a screenshot now."
        )

        # Attach initial screenshot to the first message
        initial_screenshot = self._take_screenshot()
        messages: List[Dict[str, Any]] = [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": task_prompt},
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": initial_screenshot,
                        },
                    },
                ],
            }
        ]

        for iteration in range(MAX_ITERATIONS):
            log.info(
                "─── ClaudeAgent iteration %d / %d  [group=%s] ───",
                iteration + 1, MAX_ITERATIONS, group_name,
            )

            try:
                response = self._client.beta.messages.create(
                    model=self._model,
                    max_tokens=4096,
                    tools=[computer_tool],
                    messages=messages,
                    betas=["computer-use-2025-01-24"],
                )
            except anthropic.APIError as exc:
                log.error("ClaudeAgent API error (iter %d): %s", iteration + 1, exc)
                break

            # Log usage if available
            if hasattr(response, "usage") and response.usage:
                u = response.usage
                log.info(
                    "  [tokens] input=%s output=%s cache_read=%s cache_create=%s",
                    getattr(u, "input_tokens", "?"),
                    getattr(u, "output_tokens", "?"),
                    getattr(u, "cache_read_input_tokens", "-"),
                    getattr(u, "cache_creation_input_tokens", "-"),
                )

            log.info(
                "  [response] stop_reason=%s blocks=%d",
                response.stop_reason, len(response.content),
            )

            # Add assistant turn to conversation
            messages.append({"role": "assistant", "content": response.content})

            # Log and check any text blocks
            for block in response.content:
                if hasattr(block, "text") and block.text:
                    log.info("  [Claude] %s", block.text.strip())

            # Check for task-complete signal
            if response.stop_reason == "end_turn":
                full_text = " ".join(
                    block.text for block in response.content if hasattr(block, "text")
                )
                if "TASK COMPLETE" in full_text.upper():
                    log.info("ClaudeAgent: TASK COMPLETE for group '%s'.", group_name)
                else:
                    log.info("ClaudeAgent: end_turn for group '%s'.", group_name)
                break

            if response.stop_reason != "tool_use":
                log.warning(
                    "ClaudeAgent: unexpected stop_reason=%r — stopping.", response.stop_reason
                )
                break

            # Execute tool calls and collect results
            tool_results: List[Dict[str, Any]] = []
            for block in response.content:
                if not hasattr(block, "type") or block.type != "tool_use":
                    continue
                if block.name != "computer":
                    log.warning("ClaudeAgent: unexpected tool name=%r", block.name)
                    continue

                result = self._execute_action(block.input)

                if result is None or (isinstance(result, str) and result.startswith("Cursor")):
                    # Take a fresh screenshot after the action
                    time.sleep(0.5)
                    screenshot_b64 = self._take_screenshot()
                    content: Any = [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": screenshot_b64,
                            },
                        }
                    ]
                    if result and result.startswith("Cursor"):
                        content.insert(0, {"type": "text", "text": result})
                else:
                    # result is already a screenshot base64
                    content = [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": result,
                            },
                        }
                    ]

                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": content,
                })

            if tool_results:
                messages.append({"role": "user", "content": tool_results})

        else:
            log.warning(
                "ClaudeAgent: hit MAX_ITERATIONS (%d) for group '%s'.",
                MAX_ITERATIONS, group_name,
            )

        # Save action log for this session
        log_path = self._session_dir / "actions.jsonl"
        with open(log_path, "a", encoding="utf-8") as f:
            for entry in self._action_log:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        self._action_log.clear()
        log.info("ClaudeAgent: actions saved → %s", log_path)
