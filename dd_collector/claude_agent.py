"""Autonomous DingTalk file-download agent using Claude computer use API."""

from __future__ import annotations

import base64
import io
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

import anthropic
import pyautogui

log = logging.getLogger("dd_collector")

# ── Computer use tool definition ─────────────────────────────

SCREEN_WIDTH = 1280
SCREEN_HEIGHT = 800

COMPUTER_TOOL: Dict[str, Any] = {
    "type": "computer_20250124",
    "name": "computer",
    "display_width_px": SCREEN_WIDTH,
    "display_height_px": SCREEN_HEIGHT,
}

# Safety cap: agent will stop after this many API round-trips
MAX_ITERATIONS = 60


# ── Agent ─────────────────────────────────────────────────────

class ClaudeAgent:
    """Autonomous agent using Claude's computer use API.

    Uses OAuth token (``sk-ant-oat01-...``) via the ``auth_token`` parameter.
    """

    def __init__(
        self,
        oauth_token: str,
        model: str = "claude-sonnet-4-5",
    ) -> None:
        # OAuth tokens (sk-ant-oat01-…) are Claude Code CLI tokens and are NOT
        # accepted by the Anthropic REST API. Regular API keys (sk-ant-api03-…)
        # from console.anthropic.com are required for direct API access.
        if oauth_token.startswith("sk-ant-oat"):
            raise ValueError(
                "The provided token is a Claude Code OAuth token (sk-ant-oat01-…), "
                "which is NOT accepted by the Anthropic REST API.\n"
                "Please obtain a regular API key (sk-ant-api03-…) from "
                "https://console.anthropic.com and set it as claude.oauth_token "
                "in config.yaml (or export ANTHROPIC_AUTH_TOKEN)."
            )
        # Regular API keys (sk-ant-api03-…) use api_key parameter
        self._client = anthropic.Anthropic(api_key=oauth_token)
        self._model = model

    # ── Screenshot / action helpers ──────────────────────────

    def _take_screenshot(self) -> str:
        """Take a full-screen screenshot and return base64-encoded PNG."""
        screenshot = pyautogui.screenshot()
        buf = io.BytesIO()
        screenshot.save(buf, format="PNG")
        return base64.standard_b64encode(buf.getvalue()).decode("utf-8")

    def _execute_action(self, action: Dict[str, Any]) -> Optional[str]:
        """Execute a single computer-use action.

        Returns:
            Base64 PNG string for 'screenshot'/'cursor_position' actions,
            None for everything else (caller takes a new screenshot).
        """
        act = action.get("action", "")

        if act == "screenshot":
            return self._take_screenshot()

        if act in ("left_click", "right_click", "double_click", "mouse_move"):
            coord = action.get("coordinate", [0, 0])
            x, y = int(coord[0]), int(coord[1])
            if act == "left_click":
                pyautogui.click(x, y, button="left")
            elif act == "right_click":
                pyautogui.click(x, y, button="right")
            elif act == "double_click":
                pyautogui.doubleClick(x, y)
            elif act == "mouse_move":
                pyautogui.moveTo(x, y)
            time.sleep(0.4)
            return None

        if act == "left_click_drag":
            start = action.get("start_coordinate", [0, 0])
            end = action.get("coordinate", [0, 0])
            pyautogui.mouseDown(int(start[0]), int(start[1]))
            time.sleep(0.1)
            pyautogui.mouseUp(int(end[0]), int(end[1]))
            time.sleep(0.4)
            return None

        if act == "type":
            text = action.get("text", "")
            # Use pyperclip + paste for non-ASCII (Chinese filenames etc.)
            try:
                import pyperclip
                pyperclip.copy(text)
                pyautogui.hotkey("ctrl", "v")
            except Exception:
                pyautogui.typewrite(text, interval=0.05)
            time.sleep(0.3)
            return None

        if act == "key":
            key_str = action.get("text", "")
            # Normalize key names: "Return" → "enter", "ctrl+c" → hotkey
            key_str = key_str.strip()
            parts = [k.strip().lower() for k in key_str.split("+")]
            # Map common names
            key_map = {
                "return": "enter", "super": "win", "ctrl": "ctrl",
                "alt": "alt", "shift": "shift", "escape": "escape",
                "esc": "escape", "backspace": "backspace", "tab": "tab",
                "delete": "delete", "del": "delete",
                "page_up": "pageup", "page_down": "pagedown",
                "home": "home", "end": "end",
                "up": "up", "down": "down", "left": "left", "right": "right",
            }
            parts = [key_map.get(p, p) for p in parts]
            if len(parts) == 1:
                pyautogui.press(parts[0])
            else:
                pyautogui.hotkey(*parts)
            time.sleep(0.3)
            return None

        if act == "scroll":
            coord = action.get("coordinate", [640, 400])
            x, y = int(coord[0]), int(coord[1])
            direction = action.get("direction", "down")
            amount = int(action.get("amount", 3))
            if direction == "down":
                pyautogui.scroll(-amount, x=x, y=y)
            elif direction == "up":
                pyautogui.scroll(amount, x=x, y=y)
            elif direction == "right":
                pyautogui.hscroll(amount, x=x, y=y)
            elif direction == "left":
                pyautogui.hscroll(-amount, x=x, y=y)
            time.sleep(0.4)
            return None

        if act == "cursor_position":
            x, y = pyautogui.position()
            return f"Cursor at ({x}, {y})"

        if act == "hold_key":
            key_str = action.get("text", "")
            duration = float(action.get("duration", 0.5))
            pyautogui.keyDown(key_str)
            time.sleep(duration)
            pyautogui.keyUp(key_str)
            time.sleep(0.2)
            return None

        log.warning("Unknown computer-use action: %s", act)
        return None

    # ── Main agent loop ───────────────────────────────────────

    def run_download_task(
        self,
        group_name: str,
        download_dir: str,
        already_downloaded: List[str],
        max_scrolls: int = 5,
    ) -> None:
        """Run the autonomous download agent for one DingTalk group.

        Args:
            group_name:         Display name of the DingTalk group.
            download_dir:       Where DingTalk saves downloaded files.
            already_downloaded: Filenames already downloaded (to skip).
            max_scrolls:        Hint for how many screens to scroll.
        """
        already_str = (
            "\n".join(f"  - {f}" for f in already_downloaded[:50])
            if already_downloaded
            else "  (none)"
        )

        task_prompt = (
            f"You are automating the DingTalk (钉钉) PC application on Windows "
            f"to download files from a group chat.\n\n"
            f"TASK: Download all NEW files from DingTalk group \"{group_name}\".\n\n"
            f"FILES ALREADY DOWNLOADED (skip these):\n{already_str}\n\n"
            f"STEP-BY-STEP INSTRUCTIONS:\n"
            f"1. Take a screenshot to see the current screen state.\n"
            f"2. Find the DingTalk window (class 'StandardFrame_DingTalk'). "
            f"   If minimized, click its taskbar button to restore it.\n"
            f"3. Navigate to the group \"{group_name}\" in DingTalk:\n"
            f"   a. Look in the left chat list sidebar for the group name.\n"
            f"   b. If visible, click directly on it to open the group chat.\n"
            f"   c. If NOT visible: look for the search/input box at the TOP of the "
            f"      sidebar (NOT the main search bar that opens a full overlay). "
            f"      Type the group name there. When results appear, click the matching group.\n"
            f"   d. If a full-screen 'Search or Ask' overlay appears, press Escape to "
            f"      dismiss it, then try clicking a group directly in the sidebar.\n"
            f"4. Once the group chat is open, scan for file attachment cards in the chat:\n"
            f"   - File cards show: file icon, filename, file size, and a Download button.\n"
            f"   - Hover over each card to reveal the Download button if needed.\n"
            f"5. For each file card NOT in the already-downloaded list above:\n"
            f"   a. Click the Download (下载) button on the file card.\n"
            f"   b. If a save/confirm dialog appears, accept it.\n"
            f"   c. Wait 1-2 seconds after clicking.\n"
            f"6. Scroll UP in the chat to see older messages and find more files. "
            f"   Scroll through approximately {max_scrolls} screens of history.\n"
            f"7. Repeat steps 4-6 until all visible file cards have been processed.\n"
            f"8. When done, say exactly: TASK COMPLETE\n\n"
            f"IMPORTANT NOTES:\n"
            f"- DingTalk file cards may say '下载' (Chinese) instead of 'Download'.\n"
            f"- Do NOT click 'Open File' or 'Open Folder' after a download starts.\n"
            f"- Do NOT try to navigate to the Files tab — only scan the chat messages.\n"
            f"- Download directory: {download_dir}\n"
            f"- If you see a pop-up update/notification dialog, dismiss it first "
            f"  by clicking 'Later', '稍后', 'Cancel', '取消', or pressing Escape.\n"
            f"- Start immediately by taking a screenshot."
        )

        # Initial message with screenshot already attached
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

        log.info(
            "Starting Claude computer-use agent for group: %s (model=%s)",
            group_name, self._model,
        )

        for iteration in range(MAX_ITERATIONS):
            log.info("Agent iteration %d / %d", iteration + 1, MAX_ITERATIONS)

            try:
                response = self._client.beta.messages.create(
                    model=self._model,
                    max_tokens=4096,
                    tools=[COMPUTER_TOOL],
                    messages=messages,
                    betas=["computer-use-2025-01-24"],
                )
            except anthropic.APIError as exc:
                log.error("Claude API error on iteration %d: %s", iteration + 1, exc)
                break

            log.debug("stop_reason=%s  content_blocks=%d", response.stop_reason, len(response.content))

            # Add assistant turn to conversation
            messages.append({"role": "assistant", "content": response.content})

            # Check if Claude declared it's done
            if response.stop_reason == "end_turn":
                last_text = ""
                for block in response.content:
                    if hasattr(block, "text"):
                        last_text += block.text
                if "TASK COMPLETE" in last_text.upper():
                    log.info("Agent signalled TASK COMPLETE for group: %s", group_name)
                else:
                    log.info("Agent finished (end_turn) for group: %s", group_name)
                break

            if response.stop_reason != "tool_use":
                log.warning(
                    "Unexpected stop_reason '%s' — stopping agent.", response.stop_reason
                )
                break

            # Execute all tool-use blocks and collect results
            tool_results: List[Dict[str, Any]] = []
            for block in response.content:
                if not hasattr(block, "type") or block.type != "tool_use":
                    continue
                if block.name != "computer":
                    log.warning("Unexpected tool name: %s", block.name)
                    continue

                action = block.input
                log.info("  action=%s  args=%s", action.get("action"), {
                    k: v for k, v in action.items() if k != "action"
                })

                result = self._execute_action(action)

                # Always send a screenshot back unless we already got one
                if result is None or not isinstance(result, str) or result.startswith("Cursor"):
                    # Take a new screenshot after the action
                    time.sleep(0.5)
                    screenshot_b64 = self._take_screenshot()
                    tool_result_content: Any = [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": "image/png",
                                "data": screenshot_b64,
                            },
                        }
                    ]
                    if result is not None and result.startswith("Cursor"):
                        # Prepend cursor position as text
                        tool_result_content.insert(0, {"type": "text", "text": result})
                else:
                    # result IS the screenshot
                    tool_result_content = [
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
                    "content": tool_result_content,
                })

            if tool_results:
                messages.append({"role": "user", "content": tool_results})

        else:
            log.warning(
                "Agent hit MAX_ITERATIONS (%d) for group: %s", MAX_ITERATIONS, group_name
            )

        log.info("Agent loop ended for group: %s", group_name)
