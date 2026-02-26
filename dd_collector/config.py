"""Load and validate config.yaml into typed dataclasses."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml


# ── Dataclasses ──────────────────────────────────────────────

@dataclass
class DingTalkConfig:
    window_class: str = "StandardFrame_DingTalk"
    download_dir: str = ""
    timeout: int = 10
    download_wait: int = 5
    download_icon_offset: int = 8
    # Full path to DingTalk.exe — used for auto-launch when not running.
    # Leave empty to disable auto-launch (requires manual start).
    exe_path: str = ""


@dataclass
class GroupConfig:
    name: str = ""
    alias: str = ""

    def __post_init__(self):
        if not self.alias:
            self.alias = self.name


@dataclass
class GDriveConfig:
    base_path: str = ""


@dataclass
class PollingConfig:
    interval_minutes: int = 30
    max_downloads_per_group: int = 50
    # Number of scroll-up iterations when scanning a group chat for attachments
    chat_scroll_pages: int = 3


@dataclass
class DedupConfig:
    path: str = "data/downloaded.json"


@dataclass
class LoggingConfig:
    dir: str = "logs"
    level: str = "INFO"
    max_bytes: int = 10_485_760
    backup_count: int = 5


@dataclass
class SelectorConfig:
    """A single UI selector entry."""
    control_type: str = ""
    name: str = ""
    fallback_control_type: str = ""
    class_name: str = ""  # ClassName for controls identified by class rather than name


@dataclass
class UISelectorsConfig:
    search_box: SelectorConfig = field(default_factory=SelectorConfig)
    search_result_item: SelectorConfig = field(default_factory=SelectorConfig)
    files_tab: SelectorConfig = field(default_factory=SelectorConfig)
    file_list: SelectorConfig = field(default_factory=SelectorConfig)
    file_item: SelectorConfig = field(default_factory=SelectorConfig)
    download_button: SelectorConfig = field(default_factory=SelectorConfig)
    context_menu_download: SelectorConfig = field(default_factory=SelectorConfig)
    dismiss_buttons: List[str] = field(default_factory=list)


@dataclass
class VLMConfig:
    """Vision Language Model settings for UI element detection."""
    api_key: str = ""
    model: str = "qwen/qwen3.5-35b-a3b"
    base_url: str = "https://openrouter.ai/api/v1"
    # Pixels around the right-click point to capture for context menu detection
    capture_margin: int = 300


@dataclass
class ClaudeConfig:
    """Claude vision scanner settings.

    Only the scanner_model is used for routine scans (Haiku — fast and cheap).
    A more capable model can be set here for debugging or edge cases.
    """
    oauth_token: str = ""
    # claude-haiku-4-5-20251001 is ~20x cheaper than Sonnet for this task
    model: str = "claude-haiku-4-5-20251001"
    # Number of scroll-up passes per group when scanning for file cards
    max_scrolls: int = 5


@dataclass
class TriggerConfig:
    """Settings for --trigger watch-and-fire mode."""
    check_interval_seconds: int = 60   # how often to poll for unread signals
    red_pixel_threshold: int = 50      # min red sidebar pixels → badge present
    cooldown_seconds: int = 10         # wait after download before resuming watch


@dataclass
class AppConfig:
    dingtalk: DingTalkConfig = field(default_factory=DingTalkConfig)
    groups: List[GroupConfig] = field(default_factory=list)
    gdrive: GDriveConfig = field(default_factory=GDriveConfig)
    polling: PollingConfig = field(default_factory=PollingConfig)
    dedup: DedupConfig = field(default_factory=DedupConfig)
    logging: LoggingConfig = field(default_factory=LoggingConfig)
    ui_selectors: UISelectorsConfig = field(default_factory=UISelectorsConfig)
    vlm: VLMConfig = field(default_factory=VLMConfig)
    claude: ClaudeConfig = field(default_factory=ClaudeConfig)
    trigger: TriggerConfig = field(default_factory=TriggerConfig)


# ── Loader ───────────────────────────────────────────────────

def _build_selector(data: Any) -> SelectorConfig:
    if not isinstance(data, dict):
        return SelectorConfig()
    return SelectorConfig(
        control_type=data.get("control_type", ""),
        name=data.get("name", ""),
        fallback_control_type=data.get("fallback_control_type", ""),
        class_name=data.get("class_name", ""),
    )


def _build_ui_selectors(data: Dict[str, Any]) -> UISelectorsConfig:
    sel = UISelectorsConfig()
    sel.search_box = _build_selector(data.get("search_box"))
    sel.search_result_item = _build_selector(data.get("search_result_item"))
    sel.files_tab = _build_selector(data.get("files_tab"))
    sel.file_list = _build_selector(data.get("file_list"))
    sel.file_item = _build_selector(data.get("file_item"))
    sel.download_button = _build_selector(data.get("download_button"))
    sel.context_menu_download = _build_selector(data.get("context_menu_download"))
    sel.dismiss_buttons = data.get("dismiss_buttons", [])
    return sel


def load_config(path: str = "config.yaml") -> AppConfig:
    """Load config.yaml and return a validated AppConfig."""
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path.resolve()}")

    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    cfg = AppConfig()

    # DingTalk
    dt = raw.get("dingtalk", {})
    cfg.dingtalk = DingTalkConfig(
        window_class=dt.get("window_class", cfg.dingtalk.window_class),
        download_dir=dt.get("download_dir", cfg.dingtalk.download_dir),
        timeout=int(dt.get("timeout", cfg.dingtalk.timeout)),
        download_wait=int(dt.get("download_wait", cfg.dingtalk.download_wait)),
        download_icon_offset=int(dt.get("download_icon_offset", cfg.dingtalk.download_icon_offset)),
        exe_path=dt.get("exe_path", cfg.dingtalk.exe_path),
    )

    # Groups
    for g in raw.get("groups", []):
        cfg.groups.append(GroupConfig(
            name=g.get("name", ""),
            alias=g.get("alias", ""),
        ))

    # Google Drive
    gd = raw.get("gdrive", {})
    cfg.gdrive = GDriveConfig(base_path=gd.get("base_path", ""))

    # Polling
    p = raw.get("polling", {})
    cfg.polling = PollingConfig(
        interval_minutes=int(p.get("interval_minutes", cfg.polling.interval_minutes)),
        max_downloads_per_group=int(p.get("max_downloads_per_group", cfg.polling.max_downloads_per_group)),
        chat_scroll_pages=int(p.get("chat_scroll_pages", cfg.polling.chat_scroll_pages)),
    )

    # Dedup
    d = raw.get("dedup", {})
    cfg.dedup = DedupConfig(path=d.get("path", cfg.dedup.path))

    # Logging
    lg = raw.get("logging", {})
    cfg.logging = LoggingConfig(
        dir=lg.get("dir", cfg.logging.dir),
        level=lg.get("level", cfg.logging.level),
        max_bytes=int(lg.get("max_bytes", cfg.logging.max_bytes)),
        backup_count=int(lg.get("backup_count", cfg.logging.backup_count)),
    )

    # UI Selectors
    cfg.ui_selectors = _build_ui_selectors(raw.get("ui_selectors", {}))

    # VLM
    v = raw.get("vlm", {})
    cfg.vlm = VLMConfig(
        api_key=v.get("api_key", os.environ.get("OPENROUTER_API_KEY", "")),
        model=v.get("model", cfg.vlm.model),
        base_url=v.get("base_url", cfg.vlm.base_url),
        capture_margin=int(v.get("capture_margin", cfg.vlm.capture_margin)),
    )

    # Claude computer-use agent
    cl = raw.get("claude", {})
    cfg.claude = ClaudeConfig(
        oauth_token=cl.get("oauth_token", os.environ.get("ANTHROPIC_AUTH_TOKEN", "")),
        model=cl.get("model", cfg.claude.model),
        max_scrolls=int(cl.get("max_scrolls", cfg.claude.max_scrolls)),
    )

    # Trigger watcher
    tr = raw.get("trigger", {})
    cfg.trigger = TriggerConfig(
        check_interval_seconds=int(tr.get("check_interval_seconds", cfg.trigger.check_interval_seconds)),
        red_pixel_threshold=int(tr.get("red_pixel_threshold", cfg.trigger.red_pixel_threshold)),
        cooldown_seconds=int(tr.get("cooldown_seconds", cfg.trigger.cooldown_seconds)),
    )

    return cfg
