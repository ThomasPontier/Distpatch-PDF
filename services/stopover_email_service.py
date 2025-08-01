"""Service for managing stopover-specific email configurations backed by ConfigManager."""

from typing import Dict, List, Optional, Any
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from .config_manager import get_config_manager


@dataclass
class StopoverEmailConfig:
    """Configuration for stopover-specific email settings."""
    
    stopover_code: str
    subject_template: str = "Stopover Report - {{stopover_code}}"
    body_template: str = """Dear Team,

Please find attached the stopover report for {{stopover_code}}.

This report contains all relevant information for this stopover location.

Best regards,
PDF Stopover Analyzer"""
    recipients: List[str] = None
    cc_recipients: List[str] = None
    bcc_recipients: List[str] = None
    is_enabled: bool = True
    last_sent_at: Optional[str] = None  # ISO 8601 UTC timestamp e.g. "2025-07-31T10:22:45Z"
    
    def __post_init__(self):
        """Initialize default lists if None."""
        if self.recipients is None:
            self.recipients = []
        if self.cc_recipients is None:
            self.cc_recipients = []
        if self.bcc_recipients is None:
            self.bcc_recipients = []
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'StopoverEmailConfig':
        """Create StopoverEmailConfig from dictionary (backward compatible)."""
        # Fill missing fields with defaults
        data = dict(data or {})
        data.setdefault("recipients", [])
        data.setdefault("cc_recipients", [])
        data.setdefault("bcc_recipients", [])
        data.setdefault("is_enabled", True)
        data.setdefault("last_sent_at", None)
        return cls(**data)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert StopoverEmailConfig to dictionary."""
        return asdict(self)


class StopoverEmailService:
    """Facade over ConfigManager for stopover emails (mappings/stopovers/last_sent/templates)."""

    def __init__(self, config_dir: str = "config"):
        # config_dir retained for compatibility; not used for persistence anymore.
        self._manager = get_config_manager()

    def get_config(self, stopover_code: str) -> StopoverEmailConfig:
        """Build StopoverEmailConfig based on unified config state."""
        code = str(stopover_code).upper()
        t = self._manager.get_templates()
        maps = self._manager.get_mappings()
        last = self._manager.get_last_sent()
        stopovers = self._manager.get_stopovers()
        return StopoverEmailConfig(
            stopover_code=code,
            subject_template=t.get("subject", "Stopover Report - {{stopover_code}}"),
            body_template=t.get("body", ""),
            recipients=list(maps.get(code, [])),
            cc_recipients=[],
            bcc_recipients=[],
            is_enabled=code in stopovers,
            last_sent_at=last.get(code),
        )

    def save_config(self, config: StopoverEmailConfig) -> bool:
        """Persist recipients/enablement/last_sent via ConfigManager."""
        code = str(config.stopover_code).upper()
        # enable/disable
        if config.is_enabled:
            self._manager.add_stopover(code)
        else:
            self._manager.remove_stopover(code)
        # recipients
        self._manager.set_mapping(code, list(config.recipients or []))
        # last sent
        if config.last_sent_at:
            self._manager.set_last_sent(code, config.last_sent_at)
        return True

    def get_all_configs(self) -> Dict[str, StopoverEmailConfig]:
        """Return all known stopovers from union of stopovers and mappings keys."""
        stopovers = set(self._manager.get_stopovers())
        stopovers.update(self._manager.get_mappings().keys())
        result: Dict[str, StopoverEmailConfig] = {}
        for code in sorted(stopovers):
            result[code] = self.get_config(code)
        return result

    def delete_config(self, stopover_code: str) -> bool:
        code = str(stopover_code).upper()
        changed = False
        # remove mapping
        maps = self._manager.get_mappings()
        if code in maps:
            self._manager.remove_mapping(code)
            changed = True
        # remove from stopovers
        stops = self._manager.get_stopovers()
        if code in stops:
            self._manager.remove_stopover(code)
            changed = True
        # clear last sent
        last = self._manager.get_last_sent()
        if code in last:
            self._manager.clear_last_sent(code)
            changed = True
        return changed

    def config_exists(self, stopover_code: str) -> bool:
        code = str(stopover_code).upper()
        return code in self._manager.get_stopovers() or code in self._manager.get_mappings()

    def get_enabled_configs(self) -> Dict[str, StopoverEmailConfig]:
        enabled: Dict[str, StopoverEmailConfig] = {}
        for code in self._manager.get_stopovers():
            enabled[code] = self.get_config(code)
        return enabled

    # New: last sent date tracking
    def set_last_sent_now(self, stopover_code: str) -> None:
        code = str(stopover_code).upper()
        self._manager.set_last_sent(code)

    def get_last_sent(self, stopover_code: str) -> Optional[str]:
        code = str(stopover_code).upper()
        return self._manager.get_last_sent().get(code)

    # Templates loader now defers to ConfigManager
    def _load_templates_json(self) -> "tuple[str, str]":
        t = self._manager.get_templates()
        return t.get("subject", "Stopover Report - {{stopover_code}}"), t.get("body", "")
