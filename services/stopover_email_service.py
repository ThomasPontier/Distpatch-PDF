"""Service for managing stopover-specific email configurations backed by ConfigManager."""

from dataclasses import asdict, dataclass
from typing import Any, Dict, List, Optional

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
    def from_dict(cls, data: Dict[str, Any]) -> "StopoverEmailConfig":
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
    """Facade over ConfigManager for stopover emails (mappings/stopovers/last_sent/templates).

    DRY contract:
    - Single source of truth remains ConfigManager.
    - Backward compatible with mappings: list[str] (legacy 'To' only).
    - Adds CC/BCC persistence by encoding them into mappings using a tagged shape:
        - If mappings[CODE] is a list[str], treat as 'To' only (legacy).
        - If mappings[CODE] is a dict, we expect keys: {"to": [...], "cc": [...], "bcc": [...]}.
      ConfigManager currently sanitizes mappings to lists; we therefore store cc/bcc in a compact
      sidecar convention inside the 'to' list using sentinel tags. This keeps a single field and
      avoids schema migration while staying fully reversible:
         - "__CC__:" + email value for each CC
         - "__BCC__:" + email value for each BCC
      On read, we split these back out; on write, we re-encode. Plain emails remain 'To'.
    """

    _CC_TAG = "__CC__:"
    _BCC_TAG = "__BCC__:"

    def __init__(self, config_dir: str = "config"):
        # config_dir retained for compatibility; not used for persistence anymore.
        self._manager = get_config_manager()

    @staticmethod
    def _encode_recipients(to_list, cc_list, bcc_list) -> list:
        encoded = []
        for e in to_list or []:
            s = str(e).strip()
            if s:
                encoded.append(s)
        for e in cc_list or []:
            s = str(e).strip()
            if s:
                encoded.append(f"{StopoverEmailService._CC_TAG}{s}")
        for e in bcc_list or []:
            s = str(e).strip()
            if s:
                encoded.append(f"{StopoverEmailService._BCC_TAG}{s}")
        return encoded

    @staticmethod
    def _decode_recipients(encoded_list) -> tuple[list, list, list]:
        to_list, cc_list, bcc_list = [], [], []
        for raw in encoded_list or []:
            s = str(raw).strip()
            if not s:
                continue
            if s.startswith(StopoverEmailService._CC_TAG):
                cc = s[len(StopoverEmailService._CC_TAG):].strip()
                if cc:
                    cc_list.append(cc)
            elif s.startswith(StopoverEmailService._BCC_TAG):
                bcc = s[len(StopoverEmailService._BCC_TAG):].strip()
                if bcc:
                    bcc_list.append(bcc)
            else:
                to_list.append(s)
        return to_list, cc_list, bcc_list

    def get_config(self, stopover_code: str) -> StopoverEmailConfig:
        """Build StopoverEmailConfig based on unified config state (with CC/BCC decode).

        Note: Return effective templates so the UI reflects the same subject/body that
        EmailService will actually use when sending (persisted value or default fallback)."""
        code = str(stopover_code).upper()
        # Raw templates snapshot (kept for future use if needed)
        _ = self._manager.get_templates()
        maps = self._manager.get_mappings()
        last = self._manager.get_last_sent()
        is_enabled = self._manager.is_stopover_enabled(code)

        encoded = list(maps.get(code, []))
        to_list, cc_list, bcc_list = self._decode_recipients(encoded)

        # Use effective templates (persisted-or-default) to match sending path
        subject_eff, body_eff = self._manager.get_effective_templates()

        return StopoverEmailConfig(
            stopover_code=code,
            subject_template=subject_eff,
            body_template=body_eff,
            recipients=to_list,
            cc_recipients=cc_list,
            bcc_recipients=bcc_list,
            is_enabled=is_enabled,
            last_sent_at=last.get(code),
        )

    def save_config(self, config: StopoverEmailConfig) -> bool:
        """Persist recipients/enablement/last_sent via ConfigManager (with CC/BCC encode)."""
        code = str(config.stopover_code).upper()
        if config.is_enabled:
            self._manager.add_stopover(code)
        else:
            self._manager.remove_stopover(code)

        encoded = self._encode_recipients(
            list(config.recipients or []),
            list(config.cc_recipients or []),
            list(config.bcc_recipients or []),
        )
        # Store back as a flat list; ConfigManager schema remains unchanged
        self._manager.set_mapping(code, encoded)

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
        maps = self._manager.get_mappings()
        if code in maps:
            self._manager.remove_mapping(code)
            changed = True
        stops = self._manager.get_stopovers()
        if code in stops:
            self._manager.remove_stopover(code)
            changed = True
        last = self._manager.get_last_sent()
        if code in last:
            self._manager.clear_last_sent_normalized(code)
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

    def set_last_sent_now(self, stopover_code: str) -> None:
        code = str(stopover_code).upper()
        self._manager.set_last_sent(code)

    def get_last_sent(self, stopover_code: str) -> Optional[str]:
        code = str(stopover_code).upper()
        return self._manager.get_last_sent().get(code)

    def _load_templates_json(self) -> "tuple[str, str]":
        """Load subject/body from unified ConfigManager templates."""
        t = self._manager.get_templates()
        return t.get("subject", "Stopover Report - {{stopover_code}}"), t.get("body", "")
