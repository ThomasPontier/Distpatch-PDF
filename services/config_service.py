"""Configuration service for managing application settings (backward-compat shim around ConfigManager)."""

from pathlib import Path
from typing import Any, Dict

from .config_manager import get_config_manager


class ConfigService:
    """Backward compatibility facade.

    This class now delegates persisted configuration responsibilities to the
    centralized ConfigManager. Legacy methods return empty/defaults or no-ops
    to avoid persisting deprecated state.
    """
    def __init__(self, config_dir: str = "config"):
        # Keep minimal compatibility for callers expecting a path
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(exist_ok=True)
        self.default_config: Dict[str, Any] = {}
        self._manager = get_config_manager()
    
    def _ensure_config_file(self):
        """Deprecated: no file creation required (centralized in ConfigManager)."""
        return
    
    def load_config(self) -> Dict[str, Any]:
        """Return unified config snapshot from ConfigManager."""
        return self._manager.get_all()
    
    def save_config(self, config: Dict[str, Any]):
        """Save by writing through ConfigManager domain setters to keep integrity.
        
        Note: To avoid partial writes and schema drift, callers should prefer
        specific setters on services using ConfigManager directly. This method
        attempts to apply common fields if present.
        """
        applied = False
        if "stopovers" in config:
            try:
                self._manager.set_stopovers(list(config.get("stopovers") or []))
                applied = True
            except Exception:
                pass
        if "mappings" in config:
            try:
                maps = config.get("mappings") or {}
                if isinstance(maps, dict):
                    for k, v in maps.items():
                        emails = v if isinstance(v, list) else [v] if isinstance(v, str) else []
                        self._manager.set_mapping(str(k), [str(e) for e in emails])
                    applied = True
            except Exception:
                pass
        if "templates" in config:
            try:
                t = config.get("templates") or {}
                subject = t.get("subject", "")
                body = t.get("body", "")
                self._manager.set_templates(str(subject), str(body))
                applied = True
            except Exception:
                pass
        if "last_sent" in config:
            try:
                last = config.get("last_sent") or {}
                if isinstance(last, dict):
                    for k, v in last.items():
                        if isinstance(v, str):
                            self._manager.set_last_sent(str(k), v)
                    applied = True
            except Exception:
                pass
        if not applied:
            # no-op for deprecated/unrecognized sections
            return
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value from unified snapshot."""
        try:
            snap = self._manager.get_all()
            return snap.get(key, default)
        except Exception:
            return default
    
    def set(self, key: str, value: Any):
        """Set a configuration value by key (limited support).
        
        Supports keys: 'templates', 'mappings', 'stopovers', 'last_sent'.
        Other keys are ignored to avoid schema drift.
        """
        try:
            if key == "templates" and isinstance(value, dict):
                subject = value.get("subject", "")
                body = value.get("body", "")
                self._manager.set_templates(str(subject), str(body))
            elif key == "mappings" and isinstance(value, dict):
                for k, v in value.items():
                    emails = v if isinstance(v, list) else [v] if isinstance(v, str) else []
                    self._manager.set_mapping(str(k), [str(e) for e in emails])
            elif key == "stopovers" and isinstance(value, list):
                self._manager.set_stopovers([str(s) for s in value])
            elif key == "last_sent" and isinstance(value, dict):
                for k, v in value.items():
                    if isinstance(v, str):
                        self._manager.set_last_sent(str(k), v)
            else:
                return
        except Exception:
            return
    
    def get_window_config(self) -> Dict[str, int]:
        """Deprecated: window configuration is no longer persisted."""
        return {}
    
    def get_pdf_config(self) -> Dict[str, int]:
        """Deprecated: pdf configuration is no longer persisted."""
        return {}
    
    def get_email_config(self) -> Dict[str, str]:
        """Deprecated: template path is removed; use ConfigManager templates."""
        t = self._manager.get_templates()
        return {"subject": t.get("subject", ""), "body": t.get("body", "")}
    
    def get_ui_config(self) -> Dict[str, int]:
        """Deprecated: ui configuration is no longer persisted."""
        return {}
    
    def reset_to_defaults(self):
        """Reset to defaults through ConfigManager."""
        self._manager.set_stopovers([])
        # clear mappings
        maps = list(self._manager.get_mappings().keys())
        for m in maps:
            self._manager.remove_mapping(m)
        # clear templates
        self._manager.set_templates("", "")
        # clear last_sent
        last = list(self._manager.get_last_sent().keys())
        for s in last:
            self._manager.clear_last_sent(s)
    
    def _merge_configs(self, default: Dict[str, Any], user: Dict[str, Any]) -> Dict[str, Any]:
        """Legacy helper (unused)."""
        result = default.copy()
        result.update(user or {})
        return result
    
    def get_config_path(self) -> str:
        """Return unified app_config.json path from ConfigManager."""
        # Derive path via snapshot source; we know where manager stores it.
        from .config_manager import ConfigManager as _CM
        return _CM.APP_CONFIG_PATH
    
    def config_exists(self) -> bool:
        """Check if unified configuration exists."""
        from .config_manager import ConfigManager as _CM
        return Path(_CM.APP_CONFIG_PATH).exists()
