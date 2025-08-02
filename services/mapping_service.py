"""Service for managing stopover-to-email mappings via ConfigManager (unified source of truth)."""

from typing import Dict, List
from .config_manager import get_config_manager


class MappingService:
    """
    Facade over ConfigManager for mappings.
    This eliminates divergence between config/mappings.json and the unified app_config.json.
    """

    def __init__(self, config_dir: str = "config"):
        # config_dir retained for compatibility but unused
        self._manager = get_config_manager()

    def get_emails_for_stopover(self, stopover_code: str) -> List[str]:
        code = str(stopover_code).upper()
        return list(self._manager.get_mappings().get(code, []))

    def add_mapping(self, stopover_code: str, email: str) -> bool:
        code = str(stopover_code).upper()
        email = str(email).strip()
        maps = self._manager.get_mappings()
        current = list(maps.get(code, []))
        if email and email not in current:
            current.append(email)
            self._manager.set_mapping(code, current)
            self._manager.add_stopover(code)  # ensure enabled
            return True
        return False

    def remove_mapping(self, stopover_code: str, email: str) -> bool:
        code = str(stopover_code).upper()
        email = str(email).strip()
        maps = self._manager.get_mappings()
        if code in maps and email in (maps.get(code) or []):
            new_list = [e for e in maps.get(code, []) if e != email]
            if new_list:
                self._manager.set_mapping(code, new_list)
            else:
                # remove entire mapping if empty
                self._manager.remove_mapping(code)
            return True
        return False

    def get_all_mappings(self) -> Dict[str, List[str]]:
        return self._manager.get_mappings()

    def get_mapped_stopovers(self) -> List[str]:
        return sorted(list(self._manager.get_mappings().keys()))

    def has_mapping(self, stopover_code: str) -> bool:
        code = str(stopover_code).upper()
        maps = self._manager.get_mappings()
        return code in maps and len(maps.get(code) or []) > 0

    def update_mappings(self, new_mappings: Dict[str, List[str]]):
        # Normalize and persist each mapping via ConfigManager setters
        for code, emails in (new_mappings or {}).items():
            c = str(code).upper()
            dedup = []
            for e in emails or []:
                s = str(e).strip()
                if s and s not in dedup:
                    dedup.append(s)
            if dedup:
                self._manager.set_mapping(c, dedup)
                self._manager.add_stopover(c)
            else:
                self._manager.remove_mapping(c)
