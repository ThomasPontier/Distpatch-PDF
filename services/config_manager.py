import json
import os
import shutil
import threading
import sys
from typing import Any, Callable, ClassVar, Dict, List, Optional, Tuple
from datetime import datetime


def _ensure_dir(path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)


def _atomic_write_json(path: str, data: Dict[str, Any]) -> None:
    """
    Atomically write JSON content to disk with a backup and crash-safety.
    - Write to .tmp then replace
    - Also create a .bak backup of the previous file before replacing
    """
    _ensure_dir(path)
    tmp_path = f"{path}.tmp"
    bak_path = f"{path}.bak"
    try:
        # If an existing file is present, copy to .bak first
        if os.path.exists(path):
            try:
                shutil.copy2(path, bak_path)
            except Exception:
                pass
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        # Atomic replace on same filesystem
        os.replace(tmp_path, path)
    finally:
        # Best-effort cleanup
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except Exception:
            pass


class ConfigManager:
    """
    Centralized configuration manager (Option A: Single JSON + simple observer).
    - Single source of truth stored in JSON on disk.
    - Immediate persistence on each mutation (atomic writes).
    - Simple observer callbacks per domain and global.
    - Loads from unified file if present; otherwise migrates from legacy fragments.
    - Keeps schema clean: no deprecated keys (check_vars, global_subject, PDF paths).
    - Accounts are left untouched for now (migrate later).
    """

    _instance: ClassVar[Optional["ConfigManager"]] = None

    # Unified config path
    @staticmethod
    def _get_app_config_path() -> str:
        """Get the appropriate config path based on execution context."""
        if getattr(sys, 'frozen', False):
            # Running as compiled executable - use APPDATA
            appdata_dir = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'Distpatch-PDF')
            return os.path.join(appdata_dir, 'config', 'app_config.json')
        else:
            # Running as script - use local config directory
            return os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'app_config.json')
    
    APP_CONFIG_PATH: ClassVar[str] = _get_app_config_path.__func__()

    # Legacy paths to migrate from
    LEGACY_PATHS: ClassVar[Dict[str, List[str]]] = {
        "root": [
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "stopover_emails.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "mappings.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "templates.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "config.json"),
        ],
        "pkg": [
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "stopover_emails.json"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "mappings.json"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "templates.json"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "config.json"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "email_template.txt"),
        ],
    }

    def __init__(self) -> None:
        # in-memory state
        self._config: Dict[str, Any] = self._default_config()
        self._lock = threading.RLock()

        # observers
        self._obs_mappings: List[Callable[[Dict[str, List[str]]], None]] = []
        self._obs_stopovers: List[Callable[[List[str]], None]] = []
        self._obs_templates: List[Callable[[Dict[str, str]], None]] = []
        self._obs_last_sent: List[Callable[[Dict[str, str]], None]] = []
        self._obs_all: List[Callable[[Dict[str, Any]], None]] = []

        # load or migrate
        self._load_or_migrate()

    @classmethod
    def instance(cls) -> "ConfigManager":
        if cls._instance is None:
            cls._instance = ConfigManager()
        return cls._instance

    # Defaults
    # Centralized default subject template
    DEFAULT_SUBJECT: ClassVar[str] = "Rapport d’escale - {{stopover_code}}"

    def _default_config(self) -> Dict[str, Any]:
        return {
            "version": 1,
            "stopovers": [],
            "mappings": {},  # stopover: list[email]
            "templates": {
                "subject": "",
                "body": "",
            },
            "last_sent": {},  # stopover: ISO8601 str
        }

    # Persistence
    def _save(self) -> None:
        with self._lock:
            try:
                _atomic_write_json(self.APP_CONFIG_PATH, self._config)
            except Exception:
                # Avoid raising to keep UI responsive; config remains in memory
                pass

    def _load_or_migrate(self) -> None:
        # Try unified file first
        try:
            if os.path.exists(self.APP_CONFIG_PATH):
                with open(self.APP_CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    self._config = self._sanitize_loaded_config(data)
                    return
        except Exception:
            # Fallback to migration/defaults
            pass

        # Migrate from legacy fragments
        migrated = self._default_config()

        # 1) Stopover mappings and stopover list sources
        legacy_mapping_files = [
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "stopover_emails.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "stopover_emails.json"),
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "mappings.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "mappings.json"),
        ]
        mappings, stopovers = self._read_legacy_mappings(legacy_mapping_files)
        migrated["mappings"] = mappings
        migrated["stopovers"] = stopovers or list(mappings.keys())

        # 2) Templates (subject+body) from templates.json or email_template.txt
        legacy_template_json_candidates = [
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "templates.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "templates.json"),
        ]
        legacy_template_txt = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "email_template.txt")
        subject, body = self._read_legacy_templates(legacy_template_json_candidates, legacy_template_txt)
        migrated["templates"]["subject"] = subject or ""
        migrated["templates"]["body"] = body or ""

        # 3) last_sent if present in any legacy config/config.json (optional)
        legacy_config_candidates = [
            os.path.join(os.path.dirname(os.path.dirname(__file__)), "config", "config.json"),
            os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config", "config.json"),
        ]
        last_sent = self._read_legacy_last_sent(legacy_config_candidates)
        migrated["last_sent"] = last_sent

        # Save migrated unified config
        self._config = self._sanitize_loaded_config(migrated)
        self._save()
        # After migration, proactively remove legacy fragments to avoid future drift
        try:
            for group in self.LEGACY_PATHS.values():
                for p in group:
                    if os.path.exists(p):
                        try:
                            os.remove(p)
                        except Exception:
                            pass
        except Exception:
            pass

    def _sanitize_loaded_config(self, cfg: Dict[str, Any]) -> Dict[str, Any]:
        # Enforce schema and types, drop deprecated keys
        sanitized = self._default_config()
        sanitized["version"] = int(cfg.get("version", 1))

        # stopovers
        stopovers = cfg.get("stopovers", [])
        if isinstance(stopovers, list):
            sanitized["stopovers"] = [str(s) for s in stopovers if isinstance(s, (str, int))]

        # mappings
        mappings = cfg.get("mappings", {})
        if isinstance(mappings, dict):
            fixed_map: Dict[str, List[str]] = {}
            for k, v in mappings.items():
                if isinstance(k, str):
                    if isinstance(v, list):
                        emails = [str(x) for x in v if isinstance(x, (str, int))]
                    elif isinstance(v, str):
                        emails = [v]
                    else:
                        emails = []
                    fixed_map[k] = emails
            sanitized["mappings"] = fixed_map

        # templates
        templates = cfg.get("templates", {})
        subject = ""
        body = ""
        if isinstance(templates, dict):
            sub = templates.get("subject", "")
            bod = templates.get("body", "")
            subject = str(sub) if isinstance(sub, (str, int)) else ""
            body = str(bod) if isinstance(bod, (str, int)) else ""
        sanitized["templates"] = {"subject": subject, "body": body}

        # last_sent
        last_sent = cfg.get("last_sent", {})
        if isinstance(last_sent, dict):
            clean_last: Dict[str, str] = {}
            for k, v in last_sent.items():
                if isinstance(k, str) and isinstance(v, str):
                    # do a light validation for ISO format; if invalid, ignore
                    clean_last[k] = v
            sanitized["last_sent"] = clean_last

        return sanitized

    def _read_json_safely(self, path: str) -> Optional[Dict[str, Any]]:
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            return None
        return None

    def _read_text_safely(self, path: str) -> Optional[str]:
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    return f.read()
        except Exception:
            return None
        return None

    def _read_legacy_mappings(self, candidates: List[str]) -> Tuple[Dict[str, List[str]], List[str]]:
        mappings: Dict[str, List[str]] = {}
        stopovers: List[str] = []
        for path in candidates:
            data = self._read_json_safely(path)
            if not data:
                continue
            # try common shapes:
            # { "ABJ": ["a@b.com"] }  or { "mappings": {...}, "stopovers": [...] }
            if isinstance(data, dict):
                if all(isinstance(v, list) for v in data.values()):
                    # flat mapping
                    for k, v in data.items():
                        if isinstance(k, str) and isinstance(v, list):
                            mappings[k] = [str(x) for x in v if isinstance(x, (str, int))]
                else:
                    if "mappings" in data and isinstance(data["mappings"], dict):
                        for k, v in data["mappings"].items():
                            if isinstance(k, str):
                                if isinstance(v, list):
                                    emails = [str(x) for x in v if isinstance(x, (str, int))]
                                elif isinstance(v, str):
                                    emails = [v]
                                else:
                                    emails = []
                                mappings[k] = emails
                    if "stopovers" in data and isinstance(data["stopovers"], list):
                        stopovers = [str(s) for s in data["stopovers"] if isinstance(s, (str, int))]
        return mappings, stopovers

    def _read_legacy_templates(self, json_candidates: List[str], txt_candidate: str) -> Tuple[Optional[str], Optional[str]]:
        subject: Optional[str] = None
        body: Optional[str] = None

        # Prefer JSON if present
        for path in json_candidates:
            data = self._read_json_safely(path)
            if data and isinstance(data, dict):
                subj = data.get("subject")
                bod = data.get("body")
                if isinstance(subj, (str, int)):
                    subject = str(subj)
                if isinstance(bod, (str, int)):
                    body = str(bod)
                if subject is not None or body is not None:
                    break

        # Fallback to email_template.txt for body if JSON body missing
        if body is None:
            text = self._read_text_safely(txt_candidate)
            if text:
                body = text

        if subject is None:
            subject = ""

        if body is None:
            body = ""

        return subject, body

    def _read_legacy_last_sent(self, candidates: List[str]) -> Dict[str, str]:
        last_sent: Dict[str, str] = {}
        for path in candidates:
            data = self._read_json_safely(path)
            if not data or not isinstance(data, dict):
                continue
            # Look for possible last_sent locations
            if "last_sent" in data and isinstance(data["last_sent"], dict):
                for k, v in data["last_sent"].items():
                    if isinstance(k, str) and isinstance(v, str):
                        last_sent[k] = v
        return last_sent

    # Observer registration
    def on_mappings_changed(self, cb: Callable[[Dict[str, List[str]]], None]) -> None:
        self._obs_mappings.append(cb)

    def on_stopovers_changed(self, cb: Callable[[List[str]], None]) -> None:
        self._obs_stopovers.append(cb)

    def on_templates_changed(self, cb: Callable[[Dict[str, str]], None]) -> None:
        self._obs_templates.append(cb)

    def on_last_sent_changed(self, cb: Callable[[Dict[str, str]], None]) -> None:
        self._obs_last_sent.append(cb)

    def on_config_changed(self, cb: Callable[[Dict[str, Any]], None]) -> None:
        self._obs_all.append(cb)

    # Notify helpers
    def _emit_mappings(self) -> None:
        for cb in list(self._obs_mappings):
            try:
                cb(self.get_mappings())
            except Exception:
                pass
        self._emit_all()

    def _emit_stopovers(self) -> None:
        for cb in list(self._obs_stopovers):
            try:
                cb(self.get_stopovers())
            except Exception:
                pass
        self._emit_all()

    def _emit_templates(self) -> None:
        for cb in list(self._obs_templates):
            try:
                cb(self.get_templates())
            except Exception:
                pass
        self._emit_all()

    def _emit_last_sent(self) -> None:
        for cb in list(self._obs_last_sent):
            try:
                cb(self.get_last_sent())
            except Exception:
                pass
        self._emit_all()

    def _emit_all(self) -> None:
        snapshot = self.get_all()
        for cb in list(self._obs_all):
            try:
                cb(snapshot)
            except Exception:
                pass

    # Public getters
    def get_all(self) -> Dict[str, Any]:
        with self._lock:
            # deep copy via json round-trip for immutability outside
            return json.loads(json.dumps(self._config))

    def get_stopovers(self) -> List[str]:
        with self._lock:
            return list(self._config.get("stopovers", []))

    def get_mappings(self) -> Dict[str, List[str]]:
        with self._lock:
            # deep copy
            return {k: list(v) for k, v in self._config.get("mappings", {}).items()}

    def get_templates(self) -> Dict[str, str]:
        with self._lock:
            t = self._config.get("templates", {}) or {}
            return {"subject": t.get("subject", ""), "body": t.get("body", "")}

    def get_last_sent(self) -> Dict[str, str]:
        with self._lock:
            return dict(self._config.get("last_sent", {}))

    # Public setters (persist and emit)
    def set_stopovers(self, stopovers: List[str]) -> None:
        with self._lock:
            # normalize input
            desired = [str(s).upper() for s in stopovers]
            self._config["stopovers"] = desired
            # prune mappings/last_sent for non-present stopovers
            maps = self._config.get("mappings", {})
            self._config["mappings"] = {k.upper(): list(v) for k, v in maps.items() if str(k).upper() in desired}
            last = self._config.get("last_sent", {})
            self._config["last_sent"] = {k.upper(): v for k, v in last.items() if str(k).upper() in desired}
            self._save()
        self._emit_stopovers()
        self._emit_mappings()
        self._emit_last_sent()

    def add_stopover(self, stopover: str) -> None:
        with self._lock:
            s = str(stopover)
            lst = self._config.get("stopovers", [])
            if s not in lst:
                lst.append(s)
                self._config["stopovers"] = lst
                self._save()
        self._emit_stopovers()

    def remove_stopover(self, stopover: str) -> None:
        with self._lock:
            s = str(stopover)
            lst = self._config.get("stopovers", [])
            if s in lst:
                lst.remove(s)
                self._config["stopovers"] = lst
            # also clean mapping and last_sent
            if s in self._config.get("mappings", {}):
                self._config["mappings"].pop(s, None)
            if s in self._config.get("last_sent", {}):
                self._config["last_sent"].pop(s, None)
            self._save()
        self._emit_stopovers()
        self._emit_mappings()
        self._emit_last_sent()

    def set_mapping(self, stopover: str, emails: List[str]) -> None:
        with self._lock:
            s = str(stopover).upper()
            ems = [str(e) for e in emails]
            maps = self._config.get("mappings", {})
            maps[s] = ems
            self._config["mappings"] = maps
            # ensure stopover list contains it
            lst = self._config.get("stopovers", [])
            if s not in lst:
                lst.append(s)
                self._config["stopovers"] = lst
            self._save()
        self._emit_mappings()
        self._emit_stopovers()

    def remove_mapping(self, stopover: str) -> None:
        with self._lock:
            s = str(stopover)
            if s in self._config.get("mappings", {}):
                self._config["mappings"].pop(s, None)
                self._save()
        self._emit_mappings()

    def set_subject(self, subject: str) -> None:
        with self._lock:
            t = self._config.get("templates", {})
            t["subject"] = str(subject)
            self._config["templates"] = t
            self._save()
        self._emit_templates()

    def set_body(self, body: str) -> None:
        with self._lock:
            t = self._config.get("templates", {})
            t["body"] = str(body)
            self._config["templates"] = t
            self._save()
        self._emit_templates()

    def set_templates(self, subject: str, body: str) -> None:
        with self._lock:
            self._config["templates"] = {"subject": str(subject), "body": str(body)}
            self._save()
        self._emit_templates()

    def set_last_sent(self, stopover: str, iso_ts: Optional[str] = None) -> None:
        with self._lock:
            s = str(stopover).upper()
            ts = iso_ts if isinstance(iso_ts, str) and iso_ts else datetime.utcnow().isoformat() + "Z"
            last = self._config.get("last_sent", {})
            last[s] = ts
            self._config["last_sent"] = last
            self._save()
        self._emit_last_sent()

    def clear_last_sent(self, stopover: str) -> None:
        with self._lock:
            s = str(stopover)
            if s in self._config.get("last_sent", {}):
                self._config["last_sent"].pop(s, None)
                self._save()
        self._emit_last_sent()

    # --------- DRY helpers (no behavior change for existing APIs) ---------

    def get_effective_templates(self) -> tuple[str, str]:
        """
        Return (subject, body) with defaults applied:
          - subject: templates["subject"] or DEFAULT_SUBJECT
          - body: templates["body"] or the literal default body template currently used by EmailService._get_default_template()
        This is a pure accessor; no changes to persisted config.
        """
        t = self.get_templates()
        subject = t.get("subject") or self.DEFAULT_SUBJECT
        # Copy the literal default body from EmailService._get_default_template() to avoid cross-dependency
        default_body = """Bonjour,

Veuillez trouver en pièce jointe le rapport d’escale pour {{stopover_code}}.

Cordialement,
Distpatch PDF"""
        body = t.get("body") or default_body
        return subject, body

    def is_stopover_enabled(self, code: str) -> bool:
        """Return True if normalized code is present in stopovers list."""
        if code is None:
            return False
        cu = str(code).upper()
        return cu in self.get_stopovers()

    def clear_last_sent_normalized(self, code: str) -> None:
        """Uppercase code internally then clear last_sent for that key."""
        if code is None:
            return
        cu = str(code).upper()
        # Direct manipulation preserving semantics of clear_last_sent
        with self._lock:
            if cu in self._config.get("last_sent", {}):
                self._config["last_sent"].pop(cu, None)
                self._save()
        self._emit_last_sent()


# Convenience module-level accessor
def get_config_manager() -> ConfigManager:
    return ConfigManager.instance()


# One-shot normalization helper to replace config content with a provided dict
def replace_all_config(stopovers: List[str], mappings: Dict[str, List[str]], last_sent: Dict[str, str]) -> None:
    """
    Replace stopovers, mappings, and last_sent atomically and notify observers.
    - Keys are normalized to upper-case for stopovers.
    - Unknown mapping keys are pruned to the stopovers list.
    """
    mgr = get_config_manager()
    # Normalize input
    normalized_stopovers = [str(s).upper() for s in (stopovers or [])]
    normalized_mappings: Dict[str, List[str]] = {}
    for k, v in (mappings or {}).items():
        ku = str(k).upper()
        if ku in normalized_stopovers:
            normalized_mappings[ku] = [str(e) for e in (v or [])]
    normalized_last: Dict[str, str] = {}
    for k, v in (last_sent or {}).items():
        ku = str(k).upper()
        if ku in normalized_stopovers and isinstance(v, str):
            normalized_last[ku] = v

    with mgr._lock:
        mgr._config["stopovers"] = normalized_stopovers
        mgr._config["mappings"] = normalized_mappings
        mgr._config["last_sent"] = normalized_last
        mgr._save()
    # Emit all domains
    mgr._emit_stopovers()
    mgr._emit_mappings()
    mgr._emit_last_sent()
