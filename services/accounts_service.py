from __future__ import annotations
import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, List, Optional


ACCOUNTS_PATH = Path("config") / "accounts.json"


@dataclass
class SenderAccount:
    id: str
    type: str  # "outlook" | "custom"
    name: str
    email: str


@dataclass
class AccountsModel:
    version: int
    senders: List[SenderAccount]
    selected_sender_id: Optional[str]


class AccountsService:
    """
    Robust manager for config/accounts.json with strict schema, validation, and migration.
    """

    def __init__(self, path: Path = ACCOUNTS_PATH) -> None:
        self.path = path
        self._data: AccountsModel = self._load_or_init()

    # ---------------- Public API ----------------

    def get_senders(self) -> List[SenderAccount]:
        return list(self._data.senders)

    def get_selected_sender(self) -> Optional[SenderAccount]:
        sid = self._data.selected_sender_id
        if not sid:
            return None
        return next((s for s in self._data.senders if s.id == sid), None)

    def set_selected_sender(self, sender_id: Optional[str]) -> None:
        if sender_id is not None and not any(s.id == sender_id for s in self._data.senders):
            raise ValueError(f"Unknown sender id: {sender_id}")
        self._data.selected_sender_id = sender_id
        self._save()

    def add_or_update_sender(self, sender: SenderAccount) -> None:
        self._validate_sender(sender)
        # update or insert
        existing_idx = next((i for i, s in enumerate(self._data.senders) if s.id == sender.id), None)
        if existing_idx is not None:
            self._data.senders[existing_idx] = sender
        else:
            self._data.senders.append(sender)
        # If no selection yet, select this one by default
        if not self._data.selected_sender_id:
            self._data.selected_sender_id = sender.id
        self._save()

    def remove_sender(self, sender_id: str) -> None:
        before = len(self._data.senders)
        self._data.senders = [s for s in self._data.senders if s.id != sender_id]
        if len(self._data.senders) == before:
            # nothing removed
            return
        if self._data.selected_sender_id == sender_id:
            self._data.selected_sender_id = None
        self._save()

    def ensure_custom_sender(self, email: str) -> SenderAccount:
        email = (email or "").strip()
        if not email:
            raise ValueError("email is empty")
        # Reuse existing by email if any
        for s in self._data.senders:
            if s.email.lower() == email.lower():
                return s
        # Otherwise create a new one
        new_id = self._next_id(prefix="custom_")
        acc = SenderAccount(id=new_id, type="custom", name=email, email=email)
        self.add_or_update_sender(acc)
        return acc

    def ensure_outlook_sender(self, display_name: str, email: str) -> SenderAccount:
        display_name = (display_name or "").strip() or "Outlook User"
        email = (email or "").strip()
        # Reuse existing by email if any
        for s in self._data.senders:
            if s.type == "outlook" and s.email.lower() == email.lower():
                # maybe update display name
                if s.name != display_name and display_name:
                    self.add_or_update_sender(SenderAccount(id=s.id, type=s.type, name=display_name, email=s.email))
                return s
        new_id = self._next_id(prefix="outlook_")
        acc = SenderAccount(id=new_id, type="outlook", name=display_name, email=email)
        self.add_or_update_sender(acc)
        return acc

    # ---------------- Internal ----------------

    def _default(self) -> AccountsModel:
        return AccountsModel(version=1, senders=[], selected_sender_id=None)

    def _save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        wire = {
            "version": self._data.version,
            "senders": [asdict(s) for s in self._data.senders],
            "selected_sender_id": self._data.selected_sender_id,
        }
        tmp = self.path.with_suffix(self.path.suffix + ".tmp")
        bak = self.path.with_suffix(self.path.suffix + ".bak")
        try:
            if self.path.exists():
                try:
                    bak.write_bytes(self.path.read_bytes())
                except Exception:
                    pass
            tmp.write_text(json.dumps(wire, indent=2, ensure_ascii=False), encoding="utf-8")
            os.replace(tmp, self.path)
        finally:
            try:
                if tmp.exists():
                    tmp.unlink()
            except Exception:
                pass

    def _load_or_init(self) -> AccountsModel:
        # Try new schema
        if self.path.exists():
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
                model = self._from_wire(data)
                if model:
                    return model
            except Exception:
                pass
        # Try migrate from legacy structure
        legacy = self._try_load_legacy()
        if legacy:
            self._data = legacy
            self._save()
            return legacy
        # default
        m = self._default()
        self._data = m
        self._save()
        return m

    def _from_wire(self, data: Any) -> Optional[AccountsModel]:
        try:
            if not isinstance(data, dict):
                return None
            version = int(data.get("version", 1))
            raw_senders = data.get("senders", [])
            selected_sender_id = data.get("selected_sender_id")
            senders: List[SenderAccount] = []
            if isinstance(raw_senders, list):
                for x in raw_senders:
                    if not isinstance(x, dict):
                        continue
                    sid = str(x.get("id", "")).strip()
                    stype = str(x.get("type", "")).strip().lower()
                    name = str(x.get("name", "")).strip()
                    email = str(x.get("email", "")).strip()
                    s = SenderAccount(id=sid, type=stype, name=name, email=email)
                    self._validate_sender(s)
                    senders.append(s)
            selected_id = str(selected_sender_id) if isinstance(selected_sender_id, (str, int)) else None
            # sanitize selection
            if selected_id and not any(s.id == selected_id for s in senders):
                selected_id = None
            # auto-select first if none
            if not selected_id and senders:
                selected_id = senders[0].id
            return AccountsModel(version=version, senders=senders, selected_sender_id=selected_id)
        except Exception:
            return None

    def _try_load_legacy(self) -> Optional[AccountsModel]:
        """
        Migrate from legacy:
          {"outlook_accounts": {id: {type,name,email}}, "custom_senders": [email,...], "selected_sender": "email"}
        """
        if not self.path.exists():
            return None
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            return None
        if not isinstance(data, dict):
            return None
        if "senders" in data:
            # already new schema
            return None
        outlook_accounts = data.get("outlook_accounts", {})
        custom_senders = data.get("custom_senders", [])
        selected_sender = data.get("selected_sender")
        senders: List[SenderAccount] = []
        # migrate outlook accounts
        if isinstance(outlook_accounts, dict):
            for k, v in outlook_accounts.items():
                if not isinstance(v, dict):
                    continue
                name = str(v.get("name", "")).strip() or "Outlook User"
                email = str(v.get("email", "")).strip()
                sid = str(k).strip() if k else self._next_id(prefix="outlook_")
                senders.append(SenderAccount(id=sid, type="outlook", name=name, email=email))
        # migrate custom senders
        if isinstance(custom_senders, list):
            for em in custom_senders:
                email = str(em).strip()
                if not email:
                    continue
                sid = self._next_id(prefix="custom_")
                senders.append(SenderAccount(id=sid, type="custom", name=email, email=email))
        # selection: choose the first matching by email
        selected_id = None
        if isinstance(selected_sender, str):
            for s in senders:
                if s.email.lower() == selected_sender.strip().lower():
                    selected_id = s.id
                    break
        if not selected_id and senders:
            selected_id = senders[0].id
        return AccountsModel(version=1, senders=senders, selected_sender_id=selected_id)

    def _validate_sender(self, s: SenderAccount) -> None:
        if not s.id or not isinstance(s.id, str):
            raise ValueError("sender.id must be non-empty string")
        if s.type not in ("outlook", "custom"):
            raise ValueError("sender.type must be 'outlook' or 'custom'")
        if not isinstance(s.name, str):
            raise ValueError("sender.name must be string")
        if not s.email or "@" not in s.email:
            # we allow missing @ for some Outlook on-prem cases, but keep basic check
            pass

    def _next_id(self, prefix: str) -> str:
        base = prefix
        i = 1
        existing = {s.id for s in self._data.senders} if hasattr(self, "_data") and self._data else set()
        while True:
            cand = f"{base}{i}"
            if cand not in existing:
                return cand
            i += 1
