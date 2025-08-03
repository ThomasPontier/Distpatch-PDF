"""Enhanced email service with Outlook account management (uses ConfigManager for templates)."""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from .config_manager import get_config_manager

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    _WIN32COM_AVAILABLE = True
except Exception as e:
    # Keep: platform-specific import handling for Windows Outlook integration
    print(f"[EmailService][DEBUG] pywin32 not available: {e}")
    win32com = None  # type: ignore
    pythoncom = None  # type: ignore
    _WIN32COM_AVAILABLE = False


class EmailService:
    """Email service supporting Outlook with multi-account transparency and best-effort selection.

    DEBUG: Verbose logging is enabled to trace Outlook integration and sending flow.
    """

    def __init__(self):
        """Initialize the email service."""
        # Legacy-selected sender (informational; kept for compatibility with AccountsService)
        self.current_account_email: Optional[str] = None
        # Outlook COM objects/state
        self.outlook = None
        self.outlook_user = None
        self.is_connected = False
        self.current_email_address = None
        # Informational selected sender (from AccountsService)
        self._preferred_outlook_account_id: Optional[str] = None
        # Multi-account: preferred Outlook account id (session-scoped)
        self._preferred_account_id: Optional[str] = None
        # Cached enumeration of Outlook accounts
        self._accounts_cache: List[Dict[str, Optional[str]]] = []
        # Last send context for UI transparency
        self._last_send_context: Dict[str, Any] = {}

    # ---------- Compatibility with existing UI (kept but tightened) ----------
    def set_current_account(self, account_id: str, account_info: Dict[str, Any]):
        """
        Store an informational selected sender (email). Does not enforce sending account.
        Resets Outlook connection so that any change is reflected on next action.
        """
        try:
            print(f"[EmailService][DEBUG] set_current_account called: id={account_id}, info={account_info}")
            self.current_account_email = (account_info or {}).get("email", None)
            self.current_email_address = self.current_account_email
            self._selected_sender_email = self.current_account_email
            print(f"[EmailService][DEBUG] selected_sender_email set to: {self._selected_sender_email}")
        except Exception as e:
            print(f"[EmailService][DEBUG] Error parsing account_info: {e}")
            self.current_account_email = None
            self.current_email_address = None
            self._selected_sender_email = None
        self._reset_connection()

    def get_current_account(self) -> Optional[str]:
        """Return the informational selected account email (not enforced)."""
        return self.current_account_email

    def get_current_account_name(self) -> str:
        """Display name for status bar when no Outlook info is available."""
        return self.current_email_address or "No Account Connected"

    def get_current_email_address(self) -> Optional[str]:
        return self.current_email_address

    # ---------- Core connection ----------
    def _reset_connection(self):
        print("[EmailService][DEBUG] Resetting Outlook connection state")
        self.outlook = None
        self.outlook_user = None
        self.is_connected = False
        self.current_email_address = None
        self._accounts_cache = []

    def connect_to_outlook(self) -> bool:
        """Connect to Outlook application and get current user info."""
        print("[EmailService][DEBUG] Attempting to connect to Outlook...")
        if not _WIN32COM_AVAILABLE:
            print("[EmailService][DEBUG] pywin32 not installed or not importable. Install: pip install pywin32")
            self._reset_connection()
            return False
        try:
            try:
                pythoncom.CoInitialize()
                print("[EmailService][DEBUG] pythoncom.CoInitialize() succeeded")
            except Exception as e_ci:
                print(f"[EmailService][DEBUG] pythoncom.CoInitialize() failed or already initialized: {e_ci}")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("[EmailService][DEBUG] Outlook.Application dispatch created")

            try:
                namespace = self.outlook.GetNamespace("MAPI")
                print("[EmailService][DEBUG] Got MAPI namespace")
                current_user = namespace.CurrentUser
                print(f"[EmailService][DEBUG] CurrentUser acquired: {current_user}")
                self.outlook_user = getattr(current_user, "Name", None)
                detected_email = None
                try:
                    detected_email = getattr(current_user, "Address", None)
                except Exception as e_addr:
                    print(f"[EmailService][DEBUG] Failed to get current_user.Address: {e_addr}")
                self.current_email_address = detected_email
                print(f"[EmailService][DEBUG] outlook_user={self.outlook_user}, detected_email={self.current_email_address}")
            except Exception as e_ns:
                print(f"[EmailService][DEBUG] Namespace/CurrentUser failed: {e_ns}")
                self.outlook_user = "Outlook User"
                self.current_email_address = None

            # Refresh accounts cache after connection
            try:
                self._accounts_cache = self._enumerate_outlook_accounts_internal()
                print(f"[EmailService][DEBUG] Enumerated {len(self._accounts_cache)} Outlook account(s)")
            except Exception as e_list:
                print(f"[EmailService][DEBUG] Failed to enumerate Outlook accounts: {e_list}")
                self._accounts_cache = []

            self.is_connected = True
            print("[EmailService][DEBUG] Outlook connection established")
            return True

        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to connect to Outlook: {e}")
            self._reset_connection()
            return False

    def disconnect_from_outlook(self):
        print("[EmailService][DEBUG] Disconnecting from Outlook")
        self._reset_connection()

    def get_current_user(self) -> Optional[str]:
        """Get the current Outlook user's display name."""
        if not self.is_connected:
            print("[EmailService][DEBUG] get_current_user requires connection; attempting connect")
            if not self.connect_to_outlook():
                return None
        print(f"[EmailService][DEBUG] get_current_user -> {self.outlook_user}")
        return self.outlook_user

    def is_outlook_available(self) -> bool:
        """Check if Outlook is available and running."""
        try:
            if not _WIN32COM_AVAILABLE:
                print("[EmailService][DEBUG] pywin32 not available -> Outlook not available")
                return False
            win32com.client.Dispatch("Outlook.Application")
            print("[EmailService][DEBUG] Outlook appears available")
            return True
        except Exception as e:
            print(f"[EmailService][DEBUG] Outlook not available: {e}")
            return False

    # ---------- Transparency helpers ----------
    def get_effective_sender(self) -> Dict[str, Optional[str]]:
        """Return the best-known effective sender from Outlook context."""
        return {"name": self.outlook_user, "email": self.current_email_address}

    def get_last_send_context(self) -> Dict[str, Any]:
        """Small info dict for UI transparency after send."""
        return dict(self._last_send_context or {})

    # ---------- Accounts enumeration and selection ----------
    def list_outlook_accounts(self) -> List[Dict[str, Optional[str]]]:
        """
        Return cached list of Outlook accounts as dicts:
          { "id": str, "display_name": str, "smtp_address": Optional[str] }
        """
        if not self.is_connected:
            if not self.connect_to_outlook():
                return []
        if not self._accounts_cache:
            try:
                self._accounts_cache = self._enumerate_outlook_accounts_internal()
            except Exception:
                self._accounts_cache = []
        return list(self._accounts_cache)

    def set_preferred_outlook_account(self, account_id: Optional[str]) -> None:
        """Set the preferred Outlook account id to use for sending (best effort)."""
        self._preferred_account_id = account_id

    def get_preferred_outlook_account(self) -> Optional[str]:
        return self._preferred_account_id

    def _enumerate_outlook_accounts_internal(self) -> List[Dict[str, Optional[str]]]:
        """Internal: enumerate Outlook.Session.Accounts best effort."""
        result: List[Dict[str, Optional[str]]] = []
        try:
            if not self.outlook:
                return result
            session = self.outlook.Session
            accounts = getattr(session, "Accounts", None)
            if not accounts:
                return result
            # Accounts is 1-based in Outlook
            count = int(getattr(accounts, "Count", 0) or 0)
            for i in range(1, count + 1):
                try:
                    acc = accounts.Item(i)
                    display_name = getattr(acc, "DisplayName", None) or f"Outlook Account {i}"
                    smtp = None
                    try:
                        smtp = getattr(acc, "SmtpAddress", None)
                    except Exception:
                        smtp = None
                    # Some Exchange accounts do not expose SmtpAddress; leave None
                    acc_id = None
                    try:
                        acc_id = getattr(acc, "EntryID", None)
                    except Exception:
                        acc_id = None
                    if not acc_id:
                        acc_id = f"acc:{i}"
                    result.append(
                        {"id": str(acc_id), "display_name": str(display_name), "smtp_address": str(smtp) if smtp else None}
                    )
                except Exception as e_item:
                    print(f"[EmailService][DEBUG] Failed to read account {i}: {e_item}")
                    continue
        except Exception as e_all:
            print(f"[EmailService][DEBUG] Accounts enumeration error: {e_all}")
        return result

    def _resolve_account_by_id(self, account_id: str):
        """Return the COM Account object matching the cached id if possible."""
        try:
            if not self.outlook or not account_id:
                return None
            session = self.outlook.Session
            accounts = getattr(session, "Accounts", None)
            if not accounts:
                return None
            count = int(getattr(accounts, "Count", 0) or 0)
            # Try by EntryID match first
            for i in range(1, count + 1):
                acc = accounts.Item(i)
                try:
                    entry_id = getattr(acc, "EntryID", None)
                except Exception:
                    entry_id = None
                if entry_id and str(entry_id) == str(account_id):
                    return acc
            # Fallback: if account_id is acc:{i}
            if str(account_id).startswith("acc:"):
                try:
                    idx = int(str(account_id).split(":")[1])
                    if 1 <= idx <= count:
                        return accounts.Item(idx)
                except Exception:
                    return None
        except Exception:
            return None
        return None

    # ---------- Sending ----------
    def send_email(
        self,
        to_emails: List[str],
        subject: str,
        body: str,
        attachment_path: Optional[str] = None,
        cc_emails: Optional[List[str]] = None,
        bcc_emails: Optional[List[str]] = None,
    ) -> bool:
        """Send an email using Outlook. Best-effort apply preferred Outlook account if set."""
        print(f"[EmailService][DEBUG] send_email called")
        print(f"[EmailService][DEBUG]   to={to_emails}")
        print(f"[EmailService][DEBUG]   cc={cc_emails}")
        print(f"[EmailService][DEBUG]   bcc={bcc_emails}")
        print(f"[EmailService][DEBUG]   subject={subject}")
        print(f"[EmailService][DEBUG]   attachment={attachment_path}")
        print(f"[EmailService][DEBUG]   preferred_outlook_account_id={self._preferred_account_id}")

        self._last_send_context = {}

        if not self.is_connected:
            print("[EmailService][DEBUG] Not connected; attempting to connect...")
            if not self.connect_to_outlook():
                print("[EmailService][DEBUG] Connection failed; aborting send")
                return False
        if not _WIN32COM_AVAILABLE:
            print("[EmailService][DEBUG] Cannot send: pywin32 not available")
            return False

        try:
            if not self.outlook:
                print("[EmailService][DEBUG] Outlook object is None")
                return False
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            print("[EmailService][DEBUG] Mail item created")
            mail.To = "; ".join(to_emails or [])
            if cc_emails:
                mail.CC = "; ".join(cc_emails)
            if bcc_emails:
                mail.BCC = "; ".join(bcc_emails)
            mail.Subject = subject or ""
            mail.Body = body or ""

            if attachment_path:
                if os.path.exists(attachment_path):
                    mail.Attachments.Add(attachment_path)
                    print(f"[EmailService][DEBUG] Attachment added: {attachment_path}")
                else:
                    print(f"[EmailService][DEBUG] Attachment path does not exist: {attachment_path}")

            # Best-effort: set SendUsingAccount if preferred account is set and resolvable
            attempted_account_id = None
            applied_account = False
            if self._preferred_account_id:
                attempted_account_id = self._preferred_account_id
                try:
                    acc_obj = self._resolve_account_by_id(self._preferred_account_id)
                    if acc_obj is not None:
                        try:
                            mail.SendUsingAccount = acc_obj
                            applied_account = True
                            print("[EmailService][DEBUG] SendUsingAccount applied")
                        except Exception as e_set:
                            print(f"[EmailService][DEBUG] Failed to set SendUsingAccount: {e_set}")
                    else:
                        print("[EmailService][DEBUG] Preferred account id could not be resolved; using Outlook default")
                except Exception as e_res:
                    print(f"[EmailService][DEBUG] Error resolving preferred account: {e_res}")

            # Snapshot effective context for transparency
            self._last_send_context = {
                "attempted_account_id": attempted_account_id,
                "applied_preferred": applied_account,
                "effective_sender": {"name": self.outlook_user, "email": self.current_email_address},
            }

            mail.Send()
            print("[EmailService][DEBUG] Mail.Send() invoked successfully")
            return True

        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to send email via Outlook: {e}")
            return False

    # ---------- Stopover helper and templates ----------
    def send_stopover_email(
        self,
        stopover_code: str,
        recipient_emails: List[str],
        pdf_path: str,
        template_path: str = "config/email_template.txt",
    ) -> bool:
        """Send email for a specific stopover with PDF attachment."""
        if not os.path.exists(pdf_path):
            print(f"PDF file not found: {pdf_path}")
            return False

        subject, body_template = self._load_templates_json()
        body = body_template.replace("{{stopover_code}}", stopover_code)
        subject = subject.replace("{{stopover_code}}", stopover_code)

        return self.send_email(recipient_emails, subject, body, pdf_path)

    def _load_template(self, template_path: str) -> Optional[str]:
        """Deprecated: legacy template loader retained for compatibility."""
        try:
            template_file = Path(template_path)
            if template_file.exists():
                return template_file.read_text(encoding="utf-8")
        except Exception as e:
            print(f"Failed to load template: {e}")
        return None

    def _get_default_template(self) -> str:
        """Get default email template."""
        return """Dear Team,

Please find attached the stopover report for {{stopover_code}}.

This report contains all relevant information for this stopover location.

Best regards,
PDF Stopover Analyzer"""

    def _load_templates_json(self) -> tuple[str, str]:
        """Load subject and body from centralized ConfigManager (JSON-backed)."""
        try:
            manager = get_config_manager()
            subject, body = manager.get_effective_templates()
            print("[EmailService][DEBUG] Templates loaded")
            return subject, body
        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to load templates from ConfigManager: {e}")
            return "Stopover Report - {{stopover_code}}", self._get_default_template()
