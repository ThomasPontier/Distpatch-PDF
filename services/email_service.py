"""Enhanced email service with Outlook account management (uses ConfigManager for templates)."""

import os
from typing import List, Optional, Dict, Any
from pathlib import Path
from .config_manager import get_config_manager

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
    _WIN32COM_AVAILABLE = True
except Exception as e:
    print(f"[EmailService][DEBUG] pywin32 not available: {e}")
    win32com = None  # type: ignore
    pythoncom = None  # type: ignore
    _WIN32COM_AVAILABLE = False


class EmailService:
    """Email service supporting Outlook with a single connected account at a time.

    DEBUG: Verbose logging is enabled to trace Outlook integration and sending flow.
    """
    
    def __init__(self):
        """Initialize the email service."""
        # Single-account state (no multi-account switching UI/logic)
        self.current_account_email: Optional[str] = None
        self.outlook = None
        self.outlook_user = None
        self.is_connected = False
        self.current_email_address = None
        # Track the user-selected sender for debug/traceability (not forced)
        self._selected_sender_email: Optional[str] = None
    
    def set_current_account(self, account_id: str, account_info: Dict[str, Any]):
        """
        Set the current email account (single account only).
        Kept for backward compatibility; stores only an email string.
        """
        try:
            print(f"[EmailService][DEBUG] set_current_account called: id={account_id}, info={account_info}")
            self.current_account_email = (account_info or {}).get('email', None)
            self.current_email_address = self.current_account_email
            self._selected_sender_email = self.current_account_email
            print(f"[EmailService][DEBUG] current_account_email set to: {self.current_account_email}")
            print(f"[EmailService][DEBUG] selected_sender_email set to: {self._selected_sender_email}")
        except Exception as e:
            print(f"[EmailService][DEBUG] Error parsing account_info: {e}")
            self.current_account_email = None
            self.current_email_address = None
            self._selected_sender_email = None
        self._reset_connection()
    
    def get_current_account(self) -> Optional[str]:
        """Get the current account identifier (email)."""
        return self.current_account_email
    
    def get_current_account_name(self) -> str:
        """Get the display name of the current account (email only)."""
        return self.current_email_address or 'No Account Connected'
    
    def get_current_email_address(self) -> Optional[str]:
        """Get the current email address."""
        return self.current_email_address
    
    def _reset_connection(self):
        """Reset connection state."""
        print("[EmailService][DEBUG] Resetting Outlook connection state")
        self.outlook = None
        self.outlook_user = None
        self.is_connected = False
    
    def connect_to_outlook(self) -> bool:
        """Connect to Outlook application and get user info."""
        print("[EmailService][DEBUG] Attempting to connect to Outlook...")
        if not _WIN32COM_AVAILABLE:
            print("[EmailService][DEBUG] pywin32 not installed or not importable. Install: pip install pywin32")
            self.outlook = None
            self.outlook_user = None
            # Do not override the user-selected sender; keep for debug
            self.is_connected = False
            return False
        try:
            # Ensure COM is initialized on this thread
            try:
                pythoncom.CoInitialize()
                print("[EmailService][DEBUG] pythoncom.CoInitialize() succeeded")
            except Exception as e_ci:
                print(f"[EmailService][DEBUG] pythoncom.CoInitialize() failed or already initialized: {e_ci}")
            # Try to connect to Outlook
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("[EmailService][DEBUG] Outlook.Application dispatch created")

            # Try to get current user info
            try:
                namespace = self.outlook.GetNamespace("MAPI")
                print("[EmailService][DEBUG] Got MAPI namespace")
                current_user = namespace.CurrentUser
                print(f"[EmailService][DEBUG] CurrentUser acquired: {current_user}")
                self.outlook_user = getattr(current_user, "Name", None)
                # Try to get email address actually used by Outlook default context
                detected_email = None
                try:
                    detected_email = getattr(current_user, "Address", None)
                except Exception as e_addr:
                    print(f"[EmailService][DEBUG] Failed to get current_user.Address: {e_addr}")
                self.current_email_address = detected_email
                print(f"[EmailService][DEBUG] outlook_user={self.outlook_user}, detected_email={self.current_email_address}")
                if self._selected_sender_email and self.current_email_address and (self._selected_sender_email.lower() != str(self.current_email_address).lower()):
                    print(f"[EmailService][DEBUG] NOTE: selected sender '{self._selected_sender_email}' differs from Outlook detected '{self.current_email_address}'. Will not force; Outlook will choose the actual sending account.")
                elif self._selected_sender_email and not self.current_email_address:
                    print(f"[EmailService][DEBUG] NOTE: Outlook did not provide a detected email. Selected sender is '{self._selected_sender_email}'.")
            except Exception as e_ns:
                # Fallback to generic connection
                print(f"[EmailService][DEBUG] Namespace/CurrentUser failed: {e_ns}")
                self.outlook_user = "Outlook User"
                # Do not override the user-selected sender
                if self._selected_sender_email:
                    print(f"[EmailService][DEBUG] Using selected sender (for debug only): {self._selected_sender_email}")
                else:
                    self.current_email_address = None
            
            self.is_connected = True
            print("[EmailService][DEBUG] Outlook connection established")
            return True
            
        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to connect to Outlook: {e}")
            self.outlook = None
            self.outlook_user = None
            # Keep selected sender for debug visibility
            self.is_connected = False
            return False
    
    def disconnect_from_outlook(self):
        """Disconnect from Outlook."""
        print("[EmailService][DEBUG] Disconnecting from Outlook")
        self.outlook = None
        self.outlook_user = None
        self.current_email_address = None
        self.is_connected = False
    
    def get_current_user(self) -> Optional[str]:
        """Get the current Outlook user's email address."""
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
            # Try to connect to Outlook without storing the connection
            win32com.client.Dispatch("Outlook.Application")
            # If we get here, Outlook is available
            print("[EmailService][DEBUG] Outlook appears available")
            return True
        except Exception as e:
            print(f"[EmailService][DEBUG] Outlook not available: {e}")
            return False
    
    def send_email(self, 
                   to_emails: List[str], 
                   subject: str, 
                   body: str, 
                   attachment_path: Optional[str] = None,
                   cc_emails: Optional[List[str]] = None,
                   bcc_emails: Optional[List[str]] = None) -> bool:
        """Send an email using Outlook. Does not force the sending account; logs the selected sender for debug."""
        print(f"[EmailService][DEBUG] send_email called")
        print(f"[EmailService][DEBUG]   to={to_emails}")
        print(f"[EmailService][DEBUG]   cc={cc_emails}")
        print(f"[EmailService][DEBUG]   bcc={bcc_emails}")
        print(f"[EmailService][DEBUG]   subject={subject}")
        print(f"[EmailService][DEBUG]   attachment={attachment_path}")
        print(f"[EmailService][DEBUG]   selected_sender(for debug)={self._selected_sender_email}")
        # Always use Outlook for sending emails
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
            
            # Do not force the sending account; only log context for testing
            try:
                effective_sender = None
                try:
                    # Some Outlook setups expose SenderEmailAddress after display
                    _ = mail  # placeholder to avoid linter warnings
                except Exception:
                    pass
                print(f"[EmailService][DEBUG] Effective sender (Outlook will decide). Selected (debug)={self._selected_sender_email}, detected_current_user={self.current_email_address}")
            except Exception as e_sender_info:
                print(f"[EmailService][DEBUG] Unable to query effective sender info: {e_sender_info}")
            
            mail.Send()
            print("[EmailService][DEBUG] Mail.Send() invoked successfully")
            return True
            
        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to send email via Outlook: {e}")
            return False
    
    def send_stopover_email(self, 
                          stopover_code: str, 
                          recipient_emails: List[str], 
                          pdf_path: str,
                          template_path: str = "config/email_template.txt") -> bool:
        """Send email for a specific stopover with PDF attachment.
        
        Note: template_path is ignored; subject/body now come from templates.json.
        """
        if not os.path.exists(pdf_path):
            print(f"PDF file not found: {pdf_path}")
            return False

        # Load consolidated templates (subject/body) from templates.json
        subject, body_template = self._load_templates_json()

        # Replace placeholders
        body = body_template.replace("{{stopover_code}}", stopover_code)
        subject = subject.replace("{{stopover_code}}", stopover_code)

        return self.send_email(recipient_emails, subject, body, pdf_path)
    
    def _load_template(self, template_path: str) -> Optional[str]:
        """Deprecated: legacy template loader retained for compatibility."""
        try:
            template_file = Path(template_path)
            if template_file.exists():
                return template_file.read_text(encoding='utf-8')
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
    
    def _load_templates_json(self) -> (str, str):
        """Load subject and body from centralized ConfigManager (JSON-backed)."""
        try:
            manager = get_config_manager()
            t = manager.get_templates()
            subject = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body = t.get("body") or self._get_default_template()
            print("[EmailService][DEBUG] Templates loaded")
            return subject, body
        except Exception as e:
            print(f"[EmailService][DEBUG] Failed to load templates from ConfigManager: {e}")
            return "Stopover Report - {{stopover_code}}", self._get_default_template()
