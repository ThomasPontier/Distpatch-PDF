"""Dialog for editing stopover-specific email configurations."""

import tkinter as tk
from tkinter import ttk, messagebox
import re
from typing import Optional, List
from services.stopover_email_service import StopoverEmailConfig, StopoverEmailService
from services.mapping_service import MappingService


class StopoverEmailDialog:
    """Dialog for editing stopover email configurations."""
    
    def __init__(self, parent, stopover_code: str, stopover_email_service: StopoverEmailService):
        """Initialize the stopover email dialog."""
        self.parent = parent
        self.stopover_code = stopover_code.upper()
        self.stopover_email_service = stopover_email_service
        self.mapping_service = MappingService()
        
        # Load existing configuration or create new one
        self.config = self.stopover_email_service.get_config(self.stopover_code)
        
        # Check if this is a new configuration (no recipients and default templates)
        self.is_new_config = (
            not self.config.recipients and 
            self.config.subject_template == "Stopover Report - {{stopover_code}}" and
            "Dear Team" in self.config.body_template and
            not self.config.cc_recipients and
            not self.config.bcc_recipients
        )
        
        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Email Configuration - {self.stopover_code}")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry("+%d+%d" % (
            parent.winfo_rootx() + 50,
            parent.winfo_rooty() + 50
        ))
        
        # Create UI components
        self._create_widgets()
        self._populate_fields()
    
    def _create_widgets(self):
        """Create all UI widgets."""
        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Removed per-stopover enable checkbox to make template/config global and always active
        self.enable_var = tk.BooleanVar(value=True)
        
        # Subject template
        subject_frame = ttk.LabelFrame(main_frame, text="Subject Template", padding="5")
        subject_frame.pack(fill="x", pady=(0, 10))
        
        self.subject_entry = ttk.Entry(subject_frame, width=50)
        self.subject_entry.pack(fill="x")
        
        # Body template
        body_frame = ttk.LabelFrame(main_frame, text="Email Body Template", padding="5")
        body_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.body_text = tk.Text(body_frame, height=10, width=50)
        body_scrollbar = ttk.Scrollbar(body_frame, orient="vertical", command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=body_scrollbar.set)
        
        self.body_text.pack(side="left", fill="both", expand=True)
        body_scrollbar.pack(side="right", fill="y")
        
        # Recipients section
        recipients_frame = ttk.LabelFrame(main_frame, text="Recipients", padding="5")
        recipients_frame.pack(fill="x", pady=(0, 10))
        
        # To recipients
        to_frame = ttk.Frame(recipients_frame)
        to_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(to_frame, text="To:").pack(anchor="w")
        self.to_text = tk.Text(to_frame, height=3, width=50)
        to_scrollbar = ttk.Scrollbar(to_frame, orient="vertical", command=self.to_text.yview)
        self.to_text.configure(yscrollcommand=to_scrollbar.set)
        self.to_text.pack(side="left", fill="x", expand=True, pady=(2, 0))
        to_scrollbar.pack(side="right", fill="y")
        
        # CC recipients
        cc_frame = ttk.Frame(recipients_frame)
        cc_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(cc_frame, text="CC:").pack(anchor="w")
        self.cc_text = tk.Text(cc_frame, height=2, width=50)
        cc_scrollbar = ttk.Scrollbar(cc_frame, orient="vertical", command=self.cc_text.yview)
        self.cc_text.configure(yscrollcommand=cc_scrollbar.set)
        self.cc_text.pack(side="left", fill="x", expand=True, pady=(2, 0))
        cc_scrollbar.pack(side="right", fill="y")
        
        # BCC recipients
        bcc_frame = ttk.Frame(recipients_frame)
        bcc_frame.pack(fill="x")
        ttk.Label(bcc_frame, text="BCC:").pack(anchor="w")
        self.bcc_text = tk.Text(bcc_frame, height=2, width=50)
        bcc_scrollbar = ttk.Scrollbar(bcc_frame, orient="vertical", command=self.bcc_text.yview)
        self.bcc_text.configure(yscrollcommand=bcc_scrollbar.set)
        self.bcc_text.pack(side="left", fill="x", expand=True, pady=(2, 0))
        bcc_scrollbar.pack(side="right", fill="y")
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Save button
        self.save_button = ttk.Button(
            button_frame,
            text="Save",
            command=self._save_config
        )
        self.save_button.pack(side="left", padx=(0, 5))
        
        # Cancel button
        self.cancel_button = ttk.Button(
            button_frame,
            text="Cancel",
            command=self._cancel
        )
        self.cancel_button.pack(side="left", padx=(0, 5))
        
        # Reset to default button
        self.reset_button = ttk.Button(
            button_frame,
            text="Reset to Default",
            command=self._reset_to_default
        )
        self.reset_button.pack(side="left")
        
        # Bind Enter key to save
        self.dialog.bind('<Return>', lambda e: self._save_config())
        self.dialog.bind('<Escape>', lambda e: self._cancel())
    
    def _populate_fields(self):
        """Populate fields with current configuration."""
        # Enable is always true in global template mode
        self.enable_var.set(True)
        
        # Subject template
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.insert(0, self.config.subject_template)
        
        # Body template
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", self.config.body_template)
        
        # Recipients
        self.to_text.delete("1.0", tk.END)
        if self.config.recipients:
            self.to_text.insert("1.0", "\n".join(self.config.recipients))
        else:
            # Always show the current email mapping for this stopover
            default_emails = self.mapping_service.get_emails_for_stopover(self.stopover_code)
            if default_emails:
                self.to_text.insert("1.0", "\n".join(default_emails))
                # Update the config with these default emails if it's a new config
                if self.is_new_config:
                    self.config.recipients = default_emails
            else:
                # Show a message indicating no emails are configured
                self.to_text.insert("1.0", f"No email configured for {self.stopover_code}")
        
        self.cc_text.delete("1.0", tk.END)
        if self.config.cc_recipients:
            self.cc_text.insert("1.0", "\n".join(self.config.cc_recipients))
        
        self.bcc_text.delete("1.0", tk.END)
        if self.config.bcc_recipients:
            self.bcc_text.insert("1.0", "\n".join(self.config.bcc_recipients))
    
    def _validate_email_list(self, email_text: str) -> List[str]:
        """Validate and parse email list from text."""
        emails = []
        lines = email_text.strip().split('\n')
        for line in lines:
            line = line.strip()
            if line:
                # Skip the "No email configured" placeholder message
                if line.startswith("No email configured for"):
                    continue
                # Split by comma or semicolon
                parts = re.split(r'[,;]', line)
                for part in parts:
                    email = part.strip()
                    if email:
                        if self._validate_email(email):
                            emails.append(email)
                        else:
                            raise ValueError(f"Invalid email address: {email}")
        return emails
    
    def _validate_email(self, email: str) -> bool:
        """Validate email address format."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def _save_config(self):
        """Save the configuration."""
        try:
            # Get values from fields
            is_enabled = self.enable_var.get()
            subject_template = self.subject_entry.get().strip()
            body_template = self.body_text.get("1.0", tk.END).strip()
            
            # Parse recipients
            to_recipients = self._validate_email_list(self.to_text.get("1.0", tk.END))
            cc_recipients = self._validate_email_list(self.cc_text.get("1.0", tk.END))
            bcc_recipients = self._validate_email_list(self.bcc_text.get("1.0", tk.END))
            
            # Validate required fields
            if not to_recipients:
                messagebox.showerror(
                    "Error", 
                    "At least one 'To' recipient is required."
                )
                return
            
            # Update config
            # is_enabled is always true in global mode
            self.config.is_enabled = True
            self.config.subject_template = subject_template
            self.config.body_template = body_template
            self.config.recipients = to_recipients
            self.config.cc_recipients = cc_recipients
            self.config.bcc_recipients = bcc_recipients
            
            # Save configuration
            self.stopover_email_service.save_config(self.config)
            
            # Close dialog
            self.dialog.destroy()
            
        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
    
    def _cancel(self):
        """Cancel and close dialog."""
        self.dialog.destroy()
    
    def _reset_to_default(self):
        """Reset configuration to default values."""
        result = messagebox.askyesno(
            "Reset to Default",
            f"Reset all email settings for {self.stopover_code} to default values?"
        )
        
        if result:
            # Create new default config
            self.config = StopoverEmailConfig(stopover_code=self.stopover_code)
            self._populate_fields()
    
    def show(self) -> Optional[StopoverEmailConfig]:
        """Show the dialog and return the configuration if saved."""
        self.dialog.wait_window()
        return self.config if self.stopover_email_service.config_exists(self.stopover_code) else None
