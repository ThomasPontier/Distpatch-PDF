"""Outlook account management dialog.

Extended to allow selecting the sender email, adding custom sender emails if needed,
and persisting everything in config/accounts.json via the existing mechanisms.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from typing import List, Dict, Optional
from pathlib import Path
import win32com.client


class AccountDialog:
    """Dialog for managing Outlook email accounts and selecting the sender."""
    
    def __init__(self, parent, email_service):
        """Initialize the account dialog."""
        self.parent = parent
        self.email_service = email_service
        print("[AccountDialog][DEBUG] init dialog")
        
        # Load account configurations
        self.accounts = self._load_accounts()
        print(f"[AccountDialog][DEBUG] loaded accounts: {self.accounts}")
        # Keep track of selected sender (email string) from the accounts file if present
        try:
            # If service has a notion of current account, keep it; else prefer accounts.json selected sender
            self.selected_account = self.email_service.get_current_account()
        except Exception:
            self.selected_account = None
        print(f"[AccountDialog][DEBUG] initial selected_account from service: {self.selected_account}")
        # If service has none, try accounts.json selected_sender
        if not self.selected_account:
            try:
                sel = self.accounts.get("selected_sender")
                if sel:
                    print(f"[AccountDialog][DEBUG] fallback selected_sender from accounts.json: {sel}")
                    self.email_service.set_current_account("custom", {"type": "custom", "name": sel, "email": sel})
                    self.selected_account = sel
            except Exception as e:
                print(f"[AccountDialog][DEBUG] failed to fallback selected sender: {e}")
        
        # Create dialog window
        self._create_dialog()
        self._setup_ui()
        self._populate_accounts()
    
    def _create_dialog(self):
        """Create the dialog window."""
        self.window = tk.Toplevel(self.parent)
        self.window.title("Outlook Account Management")
        self.window.geometry("400x300")
        self.window.minsize(350, 250)
        
        # Make dialog modal
        self.window.transient(self.parent)
        self.window.grab_set()
        
        # Center dialog
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def _setup_ui(self):
        """Setup the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Account selection frame
        selection_frame = ttk.LabelFrame(main_frame, text="Available Outlook Accounts", padding="5")
        selection_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Account list
        self.account_listbox = tk.Listbox(selection_frame, height=6)
        self.account_listbox.pack(fill="both", expand=True, pady=(0, 5))
        
        # Buttons frame
        buttons_frame = ttk.Frame(selection_frame)
        buttons_frame.pack(fill="x")
        
        # Account management buttons
        self.add_button = ttk.Button(
            buttons_frame,
            text="Connect to Outlook",
            command=self._connect_to_outlook
        )
        self.add_button.pack(side="left", padx=(0, 5))
        
        self.remove_button = ttk.Button(
            buttons_frame,
            text="Disconnect",
            command=self._disconnect_account,
            state="disabled"
        )
        self.remove_button.pack(side="left")
        
        # Custom sender frame
        custom_frame = ttk.LabelFrame(main_frame, text="Custom Sender Emails", padding="5")
        custom_frame.pack(fill="x", pady=(0, 10))
        entry_row = ttk.Frame(custom_frame)
        entry_row.pack(fill="x", pady=(0, 5))
        ttk.Label(entry_row, text="Add sender email:").pack(side="left")
        self.custom_email_entry = ttk.Entry(entry_row)
        self.custom_email_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        self.add_custom_btn = ttk.Button(entry_row, text="Add", command=self._add_custom_sender)
        self.add_custom_btn.pack(side="left")
        self.custom_listbox = tk.Listbox(custom_frame, height=4)
        self.custom_listbox.pack(fill="both", expand=True)
        rm_row = ttk.Frame(custom_frame)
        rm_row.pack(fill="x", pady=(5, 0))
        self.remove_custom_btn = ttk.Button(rm_row, text="Remove Selected", command=self._remove_custom_sender)
        self.remove_custom_btn.pack(side="right")
        
        # Bind selection event
        self.account_listbox.bind('<<ListboxSelect>>', self._on_account_select)
        
        # Current account frame
        current_frame = ttk.Frame(main_frame)
        current_frame.pack(fill="x", pady=(0, 10))
        
        self.current_label = ttk.Label(
            current_frame,
            text=f"Currently selected: {self.selected_account or 'None'}"
        )
        self.current_label.pack(side="left")
        
        # Dialog buttons
        dialog_buttons_frame = ttk.Frame(main_frame)
        dialog_buttons_frame.pack(fill="x")
        
        self.select_button = ttk.Button(
            dialog_buttons_frame,
            text="Use Selected Sender",
            command=self._select_account,
            state="disabled"
        )
        self.select_button.pack(side="right", padx=(5, 0))
        
        self.cancel_button = ttk.Button(
            dialog_buttons_frame,
            text="Close",
            command=self.window.destroy
        )
        self.cancel_button.pack(side="right")
        
        # Bind window close event
        self.window.protocol("WM_DELETE_WINDOW", self.window.destroy)
        
        # Load custom senders if present
        for em in self.accounts.get("custom_senders", []):
            if not isinstance(em, str):
                continue
        # If there is a selected sender, reflect it
        sel = self.accounts.get("selected_sender")
        if sel:
            self.current_label.config(text=f"Currently selected: {sel}")
    
    def _load_accounts(self) -> Dict:
        """Load account configurations from file with support for multiple senders."""
        config_dir = Path("config")
        config_dir.mkdir(exist_ok=True)
        accounts_file = config_dir / "accounts.json"
        
        data: Dict = {}
        if accounts_file.exists():
            try:
                with open(accounts_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except (json.JSONDecodeError, IOError):
                data = {}
        # Normalize structure
        if not isinstance(data, dict):
            data = {}
        data.setdefault("outlook_accounts", {})  # id -> {type,name,email}
        data.setdefault("custom_senders", [])     # list of email strings
        data.setdefault("selected_sender", None)  # selected email string
        return data
    
    def _save_accounts(self):
        """Save account configurations to file."""
        config_dir = Path("config")
        config_dir.mkdir(exist_ok=True)
        accounts_file = config_dir / "accounts.json"
        
        try:
            with open(accounts_file, 'w', encoding='utf-8') as f:
                json.dump(self.accounts, f, indent=2, ensure_ascii=False)
            print(f"[AccountDialog][DEBUG] saved accounts -> {accounts_file}")
            print(f"[AccountDialog][DEBUG] content: {self.accounts}")
        except IOError as e:
            print(f"[AccountDialog][DEBUG] save accounts failed: {e}")
            messagebox.showerror("Error", f"Failed to save accounts: {e}")
    
    def _populate_accounts(self):
        """Populate the account list and custom sender list."""
        print("[AccountDialog][DEBUG] populate accounts")
        self.account_listbox.delete(0, tk.END)
        # Accounts from Outlook connections
        for account_id, account_info in self.accounts.get("outlook_accounts", {}).items():
            display_name = account_info.get('name', account_id)
            email_address = account_info.get('email', '')
            label = f"{display_name} ({email_address})" if email_address else display_name
            self.account_listbox.insert(tk.END, label)
        # Custom senders
        self.custom_listbox.delete(0, tk.END)
        for em in self.accounts.get("custom_senders", []):
            self.custom_listbox.insert(tk.END, em)
        # No duplicate re-insert; keep list aligned to dict order
        print(f"[AccountDialog][DEBUG] listbox counts -> outlook={len(self.accounts.get('outlook_accounts', {}))}, custom={len(self.accounts.get('custom_senders', []))}")
    
    def _on_account_select(self, event):
        """Handle account selection."""
        selection = self.account_listbox.curselection()
        if selection:
            self.remove_button.config(state="normal")
            self.select_button.config(state="normal")
        else:
            self.remove_button.config(state="disabled")
            self.select_button.config(state="disabled")
    
    def _connect_to_outlook(self):
        """Connect to Outlook and add account."""
        try:
            print("[AccountDialog][DEBUG] connect_to_outlook requested")
            # Try to connect to Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            current_user = namespace.CurrentUser
            print(f"[AccountDialog][DEBUG] CurrentUser: {current_user}")
            
            # Get user email address
            email_address = current_user.Address if hasattr(current_user, 'Address') else "outlook@example.com"
            display_name = current_user.Name if hasattr(current_user, 'Name') else "Outlook User"
            print(f"[AccountDialog][DEBUG] detected outlook user -> name={display_name}, email={email_address}")
            
            # Add account
            account_id = f"outlook_{len(self.accounts.get('outlook_accounts', {})) + 1}"
            self.accounts["outlook_accounts"][account_id] = {
                "type": "outlook",
                "name": display_name,
                "email": email_address
            }
            # Also set as selected sender by default
            self.accounts["selected_sender"] = email_address
            
            self._save_accounts()
            self._populate_accounts()
            
            # Reflect in UI
            self.current_label.config(text=f"Currently selected: {display_name} ({email_address})")
            print(f"[AccountDialog][DEBUG] selecting account in service -> {account_id} {email_address}")
            self.email_service.set_current_account(account_id, {"type": "outlook", "name": display_name, "email": email_address})
            messagebox.showinfo("Success", f"Connected to Outlook as {display_name} ({email_address})")
            
        except Exception as e:
            print(f"[AccountDialog][DEBUG] connect_to_outlook failed: {e}")
            messagebox.showerror("Error", f"Failed to connect to Outlook: {str(e)}")
    
    def _add_custom_sender(self):
        """Add a custom sender email to the configuration."""
        email = (self.custom_email_entry.get() or "").strip()
        if not email:
            messagebox.showinfo("Info", "Please enter an email address to add.")
            return
        print(f"[AccountDialog][DEBUG] add custom sender -> {email}")
        # Initialize list if missing
        self.accounts.setdefault("custom_senders", [])
        if email in self.accounts["custom_senders"]:
            messagebox.showinfo("Info", "This email is already in the custom senders.")
            return
        self.accounts["custom_senders"].append(email)
        # Also set as selected sender by default after adding
        self.accounts["selected_sender"] = email
        self._save_accounts()
        self._populate_accounts()
        self.custom_email_entry.delete(0, tk.END)
        self.current_label.config(text=f"Currently selected: {email}")
        try:
            self.email_service.set_current_account("custom", {"type": "custom", "name": email, "email": email})
        except Exception as e:
            print(f"[AccountDialog][DEBUG] failed to set service current account for custom sender: {e}")

    def _remove_custom_sender(self):
        """Remove selected custom sender email from the configuration."""
        sel = self.custom_listbox.curselection()
        if not sel:
            messagebox.showinfo("Info", "Please select a custom sender to remove.")
            return
        email = self.custom_listbox.get(sel[0])
        print(f"[AccountDialog][DEBUG] remove custom sender -> {email}")
        try:
            self.accounts.setdefault("custom_senders", [])
            if email in self.accounts["custom_senders"]:
                self.accounts["custom_senders"].remove(email)
                # If the removed one was selected, clear selection
                if self.accounts.get("selected_sender") == email:
                    self.accounts["selected_sender"] = None
                self._save_accounts()
                self._populate_accounts()
                self.current_label.config(text="Currently selected: " + (self.accounts.get("selected_sender") or "None"))
        except Exception as e:
            print(f"[AccountDialog][DEBUG] remove custom sender failed: {e}")
            messagebox.showerror("Error", f"Failed to remove custom sender: {e}")

    def _disconnect_account(self):
        """Disconnect selected Outlook account."""
        selection = self.account_listbox.curselection()
        if not selection:
            return
        print(f"[AccountDialog][DEBUG] disconnect outlook account at index {selection[0]}")
        
        # Map UI index to account id
        account_ids = list(self.accounts.get("outlook_accounts", {}).keys())
        if selection[0] >= len(account_ids):
            return
        account_id = account_ids[selection[0]]
        removed_email = self.accounts["outlook_accounts"].get(account_id, {}).get("email")
        print(f"[AccountDialog][DEBUG] removing account_id={account_id}, email={removed_email}")
        
        # Remove account
        self.accounts["outlook_accounts"].pop(account_id, None)
        # If selected sender equals removed email, clear it
        if self.accounts.get("selected_sender") == removed_email:
            self.accounts["selected_sender"] = None
        
        self._save_accounts()
        self._populate_accounts()
        
        self.current_label.config(text="Currently selected: " + (self.accounts.get("selected_sender") or "None"))
        self.select_button.config(state="disabled")
        self.remove_button.config(state="disabled")
    
    def _select_account(self):
        """Select the chosen sender (from Outlook accounts or custom list)."""
        print("[AccountDialog][DEBUG] select_account")
        # Priority: if a custom sender is selected, use it; otherwise use Outlook account selection.
        custom_sel = self.custom_listbox.curselection()
        if custom_sel:
            sender_email = self.custom_listbox.get(custom_sel[0]).strip()
            if sender_email:
                print(f"[AccountDialog][DEBUG] selecting custom sender -> {sender_email}")
                self.accounts["selected_sender"] = sender_email
                self._save_accounts()
                self.current_label.config(text=f"Currently selected: {sender_email}")
                # Update service with a pseudo account using this sender
                try:
                    self.email_service.set_current_account("custom", {"type": "custom", "name": sender_email, "email": sender_email})
                except Exception as e:
                    print(f"[AccountDialog][DEBUG] set_current_account(custom) failed: {e}")
                return
        
        selection = self.account_listbox.curselection()
        if not selection:
            return
        
        # Map UI index to account id
        account_ids = list(self.accounts.get("outlook_accounts", {}).keys())
        if selection[0] >= len(account_ids):
            return
        account_id = account_ids[selection[0]]
        account_info = self.accounts["outlook_accounts"][account_id]
        
        # Set selected sender to the account email
        sender_email = account_info.get("email", "")
        print(f"[AccountDialog][DEBUG] selecting outlook account -> {account_id} {sender_email}")
        self.accounts["selected_sender"] = sender_email or None
        self._save_accounts()
        
        label = sender_email or account_info.get('name', account_id)
        self.current_label.config(text=f"Currently selected: {label}")
        # Update email service
        try:
            self.email_service.set_current_account(account_id, account_info)
        except Exception as e:
            print(f"[AccountDialog][DEBUG] set_current_account(outlook) failed: {e}")
    
    def show(self):
        """Show the dialog and wait for it to close. Return selected sender email."""
        self.window.wait_window()
        return self.accounts.get("selected_sender")
