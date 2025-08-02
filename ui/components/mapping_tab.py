"""Mapping tab component for managing stopover-to-email mappings."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Set, Optional
import re

from services.mapping_service import MappingService
from services.stopover_email_service import StopoverEmailService
from services.config_manager import get_config_manager
from .base_component import BaseUIComponent


class MappingTabComponent(BaseUIComponent):
    """UI component for the stopover mappings tab."""
    
    def __init__(self, parent, on_mappings_change: callable = None):
        """Initialize the mapping tab component."""
        super().__init__(parent)
        self.parent = parent
        self.on_mappings_change = on_mappings_change
        self.mapping_service = MappingService()
        self.found_stopover_codes: Set[str] = set()
        
        # Create the tab content
        self.create_tab_content()
    
    def create_tab_content(self):
        """Create the content for the mapping tab."""
        # Main content frame for mappings tab
        self.main_frame = self.create_frame(self.parent, padding="10")
        
        # Configure grid weights
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(1, weight=1)
        
        # Create input section
        self._create_input_section()
        
        # Create mappings display
        self._create_mappings_display()
        
        # Create action buttons
        self._create_action_buttons()
        # Subscribe to ConfigManager changes to keep table reactive
        try:
            self._config_manager = get_config_manager()
            self._config_manager.on_mappings_changed(lambda _m: self._load_mappings())
            self._config_manager.on_stopovers_changed(lambda _s: self._load_mappings())
            self._config_manager.on_last_sent_changed(lambda _l: self._load_mappings())
        except Exception:
            pass
        # Initial load to ensure UI reflects current unified config
        try:
            self._load_mappings()
        except Exception:
            pass
    
    def _create_input_section(self):
        """Create the input section for adding new mappings."""
        # Input frame
        self.input_frame = ttk.LabelFrame(self.main_frame, text="Add New Mapping", padding="5")
        
        # Stopover code input
        self.code_label = self.create_label(self.input_frame, text="Stopover Code:")
        self.code_entry = self.create_entry(self.input_frame, width=10)
        self.code_entry.bind('<KeyRelease>', self._on_code_change)
        
        # Email input
        self.email_label = self.create_label(self.input_frame, text="Email Address:")
        self.email_entry = self.create_entry(self.input_frame, width=30)
        
        # Add button
        self.add_button = self.create_button(
            self.input_frame, 
            text="Add Mapping", 
            command=self._add_mapping
        )
        
        # Layout input widgets
        self.code_label.grid(row=0, column=0, padx=(0, 5), sticky="e")
        self.code_entry.grid(row=0, column=1, padx=(0, 10), sticky="w")
        self.email_label.grid(row=0, column=2, padx=(0, 5), sticky="e")
        self.email_entry.grid(row=0, column=3, padx=(0, 10), sticky="w")
        self.add_button.grid(row=0, column=4, sticky="w")
        
        # Add to main frame
        self.input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    
    def _create_mappings_display(self):
        """Create the mappings display area."""
        # Mappings frame
        self.mappings_frame = ttk.LabelFrame(self.main_frame, text="Stopover Mappings", padding="5")
        
        # Treeview for mappings with status column
        self.mapping_tree = ttk.Treeview(
            self.mappings_frame,
            columns=("stopover", "emails", "status", "last_sent"),
            show="headings",
            height=15
        )
        
        # Configure tree columns
        self.mapping_tree.heading("stopover", text="Stopover Code")
        self.mapping_tree.heading("emails", text="Email Addresses")
        self.mapping_tree.heading("status", text="Found in PDF")
        self.mapping_tree.heading("last_sent", text="Last Sent")
        self.mapping_tree.column("stopover", width=100, minwidth=80)
        self.mapping_tree.column("emails", width=300, minwidth=200)
        self.mapping_tree.column("status", width=100, minwidth=80)
        self.mapping_tree.column("last_sent", width=150, minwidth=120)
        
        # Scrollbar for tree
        self.mapping_scrollbar = ttk.Scrollbar(
            self.mappings_frame,
            orient="vertical",
            command=self.mapping_tree.yview
        )
        self.mapping_tree.configure(yscrollcommand=self.mapping_scrollbar.set)
        
        # Layout mappings widgets
        self.mapping_tree.pack(side="left", fill="both", expand=True)
        self.mapping_scrollbar.pack(side="right", fill="y")
        
        # Add to main frame
        self.mappings_frame.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
    
    def _create_action_buttons(self):
        """Create action buttons for mappings."""
        # Buttons frame
        self.buttons_frame = ttk.Frame(self.main_frame)
        
        # Action buttons
        self.remove_button = self.create_button(
            self.buttons_frame,
            text="Remove Selected",
            command=self._remove_selected_mapping
        )
        
        self.edit_button = self.create_button(
            self.buttons_frame,
            text="Edit Selected",
            command=self._edit_selected_mapping
        )
        
        # Layout buttons
        self.remove_button.pack(side="left", padx=(0, 5))
        self.edit_button.pack(side="left", padx=(0, 5))
        
        # Add to main frame
        self.buttons_frame.grid(row=2, column=0, sticky="ew")
    
    def set_found_stopovers(self, stopover_codes: Set[str]):
        """Update the found stopover codes and refresh the display."""
        self.found_stopover_codes = stopover_codes
        self._update_mapping_status()
    
    def _load_mappings(self):
        """Load and display all mappings, including found stopovers."""
        # Get existing mappings from unified ConfigManager (service now uses it)
        try:
            service_mappings = self.mapping_service.get_all_mappings()
            service_mappings = {str(k).upper(): list(v or []) for k, v in (service_mappings or {}).items()}
        except Exception:
            service_mappings = {}

        # Global subject field removed from this tab per request.
        
        # Combine service mappings with found stopovers that might not have mappings
        all_mappings = service_mappings.copy()
        for stopover_code in self.found_stopover_codes:
            if stopover_code not in all_mappings:
                all_mappings[stopover_code] = []
        
        # Clear existing items
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)
        
        # Add to tree using last_sent from centralized manager
        ses = StopoverEmailService()
        for stopover_code, emails in sorted(all_mappings.items()):
            emails_str = "; ".join(emails) if emails else ""
            status = "✓ Found" if stopover_code in self.found_stopover_codes else "○ Not Found"
            raw_last = ses.get_last_sent(stopover_code) or ""
            def _fmt(val: str) -> str:
                try:
                    s = (val or "").strip()
                    if not s:
                        return ""
                    if s.endswith("Z"):
                        s = s[:-1] + "+00:00"
                    from datetime import datetime
                    try:
                        dt = datetime.fromisoformat(s)
                    except Exception:
                        dt = datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
                    return dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    return val or ""
            last_sent = _fmt(raw_last) if raw_last else ""
            self.mapping_tree.insert("", "end", values=(stopover_code, emails_str, status, last_sent))
    
    def _update_mapping_status(self):
        """Update the status column in the mapping tree."""
        # Get current mappings from tree
        current_mappings = {}
        for item in self.mapping_tree.get_children():
            values = self.mapping_tree.item(item, 'values')
            if values:
                stopover_code = values[0]
                emails_str = values[1] if len(values) > 1 else ""
                current_mappings[stopover_code] = emails_str
        
        # Add any new stopovers found in PDF that aren't already in mappings
        for stopover_code in self.found_stopover_codes:
            if stopover_code not in current_mappings:
                # Add new stopover with no emails
                current_mappings[stopover_code] = ""
        
        # Clear and rebuild the tree with updated data
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)
        
        # Add all mappings (existing + new) with updated status
        ses = StopoverEmailService()
        for stopover_code, emails_str in sorted(current_mappings.items()):
            status = "✓ Found" if stopover_code in self.found_stopover_codes else "○ Not Found"
            raw_last = ses.get_last_sent(stopover_code) or ""
            def _fmt(val: str) -> str:
                try:
                    s = (val or "").strip()
                    if not s:
                        return ""
                    if s.endswith("Z"):
                        s = s[:-1] + "+00:00"
                    from datetime import datetime
                    try:
                        dt = datetime.fromisoformat(s)
                    except Exception:
                        dt = datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
                    return dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    return val or ""
            last_sent = _fmt(raw_last) if raw_last else ""
            self.mapping_tree.insert("", "end", values=(stopover_code, emails_str, status, last_sent))
    
    def _validate_email(self, email: str) -> bool:
        """Validate email address format."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def _validate_stopover_code(self, code: str) -> bool:
        """Validate stopover code format (3 letters)."""
        return len(code) == 3 and code.isalpha()
    
    def _add_mapping(self):
        """Add a new mapping."""
        code = self.code_entry.get().strip().upper()
        email = self.email_entry.get().strip()
        
        # Validate inputs
        if not code:
            messagebox.showerror("Error", "Please enter a stopover code.")
            return
        
        if not self._validate_stopover_code(code):
            messagebox.showerror("Error", "Stopover code must be exactly 3 letters.")
            return
        
        if not email:
            messagebox.showerror("Error", "Please enter an email address.")
            return
        
        if not self._validate_email(email):
            messagebox.showerror("Error", "Please enter a valid email address.")
            return
        
        # Add mapping via centralized manager to ensure immediate persistence/signals
        try:
            mgr = get_config_manager()
            existing = mgr.get_mappings().get(code, [])
            if email in existing:
                messagebox.showinfo("Info", f"Email {email} already exists for {code}")
                return
            new_emails = existing + [email]
            mgr.set_mapping(code, new_emails)
            mgr.add_stopover(code)  # ensure stopover is listed/enabled
            self._load_mappings()
            self.code_entry.delete(0, tk.END)
            self.email_entry.delete(0, tk.END)
            self.code_entry.focus()
            if self.on_mappings_change:
                self.on_mappings_change()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add mapping: {e}")
    
    def _remove_selected_mapping(self):
        """Remove the selected mapping."""
        selection = self.mapping_tree.selection()
        if not selection:
            messagebox.showinfo("Info", "Please select a mapping to remove.")
            return
        
        item = selection[0]
        values = self.mapping_tree.item(item, 'values')
        if values:
            stopover_code = values[0]
            emails_str = values[1]
            
            # Ask for confirmation
            response = messagebox.askyesno(
                "Confirm Removal",
                f"Remove all email mappings for {stopover_code}?"
            )
            
            if response:
                # Remove mapping in centralized manager
                try:
                    mgr = get_config_manager()
                    # Remove entire mapping key and stopover entry
                    mgr.remove_mapping(stopover_code)
                    # Keep stopover if you prefer, or remove entirely if no mapping needed
                    # Here we keep stopover presence unchanged.
                    self._load_mappings()
                    if self.on_mappings_change:
                        self.on_mappings_change()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to remove mapping: {e}")
    
    def _edit_selected_mapping(self):
        """Edit the selected mapping."""
        selection = self.mapping_tree.selection()
        if not selection:
            messagebox.showinfo("Info", "Please select a mapping to edit.")
            return
        
        item = selection[0]
        values = self.mapping_tree.item(item, 'values')
        if values:
            stopover_code = values[0]
            emails_str = values[1]
            emails = [e.strip() for e in emails_str.split(';') if e.strip()]
            
            # Open edit dialog
            self._open_edit_dialog(stopover_code, emails)
    
    def _open_edit_dialog(self, stopover_code: str, current_emails: List[str]):
        """Open dialog to edit mappings for a stopover."""
        dialog = tk.Toplevel(self.parent)
        dialog.title(f"Edit {stopover_code} Mappings")
        dialog.geometry("400x300")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        # Frame
        frame = ttk.Frame(dialog, padding="10")
        frame.pack(fill="both", expand=True)
        
        # Label
        label = self.create_label(frame, text=f"Email addresses for {stopover_code}:")
        label.pack(pady=(0, 5))
        
        # Text widget for emails
        text = tk.Text(frame, height=10, width=40)
        text.pack(fill="both", expand=True, pady=(0, 10))
        
        # Insert current emails (deduplicated, non-empty)
        clean_current = []
        for e in current_emails or []:
            s = e.strip()
            if s and s not in clean_current:
                clean_current.append(s)
        text.insert("1.0", "\n".join(clean_current))
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x")
        
        def save_changes():
            # Get emails from text widget
            content = text.get("1.0", tk.END).strip()
            new_emails = [e.strip() for e in content.split('\n') if e.strip()]

            # Validate stopover code and emails
            code = str(stopover_code).upper().strip()
            if not code or not self._validate_stopover_code(code):
                messagebox.showerror("Error", "Stopover code must be exactly 3 letters (A-Z).")
                return

            valid_emails = []
            seen = set()
            for email in new_emails:
                if self._validate_email(email):
                    if email not in seen:
                        valid_emails.append(email)
                        seen.add(email)
                else:
                    messagebox.showwarning("Warning", f"Invalid email skipped: {email}")

            # If empty, confirm removal of mapping
            if not valid_emails:
                resp = messagebox.askyesno("Confirm", f"No valid emails remain for {code}. Remove mapping?")
                if not resp:
                    return
                try:
                    mgr = get_config_manager()
                    mgr.remove_mapping(code)
                    self._load_mappings()
                    if self.on_mappings_change:
                        self.on_mappings_change()
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to remove mapping: {e}")
                dialog.destroy()
                return

            # Update mappings centrally
            try:
                mgr = get_config_manager()
                mgr.set_mapping(code, valid_emails)
                mgr.add_stopover(code)  # ensure presence/enabled
                self._load_mappings()
                if self.on_mappings_change:
                    self.on_mappings_change()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to update mappings: {e}")
            finally:
                dialog.destroy()
        
        def cancel():
            dialog.destroy()
        
        save_btn = self.create_button(button_frame, text="Save", command=save_changes)
        save_btn.pack(side="left", padx=(0, 5))
        
        cancel_btn = self.create_button(button_frame, text="Cancel", command=cancel)
        cancel_btn.pack(side="left")
    
    def _on_code_change(self, event):
        """Auto-uppercase the stopover code."""
        current = self.code_entry.get()
        if current != current.upper():
            self.code_entry.delete(0, tk.END)
            self.code_entry.insert(0, current.upper())
    
    def get_mapped_stopovers(self) -> List[str]:
        """Get all stopover codes that have email mappings."""
        try:
            return sorted(list(get_config_manager().get_mappings().keys()))
        except Exception:
            return []
    
    def has_mapping(self, stopover_code: str) -> bool:
        """Check if a stopover code has any email mappings."""
        try:
            maps = get_config_manager().get_mappings()
            return stopover_code in maps and len(maps.get(stopover_code) or []) > 0
        except Exception:
            return False
    
    def get_emails_for_stopover(self, stopover_code: str) -> List[str]:
        """Get email addresses for a stopover code."""
        try:
            return list(get_config_manager().get_mappings().get(stopover_code, []))
        except Exception:
            return []
    
    def destroy(self):
        """Clean up resources when destroying the component."""
        if self.main_frame:
            self.main_frame.destroy()
