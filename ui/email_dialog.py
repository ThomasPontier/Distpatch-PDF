"""Email dispatch dialog for selecting stopovers to send emails."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict
import threading

from models.stopover import Stopover
from controllers.app_controller import AppController


class EmailDispatchDialog:
    """Dialog for selecting which stopovers to send emails for."""
    
    def __init__(self, parent, stopovers: List[Stopover], pdf_path: str):
        """Initialize the email dispatch dialog."""
        self.parent = parent
        self.stopovers = stopovers
        self.pdf_path = pdf_path
        self.controller = AppController()
        
        # Filter stopovers that have email mappings
        self.mapped_stopovers = [
            stopover for stopover in stopovers
            if self.controller.has_mapping(stopover.code)
        ]
        
        # Create window
        self.window = tk.Toplevel(parent)
        self.window.title("Send Emails")
        self.window.geometry("500x400")
        self.window.minsize(400, 300)
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
        # Create widgets
        self._create_widgets()
        self._layout_widgets()
        
        # Check if any stopovers have mappings
        if not self.mapped_stopovers:
            messagebox.showinfo(
                "No Mappings",
                "No stopovers found with email mappings.\n\n"
                "Please configure email mappings first using 'Manage Mappings'."
            )
            self.window.destroy()
            return
    
    def _create_widgets(self):
        """Create all GUI widgets."""
        # Main frame
        self.main_frame = ttk.Frame(self.window, padding="10")
        
        # Info frame
        self.info_frame = ttk.LabelFrame(self.main_frame, text="Email Summary", padding="5")
        
        # Info labels
        total_label = ttk.Label(
            self.info_frame,
            text=f"Total stopovers found: {len(self.stopovers)}"
        )
        
        mapped_label = ttk.Label(
            self.info_frame,
            text=f"Stopovers with email mappings: {len(self.mapped_stopovers)}"
        )
        
        total_label.pack(anchor="w")
        mapped_label.pack(anchor="w")
        
        # Selection frame
        self.selection_frame = ttk.LabelFrame(
            self.main_frame, 
            text="Select Stopovers to Email", 
            padding="5"
        )
        
        # Create checkboxes for each mapped stopover
        self.check_vars = {}
        self.checkbuttons = {}
        
        for stopover in self.mapped_stopovers:
            var = tk.BooleanVar(value=True)  # Default to checked
            self.check_vars[stopover.code] = var
            
            # Get emails for this stopover
            emails = self.controller.get_emails_for_stopover(stopover.code)
            emails_str = ", ".join(emails[:2])  # Show first 2 emails
            if len(emails) > 2:
                emails_str += f" (+{len(emails) - 2} more)"
            
            text = f"{stopover.code} (Page {stopover.page_number}) → {emails_str}"
            
            cb = ttk.Checkbutton(
                self.selection_frame,
                text=text,
                variable=var
            )
            cb.pack(anchor="w", padx=5, pady=2)
            self.checkbuttons[stopover.code] = cb
        
        # Scrollbar if needed
        self.canvas = tk.Canvas(self.selection_frame, height=200)
        self.scrollbar = ttk.Scrollbar(
            self.selection_frame,
            orient="vertical",
            command=self.canvas.yview
        )
        
        # Buttons frame
        self.buttons_frame = ttk.Frame(self.main_frame)
        
        self.select_all_btn = ttk.Button(
            self.buttons_frame,
            text="Select All",
            command=self._select_all
        )
        
        self.deselect_all_btn = ttk.Button(
            self.buttons_frame,
            text="Deselect All",
            command=self._deselect_all
        )
        
        self.send_btn = ttk.Button(
            self.buttons_frame,
            text="Send Emails",
            command=self._send_emails
        )
        
        self.cancel_btn = ttk.Button(
            self.buttons_frame,
            text="Cancel",
            command=self.window.destroy
        )
        
        # Progress frame
        self.progress_frame = ttk.Frame(self.main_frame)
        self.progress_label = ttk.Label(self.progress_frame, text="")
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            mode='indeterminate',
            length=300
        )
    
    def _layout_widgets(self):
        """Layout all widgets."""
        self.main_frame.pack(fill="both", expand=True)
        
        # Info frame
        self.info_frame.pack(fill="x", pady=(0, 10))
        
        # Selection frame
        self.selection_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Buttons frame
        self.buttons_frame.pack(fill="x")
        self.select_all_btn.pack(side="left", padx=(0, 5))
        self.deselect_all_btn.pack(side="left", padx=(0, 5))
        self.cancel_btn.pack(side="right")
        self.send_btn.pack(side="right", padx=(0, 5))
        
        # Progress frame (initially hidden)
        self.progress_frame.pack(fill="x", pady=(5, 0))
        self.progress_label.pack()
        self.progress_bar.pack()
        self.progress_frame.pack_forget()  # Hide initially
    
    def _select_all(self):
        """Select all checkboxes."""
        for var in self.check_vars.values():
            var.set(True)
    
    def _deselect_all(self):
        """Deselect all checkboxes."""
        for var in self.check_vars.values():
            var.set(False)
    
    def _send_emails(self):
        """Send emails for selected stopovers."""
        # Get selected stopovers
        selected_stopovers = [
            stopover for stopover in self.mapped_stopovers
            if self.check_vars[stopover.code].get()
        ]
        
        if not selected_stopovers:
            messagebox.showinfo("Info", "Please select at least one stopover to send emails.")
            return
        
        # Confirm with user
        response = messagebox.askyesno(
            "Confirm Send",
            f"Send emails for {len(selected_stopovers)} stopover(s)?"
        )
        
        if response:
            # Start sending in background
            self._start_sending(selected_stopovers)
    
    def _start_sending(self, selected_stopovers: List[Stopover]):
        """Start sending emails in background thread."""
        # Show progress
        self.progress_frame.pack(fill="x", pady=(5, 0))
        self.progress_label.config(text="Sending emails...")
        self.progress_bar.start()
        
        # Disable buttons
        self.send_btn.config(state="disabled")
        self.select_all_btn.config(state="disabled")
        self.deselect_all_btn.config(state="disabled")
        
        # Start sending in thread
        thread = threading.Thread(
            target=self._send_emails_thread,
            args=(selected_stopovers,)
        )
        thread.daemon = True
        thread.start()
    
    def _send_emails_thread(self, selected_stopovers: List[Stopover]):
        """Send emails in background thread."""
        success_count, total_count = self.controller.send_stopover_emails(
            selected_stopovers, 
            self.pdf_path
        )
        
        # Update UI in main thread
        self.window.after(0, lambda: self._sending_complete(success_count, total_count))
    
    def _sending_complete(self, success_count: int, total_count: int):
        """Called when email sending is complete."""
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        
        # Re-enable buttons
        self.send_btn.config(state="normal")
        self.select_all_btn.config(state="normal")
        self.deselect_all_btn.config(state="normal")
        
        # Show result
        if success_count == total_count:
            messagebox.showinfo(
                "Success",
                f"Successfully sent {success_count} email(s)."
            )
            self.window.destroy()
        else:
            messagebox.showwarning(
                "Partial Success",
                f"Sent {success_count} out of {total_count} email(s).\n"
                f"Check console for details."
            )
    
    def _sending_error(self, error_msg: str):
        """Called when email sending fails."""
        self.progress_bar.stop()
        self.progress_frame.pack_forget()
        
        # Re-enable buttons
        self.send_btn.config(state="normal")
        self.select_all_btn.config(state="normal")
        self.deselect_all_btn.config(state="normal")
        
        messagebox.showerror(
            "Error",
            f"Failed to send emails: {error_msg}"
        )
