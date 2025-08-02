"""Main GUI window for the PDF Stopover Analyzer."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import List, Optional, Set
import threading
import sys

from controllers.app_controller import AppController
from services.config_service import ConfigService
from ui.components.stopover_tab import StopoverTabComponent
from ui.components.mapping_tab import MappingTabComponent
from ui.components.email_preview_tab import EmailPreviewTabComponent
from ui.email_dialog import EmailDispatchDialog
from ui.account_dialog import AccountDialog
from ui.stopover_email_dialog import StopoverEmailDialog
from utils.file_utils import resource_path


class MainWindow:
    """Main application window using modular components."""
    
    def __init__(self):
        """Initialize the main window."""
        # Initialize services
        self.config_service = ConfigService()
        self.controller = AppController()
        
        # Setup window
        self._setup_window()
        
        # Initialize UI components
        self._setup_components()
        
        # Setup controller callbacks
        self._setup_controller_callbacks()
        
        # Initialize application state
        self._initialize_state()
        
        # Layout widgets
        self._layout_widgets()
        
        # Configure window closing
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _setup_window(self):
        """Setup the main window properties."""
        window_config = self.config_service.get_window_config()
        
        self.root = tk.Tk()
        self.root.title("PDF Stopover Analyzer")
        
        # Set window icon if available
        try:
            icon_path = resource_path("assets/app.ico")
            if sys.platform == "win32":
                self.root.iconbitmap(icon_path)
            else:
                # For non-Windows platforms, use iconphoto with PhotoImage
                from tkinter import PhotoImage
                icon = PhotoImage(file=icon_path)
                self.root.iconphoto(True, icon)
        except Exception:
            pass  # Ignore if icon is not available
        
        # Launch maximized (not borderless fullscreen) to keep system title bar and window controls visible
        try:
            self.root.state("zoomed")
        except Exception:
            # Fallback: set geometry to near-fullscreen if zoomed is unavailable
            try:
                sw = self.root.winfo_screenwidth()
                sh = self.root.winfo_screenheight()
                self.root.geometry(f"{max(800, sw-20)}x{max(600, sh-60)}+0+0")
            except Exception:
                # Final fallback to configured geometry
                self.root.geometry(f"{window_config.get('width', 1200)}x{window_config.get('height', 800)}")
        # Keep minimum size constraints
        self.root.minsize(window_config.get('min_width', 1000), window_config.get('min_height', 700))
    
    def _setup_components(self):
        """Setup UI components."""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        
        # Create tab components
        self.stopover_tab_component = StopoverTabComponent(
            self.notebook,
            on_stopover_select=self._on_stopover_select,
            controller=self.controller
        )
        
        self.mapping_tab_component = MappingTabComponent(
            self.notebook,
            on_mappings_change=self._on_mappings_change
        )
        
        self.email_preview_component = EmailPreviewTabComponent(
            self.notebook,
            controller=self.controller
        )
        
        # Add tabs to notebook
        self.notebook.add(self.stopover_tab_component.main_frame, text="Stopover Pages")
        self.notebook.add(self.mapping_tab_component.main_frame, text="Stopover Mappings")
        self.notebook.add(self.email_preview_component.main_frame, text="Email Preview")
        
        # Create file selection frame
        self._create_file_selection_frame()
        
        # Create status bar
        self._create_status_bar()
    
    def _create_file_selection_frame(self):
        """Create the file selection frame."""
        self.file_frame = ttk.Frame(self.root, padding="10")
        
        # File label and select button
        self.file_label = ttk.Label(self.file_frame, text="No PDF selected")
        self.select_button = ttk.Button(
            self.file_frame, 
            text="Select PDF", 
            command=self._select_pdf
        )
        
        
        
        # Account management frame
        self.account_frame = ttk.Frame(self.file_frame)
        self.account_status_label = ttk.Label(self.account_frame, text="Account: None")
        self.account_button = ttk.Button(
            self.account_frame,
            text="Manage Accounts",
            command=self._open_account_dialog
        )
    
    def _create_status_bar(self):
        """Create the status bar."""
        self.status_frame = ttk.Frame(self.root, padding="5")
        self.status_label = ttk.Label(self.status_frame, text="Ready")
        
        progress_config = self.config_service.get_ui_config()
        self.progress_bar = ttk.Progressbar(
            self.status_frame,
            mode='indeterminate',
            length=200
        )
        # Only visible during active processing
        try:
            self.progress_bar.pack_forget()
        except Exception:
            pass
    
    def _layout_widgets(self):
        """Layout all widgets in the window."""
        # File selection frame
        self.file_frame.pack(fill="x", padx=10, pady=(10, 5))
        
        # Left side of file frame
        self.file_label.pack(side="left", padx=(0, 10))
        self.select_button.pack(side="left", padx=(0, 5))
        
        
        # Right side of file frame (Account management)
        self.account_frame.pack(side="right")
        self.account_status_label.pack(side="left", padx=(0, 10))
        self.account_button.pack(side="left")
        # Connect/Disconnect Outlook button
        self.outlook_button = ttk.Button(
            self.file_frame,
            text="Connect Outlook",
            command=self._toggle_outlook_connection
        )
        self.outlook_button.pack(side="right", padx=(10, 0))
        
        # Notebook
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Status bar
        self.status_frame.pack(fill="x", padx=10, pady=(5, 10))
        self.status_label.pack(side="left")
        # Progress bar is packed dynamically only when active in _start_progress
    
    def _open_account_dialog(self):
        """Open the account management dialog."""
        dialog = AccountDialog(self.root, self.controller.email_service)
        selected_account = dialog.show()
        
        # Update account status display
        if selected_account:
            account_name = self.controller.email_service.get_current_account_name()
            self.account_status_label.config(text=f"Account: {account_name}")
        else:
            self.account_status_label.config(text="Account: None")
    
    def _setup_controller_callbacks(self):
        """Setup callbacks from the controller to update the UI."""
        self.controller.on_status_update = self._update_status
        self.controller.on_progress_start = self._start_progress
        self.controller.on_progress_stop = self._stop_progress
        self.controller.on_analysis_complete = self._on_analysis_complete
        self.controller.on_outlook_connection_change = self._on_outlook_connection_change
    
    def _initialize_state(self):
        """Initialize application state."""
        # Don't automatically connect to Outlook on startup
        # Let user explicitly connect when they want to
        self.account_status_label.config(text="Account: None")
        # Initialize Outlook button label
        try:
            self.outlook_button.config(text="Connect Outlook")
        except Exception:
            pass
        
        # Load initial mappings
        self.mapping_tab_component._load_mappings()
    
    def _select_pdf(self):
        """Handle PDF file selection."""
        filename = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        if filename:
            if self.controller.set_pdf_path(filename):
                import os
                self.file_label.config(text=f"Selected: {os.path.basename(filename)}")
                # Immediately start analysis
                # analyze_button removed
                
                # Clear stopover tab FIRST
                self.stopover_tab_component.clear()
                
                # Clear email preview tab
                self.email_preview_component.clear()
                
                # Set PDF path on stopover tab component
                self.stopover_tab_component.set_pdf_path(filename)
                
                # Update mapping tab
                self.mapping_tab_component.set_found_stopovers(set())
                
                # Automatically start analysis
                self.status_label.config(text="Analyzing PDF...")
                self._start_progress()
                self.controller.analyze_pdf()
            else:
                messagebox.showerror("Error", "Please select a valid PDF file.")
    
        
    
    def _on_analysis_complete(self, stopovers: List):
        """Handle analysis completion from controller."""
        # Update stopover tab
        self.stopover_tab_component.set_stopovers(stopovers)
        
        # Update mapping tab with found stopovers
        found_codes = {s.code for s in stopovers}
        self.mapping_tab_component.set_found_stopovers(found_codes)
        
        # Update email preview tab
        self.email_preview_component.set_stopovers(stopovers)
        self.email_preview_component.set_pdf_path(self.controller.current_pdf_path)
        
        # Update status
        self.status_label.config(text=f"Found {len(stopovers)} stopover(s)")
        
        # Stop progress
        self._stop_progress()
    
    def _update_status(self, message: str):
        """Update status message."""
        self.status_label.config(text=message)
    
    def _start_progress(self):
        """Start progress animation."""
        try:
            # Ensure it is visible and positioned
            self.progress_bar.pack(side="right")
        except Exception:
            pass
        self.progress_bar.start()
    
    def _stop_progress(self):
        """Stop progress animation."""
        try:
            self.progress_bar.stop()
        except Exception:
            pass
        # Hide when not actively processing
        try:
            self.progress_bar.pack_forget()
        except Exception:
            pass
        
    
    def _on_mappings_change(self):
        """Handle mappings change event."""
        # Refresh mapping display
        self.mapping_tab_component._load_mappings()
        # Immediately reflect changes in Email Preview tab (no manual refresh)
        try:
            self.email_preview_component.refresh_recipients_from_configs()
        except Exception:
            # Safe-guard: rebuild sections if incremental refresh fails
            try:
                self.email_preview_component.set_stopovers(self.controller.stopovers or [])
                if self.controller.current_pdf_path:
                    self.email_preview_component.set_pdf_path(self.controller.current_pdf_path)
            except Exception:
                pass
    
    def _on_stopover_select(self, stopover):
        """Handle stopover selection from stopover tab."""
        # Load page preview through stopover tab component
        def progress_callback(message):
            self._update_status(message)
        
        # Use the stopover tab component's built-in preview loading
        self.stopover_tab_component.load_page_preview(stopover, progress_callback)
        
        # Remove verbose preview size status; keep concise selection feedback if needed
        # (No status update for preview size to avoid clutter)
    
    def _on_outlook_connection_change(self, connected: bool, user: Optional[str]):
        """Update UI when Outlook connection state changes (hide email from UI)."""
        if connected:
            # Do not display connected email or username in the UI anymore
            self.account_status_label.config(text="Account: Connected")
            try:
                self.outlook_button.config(text="Disconnect Outlook")
            except Exception:
                pass
        else:
            self.account_status_label.config(text="Account: None")
            try:
                self.outlook_button.config(text="Connect Outlook")
            except Exception:
                pass

    def _toggle_outlook_connection(self):
        """Connect or disconnect Outlook based on current state."""
        # Let controller manage connect/disconnect and emit callback
        self.controller.toggle_outlook_connection()

    def _open_email_dialog(self):
        """Open the email dispatch dialog."""
        if not self.controller.stopovers:
            messagebox.showinfo("Info", "Please analyze a PDF first to find stopovers.")
            return
        
        # Check if Outlook is connected before sending emails
        if not self.controller.outlook_connected:
            response = messagebox.askyesno(
                "Outlook Not Connected",
                "Outlook is not connected. Do you want to connect now?"
            )
            if response:
                if not self.controller.connect_to_outlook():
                    messagebox.showerror("Error", "Failed to connect to Outlook.")
                    return
        
        EmailDispatchDialog(self.root, self.controller.stopovers, self.controller.current_pdf_path)
    
    
    def _on_closing(self):
        """Handle window closing."""
        # Clean up controller resources
        self.controller.destroy()
        
        # Destroy window
        self.root.destroy()
    
    def run(self):
        """Start the GUI event loop."""
        self.root.mainloop()
