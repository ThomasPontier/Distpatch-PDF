"""Stopover tab component for displaying stopover pages and previews."""

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from typing import List, Optional, Callable
import threading

from core.pdf_renderer import PDFRenderer
from models.stopover import Stopover
from utils.file_utils import validate_pdf_file
from .base_component import BaseUIComponent
from ..stopover_email_dialog import StopoverEmailDialog


class StopoverTabComponent(BaseUIComponent):
    """UI component for the stopover pages tab."""
    
    def __init__(self, parent, on_stopover_select: Callable[[Stopover], None] = None, controller=None):
        """Initialize the stopover tab component."""
        super().__init__(parent)
        self.parent = parent
        self.on_stopover_select = on_stopover_select
        self.controller = controller
        self.stopovers: List[Stopover] = []
        self.pdf_renderer: Optional[PDFRenderer] = None
        self.current_pdf_path: Optional[str] = None
        
        # Create the tab content
        self.create_tab_content()
    
    def create_tab_content(self):
        """Create the content for the stopover tab."""
        # Main content frame for stopover tab
        self.main_frame = self.create_frame(self.parent, padding="10")
        
        # Configure grid weights for responsive layout
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=3)
        self.main_frame.rowconfigure(0, weight=1)
        
        # Create stopover list and preview
        self._create_stopover_list()
        self._create_preview_area()
    
    def _create_stopover_list(self):
        """Create the stopover list widget."""
        # Stopover list frame
        self.list_frame = ttk.LabelFrame(self.main_frame, text="Stopover Pages", padding="5")
        
        # Listbox with scrollbar
        self.stopover_listbox = tk.Listbox(
            self.list_frame,
            height=20,
            width=30,
            selectmode=tk.SINGLE
        )
        
        self.stopover_scrollbar = ttk.Scrollbar(
            self.list_frame,
            orient="vertical",
            command=self.stopover_listbox.yview
        )
        self.stopover_listbox.configure(yscrollcommand=self.stopover_scrollbar.set)
        
        # Bind double-click event
        self.stopover_listbox.bind('<Double-Button-1>', self._on_stopover_double_click)
        
        # Bind right-click event for context menu
        self.stopover_listbox.bind('<Button-3>', self._on_stopover_right_click)
        
        # Create context menu
        self.context_menu = tk.Menu(self.stopover_listbox, tearoff=0)
        self.context_menu.add_command(label="Configure Email Settings", command=self._configure_email_settings)
        
        # Layout list widgets
        self.stopover_listbox.pack(side="left", fill="both", expand=True)
        self.stopover_scrollbar.pack(side="right", fill="y")
        
        # Add to main frame
        self.list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
    
    def _create_preview_area(self):
        """Create the page preview area."""
        # Preview frame
        self.preview_frame = ttk.LabelFrame(self.main_frame, text="Page Preview", padding="5")
        
        # Preview label
        self.preview_label = ttk.Label(
            self.preview_frame,
            text="Select a stopover to preview its page",
            anchor="center"
        )
        
        # Layout preview
        self.preview_label.pack(fill="both", expand=True)

        # Track last rendered PIL image to rescale on container resize
        self._last_rendered_image = None
        self.preview_frame.bind("<Configure>", self._on_preview_resize)
        
        # Add to main frame
        self.preview_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
    
    def set_pdf_path(self, pdf_path: str):
        """Set the current PDF path and update the UI."""
        if validate_pdf_file(pdf_path):
            self.current_pdf_path = pdf_path
            self.close_pdf_renderer()
            return True
        return False
    
    def set_stopovers(self, stopovers: List[Stopover]):
        """Update the stopover list with new data."""
        self.stopovers = stopovers
        self._update_stopover_list()
    
    def _update_stopover_list(self):
        """Update the stopover list with current data."""
        self.stopover_listbox.delete(0, tk.END)
        
        if self.stopovers:
            # Show only stopover codes, not page numbers
            for stopover in self.stopovers:
                self.stopover_listbox.insert(tk.END, stopover.code)
    
    def _on_stopover_double_click(self, event):
        """Handle double-click on stopover list."""
        selection = self.stopover_listbox.curselection()
        if not selection:
            return
        
        # Find the actual stopover object
        selected_code = self.stopover_listbox.get(selection[0])
        selected_stopover = None
        for stopover in self.stopovers:
            if stopover.code == selected_code:
                selected_stopover = stopover
                break
        
        if selected_stopover and self.on_stopover_select:
            self.on_stopover_select(selected_stopover)
    
    def load_page_preview(self, stopover: Stopover, progress_callback: Callable[[str], None] = None):
        """Load and display page preview for a stopover."""
        if not self.current_pdf_path:
            return
        
        try:
            # Show progress if callback provided
            if progress_callback:
                progress_callback("Loading page preview...")
            
            # Create or update PDF renderer
            if not self.pdf_renderer or self.pdf_renderer.pdf_path != self.current_pdf_path:
                self.close_pdf_renderer()
                self.pdf_renderer = PDFRenderer(self.current_pdf_path)
            
            # Get base page image at a good quality (will scale to fit later)
            img = self.pdf_renderer.get_page_image(stopover.page_number, max_width=1600, max_height=1600)
            # Keep the PIL image to enable responsive resizing
            self._last_rendered_image = img
            
            # Update preview in main thread (image scaling happens inside)
            self.parent.after(0, self._fit_and_update_preview)
            
        except Exception as e:
            error_msg = f"Error loading page preview: {str(e)}"
            if progress_callback:
                progress_callback(error_msg)
            self.parent.after(0, lambda: self._show_error(error_msg))
    
    def _fit_and_update_preview(self):
        """Scale last rendered image to fit available area while maintaining aspect ratio and update label."""
        try:
            if self._last_rendered_image is None:
                return
            # Compute available area inside preview_frame
            avail_w = max(1, self.preview_frame.winfo_width() - 16)  # some padding
            avail_h = max(1, self.preview_frame.winfo_height() - 32)
            img = self._last_rendered_image
            iw, ih = img.size
            if iw <= 0 or ih <= 0:
                return
            # Maintain aspect ratio: fit within available area
            scale = min(avail_w / iw, avail_h / ih)
            target_w = max(1, int(iw * scale))
            target_h = max(1, int(ih * scale))
            if target_w <= 0 or target_h <= 0:
                return
            # Use high-quality resize
            resized = img.resize((target_w, target_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(resized)
            self.preview_label.configure(image=photo, text="")
            self.preview_label.image = photo  # Keep a reference
        except Exception as _:
            # Fallback to text if scaling fails
            self.preview_label.configure(image="", text="Preview unavailable")

    def _on_preview_resize(self, event):
        """Handle container resize and refit the preview image."""
        try:
            # Debounce by scheduling on next loop, ensures stable dimensions
            if hasattr(self, "_resize_after_id") and self._resize_after_id:
                self.parent.after_cancel(self._resize_after_id)
        except Exception:
            pass
        self._resize_after_id = self.parent.after(50, self._fit_and_update_preview)
    
    def _show_error(self, error_msg: str):
        """Show error message in preview area."""
        self.preview_label.configure(image="", text=error_msg)
    
    def close_pdf_renderer(self):
        """Close the PDF renderer if it's open."""
        if self.pdf_renderer:
            self.pdf_renderer.close()
            self.pdf_renderer = None
    
    def clear(self):
        """Clear all data and reset the UI."""
        self.stopovers = []
        self.current_pdf_path = None
        self.stopover_listbox.delete(0, tk.END)
        self.preview_label.configure(image="", text="Select a stopover to preview its page")
        self.close_pdf_renderer()
    
    def destroy(self):
        """Clean up resources when destroying the component."""
        self.close_pdf_renderer()
        if self.main_frame:
            self.main_frame.destroy()
    
    def _on_stopover_right_click(self, event):
        """Handle right-click on stopover list to show context menu."""
        # Select the item under the cursor
        index = self.stopover_listbox.nearest(event.y)
        if index >= 0:
            self.stopover_listbox.selection_clear(0, tk.END)
            self.stopover_listbox.selection_set(index)
            self.stopover_listbox.activate(index)
            
            # Show context menu
            try:
                self.context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.context_menu.grab_release()
    
    def _configure_email_settings(self):
        """Open email configuration dialog for selected stopover."""
        if not self.controller:
            messagebox.showerror("Error", "Controller not available")
            return
        
        selection = self.stopover_listbox.curselection()
        if not selection:
            messagebox.showinfo("Info", "Please select a stopover first")
            return
        
        # Find the actual stopover object
        selected_code = self.stopover_listbox.get(selection[0])
        selected_stopover = None
        for stopover in self.stopovers:
            if stopover.code == selected_code:
                selected_stopover = stopover
                break
        
        if not selected_stopover:
            messagebox.showerror("Error", "Selected stopover not found")
            return
        
        # Open email configuration dialog
        try:
            dialog = StopoverEmailDialog(self.parent, selected_stopover.code, self.controller.stopover_email_service)
            dialog.show()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open email configuration dialog: {str(e)}")
