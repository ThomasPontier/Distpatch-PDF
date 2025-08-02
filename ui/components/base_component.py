"""Base UI component for consistent UI element creation."""

import tkinter as tk
from tkinter import ttk
from typing import Optional


class BaseUIComponent:
    """Base class for UI components providing common functionality."""
    
    def __init__(self, parent):
        """Initialize the base component."""
        self.parent = parent
        self.frame: Optional[ttk.Frame] = None
    
    def create_frame(self, parent, **kwargs) -> ttk.Frame:
        """Create a frame with common styling."""
        frame = ttk.Frame(parent, **kwargs)
        self.frame = frame
        return frame
    
    def create_label(self, parent, text: str, **kwargs) -> ttk.Label:
        """Create a label with common styling."""
        return ttk.Label(parent, text=text, **kwargs)
    
    def create_button(self, parent, text: str, command, **kwargs) -> ttk.Button:
        """Create a button with common styling."""
        return ttk.Button(parent, text=text, command=command, **kwargs)
    
    def create_entry(self, parent, **kwargs) -> ttk.Entry:
        """Create an entry widget with common styling."""
        return ttk.Entry(parent, **kwargs)
