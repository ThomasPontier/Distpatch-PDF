"""Email preview tab component for reviewing and sending stopover emails."""

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
from typing import List, Dict, Optional
import threading
import os

from models.stopover import Stopover
from services.pdf_attachment_service import PDFAttachmentService
from services.stopover_email_service import StopoverEmailService
from services.email_service import EmailService
from services.mapping_service import MappingService
from services.config_manager import get_config_manager
from .base_component import BaseUIComponent


class EmailPreviewTabComponent(BaseUIComponent):
    """UI component for the email preview tab (unified scrollable layout)."""
    
    def __init__(self, parent, controller=None):
        """Initialize the email preview tab component."""
        super().__init__(parent)
        self.parent = parent
        self.controller = controller
        self.stopovers: List[Stopover] = []
        self.current_pdf_path: Optional[str] = None

        # Services
        self.pdf_attachment_service = PDFAttachmentService()
        self.stopover_email_service = StopoverEmailService() if controller is None else controller.stopover_email_service
        # Mapping is the source of truth for recipients shown in preview
        self.mapping_service = MappingService() if controller is None else controller.mapping_service

        # UI state
        self.stopover_sections: Dict[str, Dict[str, any]] = {}

        # Subscribe to ConfigManager to keep preview reactive
        self._config_manager = get_config_manager()
        # Use simple function references to avoid Qt signal dependency
        self._config_manager.on_templates_changed(lambda _t: self._refresh_all_headers_from_config())
        self._config_manager.on_mappings_changed(lambda _m: self.refresh_recipients_from_configs())
        self._config_manager.on_stopovers_changed(lambda _s: self._refresh_all_headers_from_config())
        self._config_manager.on_last_sent_changed(lambda _l: self._refresh_all_headers_from_config())
        # Create the unified page
        self.create_tab_content()
    
    def create_tab_content(self):
        """Create the unified scrollable page with global controls and per-stopover sections."""
        self.main_frame = self.create_frame(self.parent, padding="10")
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(2, weight=1)

        # Top: Filter + Send all (stretches full width)
        top_frame = ttk.Frame(self.main_frame)
        top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        top_frame.columnconfigure(0, weight=1)

        # Filter: select a specific stopover
        filter_frame = ttk.Frame(top_frame)
        filter_frame.pack(side="left", padx=(0, 8))
        ttk.Label(filter_frame, text="Filter:").pack(side="left")
        self.filter_var = tk.StringVar(value="")
        self.filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, state="readonly", width=20)
        self.filter_combo.pack(side="left", padx=(4, 4))
        self.filter_combo.bind("<<ComboboxSelected>>", lambda e: self._apply_filter())

        # Clear filter button
        self.clear_filter_btn = self.create_button(filter_frame, text="Clear", command=self._clear_filter)
        self.clear_filter_btn.pack(side="left")

        self.send_all_btn = self.create_button(
            top_frame,
            text="Send to All Stopovers",
            command=self._send_all_stopovers_individually
        )
        self.send_all_btn.pack(side="left")
        # PDF file name (not full path)
        self.pdf_name_var = tk.StringVar(value="")
        self.pdf_name_label = ttk.Label(top_frame, textvariable=self.pdf_name_var)
        self.pdf_name_label.pack(side="left", padx=(12, 0))

        # Spacer to ensure full-width stretch
        ttk.Label(top_frame, text=" ").pack(side="left", fill="x", expand=True)

        # Global subject + body editor (stretches to full width)
        template_frame = ttk.LabelFrame(self.main_frame, text="Global Email Subject and Body (applies to all stopovers)", padding="5")
        template_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        template_frame.columnconfigure(1, weight=1)

        # Subject editor row
        ttk.Label(template_frame, text="Subject:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.subject_var = tk.StringVar(value="Stopover Report - {{stopover_code}}")
        self.subject_entry = ttk.Entry(template_frame, textvariable=self.subject_var)
        self.subject_entry.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        self.subject_entry.bind("<KeyRelease>", lambda e: self._save_global_subject_and_overwrite_all())

        # Body editor (increase visible space for editing template)
        self.template_text = tk.Text(template_frame, height=12, wrap="word")
        self.template_text.grid(row=1, column=0, columnspan=2, sticky="nsew")
        # Allow the text area to expand with the window
        try:
            template_frame.rowconfigure(1, weight=1)
        except Exception:
            pass
        self._load_global_subject_and_template()
        self.template_text.bind("<KeyRelease>", lambda e: self._save_template_and_overwrite_all())

        # Scrollable per-stopover sections
        container = ttk.Frame(self.main_frame)
        container.grid(row=2, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)
        # Let the template_frame grow as well to give more space to template editing
        try:
            self.main_frame.rowconfigure(1, weight=1)
        except Exception:
            pass

        self.canvas = tk.Canvas(container, highlightthickness=0)
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)

        self.scrollable = ttk.Frame(self.canvas)
        # Ensure scrollable content uses full available width
        self.scrollable.bind(
            "<Configure>",
            lambda e: (
                self.canvas.configure(scrollregion=self.canvas.bbox("all")),
                self.canvas.itemconfig(self._canvas_window, width=self.canvas.winfo_width())
            )
        )
        self._canvas_window = self.canvas.create_window((0, 0), window=self.scrollable, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.v_scrollbar = vsb

        # Mouse wheel: scroll whole page and trigger lazy loading
        def _on_mousewheel(event):
            delta = 0
            if hasattr(event, "delta") and event.delta != 0:
                # Windows
                delta = -1 * int(event.delta / 120) * 30
            elif hasattr(event, "num"):
                # Linux (Button-4 up / Button-5 down)
                if event.num == 4:
                    delta = -30
                elif event.num == 5:
                    delta = 30
            if delta != 0:
                self.canvas.yview_scroll(int(delta / 30), "units")
                self._lazy_load_visible_previews()
            return "break"

        self.canvas.bind("<Enter>", lambda e: self.canvas.focus_set())
        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.canvas.bind_all("<Button-4>", _on_mousewheel)
        self.canvas.bind_all("<Button-5>", _on_mousewheel)
        self.canvas.bind("<Configure>", lambda e: self._lazy_load_visible_previews())
    
    def set_stopovers(self, stopovers: List[Stopover]):
        """Update the stopovers and rebuild sections."""
        self.stopovers = stopovers
        # Update filter dropdown
        try:
            codes = [s.code for s in self.stopovers]
            self.filter_combo["values"] = [""] + codes  # "" = no filter
        except Exception:
            pass
        self._build_stopover_sections()
    
    def set_pdf_path(self, pdf_path: str):
        """Set the current PDF path and rebuild sections if needed."""
        self.current_pdf_path = pdf_path
        try:
            # Show only the filename in UI
            self.pdf_name_var.set(f"PDF: {os.path.basename(pdf_path)}" if pdf_path else "")
        except Exception:
            self.pdf_name_var.set("")
        if self.stopovers:
            self._build_stopover_sections()

    def _build_stopover_sections(self):
        """Build UI rows for each stopover: left email editor, right PDF preview, send button."""
        for child in self.scrollable.winfo_children():
            child.destroy()
        self.stopover_sections.clear()

        if not self.stopovers:
            ttk.Label(self.scrollable, text="No stopovers found").pack(pady=20)
            # Reflect latest subject/body in the global editors too
            try:
                t = get_config_manager().get_templates()
                self.subject_var.set(t.get("subject") or "Stopover Report - {{stopover_code}}")
                self.template_text.delete("1.0", tk.END)
                self.template_text.insert("1.0", t.get("body") or EmailService()._get_default_template())
            except Exception:
                pass
            return
        # Reflect latest subject/body in the global editors too
        try:
            t = get_config_manager().get_templates()
            self.subject_var.set(t.get("subject") or "Stopover Report - {{stopover_code}}")
            self.template_text.delete("1.0", tk.END)
            self.template_text.insert("1.0", t.get("body") or EmailService()._get_default_template())
        except Exception:
            pass

        for stopover in self.stopovers:
            row = ttk.LabelFrame(self.scrollable, text=f"Stopover {stopover.code}", padding="5")
            row.pack(fill="x", expand=True, pady=6)

            # Two-column layout; stretch to full width
            row.columnconfigure(0, weight=1)
            row.columnconfigure(1, weight=1)

            # Left: Email editor
            left = ttk.Frame(row)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
            email_text = tk.Text(left, height=12, wrap="word")
            email_scroll = ttk.Scrollbar(left, orient="vertical", command=email_text.yview)
            email_text.configure(yscrollcommand=email_scroll.set)
            email_text.grid(row=0, column=0, sticky="nsew")
            email_scroll.grid(row=0, column=1, sticky="ns")
            left.rowconfigure(0, weight=1)
            left.columnconfigure(0, weight=1)

            # Header + send button + status
            header = ttk.Frame(left)
            header.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
            header.columnconfigure(0, weight=1)
            send_btn = self.create_button(header, text="Send Individually", command=lambda c=stopover.code: self._send_single_stopover(c))
            send_btn.pack(side="right")

            # Populate email content from per-stopover config and current mapping
            config = self.stopover_email_service.get_config(stopover.code)
            # Subject is customizable via subject template
            subject = (config.subject_template or "Stopover Report - {{stopover_code}}").replace("{{stopover_code}}", stopover.code)
            body = config.body_template.replace("{{stopover_code}}", stopover.code)

            # To: read live recipients from mapping service to reflect latest configuration instantly
            recipients = self.mapping_service.get_emails_for_stopover(stopover.code)
            to_line = ", ".join(recipients) if recipients else "No email address"

            # Last sent info
            last_sent = self.stopover_email_service.get_last_sent(stopover.code)
            # Format expected input like "2025-07-31T10:22:45Z" or "2025-07-31T10:22:45+02:00" to "dd/MM/yyyy HH:mm"
            def _format_last_sent(val: Optional[str]) -> Optional[str]:
                try:
                    if not val:
                        return None
                    s = str(val).strip()
                    if not s:
                        return None
                    if s.endswith("Z"):
                        s = s[:-1] + "+00:00"
                    from datetime import datetime
                    dt = None
                    # Try with offset
                    try:
                        dt = datetime.fromisoformat(s)
                    except Exception:
                        # Try basic without offset
                        try:
                            dt = datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
                        except Exception:
                            return val
                    return dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    return val
            formatted_last = _format_last_sent(last_sent)
            last_sent_line = f"Dernier envoi : {formatted_last}" if formatted_last else "Dernier envoi : Jamais"

            header_lines = [f"Objet : {subject}", f"À : {to_line}", last_sent_line]
            header_text = "\n".join(header_lines) + "\n\n"
            email_text.insert("1.0", header_text + body)

            # Status indicator for eligibility (PDF page exists + recipient email available)
            status_text = self._compute_status_text(stopover, has_email=(len(recipients) > 0))
            status_label = ttk.Label(header, text=status_text, foreground=("green" if "OK" in status_text else "red"))
            status_label.pack(side="left")

            # Manual persistence bind (debounced) - body only (subject is edited globally)
            def on_edit(evt=None, code=stopover.code, text_widget=email_text):
                if hasattr(text_widget, "_after_id") and text_widget._after_id:
                    try:
                        self.parent.after_cancel(text_widget._after_id)
                    except Exception:
                        pass
                text_widget._after_id = self.parent.after(300, lambda: self._persist_manual_text(code, text_widget))
            email_text.bind("<KeyRelease>", on_edit)

            # Right: PDF preview area (reduced size to leave more room for email content)
            right = ttk.Frame(row)
            right.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
            # Make right column expand to use available space
            try:
                row.columnconfigure(1, weight=1)
            except Exception:
                pass

            preview_label = ttk.Label(right, text="Aperçu PDF affiché ici", anchor="center")
            # Let preview grow with its container both horizontally and vertically
            preview_label.pack(fill="both", expand=True)

            # Track last rendered PIL image for dynamic fit
            # We maintain one image per section to scale on container resize
            self._section_last_image = getattr(self, "_section_last_image", {})
            self._resize_after_ids = getattr(self, "_resize_after_ids", {})

            def _on_resize(event, code=stopover.code):
                # Debounce resize per-section
                try:
                    if code in self._resize_after_ids and self._resize_after_ids[code]:
                        self.parent.after_cancel(self._resize_after_ids[code])
                except Exception:
                    pass
                self._resize_after_ids[code] = self.parent.after(50, lambda c=code: self._fit_and_update_section_preview(c))

            # Bind resize on the right container (better than label for size)
            try:
                right.bind("<Configure>", _on_resize)
            except Exception:
                pass

            self.stopover_sections[stopover.code] = {
                "frame": row,
                "email_text": email_text,
                "pdf_label": preview_label,
                "pdf_loaded": False,
                "status_label": status_label
            }

        # Initial lazy load
        self._lazy_load_visible_previews()
    
    def _persist_manual_text(self, code: str, text_widget: tk.Text):
        """Persist manual edits from a section's editor to the corresponding stopover config."""
        try:
            content = text_widget.get("1.0", tk.END)
            parts = content.split("\n\n", 1)
            body = parts[1] if len(parts) > 1 else content
            config = self.stopover_email_service.get_config(code)
            config.body_template = body.strip()
            self.stopover_email_service.save_config(config)
        except Exception as e:
            print(f"Error persisting manual text for {code}: {e}")

    def _load_global_subject_and_template(self):
        """Load global subject and body template into editors."""
        # Load subject/body from centralized ConfigManager
        try:
            t = get_config_manager().get_templates()
            subject_value = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body_value = t.get("body") or EmailService()._get_default_template()
            self.subject_var.set(subject_value)
            self.template_text.delete("1.0", tk.END)
            self.template_text.insert("1.0", body_value)
        except Exception as e:
            self.subject_var.set("Stopover Report - {{stopover_code}}")
            self.template_text.delete("1.0", tk.END)
            self.template_text.insert("1.0", f"Error loading template from ConfigManager: {e}")

    def _save_global_subject_and_overwrite_all(self):
        """Save global subject template and update all stopovers' headers immediately."""
        subject_template = self.subject_var.get() or "Stopover Report - {{stopover_code}}"
        try:
            # Persist template subject to ConfigManager once (applies to all stopovers)
            t = get_config_manager().get_templates()
            body_current = t.get("body") or self.template_text.get("1.0", tk.END)
            get_config_manager().set_templates(subject_template, body_current)
            # Refresh headers for all sections
            self._refresh_all_headers_from_config()
        except Exception as e:
            print(f"Error saving global subject: {e}")

    def _save_template_and_overwrite_all(self):
        """Save global body template and overwrite all stopovers' email bodies immediately."""
        try:
            content = self.template_text.get("1.0", tk.END)
            subject_template = self.subject_var.get() or "Stopover Report - {{stopover_code}}"
            # Persist to centralized templates once
            get_config_manager().set_templates(subject_template, content)
            # Refresh headers and email bodies in UI (body content shown is from template)
            self._refresh_all_headers_from_config()
        except Exception as e:
            print(f"Error saving template to ConfigManager: {e}")

    def _lazy_load_visible_previews(self):
        """Lazy load PDF previews for sections currently visible (idempotent)."""
        if not self.current_pdf_path or not os.path.exists(self.current_pdf_path):
            return
        try:
            for code, section in self.stopover_sections.items():
                if section["pdf_loaded"]:
                    continue
                # Always attempt to load previews when invoked; idempotent and robust
                self._load_pdf_preview(code)
        except Exception as e:
            print(f"Lazy-load error: {e}")

    def _compute_status_text(self, stopover: Stopover, has_email: bool) -> str:
        """Return status string for a stopover: OK or reason missing."""
        page_ok = True
        try:
            # Fast check: ensure PDF path exists; deeper render errors handled in preview
            if not self.current_pdf_path or not os.path.exists(self.current_pdf_path):
                page_ok = False
        except Exception:
            page_ok = False

        reasons = []
        if not page_ok:
            reasons.append("PDF page missing")
        if not has_email:
            reasons.append("No recipient")

        return "Status: OK" if not reasons else f"Status: {'; '.join(reasons)}"

    def clear(self):
        """Clear all sections and reset state."""
        try:
            for child in self.scrollable.winfo_children():
                child.destroy()
        except Exception:
            pass
        self.stopovers = []
        self.current_pdf_path = None
        self.stopover_sections.clear()
        ttk.Label(self.scrollable, text="No stopovers found").pack(pady=20)

    def _load_pdf_preview(self, code: str):
        """Render and set the PDF preview image for a stopover (reuse StopoverPage logic: 1200x800, KeepAspect)."""
        try:
            # Find stopover by code
            target = next((s for s in self.stopovers if s.code == code), None)
            if not target:
                return
            # Use smaller renderer target to keep preview compact in this tab
            from core.pdf_renderer import PDFRenderer
            renderer = PDFRenderer(self.current_pdf_path)
            # Render at high resolution; actual fit done dynamically
            img = renderer.get_page_image(target.page_number, max_width=1600, max_height=1600)
            renderer.close()

            # Store original PIL image to allow dynamic scaling on resize
            self._section_last_image[code] = img
            section = self.stopover_sections.get(code)
            if section:
                section["pdf_loaded"] = True
                # Initial fit to container
                self._fit_and_update_section_preview(code)
        except Exception as e:
            section = self.stopover_sections.get(code)
            if section:
                section["pdf_label"].configure(text=f"Error loading PDF preview: {e}")

    def _send_single_stopover(self, code: str):
        """Send email for a single stopover using current config and generated attachment."""
        if not self.current_pdf_path:
            messagebox.showerror("Error", "No PDF file selected.")
            return
        try:
            stopover = next((s for s in self.stopovers if s.code == code), None)
            if not stopover:
                messagebox.showerror("Error", f"Stopover {code} not found.")
                return
            # Use mapping as source of truth for recipients
            recipients = self.mapping_service.get_emails_for_stopover(code)
            if not recipients:
                messagebox.showerror("Error", f"No recipients configured for {code}.")
                return
            config = self.stopover_email_service.get_config(code)
            attachment_path = PDFAttachmentService.create_stopover_attachment(self.current_pdf_path, stopover)
            if not attachment_path:
                messagebox.showerror("Error", f"Failed to create attachment for {code}.")
                return
            subject = (config.subject_template or "Stopover Report - {{stopover_code}}").replace("{{stopover_code}}", code)
            body = config.body_template.replace("{{stopover_code}}", code)
            success = False
            if self.controller:
                success = self.controller.email_service.send_email(
                    to_emails=recipients,
                    subject=subject,
                    body=body,
                    attachment_path=attachment_path,
                    cc_emails=config.cc_recipients,
                    bcc_emails=config.bcc_recipients
                )
            PDFAttachmentService.cleanup_temp_attachments([attachment_path])
            if success:
                try:
                    # Persist last sent date for this stopover
                    self.stopover_email_service.set_last_sent_now(code)
                except Exception as _e:
                    # Non-fatal persistence error
                    print(f"Warning: failed to persist last_sent_at for {code}: {_e}")
                messagebox.showinfo("Success", f"Email sent for {code}")
            else:
                messagebox.showerror("Error", f"Failed to send email for {code}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email for {code}: {e}")

    def _fit_and_update_section_preview(self, code: str):
        """Scale stored PIL image for a section to fit its right container while keeping aspect ratio."""
        try:
            section = self.stopover_sections.get(code)
            if not section:
                return
            img = self._section_last_image.get(code)
            if img is None:
                return
            # Available size from the parent container of the label (better reflects space)
            label = section["pdf_label"]
            parent = label.master
            avail_w = max(1, parent.winfo_width() - 8)
            avail_h = max(1, parent.winfo_height() - 8)
            iw, ih = img.size
            if iw <= 0 or ih <= 0:
                return
            scale = min(avail_w / iw, avail_h / ih)
            target_w = max(1, int(iw * scale))
            target_h = max(1, int(ih * scale))
            if target_w <= 0 or target_h <= 0:
                return
            from PIL import Image
            resized = img.resize((target_w, target_h), Image.LANCZOS)
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(resized)
            label.configure(image=photo, text="")
            label.image = photo
        except Exception:
            try:
                section = self.stopover_sections.get(code)
                if section:
                    section["pdf_label"].configure(image="", text="Aperçu PDF indisponible")
            except Exception:
                pass

    # Public API to refresh recipients from external mapping changes
    def refresh_recipients_from_configs(self):
        """Re-read latest mapping + per-stopover templates and update headers + status in-place."""
        for stopover in self.stopovers:
            section = self.stopover_sections.get(stopover.code)
            if not section:
                continue
            cfg = self.stopover_email_service.get_config(stopover.code)
            recipients = self.mapping_service.get_emails_for_stopover(stopover.code)
            subject = (cfg.subject_template or "Stopover Report - {{stopover_code}}").replace("{{stopover_code}}", stopover.code)
            to_line = ", ".join(recipients) if recipients else "No email address"
            t = get_config_manager().get_templates()
            subject_template = t.get("subject") or "Stopover Report - {{stopover_code}}"
            subject = subject_template.replace("{{stopover_code}}", stopover.code)
            header_lines = [f"Subject: {subject}", f"To: {to_line}"]
            header_text = "\n".join(header_lines) + "\n\n"

            text_widget = section["email_text"]
            current = text_widget.get("1.0", tk.END)
            parts = current.split("\n\n", 1)
            body = parts[1] if len(parts) > 1 else ""
            text_widget.delete("1.0", tk.END)
            text_widget.insert("1.0", header_text + body)

            # Update status label
            status_text = self._compute_status_text(stopover, has_email=(len(recipients) > 0))
            lbl = section.get("status_label")
            if lbl:
                lbl.configure(text=status_text, foreground=("green" if "OK" in status_text else "red"))

    def _send_all_stopovers_individually(self):
        """Send emails for all stopovers sequentially, and summarize failures."""
        if not self.current_pdf_path:
            messagebox.showerror("Error", "No PDF file selected.")
            return
        # Confirm
        if not messagebox.askyesno("Confirm Send", f"Send emails for {len(self.stopovers)} stopover(s)?"):
            return

        # Background thread
        def worker():
            success = 0
            total = len(self.stopovers)
            failures = []  # list of dicts: {code, reasons: [..]}
            attachments = []
            try:
                for s in self.stopovers:
                    reasons = []
                    try:
                        cfg = self.stopover_email_service.get_config(s.code)
                        recipients = self.mapping_service.get_emails_for_stopover(s.code)
                        if not recipients:
                            reasons.append("missing email address")
                        attachment_path = None
                        if not reasons:
                            attachment_path = PDFAttachmentService.create_stopover_attachment(self.current_pdf_path, s)
                            if not attachment_path:
                                reasons.append("missing PDF page or failed to create attachment")
                        if reasons:
                            failures.append({"code": s.code, "reasons": reasons})
                            continue
                        attachments.append(attachment_path)
                        # Always read latest subject/body from ConfigManager (single source of truth)
                        t = get_config_manager().get_templates()
                        subject = (t.get("subject") or "Stopover Report - {{stopover_code}}").replace("{{stopover_code}}", s.code)
                        body = (t.get("body") or EmailService()._get_default_template()).replace("{{stopover_code}}", s.code)
                        sent = False
                        if self.controller:
                            try:
                                sent = self.controller.email_service.send_email(
                                    to_emails=recipients,
                                    subject=subject,
                                    body=body,
                                    attachment_path=attachment_path,
                                    cc_emails=cfg.cc_recipients,
                                    bcc_emails=cfg.bcc_recipients
                                )
                            except Exception as ex:
                                reasons.append(f"Outlook error: {ex}")
                                sent = False
                        if sent:
                            success += 1
                            try:
                                # Persist last sent date for this stopover
                                self.stopover_email_service.set_last_sent_now(s.code)
                            except Exception as _e:
                                print(f"Warning: failed to persist last_sent_at for {s.code}: {_e}")
                        else:
                            if not reasons:
                                reasons.append("send failed")
                            failures.append({"code": s.code, "reasons": reasons})
                    except Exception as e:
                        failures.append({"code": s.code, "reasons": [f"unexpected error: {e}"]})
            finally:
                PDFAttachmentService.cleanup_temp_attachments(attachments)
                def _show_result():
                    # Always show completion
                    messagebox.showinfo("Send Complete", f"Sent {success}/{total} email(s).")
                    # If there are failures, show a detailed summary
                    if failures:
                        lines = []
                        for f in failures:
                            lines.append(f"{f['code']}: {', '.join(f['reasons'])}")
                        messagebox.showwarning("Failed Stopovers", "The following stopovers failed:\n\n" + "\n".join(lines))
                self.parent.after(0, _show_result)

        t = threading.Thread(target=worker, daemon=True)
        t.start()
    
    def _refresh_all_headers_from_config(self):
        """Recompute and update headers/body preview for all sections from ConfigManager."""
        try:
            t = get_config_manager().get_templates()
            subject_template = t.get("subject") or "Stopover Report - {{stopover_code}}"
            body_template = t.get("body") or EmailService()._get_default_template()
            for stopover in self.stopovers:
                section = self.stopover_sections.get(stopover.code)
                if not section:
                    continue
                recipients = self.mapping_service.get_emails_for_stopover(stopover.code)
                subject = subject_template.replace("{{stopover_code}}", stopover.code)
                to_line = ", ".join(recipients) if recipients else "Aucune adresse email"
                last_sent = self.stopover_email_service.get_last_sent(stopover.code)
                def _format_last_sent(val: Optional[str]) -> Optional[str]:
                    try:
                        if not val:
                            return None
                        s = str(val).strip()
                        if not s:
                            return None
                        if s.endswith("Z"):
                            s = s[:-1] + "+00:00"
                        from datetime import datetime
                        dt = None
                        try:
                            dt = datetime.fromisoformat(s)
                        except Exception:
                            try:
                                dt = datetime.strptime(s, "%Y-%m-%dT%H:%M:%S")
                            except Exception:
                                return val
                        return dt.strftime("%d/%m/%Y %H:%M")
                    except Exception:
                        return val
                formatted_last = _format_last_sent(last_sent)
                last_sent_line = f"Dernier envoi : {formatted_last}" if formatted_last else "Dernier envoi : Jamais"
                header_text = f"Objet : {subject}\nÀ : {to_line}\n{last_sent_line}\n\n"
                text_widget = section["email_text"]
                current = text_widget.get("1.0", tk.END)
                parts = current.split("\n\n", 1)
                body_current = parts[1] if len(parts) > 1 else ""
                # Replace body with template-rendered content to keep in sync
                rendered_body = body_template.replace("{{stopover_code}}", stopover.code)
                text_widget.delete("1.0", tk.END)
                text_widget.insert("1.0", header_text + rendered_body)
                # Update status label color/text
                status_text = self._compute_status_text(stopover, has_email=(len(recipients) > 0))
                lbl = section.get("status_label")
                if lbl:
                    lbl.configure(text=status_text, foreground=("green" if "OK" in status_text else "red"))
        except Exception as e:
            print(f"Error refreshing headers from ConfigManager: {e}")

    # Filtering helpers
    def _apply_filter(self):
        """Show only the selected stopover code if filter is set."""
        code = (self.filter_var.get() or "").strip()
        for s_code, section in self.stopover_sections.items():
            visible = (code == "" or s_code == code)
            try:
                if visible:
                    section["frame"].pack(fill="x", expand=True, pady=6)
                else:
                    section["frame"].pack_forget()
            except Exception:
                pass
        # Scroll back to top automatically
        try:
            self.canvas.yview_moveto(0.0)
            if hasattr(self, "v_scrollbar") and self.v_scrollbar:
                self.v_scrollbar.set(0.0, 0.1)
        except Exception:
            pass
        # Ensure previews load for the filtered result and refresh header from current configs instantly
        self._lazy_load_visible_previews()
        self.refresh_recipients_from_configs()
        self._refresh_all_headers_from_config()

    def _clear_filter(self):
        """Reset filter and show all stopovers."""
        try:
            self.filter_var.set("")
            self._apply_filter()
        except Exception:
            pass
