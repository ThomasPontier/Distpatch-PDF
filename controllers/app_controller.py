"""Application controller to coordinate between UI and services."""

import threading
from typing import List, Set, Optional, Callable, Dict
from core.pdf_processor import PDFProcessor
from core.pdf_renderer import PDFRenderer
from models.stopover import Stopover
from services.email_service import EmailService
from services.mapping_service import MappingService
from services.stopover_email_service import StopoverEmailService, StopoverEmailConfig
from utils.file_utils import validate_pdf_file


class AppController:
    """Controller class to manage application logic and coordinate components."""
    
    def __init__(self):
        """Initialize the application controller."""
        # Initialize services
        self.pdf_processor = PDFProcessor()
        self.email_service = EmailService()
        self.mapping_service = MappingService()
        self.stopover_email_service = StopoverEmailService()
        
        # State variables
        self.current_pdf_path: Optional[str] = None
        self.stopovers: List[Stopover] = []
        self.found_stopover_codes: Set[str] = set()
        self.pdf_renderer: Optional[PDFRenderer] = None
        self.outlook_connected = False
        self.outlook_user: Optional[str] = None
        
        # Callbacks for UI updates
        self.on_status_update: Optional[Callable[[str], None]] = None
        self.on_progress_start: Optional[Callable[[], None]] = None
        self.on_progress_stop: Optional[Callable[[], None]] = None
        self.on_analysis_complete: Optional[Callable[[List[Stopover]], None]] = None
        self.on_outlook_connection_change: Optional[Callable[[bool, Optional[str]], None]] = None
    
    def set_pdf_path(self, pdf_path: str) -> bool:
        """Set the current PDF path and validate it."""
        if validate_pdf_file(pdf_path):
            self.current_pdf_path = pdf_path
            self.stopovers.clear()
            self.found_stopover_codes.clear()
            self.close_pdf_renderer()
            return True
        return False
    
    def analyze_pdf(self) -> bool:
        """Start PDF analysis in a background thread."""
        if not self.current_pdf_path:
            return False
        
        # Start analysis in a separate thread to keep GUI responsive
        if self.on_progress_start:
            self.on_progress_start()
        
        thread = threading.Thread(target=self._analyze_pdf_thread)
        thread.daemon = True
        thread.start()
        return True
    
    def _analyze_pdf_thread(self):
        """Thread function for PDF analysis."""
        try:
            self.stopovers = self.pdf_processor.analyze_pdf(self.current_pdf_path)
            
            # Extract stopover codes from found stopovers
            self.found_stopover_codes = {stopover.code for stopover in self.stopovers}
            
            # Update UI in the main thread
            if self.on_status_update:
                self.on_status_update(f"Escales détectées : {len(self.stopovers)}")
            
            if self.on_analysis_complete:
                self.on_analysis_complete(self.stopovers)
                
        except Exception as e:
            error_msg = f"Erreur d’analyse du PDF : {str(e)}"
            if self.on_status_update:
                self.on_status_update(error_msg)
        finally:
            if self.on_progress_stop:
                self.on_progress_stop()
    
    def load_page_preview(self, stopover: Stopover, progress_callback: Optional[Callable[[str], None]] = None) -> bool:
        """Load and display page preview for a stopover."""
        if not self.current_pdf_path:
            return False
        
        try:
            # Show progress if callback provided
            if progress_callback:
                progress_callback("Chargement de l’aperçu de la page…")
            
            # Create or update PDF renderer
            if not self.pdf_renderer or self.pdf_renderer.pdf_path != self.current_pdf_path:
                self.close_pdf_renderer()
                self.pdf_renderer = PDFRenderer(self.current_pdf_path)
            
            # Get page image
            img = self.pdf_renderer.get_page_image(stopover.page_number)
            
            return True
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"Erreur de chargement de l’aperçu de la page : {str(e)}")
            return False
    
    def check_outlook_connection(self) -> bool:
        """Check if Outlook is available and try to connect."""
        if self.email_service.is_outlook_available():
            return self.connect_to_outlook()
        return False
    
    def connect_to_outlook(self) -> bool:
        """Connect to Outlook and update state."""
        try:
            if self.email_service.connect_to_outlook():
                self.outlook_connected = True
                self.outlook_user = self.email_service.get_current_user()
                
                # Notify UI of connection change
                if self.on_outlook_connection_change:
                    self.on_outlook_connection_change(True, self.outlook_user)
                
                return True
            else:
                self.outlook_connected = False
                self.outlook_user = None
                
                # Notify UI of connection change
                if self.on_outlook_connection_change:
                    self.on_outlook_connection_change(False, None)
                
                return False
        except Exception as e:
            print(f"Error connecting to Outlook: {e}")
            self.outlook_connected = False
            self.outlook_user = None
            
            # Notify UI of connection change
            if self.on_outlook_connection_change:
                self.on_outlook_connection_change(False, None)
            
            return False
    
    def disconnect_from_outlook(self):
        """Disconnect from Outlook and update state."""
        self.email_service.disconnect_from_outlook()
        self.outlook_connected = False
        self.outlook_user = None
        
        # Notify UI of connection change
        if self.on_outlook_connection_change:
            self.on_outlook_connection_change(False, None)
    
    def toggle_outlook_connection(self) -> bool:
        """Toggle Outlook connection state."""
        if self.outlook_connected:
            self.disconnect_from_outlook()
            return False
        else:
            return self.connect_to_outlook()
    
    def get_outlook_status(self) -> tuple[bool, Optional[str]]:
        """Get current Outlook connection status."""
        return self.outlook_connected, self.outlook_user
    
    def get_mapped_stopovers(self) -> List[str]:
        """Get all stopover codes that have email mappings."""
        return self.mapping_service.get_mapped_stopovers()
    
    def has_mapping(self, stopover_code: str) -> bool:
        """Check if a stopover code has any email mappings."""
        return self.mapping_service.has_mapping(stopover_code)
    
    def get_emails_for_stopover(self, stopover_code: str) -> List[str]:
        """Get email addresses for a stopover code."""
        return self.mapping_service.get_emails_for_stopover(stopover_code)
    
    def add_mapping(self, stopover_code: str, email: str) -> bool:
        """Add a new email mapping for a stopover code."""
        return self.mapping_service.add_mapping(stopover_code, email)
    
    def remove_mapping(self, stopover_code: str, email: str) -> bool:
        """Remove an email mapping for a stopover code."""
        return self.mapping_service.remove_mapping(stopover_code, email)
    
    def update_mappings(self, new_mappings: dict) -> bool:
        """Update multiple mappings at once."""
        self.mapping_service.update_mappings(new_mappings)
        return True
    
    def get_all_mappings(self) -> dict:
        """Get all stopover-to-email mappings."""
        return self.mapping_service.get_all_mappings()
    
    def send_stopover_emails(self, stopovers: List[Stopover], pdf_path: str) -> tuple[int, int]:
        """
        Send emails for selected stopovers.
        
        Returns:
            tuple: (success_count, total_count)
        """
        success_count = 0
        total_count = len(stopovers)
        
        try:
            for stopover in stopovers:
                # Get emails for this stopover
                emails = self.get_emails_for_stopover(stopover.code)
                
                if emails:
                    # Send email
                    success = self.email_service.send_stopover_email(
                        stopover.code,
                        emails,
                        pdf_path
                    )
                    
                    if success:
                        success_count += 1
                        print(f"Email envoyé pour {stopover.code}")
                    else:
                        print(f"Échec d’envoi pour {stopover.code}")
            
            return success_count, total_count
            
        except Exception as e:
            print(f"Erreur lors de l’envoi des emails : {e}")
            return success_count, total_count
    
    def set_current_account(self, account_id: str, account_info: dict):
        """Set the current email account."""
        self.email_service.set_current_account(account_id, account_info)
    
    def get_current_account_name(self) -> str:
        """Get the current account name."""
        return self.email_service.get_current_account_name()
    
    def get_current_email_address(self) -> Optional[str]:
        """Get the current email address."""
        return self.email_service.get_current_email_address()
    
    def close_pdf_renderer(self):
        """Close the PDF renderer if it's open."""
        if self.pdf_renderer:
            self.pdf_renderer.close()
            self.pdf_renderer = None
    
    def clear_state(self):
        """Clear all application state."""
        self.current_pdf_path = None
        self.stopovers.clear()
        self.found_stopover_codes.clear()
        self.close_pdf_renderer()
    
    def destroy(self):
        """Clean up resources when destroying the controller."""
        self.close_pdf_renderer()
        self.disconnect_from_outlook()
    
    def get_stopover_email_config(self, stopover_code: str) -> StopoverEmailConfig:
        """Get email configuration for a stopover."""
        return self.stopover_email_service.get_config(stopover_code)
    
    def save_stopover_email_config(self, config: StopoverEmailConfig) -> bool:
        """Save email configuration for a stopover."""
        return self.stopover_email_service.save_config(config)
    
    def get_all_stopover_email_configs(self) -> Dict[str, StopoverEmailConfig]:
        """Get all stopover email configurations."""
        return self.stopover_email_service.get_all_configs()
    
    def delete_stopover_email_config(self, stopover_code: str) -> bool:
        """Delete email configuration for a stopover."""
        return self.stopover_email_service.delete_config(stopover_code)
    
    def stopover_email_config_exists(self, stopover_code: str) -> bool:
        """Check if email configuration exists for a stopover."""
        return self.stopover_email_service.config_exists(stopover_code)
    
    def get_enabled_stopover_email_configs(self) -> Dict[str, StopoverEmailConfig]:
        """Get all enabled stopover email configurations."""
        return self.stopover_email_service.get_enabled_configs()
