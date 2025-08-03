"""Service for creating PDF attachments for individual stopovers."""

import os
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
from models.stopover import Stopover


class PDFAttachmentService:
    """Service for creating PDF attachments with proper naming conventions."""

    @staticmethod
    def create_stopover_attachment(
        pdf_path: str, stopover: Stopover, output_dir: str = None
    ) -> Optional[str]:
        """
        Create a PDF attachment for a specific stopover page.

        Args:
            pdf_path: Path to the source PDF file
            stopover: Stopover object containing code and page number
            output_dir: Directory to save the attachment (defaults to temp directory)

        Returns:
            Path to the created attachment file, or None if failed
        """
        try:
            # Create output directory if specified
            if output_dir:
                Path(output_dir).mkdir(parents=True, exist_ok=True)
                output_path = Path(output_dir)
            else:
                # Use system temp directory
                output_path = Path(os.environ.get("TEMP", "."))

            # Generate filename with required format
            filename = f"Bilan_Enquete_Satisfaction_{stopover.code}.pdf"
            attachment_path = output_path / filename

            # Extract the specific page from the source PDF
            with fitz.open(pdf_path) as source_doc:
                if 1 <= stopover.page_number <= len(source_doc):
                    # Create new PDF with just this page
                    new_doc = fitz.open()
                    new_doc.insert_pdf(
                        source_doc,
                        from_page=stopover.page_number - 1,
                        to_page=stopover.page_number - 1,
                    )
                    new_doc.save(str(attachment_path))
                    new_doc.close()
                    return str(attachment_path)

            return None

        except Exception as e:
            # Keep: debug print for failure transparency
            print(f"Error creating PDF attachment for {stopover.code}: {e}")
            return None

    @staticmethod
    def create_attachment_filename(stopover_code: str) -> str:
        """
        Generate the standard attachment filename for a stopover.

        Args:
            stopover_code: 3-letter stopover code

        Returns:
            Formatted filename string
        """
        return f"Bilan_Enquete_Satisfaction_{stopover_code}.pdf"

    @staticmethod
    def cleanup_temp_attachments(attachment_paths: list):
        """
        Clean up temporary attachment files.

        Args:
            attachment_paths: List of file paths to delete
        """
        for path in attachment_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                # Keep: warning log on failure
                print(f"Warning: Failed to delete temporary file {path}: {e}")
