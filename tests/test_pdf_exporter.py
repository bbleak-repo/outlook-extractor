"""Comprehensive tests for the PDF exporter module.

This module contains tests for the PDFExporter class, including:
- Basic PDF generation
- Formatting and styling
- Handling of different content types
- Performance with various data sizes
- Error handling and validation
"""
import os
import sys
import tempfile
import pytest
import io
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Any

from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter

from outlook_extractor.export.pdf_exporter import PDFExporter
from outlook_extractor.export.constants import EXPORT_FIELDS_V1

# Re-export fixtures from conftest for better IDE support
from tests.conftest import (
    sample_emails, empty_emails, invalid_emails,
    pdf_exporter, create_test_emails
)

class TestPDFExporter:
    """Comprehensive test suite for the PDFExporter class."""
    
    # --- Test Data ---
    
    @pytest.fixture
    def special_chars_email(self):
        """Email with special characters that might cause encoding issues."""
        return {
            'id': 'special-1',
            'conversation_id': 'conv-special',
            'subject': 'Email with special chars: 日本語, русский, 中文, ελληνικά',
            'sender_name': 'Tést Sènder',
            'sender_email': 'test.sender@example.com',
            'to_recipients': 'recipient@example.com',
            'cc_recipients': '',
            'bcc_recipients': '',
            'received_time': datetime.now() - timedelta(days=1),
            'sent_time': datetime.now() - timedelta(days=2),
            'categories': 'Test;Unicode;Special',
            'importance': 'Normal',
            'sensitivity': 0,
            'has_attachments': False,
            'is_read': True,
            'is_flagged': False,
            'is_priority': False,
            'is_admin': False,
            'body': 'This email contains special characters:\n• Bullet points\n• And emojis 😊\n• And quotes: "Hello"',
            'html_body': '<p>This email contains special characters:<br>• Bullet points<br>• And emojis 😊<br>• And quotes: &quot;Hello&quot;</p>',
            'folder_path': 'Inbox/Test',
            'thread_id': 'thread-special',
            'thread_depth': 0,
            'size': 1024
        }
    
    # --- Basic Export Tests ---
    
    @pytest.mark.parametrize("email_count", [1, 3, 5])
    def test_export_various_sizes(self, email_count, tmp_path, pdf_exporter):
        """Test export with different numbers of emails."""
        # Setup
        output_file = tmp_path / f"test_export_{email_count}.pdf"
        emails = create_test_emails(email_count)
        
        # Test
        result = pdf_exporter.export_emails(
            emails=emails,
            output_path=output_file
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        # Verify PDF structure
        self._verify_pdf_structure(output_file, email_count)
    
    def test_export_with_special_chars(self, special_chars_email, tmp_path, pdf_exporter):
        """Test export with special characters in email content."""
        # Setup
        output_file = tmp_path / "special_chars.pdf"
        
        # Test
        result = pdf_exporter.export_emails(
            emails=[special_chars_email],
            output_path=output_file
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        # Verify PDF structure
        self._verify_pdf_structure(output_file, 1)
    
    # --- Formatting Tests ---
    
    def test_export_with_custom_styles(self, sample_emails, tmp_path):
        """Test export with custom styles."""
        # Setup
        output_file = tmp_path / "custom_styles.pdf"
        
        # Create custom styles
        styles = {
            'title': {'fontName': 'Helvetica-Bold', 'fontSize': 16, 'textColor': (0, 0, 0.8)},
            'header': {'fontName': 'Helvetica-Bold', 'fontSize': 12, 'textColor': (0.2, 0.2, 0.2)},
            'normal': {'fontName': 'Helvetica', 'fontSize': 10},
            'footer': {'fontName': 'Helvetica-Oblique', 'fontSize': 8, 'textColor': (0.4, 0.4, 0.4)}
        }
        
        pdf_exporter = PDFExporter(
            page_size=letter,
            margins=(72, 72, 72, 72),  # 1 inch margins
            styles=styles,
            title="Custom Styled Export"
        )
        
        # Test
        result = pdf_exporter.export_emails(
            emails=sample_emails[:2],  # Just test with 2 emails
            output_path=output_file
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
    
    # --- Error Handling Tests ---
    
    def test_export_empty_list(self, empty_emails, tmp_path, pdf_exporter):
        """Test export with empty email list."""
        output_file = tmp_path / "empty_export.pdf"
        
        with pytest.raises(ValueError, match="No emails to export"):
            pdf_exporter.export_emails(
                emails=empty_emails,
                output_path=output_file
            )
        
        # File should not be created
        assert not output_file.exists()
    
    def test_export_invalid_emails(self, invalid_emails, tmp_path, pdf_exporter):
        """Test export with invalid email data."""
        output_file = tmp_path / "invalid_export.pdf"
        
        with pytest.raises((ValueError, TypeError, AttributeError)):
            pdf_exporter.export_emails(
                emails=invalid_emails,
                output_path=output_file
            )
        
        # File should not be created
        assert not output_file.exists()
    
    # --- Performance Tests ---
    
    @pytest.mark.parametrize("email_count", [5, 10, 20])  # Keep numbers small for PDF generation speed
    def test_export_performance(self, email_count, tmp_path, pdf_exporter, benchmark):
        """Test export performance with varying numbers of emails."""
        # Setup - create test emails
        emails = create_test_emails(email_count)
        output_file = tmp_path / f"perf_test_{email_count}.pdf"
        
        # Benchmark the export
        def _run_export():
            return pdf_exporter.export_emails(
                emails=emails,
                output_path=output_file
            )
        
        # Run benchmark
        result = benchmark(_run_export)
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
    
    # --- File Operation Tests ---
    
    def test_export_to_nonexistent_directory(self, sample_emails, tmp_path, pdf_exporter):
        """Test export to a non-existent directory."""
        output_file = tmp_path / "nonexistent" / "test_export.pdf"
        
        result = pdf_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file
        )
        
        assert result is True
        assert output_file.exists()
        assert output_file.parent.exists()
    
    def test_export_with_invalid_path(self, sample_emails, pdf_exporter):
        """Test export with invalid file path."""
        # Test with None path
        with pytest.raises(ValueError):
            pdf_exporter.export_emails(
                emails=sample_emails,
                output_path=None
            )
        
        # Test with directory path
        with tempfile.TemporaryDirectory() as temp_dir:
            with pytest.raises(IsADirectoryError):
                pdf_exporter.export_emails(
                    emails=sample_emails,
                    output_path=Path(temp_dir)
                )
    
    # --- Helper Methods ---
    
    def _verify_pdf_structure(self, pdf_path: Path, expected_email_count: int):
        """Verify the basic structure of the generated PDF."""
        # Check file exists and has content
        assert pdf_path.exists()
        assert pdf_path.stat().st_size > 0
        
        # Basic PDF validation using PyPDF2
        try:
            with open(pdf_path, 'rb') as f:
                reader = PdfReader(f)
                
                # Check that PDF is not empty
                assert len(reader.pages) > 0
                
                # For PDFs with multiple emails, there should be at least one page per email
                # (though some emails might span multiple pages)
                assert len(reader.pages) >= expected_email_count
                
                # Check that the PDF contains some expected text
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                
                # Check for some common elements
                assert "Subject:" in text
                assert "From:" in text
                assert "Date:" in text
                
        except Exception as e:
            pytest.fail(f"Failed to parse PDF: {e}")


if __name__ == "__main__":
    pytest.main(["-v", "test_pdf_exporter.py"])
