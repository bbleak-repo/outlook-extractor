"""Tests for PDF export functionality."""
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

import pandas as pd
from reportlab.lib.pagesizes import letter

from outlook_extractor.export.pdf_exporter import PDFExporter

class TestPDFExporter(unittest.TestCase):
    """Test cases for PDFExporter class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.exporter = PDFExporter()
        self.test_dir = tempfile.mkdtemp()
        
        # Create test emails
        self.test_emails = [
            {
                'id': 'email_1',
                'subject': 'Test Email 1',
                'sender': 'sender1@example.com',
                'to': 'recipient1@example.com',
                'received_date': '2025-07-15T10:30:00-04:00',
                'body': 'This is a test email body.\nWith multiple lines.',
                'has_attachments': False,
                'importance': 'Normal',
                'is_read': True,
                'categories': ['Test', 'Category1']
            },
            {
                'id': 'email_2',
                'subject': 'Test Email 2',
                'sender': 'sender2@example.com',
                'to': 'recipient2@example.com',
                'cc': 'cc@example.com',
                'received_date': '2025-07-16T14:30:00-04:00',
                'body': 'Another test email with <b>HTML</b> content.',
                'has_attachments': True,
                'attachments': [
                    {'name': 'document.pdf', 'size': 1024},
                    {'name': 'image.jpg', 'size': 2048}
                ],
                'importance': 'High',
                'is_read': False,
                'categories': ['Important']
            }
        ]
    
    def tearDown(self):
        """Clean up after tests."""
        # Clean up test files
        for file_path in Path(self.test_dir).glob('test_*.pdf'):
            try:
                file_path.unlink()
            except OSError:
                pass
    
    def test_export_single_email(self):
        """Test exporting a single email to PDF."""
        output_path = Path(self.test_dir) / 'test_single.pdf'
        success, message = self.exporter.export_emails(
            emails=[self.test_emails[0]],
            output_path=output_path
        )
        
        self.assertTrue(success)
        self.assertTrue(output_path.exists())
        self.assertGreater(output_path.stat().st_size, 1000)  # Reasonable file size
    
    def test_export_multiple_emails(self):
        """Test exporting multiple emails to a single PDF."""
        output_path = Path(self.test_dir) / 'test_multiple.pdf'
        success, message = self.exporter.export_emails(
            emails=self.test_emails,
            output_path=output_path,
            include_summary=True
        )
        
        self.assertTrue(success)
        self.assertTrue(output_path.exists())
        self.assertGreater(output_path.stat().st_size, 2000)  # Should be larger than single
    
    def test_export_to_buffer(self):
        """Test exporting to an in-memory buffer."""
        pdf_data = self.exporter.export_to_buffer(emails=self.test_emails)
        self.assertIsInstance(pdf_data, bytes)
        self.assertGreater(len(pdf_data), 1000)  # Reasonable data size
    
    def test_export_with_missing_fields(self):
        """Test exporting emails with missing fields."""
        incomplete_email = {
            'subject': 'Incomplete Email',
            'body': 'This email is missing some fields'
        }
        
        output_path = Path(self.test_dir) / 'test_incomplete.pdf'
        success, message = self.exporter.export_emails(
            emails=[incomplete_email],
            output_path=output_path
        )
        
        self.assertTrue(success)
        self.assertTrue(output_path.exists())
    
    @patch('reportlab.platypus.SimpleDocTemplate.build')
    def test_export_cancellation(self, mock_build):
        """Test export cancellation."""
        from threading import Event
        
        # Create a cancellation event and set it immediately
        cancel_event = Event()
        cancel_event.set()
        
        output_path = Path(self.test_dir) / 'test_cancelled.pdf'
        success, message = self.exporter.export_emails(
            emails=self.test_emails,
            output_path=output_path,
            cancel_event=cancel_event
        )
        
        self.assertFalse(success)
        self.assertIn('cancelled', message.lower())
        mock_build.assert_not_called()
    
    def test_format_field_value(self):
        """Test the _format_field_value method."""
        # Test datetime formatting
        dt_str = '2025-07-15T10:30:00-04:00'
        formatted = self.exporter._format_field_value('received_date', dt_str)
        self.assertIn('2025-07-15 10:30:00', formatted)
        
        # Test boolean formatting
        self.assertEqual(self.exporter._format_field_value('is_read', True), 'Yes')
        self.assertEqual(self.exporter._format_field_value('is_read', False), 'No')
        
        # Test None value
        self.assertEqual(self.exporter._format_field_value('body', None), '')
        
        # Test list value
        self.assertEqual(
            self.exporter._format_field_value('categories', ['Test', 'Category1']),
            "['Test', 'Category1']"
        )

if __name__ == '__main__':
    unittest.main()
