"""Comprehensive tests for the CSV exporter module.

This module contains tests for the CSVExporter class, including:
- Basic export functionality
- Edge cases and error handling
- Performance with various data sizes
- File encoding support
- Header inclusion/exclusion
"""
import csv
import os
import sys
import tempfile
import pytest
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Any

import pandas as pd
import pytest

from outlook_extractor.export.csv_exporter import CSVExporter
from outlook_extractor.export.constants import EXPORT_FIELDS_V1

# Re-export fixtures from conftest for better IDE support
from tests.conftest import (
    sample_emails, empty_emails, invalid_emails,
    csv_exporter, include_headers, encoding
)

class TestCSVExporter:
    """Comprehensive test suite for the CSVExporter class."""
    
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
    
    @pytest.mark.parametrize("email_count", [1, 5, 10])
    def test_export_various_sizes(self, email_count, tmp_path, csv_exporter):
        """Test export with different numbers of emails."""
        # Setup
        output_file = tmp_path / f"test_export_{email_count}.csv"
        emails = create_test_emails(email_count)
        
        # Test
        result = csv_exporter.export_emails(
            emails=emails,
            output_path=output_file,
            include_headers=True
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        # Check file contents
        with open(output_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            
            # Check headers and row count
            assert reader.fieldnames == EXPORT_FIELDS_V1
            assert len(rows) == email_count
    
    def test_export_with_special_chars(self, special_chars_email, tmp_path, csv_exporter):
        """Test export with special characters in email content."""
        # Setup
        output_file = tmp_path / "special_chars.csv"
        
        # Test
        result = csv_exporter.export_emails(
            emails=[special_chars_email],
            output_path=output_file,
            include_headers=True
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        
        # Check file contents with different encodings
        for enc in ['utf-8', 'utf-8-sig', 'latin-1']:
            try:
                with open(output_file, 'r', encoding=enc) as f:
                    content = f.read()
                    assert special_chars_email['subject'] in content
                    assert '😊' in content  # Emoji
                    assert '•' in content   # Bullet point
                    break  # Stop at first successful encoding
            except UnicodeDecodeError:
                continue
    
    @pytest.mark.parametrize("include_headers", [True, False])
    def test_export_with_and_without_headers(self, sample_emails, tmp_path, 
                                          csv_exporter, include_headers):
        """Test export with and without headers."""
        # Setup
        suffix = "with_headers" if include_headers else "without_headers"
        output_file = tmp_path / f"test_export_{suffix}.csv"
        
        # Test
        result = csv_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=include_headers
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        
        # Check file contents
        with open(output_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            if include_headers:
                # First row should be headers
                assert rows[0] == list(EXPORT_FIELDS_V1)
                assert len(rows) == len(sample_emails) + 1
            else:
                # First row should be data
                assert len(rows) == len(sample_emails)
    
    # --- Edge Case Tests ---
    
    def test_export_empty_list(self, empty_emails, tmp_path, csv_exporter):
        """Test export with empty email list."""
        # Setup
        output_file = tmp_path / "empty_export.csv"
        
        # Test & Verify
        with pytest.raises(ValueError, match="No emails to export"):
            csv_exporter.export_emails(
                emails=empty_emails,
                output_path=output_file,
                include_headers=True
            )
        
        # File should not be created
        assert not output_file.exists()
    
    def test_export_invalid_emails(self, invalid_emails, tmp_path, csv_exporter):
        """Test export with invalid email data."""
        # Setup
        output_file = tmp_path / "invalid_export.csv"
        
        # Test
        with pytest.raises((ValueError, TypeError, AttributeError)):
            csv_exporter.export_emails(
                emails=invalid_emails,
                output_path=output_file,
                include_headers=True
            )
        
        # File should not be created
        assert not output_file.exists()
    
    def test_export_to_nonexistent_directory(self, sample_emails, tmp_path, csv_exporter):
        """Test export to a non-existent directory."""
        # Setup - create a non-existent subdirectory
        output_file = tmp_path / "nonexistent" / "test_export.csv"
        
        # Test
        result = csv_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=True
        )
        
        # Verify - should create the directory and succeed
        assert result is True
        assert output_file.exists()
        assert output_file.parent.exists()
    
    # --- Performance Tests ---
    
    @pytest.mark.parametrize("email_count", [10, 100, 1000])
    def test_export_performance(self, email_count, tmp_path, csv_exporter, benchmark):
        """Test export performance with varying numbers of emails."""
        # Setup - create test emails
        emails = create_test_emails(email_count)
        output_file = tmp_path / f"perf_test_{email_count}.csv"
        
        # Benchmark the export
        def _run_export():
            return csv_exporter.export_emails(
                emails=emails,
                output_path=output_file,
                include_headers=True
            )
        
        # Run benchmark
        result = benchmark(_run_export)
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
    
    # --- File Operation Tests ---
    
    def test_export_to_readonly_directory(self, sample_emails, tmp_path, csv_exporter):
        """Test export to a read-only directory."""
        if sys.platform == 'win32':
            pytest.skip("Read-only directory test not supported on Windows")
            
        # Setup - create a read-only directory
        readonly_dir = tmp_path / "readonly"
        readonly_dir.mkdir(mode=0o555)  # Read-only permissions
        output_file = readonly_dir / "test_export.csv"
        
        # Test
        with pytest.raises(PermissionError):
            csv_exporter.export_emails(
                emails=sample_emails,
                output_path=output_file,
                include_headers=True
            )
        
        # File should not be created
        assert not output_file.exists()
    
    @pytest.mark.parametrize("encoding", [None, 'utf-8', 'utf-8-sig', 'latin-1'])
    def test_export_with_different_encodings(self, sample_emails, tmp_path, 
                                           csv_exporter, encoding):
        """Test export with different file encodings."""
        # Setup
        enc_suffix = encoding.replace('-', '_') if encoding else 'default'
        output_file = tmp_path / f"test_export_{enc_suffix}.csv"
        
        # Test
        result = csv_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=True,
            encoding=encoding
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        
        # Try reading with the same encoding
        try:
            with open(output_file, 'r', encoding=encoding or 'utf-8') as f:
                content = f.read()
                assert len(content) > 0
        except UnicodeDecodeError:
            pytest.fail(f"Failed to read file with encoding: {encoding}")
    
    def test_export_emails_no_headers(self, sample_emails, tmp_path):
        """Test export without headers."""
        # Setup
        output_file = tmp_path / "test_export_no_headers.csv"
        exporter = CSVExporter()
        
        # Test
        result = exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=False
        )
        
        # Verify
        assert result is True
        
        # Check file contents
        with open(output_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            rows = list(reader)
            
            # First row should be data, not headers
            assert rows[0][2] == 'Test Email 1'  # subject is 3rd field
    
    def test_export_empty_list(self, tmp_path):
        """Test export with empty email list."""
        # Setup
        output_file = tmp_path / "empty_export.csv"
        exporter = CSVExporter()
        
        # Test
        result = exporter.export_emails(
            emails=[],
            output_path=output_file,
            include_headers=True
        )
        
        # Verify
        assert result is False
        assert not output_file.exists()
    
    def test_export_to_nonexistent_directory(self, sample_emails, tmp_path):
        """Test export to a non-existent directory."""
        # Setup - create a non-existent subdirectory
        output_file = tmp_path / "nonexistent" / "test_export.csv"
        exporter = CSVExporter()
        
        # Test
        result = exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=True
        )
        
        # Verify - should create the directory and succeed
        assert result is True
        assert output_file.exists()
    
    def test_export_with_invalid_data(self, tmp_path):
        """Test export with invalid email data."""
        # Setup - create email data with missing fields
        invalid_emails = [
            {'subject': 'Test', 'sender_email': 'test@example.com'},  # Missing required fields
            None,  # None value in the list
            12345  # Invalid type
        ]
        
        output_file = tmp_path / "invalid_export.csv"
        exporter = CSVExporter()
        
        # Test - should handle gracefully
        result = exporter.export_emails(
            emails=invalid_emails,
            output_path=output_file,
            include_headers=True
        )
        
        # Verify - should still create the file with valid rows
        assert result is True
        assert output_file.exists()
        
        # Check that the file contains the valid row
        with open(output_file, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            rows = list(reader)
            assert len(rows) == 1  # Only one valid row

    # --- Helper Method Tests ---
    
    def test_format_field_value(self, csv_exporter):
        """Test the _format_field_value method."""
        # Test with different field types
        assert csv_exporter._format_field_value('is_read', True) == 'Yes'
        assert csv_exporter._format_field_value('is_read', False) == 'No'
        assert csv_exporter._format_field_value('is_read', None) == 'No'
        
        # Test with datetime
        now = datetime.now()
        assert csv_exporter._format_field_value('received_time', now) == now.isoformat()
        
        # Test with list
        assert csv_exporter._format_field_value('categories', ['Test', 'Important']) == 'Test; Important'
        
        # Test with special characters
        assert csv_exporter._format_field_value('subject', 'Test "quotes"') == 'Test ""quotes""'
    
    # --- Integration Tests ---
    
    def test_export_then_import(self, sample_emails, tmp_path, csv_exporter):
        """Test that exported CSV can be imported back."""
        # Setup
        output_file = tmp_path / "roundtrip_test.csv"
        
        # Export to CSV
        result = csv_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            include_headers=True
        )
        assert result is True
        
        # Import back using pandas
        try:
            df = pd.read_csv(output_file, encoding='utf-8-sig')
            assert len(df) == len(sample_emails)
            
            # Check that all expected columns are present
            for field in EXPORT_FIELDS_V1:
                assert field in df.columns
                
        except Exception as e:
            pytest.fail(f"Failed to import exported CSV: {e}")
    
    # --- Error Handling Tests ---
    
    def test_export_with_invalid_path(self, sample_emails, csv_exporter):
        """Test export with invalid file path."""
        # Test with None path
        with pytest.raises(ValueError):
            csv_exporter.export_emails(
                emails=sample_emails,
                output_path=None,
                include_headers=True
            )
        
        # Test with directory path
        with tempfile.TemporaryDirectory() as temp_dir:
            with pytest.raises(IsADirectoryError):
                csv_exporter.export_emails(
                    emails=sample_emails,
                    output_path=Path(temp_dir),
                    include_headers=True
                )
    
    @pytest.mark.parametrize("invalid_input", [None, "not_a_list", 123])
    def test_export_with_invalid_emails_param(self, invalid_input, tmp_path, csv_exporter):
        """Test export with invalid emails parameter."""
        output_file = tmp_path / "invalid_input_test.csv"
        
        with pytest.raises((TypeError, ValueError)):
            csv_exporter.export_emails(
                emails=invalid_input,
                output_path=output_file,
                include_headers=True
            )
        
        # File should not be created
        assert not output_file.exists()
