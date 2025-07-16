"""Comprehensive tests for the JSON exporter module.

This module contains tests for the JSONExporter class, including:
- Basic export functionality
- Data integrity and formatting
- Handling of different data types
- Performance with various data sizes
- Error handling and validation
"""
import json
import os
import sys
import tempfile
import pytest
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Any

from jsonschema import validate, ValidationError

from outlook_extractor.export.json_exporter import JSONExporter
from outlook_extractor.export.constants import EXPORT_FIELDS_V1

# Re-export fixtures from conftest for better IDE support
from tests.conftest import (
    sample_emails, empty_emails, invalid_emails,
    json_exporter, create_test_emails
)

# JSON Schema for validation
EMAIL_SCHEMA = {
    "type": "object",
    "properties": {
        "id": {"type": "string"},
        "conversation_id": {"type": "string"},
        "subject": {"type": "string"},
        "sender_name": {"type": "string"},
        "sender_email": {"type": "string"},
        "to_recipients": {"type": "string"},
        "cc_recipients": {"type": "string"},
        "bcc_recipients": {"type": "string"},
        "received_time": {"type": "string", "format": "date-time"},
        "sent_time": {"type": "string", "format": "date-time"},
        "categories": {"type": "string"},
        "importance": {"type": "string"},
        "sensitivity": {"type": "integer"},
        "has_attachments": {"type": "boolean"},
        "is_read": {"type": "boolean"},
        "is_flagged": {"type": "boolean"},
        "is_priority": {"type": "boolean"},
        "is_admin": {"type": "boolean"},
        "body": {"type": "string"},
        "html_body": {"type": ["string", "null"]},
        "folder_path": {"type": "string"},
        "thread_id": {"type": "string"},
        "thread_depth": {"type": "integer"},
        "size": {"type": "integer"}
    },
    "required": ["id", "subject", "sender_email", "received_time"],
    "additionalProperties": True
}

class TestJSONExporter:
    """Comprehensive test suite for the JSONExporter class."""
    
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
    def test_export_various_sizes(self, email_count, tmp_path, json_exporter):
        """Test export with different numbers of emails."""
        # Setup
        output_file = tmp_path / f"test_export_{email_count}.json"
        emails = create_test_emails(email_count)
        
        # Test
        result = json_exporter.export_emails(
            emails=emails,
            output_path=output_file
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
        
        # Check file contents
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert isinstance(data, list)
            assert len(data) == email_count
    
    def test_export_with_special_chars(self, special_chars_email, tmp_path, json_exporter):
        """Test export with special characters in email content."""
        # Setup
        output_file = tmp_path / "special_chars.json"
        
        # Test
        result = json_exporter.export_emails(
            emails=[special_chars_email],
            output_path=output_file
        )
        
        # Verify
        assert result is True
        assert output_file.exists()
        
        # Check file contents
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert data[0]['subject'] == special_chars_email['subject']
            assert data[0]['body'] == special_chars_email['body']
    
    # --- Formatting Tests ---
    
    def test_export_with_pretty_print(self, sample_emails, tmp_path, json_exporter):
        """Test export with pretty printing."""
        # Setup
        output_file = tmp_path / "pretty_print.json"
        
        # Test with pretty print
        result = json_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file,
            indent=2,
            ensure_ascii=False
        )
        
        # Verify
        assert result is True
        
        # Check that the file has pretty-printed JSON (has newlines)
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert '\n' in content  # Should have newlines for pretty printing
            data = json.loads(content)
            assert len(data) == len(sample_emails)
    
    # --- Error Handling Tests ---
    
    def test_export_empty_list(self, empty_emails, tmp_path, json_exporter):
        """Test export with empty email list."""
        output_file = tmp_path / "empty_export.json"
        
        with pytest.raises(ValueError, match="No emails to export"):
            json_exporter.export_emails(
                emails=empty_emails,
                output_path=output_file
            )
        
        # File should not be created
        assert not output_file.exists()
    
    def test_export_invalid_emails(self, invalid_emails, tmp_path, json_exporter):
        """Test export with invalid email data."""
        output_file = tmp_path / "invalid_export.json"
        
        with pytest.raises((ValueError, TypeError, AttributeError)):
            json_exporter.export_emails(
                emails=invalid_emails,
                output_path=output_file
            )
        
        # File should not be created
        assert not output_file.exists()
    
    # --- Schema Validation Tests ---
    
    def test_export_against_schema(self, sample_emails, tmp_path, json_exporter):
        """Test that exported JSON validates against the schema."""
        # Setup
        output_file = tmp_path / "schema_validation.json"
        
        # Test
        result = json_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file
        )
        
        # Verify
        assert result is True
        
        # Validate against schema
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            for email in data:
                validate(instance=email, schema=EMAIL_SCHEMA)
    
    # --- Performance Tests ---
    
    @pytest.mark.parametrize("email_count", [10, 100, 1000])
    def test_export_performance(self, email_count, tmp_path, json_exporter, benchmark):
        """Test export performance with varying numbers of emails."""
        # Setup - create test emails
        emails = create_test_emails(email_count)
        output_file = tmp_path / f"perf_test_{email_count}.json"
        
        # Benchmark the export
        def _run_export():
            return json_exporter.export_emails(
                emails=emails,
                output_path=output_file,
                indent=None  # Disable pretty printing for performance
            )
        
        # Run benchmark
        result = benchmark(_run_export)
        
        # Verify
        assert result is True
        assert output_file.exists()
        assert output_file.stat().st_size > 0
    
    # --- File Operation Tests ---
    
    def test_export_to_nonexistent_directory(self, sample_emails, tmp_path, json_exporter):
        """Test export to a non-existent directory."""
        output_file = tmp_path / "nonexistent" / "test_export.json"
        
        result = json_exporter.export_emails(
            emails=sample_emails,
            output_path=output_file
        )
        
        assert result is True
        assert output_file.exists()
        assert output_file.parent.exists()
    
    def test_export_with_invalid_path(self, sample_emails, json_exporter):
        """Test export with invalid file path."""
        # Test with None path
        with pytest.raises(ValueError):
            json_exporter.export_emails(
                emails=sample_emails,
                output_path=None
            )
        
        # Test with directory path
        with tempfile.TemporaryDirectory() as temp_dir:
            with pytest.raises(IsADirectoryError):
                json_exporter.export_emails(
                    emails=sample_emails,
                    output_path=Path(temp_dir)
                )
    
    # --- Helper Method Tests ---
    
    def test_serialize_datetime(self, json_exporter):
        """Test the _serialize_datetime method."""
        # Test with datetime
        now = datetime.now()
        assert json_exporter._serialize_datetime(now) == now.isoformat()
        
        # Test with date
        from datetime import date
        today = date.today()
        assert json_exporter._serialize_datetime(today) == today.isoformat()
        
        # Test with string (should pass through)
        assert json_exporter._serialize_datetime("2023-01-01") == "2023-01-01"
        
        # Test with None
        assert json_exporter._serialize_datetime(None) is None


if __name__ == "__main__":
    pytest.main(["-v", "test_json_exporter.py"])
