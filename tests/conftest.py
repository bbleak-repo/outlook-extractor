"""Test configuration and fixtures for the Outlook Extractor test suite."""
import os
import sys
import tempfile
import json
import csv
import pytest
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional, Union
from unittest.mock import MagicMock, patch

# Import mock_win32com to set up mocks for Windows-specific modules
if not sys.platform.startswith('win'):
    from tests.mock_win32com import *  # noqa: F403, F401

# Import modules that might use the mocked modules after setting up mocks
from outlook_extractor.core.outlook_client import OutlookClient  # noqa: E402
from outlook_extractor.processors.email_processor import EmailProcessor  # noqa: E402
from outlook_extractor.export import CSVExporter, ExcelExporter, JSONExporter, PDFExporter
from outlook_extractor.config import ConfigManager

# --- Test Data Factories ---

def create_test_email(
    email_id: Union[str, int] = 1,
    subject: str = "Test Email",
    sender: str = "sender@example.com",
    recipients: List[str] = None,
    cc_recipients: List[str] = None,
    bcc_recipients: List[str] = None,
    body: str = "Test email body",
    html_body: str = None,
    has_attachments: bool = False,
    is_read: bool = True,
    is_flagged: bool = False,
    importance: str = "Normal",
    categories: List[str] = None,
    received_days_ago: int = 0,
    sent_days_ago: int = 1,
    **kwargs
) -> Dict[str, Any]:
    """Create a test email dictionary with realistic data.
    
    Args:
        email_id: Unique identifier for the email
        subject: Email subject
        sender: Sender's email address
        recipients: List of recipient email addresses
        cc_recipients: List of CC recipient email addresses
        bcc_recipients: List of BCC recipient email addresses
        body: Plain text email body
        html_body: HTML email body (if None, will be generated from body)
        has_attachments: Whether the email has attachments
        is_read: Whether the email is marked as read
        is_flagged: Whether the email is flagged
        importance: Email importance (Low, Normal, High)
        categories: List of categories/tags
        received_days_ago: Days ago the email was received
        sent_days_ago: Days ago the email was sent
        **kwargs: Additional fields to include in the email
        
    Returns:
        Dict containing email data
    """
    if recipients is None:
        recipients = ["recipient@example.com"]
    if cc_recipients is None:
        cc_recipients = []
    if bcc_recipients is None:
        bcc_recipients = []
    if categories is None:
        categories = ["Test"]
    if html_body is None:
        html_body = f"<p>{body.replace('\n', '<br>')}</p>"
    
    received_date = datetime.now() - timedelta(days=received_days_ago)
    sent_date = datetime.now() - timedelta(days=sent_days_ago)
    
    email = {
        'id': str(email_id),
        'conversation_id': f'conv-{email_id}',
        'subject': subject,
        'sender_name': sender.split('@')[0].title(),
        'sender_email': sender,
        'to_recipients': '; '.join(recipients),
        'cc_recipients': '; '.join(cc_recipients),
        'bcc_recipients': '; '.join(bcc_recipients),
        'received_time': received_date,
        'sent_time': sent_date,
        'categories': '; '.join(categories),
        'importance': importance,
        'sensitivity': 0,
        'has_attachments': has_attachments,
        'is_read': is_read,
        'is_flagged': is_flagged,
        'is_priority': importance.lower() == 'high',
        'is_admin': 'admin' in sender.lower(),
        'body': body,
        'html_body': html_body,
        'folder_path': 'Inbox',
        'thread_id': f'thread-{email_id}',
        'thread_depth': 0,
        'size': len(body) * 2,  # Rough estimate
        **kwargs
    }
    
    # Add attachments if needed
    if has_attachments:
        email['attachments'] = [
            {
                'name': f'document_{email_id}.pdf',
                'size': 1024 * 1024,  # 1MB
                'content_type': 'application/pdf',
                'content': b'%PDF-1.4\n...'  # Minimal PDF header
            },
            {
                'name': f'spreadsheet_{email_id}.xlsx',
                'size': 512 * 1024,  # 512KB
                'content_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'content': b'PK\x03\x04...'  # Minimal ZIP header
            }
        ]
    
    return email

def create_test_emails(count: int = 3, **kwargs) -> List[Dict[str, Any]]:
    """Create multiple test emails with sequential IDs and varying data.
    
    Args:
        count: Number of emails to create
        **kwargs: Additional arguments to pass to create_test_email
        
    Returns:
        List of email dictionaries
    """
    emails = []
    for i in range(1, count + 1):
        # Vary email properties for more realistic test data
        email_kwargs = {
            'email_id': i,
            'subject': f"Test Email {i}" if i > 1 else "Important: Action Required",
            'sender': f"user{i}@example.com",
            'recipients': [f"recipient{i}@example.com"],
            'cc_recipients': [f"cc{i}@example.com"],
            'bcc_recipients': [f"bcc{i}@example.com"] if i % 2 == 0 else [],
            'body': f"This is test email {i}.\n\nIt contains multiple lines of text.",
            'has_attachments': i % 2 == 0,
            'is_read': i % 3 != 0,  # Every 3rd email is unread
            'is_flagged': i % 4 == 0,  # Every 4th email is flagged
            'importance': 'High' if i % 5 == 0 else 'Normal',
            'categories': ['Test'],
            'received_days_ago': i % 7,  # Vary received dates
            'sent_days_ago': (i % 7) + 1,
            **{k: v[i % len(v)] if isinstance(v, (list, tuple)) else v 
               for k, v in kwargs.items()}
        }
        emails.append(create_test_email(**email_kwargs))
    
    return emails

# --- Fixtures ---

@pytest.fixture(scope="session")
def test_data_dir():
    """Create and return a temporary directory for test data."""
    with tempfile.TemporaryDirectory(prefix="outlook_extractor_test_") as temp_dir:
        yield Path(temp_dir)

@pytest.fixture
def sample_emails() -> List[Dict[str, Any]]:
    """Return a list of sample emails for testing."""
    return create_test_emails(5)  # 5 sample emails with varied data

@pytest.fixture
def empty_emails() -> List[Dict[str, Any]]:
    """Return an empty list of emails."""
    return []

@pytest.fixture
def invalid_emails() -> List[Dict[str, Any]]:
    """Return a list of invalid email data."""
    return [
        {},  # Empty dict
        {'subject': 'No body'},  # Missing required fields
        None,  # None value
        123,  # Invalid type
        {'id': 'bad-email', 'subject': 12345},  # Wrong type for field
    ]

# Exporters
@pytest.fixture
def csv_exporter():
    """Return a CSVExporter instance."""
    return CSVExporter()

@pytest.fixture
def excel_exporter():
    """Return an ExcelExporter instance."""
    return ExcelExporter()

@pytest.fixture
def json_exporter():
    """Return a JSONExporter instance."""
    return JSONExporter()

@pytest.fixture
def pdf_exporter():
    """Return a PDFExporter instance."""
    return PDFExporter()

# Export configurations
@pytest.fixture(params=[True, False])
def include_headers(request):
    """Parameterized fixture for include_headers parameter."""
    return request.param

@pytest.fixture(params=[None, 'utf-8', 'utf-8-sig', 'latin-1'])
def encoding(request):
    """Parameterized fixture for different encodings."""
    return request.param

# Keep existing fixtures but update their docstrings

# ... [rest of the existing conftest.py content] ...

@pytest.fixture
def email_processor_config():
    """Return a sample configuration for the EmailProcessor."""
    return {
        'priority_addresses': ['important@example.com', 'ceo@example.com'],
        'admin_addresses': ['admin@example.com', 'it@example.com'],
        'export': {
            'format': 'csv',
            'include_headers': True,
            'output_dir': 'output',
            'filename': 'emails.csv'
        },
        'logging': {
            'level': 'INFO',
            'file': 'test.log'
        }
    }

@pytest.fixture
def email_processor(email_processor_config):
    """Create an EmailProcessor instance for testing."""
    return EmailProcessor(email_processor_config)

@pytest.fixture
def temp_config():
    """Create a temporary config file for testing."""
    with tempfile.NamedTemporaryFile(mode='w+', suffix='.ini', delete=False) as f:
        f.write("""
        [outlook]
        mailbox_name = 
        folder_patterns = Inbox,Sent Items
        max_emails = 1000
        
        [date_range]
        days_back = 30
        date_ranges = 
        
        [threading]
        enable_threading = 1
        thread_method = hybrid
        max_thread_depth = 10
        thread_timeout_days = 30
        
        [storage]
        output_dir = output
        db_filename = test_emails.db
        json_export = 1
        json_pretty_print = 1
        type = sqlite
        
        [logging]
        log_level = INFO
        log_file = test_outlook_extractor.log
        
        [email_processing]
        extract_attachments = 0
        attachment_dir = test_attachments
        extract_embedded_images = 0
        image_dir = test_images
        extract_links = 1
        extract_phone_numbers = 1
        
        [security]
        redact_sensitive_data = 1
        redaction_patterns = password,ssn,credit.?card
        """)
        config_path = f.name
    
    # Yield the path to the config file
    yield config_path
    
    # Clean up the temporary file
    try:
        os.unlink(config_path)
    except Exception as e:
        print(f"Warning: Could not delete temporary config file: {e}")

@pytest.fixture
def config_manager(temp_config):
    """Create a ConfigManager instance with test config."""
    return ConfigManager(temp_config)

@pytest.fixture
def temp_db():
    """Create a temporary SQLite database for testing."""
    import os
    
    # Create a temporary file
    fd, path = tempfile.mkstemp(suffix='.db')
    os.close(fd)
    
    # Create the database
    storage = SQLiteStorage(path)
    storage.initialize()
    storage.close()
    
    # Yield the path to the test database
    yield path
    
    # Clean up
    try:
        os.unlink(path)
    except Exception as e:
        print(f"Warning: Could not delete temporary database file: {e}")

@pytest.fixture
def temp_json():
    """Create a temporary JSON file for testing."""
    import os
    
    # Create a temporary file
    fd, path = tempfile.mkstemp(suffix='.json')
    os.close(fd)
    
    # Yield the path to the test JSON file
    yield path
    
    # Clean up
    try:
        os.unlink(path)
    except Exception as e:
        print(f"Warning: Could not delete temporary JSON file: {e}")
