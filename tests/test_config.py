# Test config
cat > tests/test_config.py << 'EOL'
import os
import pytest
from pathlib import Path
from outlook_extractor.config import ConfigManager, get_config

def test_config_loading(temp_config):
    """Test that config loads correctly from file."""
    config = ConfigManager(temp_config)
    assert config is not None
    assert config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'
    assert config.get_int('outlook', 'max_emails') == 1000
    assert config.get_boolean('threading', 'enable_threading') is True

def test_config_defaults():
    """Test that default values are set correctly."""
    config = ConfigManager()
    assert config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'
    assert config.get_int('outlook', 'max_emails') == 1000

def test_config_get_list(temp_config):
    """Test getting a list from config."""
    config = ConfigManager(temp_config)
    folders = config.get_list('outlook', 'folder_patterns')
    assert folders == ['Inbox', 'Sent Items']
    assert config.get_list('nonexistent', 'nonexistent', ['default']) == ['default']

def test_config_get_boolean(temp_config):
    """Test boolean config values."""
    config = ConfigManager(temp_config)
    assert config.get_boolean('threading', 'enable_threading') is True
    assert config.get_boolean('nonexistent', 'nonexistent', True) is True
    assert config.get_boolean('nonexistent', 'nonexistent', False) is False

def test_config_get_int(temp_config):
    """Test integer config values."""
    config = ConfigManager(temp_config)
    assert config.get_int('outlook', 'max_emails') == 1000
    assert config.get_int('nonexistent', 'nonexistent', 42) == 42

def test_config_save(temp_config, tmp_path):
    """Test saving config to a new file."""
    config = ConfigManager(temp_config)
    new_path = tmp_path / "new_config.ini"
    config.save_config(str(new_path))
    assert new_path.exists()
    
    # Verify the saved config
    new_config = ConfigManager(str(new_path))
    assert new_config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'

def test_get_config_singleton(temp_config):
    """Test that get_config returns a singleton instance."""
    config1 = get_config(temp_config)
    config2 = get_config(temp_config)
    assert config1 is config2
EOL

# Test email threading
cat > tests/test_email_threading.py << 'EOL'
import pytest
from datetime import datetime, timedelta
from outlook_extractor.core.email_threading import EmailThread, ThreadManager

@pytest.fixture
def sample_email_data():
    """Create sample email data for testing."""
    return {
        'id': 'msg123',
        'thread_id': 'thread123',
        'subject': 'Test Email',
        'sender': 'test@example.com',
        'recipients': ['recipient@example.com'],
        'cc_recipients': [],
        'bcc_recipients': [],
        'sent_date': datetime.now(),
        'received_date': datetime.now(),
        'body_text': 'This is a test email',
        'body_html': '<p>This is a test email</p>',
        'is_read': True,
        'importance': 1,
        'has_attachments': False,
        'categories': ['Test'],
        'internet_headers': {'Message-ID': '<msg123@example.com>'},
        'folder_path': 'Inbox'
    }

@pytest.fixture
def thread_manager():
    """Create a ThreadManager instance for testing."""
    return ThreadManager()

def test_email_thread_creation(sample_email_data):
    """Test creating an email thread."""
    thread = EmailThread(
        thread_id='thread123',
        subject='Test Thread',
        participants={'test@example.com'},
        message_ids={'msg123'},
        root_message_id='msg123',
        start_date=datetime.now(),
        end_date=datetime.now(),
        status='active',
        categories={'Test'}
    )
    
    assert thread.thread_id == 'thread123'
    assert 'test@example.com' in thread.participants
    assert 'msg123' in thread.message_ids
    assert thread.status == 'active'

def test_thread_manager_add_email(thread_manager, sample_email_data):
    """Test adding an email to the thread manager."""
    # Add the email
    thread_manager.add_email(sample_email_data)
    
    # Verify the thread was created
    assert 'thread123' in thread_manager.threads_by_id
    thread = thread_manager.threads_by_id['thread123']
    assert thread.subject == 'Test Email'
    assert 'test@example.com' in thread.participants
    assert 'msg123' in thread.message_ids

def test_thread_manager_get_thread(thread_manager, sample_email_data):
    """Test retrieving a thread from the thread manager."""
    # Add the email
    thread_manager.add_email(sample_email_data)
    
    # Get the thread
    thread = thread_manager.get_thread('thread123')
    assert thread is not None
    assert thread['id'] == 'thread123'
    assert thread['subject'] == 'Test Email'
    assert 'test@example.com' in thread['participants']

def test_thread_manager_thread_linking(thread_manager, sample_email_data):
    """Test that emails are properly linked in threads."""
    # First email
    email1 = sample_email_data.copy()
    email1['id'] = 'msg1'
    email1['subject'] = 'Test Email 1'
    email1['internet_headers'] = {'Message-ID': '<msg1@example.com>'}
    
    # Second email (reply to first)
    email2 = sample_email_data.copy()
    email2['id'] = 'msg2'
    email2['subject'] = 'Re: Test Email 1'
    email2['internet_headers'] = {
        'Message-ID': '<msg2@example.com>',
        'In-Reply-To': '<msg1@example.com>',
        'References': '<msg1@example.com>'
    }
    
    # Add emails to thread manager
    thread_manager.add_email(email1)
    thread_manager.add_email(email2)
    
    # Verify both emails are in the same thread
    thread1 = thread_manager.get_thread_for_message('msg1')
    thread2 = thread_manager.get_thread_for_message('msg2')
    assert thread1 is not None
    assert thread1 == thread2
    assert len(thread1['message_ids']) == 2
    assert 'msg1' in thread1['message_ids']
    assert 'msg2' in thread1['message_ids']

def test_thread_manager_to_dict(thread_manager, sample_email_data):
    """Test converting thread manager to dictionary."""
    # Add an email
    thread_manager.add_email(sample_email_data)
    
    # Convert to dict
    data = thread_manager.to_dict()
    
    # Verify the data
    assert 'threads' in data
    assert 'messages' in data
    assert len(data['threads']) == 1
    assert len(data['messages']) == 1
    assert data['threads'][0]['id'] == 'thread123'
    assert data['messages'][0]['id'] == 'msg123'

def test_thread_manager_from_dict(thread_manager, sample_email_data):
    """Test creating thread manager from dictionary."""
    # Create a thread manager with some data
    thread_manager.add_email(sample_email_data)
    data = thread_manager.to_dict()
    
    # Create a new thread manager from the data
    new_manager = ThreadManager.from_dict(data)
    
    # Verify the data was loaded correctly
    assert 'thread123' in new_manager.threads_by_id
    thread = new_manager.threads_by_id['thread123']
    assert thread.subject == 'Test Email'
    assert 'test@example.com' in thread.participants
    assert 'msg123' in thread.message_ids
EOL

# Test storage
cat > tests/test_storage.py << 'EOL'
import pytest
import os
import json
from pathlib import Path
from datetime import datetime, timedelta
from outlook_extractor.storage import SQLiteStorage, JSONStorage

@pytest.fixture
def sample_email():
    """Create a sample email for testing."""
    return {
        'id': 'test123',
        'thread_id': 'thread123',
        'subject': 'Test Email',
        'sender': 'test@example.com',
        'recipients': ['recipient@example.com'],
        'cc_recipients': [],
        'bcc_recipients': [],
        'sent_date': datetime.now(),
        'received_date': datetime.now(),
        'body_text': 'This is a test email',
        'body_html': '<p>This is a test email</p>',
        'is_read': True,
        'importance': 1,
        'has_attachments': False,
        'categories': ['Test'],
        'internet_headers': {'Message-ID': '<test123@example.com>'},
        'folder_path': 'Inbox'
    }

def test_sqlite_storage_save_email(temp_db, sample_email):
    """Test saving an email to SQLite storage."""
    storage = SQLiteStorage(db_path=temp_db)
    
    # Save the email
    assert storage.save_email(sample_email) is True
    
    # Retrieve the email
    email = storage.get_email('test123')
    assert email is not None
    assert email['id'] == 'test123'
    assert email['subject'] == 'Test Email'
    assert 'test@example.com' in email['recipients']

def test_sqlite_storage_save_emails(temp_db, sample_email):
    """Test saving multiple emails to SQLite storage."""
    storage = SQLiteStorage(db_path=temp_db)
    
    # Create multiple emails
    emails = []
    for i in range(3):
        email = sample_email.copy()
        email['id'] = f'test{i}'
        email['subject'] = f'Test Email {i}'
        emails.append(email)
    
    # Save the emails
    assert storage.save_emails(emails) == 3
    
    # Verify the emails were saved
    for i in range(3):
        email = storage.get_email(f'test{i}')
        assert email is not None
        assert email['subject'] == f'Test Email {i}'

def test_json_storage_save_email(temp_json, sample_email):
    """Test saving an email to JSON storage."""
    storage = JSONStorage(json_path=temp_json)
    
    # Save the email
    assert storage.save_email(sample_email) is True
    
    # Retrieve the email
    email = storage.get_email('test123')
    assert email is not None
    assert email['id'] == 'test123'
    assert email['subject'] == 'Test Email'

def test_json_storage_save_emails(temp_json, sample_email):
    """Test saving multiple emails to JSON storage."""
    storage = JSONStorage(json_path=temp_json)
    
    # Create multiple emails
    emails = []
    for i in range(3):
        email = sample_email.copy()
        email['id'] = f'test{i}'
        email['subject'] = f'Test Email {i}'
        emails.append(email)
    
    # Save the emails
    assert storage.save_emails(emails) == 3
    
    # Verify the emails were saved
    for i in range(3):
        email = storage.get_email(f'test{i}')
        assert email is not None
        assert email['subject'] == f'Test Email {i}'

def test_sqlite_storage_search(temp_db, sample_email):
    """Test searching emails in SQLite storage."""
    storage = SQLiteStorage(db_path=temp_db)
    
    # Add some test emails
    emails = []
    for i in range(5):
        email = sample_email.copy()
        email['id'] = f'test{i}'
        email['subject'] = f'Test Email {i}'
        email['sender'] = f'sender{i}@example.com'
        email['body_text'] = f'This is test email {i} with some unique content {i*100}'
        emails.append(email)
    
    storage.save_emails(emails)
    
    # Search by subject
    results = storage.search_emails('Test Email 1')
    assert len(results) == 1
    assert results[0]['id'] == 'test1'
    
    # Search by content
    results = storage.search_emails('unique content 300')
    assert len(results) == 1
    assert results[0]['id'] == 'test3'
    
    # Search by sender
    results = storage.search_emails('sender4@example.com')
    assert len(results) == 1
    assert results[0]['id'] == 'test4'

def test_json_storage_search(temp_json, sample_email):
    """Test searching emails in JSON storage."""
    storage = JSONStorage(json_path=temp_json)
    
    # Add some test emails
    emails = []
    for i in range(5):
        email = sample_email.copy()
        email['id'] = f'test{i}'
        email['subject'] = f'Test Email {i}'
        email['sender'] = f'sender{i}@example.com'
        email['body_text'] = f'This is test email {i} with some unique content {i*100}'
        emails.append(email)
    
    storage.save_emails(emails)
    
    # Search by subject
    results = storage.search_emails('Test Email 1')
    assert len(results) == 1
    assert results[0]['id'] == 'test1'
    
    # Search by content
    results = storage.search_emails('unique content 300')
    assert len(results) == 1
    assert results[0]['id'] == 'test3'
    
    # Search by sender
    results = storage.search_emails('sender4@example.com')
    assert len(results) == 1
    assert results[0]['id'] == 'test4'
EOL

# Test main application
cat > tests/test_main.py << 'EOL'
import pytest
from unittest.mock import MagicMock, patch
from outlook_extractor.main import OutlookExtractor

@pytest.fixture
def mock_outlook():
    """Create a mock Outlook client."""
    with patch('outlook_extractor.core.outlook_client.OutlookClient') as mock:
        yield mock

@pytest.fixture
def mock_storage():
    """Create a mock storage."""
    with patch('outlook_extractor.storage.SQLiteStorage') as mock:
        yield mock

def test_outlook_extractor_init(temp_config):
    """Test initializing the OutlookExtractor."""
    extractor = OutlookExtractor(config_path=temp_config)
    assert extractor is not None
    assert extractor.config is not None

def test_extract_emails(mock_outlook, mock_storage, temp_config):
    """Test the extract_emails method."""
    # Set up mocks
    mock_client = MagicMock()
    mock_client.connect.return_value = True
    mock_client.extract_emails.return_value = [{'id': 'test123', 'subject': 'Test Email'}]
    mock_outlook.return_value = mock_client
    
    mock_storage.return_value.save_emails.return_value = 1
    
    # Initialize and test
    extractor = OutlookExtractor(config_path=temp_config)
    result = extractor.extract_emails()
    
    # Verify results
    assert result['success'] is True
    assert result['emails_processed'] == 1
    assert result['emails_saved'] == 1
    mock_client.connect.assert_called_once()
    mock_client.extract_emails.assert_called_once()
    mock_storage.return_value.save_emails.assert_called_once()

def test_search_emails(mock_storage, temp_config):
    """Test the search_emails method."""
    # Set up mock
    mock_storage.return_value.search_emails.return_value = [
        {'id': 'test1', 'subject': 'Test 1'},
        {'id': 'test2', 'subject': 'Test 2'}
    ]
    
    # Initialize and test
    extractor = OutlookExtractor(config_path=temp_config)
    results = extractor.search_emails('test')
    
    # Verify results
    assert len(results) == 2
    assert results[0]['subject'] == 'Test 1'
    mock_storage.return_value.search_emails.assert_called_once_with('test', limit=100)

def test_get_email(mock_storage, temp_config):
    """Test the get_email method."""
    # Set up mock
    test_email = {'id': 'test123', 'subject': 'Test Email'}
    mock_storage.return_value.get_email.return_value = test_email
    
    # Initialize and test
    extractor = OutlookExtractor(config_path=temp_config)
    email = extractor.get_email('test123')
    
    # Verify results
    assert email == test_email
    mock_storage.return_value.get_email.assert_called_once_with('test123')
EOL

# Create a requirements-test.txt file
cat > requirements-test.txt << 'EOL'
pytest>=7.0.0
pytest-mock>=3.10.0
pytest-cov>=3.0.0
EOL

# Create a README.md for the tests
cat > tests/README.md << 'EOL'
# Outlook Extractor Tests

This directory contains the test suite for the Outlook Extractor application.

## Running Tests

### Install Dependencies

```bash
pip install -r requirements-test.txt
