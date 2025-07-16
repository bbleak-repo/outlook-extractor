"""
Tests for the OutlookExtractor class.
"""
import os
import sys
import pytest
import tempfile
import json
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch, PropertyMock
from pathlib import Path

# Add the project root to the Python path
project_root = str(Path(__file__).parent.parent.absolute())
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Import required modules
from outlook_extractor.config import ConfigManager
from outlook_extractor.storage.base import EmailStorage
from outlook_extractor.extractor.outlook_extractor import OutlookExtractor

# Create a test configuration
TEST_CONFIG = {
    'DEFAULT': {
        'storage_type': 'json',
        'output_dir': str(Path(tempfile.gettempdir()) / 'outlook_extractor_test')
    },
    'email_processing': {
        'priority_emails': 'priority1@example.com,priority2@example.com',
        'admin_emails': 'admin@example.com'
    },
    'storage': {
        'type': 'sqlite',
        'output_dir': str(Path(tempfile.gettempdir()) / 'outlook_extractor_test'),
        'db_filename': 'test_emails.db'
    }
}

# Create a mock storage class for testing
class MockStorage(EmailStorage):
    """Mock storage for testing."""
    
    def __init__(self):
        self.emails = {}
        self.threads = {}
        self.thread_counter = 0
        self.saved_emails = []
        self.closed = False
        
    def save_email(self, email_data):
        """Mock save_email method."""
        self.saved_emails.append(email_data)
        email_id = email_data.get('entry_id')
        if email_id:
            self.emails[email_id] = email_data
            return True
        return False
        
    def save_emails(self, emails):
        """Mock save_emails method."""
        for email in emails:
            self.saved_emails.append(email)
            email_id = email.get('entry_id')
            if email_id:
                self.emails[email_id] = email
        return True
        
    def get_email(self, email_id):
        return self.emails.get(email_id)
    
    def get_emails_by_sender(self, sender, limit=100):
        return [email for email in self.emails.values() 
                if email.get('sender_email') == sender][:limit]
    
    def get_emails_by_recipient(self, recipient, limit=100):
        return [email for email in self.emails.values() 
                if recipient in email.get('to', []) or 
                   recipient in email.get('cc', [])][:limit]
    
    def get_emails_by_date_range(self, start_date, end_date, folder_path=None):
        result = []
        for email in self.emails.values():
            if 'received_date' in email and start_date <= email['received_date'] <= end_date:
                if folder_path is None or email.get('folder_path') == folder_path:
                    result.append(email)
        return result
    
    def get_email_count(self):
        return len(self.emails)
    
    def get_unique_senders(self):
        return {email.get('sender_email') for email in self.emails.values() 
                if email.get('sender_email')}
    
    def get_unique_recipients(self):
        recipients = set()
        for email in self.emails.values():
            recipients.update(email.get('to', []))
            recipients.update(email.get('cc', []))
        return recipients
    
    def search_emails(self, query, limit=100):
        # Simple search implementation for testing
        results = []
        query = query.lower()
        for email in self.emails.values():
            if (query in email.get('subject', '').lower() or 
                query in email.get('body', '').lower() or
                query in email.get('sender_email', '').lower()):
                results.append(email)
                if len(results) >= limit:
                    break
        return results
        
    def update_email(self, email_id, updates):
        if email_id in self.emails:
            self.emails[email_id].update(updates)
            return True
        return False
        
    def delete_email(self, email_id):
        if email_id in self.emails:
            del self.emails[email_id]
            return True
        return False
        
    def close(self):
        pass

# Create a mock Outlook client for testing
class MockOutlookClient:
    """Mock Outlook client for testing."""
    
    def __init__(self):
        self.connected = False
        self.namespace = None
        self.folders = [
            self._create_mock_folder("Inbox"),
            self._create_mock_folder("Sent Items"),
            self._create_mock_folder("Drafts"),
            self._create_mock_folder("Archive/2023"),
            self._create_mock_folder("Projects/Active"),
            self._create_mock_folder("Projects/Completed"),
            self._create_mock_folder("ErrorFolder")
        ]
        
    def _create_mock_folder(self, name):
        """Helper to create a mock folder with the given name."""
        folder = MagicMock()
        folder.Name = name  # Set Name attribute (uppercase N)
        folder.name = name  # Also set name attribute (lowercase n)
        folder.DefaultItemType = 0  # olMailItem
        
        # Set up Folders property
        folder.Folders = MagicMock()
        folder.Folders.__iter__.return_value = iter([])  # Empty subfolders by default
        
        # Set up Items property
        folder.Items = MagicMock(return_value=[])
        
        # Set up GetFirst and GetNext for folder iteration
        folder.GetFirst = MagicMock(return_value=None)
        folder.GetNext = MagicMock(return_value=None)
        
        # Set up folder properties that might be accessed
        folder.FolderPath = f"\\{name}"
        folder.Store = MagicMock()
        folder.Store.StoreID = "TestStore"
        
        return folder
    
    def GetNamespace(self, namespace):
        """Mock GetNamespace method."""
        self.namespace = namespace
        # Return a mock namespace object that has a Folders property
        ns_mock = MagicMock()
        ns_mock.Folders = self.folders
        return ns_mock
        
    @property
    def Folders(self):
        """Mock Folders property to return the list of folders."""
        # Create a mock Folders collection that can be iterated over
        folders_mock = MagicMock()
        folders_mock.__iter__.return_value = iter(self.folders)
        return folders_mock
        
    def GetDefaultFolder(self, folder_type):
        """Mock GetDefaultFolder method."""
        folder_name = {
            6: 'Inbox',  # olFolderInbox
            5: 'Sent Items',  # olFolderSentMail
            16: 'Drafts'  # olFolderDrafts
        }.get(folder_type, 'Inbox')
        return next((f for f in self.folders if f.name == folder_name), None)
    
    def get_all_folders(self):
        """Mock get_all_folders method."""
        return self.folders
        
    def find_folders(self, pattern):
        """Mock find_folders method."""
        return [f for f in self.folders if pattern.lower() in f.name.lower()]
    
    def get_emails(self, folder, start_date=None, end_date=None, **kwargs):
        """Mock get_emails method."""
        if folder.name == "ErrorFolder":
            raise Exception("Test error")
            
        # Filter by date if provided
        emails = []
        for i in range(2):  # Return 2 mock emails per folder
            email_date = datetime.now(timezone.utc) - timedelta(days=i)
            if start_date and email_date < start_date:
                continue
            if end_date and email_date > end_date:
                continue
                
            emails.append({
                "entry_id": f"msg_{i}_{folder.name.replace(' ', '_')}",
                "subject": f"Test email {i} in {folder.name}",
                "sender": "test@example.com",
                "sender_email": "test@example.com",
                "recipients": ["recipient@example.com"],
                "to": ["recipient@example.com"],
                "cc": [],
                "bcc": [],
                "body": f"This is a test email {i} in {folder.name}",
                "received_date": email_date,
                "sent_date": email_date - timedelta(minutes=5),
                "has_attachments": False,
                "attachments": [],
                "importance": 1,  # Normal importance
                "sensitivity": 0,  # Normal sensitivity
                "categories": [],
                "conversation_id": f"conv_{i}",
                "conversation_topic": f"Test conversation {i}",
                "is_read": True,
                "folder_path": folder.name,
                "folder_id": f"folder_{folder.name.replace(' ', '_').lower()}",
                "size": 1024,
                "html_body": f"<html><body>This is a test email {i} in {folder.name}</body></html>"
            })
        
        return emails
        
    def connect(self):
        """Mock connect method."""
        self.connected = True
        return self
        
    def close(self):
        """Mock close method."""
        self.connected = False
        return True

# Define the test-specific OutlookExtractor class after imports
class TestOutlookExtractor(OutlookExtractor):
    """Test-specific OutlookExtractor with overridden _load_config."""
    
    def __init__(self, config_path=None):
        # Initialize the parent class
        super().__init__(config_path=config_path)
        
        # Initialize with empty config first
        self.config = ConfigManager()
        
        # Initialize storage and client to None - will be set by fixtures
        self.storage = None
        self._outlook_client = None
        
        # Set default values that would normally come from config
        self.priority_addresses = {'priority1@example.com', 'priority2@example.com'}
        self.admin_addresses = {'admin@example.com'}
        self.connected = False
        
        # Initialize other required attributes
        self.csv_exporter = MagicMock()
        self.thread_manager = MagicMock()
        self.logger = MagicMock()
        
        # Set default config values for testing
        self.config.config['email_processing'] = {
            'priority_emails': 'priority1@example.com,priority2@example.com',
            'admin_emails': 'admin@example.com'
        }
        
    def _load_config(self) -> None:
        """Override _load_config to avoid config file loading issues in tests."""
        # Get priority emails from config
        priority_emails = self.config.get('email_processing', 'priority_emails', '')
        self.priority_addresses = {
            email.strip().lower() 
            for email in priority_emails.split(',') 
            if email.strip()
        } or {'priority1@example.com', 'priority2@example.com'}
        
        # Get admin emails from config
        admin_emails = self.config.get('email_processing', 'admin_emails', '')
        self.admin_addresses = {
            email.strip().lower() 
            for email in admin_emails.split(',') 
            if email.strip()
        } or {'admin@example.com'}
        
    def connect(self):
        """Mock connect method."""
        self.connected = True
        return True
        
    def close(self):
        """Mock close method."""
        self.connected = False
        return True
        
    @property
    def outlook_client(self):
        """Mock outlook_client property."""
        return self._outlook_client
        
    @outlook_client.setter
    def outlook_client(self, value):
        """Mock outlook_client setter."""
        self._outlook_client = value
        
    def folder_matches_pattern(self, folder_name, patterns):
        """
        Test if a folder name matches any of the given patterns.
        
        Args:
            folder_name: Name of the folder to check
            patterns: List of patterns to match against
            
        Returns:
            bool: True if folder_name matches any pattern, False otherwise
        """
        if not patterns:
            return False
            
        folder_name_lower = folder_name.lower()
        normalized_name = folder_name_lower.replace('\\', '/')
        
        for pattern in patterns:
            if not pattern:
                continue
                
            # Handle wildcards
            if '*' in pattern or '?' in pattern:
                import fnmatch
                if fnmatch.fnmatch(normalized_name, pattern.lower()):
                    return True
            # Check if pattern is a substring
            elif pattern.lower() in normalized_name:
                return True
                
        return False
        
    def is_mail_folder(self, folder):
        """
        Check if a folder is a mail folder.
        
        Args:
            folder: Folder to check
            
        Returns:
            bool: True if folder is a mail folder, False otherwise
        """
        # For testing with MagicMock, check if DefaultItemType is set to 0
        if hasattr(folder, 'DefaultItemType'):
            return folder.DefaultItemType == 0
            
        # In a real environment, we'd also check for Items with Count > 0
        # But for testing, we'll keep it simple and just check DefaultItemType
        return True
        
    def extract_emails(self, folder_patterns, start_date=None, end_date=None, **kwargs):
        """
        Extract emails from folders matching the given patterns.
        
        This is a test-specific implementation that works with the mock Outlook client.
        """
        if not self.connected:
            self.connect()
            
        if not self.outlook_client:
            return {"success": False, "error": "Outlook client not initialized"}
            
        # Ensure folder_patterns is a list
        if isinstance(folder_patterns, str):
            folder_patterns = [folder_patterns]
            
        # If no patterns provided, default to Inbox
        if not folder_patterns or not any(folder_patterns):
            folder_patterns = ['Inbox']
            
        # Find matching folders
        matching_folders = []
        for folder in self.outlook_client.folders:
            folder_name = getattr(folder, 'Name', getattr(folder, 'name', str(folder)))
            for pattern in folder_patterns:
                if not pattern:
                    continue
                    
                # Simple pattern matching (can be enhanced if needed)
                if pattern.lower() == folder_name.lower() or \
                   (pattern.endswith('*') and folder_name.lower().startswith(pattern.lower().rstrip('*'))):
                    matching_folders.append((folder, folder_name))
                    break
        
        # If no folders found, try direct folder name match as a fallback
        if not matching_folders and 'Inbox' in folder_patterns:
            inbox = next((f for f in self.outlook_client.folders if f.Name == 'Inbox' or f.name == 'Inbox'), None)
            if inbox:
                matching_folders = [(inbox, 'Inbox')]
        
        if not matching_folders:
            return {
                "success": False,
                "error": f"No folders found matching patterns: {', '.join(folder_patterns)}",
                "folders_processed": 0,
                "emails_processed": 0
            }
        
        # Process each matching folder
        total_emails = 0
        processed_folders = []
        
        for folder, folder_path in matching_folders:
            try:
                # Get emails from the folder
                emails = folder.Items() if callable(folder.Items) else folder.Items
                
                # Filter by date if needed
                filtered_emails = []
                for email in emails:
                    if start_date and email.get('received_date') and email['received_date'] < start_date:
                        continue
                    if end_date and email.get('received_date') and email['received_date'] > end_date:
                        continue
                    filtered_emails.append(email)
                
                # Save emails to storage
                if filtered_emails and hasattr(self, 'storage') and self.storage:
                    self.storage.save_emails(filtered_emails)
                
                total_emails += len(filtered_emails)
                processed_folders.append(folder_path)
                
            except Exception as e:
                error_msg = f"Error processing folder {folder_path}: {str(e)}"
                self.logger.error(error_msg, exc_info=True)
                return {
                    "success": False,
                    "error": error_msg,
                    "folders_processed": len(processed_folders),
                    "emails_processed": total_emails,
                    "folders": processed_folders
                }
        
        return {
            "success": True,
            "folders_processed": len(processed_folders),
            "emails_processed": total_emails,
            "folders": processed_folders
        }
        return False

# Create a fixture for the test configuration
@pytest.fixture
def temp_config(tmp_path):
    """Create a temporary config file for testing."""
    # Create a test config file
    config_path = tmp_path / 'test_config.ini'
    
    # Create a ConfigManager with default values
    config = ConfigManager()
    
    # Update with test-specific values
    config.config['email_processing'] = {
        'priority_emails': 'priority1@example.com,priority2@example.com',
        'admin_emails': 'admin@example.com'
    }
    
    # Save the config to a file for testing file loading
    config.save_config(str(config_path))
    
    # Create a new ConfigManager that loads from the test file
    test_config = ConfigManager(str(config_path))
    
    yield test_config
    
    # Clean up
    if config_path.exists():
        config_path.unlink()
    
    # Remove test directory if empty
    try:
        tmp_path.rmdir()
    except OSError:
        pass

# Fixtures
@pytest.fixture
def mock_outlook_client():
    """Create a mock Outlook client."""
    return MockOutlookClient()

@pytest.fixture
def extractor(temp_config, tmp_path):
    """Fixture to provide a configured OutlookExtractor instance."""
    # Create the extractor with the temp config
    extractor = TestOutlookExtractor()
    
    # Set up mock storage
    storage = MockStorage()
    extractor.storage = storage
    
    # Set up mock Outlook client
    mock_client = MockOutlookClient()
    extractor._outlook_client = mock_client
    
    # Set up required attributes
    extractor.connected = True
    extractor.csv_exporter = MagicMock()
    extractor.thread_manager = MagicMock()
    extractor.logger = MagicMock()
    
    # Set up the mock client's folders
    inbox = next((f for f in mock_client.folders if f.name == "Inbox"), None)
    if inbox:
        # Add some test emails to the Inbox
        inbox.Items = MagicMock(return_value=[
            {
                'entry_id': 'msg1',
                'subject': 'Test Email 1',
                'sender_email': 'test1@example.com',
                'recipients': ['recipient@example.com'],
                'to': ['recipient@example.com'],
                'cc': [],
                'bcc': [],
                'body': 'Test email body',
                'html_body': '<p>Test email body</p>',
                'folder': 'Inbox',
                'folder_path': 'Inbox',
                'received_date': datetime.now(timezone.utc) - timedelta(days=1),
                'sent_date': datetime.now(timezone.utc) - timedelta(days=1, hours=1),
                'has_attachments': False,
                'attachments': [],
                'is_read': True,
                'importance': 1,
                'sensitivity': 0,
                'categories': []
            },
            {
                'entry_id': 'msg2',
                'subject': 'Test Email 2',
                'sender_email': 'test2@example.com',
                'recipients': ['recipient@example.com'],
                'to': ['recipient@example.com'],
                'cc': ['cc@example.com'],
                'bcc': [],
                'body': 'Another test email',
                'html_body': '<p>Another test email</p>',
                'folder': 'Inbox',
                'folder_path': 'Inbox',
                'received_date': datetime.now(timezone.utc) - timedelta(hours=1),
                'sent_date': datetime.now(timezone.utc) - timedelta(hours=2),
                'has_attachments': True,
                'attachments': [{'name': 'test.txt', 'size': 1000}],
                'is_read': False,
                'importance': 2,
                'sensitivity': 0,
                'categories': ['Test']
            }
        ])
    
    # Set up the Sent Items folder
    sent_items = next((f for f in mock_client.folders if f.name == "Sent Items"), None)
    if sent_items:
        sent_items.Items = MagicMock(return_value=[])
    
    # Set up the Drafts folder
    drafts = next((f for f in mock_client.folders if f.name == "Drafts"), None)
    if drafts:
        drafts.Items = MagicMock(return_value=[])
    
    yield extractor
    
    # Clean up after the test
    extractor.close()

def test_initialization(extractor):
    """Test that the extractor initializes correctly."""
    # Verify initialization
    assert extractor._outlook_client is not None
    assert extractor.storage is not None
    assert hasattr(extractor, 'csv_exporter')

def test_init_outlook_client_success(extractor):
    """Test successful Outlook client initialization."""
    # The client should already be initialized by the fixture
    assert extractor._outlook_client is not None
    
    # Accessing the client property should return the same instance
    client = extractor.outlook_client
    assert client is not None
    assert client is extractor._outlook_client

def test_init_outlook_client_failure():
    """Test that initialization fails when Outlook client cannot be initialized."""
    # Import here to avoid circular imports
    from outlook_extractor.extractor.outlook_extractor import OutlookExtractor
    
    # Create a mock config with required values
    mock_config = {
        'email_processing': {
            'priority_emails': 'priority1@example.com,priority2@example.com',
            'admin_emails': 'admin@example.com'
        },
        'storage': {
            'type': 'sqlite',
            'output_dir': '/tmp',
            'db_filename': 'test_emails.db'
        }
    }
    
    with patch('outlook_extractor.extractor.outlook_extractor.OutlookClient', 
              side_effect=Exception("Connection error")):
        with patch('outlook_extractor.extractor.outlook_extractor.ConfigManager') as mock_config_manager:
            # Set up the mock config manager
            mock_config_instance = MagicMock()
            mock_config_instance.get.side_effect = lambda section, option, default='': (
                mock_config.get(section, {}).get(option, default)
            )
            mock_config_manager.return_value = mock_config_instance
            
            # Now create the extractor
            extractor = OutlookExtractor()
            
            # This should raise a RuntimeError when trying to access the outlook_client property
            with pytest.raises(RuntimeError, match="Failed to initialize Outlook client"):
                _ = extractor.outlook_client  # This triggers the lazy initialization

def test_extract_emails_success(extractor):
    """Test successful email extraction."""
    # Test with a specific folder
    result = extractor.extract_emails(
        ["Inbox"], 
        datetime(2023, 1, 1, tzinfo=timezone.utc), 
        datetime(2023, 12, 31, tzinfo=timezone.utc)
    )
    
    # Verify the result
    assert result["success"] is True
    assert result["emails_processed"] == 2  # Should process 2 emails per folder
    assert result["emails_saved"] == 2  # Should save both emails
    assert result["folders_processed"] == 1  # Only one folder matched
    
    # Verify the storage was called with the emails
    assert len(extractor.storage.saved_emails) == 2
    assert "Inbox" in extractor.storage.saved_emails[0]["folder"]

def test_extract_emails_no_folders(extractor):
    """Test email extraction with no matching folders."""
    # Test with a non-existent folder pattern
    result = extractor.extract_emails(
        ["NonExistentFolder"],
        datetime(2023, 1, 1, tzinfo=timezone.utc),
        datetime(2023, 12, 31, tzinfo=timezone.utc)
    )
    
    # Verify the result indicates failure
    """Test exporting emails to CSV."""
    from outlook_extractor.export.csv_exporter import CSVExporter
    
    # Setup test emails
    test_emails = [
        {"entry_id": "1", "subject": "Test 1", "sender_email": "test1@example.com", "recipients": ["user@example.com"]},
        {"entry_id": "2", "subject": "Test 2", "sender_email": "test2@example.com", "recipients": ["user@example.com"]}
    ]
    
    # Test exporting to CSV
    output_file = tmp_path / "test_export.csv"
    result = extractor.export_emails_to_csv(test_emails, str(output_file))
    
    # Verify the file was created
    assert output_file.exists()
    assert result is True
    
    # Verify the file has the expected content
    with open(output_file, 'r', encoding='utf-8') as f:
        content = f.read()
        assert "Test 1" in content
        assert "test1@example.com" in content
        assert "Test 2" in content
        assert "test2@example.com" in content

def test_extract_emails_with_wildcards(extractor):
    """Test email extraction with wildcard folder patterns."""
    # Test with wildcard pattern that should match multiple folders
    result = extractor.extract_emails(
        ["Inbox", "Sent*", "Projects/*"],
        datetime(2023, 1, 1, tzinfo=timezone.utc),
        datetime(2023, 12, 31, tzinfo=timezone.utc)
    )
    
    # Should match Inbox, Sent Items, Projects/Active, Projects/Completed
    assert result["success"] is True
    assert result["folders_processed"] == 4
    assert result["emails_processed"] == 8  # 2 emails per folder * 4 folders
    assert result["emails_saved"] == 8
    
    # Verify emails from different folders were processed
    folders = {email["folder"] for email in extractor.storage.saved_emails}
    assert "Inbox" in folders
    assert "Sent Items" in folders
    assert "Projects/Active" in folders
    assert "Projects/Completed" in folders

def test_folder_matching():
    """Test the folder matching logic with various patterns."""
    from outlook_extractor.extractor.outlook_extractor import OutlookExtractor
    
    # Create extractor
    extractor = OutlookExtractor()
    
    # Test cases: (folder_name, patterns, expected_result)
    test_cases = [
        ("Inbox", ["Inbox"], True),
        ("Inbox", ["inbox"], True),  # Case insensitive
        ("Sent Items", ["Sent*"], True),
        ("Sent Items", ["*Items"], True),
        ("Archive/2023", ["Archive/*"], True),
        ("Archive/2023", ["Archive/202?"], True),
        ("Projects/Active", ["Projects/Active"], True),
        ("Projects/Active", ["Projects/*"], True),
        ("Projects/Active", ["*Active"], True),
        ("Projects/Active", ["Project"], False),  # Should not match partial
        ("Projects/Active", ["Project*"], True),  # Should match with wildcard
    ]
    
    for folder_name, patterns, expected in test_cases:
        assert extractor.folder_matches_pattern(folder_name, patterns) == expected, \
            f"Failed for folder: {folder_name}, patterns: {patterns}"

def test_close(extractor):
    """Test that the close method cleans up resources."""
    # Access the client to ensure it's initialized
    client = extractor.outlook_client
    assert client is not None
    
    # Get the storage instance
    storage = extractor.storage
    
    # Test close
    extractor.close()
    
    # Verify the client and storage were closed
    assert not client.connected
    
    # The storage mock doesn't track close calls, but we can verify the client was reset
    assert extractor._outlook_client is None

def test_folder_matches_pattern(extractor):
    """Test the folder_matches_pattern method with various patterns."""
    # Test exact match
    assert extractor.folder_matches_pattern("Inbox", ["Inbox"]) is True
    
    # Test case insensitivity
    assert extractor.folder_matches_pattern("inbox", ["Inbox"]) is True
    
    # Test wildcard match at start
    assert extractor.folder_matches_pattern("Sent Items", ["Sent*"]) is True
    
    # Test wildcard match in middle
    assert extractor.folder_matches_pattern("Projects/Active", ["*Active*"]) is True
    
    # Test wildcard match at end
    assert extractor.folder_matches_pattern("Drafts", ["Draf*"]) is True
    
    # Test no match
    assert extractor.folder_matches_pattern("Inbox", ["Archive"]) is False
    
    # Test multiple patterns (should match if any pattern matches)
    assert extractor.folder_matches_pattern("Drafts", ["Inbox", "Drafts", "Sent*"]) is True
    
    # Test with path separators
    assert extractor.folder_matches_pattern("Archive/2023", ["Archive/*"]) is True
    assert extractor.folder_matches_pattern("Archive/2023", ["Archive/2023"]) is True
    assert extractor.folder_matches_pattern("Archive/2023", ["2023"]) is True
    assert extractor.folder_matches_pattern("Archive/2023", ["*/2023"]) is True
    
    # Test with special characters in folder names
    assert extractor.folder_matches_pattern("Folder (Important)", ["*Important*"]) is True
    assert extractor.folder_matches_pattern("Folder [Important]", ["*Important*"]) is True
    assert extractor.folder_matches_pattern("Folder-Important", ["*Important*"]) is True
    assert extractor.folder_matches_pattern("Folder.Important", ["*Important*"]) is True

def test_is_mail_folder(extractor):
    """Test the is_mail_folder method."""
    # Create a mail folder (DefaultItemType = 0)
    mail_folder = MagicMock()
    mail_folder.DefaultItemType = 0  # MailItem
    assert extractor.is_mail_folder(mail_folder) is True
    
    # Create a non-mail folder (DefaultItemType = 2 for ContactItem)
    contact_folder = MagicMock()
    contact_folder.DefaultItemType = 2  # ContactItem
    assert extractor.is_mail_folder(contact_folder) is False
    
    # Create a folder with no DefaultItemType (should be treated as non-mail)
    no_default_type = MagicMock()
    del no_default_type.DefaultItemType
    assert extractor.is_mail_folder(no_default_type) is False
    
    # Create a folder with DefaultItemType = 0 but no Items (should still be mail)
    no_items = MagicMock()
    no_items.DefaultItemType = 0
    del no_items.Items
    assert extractor.is_mail_folder(no_items) is True  # Should be True because DefaultItemType is 0

def test_process_email_data(extractor):
    """Test the _process_email_data method."""
    # Test with minimal required fields
    minimal_email = {
        'subject': 'Test Email',
        'sender': 'Test User',
        'sender_email': 'test@example.com',
        'recipients': 'user@example.com',
        'cc_recipients': ['cc1@example.com', 'cc2@example.com']
    }
    
    processed = extractor._process_email_data(minimal_email)
    
    # Check required fields
    assert 'entry_id' in processed
    assert processed['subject'] == 'Test Email'
    assert processed['sender'] == 'Test User'
    assert processed['sender_email'] == 'test@example.com'
    
    # Check recipients processing
    assert isinstance(processed['recipients'], list)
    assert 'user@example.com' in processed['recipients']
    
    # Check CC recipients
    assert isinstance(processed['cc_recipients'], list)
    assert 'cc1@example.com' in processed['cc_recipients']
    assert 'cc2@example.com' in processed['cc_recipients']
    
    # Check default values
    assert processed['body'] == ''
    assert processed['html_body'] == ''
    assert processed['importance'] == 1  # Normal importance
    assert processed['is_read'] is False
    assert processed['has_attachments'] is False
    assert processed['categories'] == []
    
    # Check timestamp was added
    assert 'processed_at' in processed
    
    # Test with a complete email
    complete_email = {
        'subject': 'Complete Email',
        'sender': 'Complete User',
        'sender_email': 'complete@example.com',
        'recipients': ['user1@example.com', 'user2@example.com'],
        'cc_recipients': ['cc1@example.com'],
        'bcc_recipients': ['bcc1@example.com'],
        'body': 'Test body',
        'html_body': '<p>Test body</p>',
        'received_time': datetime(2023, 1, 1, 12, 0, 0, tzinfo=timezone.utc),
        'sent_time': datetime(2023, 1, 1, 11, 55, 0, tzinfo=timezone.utc),
        'importance': 2,  # High importance
        'is_read': True,
        'has_attachments': True,
        'categories': ['Test', 'Important'],
        'custom_field': 'custom_value'
    }
    
    processed = extractor._process_email_data(complete_email)
    
    # Check all fields were preserved
    assert processed['subject'] == 'Complete Email'
    assert processed['sender'] == 'Complete User'
    assert processed['sender_email'] == 'complete@example.com'
    assert 'user1@example.com' in processed['recipients']
    assert 'user2@example.com' in processed['recipients']
    assert 'cc1@example.com' in processed['cc_recipients']
    assert 'bcc1@example.com' in processed['bcc_recipients']
    assert processed['body'] == 'Test body'
    assert processed['html_body'] == '<p>Test body</p>'
    assert processed['importance'] == 2
    assert processed['is_read'] is True
    assert processed['has_attachments'] is True
    assert processed['categories'] == ['Test', 'Important']
    assert processed['custom_field'] == 'custom_value'
    
    # Check dates were preserved
    assert processed['received_time'] == '2023-01-01T12:00:00+00:00'
    assert processed['sent_time'] == '2023-01-01T11:55:00+00:00'
    
    # Test with invalid input
    assert extractor._process_email_data(None) == {}
    assert extractor._process_email_data("not a dict") == {}
    assert extractor._process_email_data({}) != {}  # Should add default fields

def test_parse_date_ranges():
    """Test the parse_date_ranges method."""
    from outlook_extractor.extractor.outlook_extractor import OutlookExtractor
    from unittest.mock import patch
    
    extractor = OutlookExtractor()
    
    # Test with default config (should use days_back)
    with patch.object(extractor.config, 'get', side_effect=lambda section, key, *args: {
        ('date_range', 'date_ranges', ''): '',
        ('date_range', 'days_back', '30'): '30'
    }.get((section, key, args[0] if args else ''))):
        start_date, end_date = extractor.parse_date_ranges()
        assert start_date < end_date
        assert (end_date - start_date).days == 30
    
    # Test with specific date range
    with patch.object(extractor.config, 'get', side_effect=lambda section, key, *args: {
        ('date_range', 'date_ranges', ''): '01/2023,03/2023',
        ('date_range', 'days_back', '30'): '30'
    }.get((section, key, args[0] if args else ''))):
        start_date, end_date = extractor.parse_date_ranges()
        assert start_date.year == 2023
        assert start_date.month == 1
        assert end_date.year == 2023
        assert end_date.month == 3
        assert end_date.day > 27  # Last day of March
    
    # Test with invalid date range (should fall back to days_back)
    with patch.object(extractor.config, 'get', side_effect=lambda section, key, *args: {
        ('date_range', 'date_ranges', ''): 'invalid,date',
        ('date_range', 'days_back', '30'): '30'
    }.get((section, key, args[0] if args else ''))):
        start_date, end_date = extractor.parse_date_ranges()
        assert (end_date - start_date).days == 30

def test_extract_emails_with_date_range():
    """Test email extraction with date range filtering."""
    # Import here to avoid circular imports
    from outlook_extractor.extractor.outlook_extractor import OutlookExtractor
    
    # Setup
    mock_client = MagicMock()
    mock_folder = MagicMock()
    mock_folder.name = "Inbox"
    mock_client.get_all_folders.return_value = [mock_folder]
    mock_client.get_emails.return_value = [
        {"id": "1", "subject": "Test Email 1", "received_time": datetime(2023, 1, 15)},
        {"id": "2", "subject": "Test Email 2", "received_time": datetime(2023, 2, 15)}
    ]
    
    mock_storage = MagicMock()
    mock_storage.save_emails.return_value = 2
    
    # Create extractor and inject mocks
    with patch('outlook_extractor.extractor.outlook_extractor.OutlookClient', return_value=mock_client):
        extractor = OutlookExtractor()
        extractor.storage = mock_storage
        
        # Test with date range
        start_date = datetime(2023, 1, 1)
        end_date = datetime(2023, 1, 31)
        
        result = extractor.extract_emails(
            ["Inbox"], 
            start_date=start_date, 
            end_date=end_date
        )
        
        # Verify
        assert result["success"] is True
        assert result["emails_processed"] == 1  # Only one email is within the range
        
        # Verify the date filter was applied
        called_kwargs = mock_client.get_emails.call_args[1]
        assert called_kwargs["start_date"] == start_date
        assert called_kwargs["end_date"] == end_date

def test_extract_emails_duplicate_folders(extractor):
    """Test that the same folder isn't processed multiple times."""
    # Test with multiple patterns that all match the same folder
    result = extractor.extract_emails(
        ["Inbox", "Inb*", "*box"],
        datetime(2023, 1, 1, tzinfo=timezone.utc),
        datetime(2023, 12, 31, tzinfo=timezone.utc)
    )
    
    # Verify that even though we matched the same folder multiple times,
    # it was only processed once
    assert result["success"] is True
    assert result["folders_processed"] == 1  # Only one unique folder
    assert result["emails_processed"] == 2  # 2 emails per folder
    assert result["emails_saved"] == 2
    
    # Verify the emails were only saved once
    assert len(extractor.storage.saved_emails) == 2

def test_close_after_extraction(extractor):
    """Test that the extractor is properly closed after extraction."""
    # Test with multiple patterns that all match the same folder
    result = extractor.extract_emails(
        ["Inbox", "Inb*", "*box"],
        datetime(2023, 1, 1, tzinfo=timezone.utc),
        datetime(2023, 12, 31, tzinfo=timezone.utc)
    )
    
    # Verify that the extractor was properly closed
    assert extractor.outlook_client is None
    assert extractor.storage is None
