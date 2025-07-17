"""Tests for the Outlook extractor functionality."""
import fnmatch
import json
import os
import pytest
import sys
import unittest
from datetime import datetime, timedelta, timezone
from unittest.mock import MagicMock, patch, call
from pathlib import Path

# Mock Windows-specific modules
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()
sys.modules['pythoncom'] = MagicMock()

# Add the project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

# Now import the extractor after setting up mocks
from outlook_extractor.extractor.outlook_extractor import OutlookExtractor
from outlook_extractor.config import ConfigManager

class MockOutlookClient:
    """Mock Outlook client for testing."""
    
    def __init__(self):
        # Create test emails with threading information
        email1 = MagicMock()
        email1.EntryID = 'email1'
        email1.Subject = 'Test Email 1'
        email1.SenderEmailAddress = 'sender@example.com'
        email1.SenderName = 'Sender'
        email1.To = 'recipient@example.com'
        email1.CC = ''
        email1.Body = 'This is a test email 1'
        email1.ReceivedTime = datetime.now(timezone.utc) - timedelta(hours=1)
        email1.SentOn = datetime.now(timezone.utc) - timedelta(hours=1, minutes=5)
        email1.UnRead = False
        email1.HasAttachments = False
        email1.Categories = ''
        email1.InReplyTo = None
        email1.ConversationID = 'conv1'
        email1.ConversationIndex = 'idx1'
        
        email2 = MagicMock()
        email2.EntryID = 'email2'
        email2.Subject = 'Re: Test Email 1'
        email2.SenderEmailAddress = 'recipient@example.com'
        email2.SenderName = 'Recipient'
        email2.To = 'sender@example.com'
        email2.CC = ''
        email2.Body = 'This is a reply to test email 1'
        email2.ReceivedTime = datetime.now(timezone.utc)
        email2.SentOn = datetime.now(timezone.utc) - timedelta(minutes=5)
        email2.UnRead = False
        email2.HasAttachments = False
        email2.Categories = ''
        # Set up the InReplyTo as a mock that returns 'email1'
        email2.InReplyTo = 'email1'  # Directly set the value to 'email1'
        email2.ConversationID = 'conv1'
        email2.ConversationIndex = 'idx1'
        
        self.folders = {
            'Inbox': [email1, email2],
            'Sent Items': [
                {
                    'id': 'msg3',
                    'subject': 'Test Email 2',
                    'sender': 'me@example.com',
                    'to': 'test2@example.com',
                    'body': 'This is another test email',
                    'receivedDateTime': (datetime.now(timezone.utc) - timedelta(hours=6)).isoformat(),
                    'conversationId': 'conv2',
                    'inReplyTo': None,
                    'hasAttachments': False,
                    'isRead': True,
                    'categories': []
                }
            ]
        }
        
        # Create MagicMock for methods we want to track
        self.GetDefaultFolder = MagicMock(side_effect=self._get_default_folder)
        self.GetNamespace = MagicMock(return_value=self)
        self.Folders = MagicMock()
        self.Folders.Item.return_value = self._get_folder('Inbox')
        
    def _get_folder(self, folder_name):
        """Get a folder by name."""
        folder = MagicMock()
        folder.Name = folder_name
        folder.Items = self.folders.get(folder_name, [])
        return folder
        
    def _get_default_folder(self, folder_type):
        """Get a default folder by type."""
        if folder_type == 6:  # olFolderInbox
            return self._get_folder('Inbox')
        elif folder_type == 5:  # olFolderSentMail
            return self._get_folder('Sent Items')
        return None

@pytest.fixture
def mock_config():
    """Create a mock ConfigManager for testing."""
    config = MagicMock()
    
    # Configure the config mock to return test values
    def mock_get(section, key, default=None):
        # Handle the case where section and key are combined with a dot
        if '.' in section:
            section, key = section.split('.', 1)
            
        if section == 'storage' and key == 'type':
            return 'sqlite'
        elif section == 'storage' and key == 'sqlite_path':
            return ':memory:'
        elif section == 'extraction' and key == 'include_folders':
            return 'Inbox, Sent Items'
        elif section == 'extraction' and key == 'batch_size':
            return '50'
        elif section == 'export' and key == 'output_dir':
            return './exports'
        elif section == 'email_processing' and key == 'priority_emails':
            return 'admin@example.com, important@example.com'
        elif section == 'email_processing' and key == 'admin_emails':
            return 'admin@example.com'
        elif section == 'threading' and key == 'enabled':
            return 'true'
        elif section == 'logging' and key == 'level':
            return 'INFO'
        return default
    
    config.get.side_effect = mock_get
    return config

# Patching at module level to avoid multiple patches
import sys

# Check if win32com is already mocked to avoid duplicate patches
if 'win32com' not in sys.modules:
    import unittest.mock as mock
    sys.modules['win32com'] = mock.MagicMock()
    sys.modules['win32com.client'] = mock.MagicMock()
    sys.modules['pythoncom'] = mock.MagicMock()

@pytest.fixture
def mock_outlook_extractor(mocker):
    """Fixture that creates a mock OutlookExtractor with mocked dependencies."""
    # Create a mock config with default values
    mock_config = MagicMock()
    
    def config_get_side_effect(*args, **kwargs):
        # Handle both dot notation (key) and separate section/key (section, key)
        if len(args) == 1:  # Called with dot notation (key)
            if '.' in args[0]:
                section, key = args[0].split('.', 1)
            else:
                section = None
                key = args[0]
            default = kwargs.get('default', '')
            fallback = kwargs.get('fallback', '')
        else:  # Called with separate section and key (section, key)
            section = args[0] if len(args) > 0 else None
            key = args[1] if len(args) > 1 else None
            default = args[2] if len(args) > 2 else kwargs.get('default', '')
            fallback = args[2] if len(args) > 2 else kwargs.get('fallback', '')
        
        # Handle special cases based on section and key
        if section == 'date_range':
            if key == 'days_back':
                return '14'  # Default to 14 days back
            return default or fallback or ''
        
        # Handle email_processing section
        if section == 'email_processing':
            if key == 'priority_emails':
                return 'priority1@example.com, priority2@example.com'
            if key == 'admin_emails':
                return 'admin@example.com'
        
        # Handle subject filter
        if key == 'subject_contains':
            return 'Test Email 1'
            
        # Default return
        return default or fallback or ''
    
    mock_config.get.side_effect = config_get_side_effect
    
    # Create a mock database
    mock_db = MagicMock()
    mock_db.save_email = MagicMock(return_value=True)
    mock_db.save_emails = MagicMock()
    mock_db.get_all_emails = MagicMock(return_value=[])
    mock_db.get_email = MagicMock(return_value=None)
    
    # Create a mock Outlook client
    mock_client = MockOutlookClient()
    
    # Patch the ConfigManager to return our mock config
    with patch('outlook_extractor.extractor.outlook_extractor.ConfigManager') as mock_config_manager:
        # Configure the mock to return our mock config
        mock_config_instance = MagicMock()
        mock_config_instance.get.side_effect = config_get_side_effect
        mock_config_manager.return_value = mock_config_instance
        
        # Create the extractor with the mocks
        extractor = OutlookExtractor()
        
        # Set up the rest of the mocks
        extractor.config = mock_config_instance
        extractor.mock_db = mock_db
        extractor.mock_client = mock_client
        
        # Patch the storage to return our mock_db
        mocker.patch.object(extractor, 'storage', mock_db)
        
        # Patch the client to return our mock_client
        mocker.patch.object(extractor, '_outlook_client', mock_client, create=True)
        
        yield extractor

class TestOutlookExtractor:
    """Test cases for the OutlookExtractor class."""
    
    def test_initialization(self, mock_outlook_extractor):
        """Test that the extractor initializes correctly."""
        extractor = mock_outlook_extractor
        assert extractor is not None
        assert extractor.outlook_client is not None
        
        # Verify the mock client was called
        extractor.mock_client.GetNamespace.assert_called_once_with('MAPI')
        
        # Verify storage was initialized
        assert extractor.mock_db is not None
        
        # Verify config was loaded
        assert extractor.config is not None
    
    def test_extract_emails(self, mock_outlook_extractor):
        """Test extracting emails from mock folders."""
        extractor = mock_outlook_extractor
        
        # Configure the mock to return test emails
        inbox_emails = extractor.mock_client.folders['Inbox']
        sent_emails = extractor.mock_client.folders['Sent Items']
        
        # Mock the folder iteration
        def get_folder_emails(folder_name):
            if 'Inbox' in folder_name:
                return inbox_emails
            elif 'Sent' in folder_name:
                return sent_emails
            return []
            
        # Mock the folder iteration
        def folder_iter_side_effect():
            for folder_name in ['Inbox', 'Sent Items']:
                folder = MagicMock()
                folder.Name = folder_name
                folder.Items = get_folder_emails(folder_name)
                yield folder
                
        extractor._iterate_folders = MagicMock(side_effect=folder_iter_side_effect)
        
        # Mock the config to return expected values
        extractor.config.get.side_effect = lambda section, key, **kwargs: ''
        
        # Run extraction with folder patterns that match our test folders
        result = extractor.extract_emails(folder_patterns=['Inbox', 'Sent Items'])
        
        # Verify results
        assert 'emails_processed' in result
        assert 'emails_saved' in result
        
        # Verify save_email was called the expected number of times
        assert extractor.mock_db.save_email.call_count > 0
        
        # Verify the structure of the saved emails
        for call in extractor.mock_db.save_email.call_args_list:
            email_data = call[0][0]  # First argument to save_email
            assert 'id' in email_data
            assert 'subject' in email_data
            assert 'sender' in email_data
            assert 'receivedDateTime' in email_data
    
    def test_thread_processing(self, mock_outlook_extractor, capsys, mocker):
        """Test email extraction with thread processing."""
        # Enable debug logging
        import logging
        logging.basicConfig(level=logging.DEBUG)
        logger = logging.getLogger(__name__)
        
        extractor = mock_outlook_extractor
        logger.debug("Initialized extractor")
        
        # Configure the mock to return test emails with threading info
        from datetime import datetime, timezone, timedelta
        
        # Create test emails with conversation IDs for threading
        now = datetime.now(timezone.utc)
        email1_time = now - timedelta(hours=2)
        email2_time = now - timedelta(hours=1)
        
        # Create mock attachments
        attachments1 = MagicMock()
        attachments1.Count = 0
        attachments2 = MagicMock()
        attachments2.Count = 0
        
        test_emails = [
            MagicMock(**{
                'id': 'msg1',
                'subject': 'Test Email',
                'sender': 'test@example.com',
                'senderEmailAddress': 'test@example.com',
                'receivedDateTime': email1_time,
                'conversationId': 'conv1',
                'body': 'Test email body',
                'to': 'recipient@example.com',
                'cc': '',
                'bcc': '',
                'ReceivedTime': email1_time,
                'Subject': 'Test Email 1',
                'SenderName': 'sender@example.com',
                'To': 'recipient@example.com',
                'Body': 'This is a test email 1',
                'EntryID': 'email1',
                'SaveAs': MagicMock(),
                'Close': MagicMock(),
                'UnRead': False,
                'HasAttachments': False,
                'Attachments': attachments1
            }),
            MagicMock(**{
                'id': 'msg2',
                'subject': 'Re: Test Email',
                'sender': 'reply@example.com',
                'senderEmailAddress': 'reply@example.com',
                'receivedDateTime': email2_time,
                'conversationId': 'conv1',
                'body': 'Reply body',
                'to': 'test@example.com',
                'cc': '',
                'bcc': '',
                'inReplyTo': 'msg1',
                'references': '<msg1@example.com>',
                'threadIndex': 'thread1',
                'hasAttachments': False,
                'isRead': True,
                'importance': 'normal',
                'ReceivedTime': email2_time,
                'ConversationID': 'conv1',
                'ConversationTopic': 'Test Email',
                'MessageClass': 'IPM.Note',
                'Subject': 'Re: Test Email 1',
                'SenderName': 'recipient@example.com',
                'To': 'sender@example.com',
                'Body': 'This is a reply to test email 1',
                'EntryID': 'email2',
                'InReplyTo': 'email1',
                'SaveAs': MagicMock(),
                'Close': MagicMock(),
                'UnRead': False,
                'HasAttachments': False,
                'Attachments': attachments2
            })
        ]
        
        # Create a mock namespace and root folder
        namespace = MagicMock()
        logger.debug("Created mock namespace")
        
        # Create mock email items
        email1, email2 = test_emails
        
        # Create a mock inbox folder with items
        inbox_folder = MagicMock()
        inbox_folder.Name = 'Inbox'
        inbox_folder.DefaultItemType = 0  # MailItem
        
        # Create a mock items collection for the inbox
        inbox_items = MagicMock()
        inbox_items.Count = len(test_emails)
        
        # Set up __getitem__ to return test emails
        def get_item(index):
            if index < len(test_emails):
                return test_emails[index]
            raise IndexError("list index out of range")
            
        inbox_items.__getitem__.side_effect = get_item
        
        # Set up __iter__ to iterate over test_emails
        inbox_items.__iter__.side_effect = lambda: iter(test_emails)
        
        # Add Sort and Restrict methods to the items collection
        def sort_items(property, descending):
            # Sort the test_emails by ReceivedTime in descending order
            def get_received_time(email):
                if hasattr(email, 'ReceivedTime'):
                    return email.ReceivedTime
                elif hasattr(email, 'get'):
                    return email.get('ReceivedTime', '')
                return ''
                
            test_emails.sort(key=get_received_time, reverse=descending)
            
        inbox_items.Sort = MagicMock(side_effect=sort_items)
        inbox_items.Restrict = MagicMock(return_value=inbox_items)  # Return self for chaining
        
        # Set up the Items property
        inbox_folder.Items = inbox_items
        
        # Make sure the folder has a Name attribute
        inbox_folder.Name = 'Inbox'
        
        # Create a mock account with Folders collection
        account = MagicMock()
        account.Name = 'Test Account'
        
        # Set up the Folders property to return a list with our inbox folder
        account.Folders = [inbox_folder]
        
        # Set up the root folder (Session.Folders)
        root_folder = MagicMock()
        root_folder.__iter__.return_value = [account]  # This makes list(root_folder) work
        
        # Mock the outlook client and its methods
        extractor._outlook_client = MagicMock()
        extractor._outlook_client.GetNamespace.return_value = namespace
        
        # Set up the session and root folders
        session = MagicMock()
        session.Folders = [account]
        extractor._outlook_client.Session = session
        
        # Mock the GetDefaultFolder method on the namespace to return our inbox folder
        namespace.GetDefaultFolder.return_value = inbox_folder
        
        # Mock the Folders property to return our account's folders
        namespace.Folders = [account]
        
        # Mock the _find_matching_folders method to return our test folder
        def find_matching_folders_side_effect(folder, patterns, current_path="", recursive=True):
            logger.debug(f"_find_matching_folders called with patterns: {patterns}, current_path: {current_path}, recursive: {recursive}")
            logger.debug(f"Folder type: {type(folder)}, dir: {[a for a in dir(folder) if not a.startswith('__')]}")
            
            # If this is the account folder, return the inbox folder
            if folder == account:
                logger.debug("Returning account folders")
                return [(inbox_folder, 'Inbox')]
            
            # If this is the inbox folder, return it if it matches the pattern
            if folder == inbox_folder and any(fnmatch.fnmatch('Inbox', p.lower()) for p in patterns):
                logger.debug("Returning inbox folder")
                return [(inbox_folder, 'Inbox')]
                
            logger.debug("No matching folders found")
            return []
            
        # Make sure we don't accidentally append to the folders list
        extractor._find_matching_folders = MagicMock(side_effect=find_matching_folders_side_effect)
            
        # Replace the actual method with our mock
        extractor._find_matching_folders = MagicMock(side_effect=find_matching_folders_side_effect)
        
        # Also mock the folder_matches_pattern method to work with our test data
        def folder_matches_pattern_side_effect(folder_name, patterns):
            logger.debug(f"folder_matches_pattern called with name: {folder_name}, patterns: {patterns}")
            if not folder_name or not patterns:
                return False
            for pattern in patterns:
                if fnmatch.fnmatch(folder_name.lower(), pattern.lower()):
                    return True
            return False
            
        extractor.folder_matches_pattern = MagicMock(side_effect=folder_matches_pattern_side_effect)
        
        # Mock the is_mail_folder method to return True for our inbox folder
        def is_mail_folder_side_effect(folder):
            logger.debug(f"is_mail_folder called with folder: {folder}")
            # Only return True for our inbox folder
            if hasattr(folder, 'Name') and folder.Name == 'Inbox':
                return True
            if hasattr(folder, 'DefaultItemType') and folder.DefaultItemType == 0:
                return True
            return False
            
        extractor.is_mail_folder = MagicMock(side_effect=is_mail_folder_side_effect)
        
        # Mock the config to return expected values
        def config_get_side_effect(section, key, **kwargs):
            if section == 'date_range' and key == 'days_back':
                return '14'
            if section == 'email_processing' and key == 'priority_emails':
                return 'priority1@example.com, priority2@example.com'
            if section == 'email_processing' and key == 'admin_emails':
                return 'admin@example.com'
            if key == 'include_threads':
                return 'true'
            return ''
            
        extractor.config.get.side_effect = config_get_side_effect
        
        # Mock the _process_email_data method to return the email as-is
        def process_email_side_effect(email_data):
            logger.debug(f"_process_email_data called with: {email_data}")
            return email_data
            
        extractor._process_email_data = MagicMock(side_effect=process_email_side_effect)
    
        # Reset the mock to track calls in this test
        extractor.mock_db.reset_mock()
        extractor.mock_db.save_email.return_value = True
        
        # Configure save_emails to return the number of emails saved
        extractor.mock_db.save_emails.return_value = len(test_emails)
        
        # Add debug print for save_email calls
        def save_email_side_effect(email):
            print(f"save_email called with: {email}")
            return True
            
        extractor.mock_db.save_email.side_effect = save_email_side_effect
        
        # Add debug print for save_emails calls
        def save_emails_side_effect(emails):
            print(f"save_emails called with {len(emails)} emails")
    
    def test_thread_processing(self, mock_outlook_extractor, capsys, mocker, tmp_path):
        """Test email extraction with thread processing."""
        # Set up the extractor with the mock database
        extractor = mock_outlook_extractor
        
        # Create a mock for attachments
        mock_attachments = MagicMock()
        mock_attachments.Count = 0
        
        # Create test emails with threading information
        email1 = MagicMock()
        email1.EntryID = 'email1'
        email1.Subject = 'Test Email 1'
        email1.SenderEmailAddress = 'sender@example.com'
        email1.SenderName = 'Sender'
        email1.To = 'recipient@example.com'
        email1.CC = ''
        email1.Body = 'This is a test email 1'
        email1.ReceivedTime = datetime.now(timezone.utc) - timedelta(hours=1)
        email1.SentOn = datetime.now(timezone.utc) - timedelta(hours=1, minutes=5)
        email1.UnRead = False
        email1.HasAttachments = False
        email1.Categories = ''
        email1.InReplyTo = None
        email1.ConversationID = 'conv1'
        email1.ConversationIndex = 'idx1'
        email1.Attachments = mock_attachments
        
        # Create email2 mock
        email2 = MagicMock()
        email2.EntryID = 'email2'
        email2.Subject = 'Re: Test Email 1'
        email2.SenderEmailAddress = 'recipient@example.com'
        email2.SenderName = 'Recipient'
        email2.To = 'sender@example.com'
        email2.CC = ''
        email2.Body = 'This is a reply to test email 1'
        email2.ReceivedTime = datetime.now(timezone.utc)
        email2.SentOn = datetime.now(timezone.utc) - timedelta(minutes=5)
        email2.UnRead = False
        email2.HasAttachments = False
        email2.Categories = ''
        email2.InReplyTo = 'email1'  # Directly set the value to 'email1'
        email2.ConversationID = 'conv1'
        email2.ConversationIndex = 'idx1'
        email2.Attachments = mock_attachments
        
        # Set up the mock folder with items
        mock_folder = MagicMock()
        mock_folder.Name = 'Inbox'
        
        # Create a list of emails for the mock items
        emails = [email1, email2]
        
        # Create a proper mock for the Items collection
        mock_items = MagicMock()
        mock_items.__iter__.return_value = emails  # For iteration
        # For indexing, adjust for 1-based indexing in Outlook
        mock_items.__getitem__.side_effect = lambda idx: emails[int(idx) - 1] if 1 <= int(idx) <= len(emails) else None
        mock_items.Count = len(emails)  # For Count property
        
        # Set up Restrict to return the same items collection
        mock_items.Restrict.return_value = mock_items
        
        # Assign the mock items to the folder
        mock_folder.Items = mock_items
        
        # Set up the mock namespace and folders
        mock_namespace = MagicMock()
        mock_account = MagicMock()
        mock_account.Name = 'Test Account'
        mock_account.Folders = [mock_folder]
        mock_namespace.Folders = [mock_account]
        
        # Configure the mock outlook client
        extractor.outlook_client.GetNamespace.return_value = mock_namespace
        
        # Set up the mock database
        extractor.mock_db.save_email.return_value = True
        extractor.mock_db.save_emails.return_value = 2
        
        # Patch the is_mail_folder method to always return True for our test folder
        with patch.object(extractor, 'is_mail_folder', return_value=True):
            # Call the method under test
            result = extractor.extract_emails(folder_patterns=['Inbox'])
        
        # Verify results
        assert 'emails_processed' in result
        assert 'emails_saved' in result
        
        # Verify the threading information was processed correctly
        # Check that save_emails was called with both emails
        assert extractor.mock_db.save_emails.call_count == 1, "Expected save_emails to be called once"
        
        # Get the saved emails from the call arguments
        saved_emails = extractor.mock_db.save_emails.call_args[0][0]
        assert len(saved_emails) == 2, f"Expected 2 emails to be saved, but {len(saved_emails)} were saved"
        
        # Print debug information about saved emails
        print("\nSaved emails:")
        for i, email in enumerate(saved_emails, 1):
            print(f"  Email {i}:")
            for key, value in email.items():
                print(f"    {key}: {value}")
        
        # Check that the second email has the correct in_reply_to field
        email2_data = next((e for e in saved_emails if e.get('entry_id') == 'email2'), None)
        assert email2_data is not None, "Could not find email2 in saved emails"
        
        # Debug: Print the email2_data to see what fields are available
        print("\nEmail2 data:")
        for key, value in email2_data.items():
            print(f"  {key}: {value}")
            
        # Verify the threading information
        assert 'in_reply_to' in email2_data, "in_reply_to field is missing from email2"
        assert email2_data['in_reply_to'] == 'email1', \
            f"Expected in_reply_to to be 'email1', got {email2_data.get('in_reply_to')}"
            
        # Verify the email was processed correctly
        assert email2_data['subject'] == 'Re: Test Email 1'
        assert email2_data['sender_email'] == 'recipient@example.com'
    
    def test_extract_with_filters(self, mock_outlook_extractor):
        """Test email extraction with filters."""
        extractor = mock_outlook_extractor
        
        # Configure the mock to return test emails with threading info
        test_emails = [
            {
                'id': 'msg1',
                'subject': 'Test Email 1',
                'sender': 'test@example.com',
                'senderEmailAddress': 'test@example.com',
                'receivedDateTime': '2025-01-01T00:00:00Z',
                'conversationId': 'conv1',
                'body': 'Test email body',
                'to': 'recipient@example.com',
                'cc': '',
                'bcc': '',
                'inReplyTo': None,
                'references': '',
                'threadIndex': 'thread1',
                'hasAttachments': False,
                'isRead': True,
                'importance': 'normal'
            },
            {
                'id': 'msg2',
                'subject': 'Other Email',
                'sender': 'other@example.com',
                'senderEmailAddress': 'other@example.com',
                'receivedDateTime': '2025-01-02T00:00:00Z',
                'conversationId': 'conv2',
                'body': 'Other email body',
                'to': 'recipient@example.com',
                'cc': '',
                'bcc': '',
                'inReplyTo': None,
                'references': '',
                'threadIndex': 'thread2',
                'hasAttachments': False,
                'isRead': True,
                'importance': 'normal'
            }
        ]
        
        # Mock the folder iteration
        def folder_iter_side_effect():
            folder = MagicMock()
            folder.Name = 'Inbox'
            folder.Items = test_emails
            yield folder
            
        extractor._iterate_folders = MagicMock(side_effect=folder_iter_side_effect)
        
        # Mock the config to return expected values
        def config_get_side_effect(section, key, **kwargs):
            if key == 'subject_contains':
                return 'Test Email 1'
            return ''
            
        extractor.config.get.side_effect = config_get_side_effect
        
        # Reset the mock to track calls in this test
        extractor.mock_db.reset_mock()
        extractor.mock_db.save_email.return_value = True
        
        # Extract with a subject filter
        result = extractor.extract_emails(
            folder_patterns=['Inbox']
        )
        
        # Verify results
        assert 'emails_processed' in result
        assert 'emails_saved' in result
        
        # Verify only the matching email was saved
        assert extractor.mock_db.save_email.call_count == 1
        
        # Verify the saved email matches the filter
        email_data = extractor.mock_db.save_email.call_args[0][0]
        assert 'Test Email 1' in email_data.get('subject', '')
