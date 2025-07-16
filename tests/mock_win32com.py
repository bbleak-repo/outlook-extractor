"""
Mock implementation of win32com for testing on non-Windows platforms.

This module provides mock implementations of the win32com.client and related modules
that are normally only available on Windows. This allows tests to run on any platform.
"""
import sys
from unittest.mock import MagicMock

# Create mock for win32com
win32com = MagicMock()
win32com.client = MagicMock()
win32com.client.Dispatch.return_value = MagicMock()

# Create mock for pythoncom
pythoncom = MagicMock()

# Add to sys.modules so they can be imported
sys.modules['win32com'] = win32com
sys.modules['win32com.client'] = win32com.client
sys.modules['pythoncom'] = pythoncom

# Mock MAPI constants
class Constants:
    olFolderInbox = 6
    olFolderSentMail = 5
    olMailItem = 0
    olAppointmentItem = 1
    olContactItem = 2
    olTaskItem = 3
    olJournalItem = 4
    olNoteItem = 5
    olPostItem = 6
    olDistributionList = 7
    olPublicFoldersAllPublicFolders = 18

# Add constants to win32com.client
win32com.client.constants = Constants()

# Create a mock for the Outlook Application
class MockOutlookApplication:
    def __init__(self):
        self.Session = MagicMock()
        self.Session.CurrentUser = MagicMock()
        self.Session.CurrentUser.Address = "test@example.com"
        self.Session.GetDefaultFolder = MagicMock(return_value=MagicMock())
        
        # Set up a mock for the Namespace
        self.Session.GetNamespace.return_value = self.Session
        
        # Set up a mock for the Inbox folder
        self.inbox = MagicMock()
        self.inbox.Name = "Inbox"
        self.inbox.FolderPath = "\\Inbox"
        self.inbox.Items = MagicMock()
        self.inbox.Folders = MagicMock()
        
        # Return the Inbox when GetDefaultFolder is called with olFolderInbox
        self.Session.GetDefaultFolder.side_effect = lambda x: self.inbox if x == 6 else MagicMock()
        
        # Set up a mock for the Items collection
        self.mock_items = MagicMock()
        self.mock_items.Count = 0
        self.mock_items.Item = MagicMock(side_effect=IndexError("Index out of range"))
        self.inbox.Items = self.mock_items

# Set up the mock Dispatch to return our mock Outlook application
win32com.client.Dispatch.return_value = MockOutlookApplication()

# Create a mock for the MailItem
class MockMailItem:
    def __init__(self):
        self.Subject = "Test Email"
        self.SenderName = "Test Sender"
        self.SenderEmailAddress = "sender@example.com"
        self.To = "recipient@example.com"
        self.CC = ""
        self.BCC = ""
        self.Body = "Test email body"
        self.HTMLBody = "<p>Test email body</p>"
        self.SentOn = "2023-01-01 12:00:00"
        self.ReceivedTime = "2023-01-01 12:05:00"
        self.Categories = ""
        self.Importance = 1  # Normal importance
        self.Sensitivity = 0  # Normal sensitivity
        self.Attachments = MagicMock()
        self.Attachments.Count = 0
        self.UnRead = True
        self.FlagStatus = 0  # Not flagged
        self.FlagRequest = ""
        self.FlagDueBy = ""
        self.FlagIcon = 0
        self.FlagMarkedAsTask = False
        self.ConversationID = "CONVERSATION_ID_123"
        self.ConversationTopic = "Test Email"
        self.EntryID = "ENTRY_ID_123"
        self.Parent = MagicMock()
        self.Parent.FolderPath = "\\Inbox"
        
        # PropertyAccessor mock
        self.PropertyAccessor = MagicMock()
        self.PropertyAccessor.GetProperty.side_effect = lambda x: {
            'http://schemas.microsoft.com/mapi/proptag/0x0FFF0102': self.EntryID,
            'http://schemas.microsoft.com/mapi/proptag/0x30130102': self.ConversationID,
            'http://schemas.microsoft.com/mapi/proptag/0x0E1D001E': self.SentOn,
            'http://schemas.microsoft.com/mapi/proptag/0x0E060040': self.ReceivedTime
        }.get(x, None)

# Add a method to add test emails to the mock
MockOutlookApplication.add_test_email = lambda self, count=1: [
    self.mock_items.Item.return_value for _ in range(count)
]
