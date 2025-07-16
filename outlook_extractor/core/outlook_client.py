"""
Outlook client for interacting with Microsoft Outlook.

This module provides a high-level interface to interact with Microsoft Outlook
using the win32com.client library. It handles connection management, folder
navigation, and email retrieval.
"""
import os
import sys
import logging
from typing import List, Tuple, Optional, Dict, Any, Union
from datetime import datetime
import win32com.client
import pythoncom
from pathlib import Path

# Import config and logging
from ..config import get_config
from ..logging_setup import get_logger

# Get logger
logger = get_logger(__name__)

if sys.platform != 'win32':
    from .mock_outlook import MockOutlookClient as OutlookClient
else:
    import win32com.client
    import pythoncom
    from typing import List, Dict, Any, Optional
    from datetime import datetime, timedelta
    import re
    import logging
    from pathlib import Path
    
class OutlookClient:
    """Client for interacting with Microsoft Outlook."""
    
    def __init__(self, config=None):
        """Initialize the Outlook client.
        
        Args:
            config: Optional ConfigManager instance. If not provided, uses default config.
        """
        self.config = config or get_config()
        self.logger = logger
        self.outlook = None
        self.namespace = None
        self.account = None
        
    def connect(self) -> bool:
        """Connect to Microsoft Outlook.
        
        Returns:
            bool: True if the connection was successful, False otherwise.
        """
        try:
            # Initialize COM for the current thread
            pythoncom.CoInitialize()
            
            self.logger.info("Connecting to Microsoft Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Get the default account or the specified mailbox
            mailbox_name = self.config.get('outlook', 'mailbox_name', '').strip()
            if mailbox_name:
                try:
                    self.account = self.namespace.Folders[mailbox_name]
                    self.logger.info(f"Connected to mailbox: {mailbox_name}")
                except Exception as e:
                    self.logger.warning(
                        f"Could not access mailbox '{mailbox_name}'. "
                        f"Using default mailbox. Error: {e}"
                    )
                    self.account = self.namespace.GetDefaultFolder(6)  # Inbox as fallback
            else:
                self.account = self.namespace.GetDefaultFolder(6)  # Inbox
                self.logger.info("Connected to default mailbox")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to connect to Outlook: {e}")
            self.disconnect()
            return False
    
    def disconnect(self):
        """Disconnect from Microsoft Outlook and clean up resources."""
        try:
            if hasattr(self, 'outlook') and self.outlook is not None:
                # Release COM objects
                self.account = None
                self.namespace = None
                self.outlook = None
                
                # Uninitialize COM for this thread
                pythoncom.CoUninitialize()
                
                self.logger.info("Disconnected from Outlook")
        except Exception as e:
            self.logger.error(f"Error disconnecting from Outlook: {e}")
    
    def get_folders(self, folder_patterns: List[str] = None) -> List[Tuple[Any, str]]:
        """Get folders matching the specified patterns.
        
        Args:
            folder_patterns: List of folder patterns to match. If None, uses patterns from config.
            
        Returns:
            List of tuples containing (folder_object, folder_path) for matching folders.
        """
        if folder_patterns is None:
            folder_patterns = self.config.get_list('outlook', 'folder_patterns', ['Inbox'])
        
        if not folder_patterns:
            self.logger.warning("No folder patterns specified in config")
            return []
        
        # Get all folders and match against patterns
        all_folders = self._get_all_folders()
        matching_folders = []
        
        for folder, path in all_folders:
            if self._folder_matches_patterns(path, folder_patterns):
                matching_folders.append((folder, path))
        
        if not matching_folders:
            self.logger.warning(f"No folders matched the patterns: {', '.join(folder_patterns)}")
        else:
            self.logger.info(f"Found {len(matching_folders)} matching folders")
            
        return matching_folders
    
    def get_all_folders(self) -> List[Any]:
        """Get all folders from the Outlook account.
        
        Returns:
            List of folder objects
        """
        try:
            if not self.namespace:
                self.connect()
                
            folders = []
            self._collect_folders_recursive(self.account, folders)
            return folders
        except Exception as e:
            self.logger.error(f"Error getting all folders: {e}")
            return []
            
    def _collect_folders_recursive(self, folder, folder_list: List[Any], current_path: str = "") -> None:
        """Recursively collect all subfolders.
        
        Args:
            folder: Current folder object
            folder_list: List to collect folders in
            current_path: Current folder path (used internally for recursion)
        """
        try:
            # Add the current folder to the list
            if hasattr(folder, 'Name'):
                folder_list.append(folder)
                
            # Recursively process subfolders
            if hasattr(folder, 'Folders'):
                for subfolder in folder.Folders:
                    try:
                        self._collect_folders_recursive(subfolder, folder_list, 
                                                     f"{current_path}/{folder.Name}" if current_path else folder.Name)
                    except Exception as e:
                        self.logger.warning(f"Error processing subfolder: {e}")
                        continue
                        
        except Exception as e:
            self.logger.error(f"Error in _collect_folders_recursive: {e}")

    def _get_all_folders(self, folder=None, current_path="") -> List[Tuple[Any, str]]:
        """Recursively get all folders in the mailbox.
        
        Args:
            folder: The current folder to start from. If None, starts from the account root.
            current_path: The current folder path (used for recursion).
            
        Returns:
            List of tuples containing (folder_object, folder_path) for all folders.
        """
        folders = []
        
        try:
            if folder is None:
                if self.account is None:
                    self.logger.error("Not connected to Outlook")
                    return []
                folder = self.account
            
            # Get the folder name and build the path
            folder_name = folder.Name
            if current_path:
                folder_path = f"{current_path}/{folder_name}"
            else:
                folder_path = folder_name
            
            # Add the current folder
            folders.append((folder, folder_path))
            
            # Recursively process subfolders
            try:
                for subfolder in folder.Folders:
                    folders.extend(self._get_all_folders(subfolder, folder_path))
            except AttributeError:
                # Some folders don't have a Folders attribute
                pass
                
        except Exception as e:
            self.logger.error(f"Error getting folders: {e}")
        
        return folders
    
    def _folder_matches_patterns(self, folder_path: str, patterns: List[str]) -> bool:
        """Check if a folder path matches any of the given patterns.
        
        Args:
            folder_path: The folder path to check.
            patterns: List of patterns to match against.
            
        Returns:
            bool: True if the folder path matches any pattern, False otherwise.
        """
        if not patterns:
            return False
            
        folder_path_lower = folder_path.lower()
        
        for pattern in patterns:
            pattern = pattern.strip()
            if not pattern:
                continue
                
            # Handle exact match
            if folder_path_lower == pattern.lower():
                return True
                
            # Handle wildcard patterns
            if '*' in pattern or '?' in pattern:
                import fnmatch
                if fnmatch.fnmatch(folder_path_lower, pattern.lower()):
                    return True
            
            # Handle partial matches (if the pattern is part of the path)
            elif pattern.lower() in folder_path_lower:
                return True
        
        return False
    
    def get_emails(self, folder, start_date=None, end_date=None, max_emails: int = None) -> List[Dict[str, Any]]:
        """Get emails from the specified folder.
        
        Args:
            folder: The Outlook folder to get emails from.
            start_date: Optional start date to filter emails.
            end_date: Optional end date to filter emails.
            max_emails: Maximum number of emails to retrieve. If None, uses config value.
            
        Returns:
            List of email dictionaries with extracted information.
        """
        if max_emails is None:
            max_emails = self.config.get_int('outlook', 'max_emails', 1000)
        
        emails = []
        count = 0
        
        try:
            self.logger.info(f"Retrieving emails from folder: {folder.Name}")
            
            # Get all items in the folder
            items = folder.Items
            
            # Sort by received time (newest first)
            items.Sort("[ReceivedTime]", True)
            
            # Apply date filter if specified
            if start_date or end_date:
                filter_str = []
                if start_date:
                    filter_str.append(f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}'")
                if end_date:
                    filter_str.append(f"[ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'")
                
                filter_str = " AND ".join(filter_str)
                filtered_items = items.Restrict(filter_str)
            else:
                filtered_items = items
            
            # Process emails
            for item in filtered_items:
                try:
                    if item.Class == 43:  # 43 is the class for MailItem
                        email = self._extract_email_info(item)
                        if email:
                            emails.append(email)
                            count += 1
                            
                            if count % 100 == 0:
                                self.logger.info(f"Processed {count} emails...")
                            
                            if max_emails > 0 and count >= max_emails:
                                self.logger.info(f"Reached maximum number of emails to process ({max_emails})")
                                break
                                
                except Exception as e:
                    self.logger.error(f"Error processing email: {e}")
                    continue
            
            self.logger.info(f"Retrieved {count} emails from folder: {folder.Name}")
            
        except Exception as e:
            self.logger.error(f"Error retrieving emails: {e}")
        
        return emails
    
    def _extract_email_info(self, mail_item) -> Optional[Dict[str, Any]]:
        """Extract information from an email message.
        
        Args:
            mail_item: The Outlook MailItem object.
            
        Returns:
            Dictionary containing extracted email information, or None if an error occurs.
        """
        try:
            # Basic email information
            email = {
                'entry_id': mail_item.EntryID,
                'subject': mail_item.Subject or '(No Subject)',
                'sender_name': mail_item.SenderName if hasattr(mail_item, 'SenderName') else 'Unknown',
                'sender_email': mail_item.SenderEmailAddress if hasattr(mail_item, 'SenderEmailAddress') else 'unknown@example.com',
                'to_recipients': self._get_recipients(mail_item, 'To'),
                'cc_recipients': self._get_recipients(mail_item, 'CC'),
                'bcc_recipients': self._get_recipients(mail_item, 'BCC'),
                'received_time': self._parse_outlook_date(mail_item.ReceivedTime),
                'sent_time': self._parse_outlook_date(mail_item.SentOnBehalfOf if hasattr(mail_item, 'SentOnBehalfOf') else mail_item.SentOnBehalfOf),
                'body': mail_item.Body or '',
                'html_body': mail_item.HTMLBody if hasattr(mail_item, 'HTMLBody') else '',
                'importance': self._get_importance(mail_item.Importance) if hasattr(mail_item, 'Importance') else 'Normal',
                'categories': mail_item.Categories.split(';') if hasattr(mail_item, 'Categories') and mail_item.Categories else [],
                'has_attachments': mail_item.Attachments.Count > 0 if hasattr(mail_item, 'Attachments') else False,
                'attachments': [],
                'conversation_id': mail_item.ConversationID if hasattr(mail_item, 'ConversationID') else '',
                'conversation_topic': mail_item.ConversationTopic if hasattr(mail_item, 'ConversationTopic') else '',
                'size': mail_item.Size if hasattr(mail_item, 'Size') else 0,
                'flags': self._get_email_flags(mail_item),
                'headers': self._get_email_headers(mail_item) if hasattr(mail_item, 'PropertyAccessor') else {},
            }
            
            # Process attachments if needed
            if self.config.get_boolean('email_processing', 'extract_attachments', False):
                email['attachments'] = self._process_attachments(mail_item)
            
            return email
            
        except Exception as e:
            self.logger.error(f"Error extracting email info: {e}")
            return None
    
    def _get_recipients(self, mail_item, recipient_type: str) -> List[Dict[str, str]]:
        """Get recipients of a specific type from an email.
        
        Args:
            mail_item: The Outlook MailItem object.
            recipient_type: Type of recipients to get ('To', 'CC', 'BCC').
            
        Returns:
            List of dictionaries with recipient name and email.
        """
        recipients = []
        
        try:
            if not hasattr(mail_item, recipient_type):
                return recipients
                
            recipient_collection = getattr(mail_item, recipient_type)
            
            for recipient in recipient_collection.split(';'):
                recipient = recipient.strip()
                if not recipient:
                    continue
                    
                # Try to parse name and email (format: "Name <email@example.com>")
                if '<' in recipient and '>' in recipient:
                    name_part, email_part = recipient.rsplit('<', 1)
                    name = name_part.strip()
                    email = email_part.rstrip('>').strip()
                else:
                    name = recipient
                    email = recipient  # Use the full string as email if no angle brackets
                
                recipients.append({
                    'name': name,
                    'email': email
                })
                
        except Exception as e:
            self.logger.warning(f"Error processing {recipient_type} recipients: {e}")
        
        return recipients
    
    def _parse_outlook_date(self, date_value) -> Optional[str]:
        """Parse an Outlook date value into an ISO format string.
        
        Args:
            date_value: The date value from Outlook.
            
        Returns:
            ISO format date string, or None if the date is invalid.
        """
        if not date_value:
            return None
            
        try:
            # If it's already a datetime object, format it
            if hasattr(date_value, 'strftime'):
                return date_value.isoformat()
            # If it's a string, try to parse it
            elif isinstance(date_value, str):
                from dateutil.parser import parse
                return parse(date_value).isoformat()
            else:
                return None
        except Exception as e:
            self.logger.warning(f"Error parsing date {date_value}: {e}")
            return None
    
    def _get_importance(self, importance_value: int) -> str:
        """Convert importance value to string.
        
        Args:
            importance_value: The importance value from Outlook.
            
        Returns:
            Importance as a string ('Low', 'Normal', 'High').
        """
        importance_map = {
            0: 'Low',
            1: 'Normal',
            2: 'High'
        }
        return importance_map.get(importance_value, 'Normal')
    
    def _get_email_flags(self, mail_item) -> Dict[str, Any]:
        """Get email flags and status.
        
        Args:
            mail_item: The Outlook MailItem object.
            
        Returns:
            Dictionary with email flags and status.
        """
        flags = {
            'is_read': bool(getattr(mail_item, 'UnRead', True) is False),
            'is_flagged': bool(getattr(mail_item, 'FlagStatus', 0) > 0),
            'is_forwarded': bool(getattr(mail_item, 'IsMarkedAsTask', False)),
            'sensitivity': getattr(mail_item, 'Sensitivity', 0),
            'sensitivity_text': self._get_sensitivity_text(getattr(mail_item, 'Sensitivity', 0)),
            'has_been_forwarded': bool(getattr(mail_item, 'IsMarkedAsTask', False)),
            'has_been_replied': bool(getattr(mail_item, 'IsMarkedAsTask', False)),
        }
        
        # Additional flags that might be useful
        for attr in ['IsMarkedAsTask', 'IsConflict', 'NoAging', 'DownloadState']:
            if hasattr(mail_item, attr):
                flags[attr.lower()] = getattr(mail_item, attr)
        
        return flags
    
    def _get_sensitivity_text(self, sensitivity_value: int) -> str:
        """Convert sensitivity value to text.
        
        Args:
            sensitivity_value: The sensitivity value from Outlook.
            
        Returns:
            Sensitivity as a string ('Normal', 'Personal', 'Private', 'Confidential').
        """
        sensitivity_map = {
            0: 'Normal',
            1: 'Personal',
            2: 'Private',
            3: 'Confidential'
        }
        return sensitivity_map.get(sensitivity_value, 'Normal')
    
    def _get_email_headers(self, mail_item) -> Dict[str, str]:
        """Get email headers.
        
        Args:
            mail_item: The Outlook MailItem object.
            
        Returns:
            Dictionary with email headers.
        """
        headers = {}
        
        try:
            if not hasattr(mail_item, 'PropertyAccessor'):
                return headers
                
            property_accessor = mail_item.PropertyAccessor
            
            # Common headers to retrieve
            header_fields = [
                'Received', 'Received-SPF', 'Authentication-Results',
                'DKIM-Signature', 'X-Received', 'X-Google-DKIM-Signature',
                'X-Google-Smtp-Source', 'MIME-Version', 'X-Originating-IP',
                'Message-ID', 'References', 'In-Reply-To', 'Thread-Index'
            ]
            
            for field in header_fields:
                try:
                    value = property_accessor.GetProperty(f"http://schemas.microsoft.com/mapi/proptag/0x007D001F")
                    if value:
                        headers[field] = value
                except:
                    continue
            
            # Special handling for some headers
            try:
                # Get the full internet headers as a string
                internet_headers = property_accessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
                )
                if internet_headers:
                    headers['all_headers'] = internet_headers
            except:
                pass
                
        except Exception as e:
            self.logger.warning(f"Error getting email headers: {e}")
        
        return headers
    
    def _process_attachments(self, mail_item) -> List[Dict[str, Any]]:
        """Process email attachments.
        
        Args:
            mail_item: The Outlook MailItem object.
            
        Returns:
            List of dictionaries with attachment information.
        """
        attachments = []
        
        if not hasattr(mail_item, 'Attachments') or mail_item.Attachments.Count == 0:
            return attachments
        
        attachment_dir = self.config.get_attachment_dir()
        
        try:
            for i in range(1, mail_item.Attachments.Count + 1):
                try:
                    attachment = mail_item.Attachments.Item(i)
                    
                    # Skip embedded images in HTML
                    if hasattr(attachment, 'Position'):
                        continue
                    
                    # Get attachment info
                    attachment_info = {
                        'filename': attachment.FileName,
                        'size': attachment.Size,
                        'content_type': getattr(attachment, 'Type', 'application/octet-stream'),
                        'content_id': getattr(attachment, 'ContentID', ''),
                        'is_inline': getattr(attachment, 'IsInline', False),
                        'saved_path': None
                    }
                    
                    # Save the attachment if needed
                    if self.config.get_boolean('email_processing', 'extract_attachments', False):
                        try:
                            # Create a safe filename
                            safe_filename = self._sanitize_filename(attachment.FileName)
                            save_path = os.path.join(attachment_dir, safe_filename)
                            
                            # Handle duplicate filenames
                            counter = 1
                            base_name, ext = os.path.splitext(safe_filename)
                            while os.path.exists(save_path):
                                new_filename = f"{base_name}_{counter}{ext}"
                                save_path = os.path.join(attachment_dir, new_filename)
                                counter += 1
                            
                            # Save the attachment
                            attachment.SaveAsFile(save_path)
                            attachment_info['saved_path'] = save_path
                            
                        except Exception as e:
                            self.logger.error(f"Error saving attachment {attachment.FileName}: {e}")
                            attachment_info['error'] = str(e)
                    
                    attachments.append(attachment_info)
                    
                except Exception as e:
                    self.logger.error(f"Error processing attachment {i}: {e}")
                    continue
                    
        except Exception as e:
            self.logger.error(f"Error processing attachments: {e}")
        
        return attachments
    
    def _sanitize_filename(self, filename: str) -> str:
        """Sanitize a filename to remove invalid characters.
        
        Args:
            filename: The original filename.
            
        Returns:
            Sanitized filename.
        """
        import re
        # Remove invalid characters
        filename = re.sub(r'[\\/*?:"<>|]', '_', filename)
        # Remove control characters
        filename = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', filename)
        # Remove leading/trailing spaces and dots
        filename = filename.strip('. ')
        # Ensure the filename is not empty
        if not filename:
            filename = f"unnamed_attachment_{int(datetime.now().timestamp())}"
        return filename


# Example usage
if __name__ == "__main__":
    # Set up logging
    from ..logging_setup import setup_logging
    setup_logging()
    
    # Create and configure client
    client = OutlookClient()
    
    try:
        # Connect to Outlook
        if not client.connect():
            print("Failed to connect to Outlook")
            sys.exit(1)
        
        # Get matching folders
        folders = client.get_folders()
        print(f"Found {len(folders)} matching folders")
        
        # Process each folder
        for folder, path in folders:
            print(f"\nProcessing folder: {path}")
            
            # Get emails from the last 7 days
            end_date = datetime.now()
            start_date = end_date - timedelta(days=7)
            
            emails = client.get_emails(folder, start_date, end_date, max_emails=10)
            print(f"  Found {len(emails)} emails")
            
            # Print some info about each email
            for i, email in enumerate(emails[:3], 1):  # Show first 3 emails
                print(f"  {i}. {email['subject']} (From: {email['sender_name']} <{email['sender_email']}>)")
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Disconnect from Outlook
        client.disconnect()
