"""
Outlook Email Extractor

This module handles the extraction of emails from Microsoft Outlook with threading support.
"""
import os
import re
import fnmatch
import logging
import sqlite3
import json
import hashlib
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import List, Dict, Any, Optional, Set, Tuple, DefaultDict
from email.utils import getaddresses, parseaddr

from ..core.outlook_client import OutlookClient
from ..core.email_threading import ThreadManager, EmailThread, THREAD_STATUS_ACTIVE
from ..storage.base import EmailStorage
from ..storage.sqlite_storage import SQLiteStorage
from ..storage.json_storage import JSONStorage
from ..export.csv_exporter import CSVExporter
from ..config import ConfigManager

logger = logging.getLogger(__name__)

class OutlookExtractor:
    """Main class for extracting emails from Outlook with threading support."""
    
    def __init__(self, config_path: str = None):
        """Initialize the OutlookExtractor.
        
        Args:
            config_path: Path to the configuration file
        """
        self.config = ConfigManager(config_path) if config_path else ConfigManager()
        self._outlook_client = None  # Make it a private attribute
        self.storage = None
        self.csv_exporter = CSVExporter(self.config)
        self.thread_manager = ThreadManager()
        self.priority_addresses = set()
        self.admin_addresses = set()
        
        # Initialize storage only, outlook_client will be initialized on demand
        self._init_storage()
        self._load_config()
    
    @property
    def outlook_client(self):
        """Lazy initialization of the Outlook client."""
        if self._outlook_client is None:
            try:
                self._outlook_client = OutlookClient()
                logger.info("Outlook client initialized successfully")
            except Exception as e:
                error_msg = f"Failed to initialize Outlook client: {e}"
                logger.error(error_msg)
                raise RuntimeError(error_msg) from e
        return self._outlook_client
        
    @outlook_client.setter
    def outlook_client(self, value):
        """Set the Outlook client directly (primarily for testing)."""
        self._outlook_client = value
        
    def _init_outlook_client(self) -> None:
        """Initialize the Outlook client."""
        # This is now a no-op since we're using lazy initialization
        # The property getter will handle the actual initialization
        _ = self.outlook_client  # This will trigger initialization
    
    def _init_storage(self) -> None:
        """Initialize the storage backend."""
        storage_type = self.config.get('storage', 'type', 'sqlite')
        
        if storage_type.lower() == 'sqlite':
            db_path = self.config.get('storage', 'sqlite_path', 'emails.db')
            self.storage = SQLiteStorage(db_path)
        else:  # Default to JSON
            json_path = self.config.get('storage', 'json_path', 'emails.json')
            self.storage = JSONStorage(json_path)
        
        logger.info(f"Initialized {storage_type} storage at {self.storage.file_path}")
    
    def folder_matches_pattern(self, folder_name: str, patterns: List[str]) -> bool:
        """Check if folder name matches any of the patterns (supports wildcards).
        
        Args:
            folder_name: Name of the folder to check
            patterns: List of patterns to match against
            
        Returns:
            bool: True if the folder matches any pattern, False otherwise
        """
        if not patterns:
            return False
            
        # Clean up folder name - remove extra whitespace and normalize case
        clean_folder_name = folder_name.strip()
        folder_name_lower = clean_folder_name.lower()
        
        logger.debug(f"Checking if folder '{clean_folder_name}' matches patterns: {patterns}")
        
        # Normalize path separators
        normalized_name = folder_name_lower.replace('\\', '/')
        
        for pattern in patterns:
            if not pattern:  # Skip empty patterns
                continue
                
            # Clean and normalize pattern
            clean_pattern = pattern.strip()
            pattern_lower = clean_pattern.lower()
            normalized_pattern = pattern_lower.replace('\\', '/')
            
            # Check for exact match (case-insensitive)
            if folder_name_lower == pattern_lower:
                logger.debug(f"Exact match: '{clean_folder_name}' == '{clean_pattern}'")
                return True
                
            # Check for wildcard match
            if fnmatch.fnmatch(folder_name_lower, pattern_lower):
                logger.debug(f"Wildcard match: '{clean_folder_name}' matches pattern '{clean_pattern}'")
                return True
                
            # Check for partial match if pattern contains * or /
            if ('*' in pattern_lower or '/' in pattern_lower) and \
               fnmatch.fnmatch(folder_name_lower, f"*{pattern_lower}*"):
                logger.debug(f"Partial match: '{clean_folder_name}' contains pattern '{clean_pattern}'")
                return True
                
        logger.debug(f"Folder '{clean_folder_name}' did not match any patterns")
        return False
        
    def _find_matching_folders(self, folder, patterns: List[str], current_path: str = "", recursive: bool = True) -> List[tuple]:
        """Recursively find all folders matching the given patterns.
        
        Args:
            folder: The root folder to search from
            patterns: List of patterns to match against folder names
            current_path: Current folder path (used internally for recursion)
            recursive: Whether to search subfolders recursively
            
        Returns:
            List of tuples (folder_object, full_path) for matching folders
        """
        matching_folders = []
        
        try:
            # Debug folder object
            logger.debug(f"Folder object type: {type(folder)}")
            logger.debug(f"Folder object dir: {[a for a in dir(folder) if not a.startswith('__')]}")
            
            # Get the folder name - try both 'Name' and 'name' attributes
            folder_name = getattr(folder, 'Name', None) or getattr(folder, 'name', None)
            if folder_name is None:
                # Try to get the name using different methods
                if hasattr(folder, '__getitem__') and 'Name' in dir(folder):
                    folder_name = folder['Name']
                elif hasattr(folder, 'Name'):
                    folder_name = folder.Name
                elif hasattr(folder, 'name'):
                    folder_name = folder.name
                else:
                    folder_name = str(folder)
                    
            full_path = f"{current_path}/{folder_name}" if current_path else folder_name
            
            logger.debug(f"Checking folder: name='{folder_name}', path='{full_path}', patterns={patterns}")
            
            # Debug: Log folder attributes
            if hasattr(folder, 'DefaultItemType'):
                logger.debug(f"DefaultItemType: {getattr(folder, 'DefaultItemType', 'N/A')}")
            if hasattr(folder, 'Item'):
                logger.debug("Folder has 'Item' attribute")
            if hasattr(folder, 'Items'):
                logger.debug("Folder has 'Items' attribute")
            
            # Check if current folder matches any pattern
            if self.folder_matches_pattern(folder_name, patterns):
                logger.info(f"Found matching folder: {full_path} (matches patterns: {patterns})")
                matching_folders.append((folder, full_path))
            else:
                logger.debug(f"Folder '{folder_name}' did not match any patterns")
                # Debug: Log why it didn't match
                for pattern in patterns:
                    logger.debug(f"Checking pattern: {pattern}")
                    logger.debug(f"Folder name lower: {folder_name.lower()}")
                    logger.debug(f"Pattern lower: {pattern.lower()}")
                    logger.debug(f"Exact match: {folder_name.lower() == pattern.lower()}")
                    logger.debug(f"Wildcard match: {fnmatch.fnmatch(folder_name.lower(), pattern.lower())}")
            
            # Process subfolders if recursive is True
            if recursive:
                # Try to get subfolders, handling both 'Folders' and 'folders' attributes
                subfolders = []
                folders_attr = None
                
                if hasattr(folder, 'Folders'):
                    folders_attr = 'Folders'
                    subfolders = getattr(folder, 'Folders', [])
                elif hasattr(folder, 'folders'):
                    folders_attr = 'folders'
                    subfolders = getattr(folder, 'folders', [])
                
                if not subfolders:
                    logger.debug(f"No subfolders found in folder: {folder_name} (checked {folders_attr or 'no folder attributes'})")
                else:
                    logger.debug(f"Processing {len(subfolders)} subfolders of '{folder_name}' (from {folders_attr})")
                    
                    for i, subfolder in enumerate(subfolders, 1):
                        try:
                            # Try both 'Name' and 'name' attributes for subfolder name
                            subfolder_name = getattr(subfolder, 'Name', None) or getattr(subfolder, 'name', f'Unknown_{i}')
                            logger.debug(f"Processing subfolder {i}/{len(subfolders)}: {subfolder_name}")
                            
                            # Recursively process subfolder
                            sub_matches = self._find_matching_folders(subfolder, patterns, full_path, recursive)
                            if sub_matches:
                                logger.debug(f"Found {len(sub_matches)} matching subfolders in '{subfolder_name}'")
                                matching_folders.extend(sub_matches)
                        except Exception as e:
                            subfolder_name = getattr(subfolder, 'Name', None) or getattr(subfolder, 'name', f'Unknown_{i}')
                            logger.error(f"Error processing subfolder '{subfolder_name}': {e}", exc_info=True)
                            continue
                        
        except Exception as e:
            folder_name = getattr(folder, 'Name', None) or getattr(folder, 'name', 'Unknown')
            logger.error(f"Error in _find_matching_folders for folder '{folder_name}': {e}", exc_info=True)
            
        return matching_folders
        
    def is_mail_folder(self, folder) -> bool:
        """Check if a folder is a mail folder that can contain messages.
        
        Args:
            folder: The folder to check
            
        Returns:
            bool: True if the folder is a mail folder, False otherwise
        """
        try:
            # Check if the folder has the DefaultItemType property (Outlook 2007+)
            if hasattr(folder, 'DefaultItemType'):
                # 0 = MailItem, 1 = AppointmentItem, 2 = ContactItem, etc.
                return folder.DefaultItemType == 0
                
            # Fallback: Check if the folder has the Items collection
            if hasattr(folder, 'Items') and hasattr(folder, 'Name'):
                # Try to access the Items collection to verify it's a mail folder
                try:
                    # This will raise an error if it's not a mail folder
                    _ = folder.Items.Count
                    return True
                except:
                    return False
                    
            return False
            
        except Exception as e:
            logger.debug(f"Error checking if folder is a mail folder: {e}")
            return False
            
    def _process_email_data(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Process and validate email data before saving.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            Processed email data with additional metadata
        """
        if not email_data or not isinstance(email_data, dict):
            logger.warning("Invalid email data: not a dictionary")
            return {}
            
        try:
            # Ensure required fields exist
            email_data.setdefault('entry_id', str(uuid.uuid4()))
            email_data.setdefault('subject', '(No Subject)')
            email_data.setdefault('sender', 'Unknown Sender')
            email_data.setdefault('sender_email', '')
            email_data.setdefault('recipients', [])
            email_data.setdefault('cc_recipients', [])
            email_data.setdefault('bcc_recipients', [])
            email_data.setdefault('received_date', None)
            email_data.setdefault('sent_date', None)
            email_data.setdefault('body', '')
            email_data.setdefault('html_body', '')
            email_data.setdefault('importance', 1)  # 0=Low, 1=Normal, 2=High
            email_data.setdefault('is_read', False)
            email_data.setdefault('has_attachments', False)
            email_data.setdefault('categories', [])
            email_data.setdefault('in_reply_to', '')
            email_data.setdefault('conversation_id', '')
            email_data.setdefault('conversation_index', '')
            email_data.setdefault('internet_headers', {})
            email_data.setdefault('message_id', '')
            email_data.setdefault('size', 0)
            email_data.setdefault('sensitivity', 0)  # 0=Normal, 1=Personal, 2=Private, 3=Confidential
            
            # Normalize email addresses to lowercase
            email_data['sender_email'] = email_data['sender_email'].lower() if email_data['sender_email'] else ''
            
            # Process recipients
            for field in ['recipients', 'cc_recipients', 'bcc_recipients']:
                if isinstance(email_data.get(field), str):
                    # Convert string to list if needed
                    email_data[field] = [addr.strip() for addr in email_data[field].split(';') if addr.strip()]
                elif not isinstance(email_data.get(field), list):
                    email_data[field] = []
                
                # Normalize email addresses
                email_data[field] = [addr.lower() for addr in email_data[field] if addr]
            
            # Set priority flags based on sender
            email_data['is_priority'] = email_data['sender_email'] in self.priority_addresses
            email_data['is_admin'] = email_data['sender_email'] in self.admin_addresses
            
            # Generate a unique message ID if not present
            if not email_data.get('message_id'):
                email_data['message_id'] = f"<{email_data['entry_id']}@outlook.extractor>"
            
            # Set processing timestamp
            email_data['processed_at'] = datetime.now(timezone.utc).isoformat()
            
            return email_data
            
        except Exception as e:
            logger.error(f"Error processing email data: {e}", exc_info=True)
            # Return the original data with error flag
            email_data['_processing_error'] = str(e)
            return email_data
            
    def parse_date_ranges(self) -> Tuple[datetime, datetime]:
        """Parse date ranges from configuration.
        
        Returns:
            Tuple of (start_date, end_date) as timezone-aware datetime objects
            
        Raises:
            ValueError: If date range configuration is invalid
        """
        try:
            # Get date range configuration
            date_ranges = self.config.get('date_range', 'date_ranges', '').strip()
            days_back = int(self.config.get('date_range', 'days_back', '30'))
            
            end_date = datetime.now(timezone.utc)
            
            if date_ranges:
                # Parse date ranges in format "MM/YYYY,MM/YYYY"
                try:
                    start_str, end_str = [s.strip() for s in date_ranges.split(',', 1)]
                    
                    # Parse start date (beginning of month)
                    start_month, start_year = map(int, start_str.split('/'))
                    start_date = datetime(start_year, start_month, 1, tzinfo=timezone.utc)
                    
                    # Parse end date (end of month)
                    end_month, end_year = map(int, end_str.split('/'))
                    
                    # Calculate last day of the end month
                    if end_month == 12:
                        next_month = 1
                        next_year = end_year + 1
                    else:
                        next_month = end_month + 1
                        next_year = end_year
                        
                    end_date = datetime(next_year, next_month, 1, tzinfo=timezone.utc) - timedelta(days=1)
                    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
                    
                    logger.info(f"Using date range from config: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                    
                except (ValueError, IndexError) as e:
                    logger.warning(f"Invalid date range format in config: {date_ranges}. Using {days_back} days back as fallback.")
                    start_date = end_date - timedelta(days=days_back)
            else:
                # Use days_back as fallback
                start_date = end_date - timedelta(days=days_back)
                logger.info(f"Using default date range: {days_back} days back from {end_date.strftime('%Y-%m-%d')}")
            
            # Ensure timezone awareness
            if start_date.tzinfo is None:
                start_date = start_date.replace(tzinfo=timezone.utc)
            if end_date.tzinfo is None:
                end_date = end_date.replace(tzinfo=timezone.utc)
            
            return start_date, end_date
            
        except Exception as e:
            logger.error(f"Error parsing date ranges: {e}", exc_info=True)
            # Fallback to default 30 days if there's an error
            end_date = datetime.now(timezone.utc)
            start_date = end_date - timedelta(days=30)
            return start_date, end_date

    def _load_config(self) -> None:
        """Load configuration settings."""
        # Load priority and admin emails from config
        self.priority_addresses = set(
            email.strip().lower() 
            for email in self.config.get('email_processing.priority_emails', '').split(',')
            if email.strip()
        )
        
        self.admin_addresses = set(
            email.strip().lower()
            for email in self.config.get('email_processing.admin_emails', '').split(',')
            if email.strip()
        )
    
    def _extract_email_headers(self, msg) -> Dict[str, Any]:
        """Extract email headers for threading.
        
        Args:
            msg: Outlook message object
            
        Returns:
            Dictionary containing extracted email headers and metadata
        """
        try:
            # Get basic email properties
            entry_id = getattr(msg, 'EntryID', str(uuid.uuid4()))
            subject = getattr(msg, 'Subject', '(No Subject)')
            sender = getattr(msg, 'SenderName', 'Unknown Sender')
            sender_email = getattr(msg, 'SenderEmailAddress', '')
                
            # Get threading headers
            in_reply_to = getattr(msg, 'InReplyTo', '')
            references = getattr(msg, 'ConversationID', '') or getattr(msg, 'ConversationTopic', '')
            thread_index = getattr(msg, 'ConversationIndex', '')
                
            # Get recipients
            to_recipients = []
            cc_recipients = []
                
            if hasattr(msg, 'Recipients'):
                for recipient in msg.Recipients:
                    try:
                        email = ''
                        if hasattr(recipient, 'GetExchangeUser') and recipient.GetExchangeUser():
                            email = recipient.GetExchangeUser().PrimarySmtpAddress
                        elif hasattr(recipient, 'Address'):
                            email = recipient.Address
                            
                        if email:
                            if recipient.Type == 1:  # To
                                to_recipients.append(email)
                            elif recipient.Type == 2:  # CC
                                cc_recipients.append(email)
                    except Exception as e:
                        logger.debug(f"Error processing recipient: {e}")
                
            # Format recipients as strings
            to_str = '; '.join(to_recipients) if to_recipients else ''
            cc_str = '; '.join(cc_recipients) if cc_recipients else ''
                
            # Get dates
            received_time = getattr(msg, 'ReceivedTime', None)
            sent_time = getattr(msg, 'SentOn', None)
                
            # Ensure timezone awareness for dates
            if received_time and hasattr(received_time, 'tzinfo') and received_time.tzinfo is None:
                received_time = received_time.replace(tzinfo=timezone.utc)
            if sent_time and hasattr(sent_time, 'tzinfo') and sent_time.tzinfo is None:
                sent_time = sent_time.replace(tzinfo=timezone.utc)
                
            # Format dates as strings for storage
            received_str = received_time.isoformat() if received_time else ''
            sent_str = sent_time.isoformat() if sent_time else ''
                
            # Get message body
            body = getattr(msg, 'Body', '')
                
            # Create email data dictionary
            email_data = {
                'entry_id': entry_id,
                'folder': '',  # Will be set by the caller
                'subject': subject,
                'sender': sender,
                'sender_email': sender_email.lower(),
                'to_recipients': to_str.lower(),
                'cc_recipients': cc_str.lower(),
                'received_time': received_str,
                'sent_on': sent_str,
                'body': body,
                'body_preview': body[:500] + '...' if len(body) > 500 else body,
                'in_reply_to': in_reply_to,
                'references': references,
                'thread_index': thread_index,
                'is_read': bool(getattr(msg, 'UnRead', 0) == 0),  # 0 means read, 1 means unread
                'has_attachments': bool(getattr(msg, 'Attachments', None) and msg.Attachments.Count > 0),
                'categories': getattr(msg, 'Categories', '')
            }
                
            return email_data
                
        except Exception as e:
            logger.error(f"Error extracting email headers: {e}", exc_info=True)
            raise

    def _get_mapi_property(self, prop_accessor, prop_name: str, default: Any = None) -> Any:
        """Safely get a MAPI property.
        
        Args:
            prop_accessor: Outlook property accessor object
            prop_name: Name of the property to get
            default: Default value if property cannot be retrieved
            
        Returns:
            The property value or default if not found
        """
        try:
            if hasattr(prop_accessor, 'PropertyAccessor'):
                return prop_accessor.PropertyAccessor.GetProperty(prop_name)
            return getattr(prop_accessor, prop_name, default)
        except Exception as e:
            logger.debug(f"Could not get property {prop_name}: {e}")
            return default

    def _process_email_data(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """Process raw email data before saving.
        
        Args:
            email_data: Raw email data from Outlook
            
        Returns:
            Processed email data
        """
        # Ensure required fields
        if 'entry_id' not in email_data:
            email_data['entry_id'] = hashlib.md5(
                f"{email_data.get('subject', '')}{email_data.get('sent_on', '')}".encode('utf-8')
            ).hexdigest()
        
        # Normalize email addresses
        for field in ['sender_email', 'to_recipients', 'cc_recipients']:
            if field in email_data and email_data[field]:
                email_data[field] = self._normalize_email_field(email_data[field])
        
        # Add priority and admin flags
        email_data['is_priority'] = self._is_priority_email(email_data)
        email_data['is_admin'] = self._is_admin_email(email_data)
        
        # Ensure thread-related fields exist
        email_data.setdefault('in_reply_to', '')
        email_data.setdefault('references', '')
        email_data.setdefault('thread_index', '')
        
        return email_data
    
    def _normalize_email_field(self, email_field: str) -> str:
        """Normalize an email field to a semicolon-separated string of addresses."""
        if not email_field:
            return ""
        
        # Handle both string (semicolon-separated) and list of addresses
        if isinstance(email_field, str):
            # Clean up and normalize the email string
            emails = [e.strip().lower() for e in email_field.split(';') if e.strip()]
        elif isinstance(email_field, list):
            emails = [e.strip().lower() for e in email_field if isinstance(e, str) and e.strip()]
        else:
            return ""
        
        # Remove duplicates while preserving order
        seen = set()
        return "; ".join([e for e in emails if not (e in seen or seen.add(e))])
        
    def parse_date_ranges(self) -> tuple[datetime, datetime]:
        """Parse date ranges from config.
        
        Returns:
            Tuple of (start_date, end_date) as timezone-aware datetime objects
        """
        try:
            # Check if date_ranges is specified in config
            date_ranges = self.config.get('date_range', 'ranges', fallback='').strip()
            if date_ranges:
                ranges = [r.strip() for r in date_ranges.split(',') if r.strip()]
                if len(ranges) == 2:
                    try:
                        # Parse MM/YYYY format
                        start_month, start_year = map(int, ranges[0].split('/'))
                        end_month, end_year = map(int, ranges[1].split('/'))
                        
                        # Create datetime objects (first day of start month, last day of end month)
                        start_date = datetime(start_year, start_month, 1, tzinfo=timezone.utc)
                        
                        # Calculate first day of next month, then subtract one day
                        if end_month == 12:
                            next_month = 1
                            next_year = end_year + 1
                        else:
                            next_month = end_month + 1
                            next_year = end_year
                            
                        end_date = datetime(next_year, next_month, 1, tzinfo=timezone.utc) - timedelta(days=1)
                        end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
                        
                        logger.info(f"Using custom date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                        return start_date, end_date
                        
                    except (ValueError, IndexError) as e:
                        logger.warning(f"Invalid date_ranges format. Using days_back instead. Error: {e}")
            
            # Fall back to days_back if no valid date_ranges
            days_back = int(self.config.get('date_range', 'days_back', fallback='14'))
            end_date = datetime.now(timezone.utc)
            start_date = end_date - timedelta(days=days_back)
            logger.info(f"Using last {days_back} days: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            return start_date, end_date
            
        except Exception as e:
            logger.error(f"Error parsing date range: {e}", exc_info=True)
            # Default to last 14 days if there's any error
            end_date = datetime.now(timezone.utc)
            start_date = end_date - timedelta(days=14)
            return start_date, end_date
    
    def _is_priority_email(self, email_data: Dict[str, Any]) -> bool:
        """Check if an email is from a priority sender."""
        if not self.priority_addresses:
            return False
            
        sender = email_data.get('sender_email', '').lower()
        if not sender:
            return False
            
        return any(priority in sender for priority in self.priority_addresses)
    
    def _is_admin_email(self, email_data: Dict[str, Any]) -> bool:
        """Check if an email is from an admin."""
        if not self.admin_addresses:
            return False
            
        sender = email_data.get('sender_email', '').lower()
        if not sender:
            return False
            
        return any(admin in sender for admin in self.admin_addresses)
    
    def extract_emails(
        self,
        folder_patterns: List[str],
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
        **kwargs
    ) -> Dict[str, Any]:
        """Extract emails from Outlook with threading support.
        
        Args:
            folder_patterns: List of folder patterns to search in (supports wildcards)
            start_date: Optional start date for filtering emails (overrides config)
            end_date: Optional end date for filtering emails (overrides config)
            **kwargs: Additional parameters
                - include_threads: bool - Whether to include thread information
                - thread_status: str - Status to set for new threads
                - recursive: bool - Whether to search subfolders recursively (default: True)
                - max_emails: int - Maximum number of emails to process (0 for no limit)
                
        Returns:
            Dictionary containing extraction results with thread information
        """
        # Initialize counters
        emails_processed = 0
        emails_saved = 0
        threads_processed = 0
        folders = []
        
        # Get options from kwargs
        include_threads = kwargs.get('include_threads', True)
        thread_status = kwargs.get('thread_status', 'active')
        recursive = kwargs.get('recursive', True)
        max_emails = int(kwargs.get('max_emails', 0))  # 0 means no limit
        
        try:
            # Get date range from config if not provided
            if start_date is None or end_date is None:
                config_start, config_end = self.parse_date_ranges()
                start_date = start_date or config_start
                end_date = end_date or config_end
            
            # Ensure dates are timezone-aware
            if start_date and start_date.tzinfo is None:
                start_date = start_date.replace(tzinfo=timezone.utc)
            if end_date and end_date.tzinfo is None:
                end_date = end_date.replace(tzinfo=timezone.utc)
            
            logger.info(f"Extracting emails from {start_date} to {end_date}")
            
            # Get the namespace and root folder
            namespace = self.outlook_client.GetNamespace("MAPI")
            root_folder = namespace.Folders
            
            # Normalize folder patterns (trim whitespace and handle case)
            folder_patterns = [p.strip() for p in folder_patterns if p and p.strip()]
            
            # If no valid patterns, use default
            if not folder_patterns:
                folder_patterns = ['Inbox']
                logger.warning("No valid folder patterns provided, defaulting to 'Inbox'")
            
            # Debug: Log the root folder structure in detail
            logger.debug(f"Root folder type: {type(root_folder)}")
            logger.debug(f"Root folder dir: {[a for a in dir(root_folder) if not a.startswith('__')]}")
            
            # Try to get the root folder contents
            try:
                if hasattr(root_folder, '__iter__') and not isinstance(root_folder, (str, bytes, dict)):
                    root_items = list(root_folder)
                    logger.debug(f"Root folder is iterable with {len(root_items)} items")
                    for i, item in enumerate(root_items, 1):
                        item_type = type(item).__name__
                        item_attrs = [a for a in dir(item) if not a.startswith('__')]
                        logger.debug(f"  Item {i}: Type={item_type}, Attrs={item_attrs}")
                        if hasattr(item, 'Name'):
                            logger.debug(f"    Name: {getattr(item, 'Name', 'N/A')}")
                        if hasattr(item, 'name'):
                            logger.debug(f"    name: {getattr(item, 'name', 'N/A')}")
                        if hasattr(item, 'Folders'):
                            folders = getattr(item, 'Folders', [])
                            if hasattr(folders, '__iter__'):
                                logger.debug(f"    Has {len(list(folders))} subfolders in 'Folders'")
            except Exception as e:
                logger.error(f"Error inspecting root folder contents: {e}", exc_info=True)
            
            # Try to iterate through root_folder items
            try:
                root_items = list(root_folder)
                logger.debug(f"Root folder contains {len(root_items)} items")
                for i, account in enumerate(root_items, 1):
                    account_name = getattr(account, 'Name', None) or getattr(account, 'name', f'Account_{i}')
                    logger.debug(f"Processing account {i}: {account_name}")
                    logger.debug(f"Account type: {type(account)}")
                    logger.debug(f"Account dir: {[a for a in dir(account) if not a.startswith('__')]}")
                    
                    # Check if account has Folders attribute
                    if hasattr(account, 'Folders'):
                        folders_attr = getattr(account, 'Folders', [])
                        logger.debug(f"Account {account_name} has {len(list(folders_attr)) if hasattr(folders_attr, '__iter__') else 'unknown'} folders in 'Folders'")
                    
                    # Try to find matching folders
                    try:
                        account_folders = self._find_matching_folders(
                            account, 
                            folder_patterns,
                            recursive=recursive
                        )
                        logger.debug(f"Found {len(account_folders)} matching folders in account {account_name}")
                        folders.extend(account_folders)
                    except Exception as e:
                        logger.error(f"Error finding folders in account {account_name}: {e}", exc_info=True)
                        
            except Exception as e:
                logger.error(f"Error iterating through root folder items: {e}", exc_info=True)
            
            if not folders:
                error_msg = f"No folders found matching patterns: {', '.join(folder_patterns)}"
                logger.warning(error_msg)
                return {
                    'success': False,
                    'error': error_msg,
                    'emails_processed': 0,
                    'emails_saved': 0,
                    'folders_processed': 0
                }
            
            logger.info(f"Found {len(folders)} folders to process")
            
            # Process each folder
            for folder, folder_path in folders:
                try:
                    if not self.is_mail_folder(folder):
                        logger.debug(f"Skipping non-mail folder: {folder_path}")
                        continue
                        
                    logger.info(f"Processing folder: {folder_path}")
                    
                    # Get all emails in the folder
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
                    
                    # Apply date filter
                    filter_str = []
                    if start_date:
                        filter_str.append(f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %H:%M %p')}'")
                    if end_date:
                        filter_str.append(f"[ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %H:%M %p')}'")
                    
                    if filter_str:
                        items = items.Restrict(' AND '.join(filter_str))
                    
                    # Get emails with progress tracking
                    total_emails = items.Count
                    logger.info(f"Found {total_emails} emails in folder {folder_path}")
                    
                    # Process emails in batches to manage memory
                    processed_emails = []
                    batch_size = 100
                    
                    for i in range(0, total_emails, batch_size):
                        # Check if we've reached the maximum number of emails to process
                        if max_emails > 0 and emails_processed >= max_emails:
                            logger.info(f"Reached maximum of {max_emails} emails to process")
                            break
                            
                        # Process a batch of emails
                        batch = []
                        for j in range(i, min(i + batch_size, total_emails)):
                            try:
                                msg = items[j + 1]  # Outlook collections are 1-based
                                batch.append(msg)
                            except Exception as e:
                                logger.error(f"Error getting email {j+1}/{total_emails}: {e}")
                        
                        # Process the batch
                        for msg in batch:
                            try:
                                # Check if we've reached the maximum number of emails to process
                                if max_emails > 0 and emails_processed >= max_emails:
                                    break
                                
                                # Extract email headers and metadata
                                email_data = self._extract_email_headers(msg)
                                if not email_data:
                                    continue
                                
                                # Set the folder path
                                email_data['folder'] = folder_path
                                
                                # Process and validate email data
                                email_data = self._process_email_data(email_data)
                                
                                # Add to thread manager if threading is enabled
                                if include_threads:
                                    self.thread_manager.add_email(email_data)
                                
                                processed_emails.append(email_data)
                                emails_processed += 1
                                
                                # Log progress periodically
                                if emails_processed % 10 == 0 or emails_processed == 1:
                                    logger.info(f"Processed {emails_processed} emails ({len(processed_emails)} in current batch)")
                                
                            except Exception as e:
                                logger.error(f"Error processing email: {e}", exc_info=True)
                                continue
                        
                        # Save the batch to storage
                        if processed_emails:
                            try:
                                saved_count = self.storage.save_emails(processed_emails)
                                emails_saved += saved_count
                                logger.info(f"Saved {saved_count} emails to storage")
                                
                                # Log memory usage
                                try:
                                    import psutil
                                    process = psutil.Process()
                                    memory_info = process.memory_info()
                                    logger.debug(
                                        f"Memory usage: {memory_info.rss / 1024 / 1024:.2f}MB "
                                        f"(virtual: {memory_info.vms / 1024 / 1024:.2f}MB)"
                                    )
                                except ImportError:
                                    pass  # psutil not available, skip memory logging
                                
                                # Clear the processed batch
                                processed_emails = []
                                
                            except Exception as e:
                                logger.error(f"Error saving emails to storage: {e}", exc_info=True)
                                # Continue processing other folders even if one fails
                    
                except Exception as e:
                    logger.error(f"Error processing folder {folder_path}: {e}", exc_info=True)
                    continue
            
            # Process threads if enabled
            threads = []
            if include_threads and self.thread_manager:
                try:
                    threads = self.thread_manager.get_threads()
                    threads_processed = len(threads)
                    logger.info(f"Processed {threads_processed} email threads")
                except Exception as e:
                    logger.error(f"Error processing threads: {e}", exc_info=True)
                    # Continue with empty threads list if there's an error
            
            # Prepare results
            result = {
                'success': True,
                'emails_processed': emails_processed,
                'emails_saved': emails_saved,
                'threads_processed': threads_processed,
                'folders_processed': len(folders)
            }
            
            # Only include threads in result if threading is enabled
            if include_threads and threads:
                # Handle case where threads might be dictionaries or objects
                result['threads'] = [
                    thread if isinstance(thread, dict) else thread.to_dict()
                    for thread in threads
                ]
                
            return result
            
        except Exception as e:
            error_msg = f"Failed to extract emails: {str(e)}"
            logger.error(error_msg, exc_info=True)
            return {
                'success': False,
                'error': error_msg,
                'emails_processed': emails_processed,
                'emails_saved': emails_saved,
                'folders_processed': len(folders)
            }
    
    def export_emails(
        self,
        emails: List[Dict[str, Any]],
        format: str = 'csv',
        export_settings: Optional[Dict[str, Any]] = None
    ) -> tuple:
        """Export emails to the specified format.
        
        Args:
            emails: List of email dictionaries to export
            format: Export format ('csv' or 'json')
            export_settings: Dictionary of export settings
            
        Returns:
            Tuple of (success, output_files)
        """
        if not emails:
            logger.warning("No emails to export")
            return False, []
        
        export_settings = export_settings or {}
        
        try:
            if format.lower() == 'csv':
                output_dir = export_settings.get('output_dir', str(Path.home() / 'email_exports'))
                file_prefix = export_settings.get('file_prefix', 'emails_')
                
                # Ensure output directory exists
                os.makedirs(output_dir, exist_ok=True)
                
                # Generate output filename with timestamp
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_file = os.path.join(output_dir, f"{file_prefix}{timestamp}.csv")
                
                # Export to CSV
                success = self.csv_exporter.export_emails(
                    emails=emails,
                    output_file=output_file,
                    include_headers=export_settings.get('include_headers', True),
                    encoding=export_settings.get('encoding', 'utf-8')
                )
                
                return success, [output_file] if success else []
                
            else:
                logger.error(f"Unsupported export format: {format}")
                return False, []
                
        except Exception as e:
            logger.error(f"Failed to export emails: {e}", exc_info=True)
            return False, []
    
    def close(self) -> bool:
        """Close the Outlook client and storage.
        
        Returns:
            bool: True if resources were closed successfully, False otherwise
        """
        try:
            if hasattr(self, '_outlook_client') and self._outlook_client:
                self._outlook_client.quit()
            if hasattr(self, 'storage') and self.storage:
                self.storage.close()
            logger.info("OutlookExtractor resources cleaned up")
            return True
        except Exception as e:
            logger.error(f"Error cleaning up resources: {e}")
            return False
