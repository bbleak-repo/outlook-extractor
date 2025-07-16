"""
Outlook Email Extraction Tool v10.0.0

This script extracts emails from Microsoft Outlook and saves them to a SQLite database and/or JSON file.
It includes folder type checking, wildcard folder matching, and configurable email processing.

Version History:
- v8.0.1: Added support for date ranges in config file (MM/YYYY format)
- v8.0.0: Added wildcard folder support, basic extraction mode, and fixed JSON duplicates
"""

import os
import sys
import re
import fnmatch
import json
import sqlite3
import logging
import win32com.client
import pythoncom
import configparser
import uuid
import hashlib
from collections import defaultdict, deque
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Set, Optional, Any, Tuple, DefaultDict, Deque
import traceback
from pathlib import Path
from dataclasses import dataclass, field
from email.utils import parseaddr, getaddresses

# Constants
VERSION = "10.0.0"
LOG_FILE = 'outlook_extraction.log'

# Threading constants
THREAD_STATUS_ACTIVE = 'active'
THREAD_STATUS_COMPLETED = 'completed'
THREAD_STATUS_FAILED = 'failed'
THREAD_STATUS_ARCHIVED = 'archived'

@dataclass
class EmailThread:
    """Represents a conversation thread of emails."""
    thread_id: str
    subject: str
    participants: Set[str] = field(default_factory=set)
    message_ids: Set[str] = field(default_factory=set)
    root_message_id: Optional[str] = None
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None
    status: str = THREAD_STATUS_ACTIVE
    categories: Set[str] = field(default_factory=set)

    def add_email(self, email_data: Dict[str, Any]):
        """Add an email to this thread."""
        self.message_ids.add(email_data['entry_id'])
        self._extract_participants(email_data)
        
        # Update thread dates
        email_date = self._parse_date(email_data.get('received_time'))
        if email_date:
            if self.start_date is None or email_date < self.start_date:
                self.start_date = email_date
            if self.end_date is None or email_date > self.end_date:
                self.end_date = email_date
    
    def _extract_participants(self, email_data: Dict[str, Any]):
        """Extract all participants from an email."""
        for field in ['sender_email', 'to_recipients', 'cc_recipients']:
            if field in email_data and email_data[field]:
                if field == 'to_recipients' or field == 'cc_recipients':
                    for addr in email_data[field].split(';'):
                        if addr.strip():
                            self.participants.add(addr.strip().lower())
                else:
                    self.participants.add(email_data[field].lower())
    
    @staticmethod
    def _parse_date(date_str: Optional[str]) -> Optional[datetime]:
        """Parse a date string into a datetime object."""
        if not date_str:
            return None
        try:
            if isinstance(date_str, str):
                return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
            elif isinstance(date_str, datetime):
                return date_str
        except (ValueError, AttributeError):
            pass
        return None


class ThreadManager:
    """Manages email threads and their relationships."""
    
    def __init__(self):
        self.threads_by_id: Dict[str, EmailThread] = {}
        self.message_to_thread: Dict[str, str] = {}
        self.threads_by_participant: DefaultDict[str, Set[str]] = defaultdict(set)
    
    def add_email(self, email_data: Dict[str, Any]):
        """Add an email to the appropriate thread."""
        # Skip if already processed
        if email_data['entry_id'] in self.message_to_thread:
            return
        
        # Extract threading headers
        in_reply_to = email_data.get('in_reply_to', '').strip()
        references = [r.strip() for r in email_data.get('references', '').split() if r.strip()]
        
        # Find existing thread or create new one
        thread_id = self._find_existing_thread(in_reply_to, references)
        if not thread_id:
            thread_id = self._create_new_thread(email_data)
        
        # Add email to thread
        self.threads_by_id[thread_id].add_email(email_data)
        self.message_to_thread[email_data['entry_id']] = thread_id
        
        # Update participant index
        for participant in self.threads_by_id[thread_id].participants:
            self.threads_by_participant[participant].add(thread_id)
    
    def _find_existing_thread(self, in_reply_to: Optional[str], references: List[str]) -> Optional[str]:
        """Find an existing thread based on reply/reference headers."""
        # Check in-reply-to
        if in_reply_to and in_reply_to in self.message_to_thread:
            return self.message_to_thread[in_reply_to]
        
        # Check references in reverse order (most recent first)
        for ref in reversed(references):
            if ref in self.message_to_thread:
                return self.message_to_thread[ref]
        
        return None
    
    def _create_new_thread(self, email_data: Dict[str, Any]) -> str:
        """Create a new thread for an email."""
        thread_id = self._generate_thread_id(email_data)
        subject = email_data.get('subject', '(no subject)')
        thread = EmailThread(thread_id=thread_id, subject=subject)
        
        # Set root message if this is a reply
        if email_data.get('in_reply_to'):
            thread.root_message_id = email_data['in_reply_to']
        
        self.threads_by_id[thread_id] = thread
        return thread_id
    
    def _generate_thread_id(self, email_data: Dict[str, Any]) -> str:
        """Generate a unique thread ID based on email headers."""
        # Try to use thread index if available
        if 'thread_index' in email_data and email_data['thread_index']:
            return hashlib.md5(email_data['thread_index'].encode('utf-8')).hexdigest()
        
        # Fall back to subject and participants
        subject = email_data.get('subject', '')
        participants = set()
        
        def add_participants(header):
            if not header:
                return
            if isinstance(header, str):
                for addr in header.split(';'):
                    if '@' in addr:
                        participants.add(addr.strip().lower())
            elif isinstance(header, (list, tuple)):
                for addr in header:
                    if isinstance(addr, str) and '@' in addr:
                        participants.add(addr.strip().lower())
        
        add_participants(email_data.get('sender_email'))
        add_participants(email_data.get('to_recipients'))
        add_participants(email_data.get('cc_recipients'))
        
        # Generate a stable ID based on subject and participants
        id_str = f"{subject}:{':'.join(sorted(participants))}"
        return hashlib.md5(id_str.encode('utf-8')).hexdigest()
    
    def _parse_references(self, references: str) -> List[str]:
        """Parse the References header into individual message IDs."""
        if not references:
            return []
        return [r.strip() for r in references.split() if r.strip()]
    
    def get_threads(self) -> List[Dict]:
        """Get all threads as a list of dictionaries."""
        return [
            {
                'thread_id': thread.thread_id,
                'subject': thread.subject,
                'message_count': len(thread.message_ids),
                'participants': sorted(thread.participants),
                'start_date': thread.start_date.isoformat() if thread.start_date else None,
                'end_date': thread.end_date.isoformat() if thread.end_date else None,
                'status': thread.status,
                'categories': sorted(thread.categories)
            }
            for thread in self.threads_by_id.values()
        ]

class OutlookExtractor:
    def __init__(self):
        self.config = self.load_config()
        self.setup_logging()
        self.logger = logging.getLogger(__name__)
        self.logger.info(f"Outlook Email Extraction Tool v{VERSION}")
        self.logger.info(f"Python {sys.version}")
        self.logger.info(f"Current directory: {os.getcwd()}")
        self.logger.info(f"Logging to {os.path.abspath(LOG_FILE)}")

    def setup_logging(self):
        """Configure logging to both file and console."""
        log_level = self.config.get('DEFAULT', 'log_level', fallback='INFO').upper()
        log_level = getattr(logging, log_level, logging.INFO)
        
        # Clear any existing handlers
        logging.getLogger().handlers = []
        
        # Configure root logger
        logging.basicConfig(
            level=getattr(logging, self.config['DEFAULT']['log_level'].upper()),
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(LOG_FILE, mode='w'),  # Overwrite log file
                logging.StreamHandler()
            ]
        )

    def load_config(self) -> configparser.ConfigParser:
        """Load configuration from INI file or create default if not exists."""
        config = configparser.ConfigParser()
        config_file = 'outlook_extraction_config.ini'
        
        # Default configuration
        if not os.path.exists(config_file):
            config['DEFAULT'] = {
                'days_back': '14',
                'log_level': 'INFO',
            }
            
            config['outlook'] = {
                'folders': 'Inbox, Sent Items, Mayor and Council, Legal*',
                'basic_extraction_only': '0'
            }
            
            config['storage'] = {
                'backend': 'sqlite,json',
                'sqlite_path': 'outlook_emails_dashboard.sqlite',
                'json_path': 'outlook_emails_dashboard.json',
            }
            
            config['priority'] = {
                'addresses': 'mayor@example.com,admin@example.com',
            }
            
            config['administrators'] = {
                'addresses': 'khumphrey@keyportonline.com,newadmin@keyportonline.com,iani@keyportonline.com'
            }
            
            # Save default config
            with open(config_file, 'w') as f:
                config.write(f)
        
        # Load config
        config.read(config_file)
        return config

    def folder_matches_pattern(self, folder_name: str, patterns: List[str]) -> bool:
        """Check if folder name matches any of the patterns (supports wildcards)."""
        if not patterns:
            return False
            
        folder_name_lower = folder_name.lower()
        self.logger.debug(f"Checking if folder '{folder_name}' matches patterns: {patterns}")
        
        for pattern in patterns:
            pattern_lower = pattern.lower()
            # Convert Windows path separators to forward slashes for consistent matching
            normalized_pattern = pattern_lower.replace('\\', '/')
            normalized_name = folder_name_lower.replace('\\', '/')
            
            # Check for direct match or wildcard match
            if (normalized_name == pattern_lower or 
                fnmatch.fnmatch(normalized_name, pattern_lower) or
                fnmatch.fnmatch(normalized_name, f"*{pattern_lower}")):
                self.logger.debug(f"Folder '{folder_name}' matches pattern: {pattern}")
                return True
                
        self.logger.debug(f"Folder '{folder_name}' did not match any patterns")
        return False

    def find_matching_folders(self, outlook, folder_patterns):
        """Find all folders matching the given patterns."""
        matching_folders = []
        patterns = [p.strip().lower() for p in folder_patterns.split(',') if p.strip()]
        self.logger.info(f"Searching for folders matching patterns: {patterns}")
        
        # First, try to find exact matches in the main account
        account = None
        for folder in outlook.Folders:
            if folder.Name.lower() == 'khumphrey@keyportonline.com':
                account = folder
                break
        
        if not account:
            self.logger.warning("Could not find the main account folder. Using first available account.")
            account = outlook.Folders[0]
        
        self.logger.info(f"Using account: {account.Name}")
        
        # Get all folders in the account
        all_folders = []
        
        def collect_folders(folder, path=''):
            current_path = f"{path}/{folder.Name}" if path else folder.Name
            all_folders.append((folder, current_path))
            for subfolder in folder.Folders:
                collect_folders(subfolder, current_path)
        
        collect_folders(account)
        
        # Log all available folders for debugging
        self.logger.info("Available folders:")
        for folder, path in all_folders:
            self.logger.info(f"- {path} (Type: {getattr(folder, 'DefaultItemType', 'Unknown')})")
        
        # Try to match folders
        for pattern in patterns:
            self.logger.info(f"Searching for pattern: {pattern}")
            
            # Try different matching strategies
            for folder, path in all_folders:
                folder_name = folder.Name.lower()
                
                # 1. Exact name match
                if folder_name == pattern.lower():
                    if folder not in matching_folders:
                        self.logger.info(f"Found exact match: {path}")
                        matching_folders.append(folder)
                
                # 2. Case-insensitive contains
                elif pattern.lower() in folder_name:
                    if folder not in matching_folders:
                        self.logger.info(f"Found partial match: {path} (contains '{pattern}')")
                        matching_folders.append(folder)
                
                # 3. Wildcard match
                elif fnmatch.fnmatch(folder_name, f"*{pattern.lower()}*"):
                    if folder not in matching_folders:
                        self.logger.info(f"Found wildcard match: {path} (matches '{pattern}')")
                        matching_folders.append(folder)
        
        self.logger.info(f"Found {len(matching_folders)} matching folders")
        for i, folder in enumerate(matching_folders, 1):
            self.logger.info(f"  {i}. {folder.Name}")
            
        if not matching_folders:
            self.logger.warning("No matching folders found. Available folders:")
            for folder, path in all_folders[:20]:  # Only show first 20 to avoid flooding
                self.logger.warning(f"- {path}")
            if len(all_folders) > 20:
                self.logger.warning(f"... and {len(all_folders) - 20} more folders")
                
        return matching_folders

    def get_folder_path(self, folder):
        """Get the full path of a folder by traversing up the folder hierarchy."""
        try:
            path_parts = [folder.Name]
            parent = folder.Parent
            
            # Traverse up the folder hierarchy
            while parent and hasattr(parent, 'Name') and parent.Name != '':
                path_parts.insert(0, parent.Name)
                parent = getattr(parent, 'Parent', None)
            
            # Remove the account name from the path
            if len(path_parts) > 1 and path_parts[0].lower() == 'khumphrey@keyportonline.com':
                path_parts = path_parts[1:]
            
            return '\\'.join(path_parts)
            
        except Exception as e:
            self.logger.warning(f"Error getting folder path: {e}")
            return getattr(folder, 'Name', 'Unknown')
    
    def get_folders_to_process(self, outlook, folder_patterns: List[str]) -> List[Any]:
        """Get all folders that match the specified patterns."""
        if not folder_patterns:
            self.logger.warning("No folder patterns specified in config")
            return []
            
        self.logger.info(f"Looking for folders matching: {', '.join(folder_patterns)}")
        
        def search_folders(folder, current_path=""):
            folders = []
            try:
                folder_name = getattr(folder, 'Name', 'Unknown')
                full_path = f"{current_path}\\{folder_name}" if current_path else folder_name
                
                # Normalize path for matching (replace backslashes with forward slashes)
                normalized_path = full_path.replace('\\', '/')
                
                # Check if current folder matches any pattern
                if self.folder_matches_pattern(normalized_path, folder_patterns):
                    self.logger.info(f"Found matching folder: {full_path}")
                    folders.append((folder, full_path))
                
                # Recursively search subfolders
                if hasattr(folder, 'Folders'):
                    for subfolder in folder.Folders:
                        try:
                            folders.extend(search_folders(subfolder, full_path))
                        except Exception as e:
                            self.logger.warning(f"Could not process subfolder in {full_path}: {e}")
                
            except Exception as e:
                self.logger.error(f"Error processing folder {getattr(folder, 'Name', 'Unknown')}: {e}")
            
            return folders
        
        # Start search from the root of each account
        matching_folders = []
        for account in outlook.Folders:
            try:
                matching_folders.extend(search_folders(account, ""))
            except Exception as e:
                self.logger.error(f"Error processing account {account.Name}: {e}")
        
        return matching_folders

    def is_mail_folder(self, folder) -> bool:
        """Check if a folder is a mail folder that can be processed."""
        try:
            folder_name = getattr(folder, 'Name', 'Unknown')
            
            # Check for required attributes
            if not hasattr(folder, 'DefaultItemType'):
                self.logger.debug(f"Folder '{folder_name}': Missing DefaultItemType attribute")
                return False
                
            # Check if it's a mail folder (5 = olMailItem)
            if folder.DefaultItemType != 5:
                self.logger.debug(f"Folder '{folder_name}': Not a mail folder (DefaultItemType={folder.DefaultItemType}, expected 5)")
                return False
                
            # Check if it has Items collection
            if not hasattr(folder, 'Items'):
                self.logger.debug(f"Folder '{folder_name}': Missing Items attribute")
                return False
                
            # Additional check for folder class
            if hasattr(folder, 'Class') and folder.Class != 18:  # 18 = olFolder
                self.logger.debug(f"Folder '{folder_name}': Incorrect folder class ({folder.Class}, expected 18)")
                return False
                
            self.logger.debug(f"Folder '{folder_name}': Valid mail folder")
            return True
            
        except Exception as e:
            folder_name = getattr(folder, 'Name', 'Unknown')
            self.logger.warning(f"Error checking folder type for '{folder_name}': {e}", exc_info=True)
            return False

    def get_folder_path(self, folder):
        """Get the full path of a folder by traversing up the folder hierarchy."""
        try:
            path_parts = [folder.Name]
            parent = folder.Parent
                
            # Traverse up the folder hierarchy
            while parent and hasattr(parent, 'Name') and parent.Name != '':
                path_parts.insert(0, parent.Name)
                parent = getattr(parent, 'Parent', None)
            
            # Remove the account name from the path
            if len(path_parts) > 1 and path_parts[0].lower() == 'khumphrey@keyportonline.com':
                path_parts = path_parts[1:]
            
            return '\\'.join(path_parts)
                
        except Exception as e:
            self.logger.warning(f"Error getting folder path: {e}")
            return getattr(folder, 'Name', 'Unknown')

    def get_folders_to_process(self, outlook, folder_patterns: List[str]) -> List[Any]:
        """Get all folders that match the specified patterns."""
        if not folder_patterns:
            self.logger.warning("No folder patterns specified in config")
            return []
                
        self.logger.info(f"Looking for folders matching: {', '.join(folder_patterns)}")
            
        def search_folders(folder, current_path=""):
            folders = []
            try:
                folder_name = getattr(folder, 'Name', 'Unknown')
                full_path = f"{current_path}\\{folder_name}" if current_path else folder_name
                    
                # Normalize path for matching (replace backslashes with forward slashes)
                normalized_path = full_path.replace('\\', '/')
                    
                # Check if current folder matches any pattern
                if self.folder_matches_pattern(normalized_path, folder_patterns):
                    self.logger.info(f"Found matching folder: {full_path}")
                    folders.append((folder, full_path))
                
                # Recursively search subfolders
                if hasattr(folder, 'Folders'):
                    for subfolder in folder.Folders:
                        try:
                            folders.extend(search_folders(subfolder, full_path))
                        except Exception as e:
                            self.logger.warning(f"Could not process subfolder in {full_path}: {e}")
                    
            except Exception as e:
                self.logger.error(f"Error processing folder {getattr(folder, 'Name', 'Unknown')}: {e}")
                
            return folders
            
        # Start search from the root of each account
        matching_folders = []
        for account in outlook.Folders:
            try:
                matching_folders.extend(search_folders(account, ""))
            except Exception as e:
                self.logger.error(f"Error processing account {account.Name}: {e}")
            
        return matching_folders

    def is_mail_folder(self, folder) -> bool:
        """Check if a folder is a mail folder that can be processed."""
        try:
            folder_name = getattr(folder, 'Name', 'Unknown')
                
            # Check for required attributes
            if not hasattr(folder, 'DefaultItemType'):
                self.logger.debug(f"Folder '{folder_name}': Missing DefaultItemType attribute")
                return False
                    
            # Check if it's a mail folder (5 = olMailItem)
            if folder.DefaultItemType != 5:
                self.logger.debug(f"Folder '{folder_name}': Not a mail folder (DefaultItemType={folder.DefaultItemType}, expected 5)")
                return False
                    
            # Check if it has Items collection
            if not hasattr(folder, 'Items'):
                self.logger.debug(f"Folder '{folder_name}': Missing Items attribute")
                return False
                    
            # Additional check for folder class
            if hasattr(folder, 'Class') and folder.Class != 18:  # 18 = olFolder
                self.logger.debug(f"Folder '{folder_name}': Incorrect folder class ({folder.Class}, expected 18)")
                return False
                    
            self.logger.debug(f"Folder '{folder_name}': Valid mail folder")
            return True
                
        except Exception as e:
            folder_name = getattr(folder, 'Name', 'Unknown')
            self.logger.warning(f"Error checking folder type for '{folder_name}': {e}", exc_info=True)
            return False

    def _extract_email_headers(self, msg) -> Dict[str, Any]:
        """Extract email headers for threading."""
        try:
            # Get basic email properties
            entry_id = msg.EntryID if hasattr(msg, 'EntryID') else str(uuid.uuid4())
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
                        self.logger.debug(f"Error processing recipient: {e}")
                
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
            body = ''
            if hasattr(msg, 'Body'):
                body = msg.Body
                
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
            self.logger.error(f"Error extracting email headers: {e}", exc_info=True)
            raise

    def _get_mapi_property(self, prop_accessor, prop_name: str, default: Any = None) -> Any:
        """Safely get a MAPI property."""
        try:
            return prop_accessor.GetProperty(prop_name)
        except Exception as e:
            self.logger.debug(f"Could not get property {prop_name}: {e}")
            return default

    def process_email(self, msg, folder_path: str, priority_addresses: set, admin_addresses: set):
        """Process a single email message with threading support."""
        try:
            # Extract email data with threading information
            email_data = self._extract_email_headers(msg)
            if not email_data:
                return None
                
            # Set folder path
            email_data['folder'] = folder_path
            
            # Check if email is from priority or admin
            sender_email = email_data['sender_email'].lower()
            email_data['is_priority'] = sender_email in priority_addresses
            email_data['is_admin'] = sender_email in admin_addresses
            
            # Extract categories if available and convert to string
            if 'categories' in email_data and email_data['categories']:
                if isinstance(email_data['categories'], str):
                    categories = [c.strip() for c in email_data['categories'].split(';') if c.strip()]
                    email_data['categories'] = ';'.join(categories)
                elif isinstance(email_data['categories'], list):
                    email_data['categories'] = ';'.join([str(c).strip() for c in email_data['categories'] if str(c).strip()])
            else:
                email_data['categories'] = ''
            
            self.logger.debug(f"Processed email: {email_data['subject']} (ID: {email_data['entry_id']})")
            return email_data
            
        except Exception as e:
            self.logger.error(f"Error processing email: {e}", exc_info=True)
            return None
            
    def process_folder(self, folder, data, priority_emails, admin_emails, processed_folders,
                     priority_addresses, admin_addresses, start_date=None, end_date=None):
        """Process a folder and its subfolders for emails."""
        try:
            # Ensure priority_emails and admin_emails are sets
            if isinstance(priority_emails, list):
                priority_emails = set(priority_emails)
            if isinstance(admin_emails, list):
                admin_emails = set(admin_emails)
                
            folder_path = self.get_folder_path(folder)
            
            # Skip if already processed
            if folder_path in processed_folders:
                self.logger.debug(f"Skipping already processed folder: {folder_path}")
                return
                
            processed_folders.add(folder_path)
            
            self.logger.info(f"Processing folder: {folder_path}")
            processed_count = 0
            skipped_count = 0
            
            # Log folder properties for debugging
            try:
                folder_class = getattr(folder, 'Class', 'Unknown')
                default_item_type = getattr(folder, 'DefaultItemType', 'Unknown')
                self.logger.debug(f"Folder properties - Class: {folder_class}, DefaultItemType: {default_item_type}")
            except Exception as e:
                self.logger.warning(f"Could not get folder properties: {e}")
            
            # Process emails in this folder
            try:
                items = folder.Items
                items.Sort("[ReceivedTime]", True)  # Sort by received time, newest first
                item_count = items.Count
                self.logger.info(f"Found {item_count} items in {folder_path}")
                
                if item_count > 0:
                    # Process in reverse order to avoid collection modification issues
                    self.logger.info(f"Processing {item_count} items in folder {folder_path}")
                    
                    # Process items using GetFirst/GetNext pattern
                    self.logger.info(f"Starting to process items in folder {folder_path}")
                    
                    try:
                        item = items.GetFirst()
                        while item:
                            try:
                                if not hasattr(item, 'ReceivedTime'):
                                    self.logger.debug("Skipping non-email item")
                                    skipped_count += 1
                                else:
                                    # Process the email
                                    try:
                                        email_data = self.process_email(item, folder_path, priority_addresses, admin_addresses)
                                        if email_data and 'entry_id' in email_data:
                                            data.append(email_data)
                                            processed_count += 1
                                            
                                            # Track priority and admin emails
                                            sender_email = email_data.get('sender_email', '').lower()
                                            if email_data.get('is_priority') and sender_email:
                                                priority_emails.add(sender_email)
                                            if email_data.get('is_admin') and sender_email:
                                                admin_emails.add(sender_email)
                                                
                                            # Log progress every 10 emails
                                            if processed_count % 10 == 0:
                                                self.logger.info(f"Processed {processed_count} emails in folder {folder_path}")
                                        else:
                                            self.logger.warning("Skipping email - missing required data")
                                            skipped_count += 1
                                            
                                    except Exception as e:
                                        self.logger.error(f"Error processing email: {str(e)}", exc_info=True)
                                        skipped_count += 1
                                
                                # Get next item
                                item = items.GetNext()
                                
                            except Exception as e:
                                self.logger.error(f"Error in email processing loop: {str(e)}", exc_info=True)
                                skipped_count += 1
                                break
                                
                    except Exception as e:
                        self.logger.error(f"Error iterating through items: {str(e)}", exc_info=True)
                    
                    # Log completion
                    self.logger.info(f"Completed processing folder {folder_path} - Processed: {processed_count}, Skipped: {skipped_count}")
                            
                    self.logger.info(f"Processed {processed_count} emails, skipped {skipped_count} in {folder_path}")
                
            except Exception as e:
                self.logger.error(f"Error processing items in folder {folder_path}: {str(e)}", exc_info=True)
            
            # Process subfolders
            if hasattr(folder, 'Folders') and folder.Folders.Count > 0:
                subfolders = list(folder.Folders)
                subfolder_count = len(subfolders)
                self.logger.debug(f"Found {subfolder_count} subfolders in {folder_path}")
                
                for i, subfolder in enumerate(subfolders, 1):
                    try:
                        subfolder_name = getattr(subfolder, 'Name', 'Unknown')
                        self.logger.debug(f"Processing subfolder {i}/{subfolder_count}: {subfolder_name}")
                        self.process_folder(
                            subfolder, data, priority_emails, admin_emails, processed_folders,
                            priority_addresses, admin_addresses, start_date, end_date
                        )
                    except Exception as e:
                        self.logger.error(f"Error processing subfolder {i}: {e}", exc_info=True)
                        
        except Exception as e:
            self.logger.error(f"Error in process_folder: {e}", exc_info=True)

    def save_to_sqlite(self, emails: List[Dict]) -> None:
        """Save emails to SQLite database."""
        db_path = self.config.get('storage', 'sqlite_path', fallback='outlook_emails_dashboard.sqlite')
        
        try:
            # Ensure the directory exists
            os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
            
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Create or recreate emails table
            cursor.execute('''
            CREATE TABLE IF NOT EXISTS emails (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_id TEXT UNIQUE,
                folder TEXT,
                subject TEXT,
                sender TEXT,
                sender_email TEXT,
                to_recipients TEXT,
                cc_recipients TEXT,
                received_time TEXT,
                sent_on TEXT,
                body TEXT,
                body_preview TEXT,
                is_read INTEGER,
                has_attachments INTEGER,
                is_priority INTEGER,
                is_admin INTEGER,
                categories TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Create indexes
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_entry_id ON emails(entry_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_is_priority ON emails(is_priority)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_folder ON emails(folder)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_is_admin ON emails(is_admin)')
            
            # Insert emails
            for email in emails:
                cursor.execute('''
                    INSERT OR REPLACE INTO emails (
                        entry_id, folder, subject, sender, sender_email, to_recipients, 
                        cc_recipients, received_time, sent_on, body, body_preview, 
                        is_read, has_attachments, is_priority, is_admin, categories
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    email['entry_id'],
                    email['folder'],
                    email['subject'],
                    email['sender'],
                    email['sender_email'],
                    email['to_recipients'],
                    email['cc_recipients'],
                    email['received_time'],
                    email['sent_on'],
                    email['body'],
                    email['body_preview'],
                    email['is_read'],
                    email['has_attachments'],
                    email['is_priority'],
                    email['is_admin'],
                    email['categories']
                ))
            
            conn.commit()
            self.logger.info(f"Saved {len(emails)} emails to SQLite database: {os.path.abspath(db_path)}")
            
        except Exception as e:
            self.logger.error(f"Error saving to SQLite: {e}")
            raise
        finally:
            if 'conn' in locals():
                conn.close()

    def save_to_json(self, emails: List[Dict]) -> Dict:
        """Save emails to JSON file with threading information."""
        try:
            # Group emails by folder for the JSON output
            emails_by_folder = {}
            email_ids = {}
            
            # Process emails
            for idx, email in enumerate(emails, 1):
                folder = email.get('folder', 'Unknown')
                if folder not in emails_by_folder:
                    emails_by_folder[folder] = []
                
                # Ensure datetime objects are properly serialized
                if 'received_time' in email and isinstance(email['received_time'], datetime):
                    if email['received_time'].tzinfo is None:
                        email['received_time'] = email['received_time'].replace(tzinfo=timezone.utc)
                    email['received_time'] = email['received_time'].isoformat()
                
                if 'sent_on' in email and isinstance(email['sent_on'], datetime):
                    if email['sent_on'].tzinfo is None:
                        email['sent_on'] = email['sent_on'].replace(tzinfo=timezone.utc)
                    email['sent_on'] = email['sent_on'].isoformat()
                
                emails_by_folder[folder].append(email)
                email_ids[email['entry_id']] = idx
            
            # Create summary
            summary = {
                'total_emails': len(emails),
                'unique_emails': len(set(e['entry_id'] for e in emails)),
                'folders_processed': len(emails_by_folder),
                'extraction_date': datetime.now(timezone.utc).isoformat(),
                'version': VERSION
            }
            
            # Prepare output
            output = {
                'version': VERSION,
                'extraction_date': datetime.now(timezone.utc).isoformat(),
                'summary': summary,
                'emails_by_folder': emails_by_folder,
                'email_ids': email_ids
            }
            
            # Save to file
            json_path = self.config.get('storage', 'json_path', fallback='outlook_emails_dashboard.json')
            
            # Ensure the directory exists
            os.makedirs(os.path.dirname(os.path.abspath(json_path)), exist_ok=True)
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(output, f, default=str, indent=2)
                
            self.logger.info(f"Saved {len(emails)} emails to JSON file: {os.path.abspath(json_path)}")
            return output
            
        except Exception as e:
            error_msg = f"Error saving to JSON: {e}"
            self.logger.error(error_msg, exc_info=True)
            raise RuntimeError(error_msg) from e

    def print_summary(self, emails: List[Dict]):
        """Print a summary of the extraction results."""
        # Count emails by type
        total_emails = len(emails)
        unique_emails = len({e['entry_id'] for e in emails})
        
        summary = [
            "",
            "=" * 80,
            "OUTLOOK EMAIL EXTRACTION SUMMARY",
            "=" * 80,
            f"Version: {VERSION}",
            f"Extraction completed at: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')}",
            "",
            f"Total Emails Processed: {total_emails}",
            f"Unique Emails: {unique_emails}",
            f"Duplicates: {total_emails - unique_emails}",
        ]
        
        # Only show priority/admin info if not in basic_extraction_only mode
        if not self.config.getboolean('outlook', 'basic_extraction_only', fallback=False):
            priority_count = sum(1 for e in emails if e.get('is_priority'))
            admin_count = sum(1 for e in emails if e.get('is_admin'))
            summary.extend([
                f"Priority Emails: {priority_count}",
                f"Administrator Emails: {admin_count}",
            ])
            
            # Get email addresses from config
            priority_addresses = set(
                a.strip().lower() 
                for a in self.config.get('priority', 'addresses', fallback='').split(',')
                if a.strip()
            )
            
            admin_addresses = set(
                a.strip().lower()
                for a in self.config.get('administrators', 'addresses', fallback='').split(',')
                if a.strip()
            )
            
            if priority_addresses:
                summary.append("\nPriority Email Addresses:")
                for addr in sorted(priority_addresses):
                    summary.append(f"- {addr}")
            
            if admin_addresses:
                summary.append("\nAdministrator Email Addresses:")
                for addr in sorted(admin_addresses):
                    summary.append(f"- {addr}")
        
        # Date range
        if emails:
            try:
                dates = [datetime.strptime(e['received_time'], '%Y-%m-%d %H:%M:%S') for e in emails if 'received_time' in e]
                if dates:
                    min_date = min(dates)
                    max_date = max(dates)
                    summary.append(f"\nDate Range: {min_date.strftime('%Y-%m-%d')} to {max_date.strftime('%Y-%m-%d')}")
            except Exception as e:
                self.logger.warning(f"Error determining date range: {e}")
        
        # Emails by folder
        emails_by_folder = {}
        for email in emails:
            folder = email.get('folder', 'Unknown')
            emails_by_folder[folder] = emails_by_folder.get(folder, 0) + 1
        
        if emails_by_folder:
            summary.append("\nEmails by Folder:")
            for folder, count in sorted(emails_by_folder.items(), key=lambda x: x[1], reverse=True):
                summary.append(f"- {folder}: {count} emails")
        
        # Output files
        summary.append("\nOutput Files:")
        backends = [b.strip().lower() for b in self.config.get('storage', 'backend', fallback='sqlite,json').split(',')]
        if 'sqlite' in backends:
            db_path = os.path.abspath(self.config.get('storage', 'sqlite_path', fallback='outlook_emails_dashboard.sqlite'))
            summary.append(f"- SQLite Database: {db_path}")
        if 'json' in backends:
            json_path = os.path.abspath(self.config.get('storage', 'json_path', fallback='outlook_emails_dashboard.json'))
            summary.append(f"- JSON Output: {json_path}")
        
        summary.append(f"- Log File: {os.path.abspath(LOG_FILE)}")
        summary.append("\n" + "=" * 80 + "\n")
        
        # Print to console and log
        print("\n".join(summary))
        self.logger.info("\n".join(line for line in summary if not line.startswith('=')))

    def parse_date_ranges(self):
        """Parse date ranges from config."""
        try:
            # Check if date_ranges is specified
            date_ranges = self.config.get('date_range', 'date_ranges', fallback='').strip()
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
                        
                        self.logger.info(f"Using custom date range: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                        return start_date, end_date
                        
                    except (ValueError, IndexError) as e:
                        self.logger.warning(f"Invalid date_ranges format. Using days_back instead. Error: {e}")
            
            # Fall back to days_back if no valid date_ranges
            days_back = int(self.config.get('date_range', 'days_back', fallback='14'))
            end_date = datetime.now(timezone.utc)
            start_date = end_date - timedelta(days=days_back)
            self.logger.info(f"Using last {days_back} days: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            return start_date, end_date
            
        except Exception as e:
            self.logger.error(f"Error parsing date range: {e}")
            # Default to last 14 days if there's any error
            end_date = datetime.now(timezone.utc)
            start_date = end_date - timedelta(days=14)
            return start_date, end_date

    def run(self):
        """Main method to run the extraction."""
        try:
            # Get date range from config
            start_date, end_date = self.parse_date_ranges()
            self.logger.info(f"Extraction window: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
            
            # Get email addresses for priority/admin
            basic_extraction = self.config.getboolean('outlook', 'basic_extraction_only', fallback=False)
            
            if basic_extraction:
                self.logger.info("Running in basic extraction mode (priority/admin processing disabled)")
                priority_addresses = set()
                admin_addresses = set()
            else:
                priority_addresses = set(
                    a.strip().lower() 
                    for a in self.config.get('priority', 'addresses', fallback='').split(',')
                    if a.strip()
                )
                admin_addresses = set(
                    a.strip().lower()
                    for a in self.config.get('administrators', 'addresses', fallback='').split(',')
                    if a.strip()
                )
            
            # Connect to Outlook
            self.logger.info("Connecting to Outlook...")
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            
            # Get folders to process
            folder_patterns = [
                f.strip() 
                for f in self.config.get('outlook', 'folders', fallback='Inbox, Sent Items').split(',')
                if f.strip()
            ]
            
            folders_to_process = self.get_folders_to_process(outlook, folder_patterns)
            
            if not folders_to_process:
                self.logger.warning("No matching folders found. Please check your configuration.")
                return
                
            self.logger.info(f"Found {len(folders_to_process)} folders to process")
            
            # Process folders
            all_emails = []
            priority_emails = []
            admin_emails = []
            processed_folders = set()
            
            for folder, folder_path in folders_to_process:
                self.process_folder(
                    folder, all_emails, priority_emails, admin_emails, processed_folders,
                    priority_addresses, admin_addresses, start_date, end_date
                )
            
            # Save results
            backends = [b.strip().lower() for b in self.config.get('storage', 'backend', fallback='sqlite,json').split(',')]
            
            if 'sqlite' in backends:
                self.save_to_sqlite(all_emails)
            
            if 'json' in backends:
                self.save_to_json(all_emails)
            
            # Print summary
            self.print_summary(all_emails)
            
            self.logger.info("Email extraction completed successfully.")
            
        except Exception as e:
            self.logger.error(f"Error during extraction: {e}", exc_info=True)
            raise

def main():
    """Main entry point for the script."""
    try:
        extractor = OutlookExtractor()
        extractor.run()
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    sys.exit(main())
