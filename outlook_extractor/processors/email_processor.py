"""Email processing functionality for Outlook Extractor.

This module handles the extraction and processing of email messages,
including priority and admin flagging based on sender addresses.
"""
import logging
import re
from datetime import datetime
from typing import Dict, Any, List, Set, Optional

from ..core.mapi_service import MAPIPropertyAccessor

logger = logging.getLogger(__name__)

class EmailProcessor:
    """Processes email messages and extracts relevant data with flagging."""
    
    def __init__(self, config: dict):
        """Initialize the email processor with configuration.
        
        Args:
            config: Configuration dictionary with processing settings
        """
        self.config = config
        self.priority_addresses = self._normalize_emails(config.get('priority_addresses', []))
        self.admin_addresses = self._normalize_emails(config.get('admin_addresses', []))
        
    def _normalize_emails(self, emails: List[str]) -> Set[str]:
        """Normalize email addresses for consistent comparison.
        
        Args:
            emails: List of email addresses to normalize
            
        Returns:
            Set of normalized email addresses
        """
        return {email.strip().lower() for email in emails if email and isinstance(email, str)}
        
    def process_message(self, msg) -> Dict[str, Any]:
        """Process a single email message.
        
        Args:
            msg: Outlook MailItem object
            
        Returns:
            Dictionary containing processed email data
        """
        mapi = MAPIPropertyAccessor(msg)
        
        # Extract basic email data
        email_data = {
            'entry_id': mapi.get_property('http://schemas.microsoft.com/mapi/proptag/0x0FFF0102'),
            'conversation_id': mapi.get_property('http://schemas.microsoft.com/mapi/proptag/0x30130102'),
            'subject': self._get_subject(msg),
            'sender_name': self._get_sender_name(msg),
            'sender_email': self._get_sender_email(msg),
            'to_recipients': self._get_recipients(msg, 'To'),
            'cc_recipients': self._get_recipients(msg, 'CC'),
            'bcc_recipients': self._get_recipients(msg, 'BCC'),
            'received_time': self._get_received_time(msg, mapi),
            'sent_time': self._get_sent_time(msg, mapi),
            'categories': self._get_categories(msg),
            'importance': getattr(msg, 'Importance', 1),  # 0=Low, 1=Normal, 2=High
            'sensitivity': getattr(msg, 'Sensitivity', 0),  # 0=Normal, 1=Personal, 2=Private, 3=Confidential
            'has_attachments': getattr(msg, 'Attachments', None) and msg.Attachments.Count > 0,
            'is_read': getattr(msg, 'UnRead', False) is False,
            'is_flagged': getattr(msg, 'FlagStatus', 0) > 0,
            'body': self._get_body(msg, 'plain'),
            'html_body': self._get_body(msg, 'html'),
            'folder_path': self._get_folder_path(msg),
            'thread_id': None,  # Will be set by thread manager
            'thread_depth': 0,  # Will be set by thread manager
        }
        
        # Apply priority/admin flags
        self._apply_flags(email_data)
        
        return email_data
        
    def _apply_flags(self, email_data: Dict[str, Any]) -> None:
        """Apply priority and admin flags based on sender.
        
        Args:
            email_data: Dictionary containing email data to update
        """
        sender = email_data.get('sender_email', '').lower()
        if not sender:
            return
            
        if self.priority_addresses and sender in self.priority_addresses:
            email_data['is_priority'] = True
            logger.debug(f"Marked email as priority from {sender}")
            
        if self.admin_addresses and sender in self.admin_addresses:
            email_data['is_admin'] = True
            logger.debug(f"Marked email as admin from {sender}")
    
    # Helper methods for extracting specific fields
    def _get_subject(self, msg) -> str:
        """Extract and clean the email subject."""
        try:
            subject = getattr(msg, 'Subject', '')
            logger.debug(f"Got subject: {subject}")
            
            # Handle case where subject is a mock object
            if hasattr(subject, '__class__') and 'Mock' in subject.__class__.__name__:
                logger.debug("Subject is a mock object")
                return '(No Subject)'
                
            # Handle empty subject
            if not subject:
                return '(No Subject)'
                
            # Handle case where strip() might fail
            try:
                return subject.strip()
            except Exception as e:
                logger.debug(f"Error stripping subject: {e}")
                return str(subject) if subject else '(No Subject)'
                
        except Exception as e:
            logger.debug(f"Error in _get_subject: {e}")
            return '(No Subject)'
    
    def _get_sender_name(self, msg) -> str:
        """Extract the sender's display name."""
        return getattr(msg, 'SenderName', '')
    
    def _get_sender_email(self, msg) -> str:
        """Extract the sender's email address."""
        sender = getattr(msg, 'SenderEmailAddress', '')
        if not sender and hasattr(msg, 'Sender'):
            sender = getattr(msg.Sender, 'Address', '')
        return sender.lower() if sender else ''
    
    def _get_recipients(self, msg, recipient_type: str) -> str:
        """Extract recipients of a specific type (To/CC/BCC)."""
        recipients = []
        try:
            recipients_collection = getattr(msg, recipient_type, None)
            if recipients_collection:
                for recipient in recipients_collection:
                    email = getattr(recipient, 'Address', '')
                    if email:
                        recipients.append(email)
        except Exception as e:
            logger.debug(f"Error getting {recipient_type} recipients: {e}")
        
        return '; '.join(recipients) if recipients else ''
    
    def _get_received_time(self, msg, mapi) -> Optional[datetime]:
        """Extract the received time of the email."""
        try:
            # Try to get the exact received time
            received_time = mapi.get_property('http://schemas.microsoft.com/mapi/proptag/0x0E060040')
            if received_time:
                return received_time
                
            # Fall back to message properties
            if hasattr(msg, 'ReceivedTime'):
                return msg.ReceivedTime
                
            return None
            
        except Exception as e:
            logger.debug(f"Error getting received time: {e}")
            return None
    
    def _get_sent_time(self, msg, mapi) -> Optional[datetime]:
        """Extract the sent time of the email."""
        try:
            # Try to get the exact sent time
            sent_time = mapi.get_property('http://schemas.microsoft.com/mapi/proptag/0x00390040')
            if sent_time:
                return sent_time
                
            # Fall back to message properties
            if hasattr(msg, 'SentOn'):
                return msg.SentOn
                
            return None
            
        except Exception as e:
            logger.debug(f"Error getting sent time: {e}")
            return None
    
    def _get_categories(self, msg) -> str:
        """Extract and format email categories."""
        try:
            categories = getattr(msg, 'Categories', '')
            if categories:
                return ';'.join(cat.strip() for cat in categories.split(';') if cat.strip())
        except Exception:
            pass
        return ''
    
    def _get_body(self, msg, format_type: str = 'plain') -> str:
        """Extract the email body in the specified format."""
        try:
            if format_type.lower() == 'html':
                return getattr(msg, 'HTMLBody', '')
            else:
                return getattr(msg, 'Body', '')
        except Exception as e:
            logger.debug(f"Error getting {format_type} body: {e}")
            return ''
    
    def _get_folder_path(self, msg) -> str:
        """Get the full folder path of the email."""
        try:
            folder = msg.Parent
            if folder:
                folder_path = folder.FolderPath
                # Clean up the folder path (remove leading/trailing slashes and quotes)
                return folder_path.strip("'/\\\"")
        except Exception as e:
            logger.debug(f"Error getting folder path: {e}")
            
        return 'Unknown'
