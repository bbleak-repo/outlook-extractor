"""
JSON storage implementation for email data.
"""
import os
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Any, Optional, Union, Set

from ..config import get_config
from .base import EmailStorage

logger = logging.getLogger(__name__)

class JSONStorage(EmailStorage):
    """JSON file-based storage implementation for email data."""
    
    def __init__(self, json_path: str = None, config=None):
        """Initialize the JSON storage.
        
        Args:
            json_path: Path to the JSON file. If None, uses the path from config.
            config: Optional ConfigManager instance. If not provided, uses default config.
        """
        self.config = config or get_config()
        self.json_path = json_path or self.config.get('storage', 'json_path', 'emails.json')
        self.data = {
            'emails': {},
            'threads': {},
            'metadata': {
                'version': '1.0',
                'created_at': datetime.now(timezone.utc).isoformat(),
                'updated_at': datetime.now(timezone.utc).isoformat(),
                'email_count': 0
            }
        }
        self._load_data()
    
    def _load_data(self) -> None:
        """Load data from the JSON file if it exists."""
        try:
            if os.path.exists(self.json_path):
                with open(self.json_path, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
                logger.info(f"Loaded {len(self.data['emails'])} emails from {self.json_path}")
            else:
                # Ensure the directory exists
                json_dir = os.path.dirname(self.json_path)
                if json_dir and not os.path.exists(json_dir):
                    os.makedirs(json_dir, exist_ok=True)
                self._save_data()
        except Exception as e:
            logger.error(f"Error loading JSON data: {e}", exc_info=True)
    
    def _save_data(self) -> None:
        """Save data to the JSON file."""
        try:
            # Update metadata
            self.data['metadata']['updated_at'] = datetime.now(timezone.utc).isoformat()
            self.data['metadata']['email_count'] = len(self.data['emails'])
            
            # Save to file atomically by writing to a temp file first
            temp_path = f"{self.json_path}.tmp"
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(
                    self.data, 
                    f, 
                    indent=2, 
                    ensure_ascii=False,
                    default=self._json_serializer
                )
            
            # Replace the original file
            if os.path.exists(self.json_path):
                os.replace(temp_path, self.json_path)
            else:
                os.rename(temp_path, self.json_path)
                
        except Exception as e:
            logger.error(f"Error saving JSON data: {e}", exc_info=True)
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
    
    def _json_serializer(self, obj):
        """JSON serializer for objects not serializable by default json code."""
        if isinstance(obj, datetime):
            return obj.isoformat()
        raise TypeError(f"Type {type(obj)} not serializable")
    
    def save_email(self, email_data: Dict[str, Any]) -> bool:
        """Save a single email to the JSON file."""
        try:
            email_id = email_data.get('id')
            if not email_id:
                logger.warning("Email data missing 'id' field, skipping")
                return False
            
            # Store the email
            self.data['emails'][email_id] = email_data
            
            # Update thread information if thread_id is present
            thread_id = email_data.get('thread_id')
            if thread_id:
                if thread_id not in self.data['threads']:
                    self.data['threads'][thread_id] = {
                        'id': thread_id,
                        'subject': email_data.get('subject', ''),
                        'participants': set(),
                        'message_ids': set(),
                        'start_date': email_data.get('sent_date') or email_data.get('received_date'),
                        'end_date': email_data.get('sent_date') or email_data.get('received_date'),
                        'status': 'active',
                        'categories': set(email_data.get('categories', [])),
                        'created_at': datetime.now(timezone.utc).isoformat(),
                        'updated_at': datetime.now(timezone.utc).isoformat()
                    }
                
                # Update thread information
                thread = self.data['threads'][thread_id]
                thread['message_ids'].add(email_id)
                
                # Update participants
                for field in ['sender', 'recipients', 'cc_recipients', 'bcc_recipients']:
                    if field in email_data and email_data[field]:
                        if isinstance(email_data[field], str):
                            thread['participants'].add(email_data[field])
                        elif isinstance(email_data[field], list):
                            thread['participants'].update(email_data[field])
                
                # Update dates
                email_date = email_data.get('sent_date') or email_data.get('received_date')
                if email_date:
                    if not thread['start_date'] or email_date < thread['start_date']:
                        thread['start_date'] = email_date
                    if not thread['end_date'] or email_date > thread['end_date']:
                        thread['end_date'] = email_date
                
                # Update categories
                if 'categories' in email_data and email_data['categories']:
                    if isinstance(email_data['categories'], list):
                        thread['categories'].update(email_data['categories'])
                    else:
                        thread['categories'].add(email_data['categories'])
                
                thread['updated_at'] = datetime.now(timezone.utc).isoformat()
            
            # Save the data
            self._save_data()
            return True
            
        except Exception as e:
            logger.error(f"Error saving email to JSON: {e}", exc_info=True)
            return False
    
    def save_emails(self, emails: List[Dict[str, Any]]) -> int:
        """Save multiple emails to the JSON file."""
        if not emails:
            return 0
            
        saved_count = 0
        for email in emails:
            if self.save_email(email):
                saved_count += 1
        return saved_count
    
    def get_email(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Retrieve a single email by its ID."""
        return self.data['emails'].get(email_id)
    
    def get_emails_by_sender(self, sender: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by sender email address."""
        results = []
        for email in self.data['emails'].values():
            if 'sender' in email and sender.lower() in email['sender'].lower():
                results.append(email)
                if len(results) >= limit:
                    break
        return results
    
    def get_emails_by_recipient(self, recipient: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by recipient email address."""
        results = []
        recipient_lower = recipient.lower()
        
        for email in self.data['emails'].values():
            found = False
            
            # Check sender
            if 'sender' in email and recipient_lower in email['sender'].lower():
                found = True
            
            # Check recipients
            for field in ['recipients', 'cc_recipients', 'bcc_recipients']:
                if field in email and email[field]:
                    if isinstance(email[field], str) and recipient_lower in email[field].lower():
                        found = True
                        break
                    elif isinstance(email[field], list) and any(
                        recipient_lower in addr.lower() for addr in email[field] if addr
                    ):
                        found = True
                        break
            
            if found:
                results.append(email)
                if len(results) >= limit:
                    break
        
        return results
    
    def get_emails_by_date_range(self, start_date: datetime, end_date: datetime, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails within a date range."""
        results = []
        
        for email in self.data['emails'].values():
            # Check sent_date or received_date
            email_date = None
            if 'sent_date' in email and email['sent_date']:
                email_date = self._parse_date(email['sent_date'])
            elif 'received_date' in email and email['received_date']:
                email_date = self._parse_date(email['received_date'])
            
            if email_date and start_date <= email_date <= end_date:
                results.append(email)
                if len(results) >= limit:
                    break
        
        # Sort by date (newest first)
        results.sort(
            key=lambda x: self._get_email_date(x) or datetime.min,
            reverse=True
        )
        
        return results
    
    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse a date string into a datetime object."""
        if not date_str:
            return None
            
        try:
            if isinstance(date_str, str):
                if 'T' in date_str:  # ISO format
                    return datetime.fromisoformat(date_str.replace('Z', '+00:00'))
                else:
                    # Try common date formats
                    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y'):
                        try:
                            return datetime.strptime(date_str, fmt)
                        except ValueError:
                            continue
        except (ValueError, TypeError):
            pass
            
        return None
    
    def _get_email_date(self, email: Dict[str, Any]) -> Optional[datetime]:
        """Get the most relevant date from an email."""
        if 'sent_date' in email and email['sent_date']:
            return self._parse_date(email['sent_date'])
        elif 'received_date' in email and email['received_date']:
            return self._parse_date(email['received_date'])
        return None
    
    def search_emails(self, query: str, fields: List[str] = None, limit: int = 100) -> List[Dict[str, Any]]:
        """Search for emails matching the query."""
        if not query:
            return []
            
        query = query.lower()
        results = []
        
        # Default fields to search if none specified
        if not fields:
            fields = ['subject', 'body_text', 'sender', 'recipients']
        
        for email in self.data['emails'].values():
            match_found = False
            
            for field in fields:
                if field not in email or not email[field]:
                    continue
                    
                field_value = email[field]
                
                # Handle different field types
                if isinstance(field_value, str):
                    if query in field_value.lower():
                        match_found = True
                        break
                elif isinstance(field_value, list):
                    # Check if any item in the list contains the query
                    if any(query in str(item).lower() for item in field_value if item):
                        match_found = True
                        break
                elif field_value is not None:
                    # Convert to string and search
                    if query in str(field_value).lower():
                        match_found = True
                        break
            
            if match_found:
                results.append(email)
                if len(results) >= limit:
                    break
        
        # Sort by date (newest first)
        results.sort(
            key=lambda x: self._get_email_date(x) or datetime.min,
            reverse=True
        )
        
        return results
    
    def get_unique_senders(self) -> Set[str]:
        """Get all unique email senders in the storage."""
        senders = set()
        for email in self.data['emails'].values():
            if 'sender' in email and email['sender']:
                senders.add(email['sender'])
        return senders
    
    def get_unique_recipients(self) -> Set[str]:
        """Get all unique email recipients in the storage."""
        recipients = set()
        
        for email in self.data['emails'].values():
            for field in ['recipients', 'cc_recipients', 'bcc_recipients']:
                if field in email and email[field]:
                    if isinstance(email[field], str):
                        recipients.add(email[field])
                    elif isinstance(email[field], list):
                        recipients.update(addr for addr in email[field] if addr)
        
        return recipients
    
    def get_email_count(self) -> int:
        """Get the total number of emails in the storage."""
        return len(self.data['emails'])
    
    def close(self) -> None:
        """Close the storage and save any pending changes."""
        self._save_data()
    
    def __del__(self):
        """Ensure data is saved when the object is destroyed."""
        self.close()
