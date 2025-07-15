"""Email threading module for organizing emails into conversation threads."""

from __future__ import annotations

import hashlib
import re
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email.utils import getaddresses
from typing import Any, Dict, List, Optional, Set, DefaultDict

# Thread status constants
THREAD_STATUS_ACTIVE = 'active'
THREAD_STATUS_RESOLVED = 'resolved'
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
    
    def add_email(self, email_data: Dict[str, Any]) -> None:
        """Add an email to this thread.
        
        Args:
            email_data: Dictionary containing email data
        """
        if not self.root_message_id and not self.message_ids:
            self.root_message_id = email_data.get('message_id')
            self.subject = email_data.get('subject', '(No Subject)')
        
        self.message_ids.add(email_data['entry_id'])
        
        # Update participants
        self.participants.update(self._extract_participants(email_data))
        
        # Update date range
        email_date = self._parse_date(email_data.get('sent_on') or email_data.get('received_time'))
        if email_date:
            if not self.start_date or email_date < self.start_date:
                self.start_date = email_date
            if not self.end_date or email_date > self.end_date:
                self.end_date = email_date
        
        # Update categories
        if email_data.get('categories'):
            self.categories.update(cat.strip() for cat in email_data['categories'].split(','))
    
    def _extract_participants(self, email_data: Dict[str, Any]) -> Set[str]:
        """Extract all participants from an email.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            Set of participant email addresses
        """
        participants = set()
        
        def add_emails(field):
            if not field:
                return
            for name, addr in getaddresses([field]):
                if '@' in addr:
                    participants.add(addr.lower())
        
        add_emails(email_data.get('sender_email', ''))
        add_emails(email_data.get('to_recipients', ''))
        add_emails(email_data.get('cc_recipients', ''))
        
        return participants
    
    def _parse_date(self, date_str: Optional[str]) -> Optional[datetime]:
        """Parse a date string into a datetime object.
        
        Args:
            date_str: Date string to parse
            
        Returns:
            datetime object or None if parsing fails
        """
        if not date_str:
            return None
        
        try:
            # Try parsing with timezone
            for fmt in (
                '%Y-%m-%d %H:%M:%S%z',
                '%Y-%m-%d %H:%M:%S',
                '%Y-%m-%d',
                '%m/%d/%Y %H:%M:%S %p',
                '%m/%d/%Y %H:%M:%S',
                '%m/%d/%Y',
            ):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    if dt.tzinfo is None:
                        dt = dt.replace(tzinfo=timezone.utc)
                    return dt
                except ValueError:
                    continue
        except Exception:
            pass
        
        return None
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert thread to dictionary for serialization.
        
        Returns:
            Dictionary representation of the thread
        """
        return {
            'thread_id': self.thread_id,
            'subject': self.subject,
            'participants': list(self.participants),
            'message_ids': list(self.message_ids),
            'root_message_id': self.root_message_id,
            'start_date': self.start_date.isoformat() if self.start_date else None,
            'end_date': self.end_date.isoformat() if self.end_date else None,
            'status': self.status,
            'categories': list(self.categories)
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'EmailThread':
        """Create an EmailThread from a dictionary.
        
        Args:
            data: Dictionary containing thread data
            
        Returns:
            EmailThread instance
        """
        thread = cls(
            thread_id=data['thread_id'],
            subject=data['subject'],
            participants=set(data.get('participants', [])),
            message_ids=set(data.get('message_ids', [])),
            root_message_id=data.get('root_message_id'),
            status=data.get('status', THREAD_STATUS_ACTIVE),
            categories=set(data.get('categories', []))
        )
        
        # Parse dates
        if data.get('start_date'):
            thread.start_date = datetime.fromisoformat(data['start_date'])
        if data.get('end_date'):
            thread.end_date = datetime.fromisoformat(data['end_date'])
            
        return thread


class ThreadManager:
    """Manages email threads and their relationships."""
    
    def __init__(self):
        self.threads_by_id: Dict[str, EmailThread] = {}
        self.message_to_thread: Dict[str, str] = {}
        self.threads_by_participant: DefaultDict[str, Set[str]] = defaultdict(set)
    
    def add_email(self, email_data: Dict[str, Any]) -> None:
        """Add an email to the appropriate thread.
        
        Args:
            email_data: Dictionary containing email data
        """
        message_id = email_data.get('message_id')
        if not message_id or message_id in self.message_to_thread:
            return  # Already processed this message
        
        # Check if this is a reply to an existing message
        in_reply_to = email_data.get('in_reply_to')
        references = self._parse_references(email_data.get('references', ''))
        
        # Find existing thread or create a new one
        thread = self._find_existing_thread(in_reply_to, references) or self._create_new_thread(email_data)
        
        # Add email to thread
        thread.add_email(email_data)
        self.message_to_thread[message_id] = thread.thread_id
        
        # Update participant index
        for participant in thread.participants:
            self.threads_by_participant[participant].add(thread.thread_id)
    
    def _find_existing_thread(self, in_reply_to: Optional[str], references: List[str]) -> Optional[EmailThread]:
        """Find an existing thread based on reply/reference headers.
        
        Args:
            in_reply_to: Message-ID this email is in reply to
            references: List of Message-IDs in the References header
            
        Returns:
            Existing EmailThread or None if not found
        """
        # Check direct reply
        if in_reply_to and in_reply_to in self.message_to_thread:
            return self.threads_by_id[self.message_to_thread[in_reply_to]]
        
        # Check references
        for ref in reversed(references):
            if ref in self.message_to_thread:
                return self.threads_by_id[self.message_to_thread[ref]]
        
        return None
    
    def _create_new_thread(self, email_data: Dict[str, Any]) -> EmailThread:
        """Create a new thread for an email.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            New EmailThread instance
        """
        thread_id = self._generate_thread_id(email_data)
        thread = EmailThread(thread_id=thread_id, subject=email_data.get('subject', '(No Subject)'))
        self.threads_by_id[thread_id] = thread
        return thread
    
    def _generate_thread_id(self, email_data: Dict[str, Any]) -> str:
        """Generate a unique thread ID based on email headers.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            Generated thread ID
        """
        # Use existing thread ID if present
        if 'thread_index' in email_data:
            return f"thread_{email_data['thread_index']}"
        
        # Fallback: Generate ID from subject and participants
        subject = email_data.get('subject', '').lower()
        participants = set()
        
        def add_participants(header):
            if not header:
                return
            for name, addr in getaddresses([header]):
                if '@' in addr:
                    participants.add(addr.lower())
        
        add_participants(email_data.get('sender_email', ''))
        add_participants(email_data.get('to_recipients', ''))
        add_participants(email_data.get('cc_recipients', ''))
        
        # Create a stable hash of subject + sorted participants
        key = f"{subject}:{':'.join(sorted(participants))}"
        return f"thread_{hashlib.md5(key.encode('utf-8')).hexdigest()[:16]}"
    
    def _parse_references(self, references: str) -> List[str]:
        """Parse the References header into individual message IDs.
        
        Args:
            references: References header value
            
        Returns:
            List of message IDs
        """
        if not references:
            return []
        return [ref.strip() for ref in re.split(r'\s+', references) if ref.strip()]
    
    def get_threads(self) -> List[Dict[str, Any]]:
        """Get all threads as a list of dictionaries.
        
        Returns:
            List of thread dictionaries
        """
        return [thread.to_dict() for thread in self.threads_by_id.values()]
    
    def get_thread(self, thread_id: str) -> Optional[Dict[str, Any]]:
        """Get a specific thread by ID.
        
        Args:
            thread_id: ID of the thread to retrieve
            
        Returns:
            Thread dictionary or None if not found
        """
        thread = self.threads_by_id.get(thread_id)
        return thread.to_dict() if thread else None
    
    def get_threads_for_participant(self, email: str) -> List[Dict[str, Any]]:
        """Get all threads involving a specific email address.
        
        Args:
            email: Email address to search for
            
        Returns:
            List of thread dictionaries
        """
        email = email.lower()
        return [
            self.threads_by_id[tid].to_dict()
            for tid in self.threads_by_participant.get(email, [])
            if tid in self.threads_by_id
        ]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert the thread manager to a dictionary for serialization.
        
        Returns:
            Dictionary representation of the thread manager
        """
        return {
            'threads': [t.to_dict() for t in self.threads_by_id.values()],
            'message_to_thread': self.message_to_thread,
            'threads_by_participant': {
                k: list(v) for k, v in self.threads_by_participant.items()
            }
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ThreadManager':
        """Create a ThreadManager from a dictionary.
        
        Args:
            data: Dictionary containing thread manager data
            
        Returns:
            ThreadManager instance
        """
        manager = cls()
        
        # Rebuild threads
        for thread_data in data.get('threads', []):
            thread = EmailThread.from_dict(thread_data)
            manager.threads_by_id[thread.thread_id] = thread
        
        # Rebuild message to thread mapping
        manager.message_to_thread = data.get('message_to_thread', {})
        
        # Rebuild participant index
        manager.threads_by_participant = defaultdict(
            set,
            {k: set(v) for k, v in data.get('threads_by_participant', {}).items()}
        )
        
        return manager
