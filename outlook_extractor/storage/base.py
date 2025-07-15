"""
Base storage interface for email data persistence.
"""
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any, Set
from datetime import datetime

class EmailStorage(ABC):    
    """Abstract base class for email storage backends."""
    
    @abstractmethod
    def save_email(self, email_data: Dict[str, Any]) -> bool:
        """Save a single email to the storage.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            bool: True if the email was saved successfully, False otherwise
        """
        pass
    
    @abstractmethod
    def save_emails(self, emails: List[Dict[str, Any]]) -> int:
        """Save multiple emails to the storage.
        
        Args:
            emails: List of email data dictionaries
            
        Returns:
            int: Number of emails successfully saved
        """
        pass
    
    @abstractmethod
    def get_email(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Retrieve a single email by its ID.
        
        Args:
            email_id: Unique identifier for the email
            
        Returns:
            Optional[Dict]: The email data if found, None otherwise
        """
        pass
    
    @abstractmethod
    def get_emails_by_sender(self, sender: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by sender email address.
        
        Args:
            sender: Email address of the sender
            limit: Maximum number of emails to return
            
        Returns:
            List of matching email data dictionaries
        """
        pass
    
    @abstractmethod
    def get_emails_by_recipient(self, recipient: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by recipient email address.
        
        Args:
            recipient: Email address of the recipient
            limit: Maximum number of emails to return
            
        Returns:
            List of matching email data dictionaries
        """
        pass
    
    @abstractmethod
    def get_emails_by_date_range(self, 
                               start_date: datetime, 
                               end_date: datetime, 
                               limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails within a date range.
        
        Args:
            start_date: Start of the date range (inclusive)
            end_date: End of the date range (inclusive)
            limit: Maximum number of emails to return
            
        Returns:
            List of matching email data dictionaries
        """
        pass
    
    @abstractmethod
    def search_emails(self, 
                     query: str, 
                     fields: List[str] = None, 
                     limit: int = 100) -> List[Dict[str, Any]]:
        """Search for emails matching the query.
        
        Args:
            query: Search query string
            fields: List of fields to search in (None for all searchable fields)
            limit: Maximum number of results to return
            
        Returns:
            List of matching email data dictionaries
        """
        pass
    
    @abstractmethod
    def get_unique_senders(self) -> Set[str]:
        """Get all unique email senders in the storage.
        
        Returns:
            Set of unique sender email addresses
        """
        pass
    
    @abstractmethod
    def get_unique_recipients(self) -> Set[str]:
        """Get all unique email recipients in the storage.
        
        Returns:
            Set of unique recipient email addresses
        """
        pass
    
    @abstractmethod
    def get_email_count(self) -> int:
        """Get the total number of emails in the storage.
        
        Returns:
            int: Total number of emails
        """
        pass
    
    @abstractmethod
    def close(self) -> None:
        """Close the storage and release any resources."""
        pass
    
    def __enter__(self):
        """Context manager entry."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - ensure resources are cleaned up."""
        self.close()
