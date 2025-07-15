"""
SQLite storage implementation for email data.
"""
import os
import sqlite3
import json
import logging
from datetime import datetime
from typing import Dict, List, Optional, Any, Set, Union, Tuple
from pathlib import Path

from ..config import get_config
from .base import EmailStorage

logger = logging.getLogger(__name__)

class SQLiteStorage(EmailStorage):
    """SQLite storage implementation for email data."""
    
    def __init__(self, db_path: str = None, config=None):
        """Initialize the SQLite storage.
        
        Args:
            db_path: Path to the SQLite database file. If None, uses the path from config.
            config: Optional ConfigManager instance. If not provided, uses default config.
        """
        self.config = config or get_config()
        self.db_path = db_path or self.config.get('storage', 'db_path', 'emails.db')
        self.conn = None
        self._ensure_db()
    
    def _ensure_db(self) -> None:
        """Ensure the database and tables exist."""
        # Ensure the directory exists
        db_dir = os.path.dirname(self.db_path)
        if db_dir and not os.path.exists(db_dir):
            os.makedirs(db_dir, exist_ok=True)
        
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row  # Enable column access by name
        
        # Create tables if they don't exist
        with self.conn:
            # Emails table
            self.conn.execute('''
            CREATE TABLE IF NOT EXISTS emails (
                id TEXT PRIMARY KEY,
                thread_id TEXT,
                subject TEXT,
                sender TEXT,
                recipients TEXT,  -- JSON array of email addresses
                cc_recipients TEXT,  -- JSON array of email addresses
                bcc_recipients TEXT,  -- JSON array of email addresses
                sent_date TIMESTAMP,
                received_date TIMESTAMP,
                body_text TEXT,
                body_html TEXT,
                is_read BOOLEAN,
                importance INTEGER,
                has_attachments BOOLEAN,
                categories TEXT,  -- JSON array of categories
                internet_headers TEXT,  -- JSON object of internet headers
                folder_path TEXT,
                raw_data TEXT,  -- Full JSON of the email
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
            
            # Create indexes for faster lookups
            self.conn.execute('CREATE INDEX IF NOT EXISTS idx_emails_sender ON emails(sender)')
            self.conn.execute('CREATE INDEX IF NOT EXISTS idx_emails_thread_id ON emails(thread_id)')
            self.conn.execute('CREATE INDEX IF NOT EXISTS idx_emails_sent_date ON emails(sent_date)')
            self.conn.execute('CREATE INDEX IF NOT EXISTS idx_emails_received_date ON emails(received_date)')
            
            # Threads table
            self.conn.execute('''
            CREATE TABLE IF NOT EXISTS threads (
                id TEXT PRIMARY KEY,
                subject TEXT,
                participants TEXT,  -- JSON array of email addresses
                message_ids TEXT,   -- JSON array of message IDs
                start_date TIMESTAMP,
                end_date TIMESTAMP,
                status TEXT,
                categories TEXT,   -- JSON array of categories
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')
    
    def _dict_factory(self, cursor, row):
        """Convert database row to dictionary."""
        d = {}
        for idx, col in enumerate(cursor.description):
            d[col[0]] = row[idx]
            
            # Deserialize JSON fields
            if col[0] in ['recipients', 'cc_recipients', 'bcc_recipients', 'categories', 
                         'internet_headers', 'participants', 'message_ids'] and row[idx]:
                try:
                    d[col[0]] = json.loads(row[idx])
                except (json.JSONDecodeError, TypeError):
                    d[col[0]] = row[idx]
        return d
    
    def save_email(self, email_data: Dict[str, Any]) -> bool:
        """Save a single email to the database."""
        try:
            # Prepare data for insertion
            email_id = email_data.get('id')
            if not email_id:
                logger.warning("Email data missing 'id' field, skipping")
                return False
                
            # Convert lists/dicts to JSON strings
            recipients = json.dumps(email_data.get('recipients', []))
            cc_recipients = json.dumps(email_data.get('cc_recipients', []))
            bcc_recipients = json.dumps(email_data.get('bcc_recipients', []))
            categories = json.dumps(email_data.get('categories', []))
            internet_headers = json.dumps(email_data.get('internet_headers', {}))
            
            # Convert dates to ISO format strings
            sent_date = email_data.get('sent_date')
            if sent_date and isinstance(sent_date, datetime):
                sent_date = sent_date.isoformat()
                
            received_date = email_data.get('received_date')
            if received_date and isinstance(received_date, datetime):
                received_date = received_date.isoformat()
            
            # Check if email already exists
            cursor = self.conn.cursor()
            cursor.execute('SELECT id FROM emails WHERE id = ?', (email_id,))
            exists = cursor.fetchone() is not None
            
            if exists:
                # Update existing email
                query = '''
                UPDATE emails 
                SET thread_id = ?, subject = ?, sender = ?, recipients = ?, 
                    cc_recipients = ?, bcc_recipients = ?, sent_date = ?, 
                    received_date = ?, body_text = ?, body_html = ?, 
                    is_read = ?, importance = ?, has_attachments = ?, 
                    categories = ?, internet_headers = ?, folder_path = ?, 
                    raw_data = ?, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                '''
                params = (
                    email_data.get('thread_id'),
                    email_data.get('subject'),
                    email_data.get('sender'),
                    recipients,
                    cc_recipients,
                    bcc_recipients,
                    sent_date,
                    received_date,
                    email_data.get('body_text'),
                    email_data.get('body_html'),
                    1 if email_data.get('is_read') else 0,
                    email_data.get('importance', 1),  # Default to normal importance
                    1 if email_data.get('has_attachments') else 0,
                    categories,
                    internet_headers,
                    email_data.get('folder_path'),
                    json.dumps(email_data),
                    email_id
                )
            else:
                # Insert new email
                query = '''
                INSERT INTO emails (
                    id, thread_id, subject, sender, recipients, cc_recipients, 
                    bcc_recipients, sent_date, received_date, body_text, body_html, 
                    is_read, importance, has_attachments, categories, internet_headers, 
                    folder_path, raw_data
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                '''
                params = (
                    email_id,
                    email_data.get('thread_id'),
                    email_data.get('subject'),
                    email_data.get('sender'),
                    recipients,
                    cc_recipients,
                    bcc_recipients,
                    sent_date,
                    received_date,
                    email_data.get('body_text'),
                    email_data.get('body_html'),
                    1 if email_data.get('is_read') else 0,
                    email_data.get('importance', 1),
                    1 if email_data.get('has_attachments') else 0,
                    categories,
                    internet_headers,
                    email_data.get('folder_path'),
                    json.dumps(email_data)
                )
            
            # Execute the query
            cursor.execute(query, params)
            self.conn.commit()
            return True
            
        except Exception as e:
            logger.error(f"Error saving email to database: {e}", exc_info=True)
            if self.conn:
                self.conn.rollback()
            return False
    
    def save_emails(self, emails: List[Dict[str, Any]]) -> int:
        """Save multiple emails to the database."""
        if not emails:
            return 0
            
        saved_count = 0
        for email in emails:
            if self.save_email(email):
                saved_count += 1
        return saved_count
    
    def get_email(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Retrieve a single email by its ID."""
        try:
            cursor = self.conn.cursor()
            cursor.row_factory = self._dict_factory
            cursor.execute('SELECT * FROM emails WHERE id = ?', (email_id,))
            row = cursor.fetchone()
            
            if not row:
                return None
                
            # Convert row to dict and deserialize JSON fields
            email = dict(row)
            
            # Convert string dates back to datetime objects
            for date_field in ['sent_date', 'received_date', 'created_at', 'updated_at']:
                if date_field in email and email[date_field]:
                    try:
                        email[date_field] = datetime.fromisoformat(email[date_field])
                    except (ValueError, TypeError):
                        pass
            
            # Deserialize raw_data if present
            if 'raw_data' in email and email['raw_data']:
                try:
                    email.update(json.loads(email['raw_data']))
                except (json.JSONDecodeError, TypeError):
                    pass
                    
            return email
            
        except Exception as e:
            logger.error(f"Error retrieving email {email_id}: {e}", exc_info=True)
            return None
    
    def get_emails_by_sender(self, sender: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by sender email address."""
        try:
            cursor = self.conn.cursor()
            cursor.row_factory = self._dict_factory
            cursor.execute(
                'SELECT * FROM emails WHERE sender LIKE ? ORDER BY sent_date DESC LIMIT ?',
                (f'%{sender}%', limit)
            )
            return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Error retrieving emails by sender {sender}: {e}", exc_info=True)
            return []
    
    def get_emails_by_recipient(self, recipient: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails by recipient email address."""
        try:
            cursor = self.conn.cursor()
            cursor.row_factory = self._dict_factory
            
            # Search in recipients, cc, and bcc fields
            query = """
            SELECT DISTINCT e.* 
            FROM emails e
            WHERE json_extract(e.recipients, '$') LIKE ?
               OR json_extract(e.cc_recipients, '$') LIKE ?
               OR json_extract(e.bcc_recipients, '$') LIKE ?
            ORDER BY e.sent_date DESC
            LIMIT ?
            """
            pattern = f'%{recipient}%'
            cursor.execute(query, (pattern, pattern, pattern, limit))
            return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Error retrieving emails by recipient {recipient}: {e}", exc_info=True)
            return []
    
    def get_emails_by_date_range(self, start_date: datetime, end_date: datetime, limit: int = 100) -> List[Dict[str, Any]]:
        """Retrieve emails within a date range."""
        try:
            cursor = self.conn.cursor()
            cursor.row_factory = self._dict_factory
            
            query = """
            SELECT * FROM emails 
            WHERE (sent_date BETWEEN ? AND ?)
               OR (received_date BETWEEN ? AND ?)
            ORDER BY COALESCE(sent_date, received_date) DESC
            LIMIT ?
            """
            
            # Convert dates to ISO format strings
            start_iso = start_date.isoformat()
            end_iso = end_date.isoformat()
            
            cursor.execute(query, (start_iso, end_iso, start_iso, end_iso, limit))
            return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"Error retrieving emails by date range: {e}", exc_info=True)
            return []
    
    def search_emails(self, query: str, fields: List[str] = None, limit: int = 100) -> List[Dict[str, Any]]:
        """Search for emails matching the query."""
        if not query:
            return []
            
        try:
            cursor = self.conn.cursor()
            cursor.row_factory = self._dict_factory
            
            # Default fields to search if none specified
            if not fields:
                fields = ['subject', 'body_text', 'sender', 'recipients']
            
            # Build the WHERE clause dynamically based on the fields to search
            where_clauses = []
            params = []
            
            for field in fields:
                if field in ['subject', 'body_text', 'sender']:
                    where_clauses.append(f"{field} LIKE ?")
                    params.append(f'%{query}%')
                elif field == 'recipients':
                    where_clauses.append("recipients LIKE ?")
                    where_clauses.append("cc_recipients LIKE ?")
                    where_clauses.append("bcc_recipients LIKE ?")
                    params.extend([f'%{query}%'] * 3)
            
            if not where_clauses:
                return []
                
            where_clause = " OR ".join(where_clauses)
            sql = f"""
            SELECT * FROM emails 
            WHERE {where_clause}
            ORDER BY COALESCE(sent_date, received_date) DESC
            LIMIT ?
            """
            
            params.append(limit)
            cursor.execute(sql, params)
            return [dict(row) for row in cursor.fetchall()]
            
        except Exception as e:
            logger.error(f"Error searching emails: {e}", exc_info=True)
            return []
    
    def get_unique_senders(self) -> Set[str]:
        """Get all unique email senders in the database."""
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT DISTINCT sender FROM emails WHERE sender IS NOT NULL')
            return {row[0] for row in cursor.fetchall() if row[0]}
        except Exception as e:
            logger.error(f"Error getting unique senders: {e}", exc_info=True)
            return set()
    
    def get_unique_recipients(self) -> Set[str]:
        """Get all unique email recipients in the database."""
        try:
            cursor = self.conn.cursor()
            
            # Get recipients from all recipient fields
            cursor.execute('SELECT recipients, cc_recipients, bcc_recipients FROM emails')
            
            recipients = set()
            for row in cursor.fetchall():
                for field in row:
                    if field:
                        try:
                            email_list = json.loads(field)
                            if isinstance(email_list, list):
                                recipients.update(email_list)
                        except (json.JSONDecodeError, TypeError):
                            pass
            
            return recipients
            
        except Exception as e:
            logger.error(f"Error getting unique recipients: {e}", exc_info=True)
            return set()
    
    def get_email_count(self) -> int:
        """Get the total number of emails in the database."""
        try:
            cursor = self.conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM emails')
            return cursor.fetchone()[0] or 0
        except Exception as e:
            logger.error(f"Error getting email count: {e}", exc_info=True)
            return 0
    
    def close(self) -> None:
        """Close the database connection."""
        try:
            if self.conn:
                self.conn.close()
                self.conn = None
        except Exception as e:
            logger.error(f"Error closing database connection: {e}", exc_info=True)
    
    def __del__(self):
        """Ensure the database connection is closed when the object is destroyed."""
        self.close()
