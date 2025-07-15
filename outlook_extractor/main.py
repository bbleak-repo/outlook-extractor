"""
Main entry point for the Outlook Extractor application.

This module provides the main entry point and application logic for the
Outlook Email Extraction Tool. It orchestrates the interaction between
the Outlook client, storage backends, and user interface.
"""
import os
import sys
import logging
import json
import csv
from typing import List, Dict, Any, Optional, Set, Tuple
from datetime import datetime, timedelta
from pathlib import Path

# Import local modules
from .config import ConfigManager, get_config, load_config
from .logging_setup import setup_logging, get_logger, logger
from .core.outlook_client import OutlookClient
from .core.email_threading import EmailThread, ThreadManager
from .storage import SQLiteStorage, JSONStorage, EmailStorage
from .export.csv_exporter import CSVExporter
from .ui import EmailExtractorUI, ExportTab

# Set up logging
setup_logging()
logger = get_logger(__name__)

class OutlookExtractor:
    """Main application class for the Outlook Email Extraction Tool."""
    
    def __init__(self, config_path: str = None):
        """Initialize the Outlook Extractor application.
        
        Args:
            config_path: Optional path to a configuration file.
        """
        # Load configuration
        self.config = load_config(config_path) if config_path else get_config()
        
        # Initialize components
        self.outlook_client = None
        self.storage = None
        self.thread_manager = ThreadManager()
        self.csv_exporter = CSVExporter(self.config)
        self.ui = EmailExtractorUI(self)
        
        # Initialize storage
        self._init_storage()
    
    def _init_storage(self) -> None:
        """Initialize the storage backend based on configuration."""
        storage_type = self.config.get('storage', 'type', 'sqlite').lower()
        
        if storage_type == 'sqlite':
            db_path = self.config.get('storage', 'db_path', 'emails.db')
            self.storage = SQLiteStorage(db_path=db_path, config=self.config)
        elif storage_type == 'json':
            json_path = self.config.get('storage', 'json_path', 'emails.json')
            self.storage = JSONStorage(json_path=json_path, config=self.config)
        else:
            raise ValueError(f"Unsupported storage type: {storage_type}")
        
        logger.info(f"Initialized {storage_type.upper()} storage at {getattr(self.storage, 'db_path', getattr(self.storage, 'json_path', 'unknown'))}")
    
    def connect_to_outlook(self) -> bool:
        """Connect to Microsoft Outlook.
        
        Returns:
            bool: True if the connection was successful, False otherwise.
        """
        try:
            self.outlook_client = OutlookClient(config=self.config)
            return self.outlook_client.connect()
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}", exc_info=True)
            return False
    
    def extract_emails(self, folder_paths: List[str] = None, 
                      start_date: datetime = None, 
                      end_date: datetime = None) -> Dict[str, Any]:
        """Extract emails from the specified folders.
        
        Args:
            folder_paths: List of folder paths to extract emails from.
                         If None, uses folders from configuration.
            start_date: Optional start date for filtering emails.
            end_date: Optional end date for filtering emails.
            
        Returns:
            Dict containing extraction results and statistics.
        """
        if not self.outlook_client or not self.outlook_client.is_connected():
            if not self.connect_to_outlook():
                return {
                    'success': False,
                    'error': 'Could not connect to Outlook',
                    'emails_processed': 0
                }
        
        # Use configured folders if none provided
        if not folder_paths:
            folder_patterns = self.config.get_list('outlook', 'folder_patterns', ['Inbox'])
            folder_paths = self.outlook_client.find_matching_folders(folder_patterns)
        
        # Set default date range if not specified
        if not start_date or not end_date:
            days_back = self.config.get_int('date_range', 'days_back', 30)
            end_date = end_date or datetime.now()
            start_date = start_date or (end_date - timedelta(days=days_back))
        
        logger.info(f"Extracting emails from {len(folder_paths)} folders between {start_date} and {end_date}")
        
        # Extract emails
        emails = []
        for folder_path in folder_paths:
            try:
                folder_emails = self.outlook_client.extract_emails(
                    folder_path=folder_path,
                    start_date=start_date,
                    end_date=end_date,
                    include_attachments=self.config.get_boolean('email_processing', 'extract_attachments', False),
                    include_embedded_images=self.config.get_boolean('email_processing', 'extract_embedded_images', False)
                )
                emails.extend(folder_emails)
                logger.info(f"Extracted {len(folder_emails)} emails from {folder_path}")
            except Exception as e:
                logger.error(f"Error extracting emails from {folder_path}: {e}", exc_info=True)
        
        # Process emails (threading, etc.)
        processed_emails = self._process_emails(emails)
        
        # Save emails to storage
        saved_count = self.storage.save_emails(processed_emails)
        
        return {
            'success': True,
            'emails_processed': len(emails),
            'emails_saved': saved_count,
            'folders_processed': len(folder_paths),
            'start_date': start_date,
            'end_date': end_date
        }
    
    def _process_emails(self, emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Process extracted emails (threading, etc.).
        
        Args:
            emails: List of email dictionaries to process.
            
        Returns:
            List of processed email dictionaries.
        """
        processed_emails = []
        
        for email in emails:
            try:
                # Add to thread manager
                self.thread_manager.add_email(email)
                
                # Add any additional processing here
                processed_emails.append(email)
                
            except Exception as e:
                logger.error(f"Error processing email {email.get('id', 'unknown')}: {e}", exc_info=True)
        
        return processed_emails
    
    def search_emails(self, query: str, limit: int = 100) -> List[Dict[str, Any]]:
        """Search for emails matching the query.
        
        Args:
            query: Search query string.
            limit: Maximum number of results to return.
            
        Returns:
            List of matching email dictionaries.
        """
        return self.storage.search_emails(query, limit=limit)
    
    def get_email(self, email_id: str) -> Optional[Dict[str, Any]]:
        """Get a single email by ID.
        
        Args:
            email_id: ID of the email to retrieve.
            
        Returns:
            Email dictionary if found, None otherwise.
        """
        return self.storage.get_email(email_id)
    
    def export_emails(self, emails: List[Dict[str, Any]], format: str = 'all', 
                     export_settings: Optional[Dict[str, Any]] = None) -> Tuple[bool, List[str]]:
        """Export extracted emails to the specified format.
        
        Args:
            emails: List of email dictionaries to export.
            format: Export format ('json', 'csv', or 'all').
            export_settings: Dictionary of export settings from the UI.
            
        Returns:
            Tuple[bool, List[str]]: (success, output_files) - success status and list of exported files.
        """
        if not export_settings:
            export_settings = self.config.get('export', {})
            
        output_files = []
        success = True
        
        try:
            output_dir = Path(export_settings.get('output_dir', self.config.get('export', 'output_dir', '.')))
            output_dir.mkdir(parents=True, exist_ok=True)
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            file_prefix = export_settings.get('file_prefix', 'emails_')
            
            # Export to JSON if requested or if format is 'all'
            if format.lower() in ('json', 'all'):
                output_file = output_dir / f'{file_prefix}{timestamp}.json'
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(emails, f, indent=2, default=str)
                logger.info(f'Exported {len(emails)} emails to {output_file}')
                output_files.append(str(output_file))
            
            # Export to CSV if requested or if format is 'all'
            if format.lower() in ('csv', 'all'):
                if not emails:
                    logger.warning('No emails to export')
                    return False, []
                
                # Use CSV exporter for advanced export options
                if export_settings.get('enable_csv', True):
                    # Basic email data export
                    if export_settings.get('export_basic', True):
                        csv_file = output_dir / f'{file_prefix}{timestamp}.csv'
                        self.csv_exporter.export_emails_to_csv(emails, str(csv_file))
                        output_files.append(str(csv_file))
                    
                    # Subject analysis export
                    if export_settings.get('export_analysis', True):
                        analysis_file = output_dir / f'subject_analysis_{timestamp}.csv'
                        self.csv_exporter.export_subject_analysis(emails, str(analysis_file))
                        output_files.append(str(analysis_file))
            
            return True, output_files
            
        except Exception as e:
            logger.error(f'Error exporting emails: {str(e)}')
            return False, output_files
            
    def get_thread(self, thread_id: str) -> Optional[Dict[str, Any]]:
        """Get a thread and its emails by thread ID.
        
        Args:
            thread_id: ID of the thread to retrieve.
            
        Returns:
            Dictionary containing thread information and emails if found, None otherwise.
        """
        # This is a simplified implementation - in a real app, you'd want to
        # store and retrieve thread information from storage
        thread_emails = self.storage.search_emails(f'thread_id:{thread_id}')
        
        if not thread_emails:
            return None
        
        # Sort emails by date
        thread_emails.sort(key=lambda x: x.get('sent_date') or x.get('received_date', ''))
        
        # Get thread info from the first email
        first_email = thread_emails[0]
        
        return {
            'id': thread_id,
            'subject': first_email.get('subject', '(No Subject)'),
            'start_date': first_email.get('sent_date') or first_email.get('received_date'),
            'end_date': thread_emails[-1].get('sent_date') or thread_emails[-1].get('received_date'),
            'email_count': len(thread_emails),
            'participants': self._get_thread_participants(thread_emails),
            'emails': thread_emails
        }
    
    def _get_thread_participants(self, emails: List[Dict[str, Any]]) -> Set[str]:
        """Get all unique participants in a thread.
        
        Args:
            emails: List of email dictionaries in the thread.
            
        Returns:
            Set of participant email addresses.
        """
        participants = set()
        
        for email in emails:
            # Add sender
            if 'sender' in email and email['sender']:
                participants.add(email['sender'])
            
            # Add recipients
            for field in ['recipients', 'cc_recipients', 'bcc_recipients']:
                if field in email and email[field]:
                    if isinstance(email[field], str):
                        participants.add(email[field])
                    elif isinstance(email[field], list):
                        participants.update(addr for addr in email[field] if addr)
        
        return participants
    
    def run(self) -> int:
        """Run the application.
        
        Returns:
            int: Exit code (0 for success, non-zero for error).
        """
        try:
            # Start the UI
            return self.ui.run()
        except Exception as e:
            logger.critical(f"Fatal error: {e}", exc_info=True)
            return 1
        finally:
            # Clean up resources
            self.close()
    
    def close(self) -> None:
        """Close resources and clean up."""
        if self.storage:
            try:
                self.storage.close()
            except Exception as e:
                logger.error(f"Error closing storage: {e}", exc_info=True)
        
        if self.outlook_client:
            try:
                self.outlook_client.disconnect()
            except Exception as e:
                logger.error(f"Error disconnecting from Outlook: {e}", exc_info=True)


def main() -> int:
    """Main entry point for the command-line interface."""
    import argparse
    
    # Parse command-line arguments
    parser = argparse.ArgumentParser(description='Outlook Email Extraction Tool')
    parser.add_argument('--config', '-c', help='Path to configuration file')
    parser.add_argument('--extract', '-e', action='store_true', help='Run extraction immediately')
    parser.add_argument('--folders', nargs='+', help='Folders to extract from (overrides config)')
    parser.add_argument('--start-date', help='Start date (YYYY-MM-DD)')
    parser.add_argument('--end-date', help='End date (YYYY-MM-DD)')
    parser.add_argument('--search', '-s', help='Search query')
    parser.add_argument('--output', '-o', help='Output format (json, csv, text)', default='text')
    parser.add_argument('--limit', '-l', type=int, default=100, help='Maximum number of results to return')
    parser.add_argument('--version', '-v', action='store_true', help='Show version and exit')
    
    args = parser.parse_args()
    
    # Show version and exit
    if args.version:
        from . import __version__
        print(f"Outlook Extractor v{__version__}")
        return 0
    
    # Initialize the application
    try:
        app = OutlookExtractor(config_path=args.config)
    except Exception as e:
        logger.critical(f"Failed to initialize application: {e}", exc_info=True)
        return 1
    
    try:
        # Handle command-line actions
        if args.search:
            # Search mode
            results = app.search_emails(args.search, limit=args.limit)
            
            # Output results
            if args.output.lower() == 'json':
                import json
                print(json.dumps(results, indent=2, default=str))
            elif args.output.lower() == 'csv':
                import csv
                import sys
                
                if not results:
                    print("No results found")
                    return 0
                
                # Get all possible field names
                fieldnames = set()
                for result in results:
                    fieldnames.update(result.keys())
                
                writer = csv.DictWriter(sys.stdout, fieldnames=sorted(fieldnames))
                writer.writeheader()
                writer.writerows(results)
            else:
                # Text output (default)
                for i, result in enumerate(results, 1):
                    print(f"{i}. {result.get('subject', '(No Subject)')}")
                    print(f"   From: {result.get('sender', 'Unknown')}")
                    print(f"   Date: {result.get('sent_date', result.get('received_date', 'Unknown'))}")
                    print(f"   ID: {result.get('id', 'N/A')}")
                    print()
            
            return 0
        
        elif args.extract:
            # Extract mode
            start_date = datetime.strptime(args.start_date, '%Y-%m-%d') if args.start_date else None
            end_date = datetime.strptime(args.end_date, '%Y-%m-%d') if args.end_date else None
            
            result = app.extract_emails(
                folder_paths=args.folders,
                start_date=start_date,
                end_date=end_date
            )
            
            if result['success']:
                print(f"Successfully processed {result['emails_processed']} emails from {result['folders_processed']} folders")
                print(f"Saved {result['emails_saved']} emails to storage")
                return 0
            else:
                print(f"Error: {result.get('error', 'Unknown error')}", file=sys.stderr)
                return 1
        
        else:
            # Interactive mode
            return app.run()
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        return 130  # SIGINT exit code
    except Exception as e:
        logger.critical(f"Fatal error: {e}", exc_info=True)
        return 1
    finally:
        app.close()


if __name__ == "__main__":
    sys.exit(main())
