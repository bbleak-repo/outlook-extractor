#!/usr/bin/env python3
"""Command-line interface for Outlook Email Extractor.

This module provides a command-line interface for extracting emails from Outlook
and exporting them to various formats.
"""
import argparse
import json
import logging
import os
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional

from outlook_extractor import __version__
from outlook_extractor.core.outlook_client import OutlookClient
from outlook_extractor.processors.email_processor import EmailProcessor
from outlook_extractor.export.csv_exporter import CSVExporter
from outlook_extractor.export.constants import EXPORT_FIELDS_V1

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class CLIApp:
    """Command-line application for Outlook Email Extractor."""
    
    def __init__(self):
        """Initialize the CLI application."""
        self.parser = self._create_parser()
        self.args = None
        self.config = {}
        
    def _create_parser(self) -> argparse.ArgumentParser:
        """Create and configure the argument parser."""
        parser = argparse.ArgumentParser(
            description='Extract emails from Outlook and export to various formats.'
        )
        
        parser.add_argument(
            '-v', '--version',
            action='version',
            version=f'%(prog)s {__version__}'
        )
        
        # Input options
        input_group = parser.add_argument_group('Input Options')
        input_group.add_argument(
            '-f', '--folder',
            type=str,
            default='Inbox',
            help='Source folder to extract emails from (default: Inbox)'
        )
        input_group.add_argument(
            '-r', '--recursive',
            action='store_true',
            help='Process subfolders recursively'
        )
        input_group.add_argument(
            '--include-read',
            action='store_true',
            help='Include emails that have been read'
        )
        
        # Filter options
        filter_group = parser.add_argument_group('Filter Options')
        filter_group.add_argument(
            '--from',
            dest='sender',
            type=str,
            help='Filter by sender email address'
        )
        filter_group.add_argument(
            '--subject',
            type=str,
            help='Filter by subject (case-insensitive substring match)'
        )
        filter_group.add_argument(
            '--after',
            type=str,
            help='Filter emails received after this date (YYYY-MM-DD)'
        )
        filter_group.add_argument(
            '--before',
            type=str,
            help='Filter emails received before this date (YYYY-MM-DD)'
        )
        
        # Output options
        output_group = parser.add_argument_group('Output Options')
        output_group.add_argument(
            '-o', '--output',
            type=str,
            default='emails.csv',
            help='Output file path (default: emails.csv)'
        )
        output_group.add_argument(
            '--format',
            choices=['csv', 'json'],
            default='csv',
            help='Output format (default: csv)'
        )
        output_group.add_argument(
            '--no-headers',
            action='store_false',
            dest='include_headers',
            help='Do not include headers in the output file'
        )
        
        # Configuration options
        config_group = parser.add_argument_group('Configuration')
        config_group.add_argument(
            '--config',
            type=str,
            help='Path to configuration file (YAML or JSON)'
        )
        config_group.add_argument(
            '--priority-emails',
            type=str,
            nargs='+',
            help='List of email addresses to mark as priority'
        )
        config_group.add_argument(
            '--admin-emails',
            type=str,
            nargs='+',
            help='List of email addresses to mark as admin'
        )
        
        # Debug options
        debug_group = parser.add_argument_group('Debugging')
        debug_group.add_argument(
            '--debug',
            action='store_true',
            help='Enable debug logging'
        )
        debug_group.add_argument(
            '--dry-run',
            action='store_true',
            help='Process emails but do not export'
        )
        
        return parser
    
    def _load_config(self) -> None:
        """Load configuration from file and command line arguments."""
        # Load config from file if specified
        if self.args.config:
            try:
                with open(self.args.config, 'r') as f:
                    if self.args.config.lower().endswith('.json'):
                        import json
                        self.config = json.load(f)
                    else:
                        import yaml
                        self.config = yaml.safe_load(f)
                logger.info(f"Loaded configuration from {self.args.config}")
            except Exception as e:
                logger.warning(f"Failed to load config file: {e}")
        
        # Override with command line arguments
        if self.args.priority_emails:
            self.config['priority_addresses'] = self.args.priority_emails
        if self.args.admin_emails:
            self.config['admin_addresses'] = self.args.admin_emails
    
    def _process_emails(self) -> List[Dict[str, Any]]:
        """Process emails based on the provided configuration."""
        logger.info("Connecting to Outlook...")
        
        try:
            with OutlookClient() as outlook:
                # Get the source folder
                folder = outlook.get_folder(self.args.folder)
                if not folder:
                    logger.error(f"Folder '{self.args.folder}' not found")
                    return []
                
                # Get all messages in the folder
                messages = outlook.get_messages(
                    folder,
                    recursive=self.args.recursive,
                    include_read=self.args.include_read
                )
                
                if not messages:
                    logger.warning("No messages found")
                    return []
                
                # Process messages
                logger.info(f"Processing {len(messages)} messages...")
                processor = EmailProcessor(self.config)
                processed_emails = []
                
                for i, msg in enumerate(messages, 1):
                    try:
                        email_data = processor.process_message(msg)
                        processed_emails.append(email_data)
                        
                        if i % 10 == 0 or i == len(messages):
                            logger.info(f"Processed {i}/{len(messages)} messages...")
                            
                    except Exception as e:
                        logger.error(f"Error processing message {i}: {e}", exc_info=self.args.debug)
                
                return processed_emails
                
        except Exception as e:
            logger.error(f"Error accessing Outlook: {e}", exc_info=self.args.debug)
            return []
    
    def _export_emails(self, emails: List[Dict[str, Any]]) -> bool:
        """Export processed emails to the specified format."""
        if not emails:
            logger.warning("No emails to export")
            return False
        
        output_path = Path(self.args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            if self.args.format == 'csv':
                exporter = CSVExporter()
                success = exporter.export_emails(
                    emails=emails,
                    output_path=output_path,
                    include_headers=self.args.include_headers
                )
            elif self.args.format == 'json':
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(emails, f, indent=2, default=str)
                success = True
            else:
                logger.error(f"Unsupported export format: {self.args.format}")
                return False
            
            if success:
                logger.info(f"Successfully exported {len(emails)} emails to {output_path}")
                return True
            else:
                logger.error("Failed to export emails")
                return False
                
        except Exception as e:
            logger.error(f"Error exporting emails: {e}", exc_info=self.args.debug)
            return False
    
    def run(self) -> int:
        """Run the CLI application."""
        # Parse command line arguments
        self.args = self.parser.parse_args()
        
        # Configure logging level
        if self.args.debug:
            logging.getLogger().setLevel(logging.DEBUG)
        
        logger.debug(f"Command line arguments: {self.args}")
        
        # Load configuration
        self._load_config()
        
        # Process emails
        emails = self._process_emails()
        
        # Apply filters
        if self.args.sender:
            emails = [e for e in emails if self.args.sender.lower() in e.get('sender_email', '').lower()]
        
        if self.args.subject:
            subject_lower = self.args.subject.lower()
            emails = [e for e in emails if subject_lower in e.get('subject', '').lower()]
        
        # Export emails if not in dry-run mode
        if not self.args.dry_run and emails:
            success = self._export_emails(emails)
            return 0 if success else 1
        
        return 0

def main() -> None:
    """Entry point for the CLI application."""
    try:
        app = CLIApp()
        sys.exit(app.run())
    except KeyboardInterrupt:
        logger.info("Operation cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"An unexpected error occurred: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
