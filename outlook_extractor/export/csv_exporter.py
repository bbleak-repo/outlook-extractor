"""CSV export functionality for Outlook Extractor.

This module provides functionality to export email data to CSV format
with proper field ordering and formatting according to the v10 specification.
Supports batch processing, progress updates, and cancellation.
"""
import csv
import json
import logging
from datetime import datetime, timezone
from pathlib import Path
from threading import Event
from typing import Any, Dict, List, Optional, Tuple, Union, Callable
from queue import Queue, Empty
import pandas as pd

from .. import constants

logger = logging.getLogger(__name__)

class CSVExporter:
    """Handles the export of email data to CSV format.
    
    This exporter ensures compatibility with the v10 specification while
    providing robust error handling and performance optimizations.
    """
    
    def __init__(self, config: Optional[dict] = None):
        """Initialize the CSV exporter.
        
        Args:
            config: Optional configuration dictionary
        """
        self.config = config or {}
        self.fields = constants.EXPORT_FIELDS_V1
        self._setup_field_formatters()
    
    def _setup_field_formatters(self) -> None:
        """Set up field formatters for CSV export."""
        self._field_formatters = {}
        
        # Define default formatters for different field types
        type_formatters = {
            'string': str,
            'number': str,
            'boolean': lambda x: str(x).lower(),
            'datetime': lambda x: x.isoformat() if hasattr(x, 'isoformat') else str(x) if x else '',
            'list': lambda x: ';'.join(str(i) for i in x) if isinstance(x, list) else str(x or ''),
            'dict': lambda x: json.dumps(x) if isinstance(x, dict) else str(x or ''),
            'text': str
        }
        
        # Set up formatters for each field
        for field in constants.EXPORT_FIELDS_V1:
            field_id = field['id']
            field_type = field.get('type', 'string')
            self._field_formatters[field_id] = type_formatters.get(field_type, str)
            
        # Set default fields if not already set
        self.fields = [field['id'] for field in constants.EXPORT_FIELDS_V1]
    
    def export_emails(
        self, 
        emails: List[Dict[str, Any]], 
        output_path: Union[str, Path],
        include_headers: bool = True,
        encoding: str = 'utf-8-sig',  # Use BOM for Excel compatibility
        batch_size: int = 1000,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[Event] = None
    ) -> Tuple[bool, str]:
        """Export emails to a CSV file with batch processing and progress updates.
        
        Args:
            emails: List of email dictionaries to export
            output_path: Path to the output CSV file
            include_headers: Whether to include headers in the CSV
            encoding: File encoding to use
            batch_size: Number of emails to process in each batch
            progress_callback: Optional callback function(processed: int, total: int)
            cancel_event: Optional threading.Event to cancel the export
            
        Returns:
            Tuple[bool, str]: (success, message) - success status and message
        """
        if not emails:
            msg = "No emails to export"
            logger.warning(msg)
            return False, msg
            
        if cancel_event and cancel_event.is_set():
            msg = "Export cancelled by user"
            logger.info(msg)
            return False, msg
            
        output_path = Path(output_path)
        total_emails = len(emails)
        processed = 0
        temp_path = output_path.with_suffix('.tmp' + output_path.suffix)
        
        try:
            # Ensure output directory exists
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with temp_path.open('w', newline='', encoding=encoding) as f:
                writer = csv.DictWriter(f, fieldnames=self.fields)
                
                if include_headers:
                    # Use the 'name' from EXPORT_FIELDS_V1 for headers if available
                    headers = {}
                    for field in constants.EXPORT_FIELDS_V1:
                        headers[field['id']] = field.get('name', field['id'])
                    writer.writerow(headers)
                
                # Process emails in batches
                for i in range(0, total_emails, batch_size):
                    if cancel_event and cancel_event.is_set():
                        msg = "Export cancelled by user"
                        logger.info(msg)
                        return False, msg
                        
                    batch = emails[i:i + batch_size]
                    for email in batch:
                        try:
                            row = self._format_email_row(email)
                            writer.writerow(row)
                            processed += 1
                            
                            # Update progress every 10 emails or on last email
                            if progress_callback and (processed % 10 == 0 or processed == total_emails):
                                progress_callback(processed, total_emails)
                                
                        except Exception as e:
                            logger.error(f"Error processing email: {e}", exc_info=True)
                            continue
            
            # Rename temp file to final name (atomic operation on most filesystems)
            if temp_path.exists():
                if output_path.exists():
                    output_path.unlink()  # Remove existing file if it exists
                temp_path.rename(output_path)
            
            msg = f"Successfully exported {processed} of {total_emails} emails to {output_path}"
            logger.info(msg)
            return True, msg
            
        except Exception as e:
            error_msg = f"Error exporting emails: {str(e)}"
            logger.error(error_msg, exc_info=True)
            
            # Clean up temp file if it exists
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception as cleanup_error:
                    logger.error(f"Failed to clean up temp file: {cleanup_error}")
            
            return False, error_msg
    
    def _format_email_row(self, email: Dict[str, Any]) -> Dict[str, str]:
        """Format an email for CSV export.
        
        Args:
            email: Email data to format
            
        Returns:
            Formatted email data
        """
        formatted = {}
        for field in self.fields:
            formatter = self._field_formatters.get(field, str)
            try:
                # Handle nested fields (e.g., 'sender.email_address')
                if '.' in field:
                    parts = field.split('.')
                    value = email
                    for part in parts:
                        if isinstance(value, dict):
                            value = value.get(part, '')
                        else:
                            value = ''
                            break
                else:
                    value = email.get(field, '')
                
                # Apply formatter if value is not None
                if value is not None:
                    formatted[field] = formatter(value)
                else:
                    formatted[field] = ''
                    
            except Exception as e:
                logger.warning(f"Error formatting field '{field}': {e}")
                formatted[field] = ''
                
        return formatted
    
    def export_to_file(
        self, 
        emails: List[Dict[str, Any]], 
        output_path: Union[str, Path],
        **kwargs
    ) -> bool:
        """Alias for export_emails for backward compatibility.
        
        Args:
            emails: List of email dictionaries to export
            output_path: Path to the output CSV file
            **kwargs: Additional arguments to pass to export_emails
            
        Returns:
            bool: True if export was successful, False otherwise
        """
        return self.export_emails(emails, output_path, **kwargs)

    def export_subject_analysis(
        self,
        emails: List[Dict],
        output_path: str
    ) -> str:
        """Generate subject analysis report."""
        try:
            # Count subjects by folder
            subject_counts = {}
            for email_data in emails:
                folder = email_data.get('parent_folder', 'Unknown')
                subject = email_data.get('subject', '(No Subject)')
                
                if folder not in subject_counts:
                    subject_counts[folder] = {}
                subject_counts[folder][subject] = subject_counts[folder].get(subject, 0) + 1
            
            # Convert to DataFrame for easy CSV export
            rows = []
            for folder, subjects in subject_counts.items():
                for subject, count in subjects.items():
                    rows.append({
                        'folder': folder,
                        'subject': subject,
                        'count': count
                    })
            
            # Sort by count descending
            rows.sort(key=lambda x: x['count'], reverse=True)
            
            # Write to CSV
            df = pd.DataFrame(rows)
            df.to_csv(output_path, index=False)
            
            logger.info(f"Subject analysis exported to {output_path}")
            return str(output_path)
            
        except Exception as e:
            logger.error(f"Error generating subject analysis: {str(e)}")
            raise