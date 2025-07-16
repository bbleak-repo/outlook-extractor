"""JSON export functionality for Outlook Extractor.

This module provides functionality to export email data to JSON format with
support for pretty-printing and custom serialization.
"""
import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from threading import Event
from typing import Any, Dict, List, Optional, Union, Callable, Tuple

from .. import constants

logger = logging.getLogger(__name__)

class JSONExporter:
    """Handles the export of email data to JSON format."""
    
    def __init__(self, config: Optional[dict] = None):
        """Initialize the JSON exporter.
        
        Args:
            config: Optional configuration dictionary
        """
        self.config = config or {}
        self.fields = constants.EXPORT_FIELDS_V1
        
    def export_emails(
        self,
        emails: List[Dict[str, Any]],
        output_path: Union[str, Path],
        pretty_print: bool = True,
        include_metadata: bool = True,
        progress_callback: Optional[Callable[[int, int], None]] = None,
        cancel_event: Optional[Event] = None
    ) -> Tuple[bool, str]:
        """Export emails to a JSON file.
        
        Args:
            emails: List of email dictionaries to export
            output_path: Path to the output JSON file
            pretty_print: Whether to format the JSON with indentation
            include_metadata: Whether to include export metadata
            progress_callback: Optional callback for progress updates
            cancel_event: Optional event to cancel the export
            
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
        temp_path = output_path.with_suffix('.tmp' + output_path.suffix)
        total_emails = len(emails)
        
        try:
            # Prepare export data
            export_data = {
                'metadata': {
                    'export_date': datetime.now(timezone.utc).isoformat(),
                    'format': 'json',
                    'version': '1.0',
                    'record_count': total_emails,
                    'fields': self.fields
                } if include_metadata else None,
                'emails': []
            }
            
            # Process emails in batches
            batch_size = 100
            for i in range(0, total_emails, batch_size):
                if cancel_event and cancel_event.is_set():
                    return False, "Export cancelled by user"
                
                batch = emails[i:i + batch_size]
                for email in batch:
                    try:
                        # Convert any non-serializable fields
                        processed_email = self._process_email(email)
                        export_data['emails'].append(processed_email)
                        
                        # Update progress
                        if progress_callback and (len(export_data['emails']) % 10 == 0 or 
                                               len(export_data['emails']) == total_emails):
                            progress_callback(len(export_data['emails']), total_emails)
                            
                    except Exception as e:
                        logger.error(f"Error processing email: {e}", exc_info=True)
                        continue
            
            # Write to temporary file first
            with open(temp_path, 'w', encoding='utf-8') as f:
                json.dump(
                    export_data,
                    f,
                    indent=2 if pretty_print else None,
                    ensure_ascii=False,
                    default=self._json_serializer
                )
            
            # Rename temp file to final name (atomic operation)
            if temp_path.exists():
                if output_path.exists():
                    output_path.unlink()
                temp_path.rename(output_path)
            
            msg = f"Successfully exported {len(emails)} emails to {output_path}"
            logger.info(msg)
            return True, msg
            
        except Exception as e:
            error_msg = f"Error exporting to JSON: {str(e)}"
            logger.error(error_msg, exc_info=True)
            
            # Clean up temp file if it exists
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception as cleanup_error:
                    logger.error(f"Failed to clean up temp file: {cleanup_error}")
            
            return False, error_msg
    
    def _process_email(self, email: Dict[str, Any]) -> Dict[str, Any]:
        """Process an email dictionary for JSON serialization."""
        processed = {}
        
        for key, value in email.items():
            try:
                # Skip None values to reduce file size
                if value is not None:
                    # Convert any non-serializable values
                    if isinstance(value, (datetime, bytes)):
                        processed[key] = str(value)
                    elif hasattr(value, '__dict__'):
                        processed[key] = str(value)
                    else:
                        processed[key] = value
            except Exception as e:
                logger.warning(f"Error processing field '{key}': {e}")
                processed[key] = "[Error processing value]"
        
        return processed
    
    @staticmethod
    def _json_serializer(obj: Any) -> Any:
        """Custom JSON serializer for non-serializable objects."""
        if isinstance(obj, (datetime, bytes)):
            return str(obj)
        elif hasattr(obj, '__dict__'):
            return str(obj)
        raise TypeError(f"Object of type {type(obj)} is not JSON serializable")
    
    def export_to_string(
        self,
        emails: List[Dict[str, Any]],
        pretty_print: bool = True,
        include_metadata: bool = True
    ) -> str:
        """Export emails to a JSON string.
        
        Args:
            emails: List of email dictionaries to export
            pretty_print: Whether to format the JSON with indentation
            include_metadata: Whether to include export metadata
            
        Returns:
            JSON string containing the exported data
        """
        export_data = {
            'metadata': {
                'export_date': datetime.now(timezone.utc).isoformat(),
                'format': 'json',
                'version': '1.0',
                'record_count': len(emails),
                'fields': self.fields
            } if include_metadata else None,
            'emails': [self._process_email(email) for email in emails]
        }
        
        return json.dumps(
            export_data,
            indent=2 if pretty_print else None,
            ensure_ascii=False,
            default=self._json_serializer
        )
