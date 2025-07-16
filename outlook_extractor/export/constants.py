"""Constants for export functionality.

This module defines standard field names and formats used for exporting
email data to various formats.
"""
from typing import List, Dict, Any, Optional

# Maintain exact field order from v10 for compatibility
EXPORT_FIELDS_V1 = [
    "entry_id", "conversation_id", "subject", "sender_name",
    "sender_email", "to_recipients", "cc_recipients",
    "bcc_recipients", "received_time", "sent_time",
    "categories", "importance", "sensitivity",
    "has_attachments", "is_read", "is_flagged",
    "is_priority", "is_admin", "body", "html_body",
    "folder_path", "thread_id", "thread_depth"
]

# Field type information for validation and formatting
FIELD_TYPES = {
    "entry_id": str,
    "conversation_id": str,
    "subject": str,
    "sender_name": str,
    "sender_email": str,
    "to_recipients": str,
    "cc_recipients": str,
    "bcc_recipients": str,
    "received_time": "datetime",
    "sent_time": "datetime",
    "categories": str,
    "importance": int,  # 0=Low, 1=Normal, 2=High
    "sensitivity": int,  # 0=Normal, 1=Personal, 2=Private, 3=Confidential
    "has_attachments": bool,
    "is_read": bool,
    "is_flagged": bool,
    "is_priority": bool,
    "is_admin": bool,
    "body": str,
    "html_body": str,
    "folder_path": str,
    "thread_id": str,
    "thread_depth": int
}

# Default datetime format for string representation
DEFAULT_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S %Z"

def get_field_formatter(field_name: str):
    """Get a formatter function for a specific field.
    
    Args:
        field_name: Name of the field to format
        
    Returns:
        A function that formats the field value appropriately
    """
    field_type = FIELD_TYPES.get(field_name, str)
    
    if field_type == "datetime":
        def format_datetime(value):
            if value is None:
                return ""
            if hasattr(value, 'strftime'):
                return value.strftime(DEFAULT_DATETIME_FORMAT)
            return str(value)
        return format_datetime
    
    elif field_type == bool:
        return lambda x: "Yes" if x else "No"
        
    # Default formatter (just convert to string)
    return str
