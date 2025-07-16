"""Constants used throughout the Outlook Extractor application."""
from typing import List, Dict, Any, Tuple

# Version information
VERSION: str = '1.0.0'

# Default export fields for different versions
EXPORT_FIELDS_V1: List[Dict[str, Any]] = [
    {
        'id': 'id',
        'name': 'ID',
        'description': 'Unique identifier for the email',
        'type': 'string',
        'default_visible': False
    },
    {
        'id': 'conversation_id',
        'name': 'Conversation ID',
        'description': 'ID of the conversation this email belongs to',
        'type': 'string',
        'default_visible': False
    },
    {
        'id': 'subject',
        'name': 'Subject',
        'description': 'Email subject',
        'type': 'string',
        'default_visible': True
    },
    {
        'id': 'sender',
        'name': 'From',
        'description': 'Email sender',
        'type': 'string',
        'default_visible': True
    },
    {
        'id': 'to_recipients',
        'name': 'To',
        'description': 'Email recipients (To field)',
        'type': 'list',
        'default_visible': True
    },
    {
        'id': 'cc_recipients',
        'name': 'CC',
        'description': 'Carbon copy recipients',
        'type': 'list',
        'default_visible': False
    },
    {
        'id': 'bcc_recipients',
        'name': 'BCC',
        'description': 'Blind carbon copy recipients',
        'type': 'list',
        'default_visible': False
    },
    {
        'id': 'received_date',
        'name': 'Received',
        'description': 'Date and time when the email was received',
        'type': 'datetime',
        'default_visible': True
    },
    {
        'id': 'sent_date',
        'name': 'Sent',
        'description': 'Date and time when the email was sent',
        'type': 'datetime',
        'default_visible': True
    },
    {
        'id': 'importance',
        'name': 'Importance',
        'description': 'Email importance level (Low, Normal, High)',
        'type': 'string',
        'default_visible': False
    },
    {
        'id': 'is_read',
        'name': 'Is Read',
        'description': 'Whether the email has been read',
        'type': 'boolean',
        'default_visible': False
    },
    {
        'id': 'has_attachments',
        'name': 'Has Attachments',
        'description': 'Whether the email has attachments',
        'type': 'boolean',
        'default_visible': True
    },
    {
        'id': 'body',
        'name': 'Body',
        'description': 'Email body content',
        'type': 'text',
        'default_visible': True
    },
    {
        'id': 'categories',
        'name': 'Categories',
        'description': 'Email categories',
        'type': 'list',
        'default_visible': False
    },
    {
        'id': 'size',
        'name': 'Size (bytes)',
        'description': 'Size of the email in bytes',
        'type': 'number',
        'default_visible': False
    },
    {
        'id': 'is_draft',
        'name': 'Is Draft',
        'description': 'Whether the email is a draft',
        'type': 'boolean',
        'default_visible': False
    },
    {
        'id': 'is_encrypted',
        'name': 'Is Encrypted',
        'description': 'Whether the email is encrypted',
        'type': 'boolean',
        'default_visible': False
    },
    {
        'id': 'is_signed',
        'name': 'Is Signed',
        'description': 'Whether the email is digitally signed',
        'type': 'boolean',
        'default_visible': False
    },
    {
        'id': 'internet_message_headers',
        'name': 'Internet Headers',
        'description': 'Email internet headers',
        'type': 'dict',
        'default_visible': False
    },
    {
        'id': 'attachments',
        'name': 'Attachments',
        'description': 'Email attachments',
        'type': 'list',
        'default_visible': False
    }
]

# Default export fields for quick selection
DEFAULT_EXPORT_FIELDS = [
    'subject', 'sender', 'to_recipients', 'received_date', 'body'
]

# Field types that should be serialized as JSON
JSON_SERIALIZABLE_TYPES = ['list', 'dict']

# Default date format for display
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# Default timezone for date handling
DEFAULT_TIMEZONE = 'UTC'

# Default encoding for file operations
DEFAULT_ENCODING = 'utf-8'

# Supported export formats
EXPORT_FORMATS = [
    {
        'id': 'csv',
        'name': 'CSV',
        'extension': '.csv',
        'description': 'Comma-separated values',
        'mime_type': 'text/csv'
    },
    {
        'id': 'excel',
        'name': 'Excel',
        'extension': '.xlsx',
        'description': 'Microsoft Excel format',
        'mime_type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    },
    {
        'id': 'json',
        'name': 'JSON',
        'extension': '.json',
        'description': 'JavaScript Object Notation',
        'mime_type': 'application/json'
    }
]

# Default export settings
DEFAULT_EXPORT_SETTINGS = {
    'format': 'csv',
    'include_headers': True,
    'encoding': 'utf-8',
    'date_format': DATE_FORMAT,
    'timezone': DEFAULT_TIMEZONE,
    'fields': DEFAULT_EXPORT_FIELDS,
    'pretty_print': True
}

# Logging configuration
LOGGING_CONFIG = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
            'datefmt': '%Y-%m-%d %H:%M:%S'
        },
    },
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'formatter': 'standard',
            'level': 'INFO',
            'stream': 'ext://sys.stdout'
        },
        'file': {
            'class': 'logging.handlers.RotatingFileHandler',
            'formatter': 'standard',
            'level': 'DEBUG',
            'filename': 'outlook_extractor.log',
            'maxBytes': 10485760,  # 10MB
            'backupCount': 5,
            'encoding': 'utf8'
        }
    },
    'loggers': {
        '': {
            'handlers': ['console', 'file'],
            'level': 'DEBUG',
            'propagate': True
        },
        'outlook_extractor': {
            'handlers': ['console', 'file'],
            'level': 'DEBUG',
            'propagate': False
        }
    }
}
