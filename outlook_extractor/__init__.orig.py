"""
Outlook Extractor - A Python package for extracting and analyzing emails from Microsoft Outlook.

This package provides tools to connect to Microsoft Outlook, extract emails based on various
criteria, process email content, and save the results in different formats.
"""

# Import main components to make them available at the package level
from .config import ConfigManager, get_config, load_config
from .logging_setup import setup_logging, get_logger, logger
from .core.outlook_client import OutlookClient
from .core.email_threading import EmailThread, ThreadManager
from .storage import EmailStorage, SQLiteStorage, JSONStorage
from .ui import EmailExtractorUI

# Set default logging handler to avoid "No handler found" warnings
import logging
from logging import NullHandler

# Set default logging handler to avoid "No handler found" warnings
logging.getLogger(__name__).addHandler(NullHandler())

# Package version
__version__ = '0.1.0'

# Package description
__description__ = 'A Python package for extracting and analyzing emails from Microsoft Outlook'

# Package author information
__author__ = 'Your Name'
__email__ = 'your.email@example.com'

# Define what gets imported with 'from outlook_extractor import *'
__all__ = [
    'ConfigManager',
    'get_config',
    'load_config',
    'setup_logging',
    'get_logger',
    'logger',
    '__version__',
    '__description__',
    '__author__',
    '__email__',
]

# Initialize logging with default configuration
setup_logging()