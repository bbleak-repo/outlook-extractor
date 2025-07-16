"""
Outlook Email Extractor - A tool to extract and process emails from Outlook.
"""

import sys
import platform
import warnings

# Version information
__version__ = "1.2.0"
__author__ = "Your Name"
__license__ = "MIT"

# Platform detection
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'
IS_LINUX = platform.system() == 'Linux'

# Initialize variables for exports
OutlookClient = None
EmailThread = None
ThreadManager = None
SQLiteStorage = None
JSONStorage = None
OutlookExtractor = None
ConfigManager = None
get_config = None
setup_logging = None
AutoUpdater = None
UpdateError = None
check_for_updates = None

# Import core functionality
try:
    from .storage.sqlite_storage import SQLiteStorage
    from .storage.json_storage import JSONStorage
    from .config import ConfigManager, get_config
    from .logging_config import setup_logging
    from .export.csv_exporter import CSVExporter
    
    # Import Windows-specific modules if on Windows
    if IS_WINDOWS:
        try:
            from .core.outlook_client import OutlookClient
            from .core.threading import EmailThread, ThreadManager
            from .main import OutlookExtractor
            from .auto_updater import AutoUpdater, UpdateError
            from .ui.update_dialog import check_for_updates
        except ImportError as e:
            warnings.warn(f"Failed to import Windows-specific modules: {e}")
    else:
        warnings.warn("Outlook integration is only available on Windows")
        
    # Import platform-agnostic modules
    from .core.threading import EmailThread, ThreadManager
    from .main import OutlookExtractor
    
except ImportError as e:
    warnings.warn(f"Failed to import some modules: {e}")

# Export public API
__all__ = [
    'OutlookClient', 'EmailThread', 'ThreadManager', 'SQLiteStorage',
    'JSONStorage', 'OutlookExtractor', 'ConfigManager', 'get_config',
    'setup_logging', 'CSVExporter', 'AutoUpdater', 'UpdateError', 
    'check_for_updates'
]

# Only import Windows-specific modules on Windows
if sys.platform == 'win32':
    try:
        from .core.outlook_client import OutlookClient
        from .core.email_threading import EmailThread, ThreadManager
        from .storage import SQLiteStorage, JSONStorage
        from .main import OutlookExtractor
        from .config import ConfigManager, get_config
        from .logging_setup import setup_logging
    except ImportError as e:
        print(f"Warning: Could not import Windows-specific modules: {e}")

# Always import export functionality as it's platform-independent
try:
    from .export import CSVExporter
except ImportError as e:
    print(f"Warning: Could not import export modules: {e}")

__version__ = '1.0.0'
__all__ = [
    'OutlookClient',
    'EmailThread',
    'ThreadManager',
    'SQLiteStorage',
    'JSONStorage',
    'OutlookExtractor',
    'ConfigManager',
    'get_config',
    'setup_logging',
    'CSVExporter'
]
