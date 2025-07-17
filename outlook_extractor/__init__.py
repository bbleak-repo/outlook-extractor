"""
Outlook Email Extractor - A tool to extract and process emails from Outlook.
"""

import sys
import platform
import warnings

# Version information
__version__ = "1.1.1"
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
CSVExporter = None
AutoUpdater = None
UpdateError = None
check_for_updates = None

# Import core functionality
try:
    # Import platform-agnostic modules first
    from .storage.base import EmailStorage as BaseStorage
    from .storage.json_storage import JSONStorage
    from .config import ConfigManager, get_config
    from .logging_config import setup_logging
    from .export.csv_exporter import CSVExporter
    
    # Import threading module (platform-agnostic)
    from .core import StoppableThread, WorkerThread, ThreadPool
    
    # Import Windows-specific modules if on Windows
    if IS_WINDOWS:
        try:
            from .core.outlook_client import OutlookClient
            from .core.email_threading import EmailThread
            from .main import OutlookExtractor
            from .auto_updater import AutoUpdater, UpdateError
            from .ui.update_dialog import check_for_updates
            
            # Set up Windows-specific exports
            ThreadManager = ThreadPool  # Use ThreadPool as the default ThreadManager
        except ImportError as e:
            warnings.warn(f"Failed to import Windows-specific modules: {e}")
    else:
        warnings.warn("Outlook integration is only available on Windows")
        
        # Provide dummy implementations for non-Windows platforms
        class DummyOutlookClient:
            def __init__(self, *args, **kwargs):
                raise NotImplementedError("Outlook integration is only available on Windows")
        
        OutlookClient = DummyOutlookClient
        EmailThread = object
        ThreadManager = ThreadPool  # Still provide thread pool functionality
        
    # Import main application
    from .main import OutlookExtractor
    
except ImportError as e:
    warnings.warn(f"Failed to import some modules: {e}")
    
    # Provide dummy implementations if imports failed
    class DummyClass: pass
    
    for name in [
        'BaseStorage', 'JSONStorage', 'ConfigManager', 'get_config',
        'setup_logging', 'CSVExporter', 'StoppableThread', 'WorkerThread',
        'ThreadPool', 'OutlookClient', 'EmailThread', 'ThreadManager',
        'OutlookExtractor', 'AutoUpdater', 'UpdateError', 'check_for_updates'
    ]:
        if name not in locals():
            locals()[name] = DummyClass

# Export public API
__all__ = [
    'OutlookClient', 'EmailThread', 'ThreadManager', 'SQLiteStorage',
    'JSONStorage', 'OutlookExtractor', 'ConfigManager', 'get_config',
    'setup_logging', 'CSVExporter', 'AutoUpdater', 'UpdateError', 
    'check_for_updates'
]

# Export the threading utilities
__all__ = [
    'OutlookClient', 'EmailThread', 'ThreadManager', 'SQLiteStorage',
    'JSONStorage', 'OutlookExtractor', 'ConfigManager', 'get_config',
    'setup_logging', 'CSVExporter', 'AutoUpdater', 'UpdateError', 
    'check_for_updates', 'StoppableThread', 'WorkerThread', 'ThreadPool'
]

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
