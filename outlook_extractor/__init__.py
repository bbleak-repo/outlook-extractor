# outlook_extractor/__init__.py
import sys

# Define default values for all platforms
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
