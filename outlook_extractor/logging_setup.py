"""
Logging configuration for the Outlook Extractor package.
"""
import logging
import logging.handlers
import os
import sys
import traceback
from pathlib import Path
from typing import Optional, Type, Any, TypeVar

# Type variable for exception handling
E = TypeVar('E', bound=Exception)

class ColoredFormatter(logging.Formatter):
    """Custom formatter that adds colors to log levels."""
    
    # ANSI color codes
    COLORS = {
        'DEBUG': '\033[36m',     # Cyan
        'INFO': '\033[32m',      # Green
        'WARNING': '\033[33m',   # Yellow
        'ERROR': '\033[31m',     # Red
        'CRITICAL': '\033[41m',  # Red background
        'RESET': '\033[0m'       # Reset to default
    }
    
    def format(self, record):
        """Format the specified record as text with colors."""
        # Get the level name and color code
        levelname = record.levelname
        color = self.COLORS.get(levelname, self.COLORS['RESET'])
        
        # Add color to the level name
        record.levelname = f"{color}{levelname}{self.COLORS['RESET']}"
        
        # Format the message with color
        return super().format(record)

def setup_logging(config=None, log_file: Optional[str] = None, 
                 log_level: Optional[str] = None) -> logging.Logger:
    """Set up logging with the specified configuration.
    
    Args:
        config: Optional ConfigManager instance. If not provided, uses default config.
        log_file: Optional path to the log file. If not provided, uses config value.
        log_level: Optional log level as a string (DEBUG, INFO, WARNING, ERROR, CRITICAL).
                  If not provided, uses config value.
                  
    Returns:
        A configured logger instance.
    """
    # Create a logger
    logger = logging.getLogger('outlook_extractor')
    
    # Don't propagate to root logger
    logger.propagate = False
    
    # Clear any existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()
    
    # Set default log level to INFO
    logger.setLevel(logging.INFO)
    
    # Default log file path if not provided
    if log_file is None and config is not None:
        try:
            log_file = config.get('logging', 'file', fallback='outlook_extractor.log')
        except Exception:
            log_file = 'outlook_extractor.log'
    
    # Default log level if not provided
    if log_level is None and config is not None:
        try:
            log_level = config.get('logging', 'level', fallback='INFO')
        except Exception:
            log_level = 'INFO'
    
    # Convert log level string to logging level
    level = getattr(logging, str(log_level).upper(), logging.INFO)
    logger.setLevel(level)
    
    # Create formatters
    console_formatter = ColoredFormatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Create a default formatter for file output
    file_formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Add console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(console_formatter)
    console_handler.setLevel(level)
    logger.addHandler(console_handler)
    
    # Add file handler if log file is specified
    if log_file:
        try:
            # Ensure the log directory exists
            log_path = Path(log_file).resolve()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Create file handler with rotation (10MB per file, keep 5 backups)
            file_handler = logging.handlers.RotatingFileHandler(
                str(log_path),
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5,
                encoding='utf-8',
                delay=True  # Don't create empty log files
            )
            file_handler.setFormatter(file_formatter)
            file_handler.setLevel(level)
            logger.addHandler(file_handler)
            
            logger.debug(f"Logging to file: {log_path}")
            
        except Exception as e:
            logger.warning(f"Failed to set up file logging: {e}")
            traceback.print_exc()
    
    # Set up uncaught exception handler
    def handle_exception(exc_type, exc_value, exc_traceback):
        """Handle uncaught exceptions and log them."""
        # Don't log keyboard interrupts
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
            
        logger.critical(
            "Uncaught exception:", 
            exc_info=(exc_type, exc_value, exc_traceback)
        )
        
        # For critical errors, also log to stderr
        if exc_type is not SystemExit:
            traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
    
    sys.excepthook = handle_exception
    
    # Log successful setup
    logger.info("Logging system initialized successfully")
    
    return logger

def get_logger(name: str = None) -> logging.Logger:
    """Get a logger with the specified name.
    
    Args:
        name: The name of the logger. If None, returns the root logger.
             If '__name__' is passed, uses the module's name.
             
    Returns:
        A configured logger instance.
    """
    if name is None:
        return logging.getLogger()
    
    # If name is '__name__', use the module's name
    if name == '__main__' or name == '__mp_main__':
        name = 'outlook_extractor'
    elif name.startswith('__') and name.endswith('__'):
        # If it's a special name like __main__, use the root logger
        return logging.getLogger()
    
    # Ensure the logger is a child of the main logger
    if not name.startswith('outlook_extractor'):
        name = f'outlook_extractor.{name}'
    
    return logging.getLogger(name)

# Default logger instance
logger = get_logger(__name__)

# Example usage:
if __name__ == '__main__':
    # Configure logging
    logger = setup_logging()
    
    # Log some messages
    logger.debug('This is a debug message')
    logger.info('This is an info message')
    logger.warning('This is a warning message')
    logger.error('This is an error message')
    logger.critical('This is a critical message')
    
    try:
        # This will raise an exception
        1 / 0
    except Exception as e:
        logger.exception('An error occurred: %s', e)
