"""
Centralized logging configuration for Outlook Extractor.

This module provides a comprehensive, production-ready logging system
with support for both file and UI logging.
"""
import json
import logging
import logging.config
import logging.handlers
import os
import queue
import sys
import threading
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, cast

import FreeSimpleGUI as sg
from typing_extensions import Literal

# Type aliases
LogLevel = Literal['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
LogRecord = Dict[str, Any]

class UILogHandler(logging.Handler):
    """Custom logging handler that updates PySimpleGUI elements."""
    
    def __init__(self, element_key: str, level: int = logging.NOTSET):
        super().__init__(level=level)
        self.element_key = element_key
        self._records: queue.Queue[LogRecord] = queue.Queue()
        self._stop_event = threading.Event()
        self._thread = threading.Thread(
            target=self._update_ui,
            name="UILogHandler-Thread",
            daemon=True
        )
        self._thread.start()
    
    def emit(self, record: logging.LogRecord) -> None:
        """Add log record to the queue for UI update."""
        try:
            # Convert log record to a serializable dict
            log_entry = {
                'timestamp': datetime.fromtimestamp(record.created).isoformat(),
                'level': record.levelname,
                'name': record.name,
                'message': record.getMessage(),
                'pathname': record.pathname,
                'lineno': record.lineno,
                'funcName': record.funcName,
                'exc_info': record.exc_info,
                'thread': record.thread,
                'thread_name': record.threadName,
                'process': record.process,
                'process_name': record.processName
            }
            self._records.put(log_entry)
        except Exception as e:
            print(f"Error in UILogHandler.emit: {e}", file=sys.stderr)
    
    def _update_ui(self) -> None:
        """Background thread to update the UI with new log entries."""
        while not self._stop_event.is_set():
            try:
                window = sg.Window._active_window or sg.Window._window_that_exited
                if window and window.TKroot and window.TKroot.winfo_exists():
                    records_to_process: List[LogRecord] = []
                    while not self._records.empty():
                        records_to_process.append(self._records.get_nowait())
                    
                    if records_to_process:
                        window.write_event_value(
                            '-LOG-UPDATE-',
                            {'handler': self.element_key, 'records': records_to_process}
                        )
                
                time.sleep(0.1)  # Prevent CPU spinning
            except Exception as e:
                print(f"Error in _update_ui: {e}", file=sys.stderr)
                time.sleep(1)  # Prevent tight loop on error
    
    def close(self) -> None:
        """Clean up resources."""
        self._stop_event.set()
        if self._thread.is_alive():
            self._thread.join(timeout=2.0)
        super().close()

class JSONFormatter(logging.Formatter):
    """Custom formatter that outputs logs in JSON format."""
    
    def format(self, record: logging.LogRecord) -> str:
        """Format the log record as a JSON string."""
        log_record = {
            'timestamp': datetime.fromtimestamp(record.created).isoformat(),
            'level': record.levelname,
            'name': record.name,
            'message': record.getMessage(),
            'pathname': record.pathname,
            'lineno': record.lineno,
            'funcName': record.funcName,
            'process': record.process,
            'thread': record.thread,
        }
        
        # Add exception info if present
        if record.exc_info:
            log_record['exc_info'] = self.formatException(record.exc_info)
        
        return json.dumps(log_record, ensure_ascii=False)

def setup_logging(
    log_level: Union[str, int] = 'INFO',
    log_file: Optional[Union[str, Path]] = None,
    max_bytes: int = 10 * 1024 * 1024,  # 10MB
    backup_count: int = 5
) -> Dict[str, Any]:
    """
    Configure the logging system with production-ready settings.
    
    Args:
        log_level: The log level as string or int
        log_file: Path to the log file
        max_bytes: Maximum log file size before rotation
        backup_count: Number of backup log files to keep
        
    Returns:
        Dict containing the logging configuration
    """
    if isinstance(log_level, str):
        log_level = getattr(logging, log_level.upper(), logging.INFO)
    
    # Default log file if not specified
    if not log_file:
        log_dir = Path.home() / '.outlook_extractor' / 'logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / 'outlook_extractor.log'
    
    log_file = str(Path(log_file).resolve())
    
    # Create log directory if it doesn't exist
    log_dir = os.path.dirname(log_file)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)
    
    # Base configuration
    config = {
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'json': {
                '()': f'{__name__}.JSONFormatter',
            },
            'console': {
                'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                'datefmt': '%Y-%m-%d %H:%M:%S',
            },
            'ui': {
                'format': '%(asctime)s - %(levelname)s - %(message)s',
                'datefmt': '%H:%M:%S',
            },
        },
        'handlers': {
            'console': {
                'class': 'logging.StreamHandler',
                'level': log_level,
                'formatter': 'console',
                'stream': 'ext://sys.stderr',
            },
            'file': {
                'class': 'logging.handlers.RotatingFileHandler',
                'level': log_level,
                'formatter': 'json',
                'filename': log_file,
                'maxBytes': max_bytes,
                'backupCount': backup_count,
                'encoding': 'utf-8',
            },
        },
        'loggers': {
            'outlook_extractor': {
                'level': log_level,
                'handlers': ['console', 'file'],
                'propagate': False,
            },
        },
        'root': {
            'level': 'WARNING',
            'handlers': ['console', 'file'],
        },
    }
    
    # Apply the configuration
    logging.config.dictConfig(config)
    
    # Set up uncaught exception handler
    def handle_exception(exc_type, exc_value, exc_traceback):
        """Handle uncaught exceptions and log them."""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
            
        logger = logging.getLogger('outlook_extractor')
        logger.critical(
            "Uncaught exception",
            exc_info=(exc_type, exc_value, exc_traceback)
        )
    
    sys.excepthook = handle_exception
    
    return config

def add_ui_handler(element_key: str, log_level: Union[str, int] = 'INFO') -> UILogHandler:
    """
    Add a UI log handler to the root logger.
    
    Args:
        element_key: The PySimpleGUI element key to update with log messages
        log_level: The log level for this handler
        
    Returns:
        The created UILogHandler instance
    """
    if isinstance(log_level, str):
        log_level = getattr(logging, log_level.upper(), logging.INFO)
    
    # Create and configure the UI handler
    handler = UILogHandler(element_key, level=log_level)
    handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    ))
    
    # Add to the root logger
    logging.getLogger('outlook_extractor').addHandler(handler)
    
    return handler

def get_logger(name: Optional[str] = None) -> logging.Logger:
    """
    Get a logger with the specified name.
    
    Args:
        name: The name of the logger. If None, returns the root logger.
        
    Returns:
        A configured Logger instance
    """
    if name is None:
        return logging.getLogger()
    
    # Ensure the logger is under our package namespace
    if not name.startswith('outlook_extractor'):
        name = f'outlook_extractor.{name}'
    
    return logging.getLogger(name)

# Context manager for logging exceptions
class LogErrors:
    """Context manager for logging exceptions."""
    
    def __init__(self, logger: logging.Logger, message: str, *args, **kwargs):
        self.logger = logger
        self.message = message
        self.args = args
        self.kwargs = kwargs
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_val is not None:
            self.logger.error(
                f"{self.message} - {exc_val}",
                *self.args,
                exc_info=(exc_type, exc_val, exc_tb),
                **self.kwargs
            )
        return False  # Don't suppress the exception

# Don't initialize logging automatically to prevent circular imports
