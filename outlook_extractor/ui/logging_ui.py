"""
UI-specific logging functionality for Outlook Extractor.

This module provides UI components and utilities for displaying logs
in the PySimpleGUI interface.
"""
import json
import logging
import queue
import threading
import time
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple, cast

import PySimpleGUI as sg

from outlook_extractor.logging_config import UILogHandler, get_logger

# Type aliases
LogRecord = Dict[str, Any]

class LogViewer:
    """A UI component for viewing and filtering log messages."""
    
    def __init__(
        self,
        element_key: str = '-LOG-',
        max_lines: int = 1000,
        level: str = 'INFO',
        width: int = 100,
        height: int = 20,
        auto_scroll: bool = True,
        show_timestamps: bool = True,
        show_levels: bool = True,
        theme: Optional[str] = None
    ):
        """
        Initialize the log viewer.
        
        Args:
            element_key: The key for the log display element
            max_lines: Maximum number of log lines to keep in memory
            level: Default log level to display
            width: Width of the log display in characters
            height: Height of the log display in lines
            auto_scroll: Whether to automatically scroll to the bottom
            show_timestamps: Whether to show timestamps in the log display
            show_levels: Whether to show log levels in the log display
            theme: PySimpleGUI theme to use for the log viewer
        """
        self.element_key = element_key
        self.max_lines = max_lines
        self.level = getattr(logging, level.upper(), logging.INFO)
        self.auto_scroll = auto_scroll
        self.show_timestamps = show_timestamps
        self.show_levels = show_levels
        self.theme = theme
        
        # Internal state
        self._log_lines: List[str] = []
        self._filtered_indices: List[int] = []
        self._current_filter: Optional[str] = None
        self._current_level = self.level
        self._ui_handler: Optional[UILogHandler] = None
        self._log_queue: queue.Queue[LogRecord] = queue.Queue()
        self._stop_event = threading.Event()
        self._update_thread: Optional[threading.Thread] = None
        
        # Create the log display element
        self.element = sg.Multiline(
            key=element_key,
            size=(width, height),
            autoscroll=auto_scroll,
            write_only=True,
            disabled=True,
            font=('Courier New', 10) if self.theme is None else None,
            text_color='black' if self.theme is None else None,
            background_color='white' if self.theme is None else None,
        )
        
        # Start the update thread
        self._start_update_thread()
    
    def _start_update_thread(self) -> None:
        """Start the background thread for updating the UI."""
        self._stop_event.clear()
        self._update_thread = threading.Thread(
            target=self._process_log_queue,
            name="LogViewer-Update-Thread",
            daemon=True
        )
        self._update_thread.start()
    
    def _process_log_queue(self) -> None:
        """Process log records from the queue and update the UI."""
        while not self._stop_event.is_set():
            try:
                # Process all available records in the queue
                records: List[LogRecord] = []
                while not self._log_queue.empty():
                    try:
                        record = self._log_queue.get_nowait()
                        records.append(record)
                        self._log_queue.task_done()
                    except queue.Empty:
                        break
                
                # Update the UI if we have records
                if records:
                    self._add_records(records)
                
                # Sleep briefly to prevent CPU spinning
                time.sleep(0.1)
                
            except Exception as e:
                # Log the error but don't crash the thread
                get_logger(__name__).error(
                    "Error in log update thread: %s",
                    str(e),
                    exc_info=True
                )
                time.sleep(1)  # Prevent tight loop on error
    
    def _add_records(self, records: List[LogRecord]) -> None:
        """Add log records to the display."""
        # Format the records as text
        formatted_lines = []
        for record in records:
            # Apply level filter
            if record.get('levelno', 0) < self._current_level:
                continue
                
            # Format the line
            parts = []
            
            # Add timestamp if enabled
            if self.show_timestamps:
                timestamp = record.get('timestamp', '')
                if timestamp:
                    try:
                        # Try to parse and reformat the timestamp
                        dt = datetime.fromisoformat(timestamp)
                        timestamp = dt.strftime('%H:%M:%S')
                    except (ValueError, TypeError):
                        pass
                parts.append(f"{timestamp}")
            
            # Add log level if enabled
            if self.show_levels:
                level = record.get('level', 'INFO')
                parts.append(f"{level:8}")
            
            # Add the message
            message = record.get('message', '')
            parts.append(message)
            
            # Add exception info if present
            exc_info = record.get('exc_info')
            if exc_info:
                parts.append(f"\n{exc_info}")
            
            formatted_lines.append(" ".join(parts).strip())
        
        # Update the log display
        if formatted_lines:
            # Add to our internal buffer
            self._log_lines.extend(formatted_lines)
            
            # Trim the buffer if it's too large
            if len(self._log_lines) > self.max_lines:
                self._log_lines = self._log_lines[-self.max_lines:]
            
            # Update the UI
            if hasattr(self, 'window') and self.window:
                try:
                    self.window.write_event_value(
                        '-LOG-UPDATE-',
                        {'element': self.element_key, 'text': '\n'.join(self._log_lines)}
                    )
                except Exception as e:
                    get_logger(__name__).error(
                        "Error updating log display: %s",
                        str(e),
                        exc_info=True
                    )
    
    def attach_to_window(self, window: sg.Window) -> None:
        """Attach the log viewer to a PySimpleGUI window."""
        self.window = window
        
        # Set up the UI log handler if not already done
        if self._ui_handler is None:
            self._ui_handler = UILogHandler(self.element_key)
            logging.getLogger('outlook_extractor').addHandler(self._ui_handler)
    
    def set_level(self, level: str) -> None:
        """Set the minimum log level to display."""
        self._current_level = getattr(logging, level.upper(), logging.INFO)
        self.refresh()
    
    def set_filter(self, text: Optional[str]) -> None:
        """Set a filter for log messages."""
        self._current_filter = text.lower() if text else None
        self.refresh()
    
    def refresh(self) -> None:
        """Refresh the log display with current filters."""
        if hasattr(self, 'window') and self.window:
            self.window[self.element_key].update('\n'.join(self._log_lines))
    
    def clear(self) -> None:
        """Clear the log display."""
        self._log_lines.clear()
        if hasattr(self, 'window') and self.window:
            self.window[self.element_key].update('')
    
    def close(self) -> None:
        """Clean up resources."""
        self._stop_event.set()
        
        # Stop the update thread
        if self._update_thread and self._update_thread.is_alive():
            self._update_thread.join(timeout=2.0)
        
        # Remove the UI handler
        if self._ui_handler:
            logging.getLogger('outlook_extractor').removeHandler(self._ui_handler)
            self._ui_handler.close()
            self._ui_handler = None
        
        # Clear the window reference
        if hasattr(self, 'window'):
            del self.window

def create_log_viewer_frame(
    title: str = "Application Logs",
    element_key: str = '-LOG-',
    width: int = 100,
    height: int = 20,
    level: str = 'INFO',
    show_controls: bool = True,
    theme: Optional[str] = None
) -> Tuple[sg.Frame, LogViewer]:
    """
    Create a log viewer frame with controls.
    
    Args:
        title: Title for the frame
        element_key: Key for the log display element
        width: Width of the log display in characters
        height: Height of the log display in lines
        level: Default log level to display
        show_controls: Whether to show log level and filter controls
        theme: PySimpleGUI theme to use
        
    Returns:
        A tuple of (frame_element, log_viewer_instance)
    """
    # Create the log viewer
    log_viewer = LogViewer(
        element_key=element_key,
        width=width,
        height=height,
        level=level,
        theme=theme
    )
    
    # Create the layout
    layout = [
        [log_viewer.element]
    ]
    
    # Add controls if requested
    if show_controls:
        controls = [
            [
                sg.Text("Log Level:"),
                sg.Combo(
                    ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                    default_value=level.upper(),
                    key=f'{element_key}-LEVEL',
                    enable_events=True,
                    readonly=True,
                    size=(10, 1)
                ),
                sg.Text("Filter:"),
                sg.Input(
                    key=f'{element_key}-FILTER',
                    size=(30, 1),
                    enable_events=True
                ),
                sg.Button("Clear", key=f'{element_key}-CLEAR')
            ]
        ]
        layout = controls + layout
    
    # Create the frame
    frame = sg.Frame(title, layout, expand_x=True, expand_y=True)
    
    return frame, log_viewer
