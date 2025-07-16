"""Main window for the Outlook Extractor application.

This module provides the main user interface for the application,
including all tabs and their associated functionality.
"""

import logging
import os
import shutil
import sys
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, cast, Union

import PySimpleGUI as sg

from outlook_extractor import __version__
from outlook_extractor.config import ConfigManager, load_config, get_config
from outlook_extractor.logging_config import setup_logging, get_logger
from outlook_extractor.logging_utils import log_errors
from outlook_extractor.ui.export_tab import ExportTab
from outlook_extractor.ui.update_dialog import check_for_updates

# Import the export tab
from .export_tab import ExportTab as ExtractionTab

class EmailExtractorUI:
    def __init__(self, config_path: str = None):
        """Initialize the main window.
        
        Args:
            config_path: Path to the configuration file
        """
        # Initialize logging first, before any other operations
        self._init_logging()
        
        # Get logger instance after logging is initialized
        self.logger = get_logger(__name__)
        self.logger.info('Initializing EmailExtractorUI')
        
        # Track if we've already checked for updates
        self._update_checked = False
        self._last_update_check = 0  # Timestamp of last update check
        self._window_initialized = False  # Track if window is fully initialized
        
        try:
            # Load configuration
            self.logger.debug('Loading configuration...')
            self.config = load_config(config_path)
            self.config_path = config_path  # Store the provided config path
            self.logger.info('Configuration loaded successfully')
            
            # Initialize UI components
            self.window = None
            self.theme = 'LightGrey1'  # Default theme
            self.current_folder_patterns = []
            
            # Set up theme and UI
            self.setup_theme()
            
            # Create the main window
            self.logger.debug('Creating main window...')
            self.window = self.create_window()
            
            # Mark window as initialized
            self._window_initialized = True
            self.logger.debug('Window initialization complete')
            
            # Initialize export tab as None, will be created when first accessed
            self.export_tab = None
            self.logger.debug('Export tab will be initialized on first access')
            
        except Exception as e:
            self.logger.critical('Failed to initialize application', exc_info=True)
            # If window was partially initialized, clean it up
            if hasattr(self, 'window') and self.window:
                try:
                    self.window.close()
                except:
                    pass
            raise
        
    def _init_logging(self) -> None:
        """
        Initialize the logging system with robust error handling.
        Ensures logging directory exists and configures basic logging before any other operations.
        """
        # Configure basic logging first as a fallback
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        
        # Get root logger and remove any existing handlers
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        try:
            # Configure log directory with proper permissions
            log_dir = Path.home() / '.outlook_extractor' / 'logs'
            try:
                log_dir.mkdir(parents=True, exist_ok=True, mode=0o755)
                if not os.access(log_dir, os.W_OK):
                    raise PermissionError(f'Cannot write to log directory: {log_dir}')
            except Exception as e:
                raise RuntimeError(f'Failed to create log directory {log_dir}: {str(e)}') from e
            
            log_file = log_dir / 'outlook_extractor.log'
            
            # Configure console handler
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)
            
            # Configure file handler with rotation
            try:
                file_handler = logging.handlers.RotatingFileHandler(
                    filename=str(log_file),
                    maxBytes=10 * 1024 * 1024,  # 10MB
                    backupCount=5,
                    encoding='utf-8',
                    delay=True  # Delay file opening until first write
                )
                file_handler.setLevel(logging.DEBUG)
                
                # Set formatters
                formatter = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                )
                console_handler.setFormatter(formatter)
                file_handler.setFormatter(formatter)
                
                # Add handlers
                root_logger.setLevel(logging.DEBUG)
                root_logger.addHandler(console_handler)
                root_logger.addHandler(file_handler)
                
                # Log successful initialization
                root_logger.info('=' * 80)
                root_logger.info('Application starting...')
                root_logger.info(f'Log file: {log_file.absolute()}')
                
            except Exception as e:
                # Fall back to basic config if file logging fails
                logging.basicConfig(
                    level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[console_handler]
                )
                logging.warning(f'Could not set up file logging: {str(e)}', exc_info=True)
                
        except Exception as e:
            # Last resort error handling if logging setup fails completely
            import sys
            sys.stderr.write(f'CRITICAL: Failed to initialize logging: {str(e)}\n')
            sys.stderr.flush()
            # Re-raise with a more descriptive error
            raise RuntimeError('Failed to initialize logging system') from e
            
    def _configure_logging(self, window) -> None:
        """
        Configure logging to include the window's log element if available.
        
        Args:
            window: The PySimpleGUI window object that may contain a log element
        """
        if not window or not hasattr(window, 'log'):
            logging.debug('Window or log element not available, skipping GUI log configuration')
            return
            
        root_logger = logging.getLogger()
        
        # Remove any existing window handlers to prevent duplicates
        for handler in root_logger.handlers[:]:
            if hasattr(handler, 'window'):
                try:
                    root_logger.removeHandler(handler)
                except Exception as e:
                    logging.error(f'Error removing existing window handler: {str(e)}')
        
        class WindowLogHandler(logging.Handler):
            """Custom logging handler that writes to a PySimpleGUI Multiline element."""
            def __init__(self, window):
                super().__init__()
                self.window = window
                self.setLevel(logging.INFO)  # Only show INFO and above in GUI
                self.setFormatter(logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%H:%M:%S'
                ))
                self._initialized = False
                
            def emit(self, record):
                if not self._initialized:
                    return
                    
                try:
                    # Format the message
                    msg = self.format(record)
                    
                    # Update the GUI element in a thread-safe way
                    if hasattr(self.window, 'write_event_value'):
                        self.window.write_event_value(('LOG_MESSAGE', msg + '\n'), None)
                    else:
                        # Fallback for older PySimpleGUI versions
                        self.window.log.update(value=msg + '\n', append=True)
                    
                    # Auto-scroll to the bottom if possible
                    try:
                        self.window.log.set_vscroll_position(1.0)
                    except Exception:
                        pass
                        
                except Exception as e:
                    # If we can't log to the window, at least don't crash
                    try:
                        sys.stderr.write(f'Error in WindowLogHandler: {str(e)}\n')
                    except:
                        pass  # Absolute last resort
            
            def set_initialized(self, value):
                self._initialized = value
        
        try:
            # Add the window handler
            window_handler = WindowLogHandler(window)
            root_logger.addHandler(window_handler)
            
            # Mark handler as initialized after a short delay to avoid early logging
            import threading
            def delayed_init():
                import time
                time.sleep(1.0)  # Wait for window to be fully initialized
                window_handler.set_initialized(True)
                logging.info('GUI logging handler initialized')
                
            init_thread = threading.Thread(target=delayed_init, daemon=True)
            init_thread.start()
            
        except Exception as e:
            logging.error(f'Failed to configure GUI logging: {str(e)}', exc_info=True)
            # Ensure we still have console logging if GUI logging fails
            if not any(isinstance(h, logging.StreamHandler) for h in root_logger.handlers):
                root_logger.addHandler(logging.StreamHandler())
        
    def setup_theme(self):
        """Setup the PySimpleGUI theme and settings."""
        sg.theme(self.theme)
        # Only set theme-specific options here
        sg.set_options(
            font=('Arial', 10),
            element_padding=(5, 5),
            button_element_size=(12, 1),
            auto_size_buttons=False
        )
    
    def create_window(self):
        """Create the main application window with robust error handling."""
        self.logger.debug('Starting window creation...')
        
        try:
            # Create a loading indicator for the export tab
            export_loading = [
                [sg.Text('Loading export settings...', font=('Arial', 12))],
                [sg.ProgressBar(100, orientation='h', size=(20, 20), key='-EXPORT_LOADING-')]
            ]
            
            # Create the tab group with all tabs
            tab_group = [
                # Extraction Tab
                sg.Tab('Extraction', [[sg.Column(
                    self._create_extraction_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-TAB_EXTRACTION-'),
                
                # Storage Tab
                sg.Tab('Storage', [[sg.Column(
                    self._create_storage_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-TAB_STORAGE-'),
                
                # Threading Tab
                sg.Tab('Threading', [[sg.Column(
                    self._create_threading_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-TAB_THREADING-'),
                
                # Email Processing Tab
                sg.Tab('Email Processing', [[sg.Column(
                    self._create_email_processing_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-TAB_EMAIL_PROCESSING-'),
                
                # Security Tab
                sg.Tab('Security', [[sg.Column(
                    self._create_security_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-TAB_SECURITY-'),
                
                # Export Tab (with loading placeholder)
                sg.Tab('Export', [
                    [sg.Column(export_loading, key='-EXPORT_LOADING-', expand_x=True, expand_y=True)]
                ], key='-EXPORT_TAB-'),
                
                # Logs Tab - Must be last to ensure it's fully initialized
                sg.Tab('Logs', [[sg.Multiline(
                    size=(80, 25),
                    autoscroll=True,
                    auto_refresh=True,
                    write_only=True,
                    key='-LOG-',
                    disabled=True,
                    background_color='#f0f0f0',
                    text_color='black',
                    font=('Courier', 10),
                    expand_x=True,
                    expand_y=True,
                    reroute_stdout=False,  # Disable automatic rerouting to prevent duplicates
                    reroute_stderr=False,  # Disable automatic rerouting to prevent duplicates
                    echo_stdout_stderr=False
                )]], key='-LOGS_TAB-'),
                
                # About Tab
                sg.Tab('About', [[sg.Column(
                    self._create_about_tab(),
                    scrollable=True,
                    vertical_scroll_only=True,
                    expand_x=True,
                    expand_y=True
                )]], key='-ABOUT_TAB-')
            ]
            
            # Define the main layout with proper expansion
            layout = [
                [self._create_menu_bar()],
                [
                    sg.Column(
                        [
                            [
                                sg.TabGroup(
                                    [tab_group],
                                    key='-TAB_GROUP-',
                                    expand_x=True,
                                    expand_y=True,
                                    tab_location='top',
                                    enable_events=True
                                    # Using minimal parameters for maximum compatibility
                                )
                            ]
                        ],
                        expand_x=True,
                        expand_y=True,
                        pad=(5, 5)
                    )
                ],
                [
                    sg.StatusBar('Ready', key='-STATUS-', size=(20, 1), expand_x=True),
                    sg.Push(),
                    sg.Button('Backup', key='-BACKUP-'),
                    sg.Button('Run Extraction', key='-RUN-', button_color=('white', 'green')),
                    sg.Button('Exit', key='-EXIT-', size=(8, 1))
                ]
            ]
            
            # Create the main window with error handling
            try:
                self.logger.debug('Creating main window...')
                window = sg.Window(
                    f'Outlook Email Extractor {__version__}',
                    layout,
                    finalize=True,
                    resizable=True,
                    size=(1200, 800),  # Slightly larger default size
                    element_justification='center',
                    font=('Arial', 10),
                    enable_close_attempted_event=True,
                    location=(None, None),  # Let the window manager handle initial position
                    margins=(10, 10),
                    element_padding=(5, 5)
                )
                
                # Set the window to be maximized if supported
                try:
                    window.maximize()
                except Exception as e:
                    self.logger.warning(f'Could not maximize window: {str(e)}')
                    window.size = (1200, 800)
                
                # Configure logging to use the window's log element
                self.logger.debug('Configuring window logging...')
                self._configure_logging(window)
                
                # Log window creation success
                self.logger.info('Main window created successfully')
                self.logger.debug(f'Window size: {window.size}')
                
                return window
                
            except Exception as e:
                self.logger.critical(f'Failed to create main window: {str(e)}', exc_info=True)
                # Try to create a minimal error window
                try:
                    sg.popup_error(
                        'Failed to create main window',
                        str(e),
                        title='Fatal Error',
                        keep_on_top=True
                    )
                except:
                    pass
                raise
                
        except Exception as e:
            self.logger.critical(f'Failed to create window layout: {str(e)}', exc_info=True)
            # If we get here, something went very wrong with the layout
            try:
                sg.popup_error(
                    'Failed to create window layout',
                    str(e),
                    title='Fatal Error',
                    keep_on_top=True
                )
            except:
                pass
            raise
                # Continue without GUI logging rather than failing
            
            # Now that the window is created, initialize the export tab
            self.export_tab = ExportTab(self.config)
            export_column = self.export_tab.get_layout()
            
            # Store window reference in export tab if needed
            if hasattr(self.export_tab, 'window'):
                self.export_tab.window = self.window
            
            # Replace the placeholder with the actual export tab content
            self.window.extend_layout(self.window['-EXPORT_TAB-'], [[export_column]])
            self.window['-EXPORT_LOADING-'].update(visible=False)
            
            # Initialize folder patterns if available
            if hasattr(self, 'window') and self.window:
                if '-FOLDER_PATTERNS-' in self.window.AllKeysDict:
                    folder_patterns = self.window['-FOLDER_PATTERNS-'].get().split(',')
                    folder_patterns = [p.strip() for p in folder_patterns if p.strip()]
                    if folder_patterns:
                        self.export_tab.update_folder_patterns(folder_patterns)
            
            self.logger.info('Export tab loaded successfully')
            
        except Exception as e:
            error_msg = f'Failed to initialize export tab: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            
            # Show error in the tab
            error_layout = [
                [sg.Text('⚠️ Error Loading Export Tab', font=('Helvetica', 12, 'bold'), text_color='red')],
                [sg.Multiline(
                    f'An error occurred while loading the export tab.\n\nError: {str(e)}',
                    size=(80, 10),
                    disabled=True,
                    text_color='red',
                    background_color='#f9f9f9',
                    expand_x=True,
                    expand_y=True
                )],
                [sg.Text('Please check the application logs for more details.', text_color='orange')],
                [sg.Button('Retry', key='-RETRY_EXPORT_TAB-'), sg.Button('Close', key='-CLOSE_ERROR-')]
            ]
            
            # If window exists, update it, otherwise show error in a popup
            if hasattr(self, 'window') and self.window:
                self.window['-EXPORT_LOADING-'].update(visible=False)
                self.window.extend_layout(self.window['-EXPORT_TAB-'], [error_layout])
            else:
                sg.popup_error('Failed to initialize export tab', str(e))
                
        # Load current config into the UI
        self._load_config_to_ui()
        
        return self.window
    
    def _create_menu_bar(self):
        """Create the menu bar."""
        menu_def = [
            ['&File', ['&Open Config', '&Save Config', '&Backup Config', '---', 'E&xit::exit']],
            ['&Tools', ['&Options', '&Backup Data']],
            ['&Help', ['Check for &Updates', '&About', '&Documentation']]
        ]
        return sg.Menubar(menu_def, tearoff=False, key='-MENUBAR-')
    
    def _create_extraction_tab(self):
        """Create the extraction settings tab."""
        return [
            [sg.Frame('Outlook Settings', [
                [sg.T('Mailbox (leave empty for default):', 
                     tooltip='Leave empty to use your default Outlook profile, or enter an email address')],
                [sg.Input(key='-MAILBOX-', size=(45, 1), 
                         tooltip='Enter an email address or leave empty for default Outlook profile')],
                [sg.T('Folder Patterns:', 
                     tooltip='Comma-separated list of folder patterns to include. Supports wildcards (e.g., *Legal*, Inbox/*)')],
                [sg.Input(key='-FOLDER_PATTERNS-', size=(45, 1), 
                         tooltip='Example: Inbox,Sent Items,*Legal*,Projects/*')],
                [sg.T('Max Emails to Extract:', 
                     tooltip='Maximum number of emails to process (0 for no limit)')],
                [sg.Spin([i*100 for i in range(0, 101)], 1000, key='-MAX_EMAILS-', size=(6, 1))]
            ])],
            [
                sg.Frame('Date Range', [
                    [sg.Radio('Last N Days', 'DATE_RANGE', default=True, key='-DATE_RANGE_DAYS-', 
                             enable_events=True, tooltip='Extract emails from the last N days')],
                    [
                        sg.T('   '),  # Indent for better visual hierarchy
                        sg.T('Days:', text_color='#404040'),
                        sg.Input('30', key='-DAYS_BACK-', size=(5, 1), disabled=False, 
                                 text_color='black', background_color='white'),
                        sg.T('  '),  # Spacer
                        sg.T('(e.g., 30 = last 30 days)', font=('Arial', 8), text_color='#666666')
                    ],
                    [sg.HorizontalSeparator()],
                    [sg.Radio('Custom Date Range', 'DATE_RANGE', key='-DATE_RANGE_CUSTOM-', 
                             enable_events=True, tooltip='Specify a custom date range')],
                    [
                        sg.T('   '),  # Indent for better visual hierarchy
                        sg.T('From:', text_color='#404040'),
                        sg.Input(key='-START_DATE-', size=(12, 1), disabled=True, 
                                text_color='#000000', background_color='#FFFFFF',
                                disabled_readonly_background_color='#F0F0F0',
                                disabled_readonly_text_color='#000000'),
                        sg.CalendarButton('📅', target='-START_DATE-', format='%Y-%m-%d', 
                                        button_color=('white', '#4B8BBE'), size=(2, 1),
                                        pad=((2, 5), (0, 0))),
                        sg.T('   To:', text_color='#404040'),
                        sg.Input(key='-END_DATE-', size=(12, 1), disabled=True, 
                               text_color='#000000', background_color='#FFFFFF',
                               disabled_readonly_background_color='#F0F0F0',
                               disabled_readonly_text_color='#000000'),
                        sg.CalendarButton('📅', target='-END_DATE-', format='%Y-%m-%d', 
                                        button_color=('white', '#4B8BBE'), size=(2, 1),
                                        pad=((2, 0), (0, 0))),
                    ]
                ], element_justification='left')
            ]
        ]
    
    def _create_storage_tab(self):
        """Create the storage settings tab."""
        return [
            [sg.Frame('Storage Settings', [
                [sg.T('Output Directory:'), 
                 sg.Input(key='-OUTPUT_DIR-', size=(40, 1), default_text='output'),
                 sg.FolderBrowse('Browse')],
                [sg.T('Storage Type:'),
                 sg.Combo(['sqlite', 'json'], default_value='sqlite', key='-STORAGE_TYPE-', 
                          readonly=True, enable_events=True)],
                [sg.T('Database File:'), 
                 sg.Input(key='-DB_FILENAME-', size=(40, 1), default_text='emails.db'),
                 sg.FileSaveAs('Save As', file_types=(('SQLite Database', '*.db'),))],
                [sg.CB('Export to JSON', default=True, key='-EXPORT_JSON-', enable_events=True),
                 sg.Input(key='-JSON_FILENAME-', size=(40, 1), default_text='emails.json', 
                          disabled=False),
                 sg.FileSaveAs('Save As', file_types=(('JSON File', '*.json'),))]
            ])]
        ]
    
    def _create_threading_tab(self):
        """Create the threading settings tab."""
        return [
            [sg.Frame('Threading Settings', [
                [sg.CB('Enable Threading', default=True, key='-ENABLE_THREADING-',
                      tooltip='Group related emails into conversation threads')],
                
                [sg.T('Threading Method:', tooltip='How emails are grouped into threads:')],
                [sg.T('   • Hybrid: Best of both methods (recommended)')],
                [sg.T('   • Subject: Groups by email subject only')],
                [sg.T('   • Conversation: Uses Outlook conversation IDs')],
                [sg.Combo(
                    ['hybrid', 'subject', 'conversation'], 
                    default_value='hybrid', 
                    key='-THREAD_METHOD-', 
                    readonly=True,
                    size=(15, 1)
                )],
                
                [sg.T('\nThread Depth Control:', 
                     tooltip='How many levels of replies to include in a thread')],
                [sg.T('   Max Depth:'), 
                 sg.Spin(
                     list(range(1, 51)), 
                     10, 
                     key='-MAX_THREAD_DEPTH-', 
                     size=(5, 1),
                     tooltip='Maximum depth of email threads to process (1-50)'
                 )],
                 
                [sg.T('\nThread Timeout:', 
                     tooltip='Maximum age difference between emails to consider them part of the same thread')],
                [sg.T('   '),
                 sg.Spin(
                     list(range(1, 366)), 
                     30, 
                     key='-THREAD_TIMEOUT_DAYS-', 
                     size=(5, 1),
                     tooltip='Maximum days between emails in a thread (1-365)'
                 ),
                 sg.T('days')],
                 
                [sg.T('\nNote: Threading can be resource-intensive for large mailboxes.')],
                [sg.T('Consider using filters to reduce the number of emails processed.')]
            ], element_justification='left')]
        ]
    
    def _create_email_processing_tab(self):
        """Create the email processing settings tab."""
        return [
            [sg.Frame('Attachment Settings', [
                [sg.CB('Extract Attachments', default=False, key='-EXTRACT_ATTACHMENTS-')],
                [sg.T('Attachment Directory:'), 
                 sg.Input(key='-ATTACHMENT_DIR-', size=(40, 1), default_text='attachments'),
                 sg.FolderBrowse('Browse')]
            ])],
            [sg.Frame('Content Extraction', [
                [sg.CB('Extract Embedded Images', default=False, key='-EXTRACT_IMAGES-')],
                [sg.T('Image Directory:'), 
                 sg.Input(key='-IMAGE_DIR-', size=(40, 1), default_text='images'),
                 sg.FolderBrowse('Browse')],
                [sg.CB('Extract Links', default=True, key='-EXTRACT_LINKS-')],
                [sg.CB('Extract Phone Numbers', default=True, key='-EXTRACT_PHONES-')]
            ])]
        ]
    
    def _create_security_tab(self):
        """Create the security settings tab."""
        return [
            [sg.Frame('Data Redaction', [
                [sg.CB('Redact Sensitive Data', default=True, key='-REDACT_SENSITIVE-')],
                [sg.T('Redaction Patterns:')],
                [sg.Multiline('password\nssn\ncredit.?card\naccount.?number', 
                             size=(60, 5), key='-REDACTION_PATTERNS-')],
                [sg.T('(One pattern per line, supports regex)')]
            ])]
        ]
    
    def _create_logs_tab(self):
        """Create the logs tab."""
        return [
            [sg.Frame('Logging', [
                [sg.T('Log Level:'),
                 sg.Combo(['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'], 
                         default_value='INFO', key='-LOG_LEVEL-', readonly=True)],
                [sg.T('Log File:'), 
                 sg.Input(key='-LOG_FILE-', size=(40, 1), default_text='outlook_extractor.log'),
                 sg.FileSaveAs('Save As', file_types=(('Log Files', '*.log'),))],
                [sg.Multiline('', size=(80, 20), key='-LOG_OUTPUT-', autoscroll=True, 
                             disabled=True, reroute_stdout=True, reroute_stderr=True)]
            ])]
        ]
    
    def _create_about_tab(self):
        """Create the about tab."""
        about_text = """
        Outlook Email Extractor
        
        Version: 1.0.0
        Author: Your Name
        License: MIT
        
        A tool to extract and analyze emails from Microsoft Outlook.
        
        Features:
        - Extract emails with various filtering options
        - Threaded conversation view
        - Multiple storage backends (SQLite, JSON)
        - Configurable data extraction
        - Secure handling of sensitive information
        
        2025 Your Company. All rights reserved.
        """
        return [
            [sg.Multiline(about_text.strip(), size=(80, 20), disabled=True, 
                         text_color='black', background_color='white')],
            [sg.T('Documentation:'), 
             sg.T('https://github.com/yourusername/outlook-extractor', 
                  enable_events=True, text_color='blue', key='-DOCS_LINK-')]
        ]
    
    def _load_config_to_ui(self):
        """Load configuration values into the UI elements."""
        if not self.window:
            return
            
        try:
            # Outlook settings
            self.window['-MAILBOX-'].update(
                self.config.get('outlook', 'mailbox', fallback=''))
            
            # Get folder patterns from config
            folder_patterns = self.config.get_list('outlook', 'folder_patterns', ['Inbox'])
            self.window['-FOLDER_PATTERNS-'].update(', '.join(folder_patterns))
            
            # Update export tab with initial folder patterns
            if hasattr(self, 'export_tab') and self.export_tab:
                logging.info(f'Initializing ExportTab with folder patterns: {folder_patterns}')
                self.export_tab.update_folder_patterns(folder_patterns)
            
            # Handle folder patterns - ensure it works with wildcards
            folder_patterns = self.config.get('outlook', 'folder_patterns', 
                                           fallback='Inbox,Sent Items')
            self.window['-FOLDER_PATTERNS-'].update(folder_patterns)
            
            # Set max emails (0 means no limit)
            max_emails = self.config.get_int('outlook', 'max_emails', fallback=1000)
            self.window['-MAX_EMAILS-'].update(max_emails if max_emails >= 0 else 0)
                
            # Date range - handle both formats (days back and custom range)
            date_ranges = self.config.get('date_range', 'date_ranges', fallback='')
            if date_ranges and '|' in date_ranges:
                # Custom date range is set
                self.window['-DATE_RANGE_CUSTOM-'].update(True)
                self.window['-DATE_RANGE_DAYS-'].update(False)
                start_date, end_date = date_ranges.split('|')
                self.window['-START_DATE-'].update(start_date)
                self.window['-END_DATE-'].update(end_date)
                self.window['-DAYS_BACK-'].update(disabled=True)
                self.window['-START_DATE-'].update(disabled=False)
                self.window['-END_DATE-'].update(disabled=False)
            else:
                # Use last N days
                self.window['-DATE_RANGE_DAYS-'].update(True)
                self.window['-DATE_RANGE_CUSTOM-'].update(False)
                self.window['-DAYS_BACK-'].update(
                    self.config.get('date_range', 'days_back', fallback='30'))
                self.window['-START_DATE-'].update(disabled=True)
                self.window['-END_DATE-'].update(disabled=True)
                
            # Storage settings
            self.window['-OUTPUT_DIR-'].update(
                self.config.get('storage', 'output_dir', fallback='output'))
            self.window['-STORAGE_TYPE-'].update(
                self.config.get('storage', 'type', fallback='sqlite').lower())
            self.window['-DB_FILENAME-'].update(
                self.config.get('storage', 'db_filename', fallback='emails.db'))
            
            # Handle JSON export
            export_json = self.config.get_boolean('storage', 'json_export', fallback=True)
            self.window['-EXPORT_JSON-'].update(export_json)
            self.window['-JSON_FILENAME-'].update(
                self.config.get('storage', 'json_filename', fallback='emails.json'))
            self.window['-JSON_FILENAME-'].update(disabled=not export_json)
                
            # Threading settings
            self.window['-ENABLE_THREADING-'].update(
                self.config.getboolean('threading', 'enable_threading', fallback=True))
            self.window['-THREAD_METHOD-'].update(
                self.config.get('threading', 'thread_method', fallback='hybrid'))
            self.window['-MAX_THREAD_DEPTH-'].update(
                self.config.getint('threading', 'max_thread_depth', fallback=10))
            self.window['-THREAD_TIMEOUT_DAYS-'].update(
                self.config.getint('threading', 'thread_timeout_days', fallback=30))
                
            # Email processing
            self.window['-EXTRACT_ATTACHMENTS-'].update(
                self.config.getboolean('email_processing', 'extract_attachments', fallback=False))
            self.window['-ATTACHMENT_DIR-'].update(
                self.config.get('email_processing', 'attachment_dir', fallback='attachments'))
            self.window['-EXTRACT_IMAGES-'].update(
                self.config.getboolean('email_processing', 'extract_embedded_images', fallback=False))
            self.window['-IMAGE_DIR-'].update(
                self.config.get('email_processing', 'image_dir', fallback='images'))
            self.window['-EXTRACT_LINKS-'].update(
                self.config.getboolean('email_processing', 'extract_links', fallback=True))
            self.window['-EXTRACT_PHONES-'].update(
                self.config.getboolean('email_processing', 'extract_phone_numbers', fallback=True))
                
            # Security
            self.window['-REDACT_SENSITIVE-'].update(
                self.config.getboolean('security', 'redact_sensitive_data', fallback=True))
            
            # Handle redaction patterns (convert list to newline-separated string)
            redaction_patterns = self.config.getlist('security', 'redaction_patterns', 
                                                   fallback=['password', 'ssn', 'credit.?card'])
            self.window['-REDACTION_PATTERNS-'].update('\n'.join(redaction_patterns))
            
            # Logging
            self.window['-LOG_LEVEL-'].update(
                self.config.get('logging', 'log_level', fallback='INFO'))
            self.window['-LOG_FILE-'].update(
                self.config.get('logging', 'log_file', fallback='outlook_extractor.log'))
                
        except Exception as e:
            sg.popup_error(f'Error loading configuration: {str(e)}', title='Error')
    
    def _save_ui_to_config(self):
        """Save UI values back to the config."""
        if not self.window:
            return
            
        try:
            values = self.window.read(timeout=100)[1]
            
            # Outlook settings
            if '-MAILBOX-' in values:
                self.config.config['outlook']['mailbox_name'] = values['-MAILBOX-'].strip()
            
            # Handle folder patterns - ensure proper formatting
            if '-FOLDER_PATTERNS-' in values:
                folder_patterns = [p.strip() for p in values['-FOLDER_PATTERNS-'].split(',') 
                                 if p.strip()]
                self.config.config['outlook']['folder_patterns'] = ','.join(folder_patterns)
            
            # Date range
            if '-DATE_RANGE_DAYS-' in values and values['-DATE_RANGE_DAYS-']:
                self.config.config['date_range']['days_back'] = str(values.get('-DAYS_BACK-', '30'))
                
            if '-DATE_RANGE_MONTHS-' in values and values['-DATE_RANGE_MONTHS-'] and '-MONTH_YEAR-' in values and values['-MONTH_YEAR-']:
                # Format: MM/YYYY,MM/YYYY
                month_year = values['-MONTH_YEAR-']
                start_month, start_year = month_year.split('/')
                end_month = str(int(start_month) + 1).zfill(2)
                end_year = start_year
                if int(start_month) == 12:  # Handle December
                    end_month = '01'
                    end_year = str(int(start_year) + 1)
                date_ranges = f"{start_month}/{start_year},{end_month}/{end_year}"
                self.config.config['date_range']['date_ranges'] = date_ranges
            
            # Storage settings
            if '-OUTPUT_DIR-' in values:
                self.config.config['storage']['output_dir'] = values['-OUTPUT_DIR-'].strip()
            if '-DB_FILENAME-' in values:
                self.config.config['storage']['db_filename'] = values['-DB_FILENAME-'].strip()
            if '-JSON_EXPORT-' in values:
                self.config.config['storage']['json_export'] = '1' if values['-JSON_EXPORT-'] else '0'
            if '-JSON_PRETTY-' in values:
                self.config.config['storage']['json_pretty_print'] = '1' if values['-JSON_PRETTY-'] else '0'
            
            # Email processing
            if '-EXTRACT_ATTACHMENTS-' in values:
                self.config.config['email_processing']['extract_attachments'] = '1' if values['-EXTRACT_ATTACHMENTS-'] else '0'
            if '-ATTACHMENT_DIR-' in values:
                self.config.config['email_processing']['attachment_dir'] = values['-ATTACHMENT_DIR-'].strip()
            if '-EXTRACT_IMAGES-' in values:
                self.config.config['email_processing']['extract_embedded_images'] = '1' if values['-EXTRACT_IMAGES-'] else '0'
            if '-IMAGE_DIR-' in values:
                self.config.config['email_processing']['image_dir'] = values['-IMAGE_DIR-'].strip()
            if '-EXTRACT_LINKS-' in values:
                self.config.config['email_processing']['extract_links'] = '1' if values['-EXTRACT_LINKS-'] else '0'
            if '-EXTRACT_PHONES-' in values:
                self.config.config['email_processing']['extract_phone_numbers'] = '1' if values['-EXTRACT_PHONES-'] else '0'
            
            # Threading
            if '-ENABLE_THREADING-' in values:
                self.config.config['threading']['enable_threading'] = '1' if values['-ENABLE_THREADING-'] else '0'
            if '-THREAD_METHOD-' in values:
                self.config.config['threading']['thread_method'] = values['-THREAD_METHOD-']
            if '-MAX_THREAD_DEPTH-' in values:
                self.config.config['threading']['max_thread_depth'] = str(values['-MAX_THREAD_DEPTH-'])
            if '-THREAD_TIMEOUT_DAYS-' in values:
                self.config.config['threading']['thread_timeout_days'] = str(values['-THREAD_TIMEOUT_DAYS-'])
            
            # Security
            if '-REDACT_SENSITIVE-' in values:
                self.config.config['security']['redact_sensitive_data'] = '1' if values['-REDACT_SENSITIVE-'] else '0'
            
            # Handle redaction patterns
            if '-REDACTION_PATTERNS-' in values:
                if values['-REDACTION_PATTERNS-']:
                    patterns = [p.strip() for p in values['-REDACTION_PATTERNS-'].split(',') 
                               if p.strip()]
                else:
                    patterns = ['password', 'ssn', 'credit.?card']
                self.config.config['security']['redaction_patterns'] = ','.join(patterns)
            
            # Logging
            if '-LOG_LEVEL-' in values:
                self.config.config['logging']['log_level'] = values['-LOG_LEVEL-']
            if '-LOG_FILE-' in values:
                self.config.config['logging']['log_file'] = values['-LOG_FILE-'].strip()
            
        except Exception as e:
            sg.popup_error(f'Error saving configuration: {str(e)}', title='Error')
            raise
    
    def backup_data(self, backup_json=True, backup_sqlite=True):
        """Backup the data files.
        
        Args:
            backup_json: Whether to backup JSON files
            backup_sqlite: Whether to backup SQLite database
            
        Returns:
            tuple: (success, message)
        """
        backup_dir = os.path.join('backups', datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
        os.makedirs(backup_dir, exist_ok=True)
        
        success = True
        message = 'Backup completed successfully.'
        
        try:
            if backup_json and self.config.getboolean('storage', 'json_export', fallback=True):
                src = os.path.join(
                    self.config.get('storage', 'output_dir', fallback='output'),
                    self.config.get('storage', 'json_filename', fallback='emails.json')
                )
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(backup_dir, os.path.basename(src)))
            
            if backup_sqlite and self.config.get('storage', 'type', fallback='sqlite') == 'sqlite':
                src = os.path.join(
                    self.config.get('storage', 'output_dir', fallback='output'),
                    self.config.get('storage', 'db_filename', fallback='emails.db')
                )
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(backup_dir, os.path.basename(src)))
        except Exception as e:
            success = False
            message = f'Error during backup: {str(e)}'
        
        return success, message
    
    @log_errors()
    def run_extraction(
        self, 
        folder_patterns: Optional[List[str]] = None, 
        start_date: Optional[datetime] = None, 
        end_date: Optional[datetime] = None, 
        values: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Run the email extraction process.
        
        This method handles the entire extraction workflow including:
        - Validating input parameters
        - Running the email extraction
        - Optionally exporting to CSV
        - Updating the UI with progress and results
        
        Args:
            folder_patterns: List of folder patterns to extract from. If None, uses from config.
            start_date: Start date for extraction. If None, calculates based on config.
            end_date: End date for extraction. If None, uses current time.
            values: Dictionary of UI values for export settings. If None, no export is performed.
            
        Returns:
            Dictionary containing extraction results with keys:
                - success: Boolean indicating if extraction was successful
                - emails_processed: Number of emails processed
                - emails_saved: Number of emails saved
                - folders_processed: Number of folders processed
                - error: Error message if extraction failed
        """
        self.logger.info('Starting email extraction')
        
        # Initialize default result
        result = {
            'success': False,
            'emails_processed': 0,
            'emails_saved': 0,
            'folders_processed': 0,
            'error': None
        }
        
        try:
            # Update UI state
            if hasattr(self, 'window') and self.window:
                self.window['-STATUS-'].update('Running extraction...')
                self.window['-RUN-'].update(disabled=True)
                self.window.refresh()
            
            # Initialize extractor
            from ..extractor import OutlookExtractor  # Lazy import to avoid circular imports
            extractor = OutlookExtractor(config_path=self.config_path)
            
            # Use provided folder patterns or get from config
            if not folder_patterns:
                folder_patterns = self.config.getlist('outlook', 'folder_patterns', ['Inbox'])
            self.logger.info('Using folder patterns: %s', folder_patterns)
                
            # Use provided date range or get from config
            if not start_date or not end_date:
                days_back = self.config.getint('date_range', 'days_back', 30)
                end_date = end_date or datetime.now()
                start_date = start_date or (end_date - timedelta(days=days_back))
            
            self.logger.info('Extracting emails from %s to %s', start_date, end_date)
            
            # Run extraction
            result = extractor.extract_emails(
                folder_patterns=folder_patterns,
                start_date=start_date,
                end_date=end_date
            )
            
            # Handle CSV export if enabled
            if result.get('success') and values and values.get('-EXPORT_CSV-', False):
                self.logger.info('Exporting emails to CSV')
                if hasattr(self, 'window') and self.window:
                    self.window['-STATUS-'].update('Exporting to CSV...')
                    self.window.refresh()
                
                export_settings = {
                    'enable_csv': values.get('-EXPORT_CSV-', True),
                    'output_dir': values.get('-CSV_OUTPUT_DIR-', str(Path.home() / 'email_exports')),
                    'file_prefix': values.get('-CSV_PREFIX-', 'emails_'),
                    'export_basic': values.get('-EXPORT_BASIC-', True),
                    'export_analysis': values.get('-EXPORT_ANALYSIS-', True),
                    'clean_bodies': values.get('-CLEAN_BODIES-', True),
                    'include_summaries': values.get('-INCLUDE_SUMMARIES-', True)
                }
                
                success, output_files = extractor.export_emails(
                    result.get('emails', []),
                    format='csv',
                    export_settings=export_settings
                )
                
                if success and output_files:
                    export_msg = "\n".join([f"- {f}" for f in output_files])
                    self.logger.info('Successfully exported %d files', len(output_files))
                    if hasattr(self, 'window') and self.window:
                        sg.popup_ok(
                            f"Extraction and export completed successfully!\n\n"
                            f"Emails processed: {result.get('emails_processed', 0)}\n"
                            f"Emails saved: {result.get('emails_saved', 0)}\n"
                            f"Folders processed: {result.get('folders_processed', 0)}\n\n"
                            f"Exported files:\n{export_msg}",
                            title='Extraction and Export Complete'
                        )
                else:
                    self.logger.warning('Export failed or no files were exported')
                    if hasattr(self, 'window') and self.window:
                        sg.popup_warning(
                            f"Extraction completed but export failed.\n\n"
                            f"Emails processed: {result.get('emails_processed', 0)}\n"
                            f"Emails saved: {result.get('emails_saved', 0)}",
                            title='Extraction Complete (Export Failed)'
                        )
            
            self.logger.info('Extraction completed successfully')
            return result
            
        except Exception as e:
            error_msg = f'Error during extraction: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            if hasattr(self, 'window') and self.window:
                sg.popup_error(error_msg, title='Error')
            return {
                'success': False,
                'error': error_msg,
                'emails_processed': result.get('emails_processed', 0),
                'emails_saved': result.get('emails_saved', 0),
                'folders_processed': result.get('folders_processed', 0)
            }
            
        finally:
            if hasattr(self, 'window') and self.window:
                self.window['-STATUS-'].update('Ready')
                self.window['-RUN-'].update(disabled=False)
                self.window.refresh()
    
    @log_errors()
    def event_loop(self):
        """
        Run the main event loop for the application with robust error handling.
        
        This method handles the main application loop, processes events, and ensures
        proper cleanup on exit. It includes comprehensive error handling to prevent
        crashes and provide meaningful feedback to the user.
        """
        if not self.window:
            self.logger.error("Cannot start event loop: window not initialized")
            return
            
        self.logger.info("Starting main event loop")
        
        # Track if we're in the process of shutting down
        self._shutting_down = False
        
        try:
            while True:
                try:
                    # Read with a timeout to allow for periodic checks
                    event, values = self.window.read(timeout=100)
                    
                    # Check for window close or exit request
                    if event in (sg.WIN_CLOSED, '-EXIT-', None):
                        self.logger.info("Exit requested by user")
                        if sg.popup_yes_no('Are you sure you want to exit?', title='Exit') == 'Yes':
                            break
                        continue
                        
                    # Skip None events (timeouts)
                    if event is None:
                        continue
                        
                    # Handle special log message events from the GUI handler
                    if isinstance(event, tuple) and len(event) == 2 and event[0] == 'LOG_MESSAGE':
                        try:
                            if hasattr(self.window, 'log') and self.window.log:
                                self.window.log.update(value=event[1], append=True)
                                self.window.log.set_vscroll_position(1.0)  # Auto-scroll
                        except Exception as e:
                            sys.stderr.write(f"Error updating log window: {str(e)}\n")
                        continue
                        
                    # Log the event at debug level to avoid noise
                    self.logger.debug('Event received: %s', event)
                    
                    # Handle the event with error handling
                    try:
                        self._handle_event(event, values)
                    except Exception as e:
                        self.logger.error(
                            'Error handling event %s: %s', 
                            event, 
                            str(e),
                            exc_info=True
                        )
                        
                        # Show error to user with more context
                        try:
                            sg.popup_error(
                                'An error occurred while processing your request.\n\n'
                                f'Error: {str(e)}\n\n'
                                'The application will continue running.\n'
                                'Please check the logs for more details.',
                                title='Error',
                                keep_on_top=True
                            )
                        except Exception as popup_error:
                            self.logger.error(
                                'Failed to show error popup: %s', 
                                str(popup_error),
                                exc_info=True
                            )
                            
                except KeyboardInterrupt:
                    self.logger.info("Keyboard interrupt received, shutting down...")
                    break
                    
                except Exception as e:
                    self.logger.critical(
                        'Unexpected error in event loop: %s', 
                        str(e),
                        exc_info=True
                    )
                    
                    # Try to recover by showing error to user
                    try:
                        if sg.popup_yes_no(
                            'An unexpected error occurred in the application.\n\n'
                            f'Error: {str(e)}\n\n'
                            'Do you want to continue running the application?\n'
                            "(Selecting 'No' will close the application)",
                            title='Unexpected Error',
                            keep_on_top=True
                        ) != 'Yes':
                            break
                    except Exception as popup_error:
                        self.logger.error(
                            'Failed to show recovery popup: %s', 
                            str(popup_error),
                            exc_info=True
                        )
                        # If we can't show the popup, it's better to exit
                        break
                    except Exception:
                        pass  # If we can't show the error, at least we logged it
                        
        except Exception as e:
            self.logger.critical('Fatal error in main event loop', exc_info=True)
            raise
        try:
            self.logger.debug('Stopping background tasks...')
            # Add any background task cleanup here
            # Example: self.background_thread.stop() if hasattr(self, 'background_thread') else None
            pass
        except Exception as e:
            self.logger.error(error_msg, exc_info=True)
            cleanup_errors.append(error_msg)
            
    def _handle_event(self, event: str, values: Dict[str, Any]) -> None:
        """Handle UI events.
        
        Args:
            event: The event that occurred
            values: Dictionary of UI element values
        """
        try:
            if event == 'Save Config':
                save_path = sg.popup_get_file(
                    'Save Config As',
                    save_as=True,
                    default_extension='.ini',
                    file_types=(('INI Files', '*.ini'),)
                )
                
                if save_path:
                    try:
                        self._save_ui_to_config()
                        # Save the config using the ConfigManager instance method
                        success = self.config.save_config(save_path)
                        if success:
                            sg.popup_ok('Config saved successfully!', title='Success')
                        else:
                            sg.popup_error('Failed to save config file.', title='Error')
                    except Exception as e:
                        sg.popup_error(f'Error saving config: {str(e)}', title='Error')
            
            elif event == 'Backup Config':
                backup_path = sg.popup_get_file('Backup Config As', 
                                            save_as=True, 
                                            default_extension='.ini',
                                            file_types=(('INI Files', '*.ini'),))
                if backup_path:
                    try:
                        shutil.copy2(self.config_path, backup_path)
                        sg.popup_ok('Config backed up successfully!', title='Success')
                    except Exception as e:
                        sg.popup_error(f'Error backing up config: {str(e)}', title='Error')

            elif event == 'Backup Data':
                layout = [
                    [sg.Text('Select data to backup:')],
                    [sg.Checkbox('JSON Export', default=True, key='-BACKUP_JSON-')],
                    [sg.Checkbox('SQLite Database', default=True, key='-BACKUP_SQLITE-')],
                    [sg.Button('Backup'), sg.Button('Cancel')]
                ]

                backup_window = sg.Window('Backup Data', layout, modal=True)
                while True:
                    event, values = backup_window.read()
                    if event in (sg.WIN_CLOSED, 'Cancel'):
                        break
                    elif event == 'Backup':
                        success, message = self.backup_data(
                            values['-BACKUP_JSON-'], 
                            values['-BACKUP_SQLITE-']
                        )
                        if success:
                            sg.popup_ok('Backup completed successfully!', title='Success')
                        else:
                            sg.popup_error(f'Backup failed: {message}', title='Error')
                        break
                backup_window.close()

            # Handle tab changes
            elif event == '-TAB_GROUP-':
                active_tab = values['-TAB_GROUP-']
                self.logger.debug(f'Tab changed to: {active_tab}')
                
                # Lazy load the export tab when it's selected
                if active_tab == '-EXPORT_TAB-':
                    self._load_export_tab()
                
            # Handle theme changes
            elif event == 'Theme::Light':
                self.theme = 'LightGrey1'
                self.setup_theme()
                if hasattr(self, 'window') and self.window:
                    self.window['-THEME-'].update('Light')
            elif event == 'Theme::Dark':
                self.theme = 'DarkGrey9'
                self.setup_theme()
                if hasattr(self, 'window') and self.window:
                    self.window['-THEME-'].update('Dark')
                    
            # Handle documentation links
            elif event in ('Documentation', '-DOCS_LINK-'):
                import webbrowser
                webbrowser.open('https://github.com/yourusername/outlook-extractor')
                
            # Handle export tab events
            if hasattr(self, 'export_tab') and self.export_tab is not None:
                if event.startswith('-EXPORT_') or event.startswith('Export'):
                    try:
                        self.export_tab.handle_event(event, values)
                    except Exception as e:
                        self.logger.error(f'Error handling export event {event}: {str(e)}', exc_info=True)
                        sg.popup_error(f'Error in export tab: {str(e)}', title='Export Error')
            
            # Handle run button click
            if event == '-RUN-':
                self._handle_run_event(values)
            
        except Exception as e:
            self.logger.error(f'Error handling event {event}: {str(e)}', exc_info=True)
            sg.popup_error(f'Error: {str(e)}', title='Error')

    def _load_export_tab(self):
        """Lazy load the export tab content when the tab is first selected."""
        try:
            if self.export_tab is None:
                self.logger.debug('Loading export tab content...')
                from outlook_extractor.ui.export_tab import ExportTab
                self.export_tab = ExportTab(self.config)
                
                # Replace the loading placeholder with the actual content
                if hasattr(self, 'window') and self.window:
                    self.window['-EXPORT_LOADING-'].update(visible=False)
                    self.window.extend_layout(
                        self.window['-EXPORT_TAB-'],
                        self.export_tab.get_layout()
                    )
                    self.window.refresh()
                    self.logger.debug('Export tab content loaded successfully')
        except Exception as e:
            self.logger.error(f'Error loading export tab: {str(e)}', exc_info=True)
            sg.popup_error(f'Error loading export tab: {str(e)}', title='Error')

    def _handle_documentation_link(self, event):
        """Handle documentation link clicks."""
        if event in ('Documentation', '-DOCS_LINK-'):
            import webbrowser
            webbrowser.open('https://github.com/yourusername/outlook-extractor')
            
    def _check_for_updates(self, silent: bool = False) -> None:
        """Check for application updates.
        
        Args:
            silent: If True, don't show any UI if no update is available
        """
        try:
            # Don't check if we're already checking
            if self._update_checked:
                return
                
            from .. import check_for_updates
            
            # Get the window if it exists, otherwise use None
            parent_window = None
            if hasattr(self, 'window') and self.window is not None:
                parent_window = self.window
                self.logger.info("Window is available for update dialog")
            else:
                self.logger.warning("Window not available for update dialog")
                return  # Can't show UI without a window
                
            # Mark that we've checked for updates
            self._update_checked = True
            self._last_update_check = time.time()
                
            check_for_updates(
                parent_window=parent_window,
                repo_owner="bbleak-repo",
                repo_name="outlook-extractor",
                silent=silent
            )
        except Exception as e:
            self.logger.error(f"Error checking for updates: {e}", exc_info=True)
            if not silent and hasattr(self, 'window') and self.window is not None:
                sg.popup_error(f"Error checking for updates: {e}", title="Update Error")
            
    def run(self) -> None:
        """Run the main application loop."""
        try:
            # Start the event loop
            self.event_loop()
            
        except Exception as e:
            self.logger.critical("Fatal error in application", exc_info=True)
            sg.popup_error(
                "A fatal error occurred and the application must close.\n\n"
                f"Error: {str(e)}\n\n"
                "Please check the application logs for more details.",
                title="Fatal Error"
            )
        finally:
            self._cleanup()
            
    def _cleanup(self) -> None:
        """
        Clean up resources before exiting.
        
        This method ensures all resources are properly released, configurations are saved,
        and the application shuts down gracefully. It includes comprehensive error handling
        to ensure cleanup completes as much as possible even if some operations fail.
        """
        if hasattr(self, '_shutting_down') and self._shutting_down:
            self.logger.debug('Cleanup already in progress, skipping duplicate call')
            return
            
        self._shutting_down = True
        self.logger.info('Starting application cleanup...')
        
        # Track cleanup status
        cleanup_errors = []
        
        # 1. Save any pending configuration changes
        try:
            self.logger.debug('Saving configuration...')
            self._save_ui_to_config()
            self.logger.info('Configuration saved successfully')
        except Exception as e:
            error_msg = f'Failed to save configuration: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            cleanup_errors.append(error_msg)
        
        # 2. Stop any running background tasks or threads
        try:
            self.logger.debug('Stopping background tasks...')
            # Add any background task cleanup here
            # Example: self.background_thread.stop() if hasattr(self, 'background_thread') else None
            pass
        except Exception as e:
            error_msg = f'Error stopping background tasks: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            cleanup_errors.append(error_msg)
        
        # 3. Close the window and clean up GUI resources
        if hasattr(self, 'window') and self.window:
            try:
                self.logger.debug('Closing main window...')
                self.window.close()
                self.window = None
                self.logger.info('Main window closed')
            except Exception as e:
                error_msg = f'Error closing window: {str(e)}'
                self.logger.error(error_msg, exc_info=True)
                cleanup_errors.append(error_msg)
        
        # 4. Clean up any other resources
        try:
            self.logger.debug('Cleaning up additional resources...')
            # Add any additional cleanup code here
            # Example: self.database_connection.close() if hasattr(self, 'database_connection') else None
            pass
        except Exception as e:
            error_msg = f'Error during resource cleanup: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            cleanup_errors.append(error_msg)
        
        # 5. Flush all log handlers to ensure all messages are written
        try:
            import logging
            for handler in logging.root.handlers[:]:
                try:
                    handler.flush()
                    if hasattr(handler, 'close'):
                        handler.close()
                except Exception as e:
                    self.logger.error(f'Error closing log handler {handler}: {str(e)}', exc_info=True)
        except Exception as e:
            self.logger.error(f'Error flushing logs: {str(e)}', exc_info=True)
        
        # Log completion status
        if cleanup_errors:
            error_summary = '\n- '.join(cleanup_errors)
            self.logger.warning(f'Cleanup completed with {len(cleanup_errors)} error(s):\n- {error_summary}')
        else:
            self.logger.info('Cleanup completed successfully')
        
        # Final log message
        self.logger.info('Application shutdown complete')
        
        # Ensure all handlers are closed
        logging.shutdown()
    
    def _handle_run_event(self, values: Dict[str, Any]) -> None:
        """Handle the run button click event.
        
        Args:
            values: Dictionary containing the current values of all input elements
            
        This method is called when the user clicks the run button. It validates the input,
        processes the folder patterns and date range, and starts the extraction in a separate thread.
        """
        try:
            # Get and validate folder patterns
            folder_patterns = values.get('-FOLDER_PATTERNS-', '').split(',')
            folder_patterns = [p.strip() for p in folder_patterns if p.strip()]
            
            if not folder_patterns:
                sg.popup_error('Please specify at least one folder pattern', title='Error')
                return
                
            # Get and validate date range
            if values.get('-DATE_RANGE_LAST_N_DAYS-', True):
                try:
                    days = int(values.get('-LAST_N_DAYS-', '30'))
                    if days <= 0:
                        raise ValueError('Number of days must be positive')
                    start_date = datetime.now() - timedelta(days=days)
                    end_date = datetime.now()
                except ValueError as e:
                    sg.popup_error(f'Invalid date range: {str(e)}', title='Error')
                    return
            else:
                start_date = values.get('-START_DATE-')
                end_date = values.get('-END_DATE-')
                
                if not start_date or not end_date:
                    sg.popup_error('Please select both start and end dates', title='Error')
                    return
                    
                if start_date > end_date:
                    sg.popup_error('Start date must be before end date', title='Error')
                    return
            
            # Run extraction in a separate thread
            self.logger.info(f"Starting extraction for patterns: {folder_patterns}")
            
            def run_async():
                try:
                    self.run_extraction(
                        folder_patterns=folder_patterns,
                        start_date=start_date,
                        end_date=end_date,
                        values=values
                    )
                except Exception as e:
                    self.logger.error(f"Error in extraction thread: {str(e)}", exc_info=True)
                    if hasattr(self, 'window'):
                        sg.popup_error(
                            f'Extraction failed: {str(e)}',
                            title='Extraction Error',
                            keep_on_top=True
                        )
            
            # Start the extraction in a separate thread
            thread = threading.Thread(target=run_async, daemon=True)
            thread.start()
            
        except Exception as e:
            self.logger.error(f"Error in run handler: {str(e)}", exc_info=True)
            if hasattr(self, 'window'):
                sg.popup_error(
                    f'Failed to start extraction: {str(e)}',
                    title='Error',
                    keep_on_top=True
                )
            
            # Handle Load Config menu item
            if event == 'Load Config':
                config_file = sg.popup_get_file('Load Config', 
                                             file_types=(('INI Files', '*.ini'),))
                if config_file:
                    try:
                        self.config = get_config(config_file)
                        self.config_path = config_file
                        self._load_config_to_ui()
                        sg.popup_ok('Config loaded successfully!', title='Success')
                    except Exception as e:
                        sg.popup_error(f'Error loading config: {str(e)}', title='Error')
            
            elif event == 'Save Config':
                save_path = sg.popup_get_file('Save Config As', 
                                            save_as=True, 
                                            default_extension='.ini',
                                            file_types=(('INI Files', '*.ini'),))
                if save_path:
                    try:
                        self._save_ui_to_config()
                        # Save the config using the ConfigManager instance method
                        success = self.config.save_config(save_path)
                        if success:
                            sg.popup_ok('Config saved successfully!', title='Success')
                        else:
                            sg.popup_error('Failed to save config file.', title='Error')
                    except Exception as e:
                        sg.popup_error(f'Error saving config: {str(e)}', title='Error')
            
            elif event == 'Backup Config':
                backup_path = sg.popup_get_file('Backup Config As', 
                                              save_as=True, 
                                              default_extension='.ini',
                                              file_types=(('INI Files', '*.ini'),))
                if backup_path:
                    try:
                        shutil.copy2(self.config_path, backup_path)
                        sg.popup_ok('Config backed up successfully!', title='Success')
                    except Exception as e:
                        sg.popup_error(f'Error backing up config: {str(e)}', title='Error')
            
            elif event == 'Backup Data':
                layout = [
                    [sg.Text('Select data to backup:')],
                    [sg.Checkbox('JSON Export', default=True, key='-BACKUP_JSON-')],
                    [sg.Checkbox('SQLite Database', default=True, key='-BACKUP_SQLITE-')],
                    [sg.Button('Backup'), sg.Button('Cancel')]
                ]
                
                backup_window = sg.Window('Backup Data', layout, modal=True)
                while True:
                    event, values = backup_window.read()
                    if event in (sg.WIN_CLOSED, 'Cancel'):
                        break
                    elif event == 'Backup':
                        success, message = self.backup_data(
                            values['-BACKUP_JSON-'], 
                            values['-BACKUP_SQLITE-']
                        )
                        sg.popup_ok(message, title='Backup ' + ('Succeeded' if success else 'Failed'))
                        break
                backup_window.close()
            
            elif event == 'Options':
                # Simple options dialog
                layout = [
                    [sg.Text('Theme:'), 
                     sg.Combo(sg.theme_list(), default_value=self.theme, key='-THEME-')],
                    [sg.Button('Apply'), sg.Button('Cancel')]
                ]
                
                options_window = sg.Window('Options', layout, modal=True)
                while True:
                    event, values = options_window.read()
                    if event in (sg.WIN_CLOSED, 'Cancel'):
                        break
                    elif event == 'Apply':
                        self.theme = values['-THEME-']
                        sg.theme(self.theme)
                        sg.popup_ok('Theme will be applied after restart.', title='Info')
                        break
                options_window.close()
            
            elif event == 'About':
                sg.popup_ok(
                    'Outlook Email Extractor\n\n'
                    'Version: 1.0.0\n'
                    'A tool to extract and analyze emails from Microsoft Outlook.\n\n'
                    '© 2025 Your Company. All rights reserved.',
                    title='About'
                )
            
            elif event == 'Check for Updates':
                self._check_for_updates()
                
            elif event == 'Documentation':
                import webbrowser
                webbrowser.open('https://github.com/yourusername/outlook-extractor')
            
            elif event == '-DOCS_LINK-':
                import webbrowser
                webbrowser.open('https://github.com/yourusername/outlook-extractor')
        
        self.window.close()

def main() -> None:
    """Main entry point for the UI."""
    import sys
    
    # Set up basic logging first
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    logger = logging.getLogger(__name__)
    
    try:
        # Set up advanced logging if available
        try:
            setup_logging()
            logger = get_logger(__name__)
        except Exception as e:
            logger.warning('Could not set up advanced logging, using basic logging', exc_info=True)
        
        # Parse command line arguments
        config_path = sys.argv[1] if len(sys.argv) > 1 else None
        
        # Configure PySimpleGUI settings - only set basic options here
        sg.set_options(
            element_padding=(4, 1),
            text_justification='left',
            border_width=1,
            # Make sure these don't conflict with setup_theme
            auto_size_buttons=True
        )
        
        # Create and run the application
        logger.info('Starting Outlook Extractor')
        app = EmailExtractorUI(config_path)
        app.run()
        
    except Exception as e:
        logger.critical('Fatal error in main', exc_info=True)
        
        # Try to show error in a GUI dialog
        try:
            sg.popup_error(
                'A fatal error occurred and the application must close.\n\n'
                f'Error: {str(e)}\n\n'
                'Please check the application logs for more details.',
                title='Fatal Error',
                keep_on_top=True
            )
        except Exception as gui_error:
            print(f'Failed to show error dialog: {str(gui_error)}', file=sys.stderr)
        
        sys.exit(1)

if __name__ == '__main__':
    main()
