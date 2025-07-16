# outlook_extractor/ui/export_tab.py

import logging
import os
import threading
import uuid
import webbrowser
from datetime import datetime, timezone
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Lock
from typing import Any, Callable, Dict, List, Optional, Set, Tuple, Union

import PySimpleGUI as sg

from ..export.presets import ExportPreset, ExportPresetManager

class ExportTab:
    """Handles the export options tab in the UI."""

    def _show_error_popup(self, message: str, title: str = 'Error', exc_info: Optional[Exception] = None) -> None:
        """Show an error popup and log the error.

        Args:
            message: The error message to display
            title: The title of the error popup
            exc_info: Optional exception object for logging
        """
        self.logger.error(f'{title}: {message}', exc_info=exc_info)
        if self.window:
            try:
                sg.popup_error(message, title=title, icon=sg.ICON_ERROR, keep_on_top=True)
            except Exception as e:
                self.logger.critical(f'Failed to show error popup: {str(e)}', exc_info=True)
                # Fallback to console if UI fails
                print(f'ERROR [{title}]: {message}')

    def _show_info_popup(self, message: str, title: str = 'Information', non_blocking: bool = False) -> None:
        """Show an information popup and log the message.

        Args:
            message: The message to display
            title: The title of the popup
            non_blocking: If True, the popup won't block the UI
        """
        self.logger.info(f'{title}: {message}')
        if self.window:
            try:
                if non_blocking:
                    # Use a separate thread to show the popup without blocking
                    def show_popup():
                        try:
                            sg.popup_ok(message, title=title, icon=sg.ICON_INFORMATION, non_blocking=True, keep_on_top=True)
                        except Exception as e:
                            self.logger.error(f'Error in non-blocking popup: {str(e)}')
                    
                    popup_thread = threading.Thread(target=show_popup, daemon=True)
                    popup_thread.start()
                else:
                    sg.popup_ok(message, title=title, icon=sg.ICON_INFORMATION, keep_on_top=True)
            except Exception as e:
                self.logger.error(f'Failed to show info popup: {str(e)}', exc_info=True)
                print(f'INFO [{title}]: {message}')

    def _show_warning_popup(self, message: str, title: str = 'Warning') -> None:
        """Show a warning popup and log the warning.

        Args:
            message: The warning message to display
            title: The title of the warning popup
        """
        self.logger.warning(f'{title}: {message}')
        if self.window:
            try:
                sg.popup_ok(message, title=title, icon=sg.ICON_WARNING, keep_on_top=True)
            except Exception as e:
                self.logger.error(f'Failed to show warning popup: {str(e)}', exc_info=True)
                print(f'WARNING [{title}]: {message}')
                
    def _show_info(self, message: str, title: str = 'Information') -> None:
        """Show an information popup.
        
        Args:
            message: The message to display
            title: The window title
        """
        self._show_info_popup(message, title)
    
    def _show_error(self, message: str, title: str = 'Error') -> None:
        """Show an error popup.
        
        Args:
            message: The error message to display
            title: The window title
        """
        self._show_error_popup(message, title)

    def __init__(self, config: Dict[str, Any] = None):
        """Initialize the export tab.

        Args:
            config: Optional configuration dictionary. If None, default values will be used.
        """
        # Configure logging first
        self.logger = logging.getLogger(f'{__name__}.{self.__class__.__name__}')
        self.logger.info('=' * 50)
        self.logger.info('Initializing ExportTab')
        
        # Threading and state management
        self._export_thread = None
        self._cancel_event = Event()
        self._export_lock = Lock()
        self._export_in_progress = False
        
        # Preset management
        self._preset_manager = ExportPresetManager()
        self._preset_manager.ensure_default_presets_exist()
        self._current_preset_id = None

        try:
            self.logger.debug(f'Received config: {config}')
            self.config = config or {}
            self.window = None  # Will be set when the UI is created
            self._folder_patterns = []
            self.layout = []

            # Initialize UI components
            self.logger.debug('Initializing UI components')
            self._init_ui()

            if not self.layout:
                raise ValueError('UI layout was not properly initialized')

            self.logger.info('ExportTab initialized successfully')
            self.logger.debug(f'Initial layout: {self.layout}')

        except Exception as e:
            error_msg = f'Failed to initialize ExportTab: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            
            # Create a simple error layout that will be shown in the tab
            error_details = f'Error details: {str(e)}\n\nPlease check the logs for more information.'
            self.logger.error(f'Error details: {error_details}')
            
            self.layout = [
                [sg.Text('⚠️ Error initializing export tab', font=('Helvetica', 12, 'bold'), text_color='red')],
                [sg.Multiline(
                    error_details,
                    size=(80, 10),
                    disabled=True,
                    text_color='red',
                    background_color='#FFEBEE'
                )],
                [sg.Text('Check the application logs for more details.', text_color='orange')]
            ]

            # Log the final layout for debugging
            self.logger.debug(f'Error layout: {self.layout}')
            # Re-raise to ensure the error is not silently ignored
            raise RuntimeError(error_msg) from e

    def update_folder_patterns(self, folder_patterns: List[str]) -> None:
        """Update the folder patterns and refresh the generated filename.

        Args:
            folder_patterns: List of folder patterns to extract from
        """
        try:
            self.logger.info(f'Updating folder patterns: {folder_patterns}')
            if not isinstance(folder_patterns, list):
                self.logger.warning(f'Expected list of folder patterns, got: {type(folder_patterns)}')
                folder_patterns = [str(folder_patterns)]

            self._folder_patterns = folder_patterns

            if not folder_patterns:
                self.logger.warning('No folder patterns provided')
                return

            self.logger.debug(f'Calling _update_generated_filename with patterns: {folder_patterns}')
            self._update_generated_filename()

        except Exception as e:
            self.logger.error(f'Error updating folder patterns: {str(e)}', exc_info=True)

    def _generate_filename(self) -> str:
        """Generate a filename based on folder patterns and current timestamp.

        Returns:
            str: Generated filename with timestamp
        """
        if not self._folder_patterns:
            logging.debug('No folder patterns, using default filename')
            return 'emails_export.csv'

        try:
            # Clean folder names for use in filenames
            clean_names = []
            for pattern in self._folder_patterns:
                # Remove special characters and replace spaces with underscores
                clean = ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_' for c in pattern)
                clean = clean.replace(' ', '_').strip('_')
                if clean:  # Only add non-empty names
                    clean_names.append(clean)

            # Limit the number of folders in the filename to avoid it getting too long
            max_folders = 3
            if len(clean_names) > max_folders:
                clean_names = clean_names[:max_folders] + [f"and_{len(clean_names) - max_folders}_more"]

            # Add timestamp
            from datetime import datetime
            timestamp = datetime.now().strftime('%Y%m%d_%H%M')

            # Combine into filename
            base = '_'.join(clean_names)
            filename = f"{base}_{timestamp}.csv"

            logging.debug(f'Generated filename: {filename}')
            return filename

        except Exception as e:
            logging.error(f'Error generating filename: {str(e)}')
            return 'emails_export.csv'

    def _update_generated_filename(self) -> None:
        """Update the filename field with the generated filename."""
        if not self.window:
            logging.warning('Cannot update filename: window not initialized')
            return

        try:
            filename = self._generate_filename()
            logging.debug(f'Updating filename field to: {filename}')
            self.window['-CSV_PREFIX-'].update(filename)
        except Exception as e:
            logging.error(f'Error updating filename field: {str(e)}')

    def _get_config_value(self, section: str, key: str, default: Any = None) -> Any:
        """Safely get a value from the config with nested dictionary support."""
        if not self.config:
            return default

        # Handle nested dictionary access (e.g., 'export.enable_csv')
        keys = section.split('.')
        value = self.config

        try:
            for k in keys:
                value = value.get(k, {})
            return value.get(key, default) if isinstance(value, dict) else default
        except (AttributeError, TypeError):
            return default

    def _validate_layout_structure(self, layout, section_name="layout"):
        """Validate that layout follows PySimpleGUI structure requirements."""
        try:
            if not isinstance(layout, list):
                raise ValueError(f"{section_name} must be a list")

            for i, row in enumerate(layout):
                if not isinstance(row, list):
                    raise ValueError(f"{section_name} row {i} must be a list, got {type(row)}")

                for j, element in enumerate(row):
                    # Check if element is a PySimpleGUI element or valid container
                    if hasattr(element, '__class__') and 'PySimpleGUI' in str(element.__class__):
                        continue  # Valid PySimpleGUI element
                    elif isinstance(element, str):
                        continue  # String elements are valid
                    elif element is None:
                        self.logger.warning(f"None element found in {section_name} row {i}, col {j}")
                    else:
                        self.logger.warning(f"Unexpected element type in {section_name} row {i}, col {j}: {type(element)}")

            self.logger.debug(f"{section_name} validation passed")
            return True

        except Exception as e:
            self.logger.error(f"Layout validation failed for {section_name}: {str(e)}")
            return False

    def _init_ui(self):
        """Initialize the export tab UI elements."""
        self.logger.info('-' * 50)
        self.logger.info('Initializing ExportTab UI')

        try:
            # Get default values from config or use sensible defaults
            self.logger.debug('Loading configuration values...')
            try:
                enable_csv = self._get_config_value('export', 'enable_csv', True)
                output_dir = self._get_config_value('export', 'output_dir', str(Path.home() / 'email_exports'))
                clean_bodies = self._get_config_value('export', 'clean_bodies', True)
                include_summaries = self._get_config_value('export', 'include_summaries', True)

                self.logger.debug(f'Config loaded - enable_csv: {enable_csv}, output_dir: {output_dir}, '
                                f'clean_bodies: {clean_bodies}, include_summaries: {include_summaries}')

            except Exception as e:
                self.logger.warning(f'Error reading config: {str(e)}. Using default values.', exc_info=True)
                enable_csv = True
                output_dir = str(Path.home() / 'email_exports')
                clean_bodies = True
                include_summaries = True

            # Ensure output directory exists
            try:
                self.logger.debug(f'Verifying output directory: {output_dir}')
                output_path = Path(output_dir)
                output_path.mkdir(parents=True, exist_ok=True)
                self.logger.debug(f'Output directory verified/created: {output_path.absolute()}')
            except Exception as e:
                error_msg = f'Could not create output directory {output_dir}: {str(e)}. Using home directory.'
                self.logger.warning(error_msg, exc_info=True)
                output_dir = str(Path.home())
                self.logger.info(f'Using fallback output directory: {output_dir}')

            try:
                self.logger.debug('Creating UI layout...')
                try:
                    self.logger.debug('Creating CSV export section...')
                    # CSV export options - Simplified layout for better macOS compatibility
                    csv_section = [
                        [sg.Checkbox('Enable CSV Export',
                                   default=enable_csv,
                                   key='-EXPORT_CSV-',
                                   enable_events=True,
                                   pad=(10, (10, 5)))
                        ],
                        [sg.Text('Output Directory:', size=(15, 1), pad=(10, (10, 5))),
                         sg.In(default_text=output_dir,
                               key='-CSV_OUTPUT_DIR-',
                               disabled=not enable_csv,
                               size=(45, 1),
                               pad=(0, (10, 5))),
                         sg.FolderBrowse(disabled=not enable_csv,
                                      key='-CSV_BROWSE-',
                                      initial_folder=output_dir,
                                      size=(10, 1),
                                      pad=(5, (10, 5)))
                        ],
                        [sg.Text('Filename:', size=(15, 1), pad=(10, (5, 10))),
                         sg.In(default_text='emails_export.csv',
                               key='-CSV_PREFIX-',
                               disabled=not enable_csv,
                               size=(45, 1),
                               pad=(0, (5, 10))),
                         sg.Button('Generate',
                                 key='-GENERATE_FILENAME-',
                                 disabled=not enable_csv,
                                 size=(10, 1),
                                 pad=(5, (5, 10)))
                        ]
                    ]

                    # Validate CSV section layout
                    self.logger.debug('Validating CSV section layout...')
                    if not self._validate_layout_structure(csv_section, "csv_section"):
                        raise ValueError("CSV section layout validation failed")
                    self.logger.debug('CSV export section created and validated successfully')

                    # Export options - Simplified for macOS compatibility
                    self.logger.debug('Creating export options section...')
                    
                    # Get available presets
                    presets = self._preset_manager.get_all_presets()
                    preset_names = [p['name'] for p in presets]
                    
                    # Create preset management UI
                    preset_layout = [
                        [sg.Text('Preset:', size=(15, 1)),
                         sg.Combo(preset_names, 
                                 default_value=preset_names[0] if preset_names else '',
                                 enable_events=True,
                                 key='-PRESET_SELECT-',
                                 size=(30, 1)),
                         sg.Button('Save', key='-SAVE_PRESET-'),
                         sg.Button('Save As...', key='-SAVE_AS_PRESET-'),
                         sg.Button('Delete', key='-DELETE_PRESET-')],
                        [sg.HorizontalSeparator()]
                    ]
                    
                    # Create export format and options
                    export_options = [
                        [sg.Text('Export Options:', 
                               font=('Helvetica', 10, 'bold'), 
                               pad=(10, (10, 5)))],
                        *preset_layout,
                        [sg.Text('Format:', size=(15, 1)),
                         sg.Combo(
                             [f[0] for f in self._get_export_formats()],
                             default_value='CSV (.csv)',
                             key='-EXPORT_FORMAT-',
                             enable_events=True,
                             size=(15, 1),
                             readonly=True
                         )],
                        [sg.Checkbox('Basic Email Data',
                                   default=True,
                                   key='-EXPORT_BASIC-',
                                   disabled=not enable_csv,
                                   pad=(15, (5, 5)))],
                        [sg.Checkbox('Analysis Data',
                                   default=True,
                                   key='-EXPORT_ANALYSIS-',
                                   disabled=not enable_csv,
                                   pad=(15, (5, 5)))],
                        [sg.Checkbox('Clean HTML from Email Bodies',
                                   default=clean_bodies,
                                   key='-CLEAN_BODIES-',
                                   disabled=not enable_csv,
                                   pad=(15, (5, 5)))],
                        [sg.Checkbox('Include AI Summaries',
                                   default=include_summaries,
                                   key='-INCLUDE_SUMMARIES-',
                                   disabled=not enable_csv,
                                   pad=(15, (5, 5)))]
                    ]

                    # Validate export options layout
                    self.logger.debug('Validating export options layout...')
                    if not self._validate_layout_structure(export_options, "export_options"):
                        raise ValueError("Export options layout validation failed")
                    self.logger.debug('Export options section created and validated successfully')

                    # Assemble the main layout with better spacing for macOS
                    self.logger.debug('Assembling main layout...')
                    
                    # Create the main layout as a list of rows
                    main_layout = [
                        # Header row with padding
                        [sg.Text('Export Options', 
                               font=('Helvetica', 16, 'bold'), 
                               pad=(10, (10, 15)))],
                        
                        # CSV Export Settings Frame with padding
                        [sg.Frame('CSV Export Settings',
                                layout=csv_section,
                                expand_x=True,
                                pad=(10, 5),
                                title_location='n',
                                relief=sg.RELIEF_GROOVE)],
                        
                        # Export Options Frame with padding
                        [sg.Frame('Export Options',
                                layout=export_options,
                                expand_x=True,
                                pad=(10, 10),
                                title_location='n',
                                relief=sg.RELIEF_GROOVE)],
                        
                        # Button row with padding and centering
                        [sg.Text('', size=(1, 1))],  # Vertical spacer
                        [sg.HorizontalSeparator()],
                        [sg.Text('', size=(1, 1))],  # Vertical spacer
                        [sg.Push(),
                         sg.Button('Export to CSV',
                                 key='-EXPORT_CSV_BUTTON-',
                                 size=(20, 2),
                                 font=('Helvetica', 11, 'bold'),
                                 button_color=('white', '#0078D7'),
                                 pad=(0, (15, 10)),
                                 disabled=not enable_csv),
                         sg.Push()],
                        [sg.Text('', size=(1, 1))]  # Vertical spacer
                    ]

                    # Validate main layout structure
                    self.logger.debug('Validating main layout structure...')
                    if not self._validate_layout_structure(main_layout, "main_layout"):
                        raise ValueError("Main layout validation failed")

                    self.layout = main_layout
                    self.logger.debug('Main layout assembled and validated successfully')

                except Exception as e:
                    error_msg = f'Error creating UI components: {str(e)}'
                    self.logger.error(error_msg, exc_info=True)
                    raise RuntimeError(error_msg) from e

                self.logger.debug('UI layout created successfully')
                self.logger.info('ExportTab UI initialized successfully')

            except Exception as e:
                error_msg = f'Error creating UI layout: {str(e)}'
                self.logger.error(error_msg, exc_info=True)
                raise RuntimeError(error_msg) from e

        except Exception as e:
            error_msg = f'Critical error initializing ExportTab UI: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            
            # Create a detailed error layout
            error_details = f'Error details: {str(e)}\n\nPlease check the logs for more information.'
            self.logger.error(f'Error details: {error_details}')
            
            self.layout = [
                [sg.Text('⚠️ Error Initializing Export Tab',
                        font=('Helvetica', 12, 'bold'),
                        text_color='red')],
                [sg.Multiline(
                    error_details,
                    size=(80, 10),
                    disabled=True,
                    text_color='red',
                    background_color='#FFEBEE'
                )],
                [sg.Text('Check the application logs for more details.',
                        text_color='orange',
                        font=('Helvetica', 9, 'italic'))]
            ]

            self.logger.debug(f'Error layout created: {self.layout}')
            raise RuntimeError(error_msg) from e

    def get_layout(self):
        """Return the layout for this tab."""
        if not self.layout:
            error_msg = 'Layout is empty when get_layout() was called'
            self.logger.error(error_msg)
            return [[sg.Text(f'Error: {error_msg}', text_color='red')]]
            
        try:
            # Log the layout structure for debugging
            self.logger.debug(f'Returning layout with {len(self.layout)} rows')
            return self.layout
            
        except Exception as e:
            error_msg = f'Error in get_layout: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            return [[sg.Text(f'Error: {error_msg}', text_color='red')]]

    def handle_event(self, event: str, values: Dict[str, Any]) -> bool:
        """Handle UI events for this tab.

        Args:
            event: The event that occurred
            values: Dictionary of UI element values

        Returns:
            bool: True if the event was handled, False otherwise
        """
        # Handle preset selection
        if event == '-PRESET_SELECT-':
            preset_name = values['-PRESET_SELECT-']
            if preset_name:
                presets = self._preset_manager.get_all_presets()
                selected_preset = next((p for p in presets if p['name'] == preset_name), None)
                if selected_preset:
                    self._update_ui_from_preset(selected_preset)
            return True
            
        # Handle save preset
        elif event == '-SAVE_PRESET-':
            if self._current_preset_id:
                self._save_preset_dialog(is_new=False)
            else:
                self._save_preset_dialog(is_new=True)
            return True
            
        # Handle save as preset
        elif event == '-SAVE_AS_PRESET-':
            self._save_preset_dialog(is_new=True)
            return True
            
        # Handle delete preset
        elif event == '-DELETE_PRESET-':
            preset_name = values['-PRESET_SELECT-']
            if preset_name and sg.popup_yes_no(
                f'Are you sure you want to delete the preset "{preset_name}"?',
                title='Delete Preset'
            ) == 'Yes':
                presets = self._preset_manager.get_all_presets()
                preset_to_delete = next((p for p in presets if p['name'] == preset_name), None)
                if preset_to_delete and self._preset_manager.delete_preset(preset_to_delete['id']):
                    # Update the dropdown
                    presets = self._preset_manager.get_all_presets()
                    self.window['-PRESET_SELECT-'].update(
                        values=[p['name'] for p in presets],
                        value=presets[0]['name'] if presets else ''
                    )
            return True
        try:
            self.logger.debug(f'Handling event: {event}')
            self.logger.debug(f'Event values: {values}')
        except Exception as e:
            print(f'ERROR: Failed to log event: {str(e)}')
            print(f'Event: {event}')
            print(f'Values: {values}')

        try:
            # Handle export checkbox toggle
            if event == '-EXPORT_CSV-':
                enabled = values.get('-EXPORT_CSV-', False)
                self.logger.info(f'Toggling export controls. Enabled: {enabled}')
                try:
                    self._set_export_controls_enabled(enabled)
                    self.logger.debug('Successfully updated export controls')
                except Exception as e:
                    self.logger.error(f'Failed to update export controls: {str(e)}', exc_info=True)
                    self._show_error_popup('Failed to update export settings. Check logs for details.', 'UI Error')
                return True

            # Handle export button click
            elif event == '-EXPORT_CSV_BUTTON-':
                self.logger.info('Export to CSV button clicked')
                
                # Validate export is enabled
                if not values.get('-EXPORT_CSV-', False):
                    self._show_info_popup('Export is disabled. Please enable CSV export first.', 'Export Disabled')
                    return True

                # Get export settings
                output_dir = values.get('-CSV_OUTPUT_DIR-', str(Path.home() / 'email_exports'))
                filename = values.get('-CSV_PREFIX-', 'emails_export.csv')
                
                self.logger.info(f'Starting export to: {output_dir}/{filename}')
                self.logger.debug(f'Export settings: {values}')

                # Ensure output directory exists
                try:
                    output_path = Path(output_dir)
                    self.logger.debug(f'Verifying output directory: {output_path}')
                    output_path.mkdir(parents=True, exist_ok=True)
                    
                    if not output_path.is_dir():
                        raise NotADirectoryError(f'Output path is not a directory: {output_path}')
                    
                    if not os.access(output_path, os.W_OK):
                        raise PermissionError(f'No write permission for directory: {output_path}')
                    
                    self.logger.debug(f'Output directory verified: {output_path.absolute()}')
                    
                except Exception as e:
                    error_msg = f'Output directory error: {str(e)}'
                    self.logger.error(error_msg, exc_info=True)
                    self._show_error_popup(f'Cannot write to output directory.\n\n{str(e)}', 'Export Error')
                    return True

                # Prepare export data
                export_settings = {
                    'output_dir': output_dir,
                    'filename': filename,
                    'export_basic': values.get('-EXPORT_BASIC-', True),
                    'export_analysis': values.get('-EXPORT_ANALYSIS-', True),
                    'clean_bodies': values.get('-CLEAN_BODIES-', True),
                    'include_summaries': values.get('-INCLUDE_SUMMARIES-', True)
                }
                
                self.logger.info(f'Export settings: {export_settings}')

                try:
                    # Get the main application instance to access email data
                    from outlook_extractor.main import get_application
                    app = get_application()
                    
                    if not app or not hasattr(app, 'email_data') or not app.email_data:
                        raise ValueError('No email data available for export. Please load emails first.')
                    
                    # Get the email data
                    email_data = app.email_data
                    
                    # Initialize the CSV exporter
                    from outlook_extractor.export.csv_exporter import CSVExporter
                    exporter = CSVExporter(self.config)
                    
                    # Prepare export settings
                    export_settings = {
                        'clean_bodies': values.get('-CLEAN_BODIES-', True),
                        'include_summaries': values.get('-INCLUDE_SUMMARIES-', True)
                    }
                    
                    # Update the UI to show export in progress
                    if self.window:
                        self.window['-EXPORT_CSV_BUTTON-'].update(disabled=True, text='Exporting...')
                        self.window.refresh()
                    
                    # Prepare export in a separate thread with progress updates
                    self._export_in_progress = True
                    self._cancel_event.clear()
                    
                    # Get export format
                    # Get the selected format and corresponding extension
                    selected_format = values.get('-EXPORT_FORMAT-', 'CSV (.csv)')
                    file_ext = self._get_export_extension(selected_format)
                    
                    # Ensure the filename has the correct extension
                    base_name = os.path.splitext(filename)[0]
                    if not base_name:  # Handle case where filename is just an extension
                        base_name = 'export'
                    filename = f"{base_name}{file_ext}"
                    
                    # Create a progress window
                    progress_layout = [
                        [sg.Text('Exporting emails...', size=(40, 1), key='-PROGRESS-TEXT-')],
                        [sg.ProgressBar(len(email_data), orientation='h', size=(40, 20), key='-PROGRESS-BAR-')],
                        [sg.Text('Format:'), sg.Text(export_format.upper(), key='-EXPORT-FORMAT-')],
                        [sg.Text('File:'), sg.Text(filename, key='-EXPORT-FILENAME-', size=(40, 1))],
                        [sg.Button('Cancel', key='-CANCEL-EXPORT-', button_color=('white', 'red'))]
                    ]
                    
                    progress_window = sg.Window(
                        f'Exporting to {export_format.upper()}...', 
                        progress_layout, 
                        modal=True, 
                        keep_on_top=True,
                        finalize=True
                    )
                    
                    # Start export in a separate thread
                    self._export_thread = threading.Thread(
                        target=self._run_export,
                        args=(exporter, email_data, output_path, filename, progress_window, export_format),
                        daemon=True
                    )
                    self._export_thread.start()
                    
                    # Monitor progress
                    while True:
                        event, _ = progress_window.read(timeout=100)  # Check for events every 100ms
                        
                        if event == sg.WIN_CLOSED or event == '-CANCEL-EXPORT-':
                            if sg.popup_yes_no('Are you sure you want to cancel the export?', 
                                             title='Confirm Cancellation') == 'Yes':
                                self._cancel_event.set()
                                progress_window['-CANCEL-EXPORT-'].update(disabled=True)
                                progress_window['-PROGRESS-TEXT-'].update('Canceling export...')
                            continue
                        
                        # Check if export is complete
                        if not self._export_thread.is_alive():
                            break
                    
                    # Clean up
                    progress_window.close()
                    
                    # Show completion message if not canceled
                    if not self._cancel_event.is_set():
                        success_msg = f'Successfully exported {len(email_data)} emails to:\n{output_path / filename}'
                        self._show_info_popup(success_msg, 'Export Complete')
                        
                except Exception as e:
                    error_msg = f'Export failed: {str(e)}'
                    self.logger.error(error_msg, exc_info=True)
                    
                    # Show detailed error message
                    error_details = (
                        f'Error during export:\n\n'
                        f'Error: {str(e)}\n\n'
                        f'Please check that you have write permissions for the output directory and '
                        f'that there is enough disk space available.'
                    )
                    
                    self._show_error_popup(error_details, 'Export Failed')
                finally:
                    # Re-enable the export button and clean up
                    self._export_in_progress = False
                    if self.window:
                        self.window['-EXPORT_CSV_BUTTON-'].update(disabled=False, text='Export to CSV')
                        self.window.refresh()

                return True

            return False

        except Exception as e:
            error_msg = f'Unexpected error handling event {event}: {str(e)}'
            self.logger.critical(error_msg, exc_info=True)
            if self.window:
                self._show_error_popup(
                    'An unexpected error occurred.\n\n'
                    'Please check the application logs for more details.',
                    'Unexpected Error'
                )
            return False

    def _run_export(
        self, 
        exporter, 
        email_data: List[Dict], 
        output_dir: Path, 
        filename: str,
        progress_window: sg.Window,
        export_format: str = 'csv',
        validate_after_export: bool = True
    ) -> None:
        """Run the export in a separate thread with progress updates."""
        try:
            # Track progress
            def update_progress(processed: int, total: int) -> None:
                if self._cancel_event.is_set():
                    return
                
                if progress_window and not progress_window.was_closed():
                    progress_window['-PROGRESS-BAR-'].update_bar(processed, total)
                    progress_window['-PROGRESS-TEXT-'].update(
                        f'Exporting... {processed}/{total} emails ({processed/max(1,total)*100:.1f}%)'
                    )
            
            # Start the export
            output_path = output_dir / filename
            
            try:
                # Determine exporter based on file extension
                file_ext = os.path.splitext(filename)[1].lower()
                if file_ext == '.xlsx':
                    from ..export.excel_exporter import ExcelExporter
                    exporter = ExcelExporter()
                    export_format = 'excel'
                elif file_ext == '.json':
                    from ..export.json_exporter import JSONExporter
                    exporter = JSONExporter()
                    export_format = 'json'
                elif file_ext == '.pdf':
                    from ..export.pdf_exporter import PDFExporter
                    exporter = PDFExporter()
                    export_format = 'pdf'
                else:
                    exporter = CSVExporter()
                    export_format = 'csv'
                
                # Choose the appropriate export method based on format
                if export_format == 'excel':
                    from ..export.excel_exporter import ExcelExporter
                    excel_exporter = ExcelExporter()
                    success, message = excel_exporter.export_emails(
                        emails=email_data,
                        output_path=str(output_path),
                        include_headers=True,
                        progress_callback=update_progress,
                        cancel_event=self._cancel_event
                    )
                elif export_format == 'json':
                    from ..export.json_exporter import JSONExporter
                    json_exporter = JSONExporter()
                    success, message = json_exporter.export_emails(
                        emails=email_data,
                        output_path=str(output_path),
                        progress_callback=update_progress,
                        cancel_event=self._cancel_event
                    )
                elif export_format == 'pdf':
                    from ..export.pdf_exporter import PDFExporter
                    pdf_exporter = PDFExporter()
                    success, message = pdf_exporter.export_emails(
                        emails=email_data,
                        output_path=str(output_path),
                        progress_callback=update_progress,
                        cancel_event=self._cancel_event
                    )
                else:  # Default to CSV
                    success, message = exporter.export_emails(
                        emails=email_data,
                        output_path=str(output_path),
                        include_headers=True,
                        progress_callback=update_progress,
                        cancel_event=self._cancel_event
                    )
                
                # Show result
                if not success and not self._cancel_event.is_set():
                    self._show_error(f"Export failed: {message}")
                elif not self._cancel_event.is_set():
                    self._show_info(f"Export completed successfully: {message}")
                    
                    # Run validation if enabled
                    if validate_after_export and output_path.exists():
                        self._validate_export(output_path, export_format, len(email_data))
                            
            except Exception as e:
                self.logger.exception("Error during export")
                self._show_error(f"An error occurred during export: {str(e)}")
            finally:
                # Ensure we always close the progress window
                if progress_window and not progress_window.was_closed():
                    progress_window.write_event_value('-EXPORT-COMPLETE-', None)
    
        except Exception as e:
            self.logger.error(f'Error during export: {str(e)}', exc_info=True)
            if not self._cancel_event.is_set():
                self._show_info_popup(
                    f'An error occurred during export:\n{str(e)}', 
                    'Export Error'
                )
        finally:
            # Ensure the progress window will close
            if progress_window and not progress_window.was_closed():
                progress_window.write_event_value('-EXPORT-COMPLETE-', None)
    
    def _get_export_formats(self) -> List[Tuple[str, str]]:
        """Get available export formats."""
        return [
            ('CSV (.csv)', '*.csv'),
            ('Excel (.xlsx)', '*.xlsx'),
            ('JSON (.json)', '*.json'),
            ('PDF (.pdf)', '*.pdf'),
        ]

    def _get_export_extension(self, export_format: str) -> str:
        """Get file extension for the specified export format."""
        if 'Excel' in export_format:
            return '.xlsx'
        elif 'JSON' in export_format:
            return '.json'
        elif 'PDF' in export_format:
            return '.pdf'
        return '.csv'  # Default to CSV

    def _get_export_format_from_extension(self, filename: str) -> str:
        """Get export format from file extension."""
        filename_lower = filename.lower()
        if filename_lower.endswith('.xlsx'):
            return 'excel'
        elif filename_lower.endswith('.json'):
            return 'json'
        return 'csv'  # Default to CSV

    def _update_ui_from_preset(self, preset: Optional[ExportPreset] = None) -> None:
        """Update UI elements based on the selected preset."""
        if not preset:
            return
            
        try:
            # Update format
            if 'format' in preset:
                self.window['-EXPORT_FORMAT-'].update(value=preset['format'].upper())
                
            # Update checkboxes
            for key, value in preset.get('options', {}).items():
                if key in self.window.AllKeysDict:
                    self.window[key].update(value=value)
                    
            # Update export fields (if field selection is implemented)
            if hasattr(self, '_update_export_fields_ui'):
                self._update_export_fields_ui(preset.get('export_fields', []))
                
            self._current_preset_id = preset.get('id')
            
        except Exception as e:
            self.logger.error("Error updating UI from preset: %s", e, exc_info=True)
    
    def _get_current_export_settings(self) -> Dict[str, Any]:
        """Get current export settings from the UI."""
        values = self.window.read(timeout=100)[1] if self.window else {}
        
        return {
            'format': values.get('-EXPORT_FORMAT-', 'CSV').lower(),
            'include_headers': values.get('-INCLUDE_HEADERS-', True),
            'export_fields': self._get_selected_export_fields(),
            'options': {
                'clean_bodies': values.get('-CLEAN_BODIES-', True),
                'include_summaries': values.get('-INCLUDE_SUMMARIES-', True),
            }
        }
    
    def _get_selected_export_fields(self) -> List[str]:
        """Get the list of currently selected export fields."""
        # This should be implemented based on how fields are selected in your UI
        # For now, return all fields as a placeholder
        from .. import constants
        return constants.EXPORT_FIELDS_V1
    
    def _save_preset_dialog(self, is_new: bool = True) -> bool:
        """Show the save preset dialog."""
        current_settings = self._get_current_export_settings()
        
        layout = [
            [sg.Text('Preset Name:'),
             sg.Input(key='-PRESET_NAME-', size=(30, 1))],
            [sg.Text('Description:'),
             sg.Multiline(key='-PRESET_DESC-', size=(30, 3))],
            [sg.Checkbox('Set as default', key='-SET_AS_DEFAULT-', default=False)],
            [sg.Button('Save'), sg.Button('Cancel')]
        ]
        
        window = sg.Window('Save Preset', layout, modal=True)
        
        try:
            while True:
                event, values = window.read()
                
                if event in (sg.WIN_CLOSED, 'Cancel'):
                    return False
                    
                if event == 'Save':
                    name = values['-PRESET_NAME-'].strip()
                    if not name:
                        sg.popup_error('Please enter a name for the preset')
                        continue
                        
                    # Create or update the preset
                    preset: ExportPreset = {
                        'id': str(hash(name)) if is_new else str(self._current_preset_id or uuid.uuid4()),
                        'name': name,
                        'description': values['-PRESET_DESC-'],
                        'created_at': datetime.now(timezone.utc).isoformat(),
                        **current_settings
                    }
                    
                    self._preset_manager.save_preset(preset)
                    
                    # Update the preset dropdown
                    presets = self._preset_manager.get_all_presets()
                    self.window['-PRESET_SELECT-'].update(
                        values=[p['name'] for p in presets],
                        value=name
                    )
                    
                    if values.get('-SET_AS_DEFAULT-'):
                        # Implement setting as default if needed
                        pass
                        
                    return True
                    
        finally:
            window.close()
    
    def _validate_export(
        self, 
        file_path: Path, 
        export_format: str, 
        expected_count: int
    ) -> None:
        """Validate an exported file.
        
        Args:
            file_path: Path to the exported file
            export_format: Export format ('csv' or 'excel')
            expected_count: Expected number of records
        """
        from ..export.validation import ExportValidator
        
        # Show validation in progress
        self._show_info("Validating exported file...")
        
        # Run validation
        result = ExportValidator.validate_export(
            file_path=file_path,
            expected_format=export_format,
            expected_count=expected_count
        )
        
        # Show validation results
        if result.is_valid:
            self._show_info(
                "✓ Export validation passed!\n"
                f"File: {file_path.name}\n"
                f"Records: {result.details.get('record_count', 'N/A')}\n"
                f"Checksum: {result.details.get('checksum', 'N/A')}",
                title="Validation Successful"
            )
        else:
            error_details = "\n".join(
                f"• {k}: {v}" for k, v in result.details.items()
                if k not in ['error', 'exception']
            )
            
            self._show_error(
                "✗ Export validation failed!\n\n"
                f"File: {file_path.name}\n"
                f"Error: {result.message}\n\n"
                f"Details:\n{error_details}",
                title="Validation Failed"
            )
            
        # Generate a detailed report
        report_path = file_path.parent / f"{file_path.stem}_validation_report.txt"
        ExportValidator.generate_validation_report(
            {str(file_path): result},
            output_path=report_path
        )
        
        self.logger.info(f"Validation report saved to: {report_path}")
    
    def _set_export_controls_enabled(self, enabled: bool) -> None:
        """Enable/disable export controls based on the export checkbox.

        Args:
            enabled: Whether to enable or disable the controls
        """
        if not self.window:
            self.logger.warning('Cannot update controls: window not initialized')
            return

        try:
            # Define all control keys that should be toggled
            control_keys: Set[str] = {
                '-CSV_OUTPUT_DIR-',
                '-CSV_BROWSE-',
                '-CSV_PREFIX-',
                '-EXPORT_BASIC-',
                '-EXPORT_ANALYSIS-',
                '-CLEAN_BODIES-',
                '-INCLUDE_SUMMARIES-',
                '-EXPORT_CSV_BUTTON-'
            }

            self.logger.debug(f'Updating {len(control_keys)} controls. Enabled: {enabled}')

            # Track which controls were updated successfully
            updated: List[str] = []
            failed: Dict[str, str] = {}

            for key in control_keys:
                try:
                    if key not in self.window.AllKeysDict:
                        self.logger.warning(f'Control not found in window: {key}')
                        continue

                    element = self.window[key]
                    if element is None:
                        self.logger.warning(f'Element is None for key: {key}')
                        continue

                    # Update the element's disabled state
                    element.update(disabled=not enabled)
                    updated.append(key)

                except Exception as e:
                    error_msg = str(e)
                    self.logger.error(f'Failed to update control {key}: {error_msg}', exc_info=True)
                    failed[key] = error_msg

            # Log results
            if updated:
                self.logger.debug(f'Successfully updated {len(updated)} controls')
            
            if failed:
                self.logger.warning(f'Failed to update {len(failed)} controls')

            # Show error for critical failures
            if len(failed) > len(control_keys) // 2:  # If more than half failed
                error_msg = (
                    f'Failed to update {len(failed)} out of {len(control_keys)} controls.\n\n'
                    'This may indicate a problem with the UI state.\n'
                    'Please restart the application and try again.'
                )
                
                self._show_error_popup(error_msg, 'UI Update Error')

        except Exception as e:
            error_msg = f'Unexpected error in _set_export_controls_enabled: {str(e)}'
            self.logger.critical(error_msg, exc_info=True)
            self._show_error_popup(
                'A critical error occurred while updating the UI.\n\n'
                'Please restart the application and try again.',
                'Critical UI Error'
            )

    def get_export_settings(self, values: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
        """Get the current export settings from the UI values.
        
        Args:
            values: Dictionary containing UI element values from PySimpleGUI
            
        Returns:
            Dictionary containing export settings with the following structure:
            {
                'export': {
                    'enable_csv': bool,
                    'output_dir': str,
                    'file_prefix': str,
                    'export_basic': bool,
                    'export_analysis': bool,
                    'clean_bodies': bool,
                    'include_summaries': bool
                }
            }
            
        Raises:
            ValueError: If values is None or empty
            TypeError: If values is not a dictionary
        """
        if not values:
            self.logger.warning('Empty values dictionary provided, using defaults')
            values = {}
            
        if not isinstance(values, dict):
            error_msg = f'Expected dictionary for values, got {type(values).__name__}'
            self.logger.error(error_msg)
            raise TypeError(error_msg)
            
        try:
            # Get values with type conversion and validation
            settings = {
                'export': {
                    'enable_csv': bool(values.get('-EXPORT_CSV-', False)),
                    'output_dir': str(values.get(
                        '-CSV_OUTPUT_DIR-', 
                        str(Path.home() / 'email_exports')
                    )),
                    'file_prefix': str(values.get('-CSV_PREFIX-', 'emails_')).strip() or 'emails_',
                    'export_basic': bool(values.get('-EXPORT_BASIC-', True)),
                    'export_analysis': bool(values.get('-EXPORT_ANALYSIS-', True)),
                    'clean_bodies': bool(values.get('-CLEAN_BODIES-', True)),
                    'include_summaries': bool(values.get('-INCLUDE_SUMMARIES-', True))
                }
            }
            
            # Log the settings being returned (without sensitive data in production)
            self.logger.debug('Export settings retrieved from UI')
            
            return settings
            
        except Exception as e:
            error_msg = f'Error getting export settings: {str(e)}'
            self.logger.error(error_msg, exc_info=True)
            # Return default settings in case of error
            return {
                'export': {
                    'enable_csv': False,
                    'output_dir': str(Path.home() / 'email_exports'),
                    'file_prefix': 'emails_',
                    'export_basic': True,
                    'export_analysis': True,
                    'clean_bodies': True,
                    'include_summaries': True
                }
            }
