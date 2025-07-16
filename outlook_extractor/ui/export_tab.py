# outlook_extractor/ui/export_tab.py

import logging
import os
import FreeSimpleGUI as sg
from pathlib import Path
from typing import Dict, Any, List, Optional, Set

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
                sg.popup_error(message, title=title, icon=sg.ICON_ERROR)
            except Exception as e:
                self.logger.critical(f'Failed to show error popup: {str(e)}', exc_info=True)
                # Fallback to console if UI fails
                print(f'ERROR [{title}]: {message}')

    def _show_info_popup(self, message: str, title: str = 'Information') -> None:
        """Show an information popup and log the message.

        Args:
            message: The message to display
            title: The title of the popup
        """
        self.logger.info(f'{title}: {message}')
        if self.window:
            try:
                sg.popup_ok(message, title=title, icon=sg.ICON_INFORMATION)
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
                sg.popup_ok(message, title=title, icon=sg.ICON_WARNING)
            except Exception as e:
                self.logger.error(f'Failed to show warning popup: {str(e)}', exc_info=True)
                print(f'WARNING [{title}]: {message}')

    def __init__(self, config: Dict[str, Any] = None):
        """Initialize the export tab.

        Args:
            config: Optional configuration dictionary. If None, default values will be used.
        """
        # Configure logging first
        self.logger = logging.getLogger(f'{__name__}.{self.__class__.__name__}')
        self.logger.info('=' * 50)
        self.logger.info('Initializing ExportTab')

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
                    # Create a simpler single-column layout for better macOS compatibility
                    export_options = [
                        [sg.Text('Export Options:', 
                               font=('Helvetica', 10, 'bold'), 
                               pad=(10, (10, 5)))],
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
                    
                    # Export to CSV
                    output_file = output_path / filename
                    export_path = exporter.export_emails_to_csv(
                        emails=email_data,
                        output_path=str(output_file),
                        include_headers=True
                    )
                    
                    success_msg = f'Successfully exported {len(email_data)} emails to:\n{export_path}'
                    self.logger.info(success_msg)
                    
                    # Show success message
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
                    # Re-enable the export button
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
