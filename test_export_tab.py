"""
Test script for the ExportTab class.

This script creates a simple PySimpleGUI window with the ExportTab to test its functionality.
"""
import sys
import os
import PySimpleGUI as sg
from pathlib import Path
from typing import Dict, Any

# Add the project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '.')))

# Import the ExportTab class
from outlook_extractor.ui.export_tab import ExportTab

def main():
    """Create a test window with the ExportTab."""
    # Set up logging
    import logging
    logging.basicConfig(level=logging.DEBUG)
    
    # Create a simple layout with the ExportTab
    layout = [
        [sg.Text('Export Tab Test', font=('Helvetica', 16))],
        [sg.HorizontalSeparator()],
        [sg.Frame('Export Settings', layout=ExportTab().get_layout())],
        [sg.Button('Get Settings'), sg.Button('Exit')]
    ]
    
    # Create the window
    window = sg.Window('ExportTab Test', layout, finalize=True)
    
    # Create the ExportTab instance
    export_tab = ExportTab()
    
    # Set up folder patterns for testing
    export_tab.update_folder_patterns(['Inbox', 'Sent Items'])
    
    # Event loop
    while True:
        event, values = window.read()
        
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
            
        elif event == 'Get Settings':
            try:
                settings = export_tab.get_export_settings(values)
                sg.popup_scrolled(
                    'Current Export Settings:',
                    json.dumps(settings, indent=2),
                    title='Export Settings',
                    size=(60, 15)
                )
            except Exception as e:
                sg.popup_error(f'Error getting settings: {str(e)}')
    
    window.close()

if __name__ == '__main__':
    main()
