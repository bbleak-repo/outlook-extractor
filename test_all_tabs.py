"""
Test script for all UI tabs in the Outlook Extractor application.

This script creates a simple PySimpleGUI window with all the tabs to test their functionality.
"""
import sys
import os
import logging
import PySimpleGUI as sg
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add the project root to the Python path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '.')))

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('tab_test.log')
    ]
)
logger = logging.getLogger(__name__)

def test_tab(tab_name: str, tab_layout: List[List[sg.Element]], tab_key: str) -> sg.Tab:
    """Create a test tab with the given layout."""
    try:
        logger.info(f"Creating {tab_name} tab...")
        return sg.Tab(tab_name, tab_layout, key=f'-TAB_{tab_key.upper()}-')
    except Exception as e:
        logger.error(f"Error creating {tab_name} tab: {str(e)}", exc_info=True)
        error_layout = [
            [sg.Text(f"Error loading {tab_name} tab:", text_color='red')],
            [sg.Multiline(str(e), size=(80, 10), disabled=True)]
        ]
        return sg.Tab(tab_name, error_layout, key=f'-TAB_{tab_key.upper()}_ERROR-')

def main():
    """Create a test window with all the application tabs."""
    try:
        logger.info("Starting tab test...")
        
        # Import the main window class to access tab creation methods
        from outlook_extractor.ui.main_window import EmailExtractorUI
        
        # Create a dummy config for testing
        class DummyConfig:
            def get(self, *args, **kwargs):
                return kwargs.get('default')
                
            def getboolean(self, *args, **kwargs):
                return kwargs.get('default', False)
                
            def getint(self, *args, **kwargs):
                return kwargs.get('default', 0)
        
        # Create a minimal EmailExtractorUI instance
        class TestUI(EmailExtractorUI):
            def __init__(self):
                self.config = DummyConfig()
                self.logger = logger
                self.theme = 'LightGrey1'
                self.window = None
        
        ui = TestUI()
        
        # Create all tabs
        tabs = [
            test_tab("Extraction", ui._create_extraction_tab() or [[sg.Text("Extraction Tab")]], "extraction"),
            test_tab("Storage", ui._create_storage_tab() or [[sg.Text("Storage Tab")]], "storage"),
            test_tab("Threading", ui._create_threading_tab() or [[sg.Text("Threading Tab")]], "threading"),
            test_tab("Email Processing", ui._create_email_processing_tab() or [[sg.Text("Email Processing Tab")]], "email_processing"),
            test_tab("Security", ui._create_security_tab() or [[sg.Text("Security Tab")]], "security"),
            test_tab("Logs", ui._create_logs_tab() or [[sg.Text("Logs Tab")]], "logs"),
            test_tab("About", ui._create_about_tab() or [[sg.Text("About Tab")]], "about")
        ]
        
        # Create the tab group
        tab_group = sg.TabGroup([tabs], key='-TAB_GROUP-', enable_events=True)
        
        # Create the window layout
        layout = [
            [sg.Text('Outlook Extractor - Tab Test', font=('Helvetica', 16))],
            [sg.HorizontalSeparator()],
            [tab_group],
            [sg.Button('Refresh'), sg.Button('Exit')]
        ]
        
        # Create the window
        window = sg.Window('Tab Tester', layout, finalize=True)
        
        # Event loop
        while True:
            event, values = window.read()
            
            if event in (sg.WIN_CLOSED, 'Exit'):
                break
                
            elif event == 'Refresh':
                sg.popup_auto_close('Refreshing tabs...', auto_close_duration=1)
        
        window.close()
        logger.info("Tab test completed successfully")
        
    except Exception as e:
        logger.critical(f"Tab test failed: {str(e)}", exc_info=True)
        sg.popup_error(f"Test failed: {str(e)}", title="Test Error")

if __name__ == '__main__':
    main()
