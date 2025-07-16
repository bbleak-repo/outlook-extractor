"""
Mac-compatible launcher for Outlook Extractor.

This script provides a macOS-compatible entry point for the application,
handling platform-specific differences and providing appropriate feedback.
"""
import sys
import os
import platform
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('outlook_extractor_mac.log')
    ]
)
logger = logging.getLogger(__name__)

def check_platform():
    """Check if the current platform is supported."""
    if platform.system() != 'Windows':
        logger.warning("This application is primarily designed for Windows. "
                      "Some features may be limited on macOS.")
        return False
    return True

def main():
    """Main entry point for the application."""
    try:
        logger.info("Starting Outlook Extractor on macOS")
        
        # Check platform compatibility
        is_windows = check_platform()
        
        # Add the project root to the Python path
        project_root = os.path.dirname(os.path.abspath(__file__))
        if project_root not in sys.path:
            sys.path.insert(0, project_root)
        
        # Import the main application
        from outlook_extractor.ui.main_window import EmailExtractorUI
        
        # Create and run the application
        logger.info("Initializing UI...")
        app = EmailExtractorUI()
        
        if not is_windows:
            # Show a warning about limited functionality
            import PySimpleGUI as sg
            sg.popup(
                "Running in Compatibility Mode",
                "You are running Outlook Extractor on macOS. "
                "Some features may be limited or unavailable.\n\n"
                "Full functionality is only available on Windows.",
                title="Compatibility Notice"
            )
        
        logger.info("Starting application...")
        app.run()
        
    except Exception as e:
        logger.critical(f"Application failed to start: {str(e)}", exc_info=True)
        
        # Show error to user if possible
        try:
            import PySimpleGUI as sg
            sg.popup_error(
                f"Failed to start application: {str(e)}\n\n"
                "Please check the log file for more details.",
                title="Startup Error"
            )
        except:
            print(f"Failed to start application: {str(e)}")
            print("Please check the log file for more details.")
        
        return 1
    
    logger.info("Application closed")
    return 0

if __name__ == "__main__":
    sys.exit(main())
