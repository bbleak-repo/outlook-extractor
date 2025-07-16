#!/usr/bin/env python3
"""
Main entry point for the Outlook Extractor application.
"""
import os
import sys
import traceback
import logging
from pathlib import Path

def setup_environment():
    """Set up the Python environment and paths."""
    # Add the project root to the Python path
    project_root = str(Path(__file__).parent.parent)
    if project_root not in sys.path:
        sys.path.insert(0, project_root)

def main():
    """Main entry point for the application."""
    try:
        # Import here to ensure environment is set up first
        from outlook_extractor.ui import EmailExtractorUI
        from outlook_extractor.logging_setup import setup_logging, get_logger
        
        # Set up basic logging first
        logger = setup_logging()
        
        # Get logger after setup
        logger = get_logger()
        
        # Log startup information
        logger.info('=' * 80)
        logger.info('Starting Outlook Extractor')
        logger.info(f'Python version: {sys.version}')
        logger.info(f'Platform: {sys.platform}')
        logger.info(f'Log level: {logging.getLevelName(logging.getLogger().level)}')
        logger.info('=' * 80)
        
        # Get config path from command line or use default
        config_path = sys.argv[1] if len(sys.argv) > 1 else None
        
        try:
            # Initialize and run the UI
            logger.info("Initializing UI...")
            ui = EmailExtractorUI(config_path)
            logger.info("UI initialized successfully")
            
            # Run the application
            logger.info("Starting main application loop")
            ui.run()
            
        except Exception as e:
            logger.critical("Fatal error in main application", exc_info=True)
            sys.exit(1)
            
    except Exception as e:
        # Fallback error handling if logging setup failed
        print(f"FATAL ERROR: {str(e)}", file=sys.stderr)
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == '__main__':
    setup_environment()
    sys.exit(main())
