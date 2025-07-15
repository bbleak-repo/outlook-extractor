import sys
import os

# Add the current directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Import the UI
from outlook_extractor.ui.main_window import main

if __name__ == "__main__":
    # Set up a simple configuration for testing
    import configparser
    config = configparser.ConfigParser()
    config['outlook'] = {
        'mailbox_name': 'Test Mailbox',
        'folder_patterns': 'Inbox,Sent',
        'max_emails': '100'
    }
    # Add other default sections...
    
    # Save the test config
    config_path = 'test_config.ini'
    with open(config_path, 'w') as f:
        config.write(f)
    
    print(f"Starting UI with test configuration: {config_path}")
    main()
