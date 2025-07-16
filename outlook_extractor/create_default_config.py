"""
Script to create a default configuration file in the user's application data directory.
"""
import os
import shutil
from pathlib import Path

def create_default_config():
    """Create a default configuration file if it doesn't exist."""
    # Get the path to the default config file
    package_dir = os.path.dirname(os.path.abspath(__file__))
    default_config_path = os.path.join(package_dir, 'default_config.ini')
    
    # Determine the user's application data directory
    app_data_dir = os.path.join(os.path.expanduser('~'), '.outlook_extractor')
    os.makedirs(app_data_dir, exist_ok=True)
    
    # Path to the user's config file
    user_config_path = os.path.join(app_data_dir, 'config.ini')
    
    # Copy the default config if it doesn't exist
    if not os.path.exists(user_config_path):
        try:
            shutil.copy2(default_config_path, user_config_path)
            print(f"Created default configuration at: {user_config_path}")
            return True
        except Exception as e:
            print(f"Error creating configuration file: {e}")
            return False
    else:
        print(f"Configuration file already exists at: {user_config_path}")
        return True

if __name__ == "__main__":
    create_default_config()
