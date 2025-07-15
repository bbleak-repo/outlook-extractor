"""
Configuration management for the Outlook Extractor package.
Handles loading, validating, and providing access to configuration settings.
"""
import configparser
import os
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
import logging
from datetime import datetime, timedelta

# Default configuration values
DEFAULT_CONFIG = {
    'outlook': {
        'mailbox_name': '',  # Empty means default mailbox
        'folder_patterns': 'Inbox,Sent Items',
        'max_emails': '1000',
    },
    'date_range': {
        'days_back': '30',
        'date_ranges': '',  # Format: 'MM/YYYY,MM/YYYY'
    },
    'threading': {
        'enable_threading': '1',
        'thread_method': 'hybrid',  # 'headers', 'content', or 'hybrid'
        'max_thread_depth': '10',
        'thread_timeout_days': '30',
    },
    'storage': {
        'output_dir': 'output',
        'db_filename': 'outlook_emails.db',
        'json_export': '1',
        'json_pretty_print': '1',
    },
    'logging': {
        'log_level': 'INFO',
        'log_file': 'outlook_extractor.log',
    },
    'email_processing': {
        'extract_attachments': '0',
        'attachment_dir': 'attachments',
        'extract_embedded_images': '0',
        'image_dir': 'images',
        'extract_links': '1',
        'extract_phone_numbers': '1',
    },
    'security': {
        'redact_sensitive_data': '1',
        'redaction_patterns': 'password,ssn,credit.?card',
    },
}

class ConfigManager:
    """Manages configuration loading, validation, and access."""
    
    def __init__(self, config_path: str = None):
        """Initialize the configuration manager.
        
        Args:
            config_path: Optional path to a configuration file.
                        If not provided, uses default settings.
        """
        self.config = configparser.ConfigParser()
        self.logger = logging.getLogger(__name__)
        
        # Set default values
        for section, options in DEFAULT_CONFIG.items():
            self.config[section] = options
        
        # Load from file if provided
        if config_path and os.path.exists(config_path):
            self.load_config(config_path)
    
    def load_config(self, config_path: str) -> bool:
        """Load configuration from a file.
        
        Args:
            config_path: Path to the configuration file.
            
        Returns:
            bool: True if the configuration was loaded successfully, False otherwise.
        """
        try:
            self.config.read(config_path)
            self.logger.info(f"Loaded configuration from {config_path}")
            return True
        except Exception as e:
            self.logger.error(f"Error loading configuration from {config_path}: {e}")
            return False
    
    def save_config(self, config_path: str) -> bool:
        """Save the current configuration to a file.
        
        Args:
            config_path: Path where to save the configuration.
            
        Returns:
            bool: True if the configuration was saved successfully, False otherwise.
        """
        try:
            os.makedirs(os.path.dirname(os.path.abspath(config_path)), exist_ok=True)
            with open(config_path, 'w') as f:
                self.config.write(f)
            self.logger.info(f"Saved configuration to {config_path}")
            return True
        except Exception as e:
            self.logger.error(f"Error saving configuration to {config_path}: {e}")
            return False
    
    def get(self, section: str, option: str, fallback: Any = None) -> Any:
        """Get a configuration value.
        
        Args:
            section: The configuration section.
            option: The configuration option.
            fallback: Value to return if the section or option doesn't exist.
            
        Returns:
            The configuration value, or the fallback if not found.
        """
        try:
            return self.config.get(section, option, fallback=fallback)
        except (configparser.NoSectionError, configparser.NoOptionError):
            return fallback
    
    def get_boolean(self, section: str, option: str, fallback: bool = False) -> bool:
        """Get a boolean configuration value.
        
        Args:
            section: The configuration section.
            option: The configuration option.
            fallback: Value to return if the section or option doesn't exist.
            
        Returns:
            The boolean value of the configuration option, or the fallback if not found.
        """
        try:
            return self.config.getboolean(section, option, fallback=fallback)
        except (ValueError, AttributeError):
            return fallback
    
    def get_int(self, section: str, option: str, fallback: int = 0) -> int:
        """Get an integer configuration value.
        
        Args:
            section: The configuration section.
            option: The configuration option.
            fallback: Value to return if the section or option doesn't exist.
            
        Returns:
            The integer value of the configuration option, or the fallback if not found.
        """
        try:
            return self.config.getint(section, option, fallback=fallback)
        except (ValueError, AttributeError):
            return fallback
    
    def get_float(self, section: str, option: str, fallback: float = 0.0) -> float:
        """Get a float configuration value.
        
        Args:
            section: The configuration section.
            option: The configuration option.
            fallback: Value to return if the section or option doesn't exist.
            
        Returns:
            The float value of the configuration option, or the fallback if not found.
        """
        try:
            return self.config.getfloat(section, option, fallback=fallback)
        except (ValueError, AttributeError):
            return fallback
    
    def get_list(self, section: str, option: str, fallback: List[str] = None, 
                delimiter: str = ',') -> List[str]:
        """Get a list of strings from a configuration value.
        
        Args:
            section: The configuration section.
            option: The configuration option.
            fallback: Value to return if the section or option doesn't exist.
            delimiter: The delimiter used to split the string into a list.
            
        Returns:
            A list of strings from the configuration value, or the fallback if not found.
        """
        if fallback is None:
            fallback = []
        
        value = self.get(section, option)
        if value is None:
            return fallback
        
        return [item.strip() for item in value.split(delimiter) if item.strip()]
    
    def get_date_range(self) -> Tuple[datetime, datetime]:
        """Get the date range for email extraction.
        
        Returns:
            A tuple of (start_date, end_date) datetime objects.
        """
        end_date = datetime.now()
        
        # Try to get date ranges first
        date_ranges = self.get('date_range', 'date_ranges', '').strip()
        if date_ranges:
            try:
                # Format: 'MM/YYYY,MM/YYYY' (start_date, end_date)
                start_str, end_str = [s.strip() for s in date_ranges.split(',')][:2]
                start_date = datetime.strptime(start_str, '%m/%Y')
                end_date = datetime.strptime(end_str, '%m/%Y')
                
                # Set to the end of the month for the end date
                if end_date.month == 12:
                    end_date = end_date.replace(month=1, year=end_date.year + 1, day=1) - timedelta(days=1)
                else:
                    end_date = end_date.replace(month=end_date.month + 1, day=1) - timedelta(days=1)
                
                # Set to end of day
                end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
                
                return start_date, end_date
            except (ValueError, IndexError) as e:
                self.logger.warning(
                    f"Invalid date_ranges format: {date_ranges}. "
                    f"Falling back to days_back. Error: {e}"
                )
        
        # Fall back to days_back if date_ranges is not set or invalid
        days_back = self.get_int('date_range', 'days_back', 30)
        start_date = end_date - timedelta(days=days_back)
        
        return start_date, end_date
    
    def get_output_dir(self) -> str:
        """Get the output directory, creating it if it doesn't exist.
        
        Returns:
            The absolute path to the output directory.
        """
        output_dir = self.get('storage', 'output_dir', 'output')
        os.makedirs(output_dir, exist_ok=True)
        return os.path.abspath(output_dir)
    
    def get_db_path(self) -> str:
        """Get the path to the SQLite database file.
        
        Returns:
            The absolute path to the database file.
        """
        output_dir = self.get_output_dir()
        db_filename = self.get('storage', 'db_filename', 'outlook_emails.db')
        return os.path.join(output_dir, db_filename)
    
    def get_attachment_dir(self) -> str:
        """Get the directory for saving attachments, creating it if it doesn't exist.
        
        Returns:
            The absolute path to the attachment directory.
        """
        output_dir = self.get_output_dir()
        attachment_dir = self.get('email_processing', 'attachment_dir', 'attachments')
        attachment_path = os.path.join(output_dir, attachment_dir)
        os.makedirs(attachment_path, exist_ok=True)
        return attachment_path
    
    def get_image_dir(self) -> str:
        """Get the directory for saving embedded images, creating it if it doesn't exist.
        
        Returns:
            The absolute path to the image directory.
        """
        output_dir = self.get_output_dir()
        image_dir = self.get('email_processing', 'image_dir', 'images')
        image_path = os.path.join(output_dir, image_dir)
        os.makedirs(image_path, exist_ok=True)
        return image_path

# Default configuration instance
config = ConfigManager()

def get_config() -> ConfigManager:
    """Get the default configuration instance.
    
    Returns:
        The default ConfigManager instance.
    """
    return config

def load_config(config_path: str = None) -> ConfigManager:
    """Load configuration from a file and return a ConfigManager instance.
    
    Args:
        config_path: Path to the configuration file.
        
    Returns:
        A ConfigManager instance with the loaded configuration.
    """
    return ConfigManager(config_path)
