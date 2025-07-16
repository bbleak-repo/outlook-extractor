# tests/conftest.py
import os
import tempfile
import pytest
from pathlib import Path
from outlook_extractor.config import ConfigManager

@pytest.fixture
def temp_config():
    """Create a temporary config file for testing."""
    with tempfile.NamedTemporaryFile(mode='w+', suffix='.ini', delete=False) as f:
        f.write("""
        [outlook]
        mailbox_name = 
        folder_patterns = Inbox,Sent Items
        max_emails = 1000
        
        [date_range]
        days_back = 30
        date_ranges = 
        
        [threading]
        enable_threading = 1
        thread_method = hybrid
        max_thread_depth = 10
        thread_timeout_days = 30
        
        [storage]
        output_dir = output
        db_filename = test_emails.db
        json_export = 1
        json_pretty_print = 1
        type = sqlite
        
        [logging]
        log_level = INFO
        log_file = test_outlook_extractor.log
        
        [email_processing]
        extract_attachments = 0
        attachment_dir = test_attachments
        extract_embedded_images = 0
        image_dir = test_images
        extract_links = 1
        extract_phone_numbers = 1
        
        [security]
        redact_sensitive_data = 1
        redaction_patterns = password,ssn,credit.?card
        """)
        config_path = f.name
    
    # Yield the path to the config file
    yield config_path
    
    # Clean up the temporary file
    try:
        os.unlink(config_path)
    except Exception as e:
        print(f"Warning: Could not delete temporary config file: {e}")

@pytest.fixture
def config_manager(temp_config):
    """Create a ConfigManager instance with test config."""
    return ConfigManager(temp_config)

@pytest.fixture
def temp_db():
    """Create a temporary SQLite database for testing."""
    import os
    
    # Create a temporary file
    fd, path = tempfile.mkstemp(suffix='.db')
    os.close(fd)
    
    # Create the database
    storage = SQLiteStorage(path)
    storage.initialize()
    storage.close()
    
    # Yield the path to the test database
    yield path
    
    # Clean up
    try:
        os.unlink(path)
    except Exception as e:
        print(f"Warning: Could not delete temporary database file: {e}")

@pytest.fixture
def temp_json():
    """Create a temporary JSON file for testing."""
    import os
    
    # Create a temporary file
    fd, path = tempfile.mkstemp(suffix='.json')
    os.close(fd)
    
    # Yield the path to the test JSON file
    yield path
    
    # Clean up
    try:
        os.unlink(path)
    except Exception as e:
        print(f"Warning: Could not delete temporary JSON file: {e}")
