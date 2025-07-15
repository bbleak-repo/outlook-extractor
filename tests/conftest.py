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
    
    yield config_path
    # Cleanup
    if os.path.exists(config_path):
        os.unlink(config_path)

@pytest.fixture
def config_manager(temp_config):
    """Create a ConfigManager instance with test config."""
    return ConfigManager(temp_config)

@pytest.fixture
def temp_db():
    """Create a temporary SQLite database for testing."""
    db_path = Path("test_emails.db")
    yield str(db_path)
    # Cleanup
    if db_path.exists():
        db_path.unlink()

@pytest.fixture
def temp_json():
    """Create a temporary JSON file for testing."""
    json_path = Path("test_
