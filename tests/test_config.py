"""Tests for the configuration management system."""
import os
import pytest
from pathlib import Path
from unittest.mock import patch, mock_open

# Add the project root to the Python path
import sys
sys.path.append(str(Path(__file__).parent.parent))

from outlook_extractor.config import ConfigManager, get_config

# Sample config content for testing
SAMPLE_CONFIG = """
[outlook]
folder_patterns = Inbox,Sent Items
max_emails = 1000

[threading]
enable_threading = True
"""

@pytest.fixture
def temp_config(tmp_path):
    """Create a temporary config file for testing."""
    config_path = tmp_path / "test_config.ini"
    with open(config_path, 'w') as f:
        f.write(SAMPLE_CONFIG)
    return str(config_path)

def test_config_loading(temp_config):
    """Test that config loads correctly from file."""
    config = ConfigManager(temp_config)
    assert config is not None
    assert config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'
    assert config.get_int('outlook', 'max_emails') == 1000
    assert config.get_boolean('threading', 'enable_threading') is True

def test_config_defaults():
    """Test that default values are set correctly."""
    with patch('os.path.exists', return_value=False):
        config = ConfigManager()
        assert config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'
        assert config.get_int('outlook', 'max_emails') == 1000

def test_config_get_list(temp_config):
    """Test getting a list from config."""
    config = ConfigManager(temp_config)
    folders = config.get_list('outlook', 'folder_patterns')
    assert folders == ['Inbox', 'Sent Items']
    assert config.get_list('nonexistent', 'nonexistent', ['default']) == ['default']

def test_config_get_boolean(temp_config):
    """Test boolean config values."""
    config = ConfigManager(temp_config)
    assert config.get_boolean('threading', 'enable_threading') is True
    assert config.get_boolean('nonexistent', 'nonexistent', True) is True
    assert config.get_boolean('nonexistent', 'nonexistent', False) is False

def test_config_get_int(temp_config):
    """Test integer config values."""
    config = ConfigManager(temp_config)
    assert config.get_int('outlook', 'max_emails') == 1000
    assert config.get_int('nonexistent', 'nonexistent', 42) == 42

def test_config_save(temp_config, tmp_path):
    """Test saving config to a new file."""
    config = ConfigManager(temp_config)
    new_path = str(tmp_path / "new_config.ini")
    config.save_config(new_path)
    assert os.path.exists(new_path)
    
    # Verify the saved config
    new_config = ConfigManager(new_path)
    assert new_config.get('outlook', 'folder_patterns') == 'Inbox,Sent Items'

def test_get_config_singleton(temp_config):
    """Test that get_config returns a singleton instance."""
    config1 = get_config(temp_config)
    config2 = get_config(temp_config)
    assert config1 is config2
