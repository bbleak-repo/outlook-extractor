"""Tests for the MAPI service module."""
import pytest
from unittest.mock import MagicMock, patch

from outlook_extractor.core.mapi_service import MAPIPropertyAccessor

class TestMAPIPropertyAccessor:
    """Test cases for the MAPIPropertyAccessor class."""
    
    @pytest.fixture
    def mock_message(self):
        """Create a mock Outlook message object."""
        message = MagicMock()
        message.PropertyAccessor = MagicMock()
        return message
    
    def test_get_existing_property(self, mock_message):
        """Test getting an existing property."""
        # Setup
        expected_value = "test-value"
        prop_uri = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
        mock_message.PropertyAccessor.GetProperty.return_value = expected_value
        
        # Test
        accessor = MAPIPropertyAccessor(mock_message)
        result = accessor.get_property(prop_uri)
        
        # Verify
        mock_message.PropertyAccessor.GetProperty.assert_called_once_with(prop_uri)
        assert result == expected_value
    
    def test_get_nonexistent_property_returns_none(self, mock_message):
        """Test getting a non-existent property returns None."""
        # Setup
        prop_uri = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
        mock_message.PropertyAccessor.GetProperty.side_effect = Exception("Property not found")
        
        # Test
        accessor = MAPIPropertyAccessor(mock_message)
        result = accessor.get_property(prop_uri)
        
        # Verify
        assert result is None
    
    def test_get_property_with_default_value(self, mock_message):
        """Test getting a property with a default value."""
        # Setup
        default_value = "default"
        prop_uri = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
        mock_message.PropertyAccessor.GetProperty.side_effect = Exception("Property not found")
        
        # Test
        accessor = MAPIPropertyAccessor(mock_message)
        result = accessor.get_property(prop_uri, default=default_value)
        
        # Verify
        assert result == default_value
    
    @patch('outlook_extractor.core.mapi_service.logger')
    def test_get_property_logs_errors(self, mock_logger, mock_message):
        """Test that errors are logged when getting a property fails."""
        # Setup
        prop_uri = "http://schemas.microsoft.com/mapi/proptag/0x0FFF0102"
        error_msg = "Test error"
        mock_message.PropertyAccessor.GetProperty.side_effect = Exception(error_msg)
        
        # Test
        accessor = MAPIPropertyAccessor(mock_message)
        accessor.get_property(prop_uri)
        
        # Verify
        mock_logger.debug.assert_called()
        assert error_msg in mock_logger.debug.call_args[0][0]
