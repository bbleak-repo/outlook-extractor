"""MAPI property access service for Outlook integration.

This module provides robust access to MAPI properties with proper error handling
and fallback mechanisms for different Outlook versions and configurations.
"""
import logging
from typing import Any, Optional

logger = logging.getLogger(__name__)

class MAPIPropertyAccessor:
    """Robust MAPI property accessor with graceful fallbacks.
    
    This class provides a safe way to access MAPI properties from Outlook objects
    with comprehensive error handling and logging.
    """
    
    def __init__(self, mapi_object):
        """Initialize with a MAPI object (typically a MailItem or Folder).
        
        Args:
            mapi_object: The MAPI object to access properties from
        """
        self.mapi_object = mapi_object
        
    def get_property(self, property_tag: str, default: Any = None) -> Any:
        """Safely get a MAPI property with fallback handling.
        
        Args:
            property_tag: The MAPI property tag or name (e.g., 'PR_SUBJECT')
            default: Default value to return if property cannot be retrieved
            
        Returns:
            The property value or default if not found/accessible
        """
        try:
            if not hasattr(self.mapi_object, 'PropertyAccessor'):
                logger.debug("MAPI object has no PropertyAccessor")
                return default
                
            return self.mapi_object.PropertyAccessor.GetProperty(property_tag)
            
        except AttributeError as e:
            logger.debug(f"Property {property_tag} not accessible: {e}")
            return default
            
        except Exception as e:
            logger.warning(
                f"Error accessing property {property_tag}: {e}",
                exc_info=logger.isEnabledFor(logging.DEBUG)
            )
            return default
            
    def get_named_property(self, schema_name: str, guid: str = None) -> Any:
        """Get a named MAPI property using its schema name and GUID.
        
        Args:
            schema_name: The schema name of the property
            guid: Optional GUID string (defaults to PS_PUBLIC_STRINGS)
            
        Returns:
            The property value or None if not found
        """
        if not hasattr(self.mapi_object, 'PropertyAccessor'):
            return None
            
        try:
            if not guid:
                # Default to PS_PUBLIC_STRINGS
                guid = '{00020329-0000-0000-C000-000000000046}'
                
            prop_name = f"{{{guid}}} {schema_name}"
            return self.mapi_object.PropertyAccessor.GetProperty(prop_name)
            
        except Exception as e:
            logger.debug(f"Failed to get named property {schema_name}: {e}")
            return None
