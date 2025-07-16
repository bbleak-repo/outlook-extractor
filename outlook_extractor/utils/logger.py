"""
Mock logger utility for testing purposes.
"""
import logging

def get_logger(name):
    """Return a mock logger for testing."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Create a null handler to suppress output during tests
    if not logger.handlers:
        logger.addHandler(logging.NullHandler())
    
    return logger
