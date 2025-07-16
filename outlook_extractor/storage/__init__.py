"""
Storage module for the Outlook Extractor package.

This module provides classes for storing and retrieving email data in different formats.
"""

from .base import EmailStorage
from .sqlite_storage import SQLiteStorage
from .json_storage import JSONStorage

__all__ = ['EmailStorage', 'SQLiteStorage', 'JSONStorage']
