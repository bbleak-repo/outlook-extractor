"""
Module to hold the global prompt manager instance.
This helps avoid circular imports between modules.
"""
from .manager import PromptManager

# Create a default prompt manager instance
prompt_manager = PromptManager()
