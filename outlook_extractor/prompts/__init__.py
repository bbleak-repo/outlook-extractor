"""
Prompt management system for the Outlook Extract application.

This package provides functionality for tracking, versioning, and testing
prompts used in the application.
"""

import os
import sys
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent.parent
sys.path.append(str(project_root))

# Import the prompt manager and other utilities
from .manager import PromptManager
from .utils import PromptTemplate, generate_prompt_id, format_prompt
from .instance import prompt_manager

# Re-export track_prompt from utils after prompt_manager is defined
from .utils import track_prompt

def get_prompt_config():
    """Get the prompt management configuration."""
    # Default configuration
    config = {
        'prompts': {
            'base_dir': str(Path.home() / '.outlook_extract' / 'prompts'),
            'enable_tracking': True,
            'default_model': 'gpt-4',
            'default_parameters': {
                'temperature': 0.7,
                'max_tokens': 1000,
                'top_p': 1.0,
                'frequency_penalty': 0.0,
                'presence_penalty': 0.0
            },
            'default_tags': ['outlook-extract'],
            'experiments': {
                'default_num_runs': 3,
                'default_metrics': ['accuracy', 'relevance']
            }
        }
    }
    
    # Try to load from config file if it exists
    try:
        from ..config import get_config
        user_config = get_config().get('prompts', {})
        # Merge with defaults
        config['prompts'].update(user_config)
    except ImportError:
        pass
    
    return config

# Export the public API
__all__ = [
    'PromptManager',
    'prompt_manager',
    'PromptTemplate',
    'track_prompt',
    'generate_prompt_id',
    'format_prompt',
    'get_prompt_config'
]

# Clean up the namespace
del os, sys, Path, project_root
