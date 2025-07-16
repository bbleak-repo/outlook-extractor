"""
Configuration for the prompt management system.
"""
from pathlib import Path
from typing import Dict, Any
import os

# Default configuration for prompt management
DEFAULT_CONFIG = {
    'prompts': {
        'base_dir': str(Path.home() / '.outlook_extract' / 'prompts'),
        'enable_tracking': True,
        'auto_version': True,
        'default_model': 'gpt-4',
        'default_parameters': {
            'temperature': 0.7,
            'max_tokens': 1000,
            'top_p': 1.0,
            'frequency_penalty': 0.0,
            'presence_penalty': 0.0,
        },
        'default_tags': ['outlook-extract'],
        'experiments': {
            'default_runs': 3,
            'results_dir': 'experiments',
            'auto_save': True,
        },
        'versioning': {
            'enabled': True,
            'max_versions': 10,
            'auto_cleanup': True,
        },
    }
}

def get_prompt_config(config=None) -> Dict[str, Any]:
    """
    Get the prompt configuration, merging with defaults if needed.
    
    Args:
        config: Optional configuration (dict or ConfigManager) to merge with defaults
        
    Returns:
        Dict containing the merged configuration
    """
    # Create a deep copy of the default config
    merged_config = {**DEFAULT_CONFIG}
    
    if config is not None:
        # Handle ConfigManager object
        if hasattr(config, 'get') and callable(getattr(config, 'get')):
            try:
                # Try to get the prompts section as a dictionary
                if hasattr(config, 'has_section') and config.has_section('prompts'):
                    prompts_config = {}
                    for key, value in config.items('prompts'):
                        # Try to convert string values to appropriate types
                        if value.lower() in ('true', 'false'):
                            prompts_config[key] = value.lower() == 'true'
                        elif value.isdigit():
                            prompts_config[key] = int(value)
                        else:
                            try:
                                prompts_config[key] = float(value)
                            except (ValueError, TypeError):
                                prompts_config[key] = value
                    merged_config['prompts'].update(prompts_config)
            except (TypeError, AttributeError) as e:
                import logging
                logging.getLogger(__name__).warning(f"Error processing config: {e}")
        # Handle dictionary
        elif isinstance(config, dict) and 'prompts' in config:
            merged_config['prompts'].update(config['prompts'])
    
    # Ensure required directories exist
    base_dir = Path(merged_config['prompts']['base_dir'])
    versions_dir = base_dir / 'versions'
    experiments_dir = base_dir / 'experiments'
    
    for directory in [base_dir, versions_dir, experiments_dir]:
        directory.mkdir(parents=True, exist_ok=True)
    
    return merged_config

def update_prompt_config(new_config: Dict[str, Any]) -> None:
    """
    Update the prompt configuration.
    
    Args:
        new_config: Dictionary containing configuration updates
    """
    from ..config import update_config
    
    # Only update the prompts section
    if 'prompts' in new_config:
        update_config({'prompts': new_config['prompts']})
