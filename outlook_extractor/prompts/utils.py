"""
Utility functions for working with prompts.
"""
import inspect
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Union
import functools
import json
import hashlib

from .instance import prompt_manager
from .config import get_prompt_config

def track_prompt(
    prompt_id: str,
    model: str = None,
    parameters: Dict[str, Any] = None,
    tags: List[str] = None,
    metadata: Dict[str, Any] = None,
    enabled: bool = True
):
    """
    Decorator to track a function that generates or uses a prompt.
    
    Args:
        prompt_id: Unique identifier for the prompt
        model: AI model used (e.g., 'gpt-4', 'claude-2')
        parameters: Dictionary of generation parameters
        tags: Optional list of tags for categorization
        metadata: Additional metadata to store with the prompt
        enabled: Whether tracking is enabled for this prompt
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Get the source file and line number
            frame = inspect.currentframe()
            try:
                source_file = inspect.getsourcefile(frame)
                line_number = inspect.getlineno(frame)
            finally:
                del frame  # Avoid reference cycles
            
            # Get the prompt text from the function's docstring
            prompt_text = inspect.getdoc(func)
            if not prompt_text:
                raise ValueError(f"Function {func.__name__} must have a docstring containing the prompt")
            
            # Get default parameters from config
            config = get_prompt_config()
            default_params = config['prompts']['default_parameters'].copy()
            if parameters:
                default_params.update(parameters)
            
            # Track the prompt
            if enabled and prompt_manager._prompt_keeper_available:
                prompt_manager.track_prompt(
                    prompt_id=prompt_id,
                    prompt_text=prompt_text,
                    model=model or config['prompts']['default_model'],
                    parameters=default_params,
                    source_file=source_file,
                    line_number=line_number,
                    tags=(tags or []) + config['prompts']['default_tags'],
                    metadata=metadata
                )
            
            # Call the original function
            return func(*args, **kwargs)
        
        return wrapper
    return decorator

def generate_prompt_id(prompt_text: str, prefix: str = 'prompt_') -> str:
    """
    Generate a deterministic ID for a prompt based on its content.
    
    Args:
        prompt_text: The prompt text
        prefix: Optional prefix for the ID
        
    Returns:
        A deterministic ID for the prompt
    """
    # Normalize the prompt text
    normalized = ' '.join(prompt_text.strip().split())
    
    # Generate a hash of the normalized text
    hash_obj = hashlib.md5(normalized.encode('utf-8'))
    hash_hex = hash_obj.hexdigest()[:8]  # Use first 8 chars of hash
    
    return f"{prefix}{hash_hex}"

def format_prompt(template: str, **kwargs) -> str:
    """
    Format a prompt template with the given variables.
    
    Args:
        template: The prompt template string with {variable} placeholders
        **kwargs: Variables to format into the template
        
    Returns:
        The formatted prompt string
    """
    try:
        return template.format(**kwargs)
    except KeyError as e:
        raise ValueError(f"Missing required variable in prompt template: {e}")

class PromptTemplate:
    """
    A class to manage prompt templates with versioning and tracking.
    """
    
    def __init__(
        self,
        template: str,
        prompt_id: str = None,
        model: str = None,
        parameters: Dict[str, Any] = None,
        tags: List[str] = None,
        metadata: Dict[str, Any] = None
    ):
        """
        Initialize a prompt template.
        
        Args:
            template: The prompt template string with {variable} placeholders
            prompt_id: Unique identifier for the prompt. If None, will be generated.
            model: AI model used (e.g., 'gpt-4', 'claude-2')
            parameters: Dictionary of generation parameters
            tags: Optional list of tags for categorization
            metadata: Additional metadata to store with the prompt
        """
        self.template = template
        self.prompt_id = prompt_id or generate_prompt_id(template)
        self.model = model
        self.parameters = parameters or {}
        self.tags = tags or []
        self.metadata = metadata or {}
        
        # Get default config
        config = get_prompt_config()
        self.default_model = model or config['prompts']['default_model']
        self.default_parameters = config['prompts']['default_parameters'].copy()
        self.default_parameters.update(self.parameters)
        
        # Track the initial version
        self.track()
    
    def track(self, **kwargs) -> Optional[str]:
        """
        Track this prompt with the prompt manager.
        
        Args:
            **kwargs: Additional metadata to include
            
        Returns:
            The version hash if tracking was successful, None otherwise
        """
        if not hasattr(prompt_manager, '_prompt_keeper_available') or not prompt_manager._prompt_keeper_available:
            return None
            
        # Merge metadata
        metadata = self.metadata.copy()
        metadata.update(kwargs)
        
        # Get the caller's source file and line number
        frame = inspect.currentframe().f_back
        try:
            source_file = inspect.getsourcefile(frame)
            line_number = frame.f_lineno
        finally:
            del frame  # Avoid reference cycles
        
        # Track the prompt
        return prompt_manager.track_prompt(
            prompt_id=self.prompt_id,
            prompt_text=self.template,
            model=self.model or self.default_model,
            parameters=self.default_parameters,
            source_file=source_file or "<unknown>",
            line_number=line_number or 0,
            tags=self.tags,
            metadata=metadata
        )
    
    def format(self, **kwargs) -> str:
        """
        Format the template with the given variables and track the usage.
        
        Args:
            **kwargs: Variables to format into the template
            
        Returns:
            The formatted prompt string
        """
        # Format the prompt
        try:
            formatted = self.template.format(**kwargs)
        except KeyError as e:
            raise ValueError(f"Missing required variable in prompt template: {e}")
        
        # Track the usage
        self.track(
            variables=kwargs,
            formatted_prompt=formatted,
            timestamp=datetime.now(timezone.utc).isoformat()
        )
        
        return formatted
    
    def __str__(self) -> str:
        return self.template
    
    def __repr__(self) -> str:
        return f"<PromptTemplate id='{self.prompt_id}'>"
