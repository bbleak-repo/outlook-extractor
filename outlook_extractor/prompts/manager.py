"""
Prompt Manager implementation for the Outlook Extract application.

This module provides the PromptManager class which handles the core functionality
for managing AI prompts, including versioning, tracking, and A/B testing.
"""
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Callable, Union, Tuple

from ..config import get_config
from ..utils.logger import get_logger

# Set up logging
logger = get_logger(__name__)

class PromptManager:
    """
    Manages AI prompts including versioning, tracking, and A/B testing.
    
    This is a self-contained implementation that doesn't require external dependencies.
    """
    
    def __init__(self, config=None):
        """Initialize the Prompt Manager.
        
        Args:
            config: Optional configuration dictionary or ConfigParser object. 
                   If not provided, loads from default config.
        """
        # Default configuration
        default_config = {
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
        
        # Load configuration
        if config is not None:
            self.config = config
        else:
            try:
                self.config = get_config() or default_config
            except Exception as e:
                logger.warning(f"Failed to load config: {e}. Using default configuration.")
                self.config = default_config
        
        # Default base directory
        default_base_dir = str(Path.home() / '.outlook_extract' / 'prompts')
        
        # Handle both dictionary and ConfigParser configs
        if hasattr(self.config, 'get') and callable(getattr(self.config, 'get')):
            try:
                # Try ConfigParser style access first
                if hasattr(self.config, 'has_section') and self.config.has_section('prompts'):
                    base_dir = self.config.get('prompts', 'base_dir', 
                                            fallback=default_base_dir)
                else:
                    # Fall back to dictionary access
                    base_dir = self.config.get('prompts', {}).get('base_dir', default_base_dir)
            except (TypeError, AttributeError) as e:
                logger.warning(f"Error accessing config: {e}. Using default base directory.")
                base_dir = default_base_dir
        else:
            # Handle as dictionary
            base_dir = self.config.get('prompts', {}).get('base_dir', default_base_dir)
        
        # Set up directories
        self.base_dir = Path(base_dir)
        self.versions_dir = self.base_dir / 'versions'
        self.experiments_dir = self.base_dir / 'experiments'
        self.prompts_file = self.base_dir / 'prompts.json'
        
        # Create directories if they don't exist
        for directory in [self.base_dir, self.versions_dir, self.experiments_dir]:
            directory.mkdir(parents=True, exist_ok=True)
        
        # Initialize prompts storage
        self._prompts = self._load_prompts()
        
        logger.info(f"Prompt Manager initialized in {self.base_dir.absolute()}")
    
    def _load_prompts(self) -> Dict:
        """Load prompts from the storage file."""
        if self.prompts_file.exists():
            try:
                with open(self.prompts_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Failed to load prompts: {str(e)}")
        return {}
    
    def _save_prompts(self):
        """Save prompts to the storage file."""
        try:
            with open(self.prompts_file, 'w', encoding='utf-8') as f:
                json.dump(self._prompts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Failed to save prompts: {str(e)}")
    
    def _generate_version_hash(self, prompt_text: str) -> str:
        """Generate a deterministic hash for a prompt version."""
        return hashlib.md5(prompt_text.encode('utf-8')).hexdigest()[:8]
    
    def _get_timestamp(self) -> str:
        """Get current timestamp in ISO format with timezone."""
        return datetime.now(timezone.utc).isoformat()
    
    def track_prompt(
        self,
        prompt_id: str,
        prompt_text: str,
        model: str,
        parameters: Dict[str, Any],
        source_file: str,
        line_number: int,
        tags: List[str] = None,
        metadata: Dict[str, Any] = None
    ) -> Optional[str]:
        """
        Track a prompt with versioning.
        
        Args:
            prompt_id: Unique identifier for the prompt
            prompt_text: The actual prompt text
            model: AI model used (e.g., 'gpt-4', 'claude-2')
            parameters: Dictionary of generation parameters
            source_file: Path to the source file where the prompt is used
            line_number: Line number in the source file
            tags: Optional list of tags for categorization
            metadata: Additional metadata to store with the prompt
            
        Returns:
            str: The version hash of the tracked prompt
        """
        # Initialize prompt entry if it doesn't exist
        if prompt_id not in self._prompts:
            self._prompts[prompt_id] = {
                'versions': [],
                'latest_version': None,
                'created_at': self._get_timestamp(),
                'updated_at': self._get_timestamp(),
                'tags': tags or [],
                'usage_count': 0
            }
        
        # Generate version hash
        version_hash = self._generate_version_hash(prompt_text)
        
        # Prepare version metadata
        version_metadata = {
            'version': version_hash,
            'prompt_text': prompt_text,
            'model': model,
            'parameters': parameters,
            'source_file': source_file,
            'line_number': line_number,
            'tags': tags or [],
            'metadata': metadata or {},
            'created_at': self._get_timestamp()
        }
        
        # Add version to history
        self._prompts[prompt_id]['versions'].append(version_metadata)
        self._prompts[prompt_id]['latest_version'] = version_hash
        self._prompts[prompt_id]['updated_at'] = self._get_timestamp()
        self._prompts[prompt_id]['usage_count'] += 1
        
        # Save a copy to the versions directory
        version_file = self.versions_dir / f"{prompt_id}_{version_hash}.json"
        try:
            with open(version_file, 'w', encoding='utf-8') as f:
                json.dump(version_metadata, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Failed to save prompt version: {str(e)}")
        
        # Save updates to disk
        self._save_prompts()
        
        logger.debug("Tracked prompt '%s' with version %s", prompt_id, version_hash)
        return version_hash
    
    def run_ab_test(
        self,
        test_name: str,
        prompt_variants: List[Dict[str, Any]],
        test_inputs: List[Any],
        evaluation_metrics: Dict[str, Callable],
        num_runs: int = 1
    ) -> Dict[str, Any]:
        """
        Run an A/B test between multiple prompt variants.
        
        Args:
            test_name: Name of the test
            prompt_variants: List of prompt variants to test
            test_inputs: List of inputs to test against
            evaluation_metrics: Dictionary of metric names to evaluation functions
            num_runs: Number of runs per variant
            
        Returns:
            Dict with test results
        """
        results = {
            'test_name': test_name,
            'start_time': self._get_timestamp(),
            'variants': {},
            'metrics': {},
            'runs': []
        }
        
        # Track each variant
        for variant in prompt_variants:
            variant_name = variant.get('name', f"variant_{len(results['variants']) + 1}")
            results['variants'][variant_name] = {
                'prompt': variant.get('prompt', ''),
                'parameters': variant.get('parameters', {})
            }
        
        # Run tests
        for run in range(num_runs):
            run_results = {'run': run + 1, 'inputs': []}
            
            for input_data in test_inputs:
                input_result = {'input': input_data, 'variants': {}}
                
                for variant_name, variant_data in results['variants'].items():
                    # In a real implementation, this would call the AI model
                    # For now, we'll just store the prompt and parameters
                    variant_result = {
                        'prompt': variant_data['prompt'],
                        'parameters': variant_data['parameters'],
                        'metrics': {}
                    }
                    
                    # Calculate metrics
                    for metric_name, metric_func in evaluation_metrics.items():
                        try:
                            variant_result['metrics'][metric_name] = metric_func(
                                variant_data['prompt'].format(**input_data)
                            )
                        except Exception as e:
                            logger.error(f"Error calculating metric {metric_name}: {str(e)}")
                            variant_result['metrics'][metric_name] = None
                    
                    input_result['variants'][variant_name] = variant_result
                
                run_results['inputs'].append(input_result)
            
            results['runs'].append(run_results)
        
        # Calculate aggregate metrics
        for metric_name in evaluation_metrics.keys():
            results['metrics'][metric_name] = {}
            for variant_name in results['variants'].keys():
                metric_values = [
                    input_result['variants'][variant_name]['metrics'].get(metric_name)
                    for run in results['runs']
                    for input_result in run['inputs']
                    if input_result['variants'][variant_name]['metrics'].get(metric_name) is not None
                ]
                
                if metric_values:
                    results['metrics'][metric_name][variant_name] = {
                        'count': len(metric_values),
                        'min': min(metric_values),
                        'max': max(metric_values),
                        'avg': sum(metric_values) / len(metric_values)
                    }
        
        results['end_time'] = self._get_timestamp()
        
        # Save test results
        test_file = self.experiments_dir / f"{test_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        try:
            with open(test_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Failed to save test results: {str(e)}")
        
        logger.info(
            "Completed A/B test '%s' with %d variants and %d runs",
            test_name,
            len(prompt_variants),
            num_runs
        )
        
        return results
    
    def get_prompt_history(self, prompt_id: str) -> List[Dict]:
        """
        Get the version history of a prompt.
        
        Args:
            prompt_id: ID of the prompt to get history for
            
        Returns:
            List of version history entries
        """
        if prompt_id in self._prompts:
            return self._prompts[prompt_id].get('versions', [])
        return []
    
    def compare_versions(self, prompt_id: str, version1: str, version2: str) -> Dict:
        """
        Compare two versions of a prompt.
        
        Args:
            prompt_id: ID of the prompt
            version1: First version hash to compare
            version2: Second version hash to compare
            
        Returns:
            Dict with comparison results
        """
        history = self.get_prompt_history(prompt_id)
        v1 = next((v for v in history if v['version'] == version1), None)
        v2 = next((v for v in history if v['version'] == version2), None)
        
        if not v1 or not v2:
            return {"error": "One or both versions not found"}
        
        # Simple text diff (in a real implementation, you might want to use a proper diff library)
        from difflib import ndiff
        
        diff = list(ndiff(
            v1['prompt_text'].splitlines(keepends=True),
            v2['prompt_text'].splitlines(keepends=True)
        ))
        
        return {
            'prompt_id': prompt_id,
            'version1': version1,
            'version2': version2,
            'diff': ''.join(diff),
            'metadata_diff': {
                'model': (v1['model'], v2['model']),
                'parameters': self._dict_diff(v1['parameters'], v2['parameters']),
                'tags': (v1['tags'], v2['tags']),
                'created_at': (v1['created_at'], v2['created_at'])
            }
        }
    
    def _dict_diff(self, d1: Dict, d2: Dict) -> Dict[str, Tuple[Any, Any]]:
        """Compute differences between two dictionaries."""
        all_keys = set(d1.keys()) | set(d2.keys())
        return {
            k: (d1.get(k), d2.get(k))
            for k in all_keys
            if d1.get(k) != d2.get(k)
        }
