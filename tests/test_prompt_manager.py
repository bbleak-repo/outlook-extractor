"""
Tests for the PromptManager class.
"""
import os
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock

# Add the project root to the Python path
import sys
sys.path.append(str(Path(__file__).parent.parent))

from outlook_extractor.prompts.manager import PromptManager
from outlook_extractor.prompts.utils import PromptTemplate

class TestPromptManager(unittest.TestCase):
    """Test cases for the PromptManager class."""
    
    def setUp(self):
        """Set up test environment."""
        # Create a temporary directory for testing
        self.temp_dir = tempfile.TemporaryDirectory()
        self.base_dir = Path(self.temp_dir.name)
        
        # Initialize the prompt manager with test config
        self.config = {
            'prompts': {
                'base_dir': str(self.base_dir),
                'enable_tracking': True,
                'default_model': 'gpt-4',
                'default_parameters': {
                    'temperature': 0.7,
                    'max_tokens': 1000
                },
                'default_tags': ['test']
            }
        }
        
        self.pm = PromptManager(self.config)
    
    def tearDown(self):
        """Clean up test environment."""
        self.temp_dir.cleanup()
    
    def test_track_prompt(self):
        """Test tracking a prompt."""
        # Test data
        prompt_id = "test_prompt"
        prompt_text = "This is a test prompt with {variable}"
        model = "gpt-4"
        parameters = {"temperature": 0.8}
        source_file = "test_file.py"
        line_number = 42
        tags = ["test", "example"]
        
        # Track the prompt
        version_hash = self.pm.track_prompt(
            prompt_id=prompt_id,
            prompt_text=prompt_text,
            model=model,
            parameters=parameters,
            source_file=source_file,
            line_number=line_number,
            tags=tags
        )
        
        # Verify the version hash was returned
        self.assertIsNotNone(version_hash)
        self.assertEqual(len(version_hash), 8)  # Should be 8 chars long
        
        # Verify the prompt was saved to the versions directory
        version_file = self.base_dir / 'versions' / f"{prompt_id}_{version_hash}.json"
        self.assertTrue(version_file.exists())
        
        # Verify the content of the version file
        with open(version_file, 'r') as f:
            version_data = json.load(f)
        
        self.assertEqual(version_data['version'], version_hash)
        self.assertEqual(version_data['prompt_text'], prompt_text)
        self.assertEqual(version_data['model'], model)
        self.assertEqual(version_data['parameters'], parameters)
        self.assertEqual(version_data['source_file'], source_file)
        self.assertEqual(version_data['line_number'], line_number)
        self.assertEqual(version_data['tags'], tags)
    
    def test_get_prompt_history(self):
        """Test retrieving prompt history."""
        # Track multiple versions of a prompt
        prompt_id = "test_history"
        
        # Track version 1
        version1 = self.pm.track_prompt(
            prompt_id=prompt_id,
            prompt_text="Version 1",
            model="gpt-4",
            parameters={},
            source_file="test.py",
            line_number=1
        )
        
        # Track version 2
        version2 = self.pm.track_prompt(
            prompt_id=prompt_id,
            prompt_text="Version 2",
            model="gpt-4",
            parameters={"temperature": 0.8},
            source_file="test.py",
            line_number=2
        )
        
        # Get the history
        history = self.pm.get_prompt_history(prompt_id)
        
        # Verify the history
        self.assertEqual(len(history), 2)
        self.assertEqual(history[0]['version'], version1)
        self.assertEqual(history[1]['version'], version2)
        self.assertEqual(history[0]['prompt_text'], "Version 1")
        self.assertEqual(history[1]['prompt_text'], "Version 2")
    
    def test_run_ab_test(self):
        """Test running an A/B test."""
        # Define test variants
        variants = [
            {
                'name': 'variant1',
                'prompt': 'This is variant 1 with {input}',
                'parameters': {'temperature': 0.7}
            },
            {
                'name': 'variant2',
                'prompt': 'This is variant 2 with {input}',
                'parameters': {'temperature': 0.9}
            }
        ]
        
        # Define test inputs
        test_inputs = [
            {'input': 'test input 1'},
            {'input': 'test input 2'}
        ]
        
        # Define a simple metric function
        def length_metric(output):
            return len(output)
        
        # Run the A/B test
        results = self.pm.run_ab_test(
            test_name="test_ab_test",
            prompt_variants=variants,
            test_inputs=test_inputs,
            evaluation_metrics={'length': length_metric},
            num_runs=2
        )
        
        # Verify the results
        self.assertEqual(results['test_name'], "test_ab_test")
        self.assertEqual(len(results['variants']), 2)
        self.assertEqual(len(results['runs']), 2)
        
        # Check that metrics were calculated
        self.assertIn('length', results['metrics'])
        self.assertEqual(len(results['metrics']['length']), 2)
        
        # Check that the test results were saved
        test_files = list((self.base_dir / 'experiments').glob('test_ab_test_*.json'))
        self.assertGreaterEqual(len(test_files), 1)
    
    def test_compare_versions(self):
        """Test comparing two versions of a prompt."""
        # Track two versions of a prompt
        prompt_id = "test_compare"
        
        version1 = self.pm.track_prompt(
            prompt_id=prompt_id,
            prompt_text="Old version",
            model="gpt-3.5",
            parameters={"temperature": 0.7},
            source_file="test.py",
            line_number=1,
            tags=["old"]
        )
        
        version2 = self.pm.track_prompt(
            prompt_id=prompt_id,
            prompt_text="New version",
            model="gpt-4",
            parameters={"temperature": 0.8},
            source_file="test.py",
            line_number=2,
            tags=["new"]
        )
        
        # Compare the versions
        comparison = self.pm.compare_versions(prompt_id, version1, version2)
        
        # Verify the comparison
        self.assertEqual(comparison['prompt_id'], prompt_id)
        self.assertEqual(comparison['version1'], version1)
        self.assertEqual(comparison['version2'], version2)
        self.assertIn('diff', comparison)
        self.assertIn('metadata_diff', comparison)
        
        # Check the metadata diff
        md = comparison['metadata_diff']
        self.assertEqual(md['model'], ("gpt-3.5", "gpt-4"))
        self.assertEqual(md['parameters']['temperature'], (0.7, 0.8))
        self.assertEqual(md['tags'], (["old"], ["new"]))

if __name__ == '__main__':
    unittest.main()
