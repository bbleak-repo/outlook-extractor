"""Tests for the command-line interface."""
import json
import os
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest
import yaml

# Add the project root to the Python path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import the CLI app after modifying the path
from main_cli import CLIApp

class TestCLIApp:
    """Test cases for the CLI application."""
    
    @pytest.fixture
    def cli_app(self):
        """Create a CLI app instance for testing."""
        return CLIApp()
    
    @pytest.fixture
    def sample_config_file(self, tmp_path):
        """Create a sample config file for testing."""
        config = {
            'priority_addresses': ['important@example.com'],
            'admin_addresses': ['admin@example.com'],
            'export': {
                'format': 'csv',
                'include_headers': True
            }
        }
        
        config_path = tmp_path / 'test_config.yaml'
        with open(config_path, 'w') as f:
            yaml.dump(config, f)
            
        return config_path
    
    def test_parse_args_defaults(self, cli_app):
        """Test argument parsing with default values."""
        with patch('sys.argv', ['main_cli.py']):
            cli_app.args = cli_app.parser.parse_args()
            
        assert cli_app.args.folder == 'Inbox'
        assert cli_app.args.recursive is False
        assert cli_app.args.include_read is False
        assert cli_app.args.output == 'emails.csv'
        assert cli_app.args.format == 'csv'
        assert cli_app.args.include_headers is True
        assert cli_app.args.debug is False
        assert cli_app.args.dry_run is False
    
    def test_parse_args_custom_values(self, cli_app):
        """Test argument parsing with custom values."""
        test_args = [
            'main_cli.py',
            '--folder', 'Sent Items',
            '--recursive',
            '--include-read',
            '--output', 'output.json',
            '--format', 'json',
            '--no-headers',
            '--debug',
            '--dry-run',
            '--from', 'test@example.com',
            '--subject', 'important',
            '--after', '2023-01-01',
            '--before', '2023-12-31',
            '--priority-emails', 'p1@example.com', 'p2@example.com',
            '--admin-emails', 'admin@example.com'
        ]
        
        with patch('sys.argv', test_args):
            cli_app.args = cli_app.parser.parse_args()
            
        assert cli_app.args.folder == 'Sent Items'
        assert cli_app.args.recursive is True
        assert cli_app.args.include_read is True
        assert cli_app.args.output == 'output.json'
        assert cli_app.args.format == 'json'
        assert cli_app.args.include_headers is False
        assert cli_app.args.debug is True
        assert cli_app.args.dry_run is True
        assert cli_app.args.sender == 'test@example.com'
        assert cli_app.args.subject == 'important'
        assert cli_app.args.after == '2023-01-01'
        assert cli_app.args.before == '2023-12-31'
        assert set(cli_app.args.priority_emails) == {'p1@example.com', 'p2@example.com'}
        assert cli_app.args.admin_emails == ['admin@example.com']
    
    def test_load_config_from_file(self, cli_app, sample_config_file):
        """Test loading configuration from a file."""
        cli_app.args = MagicMock()
        cli_app.args.config = str(sample_config_file)
        cli_app.args.priority_emails = None
        cli_app.args.admin_emails = None
        
        cli_app._load_config()
        
        assert 'important@example.com' in cli_app.config['priority_addresses']
        assert 'admin@example.com' in cli_app.config['admin_addresses']
    
    @patch('main_cli.OutlookClient')
    @patch('main_cli.EmailProcessor')
    def test_process_emails(self, mock_email_processor, mock_outlook_client, cli_app, tmp_path):
        """Test processing emails with the CLI app."""
        # Setup mocks
        mock_outlook = MagicMock()
        mock_outlook.get_folder.return_value = 'mock_folder'
        mock_outlook.get_messages.return_value = ['msg1', 'msg2']
        mock_outlook_client.return_value.__enter__.return_value = mock_outlook
        
        mock_processor = MagicMock()
        mock_processor.process_message.side_effect = [
            {'subject': 'Test 1', 'sender_email': 'test1@example.com'},
            {'subject': 'Test 2', 'sender_email': 'test2@example.com'}
        ]
        mock_email_processor.return_value = mock_processor
        
        # Configure CLI app
        cli_app.args = MagicMock()
        cli_app.args.folder = 'Inbox'
        cli_app.args.recursive = False
        cli_app.args.include_read = False
        cli_app.args.debug = False
        cli_app.config = {}
        
        # Test
        emails = cli_app._process_emails()
        
        # Verify
        assert len(emails) == 2
        mock_outlook.get_folder.assert_called_once_with('Inbox')
        mock_outlook.get_messages.assert_called_once_with(
            'mock_folder',
            recursive=False,
            include_read=False
        )
        assert mock_processor.process_message.call_count == 2
    
    @patch('main_cli.CSVExporter')
    def test_export_emails_csv(self, mock_csv_exporter, cli_app, tmp_path):
        """Test exporting emails to CSV format."""
        # Setup
        output_file = tmp_path / 'test_export.csv'
        emails = [
            {'subject': 'Test 1', 'sender_email': 'test1@example.com'},
            {'subject': 'Test 2', 'sender_email': 'test2@example.com'}
        ]
        
        mock_exporter = MagicMock()
        mock_exporter.export_emails.return_value = True
        mock_csv_exporter.return_value = mock_exporter
        
        # Configure CLI app
        cli_app.args = MagicMock()
        cli_app.args.output = str(output_file)
        cli_app.args.format = 'csv'
        cli_app.args.include_headers = True
        
        # Test
        result = cli_app._export_emails(emails)
        
        # Verify
        assert result is True
        mock_csv_exporter.assert_called_once()
        mock_exporter.export_emails.assert_called_once_with(
            emails=emails,
            output_path=str(output_file),
            include_headers=True
        )
    
    def test_export_emails_json(self, cli_app, tmp_path):
        """Test exporting emails to JSON format."""
        # Setup
        output_file = tmp_path / 'test_export.json'
        emails = [
            {'subject': 'Test 1', 'sender_email': 'test1@example.com'},
            {'subject': 'Test 2', 'sender_email': 'test2@example.com'}
        ]
        
        # Configure CLI app
        cli_app.args = MagicMock()
        cli_app.args.output = str(output_file)
        cli_app.args.format = 'json'
        
        # Test
        result = cli_app._export_emails(emails)
        
        # Verify
        assert result is True
        assert output_file.exists()
        
        with open(output_file, 'r') as f:
            exported_data = json.load(f)
            assert len(exported_data) == 2
            assert exported_data[0]['subject'] == 'Test 1'
            assert exported_data[1]['sender_email'] == 'test2@example.com'
    
    @patch('main_cli.CLIApp._process_emails')
    @patch('main_cli.CLIApp._export_emails')
    def test_run_success(self, mock_export_emails, mock_process_emails, cli_app):
        """Test successful execution of the CLI app."""
        # Setup
        mock_process_emails.return_value = [{'subject': 'Test'}, {'subject': 'Test 2'}]
        mock_export_emails.return_value = True
        
        cli_app.args = MagicMock()
        cli_app.args.dry_run = False
        cli_app.args.sender = None
        cli_app.args.subject = None
        
        # Test
        result = cli_app.run()
        
        # Verify
        assert result == 0
        mock_process_emails.assert_called_once()
        mock_export_emails.assert_called_once()
    
    @patch('main_cli.CLIApp._process_emails')
    def test_run_dry_run(self, mock_process_emails, cli_app):
        """Test dry run mode."""
        # Setup
        mock_process_emails.return_value = [{'subject': 'Test'}]
        
        cli_app.args = MagicMock()
        cli_app.args.dry_run = True
        cli_app.args.sender = None
        cli_app.args.subject = None
        
        # Test
        result = cli_app.run()
        
        # Verify
        assert result == 0
        mock_process_emails.assert_called_once()
        # _export_emails should not be called in dry run mode
        assert not hasattr(mock_process_emails, 'export_emails')
    
    @patch('main_cli.CLIApp._process_emails')
    @patch('main_cli.CLIApp._export_emails')
    def test_run_no_emails(self, mock_export_emails, mock_process_emails, cli_app):
        """Test running with no emails to process."""
        # Setup
        mock_process_emails.return_value = []
        
        cli_app.args = MagicMock()
        cli_app.args.dry_run = False
        cli_app.args.sender = None
        cli_app.args.subject = None
        
        # Test
        result = cli_app.run()
        
        # Verify
        assert result == 0
        mock_process_emails.assert_called_once()
        mock_export_emails.assert_not_called()
