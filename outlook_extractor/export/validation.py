"""Validation utilities for export operations."""
import hashlib
import json
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Callable
from dataclasses import dataclass
import logging

logger = logging.getLogger(__name__)

@dataclass
class ExportValidationResult:
    """Result of an export validation."""
    is_valid: bool
    message: str
    details: Dict[str, Any]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            'is_valid': self.is_valid,
            'message': self.message,
            'details': self.details
        }

class ExportValidator:
    """Validates exported files and their contents."""
    
    @classmethod
    def validate_export(
        cls,
        file_path: Path,
        expected_format: str,
        expected_count: Optional[int] = None,
        checksum_algorithm: str = 'sha256'
    ) -> ExportValidationResult:
        """Validate an exported file.
        
        Args:
            file_path: Path to the exported file
            expected_format: Expected format ('csv' or 'excel')
            expected_count: Expected number of records (optional)
            checksum_algorithm: Algorithm for checksum calculation
            
        Returns:
            ExportValidationResult with validation details
        """
        if not file_path.exists():
            return ExportValidationResult(
                False,
                f"Export file not found: {file_path}",
                {'error': 'file_not_found'}
            )
            
        try:
            # Basic file validation
            file_size = file_path.stat().st_size
            if file_size == 0:
                return ExportValidationResult(
                    False,
                    "Export file is empty",
                    {'error': 'empty_file'}
                )
                
            # Format-specific validation
            format_lower = expected_format.lower()
            if format_lower == 'csv':
                return cls._validate_csv(file_path, expected_count, checksum_algorithm)
            elif format_lower == 'excel':
                return cls._validate_excel(file_path, expected_count, checksum_algorithm)
            elif format_lower == 'json':
                return cls._validate_json(file_path, expected_count, checksum_algorithm)
            else:
                return ExportValidationResult(
                    False,
                    f"Unsupported format: {expected_format}",
                    {'error': 'unsupported_format'}
                )
                
        except Exception as e:
            logger.exception(f"Error validating export file: {file_path}")
            return ExportValidationResult(
                False,
                f"Error validating export: {str(e)}",
                {'error': 'validation_error', 'exception': str(e)}
            )
    
    @classmethod
    def _validate_csv(
        cls,
        file_path: Path,
        expected_count: Optional[int],
        checksum_algorithm: str
    ) -> ExportValidationResult:
        """Validate a CSV export file."""
        import csv
        
        details = {
            'file_size': file_path.stat().st_size,
            'line_count': 0,
            'record_count': 0,
            'headers': [],
            'checksum': None
        }
        
        # Calculate checksum
        details['checksum'] = cls._calculate_checksum(file_path, checksum_algorithm)
        
        # Count lines and validate CSV structure
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                details['headers'] = next(reader, [])  # Read header
                details['record_count'] = sum(1 for _ in reader)
                details['line_count'] = details['record_count'] + 1  # +1 for header
        except Exception as e:
            return ExportValidationResult(
                False,
                f"Invalid CSV file: {str(e)}",
                {**details, 'error': 'invalid_csv', 'exception': str(e)}
            )
        
        # Validate record count if expected_count is provided
        if expected_count is not None and details['record_count'] != expected_count:
            return ExportValidationResult(
                False,
                f"Record count mismatch. Expected {expected_count}, got {details['record_count']}",
                {**details, 'error': 'count_mismatch'}
            )
            
        return ExportValidationResult(
            True,
            f"CSV validation successful. Found {details['record_count']} records.",
            details
        )
    
    @classmethod
    def _validate_excel(
        cls,
        file_path: Path,
        expected_count: Optional[int],
        checksum_algorithm: str
    ) -> ExportValidationResult:
        """Validate an Excel export file."""
        try:
            import pandas as pd
        except ImportError:
            return ExportValidationResult(
                False,
                "pandas is required for Excel validation",
                {'error': 'missing_dependency'}
            )
            
        details = {
            'file_size': file_path.stat().st_size,
            'sheet_count': 0,
            'sheets': {},
            'checksum': None
        }
        
        # Calculate checksum
        details['checksum'] = cls._calculate_checksum(file_path, checksum_algorithm)
        
        try:
            # Read Excel file
            with pd.ExcelFile(file_path) as xls:
                details['sheet_count'] = len(xls.sheet_names)
                
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    sheet_details = {
                        'row_count': len(df),
                        'column_count': len(df.columns),
                        'headers': list(df.columns)
                    }
                    details['sheets'][sheet_name] = sheet_details
                    
                    # Validate record count for the first sheet if expected_count is provided
                    if sheet_name == xls.sheet_names[0] and expected_count is not None:
                        if len(df) != expected_count:
                            return ExportValidationResult(
                                False,
                                f"Record count mismatch in sheet '{sheet_name}'. "
                                f"Expected {expected_count}, got {len(df)}",
                                {**details, 'error': 'count_mismatch'}
                            )
            
            return ExportValidationResult(
                True,
                f"Excel validation successful. Found {details['sheet_count']} sheets.",
                details
            )
            
        except Exception as e:
            return ExportValidationResult(
                False,
                f"Invalid Excel file: {str(e)}",
                {**details, 'error': 'invalid_excel', 'exception': str(e)}
            )
    
    @staticmethod
    def _calculate_checksum(file_path: Path, algorithm: str) -> str:
        """Calculate file checksum using the specified algorithm."""
        hash_func = getattr(hashlib, algorithm.lower(), None)
        if not hash_func:
            raise ValueError(f"Unsupported hash algorithm: {algorithm}")
            
        h = hash_func()
        with open(file_path, 'rb') as f:
            while chunk := f.read(8192):
                h.update(chunk)
                
        return h.hexdigest()
    
    @classmethod
    def _validate_json(
        cls,
        file_path: Path,
        expected_count: Optional[int],
        checksum_algorithm: str
    ) -> ExportValidationResult:
        """Validate a JSON export file."""
        details = {
            'file_size': file_path.stat().st_size,
            'record_count': 0,
            'is_valid_json': False,
            'has_metadata': False,
            'checksum': None
        }
        
        # Calculate checksum
        details['checksum'] = cls._calculate_checksum(file_path, checksum_algorithm)
        
        try:
            # Read and parse JSON file
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            details['is_valid_json'] = True
            
            # Check if it has the expected structure
            if not isinstance(data, dict):
                return ExportValidationResult(
                    False,
                    "Invalid JSON structure: Expected an object",
                    {**details, 'error': 'invalid_structure'}
                )
            
            # Check for metadata
            if 'metadata' in data and isinstance(data['metadata'], dict):
                details['has_metadata'] = True
                details['metadata'] = data['metadata']
            
            # Check for emails array
            if 'emails' not in data or not isinstance(data['emails'], list):
                return ExportValidationResult(
                    False,
                    "Invalid JSON structure: Missing or invalid 'emails' array",
                    {**details, 'error': 'missing_emails'}
                )
            
            details['record_count'] = len(data['emails'])
            
            # Validate record count if expected_count is provided
            if expected_count is not None and details['record_count'] != expected_count:
                return ExportValidationResult(
                    False,
                    f"Record count mismatch. Expected {expected_count}, got {details['record_count']}",
                    {**details, 'error': 'count_mismatch'}
                )
            
            # Sample a few records to check structure
            sample_size = min(10, details['record_count'])
            if sample_size > 0:
                sample_indices = set()
                if details['record_count'] > 0:
                    # Sample from beginning, middle, and end
                    sample_indices.update({0, details['record_count'] // 2, details['record_count'] - 1})
                    # Add a few random samples if available
                    if details['record_count'] > 3:
                        import random
                        sample_indices.update(random.sample(
                            range(1, details['record_count'] - 1),
                            min(3, details['record_count'] - 3)
                        ))
                
                details['sample_records'] = [
                    data['emails'][i] 
                    for i in sorted(sample_indices) 
                    if i < details['record_count']
                ]
            
            return ExportValidationResult(
                True,
                f"JSON validation successful. Found {details['record_count']} records.",
                details
            )
            
        except json.JSONDecodeError as e:
            return ExportValidationResult(
                False,
                f"Invalid JSON: {str(e)}",
                {**details, 'error': 'invalid_json', 'exception': str(e)}
            )
        except Exception as e:
            return ExportValidationResult(
                False,
                f"Error validating JSON: {str(e)}",
                {**details, 'error': 'validation_error', 'exception': str(e)}
            )
    
    @classmethod
    def generate_validation_report(
        cls,
        validation_results: Dict[str, ExportValidationResult],
        output_path: Optional[Path] = None
    ) -> str:
        """Generate a human-readable validation report.
        
        Args:
            validation_results: Dictionary of validation results
            output_path: Optional path to save the report
            
        Returns:
            Formatted report as a string
        """
        report = []
        report.append("=" * 80)
        report.append("EXPORT VALIDATION REPORT")
        report.append("=" * 80)
        report.append("")
        
        for file_path, result in validation_results.items():
            report.append(f"File: {file_path}")
            report.append(f"Status: {'VALID' if result.is_valid else 'INVALID'}")
            report.append(f"Message: {result.message}")
            
            if not result.is_valid:
                report.append("\nDetails:")
                for key, value in result.details.items():
                    if key not in ['error', 'exception']:
                        report.append(f"  {key}: {value}")
                
                if 'error' in result.details:
                    report.append(f"\nError: {result.details['error']}")
                if 'exception' in result.details:
                    report.append(f"Exception: {result.details['exception']}")
            
            report.append("\n" + "-" * 40 + "\n")
        
        # Add summary
        total = len(validation_results)
        valid = sum(1 for r in validation_results.values() if r.is_valid)
        report.append("=" * 80)
        report.append(f"SUMMARY: {valid} of {total} files validated successfully")
        report.append("=" * 80)
        
        report_text = "\n".join(report)
        
        # Save to file if output path is provided
        if output_path:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(report_text)
                
        return report_text
