#!/usr/bin/env python3
"""
Script to migrate from PySimpleGUI to FreeSimpleGUI.
"""
import os
import re
import sys
from pathlib import Path

# Files to process
PYTHON_FILES = [
    "test_export_tab.py",
    "outlook_extractor/ui/main_window.py",
    "outlook_extractor/ui/logging_ui.py",
    "outlook_extractor/ui/update_dialog.py",
    "outlook_extractor/ui/export_tab.py",
    "outlook_extractor/logging_config.py",
    "run_mac.py",
    "test_all_tabs.py"
]

def update_imports(file_path):
    """Update imports in a single file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace import statements
        new_content = content.replace(
            'import PySimpleGUI as sg',
            'import FreeSimpleGUI as sg'
        )
        
        # Replace any from imports
        new_content = new_content.replace(
            'from PySimpleGUI import',
            'from FreeSimpleGUI import'
        )
        
        # Check for any PySimpleGUI-specific features that might need attention
        checks = [
            ('sg.theme_previewer()', 'Not available in FreeSimpleGUI'),
            ('sg.theme_add_new', 'Custom theme addition not supported in FreeSimpleGUI'),
            ('sg.SystemTray', 'SystemTray has different implementation in FreeSimpleGUI'),
            ('sg.Print', 'Print function may have different behavior'),
            ('sg.cprint', 'cprint function may have different behavior'),
            ('sg.Clipboard', 'Clipboard handling may be different')
        ]
        
        issues = []
        for pattern, message in checks:
            if pattern in new_content:
                issues.append(f"  - {message}")
               
        # Write changes if there are any
        if new_content != content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f"Updated: {file_path}")
            if issues:
                print("  Potential issues found:")
                print("\n".join(issues))
        else:
            print(f"No changes needed: {file_path}")
            
    except Exception as e:
        print(f"Error processing {file_path}: {str(e)}")

def update_requirements():
    """Update requirements.txt to use FreeSimpleGUI."""
    req_file = "requirements.txt"
    try:
        with open(req_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Replace PySimpleGUI with FreeSimpleGUI
        new_content = content.replace(
            'PySimpleGUI>=4.60.0',
            'FreeSimpleGUI>=5.0.0'  # Using latest stable version
        )
        
        if new_content != content:
            with open(req_file, 'w', encoding='utf-8') as f:
                f.write(new_content)
            print(f"Updated: {req_file}")
        else:
            print(f"No changes needed: {req_file}")
            
    except Exception as e:
        print(f"Error updating {req_file}: {str(e)}")

def main():
    print("Starting migration from PySimpleGUI to FreeSimpleGUI...")
    
    # Process all Python files
    base_dir = Path(__file__).parent
    for file_path in PYTHON_FILES:
        full_path = base_dir / file_path
        if full_path.exists():
            update_imports(full_path)
        else:
            print(f"Warning: File not found: {file_path}")
    
    # Update requirements
    update_requirements()
    
    print("\nMigration complete!")
    print("Please review the changes and test the application thoroughly.")
    print("Note: Some PySimpleGUI features may have different behavior in FreeSimpleGUI.")

if __name__ == "__main__":
    main()
