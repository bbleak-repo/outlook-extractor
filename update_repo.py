#!/usr/bin/env python3
"""
GitHub Repository Updater Script

This script simplifies the process of updating your GitHub repository.
It will:
1. Stage all changes
2. Create a commit with a timestamp
3. Push changes to the remote repository

Usage:
    python update_repo.py [commit_message]
    
If no commit message is provided, a default one with timestamp will be used.
"""

import subprocess
import sys
from datetime import datetime
import os

def run_command(command, cwd=None):
    """Run a shell command and return the output."""
    try:
        result = subprocess.run(
            command,
            cwd=cwd,
            shell=True,
            check=True,
            text=True,
            capture_output=True
        )
        return result.stdout.strip()
    except subprocess.CalledProcessError as e:
        print(f"Error executing command: {command}")
        print(f"Error: {e.stderr}")
        sys.exit(1)

def get_git_status():
    """Get the current git status."""
    return run_command("git status --porcelain")

def update_repository(commit_message=None):
    """Update the git repository with the given commit message."""
    # Get repository root
    repo_root = os.path.dirname(os.path.abspath(__file__))
    
    # Check if this is a git repository
    if not os.path.exists(os.path.join(repo_root, '.git')):
        print("Error: This is not a git repository.")
        print("Please run 'git init' in this directory first.")
        sys.exit(1)
    
    # Check for changes
    status = get_git_status()
    if not status:
        print("No changes to commit.")
        return
    
    # Create commit message if not provided
    if not commit_message:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_message = f"Update: {timestamp}"
    
    try:
        # Stage all changes
        print("Staging changes...")
        run_command("git add .", cwd=repo_root)
        
        # Commit changes
        print("Creating commit...")
        run_command(f'git commit -m "{commit_message}"', cwd=repo_root)
        
        # Push changes
        print("Pushing to remote repository...")
        run_command("git push", cwd=repo_root)
        
        print("\n✅ Successfully updated the repository!")
        print(f"Commit: {commit_message}")
        
    except Exception as e:
        print(f"\n❌ Error updating repository: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    # Get commit message from command line arguments if provided
    commit_message = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else None
    update_repository(commit_message)
