"""
Auto-update functionality for Outlook Email Extractor.

This module handles checking for updates from GitHub releases and applying them.
"""

import os
import sys
import logging
import json
import tempfile
import platform
import subprocess
from pathlib import Path
from typing import Optional, Tuple
import requests
from packaging import version
import semver

# Configure logging
logger = logging.getLogger(__name__)

class UpdateError(Exception):
    """Custom exception for update-related errors."""
    pass

class AutoUpdater:
    """Handles checking for and applying updates from GitHub releases.
    
    Args:
        repo_owner: GitHub repository owner (username or organization)
        repo_name: GitHub repository name
        current_version: Current application version (e.g., '1.0.0')
        app_name: Name of the application (for logging and messages)
    """
    
    GITHUB_API = "https://api.github.com"
    
    def __init__(self, repo_owner: str, repo_name: str, current_version: str, app_name: str = "Outlook Email Extractor"):
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.current_version = current_version
        self.app_name = app_name
        self.update_info = None
        
        # Set up paths
        self.app_dir = Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.update_dir = self.app_dir / "updates"
        self.update_dir.mkdir(exist_ok=True)
        
        # Configure requests session
        self.session = requests.Session()
        self.session.headers.update({
            'Accept': 'application/vnd.github.v3+json',
            'User-Agent': f'{app_name}/{current_version}'
        })
    
    def check_for_updates(self) -> Tuple[bool, Optional[dict]]:
        """Check if a new version is available.
        
        Returns:
            tuple: (update_available, update_info)
                - update_available: True if an update is available
                - update_info: Dictionary with update details if available, None otherwise
        """
        try:
            logger.info("Checking for updates...")
            
            # Get the latest release from GitHub
            url = f"{self.GITHUB_API}/repos/{self.repo_owner}/{self.repo_name}/releases/latest"
            response = self.session.get(url, timeout=10)
            response.raise_for_status()
            
            latest_release = response.json()
            latest_version = latest_release.get('tag_name', '').lstrip('v')
            
            if not latest_version:
                logger.warning("No version tag found in the latest release")
                return False, None
            
            # Compare versions
            try:
                current = semver.VersionInfo.parse(self.current_version.lstrip('v'))
                latest = semver.VersionInfo.parse(latest_version)
                
                if latest > current:
                    logger.info(f"Update available: {self.current_version} -> {latest_version}")
                    self.update_info = {
                        'version': latest_version,
                        'release_notes': latest_release.get('body', ''),
                        'download_url': None,
                        'published_at': latest_release.get('published_at', '')
                    }
                    
                    # Find the appropriate asset for the current platform
                    for asset in latest_release.get('assets', []):
                        asset_name = asset.get('name', '').lower()
                        if self._is_compatible_asset(asset_name):
                            self.update_info['download_url'] = asset.get('browser_download_url')
                            break
                    
                    return True, self.update_info
                
                logger.info("No updates available")
                return False, None
                
            except ValueError as e:
                logger.error(f"Error comparing versions: {e}")
                return False, None
                
        except requests.RequestException as e:
            logger.error(f"Error checking for updates: {e}")
            return False, None
    
    def _is_compatible_asset(self, asset_name: str) -> bool:
        """Check if an asset is compatible with the current platform."""
        system = platform.system().lower()
        
        if system == 'windows':
            return asset_name.endswith('.exe') or asset_name.endswith('.msi')
        elif system == 'darwin':
            return '.dmg' in asset_name or '.pkg' in asset_name or 'macos' in asset_name
        else:
            return '.deb' in asset_name or '.rpm' in asset_name or 'linux' in asset_name
    
    def download_update(self, progress_callback=None) -> str:
        """Download the update package.
        
        Args:
            progress_callback: Optional callback function for download progress
                function(downloaded: int, total: int)
                
        Returns:
            str: Path to the downloaded update file
            
        Raises:
            UpdateError: If there's an error downloading the update
        """
        if not self.update_info or not self.update_info.get('download_url'):
            raise UpdateError("No update available or no download URL found")
        
        try:
            download_url = self.update_info['download_url']
            file_name = os.path.basename(download_url.split('?')[0])  # Remove query params
            temp_file = self.update_dir / file_name
            
            logger.info(f"Downloading update from {download_url}")
            
            # Stream the download to support progress tracking
            with self.session.get(download_url, stream=True, timeout=30) as r:
                r.raise_for_status()
                total_size = int(r.headers.get('content-length', 0))
                
                with open(temp_file, 'wb') as f:
                    downloaded = 0
                    for chunk in r.iter_content(chunk_size=8192):
                        if chunk:  # filter out keep-alive new chunks
                            f.write(chunk)
                            downloaded += len(chunk)
                            if progress_callback and total_size > 0:
                                progress_callback(downloaded, total_size)
            
            logger.info(f"Update downloaded to {temp_file}")
            return str(temp_file)
            
        except Exception as e:
            raise UpdateError(f"Error downloading update: {str(e)}")
    
    def apply_update(self, update_file: str) -> bool:
        """Apply the downloaded update.
        
        Args:
            update_file: Path to the downloaded update file
            
        Returns:
            bool: True if the update was applied successfully
            
        Raises:
            UpdateError: If there's an error applying the update
        """
        try:
            update_file = Path(update_file)
            if not update_file.exists():
                raise UpdateError(f"Update file not found: {update_file}")
            
            logger.info(f"Applying update from {update_file}")
            
            # Determine the appropriate action based on file type and platform
            if platform.system() == 'Windows':
                if update_file.suffix.lower() == '.exe':
                    # For Windows executable installers
                    subprocess.Popen([str(update_file), '/SILENT', '/NORESTART'])
                    return True
            elif platform.system() == 'Darwin':  # macOS
                if update_file.suffix.lower() == '.dmg':
                    # For macOS disk images
                    subprocess.run(['hdiutil', 'attach', str(update_file)], check=True)
                    # Additional steps would be needed to install from the mounted DMG
                    return True
                elif update_file.suffix.lower() == '.pkg':
                    # For macOS package installers
                    subprocess.run(['sudo', 'installer', '-pkg', str(update_file), '-target', '/'], check=True)
                    return True
            
            # For other cases, just log that manual installation is needed
            logger.warning(f"Automatic installation not supported for {update_file}. Please install manually.")
            return False
            
        except Exception as e:
            raise UpdateError(f"Error applying update: {str(e)}")
    
    def check_and_apply_update(self, auto_apply: bool = False) -> Tuple[bool, str]:
        """Check for updates and optionally apply them.
        
        Args:
            auto_apply: If True, automatically download and apply the update
            
        Returns:
            tuple: (update_available, message)
                - update_available: True if an update was available
                - message: Status message for the user
        """
        try:
            update_available, update_info = self.check_for_updates()
            
            if not update_available:
                return False, "You're running the latest version."
            
            if not auto_apply:
                return True, f"Version {update_info['version']} is available. Would you like to update now?"
            
            # Auto-apply the update
            try:
                update_file = self.download_update()
                if self.apply_update(update_file):
                    return True, f"Successfully updated to version {update_info['version']}. Please restart the application."
                else:
                    return True, "Update downloaded. Please install it manually."
            except UpdateError as e:
                return True, f"Error applying update: {str(e)}"
                
        except Exception as e:
            logger.exception("Error during update check")
            return False, f"Error checking for updates: {str(e)}"
