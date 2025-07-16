"""
Update dialog for the Outlook Email Extractor application.
"""

import FreeSimpleGUI as sg
from pathlib import Path
import sys
import logging
from typing import Optional, Callable, Any, Dict

from .. import __version__
from ..auto_updater import AutoUpdater, UpdateError

logger = logging.getLogger(__name__)

class UpdateDialog:
    """Handles the update notification and installation UI."""
    
    def __init__(self, repo_owner: str, repo_name: str, current_version: str = None):
        """Initialize the update dialog.
        
        Args:
            repo_owner: GitHub repository owner (username or organization)
            repo_name: GitHub repository name
            current_version: Current application version (defaults to __version__)
        """
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.current_version = current_version or __version__
        self.updater = AutoUpdater(repo_owner, repo_name, self.current_version)
        self.window = None
    
    def check_for_updates(self, silent: bool = False) -> bool:
        """Check for updates and show the update dialog if available.
        
        Args:
            silent: If True, don't show any UI if no update is available
            
        Returns:
            bool: True if an update is available and the user wants to install it
        """
        try:
            update_available, result = self.updater.check_and_apply_update(auto_apply=False)
            
            if not update_available:
                if not silent:
                    sg.popup_ok(result, title="No Updates Available")
                return False
                
            # Show update dialog
            return self.show_update_dialog(result)
            
        except Exception as e:
            logger.error(f"Error checking for updates: {e}", exc_info=True)
            if not silent:
                sg.popup_error(f"Error checking for updates: {e}", title="Update Error")
            return False
    
    def show_update_dialog(self, update_message: str) -> bool:
        """Show the update dialog.
        
        Args:
            update_message: Message to display about the available update
            
        Returns:
            bool: True if the user wants to install the update
        """
        layout = [
            [sg.Text("A new version is available!")],
            [sg.Multiline(update_message, size=(50, 10), disabled=True, autoscroll=True)],
            [sg.Button("Install Update", key="-INSTALL-"),
             sg.Button("Remind Me Later", key="-LATER-"),
             sg.Button("Skip This Version", key="-SKIP-")]
        ]
        
        self.window = sg.Window("Update Available", layout, modal=True)
        
        while True:
            event, _ = self.window.read()
            
            if event in (sg.WIN_CLOSED, "-LATER-"):
                self.window.close()
                return False
                
            elif event == "-INSTALL-":
                self.window.close()
                return self._install_update()
                
            elif event == "-SKIP-":
                # Here you could implement logic to skip this version
                # For example, save the version number to a config file
                self.window.close()
                return False
    
    def _install_update(self) -> bool:
        """Download and install the update."""
        # Create a progress window
        progress_layout = [
            [sg.Text("Downloading update...")],
            [sg.ProgressBar(100, orientation='h', size=(40, 20), key='-PROGRESS-')],
            [sg.Text("", key='-STATUS-')],
            [sg.Cancel()]
        ]
        
        progress_window = sg.Window("Updating Application", progress_layout, modal=True, finalize=True)
        progress_bar = progress_window['-PROGRESS-']
        status_text = progress_window['-STATUS-']
        
        def update_progress(downloaded: int, total: int):
            """Update the progress bar."""
            if total > 0:
                percent = int((downloaded / total) * 100)
                progress_bar.update(percent)
                status_text.update(f"Downloaded {downloaded/1024/1024:.1f}MB of {total/1024/1024:.1f}MB")
                progress_window.refresh()
        
        try:
            # Download the update
            update_file = self.updater.download_update(progress_callback=update_progress)
            
            # Close the progress window
            progress_window.close()
            
            # Ask for confirmation before applying the update
            if sg.popup_yes_no("The update has been downloaded. Apply it now?\nThe application will restart after the update.", 
                              title="Apply Update") == "Yes":
                
                # Apply the update
                if self.updater.apply_update(update_file):
                    sg.popup_ok("Update applied successfully. Please restart the application.", 
                               title="Update Complete")
                    return True
                else:
                    sg.popup_ok("Update downloaded. Please install it manually.", 
                               title="Manual Installation Required")
                    return False
                    
        except UpdateError as e:
            progress_window.close()
            sg.popup_error(f"Error installing update: {e}", title="Update Error")
            return False
            
        except Exception as e:
            logger.exception("Unexpected error during update")
            progress_window.close()
            sg.popup_error(f"An unexpected error occurred: {e}", title="Update Error")
            return False
            
        finally:
            if progress_window and not progress_window.was_closed():
                progress_window.close()
                
        return False

def check_for_updates(parent_window=None, repo_owner: str = None, repo_name: str = None, 
                     current_version: str = None, silent: bool = False) -> bool:
    """Helper function to check for updates.
    
    Args:
        parent_window: Parent window for the update dialog
        repo_owner: GitHub repository owner (username or organization)
        repo_name: GitHub repository name
        current_version: Current application version (defaults to __version__)
        silent: If True, don't show any UI if no update is available
        
    Returns:
        bool: True if an update is available and the user wants to install it
    """
    if not repo_owner or not repo_name:
        # Default to the current repository
        repo_owner = "bbleak-repo"
        repo_name = "outlook-extractor"
    
    dialog = UpdateDialog(repo_owner, repo_name, current_version)
    return dialog.check_for_updates(silent=silent)

if __name__ == "__main__":
    # Example usage
    logging.basicConfig(level=logging.INFO)
    check_for_updates()
