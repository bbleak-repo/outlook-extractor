# PySimpleGUI to FreeSimpleGUI Migration

## Summary of Changes

1. **Updated Imports**:
   - Changed all occurrences of `import PySimpleGUI as sg` to `import FreeSimpleGUI as sg`
   - Updated any `from PySimpleGUI import` statements to `from FreeSimpleGUI import`

2. **Updated Dependencies**:
   - Changed `PySimpleGUI>=4.60.0` to `FreeSimpleGUI>=5.0.0` in requirements.txt

3. **Files Modified**:
   - `outlook_extractor/logging_config.py`
   - `outlook_extractor/ui/export_tab.py`
   - `outlook_extractor/ui/logging_ui.py`
   - `outlook_extractor/ui/main_window.py`
   - `outlook_extractor/ui/update_dialog.py`
   - `requirements.txt`
   - `run_mac.py`
   - `test_all_tabs.py`
   - `test_export_tab.py`

## Testing Notes

1. **Recommended Tests**:
   - Open and close all windows/dialogs
   - Test all buttons and interactive elements
   - Verify that logging works as expected
   - Test file operations (open/save dialogs)
   - Check for any UI layout issues

2. **Known Limitations**:
   - Some PySimpleGUI features may have different behavior in FreeSimpleGUI
   - Custom themes might need adjustment
   - System tray functionality may differ

## Next Steps

1. Install the new dependency:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the application and test all functionality

3. If any issues are found, they can be fixed by:
   - Checking the FreeSimpleGUI documentation for differences
   - Adjusting UI layout if needed
   - Updating any deprecated or changed function calls

## Rollback Plan

If needed, you can rollback to the previous version using git:

```bash
git checkout backup-before-fsg-migration
```
