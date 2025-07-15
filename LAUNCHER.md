# Outlook Email Extractor - Launcher

This launcher script (`launch.py`) simplifies the process of running the Outlook Email Extractor by automatically handling dependency installation and launching the application.

## Prerequisites

- Python 3.8 or higher
- pip (Python package installer)
- Tkinter (usually included with Python, but may need to be installed separately on some Linux distributions)

## Quick Start

1. **Make the launcher executable** (Linux/macOS):
   ```bash
   chmod +x launch.py
   ```

2. **Run the launcher**:
   ```bash
   ./launch.py
   ```
   
   Or on Windows:
   ```
   python launch.py
   ```

The launcher will:
1. Check for required Python packages
2. Install any missing dependencies
3. Launch the Outlook Email Extractor UI

## Command Line Options

- `--install-deps`: Only install dependencies without launching the application
- `--no-update`: Skip checking for dependency updates

Example:
```bash
./launch.py --install-deps
```

## Troubleshooting

### Missing Tkinter
If you get an error about Tkinter not being found, install it using your system package manager:

- **Ubuntu/Debian**:
  ```bash
  sudo apt-get install python3-tk
  ```

- **Fedora/RHEL/CentOS**:
  ```bash
  sudo dnf install python3-tkinter
  ```

- **macOS**:
  ```bash
  brew install python-tk
  ```

### Permission Denied
If you get a permission denied error when running the launcher, try:
```bash
python3 launch.py
```

## Manual Installation

If you prefer to install dependencies manually:

1. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python -m outlook_extractor.run
   ```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
