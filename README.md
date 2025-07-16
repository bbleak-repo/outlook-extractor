# Outlook Email Extractor

A cross-platform application for extracting and managing emails from Microsoft Outlook with a user-friendly interface. Now with full macOS support and enhanced export capabilities.

## Features

- **Cross-Platform**: Fully supported on both Windows and macOS
- **Email Extraction**: Extract emails with various filtering options
- **Threaded Conversations**: View and manage email threads
- **Multiple Storage Backends**: SQLite and JSON support
- **Advanced Export**: Export to CSV with customizable templates
- **Data Security**: Secure handling of sensitive information
- **Performance Optimized**: Lazy loading for better performance
- **Modern UI**: Clean, responsive interface with theme support

## Prerequisites

- Python 3.8 or higher
- Microsoft Outlook (Windows) or Outlook for Mac
- Tkinter (usually included with Python)

## Installation

### Prerequisites
- Python 3.8 or higher
- Tkinter (usually included with Python)
- On macOS: Ensure Xcode command line tools are installed (`xcode-select --install`)

### Quick Start

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd outlook-extract
   ```

2. Create and activate a virtual environment:
   ```bash
   # Windows
   python -m venv venv
   .\venv\Scripts\activate

   # macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

### macOS Notes
For optimal performance on macOS:
- Use Python 3.9 or later
- Ensure you have the latest version of PySimpleGUI
- If you encounter any UI issues, try running with the `--no-console` flag

## Usage

### Basic Usage
Run the application:
```bash
# Standard mode
python -m outlook_extractor.run

# With custom config
python -m outlook_extractor.run --config /path/to/config.ini

# On macOS (without console window)
pythonw -m outlook_extractor.run
```

### Command Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `--config` | Path to a custom configuration file | `--config ~/custom_config.ini` |
| `--debug` | Enable debug mode | `--debug` |
| `--no-console` | Run without console window (macOS) | `--no-console` |
| `--version` | Show version information | `--version` |

## Configuration

The application comes with default settings that can be customized through the UI or by editing the configuration file. The configuration file is automatically created in the user's application data directory.

### Windows
`%APPDATA%\outlook_extractor\config.ini`

### macOS
`~/Library/Application Support/outlook_extractor/config.ini`

## Project Structure

```
outlook-extract/
├── outlook_extractor/
│   ├── __init__.py
│   ├── config.py          # Configuration management
│   ├── run.py             # Main entry point
│   ├── core/              # Core functionality
│   │   ├── __init__.py
│   │   ├── outlook_client.py
│   │   └── platform/      # Platform-specific implementations
│   └── ui/                # User interface components
│       ├── __init__.py
│       ├── main_window.py
│       ├── export_tab.py  # Export functionality
│       └── components/    # Reusable UI components
├── tests/                 # Unit and integration tests
│   ├── unit/
│   └── integration/
├── docs/                  # Documentation
├── requirements.txt       # Project dependencies
├── CHANGELOG.md          # Version history
└── README.md             # This file
```

## Exporting Emails

The application supports exporting emails in various formats:

1. **CSV Export**: Export email data to CSV files with customizable fields
2. **Template-Based Export**: Create custom export templates
3. **Bulk Export**: Export multiple folders at once

### Export Options
- Include/exclude email headers
- Filter by date range
- Export attachments
- Clean HTML formatting
- Include AI-generated summaries

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For support, please open an issue in the project's issue tracker.
