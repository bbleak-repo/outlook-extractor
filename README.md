# Outlook Email Extractor

A cross-platform application for extracting and managing emails from Microsoft Outlook with a user-friendly interface.

## Features

- Extract emails from Outlook with various filtering options
- Threaded conversation view
- Multiple storage backends (SQLite, JSON)
- Configurable data extraction (attachments, embedded images, links, etc.)
- Secure handling of sensitive information
- Cross-platform support (Windows, macOS)

## Prerequisites

- Python 3.8 or higher
- Microsoft Outlook (Windows) or Outlook for Mac
- Tkinter (usually included with Python)

## Installation

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

## Usage

Run the application:
```bash
python -m outlook_extractor.run
```

### Command Line Arguments

- `--config`: Path to a custom configuration file (optional)

Example:
```bash
python -m outlook_extractor.run --config /path/to/config.ini
```

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
│   │   └── outlook_client.py
│   └── ui/                # User interface components
│       ├── __init__.py
│       └── main_window.py
├── tests/                 # Unit and integration tests
├── requirements.txt       # Project dependencies
└── README.md             # This file
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

For support, please open an issue in the project's issue tracker.
