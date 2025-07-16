# Changelog

All notable changes to the Outlook Email Extractor project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0] - 2025-07-15

### Added
- macOS compatibility layer with platform-specific implementations
- Lazy loading for the Export tab to improve initial load performance
- Enhanced error handling and logging throughout the application
- Export functionality for saving email data to CSV format
- Support for custom export templates and formats
- Automated testing framework for GUI components
- Comprehensive documentation updates

### Changed
- Refactored UI components for better cross-platform compatibility
- Improved error messages and user feedback
- Optimized memory usage for large email exports
- Updated configuration handling to be more robust
- Enhanced security for handling sensitive data
- Improved macOS window management and theming

### Fixed
- Fixed layout issues in the Export tab on macOS
- Resolved configuration saving/loading problems
- Fixed threading issues in the email extraction process
- Addressed memory leaks in the UI components
- Fixed file path handling for cross-platform compatibility
- Resolved issues with special characters in email subjects and bodies

## [1.0.0] - 2025-07-14

### Added
- Initial release of Outlook Email Extractor
- Cross-platform GUI using PySimpleGUI
- Configuration management system
- Support for SQLite and JSON storage backends
- Threaded email processing
- Basic security features including data redaction

### Fixed
- N/A

### Deprecated
- N/A

### Removed
- N/A

### Security
- N/A

## [1.0.0] - 2025-07-14

### Added
- Initial release of Outlook Email Extractor

## Versioning Policy

Given a version number MAJOR.MINOR.PATCH, increment the:

1. **MAJOR** version when you make incompatible API changes
2. **MINOR** version when you add functionality in a backward compatible manner
3. **PATCH** version when you make backward compatible bug fixes

Additional labels for pre-release and build metadata are available as extensions to the MAJOR.MINOR.PATCH format.

## Security

Security issues can be reported via email to [SECURITY_EMAIL]. Please include "[SECURITY]" in the subject line.

## Contributing

Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute to this project.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
