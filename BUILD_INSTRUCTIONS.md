# Outlook Extractor - Build Instructions

This document provides comprehensive instructions for building the Outlook Extractor application on both Windows and macOS platforms.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Building on Windows](#building-on-windows)
- [Building on macOS](#building-on-macos)
- [Code Signing & Notarization](#code-signing--notarization)
- [Troubleshooting](#troubleshooting)
- [Build Artifacts](#build-artifacts)

## Prerequisites

### Common Requirements
- Python 3.8 or higher
- pip (Python package manager)
- Git
- Virtual environment (recommended)

### Platform-Specific Requirements

#### Windows
- Visual Studio Build Tools with C++ workload
- Windows 10/11 SDK

#### macOS
- Xcode Command Line Tools
- Python 3.8+ installed via Homebrew
- Xcode (for code signing and notarization)

## Building on Windows

### 1. Set Up Environment
```powershell
# Clone the repository
git clone https://github.com/yourusername/outlook-extractor.git
cd outlook-extractor

# Create and activate virtual environment
python -m venv venv
.\venv\Scripts\activate

# Install build dependencies
pip install -r requirements.txt
pip install pyinstaller
```

### 2. Build the Application
```powershell
# Development build (unsigned)
python build/scripts/build.py --clean

# Production build (signed)
python build/scripts/build.py --clean --sign
```

### 3. Run the Application
```powershell
# From source
python -m outlook_extractor.run

# From built executable
.\dist\OutlookExtractor.exe
```

## Building on macOS

### 1. Set Up Environment
```bash
# Install Homebrew if not already installed
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# Install Python and dependencies
brew install python

# Clone the repository
git clone https://github.com/yourusername/outlook-extractor.git
cd outlook-extractor

# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate

# Install build dependencies
pip install -r requirements.txt
pip install pyinstaller
```

### 2. Build the Application
```bash
# Development build (unsigned)
python3 build/scripts/build.py --clean

# Production build (signed)
python3 build/scripts/build.py --clean --sign-identity "Developer ID Application: Your Name (YOUR_TEAM_ID)"
```

### 3. Run the Application
```bash
# From source
python3 -m outlook_extractor.run

# From built application
open dist/OutlookExtractor.app
```

## Code Signing & Notarization (macOS)

### Prerequisites
1. Apple Developer ID Application certificate in Keychain
2. App-specific password for notarization
3. Valid Apple Developer account with appropriate permissions

### Steps for Notarization

1. **Build with signing identity**:
   ```bash
   python3 build/scripts/build.py --clean --sign-identity "Developer ID Application: Your Name (YOUR_TEAM_ID)"
   ```

2. **Create a ZIP archive for notarization**:
   ```bash
   ditto -c -k --keepParent "dist/OutlookExtractor.app" "dist/OutlookExtractor.zip"
   ```

3. **Upload for notarization**:
   ```bash
   xcrun notarytool submit "dist/OutlookExtractor.zip" \
       --apple-id "your@apple.id" \
       --team-id "YOUR_TEAM_ID" \
       --password "YOUR_APP_SPECIFIC_PASSWORD" \
       --wait
   ```

4. **Staple the ticket** (after successful notarization):
   ```bash
   xcrun stapler staple "dist/OutlookExtractor.app"
   ```

## Troubleshooting

### Common Issues

#### Windows
- **Missing DLL errors**: Install the Visual C++ Redistributable
- **Build failures**: Ensure all build tools are properly installed
- **Permission issues**: Run the command prompt as Administrator

#### macOS
- **Code Signing Fails**:
  - Verify the signing identity exists in Keychain
  - Ensure you have the correct permissions
  - Check that the certificate is valid and not expired

- **Notarization Fails**:
  - Check the notarization log for specific errors
  - Ensure all binaries are signed with hardened runtime
  - Verify entitlements are correctly set in the build script

- **App Won't Launch**:
  - Check Console.app for crash logs
  - Run from terminal for error messages: `/Applications/OutlookExtractor.app/Contents/MacOS/OutlookExtractor`
  - Verify all dependencies are included in the bundle

## Build Artifacts

### Windows
- `dist/OutlookExtractor.exe` - Standalone executable
- `build/` - Intermediate build files

### macOS
- `dist/OutlookExtractor.app` - Application bundle
- `dist/OutlookExtractor.zip` - Archive for notarization
- `build/` - Intermediate build files

## Distribution

For distribution, we recommend:
1. Creating a DMG (macOS) or MSI (Windows) installer
2. Including a comprehensive README file
3. Providing checksums for download verification
4. Following platform-specific distribution guidelines
EOL

zsh: permission denied: BUILD_INSTRUCTIONS.md
Exit Code 1
Open in terminal

Auto-run
I'll help you verify the current build script and provide the necessary updates. Here's what you need to do:

1. Verify the Build Script
The build script (build/scripts/build.py) has been updated to:

Support both signed and unsigned builds
Handle environment variables properly
Include better error handling
Generate build information
2. Build Instructions
Here are the build instructions you can add to your BUILD_INSTRUCTIONS.md:

markdown
CopyInsert
# Outlook Extractor - Build Instructions

## Prerequisites
- Python 3.8+
- pip
- PyInstaller: `pip install pyinstaller`
- All project dependencies: `pip install -r requirements.txt`

## Building the Executable

### Unsigned Build (Development/Testing)
```bash
# Clean and build unsigned
python build/scripts/build.py --clean
Signed Build (Production)
bash
CopyInsert
# Clean and build with code signing
python build/scripts/build.py --clean --sign-identity "Developer ID Application: Your Name (YOUR_TEAM_ID)"
Build Options
CopyInsert
--clean           Clean build artifacts before building
--onefile         Create a single executable file (default: True)
--sign-identity   Apple Developer ID for code signing
--bundle-id       Bundle identifier (default: com.outlook.extractor)
--spec-only       Only create the spec file, do not build
Code Signing & Notarization (macOS)
Prerequisites
Apple Developer ID Application certificate in Keychain
App-specific password for notarization
Steps for Notarization
Build with signing identity (as shown above)
Upload for notarization:
bash
CopyInsert
xcrun notarytool submit "build/dist/OutlookExtractor.zip" \
    --apple-id "your@apple.id" \
    --team-id "YOUR_TEAM_ID" \
    --password "YOUR_APP_SPECIFIC_PASSWORD" \
    --wait
Staple the ticket (after successful notarization):
bash
CopyInsert in Terminal
xcrun stapler staple "build/dist/OutlookExtractor.app"
Troubleshooting
Common Issues
Code Signing Fails
Verify the signing identity exists in Keychain
Ensure you have the correct permissions
Notarization Fails
Check the notarization log for specific errors
Ensure all binaries are signed with hardened runtime
Verify entitlements are correctly set
App Won't Launch
Check Console.app for crash logs
Verify all dependencies are included in the bundle
Build Artifacts
build/build/: Intermediate build files
build/dist/: Final output directory
OutlookExtractor: Standalone executable (if --onefile)
OutlookExtractor.app/: Application bundle
OutlookExtractor.zip: Archive for notarization
CopyInsert

### 3. Lint Errors Fixed

The main lint errors have been addressed by:
1. Properly handling environment variables
2. Fixing string formatting in print statements
3. Adding proper error handling
4. Ensuring all code paths have proper returns

### 4. Next Steps

1. **Test the Build**:
   ```bash
   # Test unsigned build
   python build/scripts/build.py --clean
   
   # Test signed build (replace with your actual Developer ID)
   python build/scripts/build.py --clean --sign-identity "Developer ID Application: Your Name (YOUR_TEAM_ID)"
Update Permissions: If you still have permission issues with BUILD_INSTRUCTIONS.md, you may need to:
bash
CopyInsert
sudo chown $(whoami) BUILD_INSTRUCTIONS.md
chmod u+w BUILD_INSTRUCTIONS.md
Verify Build: Check the build/dist directory for the built application and verify it runs correctly.
Would you like me to make any other adjustments to the build process or documentation?
