#!/usr/bin/env python3
"""
Build script for Outlook Extractor.
Creates standalone executables using PyInstaller.
"""
import datetime
import os
import shutil
import subprocess
import sys
from pathlib import Path

# Project root directory (where this script is located)
PROJECT_ROOT = Path(__file__).parent.parent.parent

# Build directories
BUILD_DIR = PROJECT_ROOT / "build" / "build"  # Intermediate build files
DIST_DIR = PROJECT_ROOT / "build" / "dist"     # Final output directory

# Ensure build directories exist
BUILD_DIR.mkdir(parents=True, exist_ok=True)
DIST_DIR.mkdir(parents=True, exist_ok=True)

# Spec file
SPEC_FILE = PROJECT_ROOT / "build" / "outlook_extractor.spec"

def clean_build():
    """Remove previous build artifacts."""
    print("🚮 Cleaning previous builds...")
    
    # Clean up build and dist directories
    for path in [BUILD_DIR, DIST_DIR, PROJECT_ROOT / "build" / "outlook_extractor.spec"]:
        if path.exists():
            print(f"  Removing {path}")
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            else:
                path.unlink(missing_ok=True)
    
    # Clean up any .spec files in the project root
    for spec_file in PROJECT_ROOT.glob("*.spec"):
        print(f"  Removing {spec_file}")
        spec_file.unlink(missing_ok=True)
    
    # Clean up any __pycache__ directories
    for pycache in PROJECT_ROOT.rglob("__pycache__"):
        print(f"  Removing {pycache}")
        shutil.rmtree(pycache, ignore_errors=True)

def create_spec_file():
    """Create or update the PyInstaller spec file."""
    print(f"[UPDATE] Updating spec file: {SPEC_FILE}")
    
    # Get the absolute path to the project's Python files
    outlook_extractor_path = str(PROJECT_ROOT / 'outlook_extractor')
    
    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Add the project directory to the Python path
import sys
sys.setrecursionlimit(5000)  # Increase recursion limit for PyInstaller

# The Analysis object represents the main configuration for PyInstaller
a = Analysis(
    ['{outlook_extractor_path}/__main__.py'],  # Entry point
    pathex=['{PROJECT_ROOT}'],  # Project root
    binaries=[],  # List of (src, dst) tuples for binary files
    datas=[
        # Include any non-Python files needed by the application
        # Format: (source_path, destination_path_in_bundle)
        ('{outlook_extractor_path}/ui/assets', 'outlook_extractor/ui/assets'),
    ],
    hiddenimports=[
        # List of modules that PyInstaller can't detect automatically
        'email_validator',
        'pydantic',
        'pydantic_core',
        'PySimpleGUI',
        'sqlalchemy',
        'sqlalchemy.sql.default_comparator',
        'html2text',
        'bs4',  # beautifulsoup4
        'bs4.builder._lxml',
        'tqdm',
        'tqdm.utils',
        'dateutil',
        'dateutil.parser',
        'dateutil.tz',
        'dateutil.relativedelta',
        'pytz',
        'email.mime',
        'email.mime.text',
        'email.mime.multipart',
        'email.mime.base',
        'email.encoders',
        'email.utils',
        'email.charset',
        'email.header',
        'email.message',
        'semver',
        'email.policy',
        'email._parseaddr',
        'email._policybase',
        'email.iterators',
        'email.generator',
        'html',
        'http',
        'http.client',
        'http.cookies',
        'http.cookiejar',
        'urllib3',
        'urllib3.util',
        'urllib3.exceptions',
        'urllib3.packages',
        'urllib3.contrib',
        'urllib3.contrib._appengine_environ',
        'urllib3.contrib.securetransport',
        'urllib3.contrib.socks',
        'urllib3.packages.backports',
        'urllib3.packages.six',
        'urllib3.packages.ssl_match_hostname',
        'urllib3.packages.ssltransport',
        'urllib3.util.connection',
        'urllib3.util.queue',
        'urllib3.util.request',
        'urllib3.util.response',
        'urllib3.util.retry',
        'urllib3.util.ssl_',
        'urllib3.util.timeout',
        'urllib3.util.url',
        'urllib3.util.wait',
        'urllib3.util.ssl_match_hostname',
        'urllib3.util.ssltransport',
        'urllib3.util.ssl_',
        'urllib3.util.connection',
        'urllib3.util.retry',
        'urllib3.util.ssl_',
        'urllib3.util.timeout',
        'urllib3.util.url',
        'urllib3.util.wait',
        'urllib3.contrib.emulator',
        'urllib3.contrib.emulator.handler',
        'urllib3.contrib.emulator.request',
        'urllib3.contrib.emulator.response',
        'urllib3.contrib.emulator.util',
        'urllib3.contrib.emulator.wsgi',
        'urllib3.contrib.emulator.wsgi_app',
        'urllib3.contrib.emulator.wsgi_handler',
        'urllib3.contrib.emulator.wsgi_server',
        'urllib3.contrib.emulator.wsgi_utils',
        'urllib3.contrib.emulator.wsgi_worker',
        'urllib3.contrib.emulator.wsgi_worker_pool',
        'urllib3.contrib.emulator.wsgi_worker_thread',
        'urllib3.contrib.emulator.wsgi_worker_thread_pool',
        'urllib3.contrib.emulator.wsgi_worker_thread_pool_worker',
        'urllib3.contrib.emulator.wsgi_worker_thread_worker',
        'urllib3.contrib.emulator.wsgi_worker_worker',
        'urllib3.contrib.emulator.wsgi_worker_worker_pool',
        'urllib3.contrib.emulator.wsgi_worker_worker_pool_worker',
        'urllib3.contrib.emulator.wsgi_worker_worker_worker',
    ],
    hookspath=[],  # List of paths containing hook files
    hooksconfig={{}},  # Hooks configuration
    runtime_hooks=[],  # List of custom runtime hook files
    excludes=[],  # List of module names to exclude
    win_no_prefer_redirects=False,  # Windows-specific
    win_private_assemblies=False,  # Windows-specific
    cipher=block_cipher,  # For encrypting Python bytecode
    noarchive=False,  # If True, don't create an archive
)

# Create the PYZ (Python Zip Archive)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Create the executable
exe = EXE(
    pyz,  # The PYZ archive
    a.scripts,  # Scripts to convert to executables
    a.binaries,  # Non-Python modules needed by the scripts
    a.zipfiles,  # Zip files to include
    a.datas,  # Non-binary files needed by the scripts
    name='OutlookExtractor',  # Name of the output executable
    debug=False,  # Create a debug version
    bootloader_ignore_signals=False,  # Whether to ignore signals in the bootloader
    strip=False,  # Strip debug symbols from the executable
    upx=True,  # Use UPX to compress the executable
    upx_exclude=[],  # Files to exclude from UPX compression
    runtime_tmpdir=None,  # Temporary directory for the application
    console=False,  # Whether to show a console window (False for GUI apps)
    disable_windowed_traceback=False,  # Disable traceback in windowed mode
    target_arch=None,  # Target architecture (None for auto-detect)
    codesign_identity=None,  # Code signing identity (macOS)
    entitlements_file=None,  # Entitlements file (macOS)
    icon=None,  # Path to the application icon
)
"""
    
    # Ensure the directory exists
    SPEC_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    # Write the spec file
    with open(SPEC_FILE, 'w') as f:
        f.write(spec_content)

def build_executable(onefile=True, sign_identity=None):
    """Build the executable using PyInstaller.
    
    Args:
        onefile (bool): Whether to create a single executable file
        sign_identity (str, optional): Apple Developer ID for code signing. 
            If None, the build will be unsigned.
    """
    print("[BUILD] Building executable...")
    
    # Store original environment variables
    original_env = {}
    if sign_identity:
        print(f"[SIGN] Code signing with identity: {sign_identity}")
        # Save and set environment variables for code signing
        original_env = {
            'CODESIGN_ALLOCATE': os.environ.get('CODESIGN_ALLOCATE'),
            'CODESIGN_IDENTITY': os.environ.get('CODESIGN_IDENTITY')
        }
        os.environ['CODESIGN_ALLOCATE'] = '/usr/bin/codesign_allocate'
        os.environ['CODESIGN_IDENTITY'] = sign_identity
    else:
        print("[UNSIGNED] Building unsigned application")
    
    # Ensure PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("[INSTALL] Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Create the spec file with the correct options
    create_spec_file()
    
    # Get the path to the main script
    main_script = PROJECT_ROOT / "outlook_extractor" / "__main__.py"
    if not main_script.exists():
        print(f"[ERROR] Could not find main script at {main_script}")
        return False
    
    # Get the path to the assets directory
    assets_dir = PROJECT_ROOT / "outlook_extractor" / "ui" / "assets"
    
    # Change to the project root directory
    original_dir = os.getcwd()
    os.chdir(PROJECT_ROOT)
    
    try:
        # Build command
        cmd = [
            sys.executable,
            "-m", "PyInstaller",
            "--clean",
            "--noconfirm",
            "--workpath", str(BUILD_DIR.relative_to(PROJECT_ROOT)),
            "--distpath", str(DIST_DIR.relative_to(PROJECT_ROOT)),
            "--specpath", str(SPEC_FILE.parent.relative_to(PROJECT_ROOT)),
            "--windowed",  # For GUI apps
            "--onefile" if onefile else "--onedir",
            "--name", "OutlookExtractor",
        ]
        
        # Add assets if they exist
        if assets_dir.exists():
            rel_assets = assets_dir.relative_to(PROJECT_ROOT)
            cmd.extend(["--add-data", f"{rel_assets}{os.pathsep}outlook_extractor/ui/assets"])
        
        # Add hidden imports
        hidden_imports = [
            "PySimpleGUI", "pydantic", "sqlalchemy", "html2text", "bs4", "tqdm",
            "python_dateutil", "pytz", "email_validator", "sqlalchemy.orm",
            "sqlalchemy.ext.declarative", "sqlalchemy.ext", "outlook_extractor"
        ]
        
        for imp in hidden_imports:
            cmd.extend(["--hidden-import", imp])
        
        # Add the main script
        cmd.append(str(main_script.relative_to(PROJECT_ROOT)))
        
        # Run the build
        print("[RUN] Running PyInstaller...")
        print("   " + " ".join(f'"{arg}"' if ' ' in str(arg) else str(arg) for arg in cmd))
        
        subprocess.check_call(cmd, cwd=PROJECT_ROOT)
        
        print("[SUCCESS] Build completed successfully!")
        if onefile:
            exe_path = DIST_DIR / "OutlookExtractor"
        else:
            exe_path = DIST_DIR / "OutlookExtractor" / "OutlookExtractor"
        print(f"[LOCATION] Executable location: {exe_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Build failed with error code {e.returncode}")
        return False
    finally:
        # Change back to the original directory
        os.chdir(original_dir)
    
    try:
        # Run the build
        print("[RUN] Running PyInstaller...")
        print("   " + " ".join(f'"{arg}"' if ' ' in str(arg) else str(arg) for arg in cmd))
        
        subprocess.check_call(cmd)
        
        print("[SUCCESS] Build completed successfully!")
        if onefile:
            exe_path = DIST_DIR / "OutlookExtractor"
        else:
            exe_path = DIST_DIR / "OutlookExtractor" / "OutlookExtractor"
        print(f"[LOCATION] Executable location: {exe_path}")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Build failed with error code {e.returncode}")
        return False
    
    # Run the build
    print("[RUN] Running PyInstaller...")
    print("   " + " ".join(f'"{arg}"' if ' ' in str(arg) else str(arg) for arg in cmd))
    
    try:
        subprocess.check_call(cmd)
        print("[SUCCESS] Build completed successfully!")
        print(f"[LOCATION] Executable location: {DIST_DIR / 'OutlookExtractor'}")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Build failed with error code {e.returncode}")
        sys.exit(1)

def main():
    """Main function to handle the build process."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Build Outlook Extractor')
    parser.add_argument('--onefile', action='store_true', default=True,
                      help='Create a single executable file (default: True)')
    parser.add_argument('--clean', action='store_true',
                      help='Clean build artifacts before building')
    parser.add_argument('--spec-only', action='store_true',
                      help='Only create the spec file, do not build')
    parser.add_argument('--sign-identity', type=str, default=None,
                      help='Apple Developer ID Application identity for code signing (e.g., "Developer ID Application: Your Name (XXXXXXXXXX)")')
    parser.add_argument('--bundle-id', type=str, default='com.outlook.extractor',
                      help='Bundle identifier for the application (default: com.outlook.extractor)')
    
    args = parser.parse_args()
    
    try:
        # Ensure build directories exist
        BUILD_DIR.mkdir(parents=True, exist_ok=True)
        DIST_DIR.mkdir(parents=True, exist_ok=True)
        
        # Create build info file
        with open(BUILD_DIR / 'build_info.txt', 'w') as f:
            f.write(f"Build time: {datetime.datetime.now().isoformat()}\n")
            f.write(f"Python: {sys.version}\n")
            f.write(f"Platform: {sys.platform}\n")
            f.write(f"Signed: {'Yes' if args.sign_identity else 'No'}\n")
            if args.sign_identity:
                f.write(f"Signing Identity: {args.sign_identity}\n")
        
        # Set the spec file path
        global SPEC_FILE
        SPEC_FILE = BUILD_DIR / "outlook_extractor.spec"
        
        # Clean if requested
        if args.clean:
            clean_build()
        
        # Create the spec file
        create_spec_file()
        
        # Build if not spec-only
        if not args.spec_only:
            # Add bundle identifier to Info.plist if specified
            if args.bundle_id:
                info_plist_path = PROJECT_ROOT / 'build' / 'Info.plist'
                info_plist = f'''<?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
                <plist version="1.0">
                <dict>
                    <key>CFBundleIdentifier</key>
                    <string>{args.bundle_id}</string>
                    <key>CFBundleName</key>
                    <string>Outlook Extractor</string>
                    <key>CFBundleDisplayName</key>
                    <string>Outlook Extractor</string>
                    <key>CFBundleVersion</key>
                    <string>1.0.0</string>
                    <key>CFBundleShortVersionString</key>
                    <string>1.0.0</string>
                    <key>NSHighResolutionCapable</key>
                    <true/>
                    <key>NSAppTransportSecurity</key>
                    <dict>
                        <key>NSAllowsArbitraryLoads</key>
                        <true/>
                    </dict>
                </dict>
                </plist>
                '''
                with open(info_plist_path, 'w') as f:
                    f.write(info_plist)
                
                # Add the plist file to the build command
                os.environ['PYI_OSX_BUNDLE_ID'] = args.bundle_id
                os.environ['PYI_OSX_BUNDLE_INFO_PLIST'] = str(info_plist_path)
            
            if build_executable(onefile=args.onefile, sign_identity=args.sign_identity):
                print("\n[SUCCESS] Build process completed!")
                print(f"   [LOCATION] Executable available at: {DIST_DIR / 'OutlookExtractor'}")
                
                # Sign the app bundle if identity is provided
                if args.sign_identity and (DIST_DIR / 'OutlookExtractor.app').exists():
                    print("[SIGN] Signing application bundle...")
                    try:
                        # Sign the app bundle
                        subprocess.check_call([
                            'codesign', '--deep', '--force', '--verify', '--verbose',
                            '--sign', args.sign_identity,
                            '--options', 'runtime',
                            '--entitlements', str(PROJECT_ROOT / 'build' / 'entitlements.plist'),
                            str(DIST_DIR / 'OutlookExtractor.app')
                        ])
                        print("[SUCCESS] Application bundle signed successfully!")
                        
                        # Verify the signature
                        print("[VERIFY] Verifying code signature...")
                        subprocess.check_call([
                            'codesign', '--verify', '--verbose',
                            str(DIST_DIR / 'OutlookExtractor.app')
                        ])
                        
                        # Create a ZIP file for notarization
                        print("[ZIP] Creating archive for notarization...")
                        zip_path = DIST_DIR / 'OutlookExtractor.zip'
                        subprocess.check_call([
                            'ditto', '-c', '-k', '--keepParent',
                            str(DIST_DIR / 'OutlookExtractor.app'),
                            str(zip_path)
                        ])
                        print(f"[SUCCESS] Archive created: {zip_path}")
                        print("\n[NOTARIZATION] Next steps for notarization:")
                        print("1. Upload for notarization:")
                        print(f"   xcrun altool --notarize-app --primary-bundle-id \"{args.bundle_id}\" \\")
                        print('   --username "YOUR_APPLE_ID_EMAIL" --password "@keychain:AC_PASSWORD" \\')
                        print(f'   --file "{zip_path}"')
                        print("2. Check status:")
                        print("   xcrun altool --notarization-info <request-id> -u <email>")
                        print("3. Staple the ticket:")
                        print(f"   xcrun stapler staple \"{DIST_DIR / 'OutlookExtractor.app'}\"")
                        
                    except subprocess.CalledProcessError as e:
                        print(f"[ERROR] Code signing failed: {e}")
                        # Restore original environment
                        for k, v in original_env.items():
                            if v is not None:
                                os.environ[k] = v
                            else:
                                os.environ.pop(k, None)
                        return False
                    finally:
                        # Always restore original environment
                        for k, v in original_env.items():
                            if v is not None:
                                os.environ[k] = v
                            else:
                                os.environ.pop(k, None)
                        
            else:
                print("\n[ERROR] Build failed!")
                sys.exit(1)
        else:
            print("\n[SUCCESS] Spec file created successfully!")
        
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
