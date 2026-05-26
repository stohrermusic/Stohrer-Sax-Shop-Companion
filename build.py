#!/usr/bin/env python3
"""
Build script for Stohrer Sax Shop Companion.
Creates standalone executables for Windows, macOS, and Linux using PyInstaller.

Usage:
    python build.py              # Build for current platform
    python build.py --clean      # Clean build artifacts first
    python build.py --dmg        # (macOS only) Also create a .dmg disk image
"""

import subprocess
import sys
import os
import shutil
import argparse

APP_NAME = "SaxShopCompanion"
MAIN_SCRIPT = "main.py"
DMG_NAME = f"{APP_NAME}.dmg"


def get_platform_name():
    """Return a human-readable platform name."""
    if sys.platform == 'win32':
        return 'Windows'
    elif sys.platform == 'darwin':
        return 'macOS'
    else:
        return 'Linux'


def clean_build_artifacts():
    """Remove build and dist directories."""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"Removing {dir_name}/...")
            shutil.rmtree(dir_name)

    # Remove .spec file if it exists
    spec_file = f"{APP_NAME}.spec"
    if os.path.exists(spec_file):
        print(f"Removing {spec_file}...")
        os.remove(spec_file)

    # Remove .dmg if it exists
    if os.path.exists(DMG_NAME):
        print(f"Removing {DMG_NAME}...")
        os.remove(DMG_NAME)


def create_dmg():
    """Create a .dmg disk image for macOS distribution."""
    if sys.platform != 'darwin':
        print("Error: .dmg creation is only available on macOS")
        return False

    app_path = f"dist/{APP_NAME}.app"
    if not os.path.exists(app_path):
        print(f"Error: {app_path} not found. Run build first.")
        return False

    print(f"Creating {DMG_NAME}...")

    # Remove existing .dmg if present
    if os.path.exists(DMG_NAME):
        os.remove(DMG_NAME)

    cmd = [
        'hdiutil', 'create',
        '-volname', APP_NAME,
        '-srcfolder', app_path,
        '-ov',
        '-format', 'UDZO',
        DMG_NAME
    ]

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)

    if result.returncode == 0:
        print(f"\n.dmg created successfully: {DMG_NAME}")
        return True
    else:
        print(f"\n.dmg creation failed with exit code {result.returncode}")
        return False


def _patch_macos_plist():
    """Add microphone permission to the macOS app bundle's Info.plist.

    macOS silently denies microphone access to apps that don't declare
    NSMicrophoneUsageDescription in their Info.plist. Without this,
    the tuner and toner tabs can't access the mic.
    """
    import plistlib

    plist_path = f"dist/{APP_NAME}.app/Contents/Info.plist"
    if not os.path.exists(plist_path):
        print(f"Warning: {plist_path} not found, skipping plist patch")
        return

    print("Patching Info.plist with microphone permission...")

    with open(plist_path, 'rb') as f:
        plist = plistlib.load(f)

    plist['NSMicrophoneUsageDescription'] = (
        'The tuner and tone analyzer need microphone access '
        'to detect pitch and analyze harmonics.'
    )
    plist['NSCameraUsageDescription'] = (
        'Camera access is needed to capture scrap outlines from the '
        'laser bed and to calibrate the laser-bed camera.'
    )

    with open(plist_path, 'wb') as f:
        plistlib.dump(plist, f)

    print("  Added NSMicrophoneUsageDescription + NSCameraUsageDescription to Info.plist")


def build():
    """Build the application for the current platform."""
    platform = get_platform_name()
    print(f"Building {APP_NAME} for {platform}...")

    # Base PyInstaller command
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name', APP_NAME,
        '--windowed',         # No console window (GUI app)
        '--noconfirm',        # Overwrite without asking
    ]

    # Platform-specific options
    if sys.platform == 'darwin':
        # macOS: onedir mode (.app bundle) — onefile+windowed is deprecated
        cmd.extend([
            '--onedir',
            '--osx-bundle-identifier', 'com.stohrer.saxshopcompanion',
        ])
    else:
        # Windows/Linux: single executable
        cmd.append('--onefile')

    # App icon (platform-appropriate format)
    if sys.platform == 'win32':
        icon_path = 'icon.ico'
    elif sys.platform == 'darwin':
        icon_path = 'icon.icns'
    else:
        icon_path = 'icon.ico'  # Linux uses .ico via PyInstaller
    if os.path.exists(icon_path):
        cmd.extend(['--icon', icon_path])
        print(f"  Using icon: {icon_path}")
    # Also bundle icon.ico as data so tkinter can use it at runtime
    if os.path.exists('icon.ico'):
        cmd.extend(['--add-data', f'icon.ico{os.pathsep}.'])

    # Bundle tooling asset SVGs (die organizer templates) so they're available at runtime
    if os.path.isdir('tooling_assets'):
        cmd.extend(['--add-data', f'tooling_assets{os.pathsep}tooling_assets'])

    # Bundle pad-press-spacer STL files so they can be saved to disk at runtime.
    if os.path.isdir('pad_press_spacers'):
        cmd.extend(['--add-data', f'pad_press_spacers{os.pathsep}pad_press_spacers'])

    # Bundle compiled translation catalogs (.mo files) so i18n works in frozen builds.
    # i18n._locale_dir() resolves to sys._MEIPASS/locale when frozen.
    if os.path.isdir('locale'):
        cmd.extend(['--add-data', f'locale{os.pathsep}locale'])

    # Hidden imports for optional audio dependencies (tuner/toner)
    # Only include if numpy/sounddevice are installed (not present in legacy Mac build)
    try:
        import numpy  # noqa: F401  (capability check)
        import sounddevice  # noqa: F401  (capability check)
        cmd.extend([
            '--hidden-import', 'numpy',
            '--hidden-import', 'sounddevice',
            '--hidden-import', '_sounddevice_data',
            '--collect-data', 'sounddevice',
        ])
        print("  Audio libraries found — including tuner/toner support")
    except ImportError:
        print("  Audio libraries not found — building without tuner/toner")

    # Hidden import for PIL (used by strobe tuner wheel rendering)
    try:
        import PIL  # noqa: F401  (capability check)
        cmd.extend([
            '--hidden-import', 'PIL',
            '--hidden-import', 'PIL.Image',
            '--hidden-import', 'PIL.ImageDraw',
            '--hidden-import', 'PIL.ImageTk',
        ])
        print("  Pillow found — including PIL image support")
    except ImportError:
        print("  Pillow not found — tuner will fall back to polygon rendering")

    # GPU-accelerated tuner renderer (Rust/wgpu via pyo3)
    try:
        import tuner_render  # noqa: F401  (capability check)
        cmd.extend([
            '--hidden-import', 'tuner_render',
        ])
        print("  tuner_render found — including GPU strobe renderer")
    except ImportError:
        print("  tuner_render not found — tuner will use CPU canvas rendering")

    # OpenCV — powers the camera-capture feature (Calibration Card,
    # Camera Calibration wizard, Get-from-camera button).
    try:
        import cv2  # noqa: F401
        cmd.extend([
            '--collect-all', 'cv2',
        ])
        print("  OpenCV found — including camera-capture support")
    except ImportError:
        print("  OpenCV not found — building without camera-capture")

    # pyserial — powers Frame & Cut (direct serial to Grbl-compatible lasers).
    try:
        import serial  # noqa: F401
        cmd.extend([
            '--hidden-import', 'serial',
            '--hidden-import', 'serial.tools.list_ports',
        ])
        print("  pyserial found — including Frame & Cut support")
    except ImportError:
        print("  pyserial not found — building without Frame & Cut")

    # Add the main script
    cmd.append(MAIN_SCRIPT)

    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)

    if result.returncode == 0:
        print("\nBuild successful!")
        print("Output location: dist/")

        if sys.platform == 'win32':
            print(f"  Executable: dist/{APP_NAME}.exe")
        elif sys.platform == 'darwin':
            print(f"  App Bundle: dist/{APP_NAME}.app")
            _patch_macos_plist()
        else:
            print(f"  Executable: dist/{APP_NAME}")
    else:
        print(f"\nBuild failed with exit code {result.returncode}")
        sys.exit(result.returncode)


def main():
    parser = argparse.ArgumentParser(description=f'Build {APP_NAME}')
    parser.add_argument('--clean', action='store_true', help='Clean build artifacts before building')
    parser.add_argument('--clean-only', action='store_true', help='Only clean, do not build')
    parser.add_argument('--dmg', action='store_true', help='(macOS only) Create .dmg disk image after building')
    args = parser.parse_args()

    if args.clean or args.clean_only:
        clean_build_artifacts()

    if not args.clean_only:
        build()

        # Create .dmg if requested (macOS only)
        if args.dmg:
            create_dmg()


if __name__ == '__main__':
    main()
