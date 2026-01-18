# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Stohrer Sax Shop Companion is a cross-platform desktop GUI application for saxophone repair technicians. It provides SVG generation for laser-cutting pad materials and reference databases for key heights, serial numbers, and screw specifications.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

The external dependency is `svgwrite`. The GUI uses Python's built-in `tkinter`.

## Building Executables

The app uses PyInstaller to create standalone executables. Each platform must build its own executable (no cross-compilation).

```bash
# Install build dependencies
pip install -r requirements.txt

# Build for current platform
python build.py

# Clean and rebuild
python build.py --clean

# (macOS only) Build and create .dmg disk image
python build.py --dmg

# Output locations:
#   Windows: dist/StohrerSaxShopCompanion.exe
#   macOS:   dist/StohrerSaxShopCompanion.app (or .dmg with --dmg flag)
#   Linux:   dist/StohrerSaxShopCompanion
```

## Config File Locations

The app stores settings and presets in platform-appropriate locations:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\StohrerSaxShopCompanion\` |
| macOS | `~/Library/Application Support/StohrerSaxShopCompanion/` |
| Linux | `~/.config/StohrerSaxShopCompanion/` (respects `XDG_CONFIG_HOME`) |

**Backward compatibility**: On first run, existing config files in the old location (current working directory) are automatically migrated to the new location.

## Architecture

### Module Structure

```
main.py                 → Entry point, PadSVGGeneratorApp class, tab creation
    ↓ inherits
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
    ↓ uses
config.py              → Settings I/O, constants, platform config paths, migration logic
svg_engine.py          → Pure math/SVG logic (no tkinter dependency)
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export)
serials.py             → SERIAL_DATA dictionary (manufacturer → serial ranges)
build.py               → Cross-platform PyInstaller build script
```

### Key Design Patterns

**Mixin Inheritance**: `PadSVGGeneratorApp` inherits from `LibraryFeaturesMixin` to gain database tab functionality without polluting the main module.

**Pure Logic Separation**: `svg_engine.py` contains no tkinter code, making it testable independently. All SVG generation math (star paths, nesting algorithm, sizing calculations) lives here.

**Tab-Specific Menus**: The app swaps menu bars when tabs change (`on_tab_changed` in main.py).

**Cross-Platform Helpers**: `bind_mousewheel()` in ui_dialogs.py handles platform-specific scroll behavior (Windows/macOS/Linux).

### Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` → SVG files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:
- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged

**Nesting Algorithm**: `_nest_discs()` implements greedy circle-packing, shared by both `can_all_pads_fit()` and `generate_svg()`.

### Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

### Data Files (JSON, in platform config directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications
