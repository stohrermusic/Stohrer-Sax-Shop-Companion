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

## CI/CD (GitHub Actions)

The `.github/workflows/build.yml` workflow automatically builds for all three platforms:
- Triggers on push to `main`, `beta`, or `gamma`, on release creation, or manually
- Builds Windows .exe, macOS .app (zipped), and Linux binary in parallel
- Uploads artifacts to the workflow run
- Auto-attaches binaries to GitHub Releases

## Config File Locations

The app stores settings and presets in platform-appropriate locations:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\StohrerSaxShopCompanion\` |
| macOS | `~/Library/Application Support/StohrerSaxShopCompanion/` |
| Linux | `~/.config/StohrerSaxShopCompanion/` (respects `XDG_CONFIG_HOME`) |

**Backward compatibility**: On first run, existing config files in the old location (current working directory) are automatically migrated to the new location.

**Manual import**: Users can also manually import settings from a previous installation via File → "Import Settings from Folder..." which copies config files from a selected directory.

## Architecture

### Module Structure

```
main.py                 → Entry point, PadSVGGeneratorApp class, tab creation
    ↓ inherits
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
    ↓ uses
config.py              → Settings I/O, constants, platform config paths, migration logic, import helpers
svg_engine.py          → Pure math/SVG logic (no tkinter dependency), polygon nesting
gcode_engine.py        → G-code generation for Grbl lasers, single-stroke font, circle linearization
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export, PolygonDrawWindow, GcodeSettingsWindow)
serials.py             → SERIAL_DATA dictionary (manufacturer → serial ranges)
build.py               → Cross-platform PyInstaller build script
```

### Key Design Patterns

**Mixin Inheritance**: `PadSVGGeneratorApp` inherits from `LibraryFeaturesMixin` to gain database tab functionality without polluting the main module.

**Pure Logic Separation**: `svg_engine.py` and `gcode_engine.py` contain no tkinter code, making them testable independently. All SVG/G-code generation math (star paths, nesting algorithm, sizing calculations) lives here.

**Tab-Specific Menus**: The app swaps menu bars when tabs change (`on_tab_changed` in main.py).

**Cross-Platform Helpers**: `bind_mousewheel()` in ui_dialogs.py handles platform-specific scroll behavior (Windows/macOS/Linux).

### Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` or `generate_gcode()` → output files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:
- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged

**Nesting Algorithm**: `_nest_discs()` implements greedy circle-packing, shared by both `can_all_pads_fit()` and `generate_svg()`. Supports both rectangular sheets and custom polygon shapes.

**Polygon Shape Tool**: Users can draw custom polygon shapes for irregular leather skins. Grid size adapts to unit setting: 15x15 inches (1" squares) or 40x40 cm (1cm squares). The polygon nesting algorithm (`_nest_discs_polygon()`) uses ray-casting for point-in-polygon checks and distance-to-edge calculations for circle fitting. The rectangle algorithm remains the fast path when no custom shape is defined.

### Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

### Data Files (JSON, in platform config directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications
