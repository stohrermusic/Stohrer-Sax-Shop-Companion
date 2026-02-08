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

The external dependency is `svgwrite`. The GUI uses Python's built-in `tkinter`. Requires Python 3.11+.

There is no automated test suite. Changes are verified by running the application manually.

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

**macOS Theming**: On macOS (`IS_MACOS` flag in main.py and ui_dialogs.py), the app uses native system colors instead of custom cream/beige backgrounds. This allows the app to work correctly in both macOS dark and light mode. The `DIALOG_BG` constant in ui_dialogs.py resolves to `"systemWindowBackgroundColor"` on macOS (a Tk system color that adapts to dark/light mode) and `"#F0EAD6"` on Windows/Linux. The resonance theme system is disabled on macOS. When adding new UI widgets, use `DIALOG_BG` for dialog backgrounds and `self.root.cget('bg')` for main window widgets.

### Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` or `generate_gcode()` → output files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:
- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged

**Nesting Algorithm**: `_nest_discs()` implements greedy circle-packing, shared by both `can_all_pads_fit()` and `generate_svg()`. Supports both rectangular sheets and custom polygon shapes.

**Shared Rendering Helpers**: SVG rendering is centralized in `_render_svg_discs()` and `_create_svg_drawing()` in svg_engine.py. Both `generate_svg()` and `generate_svg_from_placed()` call these. Similarly, `generate_gcode()` delegates to `generate_gcode_from_placed()` after nesting, so all G-code disc rendering logic lives in one place.

**Coordinate Systems**: SVG uses Y=0 at top (Y increases downward), while G-code uses Y=0 at bottom (Y increases upward). The `gcode_engine.py` functions flip Y coordinates (`sheet_height_mm - cy`) to ensure G-code output matches the SVG preview.

**Polygon Shape Tool**: Users can draw custom polygon shapes for irregular leather skins. Grid size adapts to unit setting: 15x15 inches (1" squares) or 40x40 cm (1cm squares). The polygon nesting algorithm (`_nest_discs_polygon()`) uses ray-casting for point-in-polygon checks and distance-to-edge calculations for circle fitting. The rectangle algorithm remains the fast path when no custom shape is defined.

**Scrap Mode**: Allows users to place pads across multiple irregular scrap pieces instead of requiring one large sheet. Key components:
- Session state in `self.scrap_session` tracks original pads, remaining pads, scrap count, locked material
- `try_nest_partial()` and `compute_remaining_pads()` in svg_engine.py handle partial placement
- `generate_svg_from_placed()` / `generate_gcode_from_placed()` generate output from pre-computed placements
- Progress popup window (`scrap_remaining_window`) shows two columns: Remaining and Done
- Materials are locked (disabled) during active session to prevent switching
- Files named with `_scrap1`, `_scrap2` suffixes

**Send to SD Card**: File → Send G-code to SD Card... provides a streamlined workflow for laser cutters that use SD cards (like the Creality Falcon2 Pro). The feature:
- Opens a folder picker to select the SD card (remembers last used location)
- Shows existing G-code files and confirms before erasing
- Opens a file picker to select the .gcode file to send
- Copies the file and safely ejects the SD card (Windows only)
- User can then physically remove the card and use the laser's built-in controls

### Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

### Data Files (JSON, in platform config directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications

## Web Data Sync ("Import Matt's")

The app can fetch reference data from https://www.stohrermusic.com:

- **Screw Specs**: File > "Import Matt's Specs" fetches `/data/screw_specs.json` and imports each model with a "(Matt's)" suffix to avoid overwriting local entries.
- **Key Heights**: File > "Import Matt's Key Heights" fetches `/data/key_height_library.json` and imports all presets into a dedicated "Matt's Library".

Both use `urllib.request` (stdlib) with a 10-second timeout. The JSON formats on the website match the app's internal format exactly, so no conversion is needed.

When adding new data to the website, update the JSON files in `C:\code\stohrermusic\static\data\` and push to deploy. App users can then re-import to get the latest.

## Screw Specs Submission Form

A web form at https://www.stohrermusic.com/articles/screw-specs-library/ (Submit Specs tab) allows colleagues to submit screw thread specifications. Submissions POST to a Google Apps Script endpoint that appends rows to a Google Sheet for review.

### Integrating Submitted Data

After reviewing submissions in the Google Sheet:

1. **Update the app's screw_specs.json**: Edit `C:\code\saxshopcompanion\static\data\screw_specs.json` (or the user's local config file) to add verified entries. Format:
   ```json
   "Manufacturer": {
     "Model": {
       "neck_screw_th": "M4x0.7",
       "neck_screw_dia": "",
       "hinge_tiny_th": "",
       "hinge_tiny_dia": "",
       "hinge_small_th": "2-56 NC",
       "hinge_small_dia": "side keys",
       ...
       "notes": "source info here"
     }
   }
   ```

2. **Update the website library**: Copy the same data to `C:\code\stohrermusic\static\data\screw_specs.json`, then commit and push to deploy.

3. **Field mapping** (Google Sheet columns → JSON keys):
   - Neck Screw Thread/Desc → `neck_screw_th`, `neck_screw_dia`
   - Hinge Rod Tiny/Small/Medium/Large → `hinge_tiny_*`, `hinge_small_*`, `hinge_med_*`, `hinge_lrg_*`
   - Pivot Small/Large → `pivot_small_*`, `pivot_lrg_*`
   - Misc 1/2 → `misc1`, `misc2`
   - Notes → `notes`

Note: The `_dia` fields in the JSON are used for descriptions despite the legacy naming.

## Related Repository

The website at https://www.stohrermusic.com is a Hugo site with the Blowfish theme, located at `C:\code\stohrermusic`. The screw specs library page there loads data from `/static/data/screw_specs.json` and must be kept in sync with this app's data.
