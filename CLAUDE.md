# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Stohrer Sax Shop Companion is a desktop GUI application for saxophone repair technicians. It provides SVG generation for laser-cutting pad materials and reference databases for key heights, serial numbers, and screw specifications.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

The only external dependency is `svgwrite`. The GUI uses Python's built-in `tkinter`.

## Architecture

### Module Structure

```
main.py                 → Entry point, PadSVGGeneratorApp class, tab creation
    ↓ inherits
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
    ↓ uses
config.py              → Settings I/O, constants, DEFAULT_SETTINGS dict
svg_engine.py          → Pure math/SVG logic (no tkinter dependency)
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export)
serials.py             → SERIAL_DATA dictionary (manufacturer → serial ranges)
```

### Key Design Patterns

**Mixin Inheritance**: `PadSVGGeneratorApp` inherits from `LibraryFeaturesMixin` to gain database tab functionality without polluting the main module.

**Pure Logic Separation**: `svg_engine.py` contains no tkinter code, making it testable independently. All SVG generation math (star paths, nesting algorithm, sizing calculations) lives here.

**Tab-Specific Menus**: The app swaps menu bars when tabs change (`on_tab_changed` in main.py).

### Data Flow

**Pad Generation**: User input → `parse_pad_list()` → `can_all_pads_fit()` check → `generate_svg()` → SVG files

**Sizing Calculations**: `get_disc_diameter()` in svg_engine.py applies material-specific offsets:
- Felt: pad_size - felt_offset
- Card: pad_size - (felt_offset + card_to_felt_offset)
- Leather: pad_size + 2*(felt_thickness + wrap) with star bonus for small pads
- Exact: pad_size unchanged

### Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

### Data Files (JSON, stored in working directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications

## Known Issues

1. **KeyError in key preset saving** (`library_features.py:260`): Doesn't create library dict before checking if preset name exists (unlike pad preset saving)

2. **SVG unit inconsistency** (`svg_engine.py:217`): Star/dart paths drawn without units; circles use "mm" units in non-compatibility mode

3. **Serial lookup assumes sorted data** (`library_features.py:50-54`): Algorithm breaks early assuming ascending order

4. **Platform mousewheel** (`ui_dialogs.py`): `event.delta/120` is Windows-specific; macOS uses different delta values

5. **Duplicated nesting algorithm**: `can_all_pads_fit()` and `generate_svg()` have nearly identical circle-packing logic
