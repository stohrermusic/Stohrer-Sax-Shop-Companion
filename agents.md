# Sax Shop Companion - Agent Context & Documentation

## Project Overview
**Sax Shop Companion** is a specialized desktop application for saxophone repair technicians, built in Python using `tkinter`. It serves as both a fabrication tool (generating SVG files for laser-cutting pads) and a reference database (key heights, serial numbers, screw specifications).

---

## Codebase Structure

### Core Files
* **`main.py`**: The application entry point. Contains the main class `PadSVGGeneratorApp`, all UI tab construction (`create_pad_generator_tab`, `create_key_library_tab`, etc.), and the core math logic for SVG generation.
* **`serials.py`**: Contains the `SERIAL_DATA` dictionary. This is a massive lookup table for manufacturer serial number ranges and years.

### Data & Configuration (JSON)
* **`app_settings.json`**: Persists user preferences (last directory, UI theme, sizing rules).
    * **Note:** "Star/Dart" settings (e.g., `darts_enabled`, `dart_threshold`) are stored at the **root level** of this JSON object, alongside standard settings.
* **`pad_presets.json`**: Stores measurements for pads.
    * *Structure:* `{"Library Name": {"Instrument Model": "Size\nSize\n..."}}`
* **`key_height_library.json`**: Stores key height measurements.
* **`screw_specs.json`**: Stores thread pitch/rod diameter data.

---

## Key Features & Logic

### 1. Pad SVG Generator (The "Factory")
This is the most complex logic in the app. It generates `.svg` files using `svgwrite`.

* **Star/Dart Pattern:** Generates "geared" or "flower" shapes for small leather pads to facilitate wrapping.
* **Key Functions:**
    * `calculate_star_path(cx, cy, outer_r, inner_r, num_points, shape_factor)`: Generates the path data. The `shape_factor` allows morphing between Sine and Square waves.
    * `leather_back_wrap()`: Calculates diameter expansion based on pad size.
    * `get_disc_diameter()`: Handles logic for Felt vs Card vs Leather (Standard) vs Leather (Dart).
* **Nesting Logic:** `can_all_pads_fit()` uses a greedy algorithm to ensure pads fit on the specified sheet dimensions before generation.

### 2. Databases (Key Heights, Serials, Screws)
* **Structure:** These tabs rely on a "Library" system. Data is loaded into nested dictionaries.
* **Serials:** The lookup logic (`lookup_serial_year`) iterates through ranges in `SERIAL_DATA` to find the start year.
* **Import/Export:** All databases support importing/exporting JSON snippets to share data between users.

---

## Development Workflows

### The "Subtraction Method" (Standalone Generation)
There is a secondary tool called the **"Standalone Pad SVG Generator"** used by colleagues.

* **To update the Standalone Tool:** Do not try to merge code into the old script.
* **Workflow:**
    1.  Take the current `main.py` from Sax Shop Companion.
    2.  Delete the `create_key_library_tab`, `create_serial_lookup_tab`, and `create_screw_specs_tab` methods and their associated classes.
    3.  Keep the `PadSVGGeneratorApp` and `OptionsWindow`.
* **Reasoning:** This ensures full interoperability of settings and presets between the full Companion and the Standalone tool.

---

## UI & Theming
* **Framework:** Standard `tkinter`.
* **Theme:** The app features a "Resonance" Easter egg. Clicking the "Resonance" button repeatedly changes the global background color (`COOL_BLUE`, `COOL_GREEN`) and window alpha transparency.
* **Widgets:** Use `ttk` for Notebooks and Comboboxes, standard `tk` for basic frames/labels to allow easier background color manipulation.

---

## Known Constraints & Gotchas
1.  **Settings Hierarchy:** Unlike early designs, Dart/Star settings are **NOT nested**. When editing `OptionsWindow` or `load_settings()`, access them directly via `self.settings["dart_threshold"]`, `self.settings["dart_shape_factor"]`, etc.
2.  **Pad Strings:** Pad lists are stored as multiline strings (`"Size x Qty"` or just `"Size"`). The parser is robust but expects standard formatting.
3.  **Imports:** Keep imports standard (`json`, `os`, `math`, `tkinter`). Only external dependency should be `svgwrite` (and `pandas` if doing bulk CSV conversions).

## Building Executables
The project uses **PyInstaller** via GitHub Actions to automatically build `.exe` files on commit. Ensure `svgwrite` is included in the requirements/spec file.
