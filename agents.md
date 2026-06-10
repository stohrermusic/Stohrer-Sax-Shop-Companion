# Sax Shop Companion - Agent Context & Documentation

## Project Overview
**Sax Shop Companion** is a specialized desktop application for saxophone repair technicians, built in Python using `tkinter`. It serves as both a fabrication tool (generating SVG files for laser-cutting pads) and a reference database (key heights, serial numbers, screw specifications).

---

## Codebase Structure (Refactored v2.0)

### Python Source Files
* **`main.py`**: The application entry point (Controller). It initializes the `PadSVGGeneratorApp`, builds the UI, and inherits extra features from `library_features.py`.
* **`config.py`**: The "spine" of the app. Contains constants (`DEFAULT_SETTINGS`, `LIGHTBURN_COLORS`), file paths, and IO functions (`load_settings`, `save_presets`).
* **`svg_engine.py`**: Pure logic and math. Contains `calculate_star_path`, nesting algorithms, and the `svgwrite` generation code. No UI dependencies.
* **`ui_dialogs.py`**: Contains all popup windows (`OptionsWindow`, `ImportPresetsWindow`, `ResonanceWindow`, etc.) to keep the main controller clean.
* **`library_features.py`**: Contains the `LibraryFeaturesMixin` class. This holds the logic for Key Heights, Serial Numbers, and Screw Specs.
* **`serials.py`**: Contains the `SERIAL_DATA` dictionary. A massive lookup table for manufacturer serial number ranges.

### Data & Configuration (JSON)
* **`app_settings.json`**: Persists user preferences.
    * **Note:** Dart settings (e.g., `darts_enabled`, `dart_*` keys) are stored at the **root level**, not nested. Internal variable names retain the `dart_` prefix; the user-facing label is just "Darts".
* **`pad_presets.json`**: Stores measurements for pads.
* **`key_height_library.json`**: Stores key height measurements.
* **`screw_specs.json`**: Stores thread pitch/rod diameter data.

---

## Key Features & Logic

### 1. Pad Maker (The "Factory")
* **Location:** Logic in `svg_engine.py`; UI in `main.py`.
* **Dart Pattern:** Generates "geared" shapes for leather pads.
    * **Math:** `calculate_star_path` uses a cosine wave modified by a `shape_factor`. Spectrum: 0.0=Triangle, 0.5=Sine, 1.0=Square (with smooth blends in between).
    * **Sizing:** `get_disc_diameter` and `leather_back_wrap` handle material expansions.
* **Nesting:** `can_all_pads_fit` uses a greedy algorithm to place circles on the sheet.

### 2. Databases (Key Heights, Serials, Screws)
* **Location:** Logic in `library_features.py` (Mixin).
* **Architecture:** The main app inherits from `LibraryFeaturesMixin`. This adds the tabs for Key Heights, Serials, and Screws.
* **Serials:** `lookup_serial_year` iterates through `SERIAL_DATA` ranges. Makers with restarted numbering (Buffet, LeBlanc/Vito, Yanagisawa) hold multiple ascending series in one list; the lookup splits on the descents and reports every series that matches (e.g. "1924 or 1978" for an ambiguous Buffet serial).

---

## Development Workflows

### The "Subtraction Method" (Standalone Generation)
There is a secondary tool called the **"Standalone Pad SVG Generator"** used by colleagues.

* **Goal:** A lightweight EXE containing *only* the Pad Generator (Tab 1), with no databases.
* **Old Method:** Deleting code lines from a monolithic file.
* **New Method (Modular):**
    1.  Create a script (e.g., `main_standalone.py`).
    2.  Import `PadSVGGeneratorApp` logic *without* inheriting from `LibraryFeaturesMixin`.
    3.  Only call `create_pad_generator_tab`.
    4.  Do *not* import `serials.py` or `library_features.py`.
* **Reasoning:** The modular structure allows building the standalone tool purely by excluding imports.

---

## UI & Theming
* **Framework:** Standard `tkinter`.
* **Theme:** The "Resonance" Easter egg is managed in `main.py` and `ui_dialogs.py`. It changes global background colors (`COOL_BLUE`, `COOL_GREEN`) and window alpha.
* **Widgets:** Use `ttk` for Notebooks and Comboboxes; standard `tk` for frames/labels to support background color changes.

---

## Known Constraints & Gotchas
1.  **Settings Hierarchy:** Dart/Star settings are at the **root** of `app_settings.json`. Access via `self.settings["dart_threshold"]`.
2.  **Imports:** `svg_engine.py` must remain "pure" (no tkinter imports) to ensure easy testing.
3.  **Pad Strings:** Pad lists are multiline strings. The parser in `main.py` expects `"Size x Qty"` (or `"Size x max"`); bare `"Size"` lines do NOT parse. Lines that fail to parse are collected and surfaced in a skip-warning before generation.

## Building Executables
The project uses **PyInstaller** via GitHub Actions.
* **Dependency Discovery:** PyInstaller automatically finds `config.py`, `svg_engine.py`, etc., by following imports from `main.py`.
* **External Libs:** Ensure `svgwrite` is installed in the build environment.
