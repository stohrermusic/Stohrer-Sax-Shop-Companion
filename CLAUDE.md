# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Stohrer Sax Shop Companion is a cross-platform desktop GUI application for saxophone repair technicians. It provides SVG/G-code generation for laser-cutting pad materials, reference databases for key heights, serial numbers, and screw specifications, a tooling tab for die inserts and holders, a chromatic strobe tuner, and a harmonic tone analyzer.

## Running the Application

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

External dependencies: `svgwrite`, `numpy`, `sounddevice` (tuner/toner). The GUI uses Python's built-in `tkinter`. Requires Python 3.11+.

There is no automated test suite. Changes are verified by running the application manually.

### Testing Before Commits

Before committing, write a temporary test script in `tools/` that exercises affected code paths non-interactively (no GUI). Test engine/logic functions directly. Print PASS/FAIL per test with a summary. Cover: the specific changes, regressions in nearby code, all affected materials, and edge cases.

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

## Branching Strategy

- **`main`**: Stable release branch. Merges from `beta` when features are tested and ready.
- **`beta`**: Active development branch. New features land here first (e.g. filled engraving, air assist toggles, cut grouping). Always work on `beta` unless told otherwise.
- CI builds trigger on push to `main`, `beta`, or `gamma`.

## Versioning

`APP_VERSION` and `APP_BUILD_DATE` are defined in `config.py`. Update both when preparing a release. The About dialog reads these constants via `ui_dialogs.py`.

## CI/CD (GitHub Actions)

The `.github/workflows/build.yml` workflow automatically builds for all three platforms:
- Triggers on push to `main`, `beta`, or `gamma`, on release creation, or manually
- Builds Windows .exe, macOS .app (universal2: Intel + Apple Silicon, zipped), and Linux binary in parallel
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
    ↓ inherits (multiple mixins)
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
tooling_tab.py          → ToolingTabMixin (Die Inserts & Die Holders tab)
tuner_tab.py            → TunerTabMixin (Chromatic strobe tuner tab, StrobeWheel class)
tuner_engine.py         → TunerEngine (FFT pitch detection, phase tracking, no tkinter)
toner_tab.py            → TonerTabMixin (Harmonic tone analyzer tab, profile system, comparison)
toner_engine.py         → TonerEngine (FFT harmonic analysis, descriptors, profile storage, no tkinter)
    ↓ uses
config.py              → Settings I/O, constants, platform config paths, migration logic, import helpers
svg_engine.py          → Pure math/SVG logic (no tkinter dependency), polygon nesting
gcode_engine.py        → G-code generation for Grbl lasers, single-stroke font, circle linearization
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export, PolygonDrawWindow, GcodeSettingsWindow, PadNotesWindow, UserGuideWindow)
serials.py             → SERIAL_DATA dictionary (manufacturer → serial ranges)
build.py               → Cross-platform PyInstaller build script
```

### Key Design Patterns

**Mixin Inheritance**: `PadSVGGeneratorApp` inherits from `LibraryFeaturesMixin`, `ToolingTabMixin`, `TunerTabMixin`, and `TonerTabMixin`. Each mixin adds one or more tabs without polluting the main module.

**Pure Logic Separation**: `svg_engine.py` and `gcode_engine.py` contain no tkinter code, making them testable independently. All SVG/G-code generation math (star paths, nesting algorithm, sizing calculations) lives here.

**Tab-Specific Menus**: The app swaps menu bars when tabs change (`on_tab_changed` in main.py).

**Tab-Aware User Guide**: `UserGuideWindow` accepts an optional `section` parameter. When opened via Help > User Guide, it shows only the content relevant to the current tab. A "Show Full Guide" button expands to the complete guide. Content is split into `_section_*()` methods in ui_dialogs.py. When adding a new tab, add a corresponding section method and wire it into `_insert_content()` and the `tab_sections` dict in `open_user_guide()`.

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

**G-code Engine Details**: `gcode_engine.py` contains two font systems: `STROKE_FONT` (single-stroke outlines for "line" engraving mode) and `FILLED_FONT` (Roboto glyph outlines for "filled" raster mode). Filled engraving uses scan-line fill with even-odd rule to convert font outlines into horizontal raster lines, with optional overscan (extending lines beyond character edges so the laser reaches full speed). `generate_gcode_from_placed()` is the main entry point — it supports two cut grouping modes ("layer": all engravings → all holes → all cuts; "pad": complete each disc before moving to the next) and per-layer air assist control (M8 on / M9 off). The internal `_collect_disc_strokes()` helper extracts per-disc stroke data for both modes.

**Polygon Shape Tool**: Users can draw custom polygon shapes for irregular leather skins. Grid size adapts to unit setting: 15x15 inches (1" squares) or 40x40 cm (1cm squares). The polygon nesting algorithm (`_nest_discs_polygon()`) uses ray-casting for point-in-polygon checks and distance-to-edge calculations for circle fitting. The rectangle algorithm remains the fast path when no custom shape is defined.

**Scrap Mode**: Allows users to place pads across multiple irregular scrap pieces instead of requiring one large sheet. Key components:
- Session state in `self.scrap_session` tracks original pads, remaining pads, scrap count, locked material
- `try_nest_partial()` and `compute_remaining_pads()` in svg_engine.py handle partial placement
- `generate_svg_from_placed()` / `generate_gcode_from_placed()` generate output from pre-computed placements
- Progress popup window (`scrap_remaining_window`) shows two columns: Remaining and Done
- Materials are locked (disabled) during active session to prevent switching
- Files named with `_scrap1`, `_scrap2` suffixes
- After each scrap with a polygon loaded, a dialog asks whether to unload or keep the shape for the next piece

**Engraving Auto-Fit**: When engraved text (pad size number) would impinge on the disc edge, both engines prefer shifting the text toward the disc center over shrinking it. Only scales down as a last resort when text can't fit even when centered. SVG engine uses bounding-box corner checks; G-code engine uses actual stroke-point measurements. Both maintain a 0.5mm clearance margin. The 80% radius check (`font_size >= r * 0.8`) disables engraving entirely before auto-fit runs.

**SD Card & Eject**: Two mechanisms:
- "Eject SD card after G-code export" checkbox (Windows only, below Generate buttons): auto-ejects removable drives after G-code generation. Uses `GetDriveTypeW` to detect removable drives; silently skips non-removable destinations.
- File → Send G-code to SD Card: guided workflow that copies a .gcode file, optionally clears old files, and ejects (Windows only via PowerShell COM object).

**Strobe Tuner**: 12-wheel chromatic tuner modeled after the Peterson Stroboconn. Architecture:
- `tuner_engine.py`: Pure audio/math — FFT-based pitch detection, per-pitch-class phase tracking, per-ring (octave) magnitude extraction. `TunerEngine` manages the sounddevice input stream; `TunerResult` holds magnitudes, phase offsets, ring magnitudes, and cents errors for all 12 pitch classes. `ReferencePlayer` outputs reference tones.
- `tuner_tab.py`: All tkinter UI — `StrobeWheel` class renders one disc (annular sector polygons with wedge mask), `TunerTabMixin` builds the tab with 3-column control panel (graphic EQ sliders | flat/pilot/sharp indicator | vintage backlit VU meter). The VU needle has damped movement (lerp toward target each frame). Theme walker is bypassed via `_skip_theme` and `_dark_canvas` flags on all dark widgets.
- The tuner auto-starts/stops when switching tabs (`_tuner_start`/`_tuner_stop` called from `on_tab_changed` in main.py).
- Transposition support: wheel labels and VU readout both apply the shift from TRANSPOSITION_SHIFTS, with octave correction when the shift wraps past C.

**Tone Analyzer (Toner)**: Real-time harmonic analyzer for saxophone. Architecture mirrors the tuner:
- `toner_engine.py`: Pure audio/math — FFT with 16384-sample window (~2.7 Hz resolution), fundamental detection via peak-picking with sub-harmonic verification and temporal hysteresis, harmonic extraction up to 12th harmonic, descriptor computation (resonance, richness, brightness, darkness, fullness). `TonerEngine` manages its own sounddevice input stream independently from the tuner. Profile storage uses `load_tone_profiles()`/`save_tone_profiles()` with nested library format `{library: {profile: data}}`.
- `toner_tab.py`: All tkinter UI — `TonerTabMixin` builds the tab with spectrum canvas (left), VU-style gauge panel (right), and control strip (bottom). Auto-capture system uses a state machine (`listening` → `delay` → `recording` → `cooldown`) that detects stable tones and records 5-second averages without button presses. Profile management, comparison tool with multi-select and filtering, import/export.
- The toner auto-starts/stops when switching tabs, same as the tuner.
- Descriptors use Benade's spectral break frequency (measured by Benade/Wolfe, JASA 1988, UNSW) to define the boundary between "bright" and "dark" harmonics. Break frequencies adapt to sax type: soprano 1300 Hz, alto 837 Hz, tenor 618 Hz, bari 450 Hz. The SAX selector in the control strip sets this.
- Bias sliders on each gauge offset the display without affecting captured data. Scale toggle switches between linear (default, true amplitude ratios) and dB (logarithmic).
- Richness uses spectral flatness (geometric mean / arithmetic mean of harmonic amplitudes) combined with a coverage factor, not a simple count.
- Three capture modes: "structured" (hold notes, 5s recording), "free" (0.5s stability, continuous micro-captures while playing naturally), and "file" (import WAV, extract stable note segments offline). Each capture is tagged with its method.
- Auto-transposition: SAX selector sets both the break frequency and the displayed note names. Written pitch is shown by default (alto shows A4 when concert C4 is played). "Concert" checkbox overrides to concert pitch.
- Coverage summary dialog appears after stopping a capture session, showing a bar chart of note distribution colored by register (low/mid/high), gap assessment, and a "Resume Capturing" button to fill underrepresented registers.
- `compute_fingerprint()` averages descriptors per-note first (equal weight per note), then across notes. This prevents register skew — a profile with mostly high-note captures won't read artificially bright.

**Audio Stream Health**: Both `tuner_engine.py` and `toner_engine.py` include stream health monitoring. The `AudioRingBuffer` tracks a write counter; if the analysis loop detects no new audio data for ~1 second, the engine automatically restarts the sounddevice stream. This recovers from silent callback death on Windows.

**Tooling Tab**: Accordion-style UI with die insert and die holder SVG/G-code generation. Small dies ≤39.5mm (50mm OD), large ≥40mm (70mm OD). Die holders 85mm OD, 4-layer stack. Canvas widgets need explicit handling in the resonance theme walker.

### Settings Backward Compatibility

`load_settings()` in config.py does a two-level deep merge: top-level keys merge with defaults, and nested dicts (like `gcode_settings.felt`) also merge key-by-key. This means adding new keys to `DEFAULT_SETTINGS` works automatically for existing users — their old config file loads, and new keys get default values filled in. When adding new settings:
- Top-level keys: just add to `DEFAULT_SETTINGS`, the merge handles it
- Nested keys (e.g. inside `gcode_settings.felt`): also handled by the two-level merge
- In engine code, still use `.get(key, default)` as a safety net for any settings read outside the merge path

### Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

**Pad preset data format**: Pad presets can be either a plain string (legacy) or a dict `{"pads": "...", "notes": "..."}`. The helper `_get_pad_preset_data()` in main.py handles both formats transparently. When saving, always use the dict format. When loading, check `isinstance(raw, str)` for backward compatibility.

### Data Files (JSON, in platform config directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications
- `tone_profiles.json` - Tone analyzer profiles (nested library format, same as presets)

### Tone Profile Data Model

Profiles use nested library format: `{library_name: {profile_name: profile_data}}`. Each profile is a fixed setup (horn + player + mouthpiece + reed). Changing any variable means creating a new profile.

A profile contains sessions (date + captures). Each capture stores `harmonics_db` (list of dB values relative to fundamental, index 0 = fundamental), `descriptors` (computed resonance/richness/brightness/darkness/fullness), `method` ("structured", "free", or "file"), `n_frames`, and metadata. Harmonics below -60 dB are filtered out as noise floor.

`compute_fingerprint()` in `toner_engine.py` aggregates all sessions into an overall average and per-note breakdown. The `per_note` dict maps note names to averaged harmonic data, enabling note-by-note comparison across horns.

Flat legacy format (profiles at top level without library wrapper) is auto-migrated on load.

### Toner Calibration

Profile notes can contain subjective tone descriptions ("rich horn", "very bright", "dark and warm"). The `tools/calibrate_toner.py` script scans all annotated profiles, extracts keywords, compares them against computed descriptors, and reports alignment. The `tools/analyze_horn_spread.py` script computes statistical spread of descriptors across all profiles — min/max/mean/stddev per descriptor, per-note variation, gauge scaling suggestions, and grouping analysis by horn type and manufacturer. Both tools help identify where the descriptor scaling constants need adjustment. Bias sliders in the UI provide per-user visual calibration without affecting captured data.

**Descriptor calibration status**: The descriptor formulas are functional but not yet calibrated against real-world data. The break frequencies are from published measurements (Benade/Wolfe) but the scaling constants are educated guesses. Pending: import a large set of recordings (same player, different horns) to establish per-note baselines and empirical descriptor ranges. The goal is gauges that show "brighter than typical for this note on this horn type" rather than raw energy ratios.

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
