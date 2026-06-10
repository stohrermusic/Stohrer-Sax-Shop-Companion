# Architecture

## Module Structure

```
main.py                 → Entry point, PadSVGGeneratorApp class, tab creation
    ↓ inherits (multiple mixins)
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
tooling_tab.py          → ToolingTabMixin (Die Inserts & Die Holders tab)
tuner_tab.py            → TunerTabMixin (Chromatic strobe tuner tab, StrobeWheel class)
tuner_engine.py         → TunerEngine (FFT pitch detection, phase tracking, no tkinter)
toner_tab.py            → TonerTabMixin (Harmonic tone analyzer tab, preset system, comparison)
toner_engine.py         → TonerEngine (FFT harmonic analysis, descriptors, preset storage, no tkinter)
audio_utils.py          → AudioRingBuffer (shared audio stream health monitoring)
tuner_renderer/         → Rust/wgpu GPU renderer for strobe tuner (PyO3 bindings; Windows/Linux only — macOS is canvas-only because Tk Aqua's winfo_id() isn't an NSView; canvas fallback when absent)
    ↓ uses
config.py              → Settings I/O, constants, platform config paths, migration logic, import helpers
svg_engine.py          → Pure math/SVG logic (no tkinter dependency), polygon nesting
gcode_engine.py        → G-code generation for Grbl lasers, single-stroke font, circle linearization
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export, PolygonDrawWindow, GcodeSettingsWindow, PadNotesWindow, UserGuideWindow, CameraCalibrationDialog, CameraCaptureDialog, FalconRunDialog, LiveCameraWindow)
serials.py             → SERIAL_DATA dictionary (manufacturer → serial ranges)
camera_capture.py       → OpenCV ChArUco calibration + scrap-polygon detection (optional dep; degrades gracefully if OpenCV missing)
falcon_sender.py        → Grbl streamer over USB serial (character-counting protocol; pyserial-optional). Includes a hard-wake-up via G4 P0.01 dwell + `$X` safety-unlock at stream start.
sleep_lock.py           → Cross-platform "keep the system awake during long operations" wrapper. Used by CameraCalibrationDialog so Windows / macOS don't suspend mid-engrave.
build.py               → Cross-platform PyInstaller build script
```

## Key Design Patterns

**Mixin Inheritance**: `PadSVGGeneratorApp` inherits from `LibraryFeaturesMixin`, `ToolingTabMixin`, `TunerTabMixin`, and `TonerTabMixin`. Each mixin adds one or more tabs without polluting the main module.

**Pure Logic Separation**: `svg_engine.py` and `gcode_engine.py` contain no tkinter code, making them testable independently. All SVG/G-code generation math (star paths, nesting algorithm, sizing calculations) lives here.

**Tab-Specific Menus**: The app swaps menu bars when tabs change (`on_tab_changed` in main.py).

**Tab-Aware User Guide**: `UserGuideWindow` accepts an optional `section` parameter. When opened via Help > User Guide, it shows only the content relevant to the current tab. A "Show Full Guide" button expands to the complete guide. Content is split into `_section_*()` methods in ui_dialogs.py. When adding a new tab, add a corresponding section method and wire it into `_insert_content()` and the `tab_sections` dict in `open_user_guide()`.

**Cross-Platform Helpers**: `bind_mousewheel()` in ui_dialogs.py handles platform-specific scroll behavior (Windows/macOS/Linux).

**macOS Theming**: On macOS (`IS_MACOS` flag in main.py and ui_dialogs.py), the app uses native system colors instead of custom cream/beige backgrounds. This allows the app to work correctly in both macOS dark and light mode. The `DIALOG_BG` constant in ui_dialogs.py resolves to `"systemWindowBackgroundColor"` on macOS (a Tk system color that adapts to dark/light mode) and `"#F0EAD6"` on Windows/Linux. The resonance theme system is disabled on macOS. When adding new UI widgets, use `DIALOG_BG` for dialog backgrounds and `self.root.cget('bg')` for main window widgets.

## Settings Backward Compatibility

`load_settings()` in config.py does a two-level deep merge: top-level keys merge with defaults, and nested dicts (like `gcode_settings.felt`) also merge key-by-key. This means adding new keys to `DEFAULT_SETTINGS` works automatically for existing users — their old config file loads, and new keys get default values filled in. When adding new settings:
- Top-level keys: just add to `DEFAULT_SETTINGS`, the merge handles it
- Nested keys (e.g. inside `gcode_settings.felt`): also handled by the two-level merge
- In engine code, still use `.get(key, default)` as a safety net for any settings read outside the merge path

## Preset/Library System

All presets use a nested dictionary structure: `{library_name: {preset_name: data}}`. Flat legacy formats are auto-migrated on load (see `load_presets()` in config.py).

**Pad preset data format**: Pad presets can be either a plain string (legacy) or a dict `{"pads": "...", "notes": "..."}`. The helper `_get_pad_preset_data()` in main.py handles both formats transparently. When saving, always use the dict format. When loading, check `isinstance(raw, str)` for backward compatibility.

## Data Files (JSON, in platform config directory)

- `app_settings.json` - User preferences
- `pad_presets.json` - Saved pad size lists
- `sizing_presets.json` - Saved Sizing Rules dialog presets (added v2.0; schema in config.py `SIZING_PRESET_KEYS`)
- `gcode_presets.json` - Per-material G-code laser presets (added v2.6; nested `{material: {preset_name: data}}`; schema in config.py `GCODE_PRESET_KEYS` / `GCODE_PRESET_MATERIALS`)
- `key_height_library.json` - Saxophone key height measurements
- `screw_specs.json` - OEM screw/rod specifications
- `toner_data.json` - Tone analyzer presets and sessions (nested library format; auto-migrated from old `tone_profiles.json`)

## Settings Persistence Pattern

Every setting read or written at runtime MUST exist in `DEFAULT_SETTINGS` in config.py. The `load_settings()` merge only preserves keys that exist in the defaults — runtime-only keys get silently dropped on next load. When adding a new setting, always add it to `DEFAULT_SETTINGS` first.

## Settings Loading Hardening

`load_settings()` defends against old/corrupted config files: null values are rejected (defaults used instead), dict-type keys that load as non-dicts are skipped, layer_colors with old LightBurn codes (e.g. "C10") are replaced with hex defaults, and `KeyError`/`AttributeError`/`ValueError` during merge fall back to full defaults. Users upgrading from very old versions (tested back to v1.0) should never crash on settings load.

## Data Safety on Write (v2.6+)

All JSON saves route through `_write_json_atomic()` in config.py — write to a sibling `.tmp` file, then `os.replace()` over the target — so a crash or power loss mid-save leaves either the old file or the complete new one, never a truncated half-write. Both `save_settings()` and `save_presets()` use it. On the read side, when a config file fails to parse, `_preserve_corrupt_file()` copies it to `<name>.corrupt.bak` before defaults take over (first corruption wins — an existing backup is never overwritten), so hand-recoverable data isn't destroyed by the next save. Both `load_settings()` and `load_presets()` call it.

## Error Logging

`setup_logging()` in config.py creates a rotating log file (`app.log`) in the config directory. `main.py` hooks both `sys.excepthook` and tkinter's `report_callback_exception` so any unhandled exception gets a full traceback written to the log and a dialog shown to the user pointing them to Help > Open Log File. The log rotates at 500KB with 1 backup.

## macOS Build

`build.py` patches Info.plist after PyInstaller to add `NSMicrophoneUsageDescription`. Without this, macOS silently denies mic access to the tuner and toner.

## Feature Set

File > Feature Set lets users show/hide tabs. As of v2.0 the Tuner is default-on (no longer marked experimental); only the Toner remains hidden by default. Toner requires a one-time beta terms acceptance dialog (scrollable text explaining beta status, three ways to use it, and setup steps: input device, recording folder, preset fields, required preset fields). Acceptance is stored in `toner_unlocked` in settings. The `visible_tabs` dict in settings controls which tabs are shown.

**Machine Integration (v2.5+, experimental, opt-in)** — File > Feature Set also exposes an `experimental_machine_menu` toggle (default False). When on, the Pad Maker tab gains an Options > Machine submenu, a Frame & Cut button, and the polygon-draw dialog grows its "Get from camera" + live overlay UI. When off, none of that surfaces — the app behaves like pre-v2.5. Helper `_machine_enabled()` in main.py is the single source of truth and is called by `_refresh_machine_ui_state()` whenever the toggle changes or the calibration file appears/disappears.

The Machine submenu has TWO-LEVEL GATING: experimental toggle on AND a valid camera calibration on disk. Without the toggle: nothing shows. With the toggle but no calibration: only "Camera Calibration..." is enabled — everything else (Home Laser, Test Connection, etc.) greys out. With both: full menu live. This is enforced in `_refresh_machine_ui_state()`. Frame & Cut button visibility follows the same logic via `_camera_capture_ready()`.

## v2.5 Machine Integration Architecture

A condensed map of the new (v2.5) cross-cutting subsystem since it touches several modules:

```
File > Feature Set toggle (experimental_machine_menu)
  ↓ (when on)
main._machine_enabled() + _refresh_machine_ui_state()
  ├─ Options > Machine submenu items
  │   ├─ Home Laser → falcon_sender.FalconSender.home($H)
  │   ├─ Test Connection → falcon_sender.FalconSender.get_status(?)
  │   ├─ Clear Errors ($X) → falcon_sender.FalconSender.unlock
  │   ├─ Reset Falcon → falcon_sender.FalconSender soft-reset (Ctrl-X)
  │   ├─ Camera Calibration → ui_dialogs.CameraCalibrationDialog
  │   │   ├─ Phase 1: engrave card (basswood preset → gcode_engine.generate_calibration_card_gcode)
  │   │   └─ Phase 2: ChArUco captures → camera_capture.calibrate_from_frames
  │   │       → saves homography_px_to_machine_mm JSON
  │   └─ Camera-Polygon Inset Margin → settings.camera_polygon_inset_mm
  ├─ Polygon-draw dialog (PolygonDrawWindow)
  │   ├─ "Show live camera underneath" → _render_overlay (cv2.warpPerspective)
  │   ├─ "Get from camera" → ui_dialogs.CameraCaptureDialog → detect_scrap_contour
  │   └─ Submit → main._adopt_camera_polygon (inset + outline split, shared origin)
  └─ Pad Maker > Frame & Cut button
      ├─ jog-to-position dialog (FalconRunDialog jog-only mode)
      │   ├─ Home Laser button
      │   ├─ Jog cluster
      │   ├─ Try Auto Locate (drives head to polygon LB-vertex)
      │   └─ Start Frame →
      ├─ Framing loop dialog (FalconRunDialog loop=True)
      │   └─ generate_polygon_framing_gcode (outline, BL-first rotation, G92 with LB offset)
      └─ Cut dialog (FalconRunDialog) — same G92 prefix as framing
```

Key invariants the subsystem keeps:
- `custom_polygon` (inset, normalized to outline's origin) drives nesting placements.
- `custom_polygon_outline` (un-inset, normalized to outline's origin) drives framing trace.
- `_custom_polygon_lb_machine` (absolute machine coords of original polygon's LB-vertex) drives Try Auto Locate.
- `_falcon_homed_this_session` (True after $H or after Machine > Reset Falcon) gates Try Auto Locate enabled state.
- Frame G-code emits `G92 X{lb_x} Y{lb_y}` (LB-vertex in Y-flipped frame) so jogging to the visible scrap corner bridges to the polygon's bbox origin.
- `falcon_sender._stream_loop` starts every stream with two safety lines BEFORE the caller's G-code: (1) `G4 P0.01` wake-up dwell to prevent Grbl parser dormancy from eating the first command, and (2) `$X` unlock so a sticky Alarm from the previous stream (e.g. soft-limit triggered by a jog-shifted cut bbox in Frame & Cut) doesn't force a Falcon power cycle. `$X` is a no-op when Grbl is already Idle.
