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
i18n.py                → gettext catalog init; must run in main.py BEFORE any UI module import
svg_engine.py          → Pure math/SVG logic (no tkinter dependency), polygon nesting
gcode_engine.py        → G-code generation for Grbl lasers, single-stroke font, circle linearization
ui_dialogs.py          → Dialog window classes (Options, Colors, Import/Export, PolygonDrawWindow, GcodeSettingsWindow, PadNotesWindow, UserGuideWindow, WelcomeWindow, QuickStartWindow, CameraCalibrationDialog, CameraCaptureDialog, FalconRunDialog, LiveCameraWindow)
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

**macOS Quit Path**: `WM_DELETE_WINDOW` only covers the window-close button. On macOS, Cmd-Q and the app menu's Quit fire Tk's `::tk::mac::Quit`, whose default exits the process without running the close handler — settings would silently never save. `main.py` registers `root.createcommand("::tk::mac::Quit", self.on_exit)` on darwin (added v2.61). Any future work on the exit path must keep both routes pointing at `on_exit`.

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

- `app_settings.json` - User preferences (including `welcomed_version`, the version whose Welcome window this install has seen)
- `pad_presets.json` - Saved pad size lists
- `sizing_presets.json` - Saved Sizing Rules dialog presets (added v2.0; schema in config.py `SIZING_PRESET_KEYS`)
- `gcode_presets.json` - Per-material G-code laser presets (added v2.6; nested `{material: {preset_name: data}}`; schema in config.py `GCODE_PRESET_KEYS` / `GCODE_PRESET_MATERIALS`)
- `job_history.json` - Log of pad jobs that reached an output stage (added v2.65; `{"version": 1, "jobs": [...]}`, newest first, capped at `JOB_HISTORY_LIMIT`)
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

`setup_logging()` in config.py creates a rotating log file (`app.log`) in the config directory. `main.py` hooks both `sys.excepthook` and tkinter's `report_callback_exception` so any unhandled exception gets a full traceback written to the log and a dialog shown to the user pointing them to Help > Open Log File. The log rotates at 2 MB (`LOG_MAX_BYTES`) with 1 backup.

The root level is WARNING, but the loggers in `config.INFO_LOGGERS` (`falcon_sender`, `camera_capture`, `ssc.machine`, `ssc.startup`) write INFO: until v2.86 every INFO line `falcon_sender` had (stream start/end, wake-up, `$X`, the M-code trace) was silently dropped because only the root level was set. Things that repeat on a timer must not log per tick: the calibration preview logs a 5 s summary, and `falcon_sender._log_throttled` holds a repeating connect message to one line per 30 s (the calibration MPos poll connects every 0.8 s). At startup `config.log_runtime_versions` writes one line with Python, OS, Tk and library versions (already-imported modules report their own `__version__`; others come from metadata without importing them), and `camera_capture` logs the OpenCV it loads. Help > Open Log File shows the file *selected* in Explorer/Finder (`reveal_in_file_manager_command`) so it can be attached to an email next to `last_calibration_frame.png` and `last_frame_cut.gcode`.

## macOS Build

`build.py` patches Info.plist after PyInstaller to add `NSMicrophoneUsageDescription` and `NSCameraUsageDescription`. Without those keys, macOS silently denies mic access to the tuner and toner (and camera access to the machine-integration features).

The keys alone are **not sufficient**: PyInstaller ad-hoc signs the bundle during build, and Info.plist is sealed into that signature. Patching the plist afterwards breaks the seal, and macOS TCC refuses to show a permission prompt for any app whose signature doesn't validate — same silent-denial symptom, keys present. Every shipped Apple Silicon build through v2.62 had a broken seal (verified against the released zip: sealed plist hash ≠ actual plist hash). `_resign_macos_app()` in build.py therefore re-signs ad-hoc (`codesign --force --deep --sign -`) immediately after the patch and verifies with `--strict`; any future work on the macOS build must keep the re-sign as the last step that touches the bundle.

Packaging must preserve symlinks too: CI zips the `.app` with `ditto -c -k --keepParent`, because `zip -r` follows the bundle's internal Frameworks↔Resources symlinks and stores them as duplicate real files (767 of them in the v2.62 zip, roughly doubling the download), which breaks the signature's resource seal on extraction. The `--dmg` path is fine as-is — hdiutil preserves symlinks. CI verifies all of it on both macOS runners: `plutil -extract` for both usage keys, `codesign --verify --deep --strict` on the built .app, and again on an unzipped copy of the final artifact. Users upgrading from a broken build may need `tccutil reset Microphone com.stohrer.saxshopcompanion` (and `Camera`) if macOS cached the old silent denial — the README documents this.

## Welcome window, Quick Start Guide, User Guide search (2026-11-09)

Matt: "each install shows a quick welcome, a version number, and thank you for installing, be sure to search the user guide with any questions, and says click here for the quick start guide", and of the guide: "whatever guide we write should be SUPER concise". He rejected captions under the Generate buttons as too busy, so nothing was added to the main tabs.

**Welcome window** (`ui_dialogs.WelcomeWindow`). Shown once for each version: `main.maybe_show_welcome()` compares `settings["welcomed_version"]` (in `DEFAULT_SETTINGS`, "" for never) with `APP_VERSION`, records the version **when the window is shown** (a crash must not bring it back every launch) and opens it. So a new install sees it once, and an existing user once after each update. Help > Welcome opens it any time (`open_welcome`, one window at a time). **`maybe_show_welcome` is called from exactly one place: the `__main__` block of main.py, under `if not tour:`, on an `after(600)`** once the main window is up. Building the app never calls it, so `--selftest`, `--tour` and every suite that constructs `PadSVGGeneratorApp` never see a Welcome window they did not ask for, and startup is no slower. Keep it that way: a call in `__init__` would pop a window in dozens of test runs.
- *Content, and nothing more*: the app name and version, "Thank you for installing.", "Questions? Search the User Guide." (a link: `open_user_guide_search`, the full guide with the cursor in its search box), "Click here for the quick start guide" (a link: `open_quick_start`), and Close.
- *Look: plain, on purpose.* The app's ordinary dialog look (`DIALOG_BG`, the usual fonts, links as in the About dialog). Not modal: it takes no grab. **Do not style it.** The first version was a nod to an old dial-up sign-on screen (three framed panels that lit up one after another, then the text); the same day Matt first asked for no animation ("just a quick popup that doesn't get in the way") and then for no styling at all: "let's have the welcome screen be simple. if I have a detailed ui request I will give it." It has no canvas, no timers and no colours of its own, and the suite checks that none come back.

**Quick Start Guide** (`ui_dialogs.QuickStartWindow`, Help > Quick Start Guide, above User Guide). One small page of labels: three ways to cut (SD card, LightBurn, Frame & Cut), "Do not mix them", and what to check before cutting. Each line is its own translatable string (`QuickStartWindow.lines`), because the wording is Matt's to change. The Frame & Cut line is only there when `machine_supported()`. No images, no scrolling (the window is as big as its text), not modal. It never opens by itself.

**User Guide search.** The Welcome window tells people to search the guide, so `UserGuideWindow` has a search box: typing highlights every match (`search_hit`) and jumps to the first (`search_current`), Enter or Next goes to the next and wraps round, the count reads "3 of 12" or "No matches", Ctrl+F (Cmd+F on macOS) puts the cursor there. Case-insensitive, plain text. It always searches the whole guide: opened on one tab's section, the first search calls `_show_all()`.

**The tour has a stop for each** ("welcome" and "quick-start", after "about"; 20 stops now), so both are in the Mac screenshots.

Tests: `tools/test_welcome.py` (9): once per version from a fresh folder and from another stored version; the single call site and `--selftest` recording nothing; the window being a plain dialog (five lines and Close, the dialog background, no canvas, nothing scheduled, no grab, none of the animated version's attributes); both links; the guide's eleven lines; Help on every tab; search (count, highlight, wrap, case, no matches, a one-tab guide expanding); and, in a German child process, both windows fitting.

## Feature Set

File > Feature Set lets users show/hide tabs. It only lists what the build can run: no Tuner/Toner on a Mac without the audio stack (`main.audio_supported`), no Machine menu toggle without pyserial (`main.machine_supported`) — i.e. the Intel Mac build. Saving merges into `visible_tabs` rather than replacing it, so a tab the build doesn't offer keeps its saved value. As of v2.0 the Tuner is default-on (no longer marked experimental); only the Toner remains hidden by default. Toner requires a one-time beta terms acceptance dialog (scrollable text explaining beta status, three ways to use it, and setup steps: input device, recording folder, preset fields, required preset fields). Acceptance is stored in `toner_unlocked` in settings. The `visible_tabs` dict in settings controls which tabs are shown.

**Machine Integration (v2.5+, experimental, opt-in)** — File > Feature Set also exposes an `experimental_machine_menu` toggle (default False). When on, the Pad Maker tab gains an Options > Machine submenu, a Frame & Cut button, and the polygon-draw dialog grows its "Get from camera" + live overlay UI. When off, none of that surfaces — the app behaves like pre-v2.5. Helper `_machine_enabled()` in main.py is the single source of truth and is called by `_refresh_machine_ui_state()` whenever the toggle changes or the calibration file appears/disappears.

**Gating: the laser and the camera are gated separately (since 2026-11-09).** Matt was asked "Should Frame & Cut require a camera calibration?" and answered "No, on either tab". `_refresh_machine_ui_state()` is the one place it is enforced:

- *Toggle off* (or no pyserial in the build, `machine_supported()`): nothing machine-related shows; the Machine menu is not built at all.
- *Toggle on*: the Machine menu's laser items are enabled (Home Laser, Test Connection, Clear Errors, Reset Falcon, Framing Power: `main.MACHINE_MENU_LASER_ITEMS`), and Camera Calibration as always. Once a laser is detected (`falcon_port`), the Frame & Cut button shows on the Pad Maker and on every Tooling section that writes G-code. One rule for both tabs, no calibration. `falcon_port` is set and cleared by the laser watch (CLAUDE-engines.md "The laser watch"), on the Tk thread, for the whole session: the button appears within seconds of the laser being plugged in after launch and goes again when its port is no longer listed (which also clears `_falcon_homed_this_session`). Before 2026-11-09 the app probed the ports until a laser answered and never looked again.
- *Toggle on and a camera calibration on disk*: everything that uses the camera. Camera-Polygon Inset Margin in the menu (`MACHINE_MENU_CAMERA_ITEMS`), the "Live camera preview" tick, and through `_camera_capture_ready()` (the toggle and `_camera_calibration_present()`) "Get from camera", the live overlay in the shape dialog, Try Auto Locate and "Re-capture from camera".

Before that date the calibration gated all of it: without one, only "Camera Calibration..." was enabled and Frame & Cut was hidden (commit `021d8df`, 2026-05-26, "Two-level machine gating"). That was a rule, not a dependency: typed dimensions or a hand-drawn shape, jog to the corner, G92, framing loop, cut never read the calibration. With no calibration a Frame & Cut run opens no camera and has no Auto Locate target (`on_frame_and_cut` asks `_camera_capture_ready()` before either), even with a stale tick or position. The toggle itself is still restart-required, as the Feature Set dialog says; the buttons and items follow laser detection and calibration live.

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
      │   ├─ Try Auto Locate (drives head to polygon anchor vertex, G53)
      │   └─ Start Frame →
      ├─ Framing loop dialog (FalconRunDialog loop=True)
      │   └─ generate_polygon_framing_gcode (outline, BL-first rotation, G92 with LB offset)
      └─ Cut dialog (FalconRunDialog) — same G92 prefix as framing
  Tooling > Frame & Cut buttons (Die Inserts, Die Holders, Die Organizer, Kerf Test,
                                 Speed & Power Test, Tonehole Crescent)
      └─ main.on_frame_and_cut(tooling_job=...) — the same routine from the
         framing lines on; bounding-box framing, G92 X0 Y0, no polygon, no
         Auto Locate
```

Since 2026-11-09 the laser branches of this map (the Machine menu's Home / Test / Clear / Reset / Framing Power, and Frame & Cut on both tabs) need only the toggle and a detected laser. The camera branches (Camera-Polygon Inset Margin, the polygon dialog's overlay and "Get from camera", Try Auto Locate) still need a calibration. See "Feature Set" above.

Key invariants the subsystem keeps:
- `custom_polygon` (inset, normalized to outline's origin) drives nesting placements.
- `custom_polygon_outline` (un-inset, normalized to outline's origin) drives framing trace.
- `_custom_polygon_lb_machine` (absolute machine coords of the outline's anchor vertex, `svg_engine.polygon_anchor_vertex` — the same vertex the G92 uses) drives Try Auto Locate.
- `_falcon_homed_this_session` (True after $H or after Machine > Reset Falcon; False again when the laser's port vanishes, since a power-cycled machine has lost its zero) gates Try Auto Locate enabled state.
- `falcon_port` comes from the laser watch and nothing else: `_on_falcon_detected` / `_on_falcon_lost` on the Tk thread, driven by `falcon_sender.LaserWatch` on a worker. `_detect_falcon_async()` is the "look now" entry (a Machine menu item or Frame & Cut with no laser, a connect that failed): it asks the watch for a full rescan and is a no-op when no watch runs, so suites and tours never open a serial port.
- Every camera window opens its camera through `camera_capture.open_camera_or_fallback(index, saved_index)`: when the saved number (`camera_index_override`) fails, the Falcon-named camera and then the last other camera are tried once, the window runs on what opened, one warning is logged, and the saved number is not changed (only a calibration save or Switch camera changes it, as before).
- Frame G-code emits `G92 X{lb_x} Y{lb_y}` (LB-vertex in Y-flipped frame) so jogging to the visible scrap corner bridges to the polygon's bbox origin.
- A job that was not nested on the loaded shape (card on "Fit card to paper", and every Tooling job) takes nothing from it: `on_shape` in `on_frame_and_cut` is False, so no outline is traced, the G92 is `X0 Y0`, the cut is not flipped with the shape's height and Try Auto Locate has no target.
- Everything after a successful `sender.connect()` runs inside the `try` whose `finally` disconnects, and the sleep lock is taken after the last early return, so no way out leaves the port open or the lock held (`tools/test_release_3_0.py`).
- `falcon_sender._stream_loop` starts every stream with two safety lines BEFORE the caller's G-code: (1) `G4 P0.01` wake-up dwell to prevent Grbl parser dormancy from eating the first command, and (2) `$X` unlock so a sticky Alarm from the previous stream (e.g. soft-limit triggered by a jog-shifted cut bbox in Frame & Cut) doesn't force a Falcon power cycle. `$X` is a no-op when Grbl is already Idle.
