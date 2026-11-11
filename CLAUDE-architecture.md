# Architecture

## Module Structure

```
main.py                 → Entry point, PadSVGGeneratorApp class, tab creation
    ↓ inherits (multiple mixins)
library_features.py     → LibraryFeaturesMixin (Key Heights, Serial Lookup, Screw Specs tabs)
tooling_tab.py          → ToolingTabMixin (Die Inserts & Die Holders tab)
tuner_tab.py            → TunerTabMixin (Chromatic strobe tuner tab, StrobeWheel class)
tuner_engine.py         → TunerEngine (FFT pitch detection, phase tracking, no tkinter)
toner_tab.py            → TonerTabMixin (the Analyzer tab: the intonation chart, the HARMONICS switch and the harmonic analyzer behind it, the preset system, comparison; the code keeps its toner_ names)
toner_engine.py         → TonerEngine (FFT harmonic analysis, descriptors, preset storage, no tkinter)
intonation.py           → each note's cents as played from the analyzer's captures (the Analyzer's intonation chart data; pure, no tkinter)
intonation_chart.py     → the chart as primitives from intonation.overlay, and three renderers of the same primitives: tk.Canvas, PNG (Pillow), a one-page PDF written by hand (no tkinter)
machine_control.py      → MachineControlMixin (Options > Machine, Frame & Cut, the laser watch, the camera plumbing: the machine half of the app, not a tab; see The machine mixin)
audio_utils.py          → AudioRingBuffer (shared audio stream health monitoring)
threads.py              → start_worker / GcHeldRLock / gc_held: every worker thread starts here, with the cyclic collector held off while the Tk thread waits on a worker (see Worker threads)
tuner_renderer/         → Rust/wgpu GPU renderer for strobe tuner (PyO3 bindings; Windows/Linux only — macOS is canvas-only because Tk Aqua's winfo_id() isn't an NSView; canvas fallback when absent)
    ↓ uses
config.py              → Settings I/O, constants, platform config paths, migration logic, import helpers
settings_schema.py     → the type of every setting (a Field per key: kind, enumeration or range, None or not; the nested blocks' own key schemas) and validate_settings, which load_settings runs at the door
i18n.py                → gettext catalog init; must run in main.py BEFORE any UI module import
svg_engine.py          → Pure math/SVG logic (no tkinter dependency), polygon nesting
gcode_engine.py        → G-code generation for Grbl lasers, single-stroke font, circle linearization
ui_dialogs/            → The dialog windows, a package (2026-11-10; moves only, cut by what the dialogs share):
    __init__.py          re-exports every name; `from ui_dialogs import X` and `ui_dialogs.X` are the API
    common.py            DIALOG_BG, IS_MACOS, Tooltip and add_tooltip(s), bind_mousewheel, ConfirmationDialog, material_display_name, gcode_file_note, mlog
    sizing.py            OptionsWindow (Sizing Rules), PadPreviewWindow, SaveSizingPresetDialog, LayerColorWindow
    gcode_settings.py    GcodeSettingsWindow
    presets_io.py        Export/Import/WebImport/ImportTarget windows, KeyLayoutWindow, PadNotesWindow
    previews.py          NestingPreviewWindow, SpeedPowerTestPreview
    history.py           JobHistoryWindow
    guide.py             UserGuideWindow, QuickStartWindow, WelcomeWindow, AboutDialog
    machine.py           PolygonDrawWindow, CameraCalibrationDialog, CameraCaptureDialog, LidClosedDialog, FalconRunDialog, LiveCameraWindow
    resonance.py         the resonance theme windows
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

**The User Guide**: `UserGuideWindow` (ui_dialogs/guide.py) is always the whole guide, with a clickable contents list, a search box and drawn figures; `section` is the key it opens at. Every paragraph is its own `_()` string. When adding a tab or feature, add a `_section_*` method (or a `_sub`), call it from `_insert_content` in order, and add the tab's key to `SECTION_KEYS` and `tab_sections`. Figures are canvases drawn from the app's own geometry helpers (`_fig_<name>`, a height in `FIGURE_HEIGHTS`), labels through `_()`, colours explicit, nothing scheduled. Drive the contents list with a click binding, never `<<TreeviewSelect>>` (they would select each other for ever). The window opens maximized and is not transient. Tests: `test_user_guide`. History: `docs/decisions/user-interface.md`.

**Cross-Platform Helpers**: `bind_mousewheel()` in ui_dialogs/common.py handles platform-specific scroll behavior (Windows/macOS/Linux).

**The ui_dialogs package** (2026-11-10): add a dialog to the module whose dialogs it shares code with, and its name to `ui_dialogs/__init__.py` (both the import and `__all__`); every caller imports from the package, never from a submodule. A dialog that another dialog constructs and that a suite stands in for on the package (`ui_dialogs.FalconRunDialog = Fake`) is looked up through the package at call time (`_pkg.FalconRunDialog(...)`), never bound at import; the five are named in each module's `_pkg` line. Why and how it was cut: `docs/decisions/user-interface.md`, "ui_dialogs becomes a package".

**The machine mixin** (2026-11-10): the Machine menu handlers, Frame & Cut, the laser watch and the camera plumbing are `machine_control.MachineControlMixin`, mixed into `PadSVGGeneratorApp` beside the tab mixins; the scrap session stays in main. The three constants `MACHINE_MENU_LASER_ITEMS`, `MACHINE_MENU_CAMERA_ITEMS` and `LASER_WATCH_PERIOD_MS` live there and main re-imports them, so `main.<name>` keeps working. A main-level name that a moved method reads and a suite replaces on main (`main.NestingPreviewWindow`, `main.machine_supported`) is read through `machine_control._FromMain` at the call; every other name is a plain import. Add a name to `_FromMain` only on both counts. `on_frame_and_cut` is a sequence of `_frame_cut_*` steps (CLAUDE-engines.md, "The steps of `on_frame_and_cut`"). Why and how it was cut: `docs/decisions/machine-integration.md`, "The machine half of main.py becomes a mixin".

**Worker threads** (2026-11-10): start every thread through `threads.start_worker()`; take the two locks a worker holds while the Tk thread may wait on them (`CameraCalibrationDialog._port_lock`, `FalconSender._lock`) as `threads.GcHeldRLock`; put any bounded `join()` on the Tk thread inside `threads.gc_held()`; otherwise never join or block on a worker from the Tk thread, poll it with `after()`. The cyclic collector is off while the Tk thread is blocked on a worker, so a worker can never finalise a tkinter Variable (a Tcl call from the wrong thread, waited on by the Tk thread) while the Tk thread waits for it. Gated by `tools/test_threads.py` (an AST scan fails on any `threading.Thread(` in a root module outside threads.py). The deadlock this prevents, its stack and the measurements: `docs/decisions/machine-integration.md`, "The Tk-thread / worker deadlock".

**macOS Quit Path**: `WM_DELETE_WINDOW` only covers the window-close button. On macOS, Cmd-Q and the app menu's Quit fire Tk's `::tk::mac::Quit`, whose default exits the process without running the close handler — settings would silently never save. `main.py` registers `root.createcommand("::tk::mac::Quit", self.on_exit)` on darwin (added v2.61). Any future work on the exit path must keep both routes pointing at `on_exit`.

**macOS Theming**: On macOS (`IS_MACOS` flag in main.py and ui_dialogs/common.py), the app uses native system colors instead of custom cream/beige backgrounds. This allows the app to work correctly in both macOS dark and light mode. The `DIALOG_BG` constant in ui_dialogs/common.py resolves to `"systemWindowBackgroundColor"` on macOS (a Tk system color that adapts to dark/light mode) and `"#F0EAD6"` on Windows/Linux. The resonance theme system is disabled on macOS. When adding new UI widgets, use `DIALOG_BG` for dialog backgrounds and `self.root.cget('bg')` for main window widgets.

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

Every setting read or written at runtime MUST exist in `DEFAULT_SETTINGS` in config.py AND in `settings_schema.SCHEMA` (its type, and its enumeration or range where the code enforces one). The `load_settings()` merge only preserves keys that exist in the defaults — runtime-only keys get silently dropped on next load (six were, until the gate found them on 2026-11-10: the live-camera tick and the Die Organizer's and Crescent's choices). `tools/test_settings_schema.py` fails on a key in `DEFAULT_SETTINGS` and not in the schema or the reverse, at every level, and on any string-literal settings key in the app's modules that the schema does not know (`settings["k"]`, `.get("k"`, `"k" in settings`, `self.settings`, `self._settings`, `DEFAULT_SETTINGS` and the local names assigned from them, one level down through the known sub-dicts). What it cannot see: a dict passed under another name (`source`, `preset`, a range rule `r`, `sizing`, `mat_settings`); it reports how many such accesses there are.

## Settings Loading Hardening

`load_settings()` defends against old/corrupted config files: null values are rejected (defaults used instead), dict-type keys that load as non-dicts are skipped, layer_colors with old LightBurn codes (e.g. "C10") are replaced with hex defaults, and `KeyError`/`AttributeError`/`ValueError` during merge fall back to full defaults. Users upgrading from very old versions (tested back to v1.0) should never crash on settings load.

**Validation at the door** (2026-11-10): after that merge and the migrations, `settings_schema.validate_settings` puts every value of the wrong type or outside its declared values back to its default and logs it once on `ssc.settings` with the key, the value and the file; an int is accepted where a float is declared, a bool never counts as a number, a nested value is replaced alone, a list is checked as a whole. A fault in the check itself is logged and the merged values are used (the file is never archived as corrupt for it). Declare a range or an enumeration only where the code already enforces or assumes it (a dialog's radios, a slider's ends, a clamp, an engine's requirement); a number nothing bounds gets its type and a comment saying so. No limit is invented in the schema. Presets are not validated on load (see the decisions file for why). History and the first scan's findings: `docs/decisions/settings.md`.

## Data Safety on Write (v2.6+)

All JSON saves route through `_write_json_atomic()` in config.py — write to a sibling `.tmp` file, then `os.replace()` over the target — so a crash or power loss mid-save leaves either the old file or the complete new one, never a truncated half-write. Both `save_settings()` and `save_presets()` use it. On the read side, when a config file fails to parse, `_preserve_corrupt_file()` copies it to `<name>.corrupt.bak` before defaults take over (first corruption wins — an existing backup is never overwritten), so hand-recoverable data isn't destroyed by the next save. Both `load_settings()` and `load_presets()` call it.

## Error Logging

`setup_logging()` in config.py creates a rotating log file (`app.log`) in the config directory. `main.py` hooks both `sys.excepthook` and tkinter's `report_callback_exception` so any unhandled exception gets a full traceback written to the log and a dialog shown to the user pointing them to Help > Open Log File. The log rotates at 2 MB (`LOG_MAX_BYTES`) with 1 backup.

The root level is WARNING, but the loggers in `config.INFO_LOGGERS` (`falcon_sender`, `camera_capture`, `ssc.machine`, `ssc.startup`) write INFO: until v2.86 every INFO line `falcon_sender` had (stream start/end, wake-up, `$X`, the M-code trace) was silently dropped because only the root level was set. Things that repeat on a timer must not log per tick: the calibration preview logs a 5 s summary, and `falcon_sender._log_throttled` holds a repeating connect message to one line per 30 s (the calibration MPos poll connects every 0.8 s). At startup `config.log_runtime_versions` writes one line with Python, OS, Tk and library versions (already-imported modules report their own `__version__`; others come from metadata without importing them), and `camera_capture` logs the OpenCV it loads. Help > Open Log File shows the file *selected* in Explorer/Finder (`reveal_in_file_manager_command`) so it can be attached to an email next to `last_calibration_frame.png` and `last_frame_cut.gcode`.

**Swallowed exceptions** (audited 2026-11-10): catch the narrowest type meant; no bare `except:`. A silent `except Exception: pass` is allowed only where `tools/test_exception_handling.py` lists it by file and function with its count; add a deliberate new one there with its reason. To make a swallowed failure visible, log one line through the module's logger and change nothing else. A handler on a timer logs its first failure with the traceback, then counts. The audit: `docs/decisions/testing-and-ci.md`.

## macOS Build

`build.py` patches Info.plist after PyInstaller to add `NSMicrophoneUsageDescription` and `NSCameraUsageDescription`. Without those keys, macOS silently denies mic access to the tuner and toner (and camera access to the machine-integration features).

Patching the plist breaks the bundle's ad-hoc signature, and macOS then silently refuses the permission prompt, so `_resign_macos_app()` re-signs ad hoc (`codesign --force --deep --sign -`) and verifies with `--strict`. Keep the re-sign as the last step that touches the bundle. Package with `ditto -c -k --keepParent`, never `zip -r` (the `--dmg` path is fine). History: `docs/decisions/build-and-release.md`.

## Welcome window and Quick Start Guide

`ui_dialogs.WelcomeWindow` shows once per version: `main.maybe_show_welcome()` compares `settings["welcomed_version"]` with `APP_VERSION` and records the version when the window is shown. Help > Welcome opens it any time. Rules:

- `maybe_show_welcome` is called from exactly one place: the `__main__` block of main.py, under `if not tour:`, on an `after(600)`. Never from `__init__`: `--selftest`, `--tour` and every suite build the app.
- It is plain on purpose: the ordinary dialog look, no canvas, timers, colours of its own or grab. Do not style it. Matt: "let's have the welcome screen be simple. if I have a detailed ui request I will give it." Content: name and version, "Thank you for installing.", a link to search the User Guide, a link to the Quick Start Guide, Close.
- `QuickStartWindow` (Help > Quick Start Guide) is one small page; each line is its own string in `QuickStartWindow.lines`; the Frame & Cut line only when `machine_supported()`. It never opens by itself.

Tests: `test_welcome`. History: `docs/decisions/user-interface.md`.

## Feature Set

File > Feature Set lets users show/hide tabs. It only lists what the build can run: no Tuner/Toner on a Mac without the audio stack (`main.audio_supported`), no Machine menu toggle without pyserial (`main.machine_supported`) — i.e. the Intel Mac build. Saving merges into `visible_tabs` rather than replacing it, so a tab the build doesn't offer keeps its saved value. As of v2.0 the Tuner is default-on (no longer marked experimental), and since 2026-11-10 so is the Analyzer (the `visible_tabs` key stays `"Toner"`; it is listed as Analyzer in the main list, and the Experimental heading is shown only with the Machine menu tick). The terms dialog (`toner_tab._toner_show_harmonics_terms`, titled experimental) gates the tab's HARMONICS switch, not the tab: the first "on" asks, a decline puts the switch back, an accept stores `toner_unlocked` and is never asked again. `config.migrate_analyzer_tab` runs once per profile at startup (`analyzer_tab_shown_v3`), in memory (the flag goes to disk with the next regular save: building the app, `--selftest` and `--tour` write nothing): a tab the old startup rule had forced hidden for a user who never unlocked comes back; one hidden on purpose by an unlocked user stays hidden. The `visible_tabs` dict in settings controls which tabs are shown. **The window resizes itself for the Analyzer** (Matt, 2026-11-10: "have it resize itself"): the start-up width is the old rule's (tab labels and the Tooling tab; `_fit_window_to_tabs` never counts the Analyzer's strip). Selecting the tab widens a narrower window to the strip's need (`_widen_for_analyzer`: `_analyzer_strip_need_px`, never past the screen, height and position unchanged, a maximized or already wide window left alone) and leaving it puts the width back (`_unwiden_after_analyzer`), but only while the window is still exactly as wide as the app made it: a size the user chose since is theirs. The restore is a geometry call checked after idle and asked again while the window still reads the widened width (`_restore_width`, up to 8 times 100 ms apart: Aqua applies a geometry late or drops one); what happened is in `_analyzer_restore`. **When even the screen cannot take one row of the strip** (German on a 1024 px Mac, 2026-11-11), the strip goes to two rows first (`toner_tab._toner_strip_two_rows`: the right column, Load / MODE / Capture, under the switches and the sax selector) and the window is widened to what two rows need; the decision is by the ONE-row need, whatever the strip stands at, and the real screen gets one row again. Capture is never the thing that falls off the page. Both run from `on_tab_changed`; every geometry change in the tests is awaited with a bounded poll, never a fixed wait, and a check reads layout state only after an idle pass.

**Machine Integration (v2.5+, experimental, opt-in)** — File > Feature Set also exposes an `experimental_machine_menu` toggle (default False). When on, the Pad Maker tab gains an Options > Machine submenu, a Frame & Cut button, and the polygon-draw dialog grows its "Get from camera" + live overlay UI. When off, none of that surfaces — the app behaves like pre-v2.5. Helper `_machine_enabled()` (machine_control.py) is the single source of truth and is called by `_refresh_machine_ui_state()` whenever the toggle changes or the calibration file appears/disappears.

**Gating: the laser and the camera are gated separately.** Frame & Cut never requires a camera calibration, on either tab. `_refresh_machine_ui_state()` is the one place the rule is enforced:

- *Toggle off* (or no pyserial in the build, `machine_supported()`): nothing machine-related shows; the Machine menu is not built at all.
- *Toggle on*: the Machine menu's laser items are enabled (Home Laser, Test Connection, Clear Errors, Reset Falcon, Framing Power: `machine_control.MACHINE_MENU_LASER_ITEMS`), and Camera Calibration as always. Once a laser is detected (`falcon_port`), the Frame & Cut button shows on the Pad Maker and on every Tooling section that writes G-code. One rule for both tabs, no calibration. `falcon_port` is set and cleared by the laser watch (CLAUDE-engines.md "The laser watch"), on the Tk thread, for the whole session: the button appears within seconds of the laser being plugged in after launch and goes again when its port is no longer listed (which also clears `_falcon_homed_this_session`). Before 2026-11-09 the app probed the ports until a laser answered and never looked again.
- *Toggle on and a camera calibration on disk*: everything that uses the camera. Camera-Polygon Inset Margin in the menu (`MACHINE_MENU_CAMERA_ITEMS`), the "Live camera preview" tick, and through `_camera_capture_ready()` (the toggle and `_camera_calibration_present()`) "Get from camera", the live overlay in the shape dialog, Try Auto Locate and "Re-capture from camera".

With no calibration a Frame & Cut run opens no camera and has no Auto Locate target. The toggle is restart-required; buttons and items follow detection and calibration live. History: `docs/decisions/machine-integration.md`.

## v2.5 Machine Integration Architecture

A condensed map of the new (v2.5) cross-cutting subsystem since it touches several modules:

```
File > Feature Set toggle (experimental_machine_menu)
  ↓ (when on)
MachineControlMixin._machine_enabled() + _refresh_machine_ui_state()   (machine_control.py)
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
      └─ MachineControlMixin.on_frame_and_cut(tooling_job=...) — the same routine from the
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
- Every camera window opens its camera through `camera_capture.open_camera_or_fallback(index, saved_index)`: when the saved number (`camera_index_override`) fails, the Falcon-named camera is tried once — and the last other camera only while no calibration is on disk (a calibration belongs to one camera; the webcam under the Falcon's homography is a believable, wrong outline) — the window runs on what opened and says so in its status line (`camera_capture.fallback_notice`), one warning is logged, and the saved number is not changed (only a calibration save or Switch camera changes it, as before). `main._release_live_camera()` closes the app's own live preview before the shape dialog or Camera Calibration opens a camera, so the app never holds the Falcon camera against itself.
- Frame G-code emits `G92 X{lb_x} Y{lb_y}` (LB-vertex in Y-flipped frame) so jogging to the visible scrap corner bridges to the polygon's bbox origin.
- A job that was not nested on the loaded shape (card on "Fit card to paper", and every Tooling job) takes nothing from it: `on_shape` in `on_frame_and_cut` is False, so no outline is traced, the G92 is `X0 Y0`, the cut is not flipped with the shape's height and Try Auto Locate has no target.
- Everything after a successful `sender.connect()` runs inside the `try` whose `finally` disconnects, and the sleep lock is taken after the last early return, so no way out leaves the port open or the lock held (`tools/test_release_3_0.py`).
- `falcon_sender._stream_loop` starts every stream with two safety lines BEFORE the caller's G-code: (1) `G4 P0.01` wake-up dwell to prevent Grbl parser dormancy from eating the first command, and (2) `$X` unlock so a sticky Alarm from the previous stream (e.g. soft-limit triggered by a jog-shifted cut bbox in Frame & Cut) doesn't force a Falcon power cycle. `$X` is a no-op when Grbl is already Idle.
