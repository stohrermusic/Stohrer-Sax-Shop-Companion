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

### Testing

Test suites live in `tools/` and run non-interactively (no GUI). Run individual suites directly:

```bash
python tools/test_toner_engine.py
python tools/test_toner_full.py
python tools/test_tuner_engine.py
python tools/test_bugfixes.py
python tools/test_config.py
```

All test suites (31 files): `test_audio_utils`, `test_autofit_shift`, `test_bugfixes`, `test_compare_filters`, `test_concert_pitch`, `test_config`, `test_dart_ranges`, `test_dart_shapes`, `test_descriptor_validity`, `test_detection_fix`, `test_edge_bias`, `test_fingerprint_filtering`, `test_goodson_import`, `test_gpu_tuner`, `test_pad_notes`, `test_pad_preview`, `test_release_1_9`, `test_sizing_presets_workflow`, `test_sizing_ranges`, `test_smoke_ui`, `test_tooling`, `test_toner_display`, `test_toner_engine`, `test_toner_full`, `test_tooltips`, `test_tuner_engine`, `test_tuner_updates`, `test_v161_compat`, `test_wav_import`, `test_wav_recording`, `test_web_pad_import`.

**Portability notes**:
- `test_descriptor_validity` hardcodes a local WAV corpus path (`C:\sax shop companion\recordings`) and only runs on Matt's workstation. Skip it in CI and clean checkouts.
- `test_smoke_ui` constructs the full `PadSVGGeneratorApp` in a withdrawn Tk root — requires a display, so works on Windows/macOS and GitHub Actions Windows runners. On headless Linux it self-skips with a "no display" message.

Before committing, run the suites affected by your changes. For releases, run all (minus `test_descriptor_validity` unless the WAV corpus is available). If adding new functionality, write a test script in `tools/` that exercises affected code paths. Test engine/logic functions directly. Print PASS/FAIL per test with a summary.

### Linting

```bash
pip install ruff
ruff check .          # zero errors expected on beta/main
ruff check --fix .    # auto-fix safe issues (unused imports, empty f-strings, etc.)
```

Config lives in `ruff.toml` (py311 target, 120-char lines, default E+F rules with E501/E701/E731/E741 relaxed, tools/ exempts E402 for `sys.path.insert` patterns). The CI `lint` job runs `ruff check .` and fails the workflow on any violation — keep the tree clean.

**Ruff gotcha for test scripts**: `ruff --fix` will strip "unused" imports even when they're the whole point (e.g. a test that verifies names import cleanly). Reference the imported names afterward (`assert all([Class1, Class2, ...])`) so ruff sees them as used. See `tools/test_smoke_ui.py` for the pattern.

## Pad Preview Window

The Sizing Rules dialog has an opt-in live preview (`PadPreviewWindow` in `ui_dialogs.py`). A "Show live pad preview" checkbox just below the preset section opens a resizable Toplevel that renders the selected pad with the parent form's current sizing rules applied. Controls: pad size (mm), per-material checkboxes (leather / felt / card / exact size), layout radio (layered concentric vs side-by-side).

Geometry comes from the same helpers as the SVG output: `svg_engine.get_disc_diameter`, `svg_engine.get_felt_thickness_mm`, and `svg_engine._wave_value` (for the dart wave). That means what you see in the preview is the exact shape that will be cut. Drawing happens on a `tk.Canvas` — fast, dependency-free, redraws on `<Configure>`.

Live updates: the window polls `parent_options._capture_form_to_dict()` every 200 ms and re-renders if the snapshot changed. No tk-var traces are used because some form state (sizing/dart range lists) lives in plain Python lists that aren't trace-able. Polling tolerates mid-edit invalid states (catches `TclError` / `ValueError`) and falls back to a placeholder message.

The preview tears down whenever the OptionsWindow is destroyed (any path) via a `<Destroy>` bind on `self.top`. Closing the preview directly resets the parent's `show_preview_var` so the checkbox stays in sync.

## Sizing Rules Presets Workflow

The Sizing Rules dialog (`OptionsWindow`) is preset-first — the preset section sits at the *top* of the form, and the bottom button is **Apply** (not "Save"). Workflow:

- **Preset dropdown + Load** at the top: explicit Load click, with a "discard unsaved changes?" confirm if the form is dirty.
- **Save Preset** opens `SaveSizingPresetDialog`, a small radio-choice modal: *Overwrite existing* (combobox of saved names) or *Save as new preset* (text entry). Defaults to overwrite when an active preset is loaded, to "new" when nothing is loaded or the library is empty.
- **Rename** prompts for a new name; refuses empty / duplicate names.
- **Delete** refuses to wipe the last preset — at least one must remain.
- **Apply** (bottom button): if the form is dirty, prompts the user to either save as a preset first or back out to keep editing — there is no path to commit unsaved-as-a-preset changes silently.
- **Cancel / window-close X**: if dirty, three-way prompt — save as preset / discard / keep editing.

Dirty detection is a `_capture_form_to_dict()` snapshot vs `self._baseline`; the baseline resets on dialog open, after Load, and after a successful Save Preset. `active_preset_name` tracks which preset's values currently sit in the form (used for the Save Preset overwrite default).

**Bootstrap**: `main.py` auto-creates a `Default` preset from current settings on first run if `sizing_presets` is empty (via `config.settings_to_sizing_preset`). The app guarantees at least one preset always exists.

## Settings-Dialog Tooltips

`ui_dialogs.py` exposes a `Tooltip` helper plus `add_tooltip(widget, text)` and `add_tooltips(text, *widgets)` convenience functions. All settings dialogs (Sizing Rules, G-code Settings, Layer Colors, Key Layout, Tuner Settings, Toner Settings) hover-explain their fields. When adding a new setting widget, attach a tooltip alongside it — short sentence, plain English, focused on *why* the user would change it. Attach to both the label and the input so users can hover either. The helper popup is overrideredirect + topmost so it appears above modal dialogs without stealing focus.

## Dart Shape Spectrum

The dart wave shape (formerly "Star/Dart"; now just "Darts" in the UI) is a single 0.0–1.0 slider that smoothly interpolates between three primitive shapes:
- 0.0 = Triangle (linear ramps between peaks/valleys)
- 0.5 = Sine (raw cosine)
- 1.0 = Square (saturating sign function via `|c|^0.05`)

Math lives in `_wave_value` in `svg_engine.py`; `calculate_star_path` calls it per sample. The default value is now `0.5` (sine). Legacy configs used a 0.0=sine, 1.0=square scale; `load_settings` migrates them once via `0.5 + 0.5 * old` and sets `dart_shape_v2: True` so the migration doesn't run again. `dart_ranges[*].shape_factor` is migrated alongside the universal value. When changing dart-shape behavior, update `tools/test_dart_shapes.py` (anchor + smoothness + migration coverage).

Internal variable names retain the `dart_` prefix (`dart_shape_factor`, `dart_threshold`, etc.); only the user-facing labels were renamed to "Darts".

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
#   Windows: dist/SaxShopCompanion.exe
#   macOS:   dist/SaxShopCompanion.app (or .dmg with --dmg flag)
#   Linux:   dist/SaxShopCompanion
```

### GPU Tuner Renderer (Rust/wgpu)

The strobe tuner has an optional GPU-accelerated renderer in `tuner_renderer/` (Rust crate using pyo3 + wgpu). CI builds this on full-build platforms, but a local checkout of `python main.py` will silently fall back to the slower canvas renderer unless you build the extension yourself:

```bash
# Requires Rust toolchain (rustup)
pip install maturin
python -m maturin build --release --manifest-path tuner_renderer/Cargo.toml
pip install --find-links tuner_renderer/target/wheels tuner_render
```

After this, `import tuner_render` succeeds and the tuner uses GPU rendering at 60-120 fps. If the import fails for any reason the tuner transparently falls back to canvas rendering — the app still runs, just slower. See the `_skip_theme` / `_dark_canvas` flag notes in the Strobe Tuner architecture section for integration details, and the GPU/Canvas constant alignment warning (`DIM_MULTIPLIER`, `BRIGHTNESS_GAMMA`) for what has to stay in sync between the Python path and the shader.

On Windows, Rust needs the MSVC linker. The one-time machine setup is `winget install Rustlang.Rustup` followed by `winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"`.

### Windows Installer (Inno Setup)

CI wraps `dist\SaxShopCompanion.exe` into a versioned `SaxShopCompanion-Windows-Setup-{version}.exe` via `installer.iss`. The installer creates Start Menu + optional desktop shortcuts, registers an uninstaller, and installs to `{autopf}\SaxShopCompanion` (requires admin UAC). User data in `%APPDATA%\StohrerSaxShopCompanion\` is untouched on uninstall. Local build:

```bash
python build.py
iscc /DAppVersion=2.0.1 installer.iss   # requires Inno Setup 6
```

**Do not change the `AppId` GUID** in `installer.iss` — Windows uses it to recognize upgrades. Changing it produces a parallel install instead of an in-place upgrade.

## Branching Strategy

- **`main`**: Stable release branch. Merges from `beta` when features are tested and ready.
- **`beta`**: Active development branch. New features land here first (e.g. filled engraving, air assist toggles, cut grouping). Always work on `beta` unless told otherwise.
- CI builds trigger on push to `main`, `beta`, or `gamma`.

## Versioning

`APP_VERSION` in `config.py` is the manual source of truth — bump it when preparing a release. `APP_BUILD_DATE` is auto-derived from `sys.executable`'s mtime in frozen builds (installer/zip copies preserve mtime) and falls back to the manual constant when running from source. The About dialog reads both via `ui_dialogs.py`.

## CI/CD (GitHub Actions)

The `.github/workflows/build.yml` workflow has two jobs:
- **`lint`** (ubuntu-latest, ~10s): runs `ruff check .` — fails the workflow on any violation
- **`build`** (4-platform matrix): Windows Inno Setup installer (the bare PyInstaller .exe is built but not published — only the installer ships), macOS Apple Silicon .app, macOS Intel .app, and Linux binary

Triggers on push to `main`, `beta`, or `gamma`, on release creation, or manually.

- macOS Intel build (`macos-15-intel` runner) installs only svgwrite+pyinstaller (no numpy/sounddevice) — tuner and toner are unavailable
- `full_build: true/false` matrix flag controls whether Rust toolchain + maturin are installed for the GPU tuner renderer
- The Windows job also installs Inno Setup 6 via Chocolatey and builds the installer; version is extracted from `config.py`'s `APP_VERSION` via PowerShell regex
- Uploads artifacts to the workflow run; on Windows that's the installer only (`SaxShopCompanion-Windows-Setup-*.exe`). All published artifacts also attach to GitHub Releases on release events.

**Action pins**: `actions/checkout@v5`, `actions/setup-python@v6`, `actions/upload-artifact@v6` — all on Node 24. Don't downgrade; GitHub removes Node 20 from runners in September 2026.

## Config File Locations

The app stores settings and presets in platform-appropriate locations:

| Platform | Location |
|----------|----------|
| Windows | `%APPDATA%\StohrerSaxShopCompanion\` |
| macOS | `~/Library/Application Support/StohrerSaxShopCompanion/` |
| Linux | `~/.config/StohrerSaxShopCompanion/` (respects `XDG_CONFIG_HOME`) |

**Backward compatibility**: On first run, existing config files in the old location (current working directory) are automatically migrated to the new location.

**Manual import**: Users can also manually import settings from a previous installation via File > "Import Settings from Folder..." which copies config files from a selected directory.
