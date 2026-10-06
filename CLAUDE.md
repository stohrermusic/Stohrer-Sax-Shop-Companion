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

Test suites live in `tools/` and run non-interactively (no GUI). **Run everything** with `python tools/run_tests.py` (one line per suite, failures' output at the end, non-zero exit on any failure; `python tools/run_tests.py zone svg` filters by name). That is what CI runs. Run individual suites directly:

```bash
python tools/test_toner_engine.py
python tools/test_toner_full.py
python tools/test_tuner_engine.py
python tools/test_bugfixes.py
python tools/test_config.py
```

All test suites (59 files): `test_audio_utils`, `test_autofit_shift`, `test_bugfixes`, `test_camera_capture`, `test_card_paper_size`, `test_compare_filters`, `test_concert_pitch`, `test_config`, `test_dart_ranges`, `test_dart_shapes`, `test_descriptor_validity`, `test_detection_fix`, `test_edge_bias`, `test_engine_parity`, `test_falcon_sender`, `test_feeds_speeds_tester`, `test_fingerprint_filtering`, `test_frame_cut_alignment`, `test_frame_cut_scrap`, `test_framing_power`, `test_gcode_passes`, `test_gcode_presets_workflow`, `test_gettext_shadowing`, `test_goodson_import`, `test_gpu_tuner`, `test_i18n`, `test_job_history`, `test_lid_confirm`, `test_library_tabs_e2e`, `test_locator_marks`, `test_mac_paths`, `test_nesting_parity`, `test_pad_maker_e2e`, `test_pad_notes`, `test_pad_preview`, `test_polygon_parity`, `test_release_1_9`, `test_scrap_live_list`, `test_scrap_partial_nest`, `test_serial_lookup`, `test_sizing_presets_workflow`, `test_sizing_ranges`, `test_smoke_ui`, `test_svg_wellformed`, `test_tooling`, `test_tooling_e2e`, `test_toner_display`, `test_toner_engine`, `test_toner_full`, `test_toner_live`, `test_tooltips`, `test_tuner_canvas`, `test_tuner_engine`, `test_tuner_updates`, `test_v161_compat`, `test_wav_import`, `test_wav_recording`, `test_web_pad_import`, `test_zone_labels`.

**SVG↔G-code parity**: `test_engine_parity` pins the contract that the SVG/preview output and the G-code output describe the same physical object (dart wave shape, engraving label placement, engine purity). The two engines render independently and have drifted before — when touching shared geometry (wave math, placement formulas, Y-flip), run this suite and extend it for any new shared shape.

**What CI runs (since 2026-10-06).** Three jobs on every push to `beta` or `main`: `lint` (ruff), `test` (every suite via `run_tests.py` on **Windows, macOS and Linux** — Windows and macOS runners have a display so the GUI suites run for real; Linux runs under `xvfb-run`; full `requirements.txt` is installed so the OpenCV and pyserial suites run instead of self-skipping), and `build`, which `needs: [lint, test]` — nothing is built or attached to a release unless every test job passed (added 2026-10-06 after the first release cut showed the test job ran *beside* the build, not before it). Before this, CI ran no tests at all, and the macOS code paths (`IS_MACOS` theming, `::tk::mac::Quit`, canvas-only tuner, the no-numpy Intel build) had never been executed anywhere — nobody on the project owns a Mac. The macOS test job is the only place they run. The repo is public, so Actions minutes are free; that is why the test job runs on every push rather than only on `main`.

**The Intel Mac is tested as it ships.** The `test` matrix includes `macos-15-intel` with `svgwrite numpy pillow babel` installed — numpy but no sounddevice, exactly the Intel build's dependency set — running `run_tests.py --allow-missing sounddevice`, which counts a suite that dies on one of those imports as skipped (reported, not hidden). Everything else runs for real on x86 macOS Tk. Until 2026-10-06 that build had no numpy, and the first Intel run measured the pure-Python nest at **774 s (`test_zone_labels`), 293 s and 260 s (the two scrap suites)** — the wait an Intel user was living with on a big fill; Matt's answer was to ship numpy in the Intel build ("include numpy for them"), after which `test_zone_labels` took 7.7 s there. The `_HAS_NUMPY` fallback now runs only in the parity suites' reference implementations and in source checkouts without numpy; no shipped build lacks it. The audio-tab suites (`test_tuner_canvas`, `test_toner_live`) self-skip when `AUDIO_AVAILABLE` is False, because that build's tuner and toner are the "not available on this Mac" panels by design.

**What the new coverage found on its first runs (2026-10-06 evening):** the no-mic tuner stop showed dark wheels and *no message* — `_tuner_build_wheels_canvas` starts with `canvas.delete("all")`, which wiped the "Audio error" text on the next resize, so a user with no (or a denied) microphone saw a dead tuner with no explanation; the toner's `_toner_build_spectrum_bars` did the same. Both tabs now keep an `_*_error_state` and redraw it after a rebuild; gated in `test_tuner_canvas` (pass 0) and `test_toner_live`. The German tour showed the pad preview's material names and legend hard-coded in English (`PadPreviewWindow.LABELS`, now `_()`-wrapped). And the "dark" Mac tour came back light: `defaults write -g AppleInterfaceStyle Dark` does not change a running login session, so the tour now forces its windows dark through Tk (`::tk::unsupported::MacWindowStyle appearance … darkaqua`, `--appearance dark`) and writes the *measured* mean brightness into `tour-done.txt`, so the artifact says which mode it really is. One Linux build job failed at "Set up job" with no log — runner infrastructure, not code; a re-run is the fix for that one. Second-round findings from the same pictures: the Mac runner *has* a silent input device, so the "tuner-no-mic" stop shows a running tuner with the pilot lit rather than the error overlay — the no-mic error itself is gated in `test_tuner_canvas` pass 0, the stop stays as "what you see before you play"; with the per-window appearance switch the main window renders dark but dialogs stay light (each Toplevel has its own appearance), so the CI dark step also tries the session-wide `osascript` appearance switch — check `tour-done.txt`'s measured brightness and the dialog shots to see which took; and the library dropdowns' "All Libraries" was untranslated because the code compares against the literal — every use is now `_("All Libraries")`, which keeps the comparisons consistent at runtime, and a stored English sentinel from an older config simply falls back to the all-libraries default.

**Mac blind spots that remain (2026-10-06), none reachable from a runner:** the microphone/camera *permission prompt* itself (the synthetic tone bypasses the mic; the signature/plist guards cover the mechanism that broke before); Retina scaling (runners are 1× displays); native file dialogs (every test stubs them, the tour never opens one); Gatekeeper's first-launch path (CI runs the binary directly, no quarantine attribute). A report in any of those areas has no automated coverage behind it.

**Windows and Linux coverage added 2026-10-06, same evening:** the Windows build job **installs the Inno Setup installer silently, runs the installed copy's `--selftest`, checks the Start Menu shortcut, uninstalls silently, and asserts the user's `%APPDATA%` config folder survived** — nothing had ever run the installer before. Both GPU-building jobs (Windows, Linux) run `test_gpu_tuner` and `test_tuner_canvas` with the freshly built wheel: runners have only a software adapter (Windows) or no Vulkan (Linux), so what that exercises is the *fallback* — `test_tuner_canvas` accepts either a live GPU renderer or the designed drop to canvas with the CPU-mode notice. The Windows build job also runs the screenshot tour **in German** (longest language) as `win-tour-de`, so clipped translated labels can be seen. **Blind spots that remain, none reachable from a runner:** display scaling above 100 % (Matt's machines run 125–250 %; runners are 100 %); SD-card eject (removable drive); Falcon serial and camera hardware (mock and synthetic only); Wayland sessions (runners are X11); real audio devices on any platform.

**Frozen-build probe.** `SaxShopCompanion --selftest` (`_selftest()` in main.py) constructs the whole app in a withdrawn root — every tab, every import — and exits 0 or 1 without saving anything or showing a dialog (the excepthook is replaced for the run). The `build` job runs it against the PyInstaller output on each platform right after building: the one check that runs the *shipped bundle* rather than the source tree, so a missing hidden import, a bad `--add-data` path, or a module bundled on the wrong platform (the macOS GPU renderer class of failure) fails the build instead of a user's first launch. The .exe is a windowed app, so the Windows step waits on it with `Start-Process -Wait` and reads `ExitCode`.

**Eyes on the Mac.** The macOS build jobs launch the built .app for real, `screencapture` the runner's screen after 12 s, and upload it as the `mac-screenshot-*` artifact — two pictures: the Pad Maker, and a second launch with `--tour tuner`, which opens the Tuner tab on a **synthetic 440 Hz tone** instead of the microphone (`TunerEngine.synthetic_hz`; `start()` opens no stream and `analyze()` feeds the tone into the ring buffer itself), so the canvas strobe and VU readout can be seen running on real macOS Tk. `tools/test_tuner_canvas.py` is the automated form of the same thing: it forces canvas mode, drives the tab with the synthetic tone through the real `_tuner_animate` loop, and requires the A wheel lit and the others dark, the readout to say A within 1 cent, and no error overlay; where the GPU wheel is built it runs the GPU path too and requires zero render failures. On the macOS CI runner that suite is the tuner's only end-to-end check. The toner has the same pair: a third Mac launch with `--tour toner` (Bb3 with six harmonics via `TonerEngine.synthetic_hz`) screenshots the analyzer, and `tools/test_toner_live.py` drives the Toner tab the same way and requires the note and frequency readouts, the fundamental within 0.5 c, and at least five harmonics. The toner is hidden behind the beta terms, so both use **`SAXSHOP_CONFIG_DIR`** — an environment override of the config folder (`config.get_config_dir()`), read before `config` is imported — to hand the app a prepared profile with the toner unlocked; `open` on macOS drops the environment, so the CI step runs the bundle's executable directly for that launch. `SAXSHOP_CONFIG_DIR` is also how any test should isolate itself from the user's real settings. The tab tests drive frames with a real `mainloop` and a quit timer, never a `root.update()` loop: the canvas tuner's frame outlasts its 16 ms interval on a large window, so `update()` keeps servicing the already-due reschedule and never returns (found 2026-10-06, 25 s stall in `canvas.coords`). Their `check()` helper fails on a returned `False` as well as on an exception — the first draft only caught exceptions, and two checks passed against an engine that had no synthetic source at all. `continue-on-error`, so it never fails a build — it is a picture for Matt to look at, the only way the Mac UI gets seen. A first-run dialog or a theme problem shows up there.

**End-to-end through the real buttons (2026-10-06).** `test_pad_maker_e2e` types a pad list into the live form, picks all four materials with zones and locator marks on, clicks Generate SVG and Generate G-code with only the folder dialog and message boxes stubbed, and checks the *output*: one file per material (G-code skips `exact_size` by design — it has no laser settings), every SVG parses with the sheet size on its root, the leather SVG and G-code carry the locator layer (engraved before any cut), every G-code move lies on the sheet and ends with the laser off, and the job history recorded both runs. `test_tooling_e2e` drives every Tooling handler the same way (dies, holders in all three variants including the too-small-sheet refusal, organizer and spacer copies byte-identical to the bundled assets, kerf tests, speed & power with its `<name>_legend.txt`). `test_library_tabs_e2e` saves, reloads and deletes a key-height set and a screw spec, on disk too, and checks the serial lookup label against `lookup_serial_year`. `test_mac_paths` pins the macOS-specific paths (quit command, system colours, GPU never on darwin, config folder) and runs the whole `--tour all` as a dialog-construction gate with screenshots off. All four use `SAXSHOP_CONFIG_DIR` and never touch the user's data. The first run of these found the three gettext-shadowing bugs above and the Aqua tab-label clipping (`_fit_window_to_tabs`) and, from the first Mac tour screenshots, the Sizing Rules dialog clipping its preset bar at a fixed 500 px (`OptionsWindow._fit_dialog_width` now widens to the content; gated in `test_mac_paths`).

**The screenshot tour.** `SaxShopCompanion --tour all --shots DIR` walks every tab and dialog (18 stops: Pad Maker, Sizing Rules with the lesser-used section open and the pad preview up, Layer Colors, G-code Settings, nesting preview, polygon draw, Job History, Feature Set, User Guide, About, the four library/tooling tabs, Key Layout, the tuner with no microphone (the audio-error overlay a user with a denied mic sees), then tuner and toner on synthetic tones), screenshots each cropped to the app's own windows, writes `tour-done.txt` (with any step errors), and exits. The macOS build jobs run it on the frozen .app twice — light mode, then with the runner switched to Dark Mode (`defaults write -g AppleInterfaceStyle Dark`) — and upload both as `mac-tour-*`; those pictures are how the Mac UI gets reviewed. `run_tour(root, app, shots_dir=None, ...)` in main.py is the same walk as a gate (`test_mac_paths`). Modal stops (nesting preview, polygon draw, job history, about) work because each stop's finish is scheduled with `after()` *before* its `open()` is called — Tk keeps firing timers inside `wait_window`.

**Gates added 2026-10-06, each measured first:** `test_svg_wellformed` parses every Pad Maker SVG variant with ElementTree and checks the root (nothing had ever parsed a generated SVG; the parity suites regex-match). `test_i18n` now checks **all four** catalogs for empty and fuzzy entries and a stale `.mo` (it checked Spanish only; the other three could ship half-translated unnoticed). `test_tooltips` walks every settings dialog and asserts each input widget has a tooltip (zero were missing when the gate landed). `test_gpu_tuner` **skips** instead of failing when `tuner_render` isn't built — a fresh clone and the CI test runners don't have the wheel, and macOS must never have it; set `SSC_REQUIRE_GPU=1` on a machine that built it to make a broken wheel fail loudly.

**Coverage snapshot, 2026-10-06** (`coverage run` over every suite): engines are well covered — svg 96 %, gcode 84 %, toner engine 77 %, tuner engine 73 %, config 77 %; UI modules are thin — toner tab 11 %, tooling tab 32 %, ui_dialogs 37 %, main 38 %, tuner tab 38 %; camera 12 % and falcon 9 % only because OpenCV and pyserial weren't installed on the measuring machine. When adding a UI feature, the smoke-style construct-and-poke test (see `test_locator_marks` GUI cases) is how that number moves.

**Portability notes**:
- `test_descriptor_validity` hardcodes a local WAV corpus path (`C:\sax shop companion\recordings`) and only runs on Matt's workstation; `run_tests.py` skips it when the corpus is absent. `test_goodson_import` reads the website checkout at `C:/code/stohrermusic/...` and self-skips elsewhere.
- `test_tuner_engine`'s reference-player start tests and `test_wav_recording`'s "Music or Documents exists" check self-skip on a machine with no audio output device / a bare home folder (CI runners). `play()` returning False there is the fallback working.
- **The first CI test run (2026-10-06) caught a shipped bug:** CI installs `opencv-python` 5.0, which returns ChArUco ids as a flat `(N,)` array where 4.x returned `(N, 1)`; `calibrate_from_frames` indexed `i[0]` on a scalar and raised `IndexError`. The Windows installer built that day bundled OpenCV 5, so camera calibration in the shipped build would have crashed at the save step. Fixed by flattening the ids (`np.asarray(ids).reshape(-1)`); `requirements.txt` keeps `>=4.7`. Nobody had run `test_camera_capture` with OpenCV installed before — on the dev machine it had always self-skipped. Its calibration fixture was also degenerate (eight fronto-parallel translations of one card cannot determine a focal length; OpenCV 5 diverged to fx ≈ 9 × 10¹⁷) and is now ten synthetic pinhole views with a known K, which the test requires `calibrate_from_frames` to recover (fx within 2 %, principal point within 3 px; measured 800.2 of 800, rms 0.04 px).
- `test_smoke_ui` constructs the full `PadSVGGeneratorApp` in a withdrawn Tk root — requires a display, so works on Windows/macOS and GitHub Actions Windows runners. On headless Linux it self-skips with a "no display" message.
- `test_zone_labels` is headless for everything except its three `preview_*` cases, which build a real `NestingPreviewWindow` and inspect its canvas. Those self-skip without a display; the other 31 always run. Note the window calls `wait_window()` in `__init__`, so the canvas inspection must be scheduled on the parent via `after()` *before* constructing it — see `_probe_preview`.

Before committing, run the suites affected by your changes. For releases, run all (minus `test_descriptor_validity` unless the WAV corpus is available). If adding new functionality, write a test script in `tools/` that exercises affected code paths. Test engine/logic functions directly. Print PASS/FAIL per test with a summary.

### Linting

```bash
pip install ruff
ruff check .          # zero errors expected on beta/main
ruff check --fix .    # auto-fix safe issues (unused imports, empty f-strings, etc.)
```

Config lives in `ruff.toml` (py311 target, 120-char lines, default E+F rules with E501/E701/E731/E741 relaxed, tools/ exempts E402 for `sys.path.insert` patterns). The CI `lint` job runs `ruff check .` and fails the workflow on any violation — keep the tree clean.

**Ruff gotcha for test scripts**: `ruff --fix` will strip "unused" imports even when they're the whole point (e.g. a test that verifies names import cleanly). Reference the imported names afterward (`assert all([Class1, Class2, ...])`) so ruff sees them as used. See `tools/test_smoke_ui.py` for the pattern.

## Job History

File > Job History (Pad Maker) opens `JobHistoryWindow` (`ui_dialogs.py`) — a log of every batch that reached an output stage. Storage is `job_history.json` via `load_job_history()` / `save_job_history()` / `append_job_history()` in config.py (newest first, trimmed to `JOB_HISTORY_LIMIT` = 300).

**Recording**: `PadSVGGeneratorApp._record_job(output, materials, pads, params, ...)` in main.py is called from exactly five places, always *after* real output exists:

| Call site | `output` | Notes |
|---|---|---|
| `on_generate_svg` | `"svg"` | after the per-material write loop |
| `_generate_svg_scrap_mode` | `"svg"` | per scrap, with `scrap_num` |
| `on_generate_gcode` | `"gcode"` | after the write loop, before the working popup closes |
| `_generate_gcode_scrap_mode` | `"gcode"` | per scrap, with `scrap_num` |
| `on_frame_and_cut` | `"laser"` | after the cut dialog; `status` carries `_final_reason`, so stopped/errored runs are logged too |

`_record_job` swallows every exception and logs it. Nothing in the app reads the history back except the dialog, so a history failure must never surface as a generation error — keep it that way when adding call sites.

**Reload**: `_load_job_into_form(job)` restores only what the user typed — pad text, materials, sheet size, center hole, base filename. It deliberately does NOT restore sizing rules, G-code settings, or `custom_polygon` (a camera-captured scrap is gone by then; the entry records `polygon_vertices` for display only). Sheet size is stored both as-typed and in mm, so a job saved in inches reloads correctly when the app is now in mm. Loading is refused while a scrap session is active (the session owns the pad list and locks the material checkboxes).

**List columns are measured, not fixed**: `_column_widths()` sizes each column from the widest header/value at refresh time. Translated headers vary a lot ("Pads" is "Zapatillas" in Spanish), and hardcoded widths truncated them. If you add a column, add it to `_headers()` and `_cells()` together — they're zipped positionally.

Tests: `tools/test_job_history.py` (storage round-trip + corruption handling, column alignment including a simulated long-translation case, and a record→reload round trip through a real form).

## Labeled Zones

Options > Sizing Rules > **Lesser-used settings** > **Labeled Zones** (moved into the collapsible 2026-10-06): an opt-in toggle plus an editable pad-size range (default 7.0–12.5mm). Pads in that range are cut in bordered blocks — one block per size, grid-packed, with the size engraved along the block's top edge. Everything outside the range nests normally at full density.

**Why it exists**: small discs are indistinguishable once they're off the laser — a 7.0 and a 7.5 look the same. Their own engraved number doesn't solve it: the font gate (`font_size >= radius * 0.8`) drops the engraving entirely below ~5mm, and above that the number is often unreadable — too small on card/felt, and buried in the darts on leather (every leather pad under `dart_threshold` gets darts). **The label cannot move to the middle of a small leather pad — that's the sealing surface, and those are usually octave pads.** So the label goes on the waste instead of the part, which is the only place it can go.

**Cost**: zones trade sheet area for legibility, hence opt-in and the "may increase material wastage" note in the dialog. Measured on 7–12.5mm pads with a 1mm gutter, rectangular zones run 53–69% full (leather ~68%, near the `π/4 × (d/(d+g))²` ceiling for the gutter). Roughly 10–15% more sheet for a zoned batch.

**One model, two packers.** A group is a compact grid of one size with a rectangle round it and the size engraved on it, and groups are nested **as units** — like oversized pads. `nest_with_zones` dispatches on sheet shape: rectangular sheets shelf-pack groups into a band along the bottom (`_shelf_pack_zones`), traced polygons first-fit them into the outline (`_nest_polygon_groups`). Both emit identical `shape: 'rect'` zone dicts, so all three renderers share one path.

**Grid shape** (`zone_grid_candidates`) scores `aspect + empty_slots`: squareness matters, but a grid with holes doesn't read as a block, so a gap costs about as much as one step of elongation. That gives 6→3×2, 9→3×3, 8→4×2 (exact, not 3×3-with-a-hole), and a prime like 7→4×2 with one gap rather than a row of seven. Callers walk the list in order, so an awkward scrap degrades to a flatter grid instead of refusing the size.

**Rejected alternatives, all measured:**
- *Circular groups* — no new packing code at all (a group is just a big disc to `_nest_discs`), but roughly half as dense (28–36% vs 53–69% fill), and far too big for real material: 8× 7mm **leather** needs a 70mm circle, 48% of a 100×80 scrap; 8× 12.5mm needs 97mm.
- *Sequential per-size placement with an inflated inter-size collision radius* — the cheapest possible "grouping", and it **doesn't group**: the 7mm set still spread 83–108mm across a 150mm scrap (baseline 84.5mm), and widening the gap made it worse.
- *One full-width horizontal band per size, clipped to the outline* — shipped briefly and reverted. Bands are fine while each size is a single row, but once pads are gridded the rows stack: two grids ate a 146mm scrap whole and the remaining sizes had nowhere to go (**15 of 25 pads placed** on a real piece). The same four groups packed as units occupy 207×54mm and all 25 fit. `_clip_polygon_y` survives from that work as a tested general utility — nothing in the app calls it anymore (the leftover region uses cover circles over the full outline instead), only its tests do.

**Key invariants**:
- `placed` keeps its plain `(pad_size, cx, cy, r)` shape. Zones travel as a **separate parallel list**, so `compute_remaining_pads`, the preview window, and job history all work untouched. Don't fold zone data into `placed`.
- Every zone is a `shape: 'rect'` dict (`x/y/w/h` + `cols/rows/label`). A `'poly'` variant existed for the band layout and was removed with it — nothing emitted it, so it was untested dead code. If a future shape (a circular group, say) is added, all three renderers — SVG `_render_svg_zones`, the G-code zone-stroke block, and the preview canvas — must learn it together, plus a parity test.
- Nothing touches the parity-pinned scan functions (`_scan_radial_*`, `_find_best_polygon_*`). Both layouts work by handing the nester a *smaller region*, never by adding obstacles. Scattering zones would leave the free area with *holes*, needing obstacle support in all six scan paths — deliberately not done.
- **A `max`-quantity size can't be grouped, and the app refuses the combination.** A group is a fixed grid, so it needs a fixed count. `_prepare_generation()` in main.py errors out when zones are on and a `max` pad's size falls inside the zone range, rather than silently cutting that size loose and unlabeled. Using `max` on a size OUTSIDE the range still works and is the intended way to fill the rest of a piece.
- **Leftover (unzoned) pads nest over the whole outline, not just below the groups.** Each placed group's footprint is reserved via `_cover_rect_with_circles` and passed to the nester as `preplaced` — a seeded `placed` list that the scan functions collision-check unchanged, so no scan math is touched and the parity contracts hold (`preplaced` defaults to `None`). Seeding the group's *discs* alone is not enough: loose pads settle into the gaps between them, inside someone else's labeled box. The nesters strip the seeds from their return value. Before this, leftovers were clipped to the strip under the lowest group — a hangover from the stacked-band layout that stranded most of a scrap, placing **zero** max-fill pads above or beside the groups.
- **Zones stand down in scrap mode** (`main.py` passes `zones=[]`): a scrap takes only part of a size's count, so a group would have to be re-sized per piece. Revisit only on real demand — the honest use case is ~10 each of a few neighbouring octave sizes, which fits one piece.
- Discs are checked with `spacing_mm + zone_edge_margin_mm` clearance while the boundary is drawn at the group edge, so the engraved line doesn't crowd the outer discs (2.5mm measured, vs 1.0mm without the margin).
- On camera-captured scraps the group rectangles land ~3mm inside the physical edge for free, because placement tests against `custom_polygon` (already inset by `camera_polygon_inset_mm`) rather than `custom_polygon_outline`. Don't "fix" this by switching to the outline. `camera_capture.inset_polygon_mm` is not usable here — it requires OpenCV, and `svg_engine` must stay dependency-light.
- **A zoned size is never cut without its group, on either path.** If a size can't get one, its pads are left unplaced (but still counted in `fixed_total`) so `can_all_pads_fit` reports the shortfall. A too-wide block first degrades through flatter grid shapes before being given up on. The tempting fallback — nest them anyway so nothing is "lost" — was tried and reverted on 2026-08-11: it dropped an unlabeled group of small discs right next to the labeled ones (exactly the confusion zones prevent) while the caller saw a full placed count and reported success. A visible refusal beats a silent unlabeled pile. Out-of-range sizes still nest normally in the leftover region.
- **Groups are placed biggest-first** (`disc_d × qty`) so the roomy parts of a scrap are claimed before small groups fill in around them. The polygon first-fit scans top-left to bottom-right at `ZONE_GRID_SEARCH_STEP_MM`, rejecting cheaply (overlap, then rectangle-in-polygon) before the expensive per-disc `_circle_fits_in_polygon` test — that ordering is what keeps it instant rather than seconds.
- **Both the discs and the drawn rectangle must be on material.** `_rect_fits_in_polygon` samples along each edge, not just the corners, so a concave notch biting into the middle of an edge is caught — otherwise the engraved boundary runs off the scrap and marks nothing.
- Grouping genuinely costs yield — 30 leather discs fit a 150×110 scrap free-nested but need 210×140 grouped. That trade was accepted deliberately: this targets tiny pads, where yield is already high and demand low.
- Borders and labels are **engraved, never cut** — a cut border would drop the zone tile through the bed slats. They're emitted as one engraving pass before any disc is cut, so the sheet is labeled before parts come loose.
- The label is digits and `.` only. `STROKE_FONT`/`FILLED_FONT` in gcode_engine carry no `x`, so a "×10" count suffix would need a new glyph in **both** fonts.

**One shape rule, both paths**: `zone_grid_candidates` is the single source of truth; `_zone_box` computes dimensions from it and accepts a `shape=` override so a caller can walk the list. The rectangular path used to score zone area independently, which silently disagreed with the polygon path — 6 came out 2×3, 8 as 2×4, and a prime like 7 as a **1×7 column** — so the same job produced different blocks depending on sheet type. Caught by the 2026-08-11 initbig audit. If you touch grid selection, `test_both_sheet_types_pick_the_same_grid` and `test_no_quantity_produces_a_single_file_strip` are the guards.

**One box geometry, both paths**: same rule, same failure mode, found the same way. `_group_metrics(settings)` returns `(gutter, border, font, group_gap)` and is the **only** place those keys are read — neither packer reads them directly, because when they did they disagreed: the border was `zone_border_mm` 1.0 on rectangular sheets but `zone_edge_margin_mm` 1.5 on polygons, and the inter-box gap was a hardcoded `ZONE_GAP_MM = 2.0` versus a `zone_band_gap_mm` setting of 6.0. `zone_edge_margin_mm` still exists but now has exactly one job — extra clearance when testing a group's discs against a traced outline — and no longer doubles as the box border. Guard: `test_both_sheet_types_space_boxes_the_same`.

**The gap between boxes is not a moat** (`zone_group_gap_mm`, default 1.0 = the disc gutter). It was 6.0 while each size got a full-width *band*, where the gap had to separate strips. Once a size became a box with its own `zone_border_mm` of white space inside it, 6.0 put ~7mm between neighbouring boxes and ~15mm between discs in adjacent groups, against 1mm between discs inside one — Matt read it as leftover moat on real material, correctly. Tightening it reclaims that: on the test scrap the four-group span went 123×61mm → 111×57mm, and on **leather** — where a 7mm pad is a 17mm disc — it's the difference between one group fitting and two. Don't take it to 0: two abutting engraved lines read as one box. `test_group_gap_is_not_a_moat` pins both ends (gap ≤ 2×border, and > 0) plus "tighter never places fewer pads".

**Preset schema**: only the three user-facing keys (`zone_labels_enabled`, `zone_label_min_size`, `zone_label_max_size`) are in `SIZING_PRESET_KEYS`; gutter/border/font/group-gap are global tuning constants with no UI. Adding keys to that tuple breaks `_detect_active_preset()` for every already-saved preset, so `config.normalize_sizing_preset()` backfills missing keys from `DEFAULT_SETTINGS` before comparing — use it whenever the schema grows.

**Keep test fixtures to real material and real pad sizes.** The two `x max` leftover tests originally filled a 302×158mm scrap with a `4.0mm` pad — a 1.2mm card disc — and placed ~2900 of them, which cost **88s and 71s** and made the suite unusable at ~170s. Matt makes nothing under 7.0mm and works on offcuts never bigger than 14×14in; the fixture is now a 254×162mm offcut filled with `MAX_FILL = 18.0`, and the suite runs in **6.5s**. The old fixture also silently distorted what was being tested: only a sub-millimetre disc fits *above* the groups (they're placed biggest-first from the top and claim the topmost material), so "fills above the groups" was an artifact of the unrealistic size, not a property worth pinning. The real signal is "fills *beside* them" plus a ≥85% comparison against the same fill with zones off.

Tests: `tools/test_zone_labels.py` (46) — opt-in/regression safety (zones off must reproduce the old nester placement-for-placement), range bounds, grid shapes and their fallbacks, containment, group non-overlap and visible separation, groups-and-boundaries-land-on-material, never-cut-unlabeled, the clip helper on concave shapes, SVG↔G-code Y-flip agreement, and three preview-canvas cases that self-skip without a display.

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

**Naming the loaded preset**: on open, `_detect_active_preset()` matches the form's snapshot against each saved preset and selects the one that fits, so the dropdown names what's loaded instead of sitting blank. There is deliberately no stored "active preset" settings key — the applied values themselves identify the preset, which can't go stale after an edit and stays honest when a config is hand-edited (the dropdown just stays blank). This relies on Apply refusing to commit changes that aren't captured in a preset, so applied settings always correspond to a saved one. `GcodeSettingsWindow` does the same per material.

**Bootstrap**: `main.py` auto-creates a `Default` preset from current settings on first run if `sizing_presets` is empty (via `config.settings_to_sizing_preset`). The app guarantees at least one preset always exists.

## G-code Settings Presets Workflow

`GcodeSettingsWindow` (Options > G-code Settings... on Pad Maker, Options > Settings... on Tooling) is preset-aware **per material**. Each material section (felt / card / leather / acrylic / basswood) has its own preset bar at the top with Load / Save / Rename / Delete, and its own active-preset name + dirty baseline. Editing felt does not dirty card. The preset library is shared across the two dialogs — saving a felt preset from Pad Maker shows up in any future dialog that includes felt.

- **Per-material storage**: `gcode_presets.json` shape is `{material: {preset_name: data}}`. Top-level keys are the 5 materials in `config.GCODE_PRESET_MATERIALS`. Inner data captures only `config.GCODE_PRESET_KEYS` (the 19 keys per material — engraving mode, line/filled speed+power+passes, fill density, hole/cut speed+power+passes, kerf, four air toggles).
- **Cross-material isolation by design**: a felt preset will not load into the acrylic slot. Materials have characteristic settings ranges and mixing them silently is dangerous; users who want to cross-apply must Save As under the target material.
- **Apply (bottom button)**: if any material is dirty, a three-way prompt (Yes / No / Cancel) — save dirty materials as preset(s) before applying, apply anyway, or keep editing.
- **Cancel / window-close X**: same three-way prompt, but the "apply anyway" branch becomes "discard and close."
- **Save**: opens `SaveSizingPresetDialog` (generalized — accepts `title`/`intro` kwargs) with material-specific copy ("Save Felt Preset" etc.). Overwrite defaults to the active preset when one is loaded.
- **Delete refuses to wipe the last preset** for that material. **Rename refuses empty / duplicate names.**

Dirty tracking uses per-material `_capture_material_to_dict(mat)` snapshots compared against `self.material_baseline[mat]`. Baselines reset on dialog open, after Load, and after a successful Save Preset. `active_preset_name[mat]` tracks which preset's values currently sit in each material's fields.

**Bootstrap**: `main.py` loads `gcode_presets.json` and, on first run *or* if any material is missing, backfills a `Default` preset for that material from the current `gcode_settings[material]` (via `config.settings_to_gcode_presets`). The app guarantees at least one preset per material always exists.

**Backward compatibility**: `GcodeSettingsWindow(..., gcode_presets=None)` (the default) disables the preset bar entirely — used by `test_smoke_ui` and any future caller that wants the bare settings dialog.

## Settings-Dialog Tooltips

`ui_dialogs.py` exposes a `Tooltip` helper plus `add_tooltip(widget, text)` and `add_tooltips(text, *widgets)` convenience functions. All settings dialogs (Sizing Rules, G-code Settings, Layer Colors, Key Layout, Tuner Settings, Toner Settings) hover-explain their fields. When adding a new setting widget, attach a tooltip alongside it — short sentence, plain English, focused on *why* the user would change it. Attach to both the label and the input so users can hover either. The helper popup is overrideredirect + topmost so it appears above modal dialogs without stealing focus.

## Dart Shape Spectrum

The dart wave shape (formerly "Star/Dart"; now just "Darts" in the UI) is a single 0.0–1.0 slider that smoothly interpolates between three primitive shapes:
- 0.0 = Triangle (linear ramps between peaks/valleys)
- 0.5 = Sine (raw cosine)
- 1.0 = Square (saturating sign function via `|c|^p`, `p = _SQUARE_POWER = 0.01`)

Math lives in `_wave_value` in `svg_engine.py`; `calculate_star_path` (SVG) and `_generate_star_points` (G-code) both call it per sample — `tools/test_engine_parity.py` pins them together. The default value is now `0.5` (sine). Legacy configs used a 0.0=sine, 1.0=square scale; `load_settings` migrates them once via `0.5 + 0.5 * old` and sets `dart_shape_v2: True` so the migration doesn't run again. `dart_ranges[*].shape_factor` is migrated alongside the universal value. When changing dart-shape behavior, update `tools/test_dart_shapes.py` (anchor + smoothness + migration coverage).

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

The strobe tuner has an optional GPU-accelerated renderer in `tuner_renderer/` (Rust crate using pyo3 + wgpu) — **Windows and Linux only**. CI builds this on the Windows and Linux runners, but a local checkout of `python main.py` will silently fall back to the slower canvas renderer unless you build the extension yourself:

```bash
# Requires Rust toolchain (rustup). Windows/Linux only — do NOT do this on
# a Mac (see the macOS warning below).
pip install maturin
python -m maturin build --release --manifest-path tuner_renderer/Cargo.toml
pip install --find-links tuner_renderer/target/wheels tuner_render
```

After this, `import tuner_render` succeeds and the tuner uses GPU rendering at 60-120 fps.

**Three renderer traps, found 2026-10-06** by an outside port of this crate into a Qt app on a 250 %-scaled laptop (Intel UHD, DX12), each verified here on the same GPU before fixing:
- **`Limits::downlevel_defaults()` caps textures at 2048 px.** A 1400 px tuner frame at 250 % is 3500 device px and `Surface::configure` panics ("maximum extent for either dimension is 2048") — in `Renderer::new` and again in `resize()` whenever the window grows past 2048. 150 % on a wide window is enough to hit it. Now `required_limits` is `downlevel_defaults().using_resolution(adapter.limits())` and `clamp_surface()` clamps width/height to `device.limits().max_texture_dimension_2d` in both places, so an oversize window stretches instead of panicking. Measured: 3500×1500 constructs and resizes.
- **A Rust panic reaches Python as `pyo3_runtime.PanicException`, a `BaseException`.** `except Exception` lets it through into Tk's callback. `tuner_tab.py` now catches `BaseException` (re-raising `KeyboardInterrupt`/`SystemExit`) at every call that can touch the surface — init, `resize`, `render` — and routes every failure through one `_tuner_gpu_fallback(reason)`: unpack the GPU frame, pack a canvas in its slot, show a CPU-mode notice that says *unavailable* (not *uninstalled*), log the reason. A single bad `render()` is a dropped frame; `GPU_RENDER_FAIL_LIMIT` (30, ~0.5 s) consecutive ones switch to the canvas. **Never retry the GPU on the same frame**: a failed `configure` leaves the surface dead. `_last_line()` trims the multi-line panic text for the log.
- **Fifo blocked `render()` for a full vsync on the UI thread.** Measured 16.66 ms/frame on Intel UHD/DX12 at 1400×800 — every tuner frame stalled Tk for 16 ms. The surface now takes `PresentMode::Mailbox` where `surface_caps.present_modes` offers it (it does on DX12), else Fifo: 0.45 ms/frame, same vsync-aligned output, no tearing. `present_mode()` reports which one is in use.

Also: **a software adapter counted as GPU mode.** If wgpu only finds Windows' Basic Render Driver or llvmpipe, `Renderer::new` succeeds and the CPU-mode hint never shows. `adapter_info()` exposes `(name, backend, device_type)`; the tab treats `device_type == "Cpu"` as no GPU and falls back. Logged at INFO on every GPU start, so `app.log` names the adapter and present mode. Gates: section 9 of `tools/test_gpu_tuner.py` (big-surface construct/resize, adapter_info/present_mode, and the render-failure fallback with a fake `BaseException`; GUI cases self-skip without a display or GPU). **JustATuner's `renderer.rs` is byte-identical to the pre-fix file and needs the same three changes** — do it there when the tuner's final home is settled, not twice. If the import fails for any reason the tuner falls back to canvas rendering on Windows/Linux — the app still runs, just slower. See the `_skip_theme` / `_dark_canvas` flag notes in the Strobe Tuner architecture section for integration details, and the GPU/Canvas constant alignment warning (`DIM_MULTIPLIER`, `BRIGHTNESS_GAMMA`) for what has to stay in sync between the Python path and the shader.

**macOS is canvas-only — never load `tuner_render` on darwin.** Tk Aqua draws all widgets into a single NSView per toplevel, and `winfo_id()` returns a pointer to Tk's internal `MacDrawable` struct, not an NSView ("the value has no meaning outside Tk" — Tk docs). `tuner_renderer/src/platform.rs` would wrap that handle as an NSView, so wgpu's Metal backend segfaults in `objc_msgSend` during surface creation — a native crash the Python init-failure `except` in `tuner_tab.py` can never catch. Three layers enforce the gate: `tuner_tab.py` skips the `tuner_render` import on darwin, `build.py` skips the `--hidden-import` on darwin, and CI skips the Rust/maturin steps on macOS runners. Even a real NSView wouldn't fix it — a CAMetalLayer on the shared per-window view would paint over the entire UI (and would still need Retina scale handling), so macOS GPU rendering is off the table by design. **History**: the macOS Apple Silicon zips for v1.95 through v2.6 shipped with the renderer bundled — on those builds, opening the Tuner tab crashes the app outright.

On Windows, Rust needs the MSVC linker. The one-time machine setup is `winget install Rustlang.Rustup` followed by `winget install Microsoft.VisualStudio.2022.BuildTools --override "--add Microsoft.VisualStudio.Workload.VCTools --includeRecommended"`.

### Windows Installer (Inno Setup)

CI wraps `dist\SaxShopCompanion.exe` into a versioned `SaxShopCompanion-Windows-Setup-{version}.exe` via `installer.iss`. The installer creates Start Menu + optional desktop shortcuts, registers an uninstaller, and installs to `{autopf}\SaxShopCompanion` (requires admin UAC). User data in `%APPDATA%\StohrerSaxShopCompanion\` is untouched on uninstall. Local build:

```bash
python build.py
iscc /DAppVersion=2.7 installer.iss     # requires Inno Setup 6
```

**Do not change the `AppId` GUID** in `installer.iss` — Windows uses it to recognize upgrades. Changing it produces a parallel install instead of an in-place upgrade.

## Bundled Runtime Assets

Files that need to be reachable at runtime in both source and frozen builds (icons, SVG templates, etc.) follow this pattern:

1. Place the asset at the repo root (single files) or in a subfolder (collections).
2. Extend `build.py`'s PyInstaller `cmd` via `--add-data`. Single files use `'<file>{os.pathsep}.'`; folders use `'<folder>{os.pathsep}<folder>'`.
3. Resolve at runtime with `base = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(__file__)`, then `os.path.join(base, ...)`.

Current examples: `icon.ico` (loaded by `main.py` for the title bar / taskbar icon), `tooling_assets/die_organizer_{upper,lower}.svg` (copied by `generate_die_organizer_svg` in `svg_engine.py`), `pad_press_spacers/*.stl` (copied by `ToolingTabMixin._save_pad_spacer_stl` in `tooling_tab.py`), and `locale/` (compiled translation catalogs resolved by `i18n._locale_dir()`).

## Internationalization (i18n)

User-facing strings are wrapped with `_("...")` and translated via GNU gettext. The catalog is initialized in `main.py` **before** any UI module is imported (so module-level `_()` calls resolve against the active language).

**Pattern in source**:
```python
# `_` and `ngettext` are installed into builtins by i18n.init_translation().
# No explicit import needed. ruff.toml whitelists them in `builtins`.
label = _("Cancel")
msg = _("Imported {n} captures").format(n=n)
plural = ngettext("{n} pad", "{n} pads", n).format(n=n)
```

**Files**:
- `i18n.py` — `init_translation(lang)`, `available_languages()`, module-level `_` and `ngettext` handles for tests
- `babel.cfg` — extraction config for pybabel
- `locale/saxshop.pot` — generated template (commit it; regenerate after adding strings)
- `locale/<lang>/LC_MESSAGES/saxshop.po` + `saxshop.mo` — per-language catalogs (commit both)

**Workflow when adding or changing strings**:
```bash
python tools/extract_strings.py        # regenerate saxshop.pot
python tools/update_translations.py    # merge .pot into existing .po files (marks new entries fuzzy)
# Edit each locale/<lang>/LC_MESSAGES/saxshop.po — translate the fuzzy entries
python tools/compile_translations.py   # rebuild all .mo files
python tools/test_i18n.py              # verify
```

**Languages**: v1 ships English (source) + Spanish / German / French / Italian. Native-name display in `i18n.LANGUAGE_NAMES`. Language switching is restart-required, picked via File > Feature Set > Language.

**Gotchas**:
- **Never assign to `_` in a function that also calls `_()`.** `_` is gettext, installed into builtins by `gettext.install()`; Python makes a name local for the WHOLE function body if it is assigned anywhere in it, so a single `for _, x in items` or `a, _, _ = f()` turns every `_("...")` in that function into `UnboundLocalError` (before the assignment) or `'tuple' object is not callable` (after it). Three shipped functions had this on 2026-10-06: the toner's Analyze dialog could not open at all, every die G-code generation ended in an "Unexpected Error" dialog after writing the file, and G-code Settings Save crashed on any invalid field. The earlier version of this note said a loop-local `_` needed no rename — that was wrong whenever the function translates anything. Use `_unused`, `_name`, `_label`. Gated by `tools/test_gettext_shadowing.py`, a scope-aware AST scan of every root module (nested defs, lambdas and comprehensions have their own `_`).
- Module-level constants with translatable strings: define as a function (e.g. `get_resonance_messages()`) so each call resolves against the *current* catalog. Lists evaluated at module import time bake the source-language values.
- Don't translate data — pad sizes, material keys (`"felt"`, `"card"`), settings keys are code. Only translate at the display layer.
- f-strings with embedded variables: `_("Imported {n} captures").format(n=n)`. Don't put the f-prefix on the gettext string itself or the placeholder gets baked.

## Branching Strategy

**ALWAYS sync with the remote before doing anything else.** Matt develops this app from several
different computers, so a local checkout is frequently behind `origin/beta` — and may also hold
unpushed local commits, making the branch *diverged* rather than simply stale. Before reading code,
answering questions about what the app does, or making any edit:

```bash
git fetch --all --prune
git status --short --branch                       # ahead / behind counts
git log --oneline beta..origin/beta               # what this machine is missing
git log --oneline origin/beta..beta               # what this machine hasn't pushed
```

- **Behind only** → `git pull --rebase` and continue.
- **Diverged (ahead *and* behind)** → `git pull --rebase` so local work replays on top; check the
  local commits for conflicts with what landed upstream before pushing.
- Never start work off a stale checkout. `APP_VERSION` in the local `config.py`, the local release
  notes, and memory files are all unreliable until this sync has run — check
  `git show origin/beta:config.py | grep APP_VERSION` and `gh release list --limit 5` for the real
  current version.

The `.claude/settings.json` `SessionStart` hook runs the fetch + divergence report automatically at
the start of each session, but it only *reports* — reconciling is still a deliberate step.

- **`main`**: Stable release branch. Merges from `beta` when features are tested and ready.
- **`beta`**: Active development branch. New features land here first (e.g. filled engraving, air assist toggles, cut grouping). Always work on `beta` unless told otherwise.
- CI builds trigger on push to `main` or `beta`.

## Versioning

`APP_VERSION` in `config.py` is the manual source of truth — bump it when preparing a release. `APP_BUILD_DATE` is auto-derived from `sys.executable`'s mtime in frozen builds (installer/zip copies preserve mtime) and falls back to the manual constant when running from source. The About dialog reads both via `ui_dialogs.py`.

## Release Notes Style

Every GitHub release uses the same body template: a short single-paragraph lead describing the app, then the **full feature overview** organized tab-by-tab (Pad Maker, Tooling, Chromatic Strobe Tuner, Harmonic Tone Analyzer, Cross-platform & General), followed by **Known limitations** and **Upgrading from v1.x**. Every release is the complete picture of the app, so a user landing on any release page cold gets the full story — no "What's new since" sections or fix logs.

Items new *in that specific release* get a plain `**(new)**` marker prefixed before the relevant bullet (or interpolated into a longer bullet that has both old and new content). Items from prior releases carry no marker. v2.1 is the canonical example — see [the v2.1 release page](https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases/tag/v2.1) for the exact format. Don't mix in `(new in v2.X.Y)` cross-version tags or "Fixes & polish" sections — both got tried and rejected.

## CI/CD (GitHub Actions)

The `.github/workflows/build.yml` workflow has two jobs:
- **`lint`** (ubuntu-latest, ~10s): runs `ruff check .` — fails the workflow on any violation
- **`build`** (4-platform matrix): Windows Inno Setup installer (the bare PyInstaller .exe is built but not published — only the installer ships), macOS Apple Silicon .app, macOS Intel .app, and Linux binary

Triggers on push to `main` or `beta`, on release creation, or manually.

- macOS Intel build (`macos-15-intel` runner) installs svgwrite+numpy+pyinstaller (no sounddevice) — tuner and toner are unavailable; nesting runs at full numpy speed (numpy added 2026-10-06 — before that the pure-Python scan took 13 minutes on one test suite there)
- `full_build: true/false` matrix flag controls whether Rust toolchain + maturin are installed for the GPU tuner renderer; the macOS runners additionally skip the Rust steps unconditionally — macOS is canvas-only (see GPU Tuner Renderer section)
- The Windows job also installs Inno Setup 6 via Chocolatey and builds the installer; version is extracted from `config.py`'s `APP_VERSION` via PowerShell regex
- macOS jobs package the `.app` with `ditto -c -k --keepParent` (never `zip -r` — it materializes bundle symlinks and breaks the code-signature resource seal) and run three guards: `plutil -extract` for the mic + camera usage keys, `codesign --verify --deep --strict` on the built .app, and the same verify on an unzipped copy of the final artifact (see "macOS Build" in CLAUDE-architecture.md)
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

## Detailed Documentation

Architecture and domain-specific guidance is split across companion files imported via `@`-statements below. Claude Code loads them as part of CLAUDE.md's context.

- **CLAUDE-architecture.md** — Module structure, design patterns, settings/presets, error logging, feature set
- **CLAUDE-engines.md** — Pad generation, G-code, SVG rendering, nesting, strobe tuner, tooling, Phil Noy credit
- **CLAUDE-toner.md** — Tone analyzer engine, data model, capture modes, analyze tool, WAV recording, calibration
- **CLAUDE-web.md** — Web data sync, screw specs submission form, related repository

---

## Subsystem imports

@CLAUDE-architecture.md
@CLAUDE-engines.md
@CLAUDE-toner.md
@CLAUDE-web.md
