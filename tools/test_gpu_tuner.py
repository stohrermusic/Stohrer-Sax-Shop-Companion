"""
Test script for GPU strobe tuner rendering integration.

Tests: tuner_render module API, tuner_tab.py GPU/canvas dual path,
layout computation, phase direction, settings, fallback behavior,
perf log compatibility, icon loading.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

passed = 0
failed = 0


def test(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


# ============================================================
# 1. tuner_render module import and API
# ============================================================
print("\n--- tuner_render Module ---")

try:
    import tuner_render
    _has_gpu = True
    test("tuner_render imports successfully", True)
except ImportError:
    _has_gpu = False
    test("tuner_render imports successfully", False)
    print("  (GPU tests will be skipped)")

if _has_gpu:
    test("TunerRenderer class exists", hasattr(tuner_render, 'TunerRenderer'))

    # Check the class has the expected methods
    # pyo3 #[new] maps to __init__, other methods keep their names
    methods = ['resize', 'set_layout', 'render',
               'set_stripe_color', 'set_faceplate_color']
    for m in methods:
        test(f"TunerRenderer has {m} method",
             hasattr(tuner_render.TunerRenderer, m) or
             callable(getattr(tuner_render.TunerRenderer, m, None)))


# ============================================================
# 2. tuner_tab.py module-level flags
# ============================================================
print("\n--- tuner_tab Module Flags ---")

import tuner_tab

test("_HAS_GPU_RENDERER flag exists", hasattr(tuner_tab, '_HAS_GPU_RENDERER'))
test("_HAS_GPU_RENDERER matches import", tuner_tab._HAS_GPU_RENDERER == _has_gpu)
test("_TUNER_IMPORTS_OK flag exists", hasattr(tuner_tab, '_TUNER_IMPORTS_OK'))
test("TRANSPOSITION_KEYS exists", hasattr(tuner_tab, 'TRANSPOSITION_KEYS'))


# ============================================================
# 4. Layout computation
# ============================================================
print("\n--- Layout Computation ---")

# We need a mock object with the method
class MockMixin:
    pass

# Bind the method to our mock
import types
mixin = MockMixin()
mixin._tuner_compute_layout = types.MethodType(
    tuner_tab.TunerTabMixin._tuner_compute_layout, mixin)

# Too small should return None
test("Layout returns None for small area (50x50)",
     mixin._tuner_compute_layout(50, 50) is None)
test("Layout returns None for small area (99x99)",
     mixin._tuner_compute_layout(99, 99) is None)

# Normal size should return 12 entries
layout = mixin._tuner_compute_layout(800, 400)
test("Layout returns 12 entries for 800x400", layout is not None and len(layout) == 12)

if layout:
    # Each entry is (pc, cx, cy, radius, is_up)
    pcs = [l[0] for l in layout]
    test("All 12 pitch classes present", sorted(pcs) == list(range(12)))

    # Check accidentals are in top row (is_up=True)
    for pc, cx, cy, radius, is_up in layout:
        if pc in {1, 3, 6, 8, 10}:
            test(f"pc {pc} (accidental) is_up=True", is_up is True)
        else:
            test(f"pc {pc} (natural) is_up=False", is_up is False)

    # All radii should be the same
    radii = [l[3] for l in layout]
    test("All wheels same radius", len(set(radii)) == 1)

    # Radius should be positive and reasonable
    test(f"Radius > 0 ({radii[0]:.1f})", radii[0] > 0)
    test(f"Radius < half height ({radii[0]:.1f} < 200)", radii[0] < 200)

    # Top row should have smaller cy than bottom row
    top_cys = [cy for pc, cx, cy, r, up in layout if up]
    bot_cys = [cy for pc, cx, cy, r, up in layout if not up]
    test("Top row above bottom row", max(top_cys) < min(bot_cys))


# ============================================================
# 5. Phase direction convention
# ============================================================
print("\n--- Phase Direction ---")

# The shader uses: rotated = angle - phase
# Sharp (positive cents) -> engine decreases phase -> phase negative
# Negative phase -> rotated = angle - (-phase) = angle + |phase| -> CW rotation
# This is verified by reading the shader source

shader_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                           'tuner_renderer', 'src', 'shader.wgsl')
if os.path.exists(shader_path):
    with open(shader_path) as f:
        shader_src = f.read()
    test("Shader uses per-ring phase 'angle - wheel.ring_phases[ring_idx]'",
         'angle - wheel.ring_phases[ring_idx]' in shader_src)
    test("Shader does NOT use 'angle + wheel.ring_phases' (wrong direction)",
         'angle + wheel.ring_phases' not in shader_src)
else:
    test("Shader file exists", False)

# Verify engine phase direction: sharp -> phase decreases
from tuner_engine import TunerEngine
engine = TunerEngine()
# The phase update is in analyze_buffer, not analyze
import inspect
src = inspect.getsource(engine.analyze_buffer)
test("Engine: phase += cents * drift (sharp -> increase -> CW)",
     'phase_offsets[pc] += cents' in src or
     '_phase_offsets[pc] += cents' in src)


# ============================================================
# 6. Constants consistency
# ============================================================
print("\n--- Constants Consistency ---")

# Shader constants should match Python constants
test("WEDGE_ANGLE = 80 degrees", tuner_tab.WEDGE_ANGLE == 80.0)
test("NUM_RINGS = 7", tuner_tab.NUM_RINGS == 7)
test("RING_SEGMENTS has 7 entries", len(tuner_tab.RING_SEGMENTS) == 7)
test("RING_SEGMENTS = [4,8,16,32,64,128,256]",
     tuner_tab.RING_SEGMENTS == [4, 8, 16, 32, 64, 128, 256])
test("CENTER_GAP_FRACTION = 0.12", tuner_tab.CENTER_GAP_FRACTION == 0.12)
test("BRIGHTNESS_GAMMA = 0.45", tuner_tab.BRIGHTNESS_GAMMA == 0.45)
test("MAGNITUDE_THRESHOLD = 0.02", tuner_tab.MAGNITUDE_THRESHOLD == 0.02)
test("DIM_MULTIPLIER = 0.08", tuner_tab.DIM_MULTIPLIER == 0.08)

if os.path.exists(shader_path):
    with open(shader_path) as f:
        shader_src = f.read()
    test("Shader CENTER_GAP = 0.12", 'CENTER_GAP: f32 = 0.12' in shader_src)
    test("Shader WEDGE_HALF_ANGLE ~= 40°",
         'WEDGE_HALF_ANGLE: f32 = 0.6981' in shader_src)
    test("Shader DIM_MULTIPLIER = 0.08", 'DIM_MULTIPLIER: f32 = 0.08' in shader_src)
    test("Shader has 7 ring segment counts",
         all(f'return {float(s)}' in shader_src
             for s in [4, 8, 16, 32, 64, 128, 256]))


# ============================================================
# 7. Settings / config defaults
# ============================================================
print("\n--- Settings ---")

from config import DEFAULT_SETTINGS

ts = DEFAULT_SETTINGS.get("tuner_settings", {})
test("tuner_settings exists in defaults", "tuner_settings" in DEFAULT_SETTINGS)
test("stripe_color default", ts.get("stripe_color") == "#00FF00")
test("faceplate_color default", ts.get("faceplate_color") == "#1A1A1A")
test("ring_brightness default", ts.get("ring_brightness") == 100)
test("overall_brightness default", ts.get("overall_brightness") == 80)
test("octave_boost default", ts.get("octave_boost") == 50)
test("fps default", ts.get("fps") == "60")
test("sensitivity default", ts.get("sensitivity") == 50)
# show_fps is not in DEFAULT_SETTINGS — loaded via .get() with fallback in tuner_tab
test("show_fps loaded via .get() fallback",
     'tuner_settings.get("show_fps", False)' in
     open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'tuner_tab.py')).read())

# octave_boost loaded from settings
test("octave_boost loaded from tuner_settings",
     'tuner_settings.get("octave_boost"' in
     open(os.path.join(os.path.dirname(os.path.dirname(__file__)), 'tuner_tab.py')).read())


# ============================================================
# 8. GPU data struct layout (Rust ↔ WGSL alignment)
# ============================================================
print("\n--- GPU Data Layout ---")

renderer_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                             'tuner_renderer', 'src', 'renderer.rs')
if os.path.exists(renderer_path):
    with open(renderer_path) as f:
        rs_src = f.read()
    test("GpuWheelData compile-time size assert (128 bytes)",
         'size_of::<GpuWheelData>() == 128' in rs_src)
    test("Globals compile-time size assert (16 bytes)",
         'size_of::<Globals>() == 16' in rs_src)
    test("sRGB to linear conversion present",
         'srgb_to_linear' in rs_src)
    test("Hex color parsing present",
         'hex_to_rgb' in rs_src)
else:
    test("renderer.rs exists", False)


# ============================================================
# 9. Platform support
# ============================================================
print("\n--- Platform Support ---")

platform_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                             'tuner_renderer', 'src', 'platform.rs')
if os.path.exists(platform_path):
    with open(platform_path) as f:
        plat_src = f.read()
    test("Windows support (Win32WindowHandle)", 'Win32WindowHandle' in plat_src)
    test("macOS support (AppKitWindowHandle)", 'AppKitWindowHandle' in plat_src)
    test("Linux support (XlibWindowHandle)", 'XlibWindowHandle' in plat_src)
    test("Linux uses c_ulong (not NonZero<u32>)", 'c_ulong' in plat_src)
else:
    test("platform.rs exists", False)


# ============================================================
# 10. Build script
# ============================================================
print("\n--- Build Script ---")

build_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'build.py')
with open(build_path) as f:
    build_src = f.read()

test("build.py has tuner_render hidden import",
     "'tuner_render'" in build_src and 'hidden-import' in build_src)
test("build.py skips tuner_render on macOS",
     'skipping tuner_render' in build_src and
     "sys.platform == 'darwin'" in build_src)
test("build.py has icon.ico support", "'icon.ico'" in build_src)
test("build.py has icon.icns support", "'icon.icns'" in build_src)
test("build.py bundles icon.ico as data", "'--add-data'" in build_src and 'icon.ico' in build_src)


# ============================================================
# 11. App icon
# ============================================================
print("\n--- App Icon ---")

repo_root = os.path.dirname(os.path.dirname(__file__))
ico_path = os.path.join(repo_root, 'icon.ico')
icns_path = os.path.join(repo_root, 'icon.icns')

test("icon.ico exists", os.path.exists(ico_path))
test("icon.icns exists", os.path.exists(icns_path))

if os.path.exists(ico_path):
    size = os.path.getsize(ico_path)
    test(f"icon.ico has multiple sizes ({size} bytes)", size > 10000)

if os.path.exists(icns_path):
    size = os.path.getsize(icns_path)
    test(f"icon.icns is non-trivial ({size} bytes)", size > 1000)

# main.py loads the icon
main_path = os.path.join(repo_root, 'main.py')
with open(main_path) as f:
    main_src = f.read()
test("main.py has iconbitmap call", 'iconbitmap' in main_src)
test("main.py handles frozen (PyInstaller) path", 'sys._MEIPASS' in main_src or '_MEIPASS' in main_src)


# ============================================================
# 12. CI workflow
# ============================================================
print("\n--- CI Workflow ---")

ci_path = os.path.join(repo_root, '.github', 'workflows', 'build.yml')
if os.path.exists(ci_path):
    with open(ci_path) as f:
        ci_src = f.read()
    test("CI installs Rust toolchain", 'rust-toolchain' in ci_src)
    test("CI installs maturin", 'maturin' in ci_src)
    test("CI builds tuner_render", 'maturin build' in ci_src)
    test("CI installs tuner_render wheel", 'tuner_render' in ci_src)
    test("Rust build only on full_build", 'matrix.full_build' in ci_src)
    test("CI skips Rust build on macOS runners",
         "!startsWith(matrix.os, 'macos')" in ci_src)
else:
    test("CI workflow exists", False)


# ============================================================
# 13. Fallback behavior
# ============================================================
print("\n--- Fallback Behavior ---")

tab_path = os.path.join(repo_root, 'tuner_tab.py')
with open(tab_path) as f:
    tab_src = f.read()

test("GPU import wrapped in try/except",
     'try:\n        import tuner_render' in tab_src)
test("_HAS_GPU_RENDERER set on ImportError",
     '_HAS_GPU_RENDERER = False' in tab_src)
# macOS: Tk Aqua's winfo_id() is not an NSView — handing it to wgpu
# segfaults natively, so darwin must never import tuner_render at all.
test("macOS never imports tuner_render (darwin gate)",
     'if IS_MACOS:\n    _HAS_GPU_RENDERER = False' in tab_src)
test("CPU-mode install hint suppressed on macOS",
     'if not IS_MACOS:' in tab_src)
test("GPU init failure falls back to canvas",
     'self._tuner_use_gpu = False' in tab_src and
     '_tuner_build_wheels_canvas' in tab_src)
test("CPU mode notice shown when no GPU",
     'CPU mode' in tab_src)
test("Canvas path preserved (StrobeWheel class exists)",
     'class StrobeWheel' in tab_src)
test("_tuner_build_wheels dispatches to both paths",
     '_tuner_build_wheels_gpu' in tab_src and
     '_tuner_build_wheels_canvas' in tab_src)

# On macOS, Cmd-Q / app-menu Quit fire ::tk::mac::Quit, which bypasses
# WM_DELETE_WINDOW — without this route, on_exit (and the settings save)
# never runs on the standard mac quit path.
test("::tk::mac::Quit routed through on_exit (Cmd-Q saves settings)",
     '"::tk::mac::Quit"' in main_src and 'createcommand' in main_src)


# ============================================================
# 14. Perf log compatibility
# ============================================================
print("\n--- Perf Log ---")

test("Perf log uses .get() for sample keys",
     "s.get(" in tab_src and "'wheels_updated'" not in
     tab_src.split('_tuner_dump_perf_log')[1].split('\n    def ')[0])
test("Perf log includes GPU flag in samples",
     "'gpu'" in tab_src)
test("FPS counter throttled (1 update/sec)",
     '_tuner_fps_last_update' in tab_src)


# ============================================================
# 15. GPU renderer functional test (if available)
# ============================================================
if _has_gpu:
    print("\n--- GPU Renderer Functional ---")
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()  # hidden window

    frame = tk.Frame(root, width=400, height=300)
    frame.pack()
    root.update()

    try:
        hwnd = frame.winfo_id()
        renderer = tuner_render.TunerRenderer(hwnd, 400, 300)
        test("TunerRenderer created from tkinter frame", True)

        renderer.set_stripe_color("#00FF00")
        test("set_stripe_color accepts hex string", True)

        renderer.set_faceplate_color("#1A1A1A")
        test("set_faceplate_color accepts hex string", True)

        renderer.set_layout([
            (100.0, 100.0, 40.0, True),
            (200.0, 200.0, 40.0, False),
        ])
        test("set_layout accepts position tuples", True)

        renderer.resize(800, 600)
        test("resize works", True)

        # Render a frame with valid data (per-ring phases: 12 lists of 7)
        ring_phases = [[float(i * 30 + r * 2) for r in range(7)] for i in range(12)]
        mags = [0.5] * 12
        rmags = [[0.3, 0.5, 0.7, 0.5, 0.3, 0.2, 0.1]] * 12
        renderer.render(ring_phases, mags, rmags, 80.0, 100.0)
        test("render() with valid data succeeds", True)

        # Render with edge cases
        renderer.render([[0.0]*7]*12, [0.0]*12, [[0.0]*7]*12, 0.0, 0.0)
        test("render() with all zeros succeeds", True)

        renderer.render([[360.0]*7]*12, [1.0]*12, [[1.0]*7]*12, 100.0, 150.0)
        test("render() with max values succeeds", True)

        # Multiple rapid frames (stress test)
        for i in range(60):
            p = [[float(i * 6 + r) for r in range(7)]] * 12
            renderer.render(p, mags, rmags, 80.0, 100.0)
        test("60 rapid frames without error", True)

    except Exception as e:
        test(f"GPU renderer functional test ({e})", False)
    finally:
        root.destroy()


# ============================================================
# 16. Tuner engine still works
# ============================================================
print("\n--- Tuner Engine ---")

try:
    from tuner_engine import TunerEngine, TunerResult

    engine = TunerEngine()
    test("TunerEngine instantiates", True)

    result = engine.analyze()
    test("analyze() returns TunerResult", isinstance(result, TunerResult))
    test("result.magnitudes has 12 entries", len(result.magnitudes) == 12)
    test("result.phase_offsets has 12 entries", len(result.phase_offsets) == 12)
    test("result.ring_magnitudes has 12 entries", len(result.ring_magnitudes) == 12)
    test("result.ring_magnitudes[0] has 7 entries", len(result.ring_magnitudes[0]) == 7)
    test("result.cents_errors has 12 entries", len(result.cents_errors) == 12)

    engine.set_reference_pitch(440.0)
    test("set_reference_pitch works", True)

    engine.set_sensitivity(75)
    test("set_sensitivity works", True)

    engine.reset_phases()
    test("reset_phases works", True)
    test("phases zeroed after reset", all(p == 0.0 for p in engine._phase_offsets))

except Exception as e:
    test(f"Tuner engine test ({e})", False)


# ============================================================
# SUMMARY
# ============================================================
print(f"\n{'='*50}")
total = passed + failed
print(f"Results: {passed}/{total} passed, {failed} failed")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print(f"FAILURES: {failed}")
    sys.exit(1)
