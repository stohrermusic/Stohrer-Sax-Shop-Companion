"""
Non-interactive test suite for Stohrer Sax Shop Companion.

Tests the code paths affected by three recent bug fixes:
1. Polygon G-code negative Y coordinates (Y-flip used wrong height)
2. Laser head not returning home (M5 before return move)
3. Polygon closing difficulty (UI-only, but we test normalization logic)

Also tests surrounding features to catch regressions.

Usage:
    python tools/test_bugfixes.py

No GUI windows are opened. All output goes to stdout.
"""

import copy
import os
import re
import sys
import tempfile

# Add parent directory to path so we can import project modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Prevent config.py from triggering tkinter messagebox on import
# by mocking the messagebox module before importing config
import types
_fake_tk = types.ModuleType("tkinter")
_fake_mb = types.ModuleType("tkinter.messagebox")
_fake_mb.showerror = lambda *a, **kw: None
_fake_mb.showinfo = lambda *a, **kw: None
_fake_mb.showwarning = lambda *a, **kw: None
_fake_tk.messagebox = _fake_mb
sys.modules.setdefault("tkinter", _fake_tk)
sys.modules.setdefault("tkinter.messagebox", _fake_mb)

from config import DEFAULT_SETTINGS
from gcode_engine import (
    generate_gcode,
    generate_gcode_footer,
    generate_gcode_from_placed,
    generate_gcode_header,
    generate_gcode_layer,
)
from svg_engine import (
    _nest_discs,
    generate_svg,
    generate_svg_from_placed,
    try_nest_partial,
)


# =============================================================================
# TEST FRAMEWORK
# =============================================================================

_results = []

def run_test(name, fn):
    """Run a test function and record PASS/FAIL."""
    try:
        fn()
        _results.append(("PASS", name, None))
        print(f"  PASS  {name}")
    except AssertionError as e:
        _results.append(("FAIL", name, str(e)))
        print(f"  FAIL  {name} -- {e}")
    except Exception as e:
        _results.append(("ERROR", name, str(e)))
        print(f"  ERROR {name} -- {type(e).__name__}: {e}")


def assert_true(condition, msg=""):
    if not condition:
        raise AssertionError(msg or "Expected True but got False")

def assert_false(condition, msg=""):
    if condition:
        raise AssertionError(msg or "Expected False but got True")

def assert_equal(a, b, msg=""):
    if a != b:
        raise AssertionError(msg or f"Expected {a!r} == {b!r}")

def assert_approx(a, b, tol=0.001, msg=""):
    if abs(a - b) > tol:
        raise AssertionError(msg or f"Expected {a} ~= {b} (tol={tol})")

def get_settings():
    """Return a deep copy of DEFAULT_SETTINGS for test isolation."""
    return copy.deepcopy(DEFAULT_SETTINGS)


def make_temp_file(suffix=".gcode"):
    """Create a temp file and return its path. Caller should delete it."""
    fd, path = tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    return path


# =============================================================================
# HELPERS: Parse G-code coordinates from a file
# =============================================================================

_COORD_RE = re.compile(
    r"(?:G[01])\s*"
    r"(?:X(-?[\d.]+))?"
    r"(?:Y(-?[\d.]+))?",
    re.IGNORECASE,
)

def parse_gcode_coordinates(filepath):
    """Extract all (x, y) coordinate pairs from G0/G1 moves in a G-code file.
    Returns list of (x, y) floats.  If a move omits X or Y, uses last known value."""
    coords = []
    last_x, last_y = 0.0, 0.0
    with open(filepath, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith(";"):
                continue
            m = _COORD_RE.match(line)
            if m:
                x_str, y_str = m.group(1), m.group(2)
                if x_str is not None:
                    last_x = float(x_str)
                if y_str is not None:
                    last_y = float(y_str)
                if x_str is not None or y_str is not None:
                    coords.append((last_x, last_y))
    return coords


def parse_gcode_lines(filepath):
    """Read all lines from a G-code file."""
    with open(filepath, "r") as f:
        return [line.rstrip("\n") for line in f]


# =============================================================================
# HELPERS: Polygon normalization (mirrors main.py logic)
# =============================================================================

def normalize_polygon_inches(polygon_grid):
    """Convert an inch-grid polygon to normalized mm polygon.
    polygon_grid: list of (x, y) in inches on the 15x15 grid.
    Returns list of (x, y) in mm with bounding box starting at (0, 0)."""
    grid_size = 15  # 15x15 inches
    raw = [(x * 25.4, (grid_size - y) * 25.4) for (x, y) in polygon_grid]
    min_x = min(p[0] for p in raw)
    min_y = min(p[1] for p in raw)
    return [(x - min_x, y - min_y) for (x, y) in raw]


def normalize_polygon_cm(polygon_grid):
    """Convert a cm-grid polygon to normalized mm polygon.
    polygon_grid: list of (x, y) in cm on the 40x40 grid.
    Returns list of (x, y) in mm with bounding box starting at (0, 0)."""
    grid_size = 40  # 40x40 cm
    raw = [(x * 10, (grid_size - y) * 10) for (x, y) in polygon_grid]
    min_x = min(p[0] for p in raw)
    min_y = min(p[1] for p in raw)
    return [(x - min_x, y - min_y) for (x, y) in raw]


# =============================================================================
# TEST GROUP: G-code Footer (Bug Fix #2 - M5 after return move)
# =============================================================================

def test_footer_m5_after_return():
    """M5 must come AFTER the G1 X0Y0 return move."""
    lines = generate_gcode_footer(return_speed=1000)
    joined = "\n".join(lines)
    idx_return = None
    idx_m5 = None
    for i, line in enumerate(lines):
        if "X0Y0" in line:
            idx_return = i
        if line.strip() == "M5":
            idx_m5 = i
    assert_true(idx_return is not None, "G1 X0Y0 return move not found in footer")
    assert_true(idx_m5 is not None, "M5 not found in footer")
    assert_true(idx_m5 > idx_return,
                f"M5 (line {idx_m5}) must come AFTER return move (line {idx_return}). Footer:\n{joined}")


def test_footer_s0_before_return():
    """G1 S0 (laser safe) must come BEFORE the return move."""
    lines = generate_gcode_footer(return_speed=1000)
    joined = "\n".join(lines)
    idx_s0 = None
    idx_return = None
    for i, line in enumerate(lines):
        if "S0" in line and "G1" in line:
            idx_s0 = i
        if "X0Y0" in line:
            idx_return = i
    assert_true(idx_s0 is not None, "G1 S0 not found in footer")
    assert_true(idx_return is not None, "G1 X0Y0 return move not found in footer")
    assert_true(idx_s0 < idx_return,
                f"G1 S0 (line {idx_s0}) must come BEFORE return move (line {idx_return}). Footer:\n{joined}")


def test_footer_m2_last():
    """M2 (program end) must be the last line of the footer."""
    lines = generate_gcode_footer(return_speed=1000)
    assert_equal(lines[-1].strip(), "M2", f"Last footer line should be M2, got: {lines[-1]}")


def test_footer_return_speed_custom():
    """Footer return move should use the specified return speed."""
    for speed in [500, 1000, 2000, 5000]:
        lines = generate_gcode_footer(return_speed=speed)
        return_line = [l for l in lines if "X0Y0" in l]
        assert_true(len(return_line) == 1, f"Expected exactly one X0Y0 line for speed={speed}")
        assert_true(f"F{speed}" in return_line[0],
                    f"Expected F{speed} in return line, got: {return_line[0]}")


def test_footer_m9_air_off():
    """M9 (air off) should be present in the footer."""
    lines = generate_gcode_footer(return_speed=1000)
    assert_true(any(l.strip() == "M9" for l in lines), "M9 not found in footer")


def test_footer_complete_sequence():
    """Verify the complete expected footer sequence."""
    lines = generate_gcode_footer(return_speed=1000)
    stripped = [l.strip() for l in lines if l.strip() and not l.strip().startswith(";")]
    # Expected order: M9, G1 S0, G90, G1 X0Y0 F1000, M5, M2
    assert_equal(stripped[0], "M9", f"Footer[0] should be M9, got {stripped[0]}")
    assert_equal(stripped[1], "G1 S0", f"Footer[1] should be G1 S0, got {stripped[1]}")
    assert_equal(stripped[2], "G90", f"Footer[2] should be G90, got {stripped[2]}")
    assert_true("X0Y0" in stripped[3], f"Footer[3] should contain X0Y0, got {stripped[3]}")
    assert_equal(stripped[4], "M5", f"Footer[4] should be M5, got {stripped[4]}")
    assert_equal(stripped[5], "M2", f"Footer[5] should be M2, got {stripped[5]}")


# =============================================================================
# TEST GROUP: Polygon Normalization (Bug Fix #3 - supports polygon closing fix)
# =============================================================================

def test_normalize_polygon_origin_inch():
    """Normalized inch polygon should have min x = 0 and min y = 0."""
    # Polygon drawn in middle of 15x15 inch grid
    grid_poly = [(5, 5), (10, 5), (10, 10), (5, 10)]
    norm = normalize_polygon_inches(grid_poly)
    min_x = min(p[0] for p in norm)
    min_y = min(p[1] for p in norm)
    assert_approx(min_x, 0.0, msg=f"Min X should be 0, got {min_x}")
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0, got {min_y}")


def test_normalize_polygon_origin_cm():
    """Normalized cm polygon should have min x = 0 and min y = 0."""
    # Polygon drawn off-center on 40x40 cm grid
    grid_poly = [(15, 10), (30, 10), (30, 25), (15, 25)]
    norm = normalize_polygon_cm(grid_poly)
    min_x = min(p[0] for p in norm)
    min_y = min(p[1] for p in norm)
    assert_approx(min_x, 0.0, msg=f"Min X should be 0, got {min_x}")
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0, got {min_y}")


def test_normalize_preserves_shape_inches():
    """Normalization should preserve relative distances between vertices (inches)."""
    grid_poly = [(3, 2), (8, 2), (8, 7), (3, 7)]
    norm = normalize_polygon_inches(grid_poly)
    # Width and height in mm should be 5 inches = 127mm
    xs = [p[0] for p in norm]
    ys = [p[1] for p in norm]
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    assert_approx(width, 5 * 25.4, tol=0.1,
                  msg=f"Width should be {5*25.4}mm, got {width}mm")
    assert_approx(height, 5 * 25.4, tol=0.1,
                  msg=f"Height should be {5*25.4}mm, got {height}mm")


def test_normalize_preserves_shape_cm():
    """Normalization should preserve relative distances between vertices (cm)."""
    grid_poly = [(10, 5), (25, 5), (25, 20), (10, 20)]
    norm = normalize_polygon_cm(grid_poly)
    # Width should be 15cm = 150mm, height should be 15cm = 150mm
    xs = [p[0] for p in norm]
    ys = [p[1] for p in norm]
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    assert_approx(width, 150.0, tol=0.1, msg=f"Width should be 150mm, got {width}mm")
    assert_approx(height, 150.0, tol=0.1, msg=f"Height should be 150mm, got {height}mm")


def test_normalize_polygon_high_y_values():
    """Polygon at bottom of grid (high Y in grid coords, near Y=0 after flip) normalizes correctly."""
    # Points near bottom of 15x15 inch grid (Y close to 15)
    grid_poly = [(2, 12), (6, 12), (6, 14.5), (2, 14.5)]
    norm = normalize_polygon_inches(grid_poly)
    min_x = min(p[0] for p in norm)
    min_y = min(p[1] for p in norm)
    assert_approx(min_x, 0.0, msg=f"Min X should be 0, got {min_x}")
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0, got {min_y}")
    # All coords should be non-negative
    for x, y in norm:
        assert_true(x >= -0.001, f"Negative X coordinate: {x}")
        assert_true(y >= -0.001, f"Negative Y coordinate: {y}")


def test_normalize_polygon_at_grid_origin():
    """Polygon touching grid origin (0,0) should still normalize correctly."""
    grid_poly = [(0, 0), (5, 0), (5, 5), (0, 5)]
    norm = normalize_polygon_inches(grid_poly)
    min_x = min(p[0] for p in norm)
    min_y = min(p[1] for p in norm)
    assert_approx(min_x, 0.0, msg=f"Min X should be 0, got {min_x}")
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0, got {min_y}")


def test_normalize_irregular_polygon():
    """Irregular (non-rectangular) polygon normalizes correctly."""
    # L-shaped polygon
    grid_poly = [(3, 3), (8, 3), (8, 6), (5, 6), (5, 8), (3, 8)]
    norm = normalize_polygon_inches(grid_poly)
    min_x = min(p[0] for p in norm)
    min_y = min(p[1] for p in norm)
    assert_approx(min_x, 0.0, msg=f"Min X should be 0, got {min_x}")
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0, got {min_y}")
    # Vertex count preserved
    assert_equal(len(norm), len(grid_poly),
                 f"Vertex count changed: {len(grid_poly)} -> {len(norm)}")


# =============================================================================
# TEST GROUP: Polygon G-code Y-flip (Bug Fix #1)
# =============================================================================

def _make_simple_polygon_mm(width_mm, height_mm):
    """Create a simple rectangular polygon in mm, already normalized to (0,0)."""
    return [(0, 0), (width_mm, 0), (width_mm, height_mm), (0, height_mm)]


def test_polygon_gcode_no_negative_y_inches():
    """G-code from inch polygon should have no negative Y coordinates."""
    # Simulate a polygon drawn in inches, offset from origin
    grid_poly = [(4, 3), (10, 3), (10, 8), (4, 8)]
    polygon = normalize_polygon_inches(grid_poly)

    settings = get_settings()
    pads = [{"size": 20, "qty": 2}]
    material = "felt"

    poly_w = max(p[0] for p in polygon)
    poly_h = max(p[1] for p in polygon)

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, material, poly_w, poly_h, filepath,
                       hole_dia=3.5, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        if not coords:
            # No pads could be placed (polygon too small) -- that's fine for this test
            return
        for x, y in coords:
            assert_true(y >= -0.01, f"Negative Y coordinate {y} in polygon G-code (inches)")
            assert_true(x >= -0.01, f"Negative X coordinate {x} in polygon G-code (inches)")
    finally:
        os.unlink(filepath)


def test_polygon_gcode_no_negative_y_cm():
    """G-code from cm polygon should have no negative Y coordinates."""
    # Simulate a polygon drawn in cm, offset from origin
    grid_poly = [(10, 8), (25, 8), (25, 22), (10, 22)]
    polygon = normalize_polygon_cm(grid_poly)

    settings = get_settings()
    pads = [{"size": 20, "qty": 2}]
    material = "felt"

    poly_w = max(p[0] for p in polygon)
    poly_h = max(p[1] for p in polygon)

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, material, poly_w, poly_h, filepath,
                       hole_dia=3.5, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        if not coords:
            return
        for x, y in coords:
            assert_true(y >= -0.01, f"Negative Y coordinate {y} in polygon G-code (cm)")
            assert_true(x >= -0.01, f"Negative X coordinate {x} in polygon G-code (cm)")
    finally:
        os.unlink(filepath)


def test_polygon_gcode_yflip_uses_bbox_height():
    """Y-flip should use polygon bounding box height, not sheet_height_mm."""
    # Create a polygon that is much smaller than sheet_height_mm
    polygon = _make_simple_polygon_mm(100, 80)
    settings = get_settings()
    pads = [{"size": 18, "qty": 1}]

    placed, _, _ = _nest_discs(pads, "felt", 100, 80, settings, polygon=polygon)
    if not placed:
        # If nesting fails, skip (polygon might be too small for pad)
        return

    filepath = make_temp_file(".gcode")
    try:
        # Pass a sheet_height_mm that is much larger than the polygon
        # If the bug were present, Y-flip would use 500 instead of 80, giving huge Y values
        generate_gcode_from_placed(placed, "felt", 100, 500, filepath,
                                   hole_dia=3.5, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        for x, y in coords:
            assert_true(y >= -0.01, f"Negative Y in polygon G-code: {y}")
            # Y should not exceed polygon height + some margin for kerf
            assert_true(y <= 85,
                        f"Y coordinate {y} exceeds polygon height 80 -- "
                        "Y-flip may be using sheet_height_mm instead of polygon bbox")
    finally:
        os.unlink(filepath)


def test_polygon_gcode_from_placed_coords():
    """generate_gcode_from_placed with polygon should produce correct Y range."""
    polygon = _make_simple_polygon_mm(120, 100)
    settings = get_settings()

    # Manually create placed discs (as if nesting produced them)
    # cx=60, cy=50 means center of polygon in SVG coords (Y=0 at top)
    placed = [(20.0, 60.0, 50.0, 9.625)]

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode_from_placed(placed, "felt", 120, 100, filepath,
                                   hole_dia=3.5, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        assert_true(len(coords) > 0, "No coordinates found in G-code output")
        for x, y in coords:
            assert_true(y >= -0.01, f"Negative Y: {y}")
            assert_true(x >= -0.01, f"Negative X: {x}")
    finally:
        os.unlink(filepath)


# =============================================================================
# TEST GROUP: Rectangle G-code (Regression)
# =============================================================================

def test_rect_gcode_no_negative_coords():
    """Standard rectangle G-code should have no negative coordinates."""
    settings = get_settings()
    pads = [{"size": 22, "qty": 2}, {"size": 16, "qty": 3}]
    width_mm = 13.5 * 25.4  # ~343mm
    height_mm = 10 * 25.4   # ~254mm

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", width_mm, height_mm, filepath,
                       hole_dia=3.5, settings=settings)
        coords = parse_gcode_coordinates(filepath)
        assert_true(len(coords) > 0, "No coordinates found in rectangle G-code")
        for x, y in coords:
            assert_true(x >= -0.01, f"Negative X in rect G-code: {x}")
            assert_true(y >= -0.01, f"Negative Y in rect G-code: {y}")
    finally:
        os.unlink(filepath)


def test_rect_gcode_footer_correct():
    """Rectangle G-code should have the correct footer structure."""
    settings = get_settings()
    pads = [{"size": 20, "qty": 1}]
    width_mm = 300
    height_mm = 200

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", width_mm, height_mm, filepath,
                       hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        # Find footer: should end with M5 then M2
        non_empty = [l.strip() for l in lines if l.strip()]
        assert_equal(non_empty[-1], "M2", f"Last line should be M2, got: {non_empty[-1]}")
        assert_equal(non_empty[-2], "M5", f"Second-to-last should be M5, got: {non_empty[-2]}")
        # The return move should be before M5
        assert_true("X0Y0" in non_empty[-3],
                     f"Expected X0Y0 return move before M5, got: {non_empty[-3]}")
    finally:
        os.unlink(filepath)


# =============================================================================
# TEST GROUP: G-code Content Validation (All Materials)
# =============================================================================

def _generate_material_gcode(material, pads=None):
    """Helper: generate G-code for a material and return (filepath, lines)."""
    settings = get_settings()
    if pads is None:
        pads = [{"size": 22, "qty": 1}, {"size": 16, "qty": 1}]
    width_mm = 300
    height_mm = 200
    filepath = make_temp_file(".gcode")
    generate_gcode(pads, material, width_mm, height_mm, filepath,
                   hole_dia=3.5, settings=settings)
    lines = parse_gcode_lines(filepath)
    return filepath, lines


def test_gcode_header_felt():
    """Felt G-code should have standard header."""
    filepath, lines = _generate_material_gcode("felt")
    try:
        joined = "\n".join(lines)
        assert_true("G00 G17 G40 G21 G54" in joined, "Missing header line G00 G17 G40 G21 G54")
        assert_true("G90" in joined, "Missing G90 in header")
        assert_true("M4" in joined, "Missing M4 in header")
    finally:
        os.unlink(filepath)


def test_gcode_header_card():
    """Card G-code should have standard header."""
    filepath, lines = _generate_material_gcode("card")
    try:
        joined = "\n".join(lines)
        assert_true("G00 G17 G40 G21 G54" in joined, "Missing header in card G-code")
    finally:
        os.unlink(filepath)


def test_gcode_header_leather():
    """Leather G-code should have standard header."""
    filepath, lines = _generate_material_gcode("leather")
    try:
        joined = "\n".join(lines)
        assert_true("G00 G17 G40 G21 G54" in joined, "Missing header in leather G-code")
    finally:
        os.unlink(filepath)


def test_gcode_layer_names_felt():
    """Felt G-code should use correct layer names: C10 (eng), C09 (hole), C00 (cut)."""
    filepath, lines = _generate_material_gcode("felt")
    try:
        joined = "\n".join(lines)
        # Engraving layer
        assert_true("Layer C10" in joined, "Missing felt engraving layer C10")
        # Hole layer (only if pad >= min_hole_size)
        assert_true("Layer C09" in joined, "Missing felt hole layer C09")
        # Cut layer
        assert_true("Layer C00" in joined, "Missing felt cut layer C00")
    finally:
        os.unlink(filepath)


def test_gcode_layer_names_card():
    """Card G-code should use correct layer names: C15 (eng), C14 (hole), C01 (cut)."""
    filepath, lines = _generate_material_gcode("card")
    try:
        joined = "\n".join(lines)
        assert_true("Layer C15" in joined, "Missing card engraving layer C15")
        assert_true("Layer C14" in joined, "Missing card hole layer C14")
        assert_true("Layer C01" in joined, "Missing card cut layer C01")
    finally:
        os.unlink(filepath)


def test_gcode_layer_names_leather():
    """Leather G-code should use correct layer names: C05 (eng), C03 (hole), C02 (cut)."""
    filepath, lines = _generate_material_gcode("leather")
    try:
        joined = "\n".join(lines)
        assert_true("Layer C05" in joined, "Missing leather engraving layer C05")
        assert_true("Layer C03" in joined, "Missing leather hole layer C03")
        assert_true("Layer C02" in joined, "Missing leather cut layer C02")
    finally:
        os.unlink(filepath)


def test_gcode_s_values_valid_range():
    """All S (power) values in G-code should be in 0-1000 range."""
    s_re = re.compile(r"S(\d+)")
    for material in ["felt", "card", "leather"]:
        filepath, lines = _generate_material_gcode(material)
        try:
            for line in lines:
                for m in s_re.finditer(line):
                    s_val = int(m.group(1))
                    assert_true(0 <= s_val <= 1000,
                                f"S value {s_val} out of range [0,1000] in {material} G-code: {line}")
        finally:
            os.unlink(filepath)


def test_gcode_footer_present_all_materials():
    """All materials should have proper footer with M5 after return move."""
    for material in ["felt", "card", "leather"]:
        filepath, lines = _generate_material_gcode(material)
        try:
            non_empty = [l.strip() for l in lines if l.strip()]
            assert_equal(non_empty[-1], "M2",
                         f"{material}: Last line should be M2, got {non_empty[-1]}")
            assert_equal(non_empty[-2], "M5",
                         f"{material}: M5 should be before M2, got {non_empty[-2]}")
            assert_true("X0Y0" in non_empty[-3],
                         f"{material}: Return move missing before M5, got {non_empty[-3]}")
        finally:
            os.unlink(filepath)


def test_gcode_air_assist_markers():
    """G-code should contain M8 (air on) or M9 (air off) for each layer."""
    filepath, lines = _generate_material_gcode("felt")
    try:
        # With default settings, air assist is on for all layers
        m8_count = sum(1 for l in lines if l.strip() == "M8")
        assert_true(m8_count >= 1, "Expected at least one M8 (air assist on) in felt G-code")
    finally:
        os.unlink(filepath)


# =============================================================================
# TEST GROUP: SVG Polygon Tests
# =============================================================================

def test_svg_polygon_generates_file():
    """SVG generation with polygon should create a non-empty file."""
    settings = get_settings()
    polygon = _make_simple_polygon_mm(200, 150)
    pads = [{"size": 20, "qty": 2}]

    filepath = make_temp_file(".svg")
    try:
        generate_svg(pads, "felt", 200, 150, filepath, 3.5, settings, polygon=polygon)
        assert_true(os.path.exists(filepath), "SVG file not created")
        size = os.path.getsize(filepath)
        assert_true(size > 100, f"SVG file too small ({size} bytes)")
    finally:
        if os.path.exists(filepath):
            os.unlink(filepath)


def test_svg_from_placed_polygon():
    """generate_svg_from_placed should work with polygon."""
    settings = get_settings()
    polygon = _make_simple_polygon_mm(200, 150)

    placed = [(20.0, 50.0, 50.0, 9.625), (18.0, 120.0, 50.0, 8.625)]

    filepath = make_temp_file(".svg")
    try:
        generate_svg_from_placed(placed, "felt", 200, 150, filepath, 3.5, settings, polygon=polygon)
        assert_true(os.path.exists(filepath), "SVG file not created")
        size = os.path.getsize(filepath)
        assert_true(size > 100, f"SVG file too small ({size} bytes)")
    finally:
        if os.path.exists(filepath):
            os.unlink(filepath)


# =============================================================================
# TEST GROUP: Edge Cases
# =============================================================================

def test_empty_pad_list_gcode():
    """Empty pad list should not crash G-code generation."""
    settings = get_settings()
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode([], "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        # File may or may not be written (no pads to place), but should not crash
    finally:
        if os.path.exists(filepath):
            os.unlink(filepath)


def test_single_pad_gcode():
    """Single pad should generate valid G-code."""
    settings = get_settings()
    pads = [{"size": 20, "qty": 1}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        assert_true(len(lines) > 10, f"G-code too short ({len(lines)} lines)")
        coords = parse_gcode_coordinates(filepath)
        assert_true(len(coords) > 0, "No coordinates in single-pad G-code")
    finally:
        os.unlink(filepath)


def test_small_polygon_single_pad():
    """A polygon just big enough for one pad should work."""
    settings = get_settings()
    # A 20mm pad in felt: diameter ~ 20 - 0.75 = 19.25mm, radius ~9.625mm
    # Need polygon slightly bigger: ~22mm square
    polygon = _make_simple_polygon_mm(25, 25)
    pads = [{"size": 20, "qty": 1}]

    placed, _, _ = _nest_discs(pads, "felt", 25, 25, settings, polygon=polygon)
    if not placed:
        # Polygon might be just slightly too small with spacing -- skip
        return

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 25, 25, filepath,
                       hole_dia=0, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        if coords:
            for x, y in coords:
                assert_true(x >= -0.01, f"Negative X in small polygon G-code: {x}")
                assert_true(y >= -0.01, f"Negative Y in small polygon G-code: {y}")
    finally:
        os.unlink(filepath)


def test_polygon_high_y_pre_normalization():
    """Polygon with vertices that would have had very high Y values pre-normalization."""
    # Draw polygon at very top of grid (low Y in grid = high Y in mm after flip)
    grid_poly = [(2, 0.5), (8, 0.5), (8, 4), (2, 4)]
    polygon = normalize_polygon_inches(grid_poly)

    # Verify normalization
    min_y = min(p[1] for p in polygon)
    assert_approx(min_y, 0.0, msg=f"Min Y should be 0 after normalization, got {min_y}")

    settings = get_settings()
    pads = [{"size": 20, "qty": 1}]

    poly_w = max(p[0] for p in polygon)
    poly_h = max(p[1] for p in polygon)

    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", poly_w, poly_h, filepath,
                       hole_dia=3.5, settings=settings, polygon=polygon)
        coords = parse_gcode_coordinates(filepath)
        if coords:
            for x, y in coords:
                assert_true(y >= -0.01, f"Negative Y with high-Y polygon: {y}")
    finally:
        os.unlink(filepath)


def test_nesting_rectangle_basic():
    """Basic rectangle nesting should place pads."""
    settings = get_settings()
    pads = [{"size": 20, "qty": 3}]
    placed, fixed_placed, fixed_total = _nest_discs(pads, "felt", 300, 200, settings)
    assert_equal(fixed_placed, 3, f"Expected 3 pads placed, got {fixed_placed}")
    assert_equal(fixed_total, 3, f"Expected 3 total, got {fixed_total}")
    assert_equal(len(placed), 3, f"Expected 3 placed discs, got {len(placed)}")


def test_nesting_polygon_basic():
    """Basic polygon nesting should place pads."""
    settings = get_settings()
    polygon = _make_simple_polygon_mm(200, 150)
    pads = [{"size": 20, "qty": 2}]
    placed, fixed_placed, fixed_total = _nest_discs(pads, "felt", 200, 150, settings, polygon=polygon)
    assert_true(len(placed) > 0, "No pads placed in polygon nesting")
    assert_equal(fixed_placed, fixed_total, "Not all fixed pads placed")


def test_try_nest_partial_basic():
    """try_nest_partial should return placed, remaining, and any_placed flag."""
    settings = get_settings()
    pads = [{"size": 20, "qty": 5}]
    # Small sheet -- probably can't fit all 5
    placed, remaining, any_placed = try_nest_partial(pads, "felt", 60, 60, settings)
    assert_true(isinstance(placed, list), "placed should be a list")
    assert_true(isinstance(remaining, list), "remaining should be a list")
    assert_true(isinstance(any_placed, bool), "any_placed should be a bool")
    if any_placed:
        assert_true(len(placed) > 0, "any_placed=True but placed is empty")


def test_gcode_cut_grouping_pad_mode():
    """G-code with 'pad' cut grouping should still have correct footer."""
    settings = get_settings()
    settings["gcode_cut_grouping"] = "pad"
    pads = [{"size": 20, "qty": 2}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        non_empty = [l.strip() for l in lines if l.strip()]
        assert_equal(non_empty[-1], "M2", f"pad mode: last should be M2, got {non_empty[-1]}")
        assert_equal(non_empty[-2], "M5", f"pad mode: M5 before M2, got {non_empty[-2]}")
        assert_true("X0Y0" in non_empty[-3],
                     f"pad mode: return move before M5, got {non_empty[-3]}")
    finally:
        os.unlink(filepath)


def test_gcode_cut_grouping_layer_mode():
    """G-code with 'layer' cut grouping should still have correct footer."""
    settings = get_settings()
    settings["gcode_cut_grouping"] = "layer"
    pads = [{"size": 20, "qty": 2}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        non_empty = [l.strip() for l in lines if l.strip()]
        assert_equal(non_empty[-1], "M2", f"layer mode: last should be M2, got {non_empty[-1]}")
        assert_equal(non_empty[-2], "M5", f"layer mode: M5 before M2, got {non_empty[-2]}")
    finally:
        os.unlink(filepath)


def test_gcode_leather_star_cuts():
    """Leather G-code for small pads should produce star/dart cuts without crashing."""
    settings = get_settings()
    settings["darts_enabled"] = True
    settings["dart_threshold"] = 18.0
    # Small pads that trigger star/dart mode
    pads = [{"size": 14, "qty": 2}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "leather", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        assert_true(len(lines) > 10, "Leather star G-code too short")
        coords = parse_gcode_coordinates(filepath)
        assert_true(len(coords) > 0, "No coordinates in leather star G-code")
        for x, y in coords:
            assert_true(x >= -0.5, f"Negative X in leather star G-code: {x}")
            assert_true(y >= -0.5, f"Negative Y in leather star G-code: {y}")
    finally:
        os.unlink(filepath)


def test_gcode_filled_engraving_mode():
    """Filled engraving mode should produce valid G-code."""
    settings = get_settings()
    settings["gcode_settings"]["felt"]["engraving_mode"] = "filled"
    pads = [{"size": 25, "qty": 1}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        assert_true(len(lines) > 10, "Filled engraving G-code too short")
    finally:
        os.unlink(filepath)


def test_gcode_filled_engraving_with_overscan():
    """Filled engraving with overscan should produce valid G-code."""
    settings = get_settings()
    settings["gcode_settings"]["felt"]["engraving_mode"] = "filled"
    settings["filled_overscan_enabled"] = True
    settings["filled_overscan_mm"] = 2.0
    pads = [{"size": 25, "qty": 1}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        assert_true(len(lines) > 10, "Overscan G-code too short")
        # S0 should appear (overscan approach/exit with laser off)
        joined = "\n".join(lines)
        assert_true("S0" in joined, "S0 missing in overscan G-code")
    finally:
        os.unlink(filepath)


def test_gcode_no_hole_small_pads():
    """Small pads below min_hole_size should have no center hole cuts."""
    settings = get_settings()
    settings["min_hole_size"] = 16.5
    # All pads below threshold
    pads = [{"size": 12, "qty": 2}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=3.5, settings=settings)
        lines = parse_gcode_lines(filepath)
        joined = "\n".join(lines)
        # Hole layer for felt is C09
        assert_false("Layer C09" in joined,
                     "Small pads should not have hole layer C09")
    finally:
        os.unlink(filepath)


def test_gcode_zero_hole_diameter():
    """Hole diameter of 0 should produce no hole cuts."""
    settings = get_settings()
    pads = [{"size": 25, "qty": 1}]
    filepath = make_temp_file(".gcode")
    try:
        generate_gcode(pads, "felt", 300, 200, filepath, hole_dia=0, settings=settings)
        lines = parse_gcode_lines(filepath)
        joined = "\n".join(lines)
        assert_false("Layer C09" in joined,
                     "Zero hole_dia should not produce hole layer")
    finally:
        os.unlink(filepath)


def test_polygon_gcode_all_materials():
    """Polygon G-code for all materials should have non-negative coords."""
    polygon = _make_simple_polygon_mm(200, 150)
    for material in ["felt", "card", "leather"]:
        settings = get_settings()
        pads = [{"size": 22, "qty": 1}]
        filepath = make_temp_file(".gcode")
        try:
            generate_gcode(pads, material, 200, 150, filepath,
                           hole_dia=3.5, settings=settings, polygon=polygon)
            coords = parse_gcode_coordinates(filepath)
            if coords:
                for x, y in coords:
                    assert_true(x >= -0.01,
                                f"Negative X ({x}) in {material} polygon G-code")
                    assert_true(y >= -0.01,
                                f"Negative Y ({y}) in {material} polygon G-code")
        finally:
            os.unlink(filepath)


def test_generate_gcode_header_bounds():
    """Header should include bounds comment."""
    lines = generate_gcode_header(5.0, 10.0, 100.0, 200.0)
    joined = "\n".join(lines)
    assert_true("Bounds" in joined, "Header missing bounds comment")
    assert_true("X5.00" in joined, "Header bounds missing X min")
    assert_true("Y10.00" in joined, "Header bounds missing Y min")


def test_generate_gcode_layer_empty():
    """Empty stroke list should produce no output."""
    result = generate_gcode_layer([], 1000, 50, "C00")
    assert_equal(len(result), 0, f"Empty strokes should produce no lines, got {len(result)}")


def test_generate_gcode_layer_single_stroke():
    """Single stroke should produce rapid + cut moves."""
    strokes = [[(10, 20), (30, 40), (50, 60)]]
    result = generate_gcode_layer(strokes, 1000, 50, "TEST")
    joined = "\n".join(result)
    assert_true("G0" in joined, "Missing rapid move G0")
    assert_true("G1" in joined, "Missing cut move G1")
    assert_true("S500" in joined, "Missing S500 (50% * 10)")
    assert_true("Layer TEST" in joined, "Missing layer name comment")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 70)
    print("Stohrer Sax Shop Companion - Bug Fix & Regression Test Suite")
    print("=" * 70)

    sections = [
        ("G-code Footer (Bug Fix #2: M5 after return move)", [
            ("M5 comes after G1 X0Y0 return move", test_footer_m5_after_return),
            ("G1 S0 comes before return move", test_footer_s0_before_return),
            ("M2 is last line", test_footer_m2_last),
            ("Return speed variations", test_footer_return_speed_custom),
            ("M9 air-off in footer", test_footer_m9_air_off),
            ("Complete footer sequence", test_footer_complete_sequence),
        ]),
        ("Polygon Normalization (Bug Fix #3: polygon closing)", [
            ("Normalize inch polygon to origin", test_normalize_polygon_origin_inch),
            ("Normalize cm polygon to origin", test_normalize_polygon_origin_cm),
            ("Shape preserved (inches)", test_normalize_preserves_shape_inches),
            ("Shape preserved (cm)", test_normalize_preserves_shape_cm),
            ("High Y values pre-normalization", test_normalize_polygon_high_y_values),
            ("Polygon at grid origin", test_normalize_polygon_at_grid_origin),
            ("Irregular polygon normalization", test_normalize_irregular_polygon),
        ]),
        ("Polygon G-code Y-flip (Bug Fix #1: negative Y)", [
            ("No negative Y in inch polygon G-code", test_polygon_gcode_no_negative_y_inches),
            ("No negative Y in cm polygon G-code", test_polygon_gcode_no_negative_y_cm),
            ("Y-flip uses polygon bbox not sheet height", test_polygon_gcode_yflip_uses_bbox_height),
            ("generate_gcode_from_placed polygon coords", test_polygon_gcode_from_placed_coords),
            ("Polygon G-code all materials", test_polygon_gcode_all_materials),
        ]),
        ("Rectangle G-code (Regression)", [
            ("No negative coords in rectangle G-code", test_rect_gcode_no_negative_coords),
            ("Correct footer in rectangle G-code", test_rect_gcode_footer_correct),
        ]),
        ("G-code Content Validation (All Materials)", [
            ("Header present - felt", test_gcode_header_felt),
            ("Header present - card", test_gcode_header_card),
            ("Header present - leather", test_gcode_header_leather),
            ("Layer names - felt (C10/C09/C00)", test_gcode_layer_names_felt),
            ("Layer names - card (C15/C14/C01)", test_gcode_layer_names_card),
            ("Layer names - leather (C05/C03/C02)", test_gcode_layer_names_leather),
            ("S values in 0-1000 range", test_gcode_s_values_valid_range),
            ("Footer present all materials", test_gcode_footer_present_all_materials),
            ("Air assist markers", test_gcode_air_assist_markers),
        ]),
        ("SVG Polygon Tests", [
            ("SVG with polygon generates file", test_svg_polygon_generates_file),
            ("SVG from placed with polygon", test_svg_from_placed_polygon),
        ]),
        ("Edge Cases", [
            ("Empty pad list G-code", test_empty_pad_list_gcode),
            ("Single pad G-code", test_single_pad_gcode),
            ("Small polygon single pad", test_small_polygon_single_pad),
            ("High Y pre-normalization polygon G-code", test_polygon_high_y_pre_normalization),
            ("Basic rectangle nesting", test_nesting_rectangle_basic),
            ("Basic polygon nesting", test_nesting_polygon_basic),
            ("try_nest_partial basic", test_try_nest_partial_basic),
            ("Cut grouping: pad mode footer", test_gcode_cut_grouping_pad_mode),
            ("Cut grouping: layer mode footer", test_gcode_cut_grouping_layer_mode),
            ("Leather star/dart cuts", test_gcode_leather_star_cuts),
            ("Filled engraving mode", test_gcode_filled_engraving_mode),
            ("Filled engraving with overscan", test_gcode_filled_engraving_with_overscan),
            ("No hole for small pads", test_gcode_no_hole_small_pads),
            ("Zero hole diameter", test_gcode_zero_hole_diameter),
            ("Header bounds format", test_generate_gcode_header_bounds),
            ("Empty stroke layer", test_generate_gcode_layer_empty),
            ("Single stroke layer", test_generate_gcode_layer_single_stroke),
        ]),
    ]

    for section_name, tests in sections:
        print(f"\n--- {section_name} ---")
        for test_name, test_fn in tests:
            run_test(test_name, test_fn)

    # Summary
    print("\n" + "=" * 70)
    passed = sum(1 for r in _results if r[0] == "PASS")
    failed = sum(1 for r in _results if r[0] == "FAIL")
    errors = sum(1 for r in _results if r[0] == "ERROR")
    total = len(_results)

    print(f"Results: {passed} passed, {failed} failed, {errors} errors out of {total} tests")

    if failed > 0 or errors > 0:
        print("\nFailed/Error tests:")
        for status, name, msg in _results:
            if status in ("FAIL", "ERROR"):
                print(f"  {status}: {name}")
                if msg:
                    print(f"         {msg}")

    print("=" * 70)
    return 1 if (failed + errors) > 0 else 0


if __name__ == "__main__":
    sys.exit(main())
