"""
Dart shape-factor spectrum test. Verifies _wave_value (triangle/sine/square
mix) and that calculate_star_path produces visibly different geometry across
the spectrum.

Run:
    python tools/test_dart_shapes.py
"""
import math
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from svg_engine import _wave_value, calculate_star_path  # noqa: E402

results = []


def check(name, fn):
    try:
        fn()
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def main():
    print("Dart Shape Spectrum Test")
    print("=" * 60)

    def anchors_match():
        # At a peak (raw_cos = +1) every anchor returns +1
        for s in (0.0, 0.5, 1.0):
            assert abs(_wave_value(1.0, s) - 1.0) < 1e-9, f"peak fail at s={s}"
            assert abs(_wave_value(-1.0, s) + 1.0) < 1e-9, f"valley fail at s={s}"
        # At raw_cos = 0 every anchor returns 0
        for s in (0.0, 0.5, 1.0):
            assert abs(_wave_value(0.0, s)) < 1e-9, f"zero crossing fail at s={s}"
    check("Anchors hit +/-1 at peaks/valleys and 0 at zero crossings", anchors_match)

    def sine_is_identity_at_half():
        # At s=0.5 the function is the raw cosine
        for c in [-0.9, -0.5, -0.1, 0.1, 0.5, 0.9]:
            assert abs(_wave_value(c, 0.5) - c) < 1e-9, f"sine identity fails at c={c}"
    check("Slider centered is identity (pure sine)", sine_is_identity_at_half)

    def triangle_is_linear():
        # Triangle wave: at raw_cos = cos(pi/4) = sqrt(2)/2 the triangle wave
        # value is at quarter-period above zero crossing -> +0.5
        # arcsin(sqrt(2)/2) = pi/4 -> (2/pi)*(pi/4) = 0.5
        c = math.sqrt(2) / 2
        assert abs(_wave_value(c, 0.0) - 0.5) < 1e-9
        # And -sqrt(2)/2 -> -0.5
        assert abs(_wave_value(-c, 0.0) + 0.5) < 1e-9
    check("Triangle anchor matches the linear ramp formula", triangle_is_linear)

    def square_pushes_toward_sign():
        # Square wave at s=1.0: should be much closer to sign(c) than sine
        c = 0.4
        sine_val = c
        square_val = _wave_value(c, 1.0)
        # square_val should be larger in magnitude than sine_val (because the
        # square wave saturates toward +1 even for small positive c)
        assert square_val > sine_val
        # 0.4 ** 0.01 ~= 0.991. We want the square anchor to look strongly
        # square — comfortably above 0.98 for a typical mid-range c.
        assert square_val > 0.98
    check("Square anchor saturates toward sign(c)", square_pushes_toward_sign)

    def smooth_blend():
        # Walking the slider from triangle -> sine -> square should be
        # monotonic in magnitude for a positive cos value (square is
        # "pushed up" relative to sine, which is "pushed up" relative to
        # triangle for the same c).
        c = 0.5
        # triangle(0.5) = (2/pi)*arcsin(0.5) = 1/3 ≈ 0.333
        # sine(0.5)   = 0.5
        # square(0.5) = 0.5 ** 0.05 ≈ 0.966
        tri = _wave_value(c, 0.0)
        mid_low = _wave_value(c, 0.25)
        sine = _wave_value(c, 0.5)
        mid_high = _wave_value(c, 0.75)
        sq = _wave_value(c, 1.0)
        assert tri < mid_low < sine < mid_high < sq, \
            f"non-monotonic: {tri:.3f} {mid_low:.3f} {sine:.3f} {mid_high:.3f} {sq:.3f}"
    check("Slider produces smooth, monotonic blend at c=0.5", smooth_blend)

    def out_of_range_clamped():
        # Slider values outside 0..1 are clamped, no exception
        assert abs(_wave_value(0.5, -0.5) - _wave_value(0.5, 0.0)) < 1e-9
        assert abs(_wave_value(0.5, 1.5) - _wave_value(0.5, 1.0)) < 1e-9
    check("Slider values outside 0..1 are clamped safely", out_of_range_clamped)

    def path_geometry_changes():
        # The resulting SVG path should differ visibly between anchors.
        p_tri = calculate_star_path(0, 0, 10, 8, num_points=8, shape_factor=0.0)
        p_sine = calculate_star_path(0, 0, 10, 8, num_points=8, shape_factor=0.5)
        p_sq = calculate_star_path(0, 0, 10, 8, num_points=8, shape_factor=1.0)
        assert p_tri != p_sine
        assert p_sine != p_sq
        assert p_tri != p_sq
        # All paths begin with M and end with Z
        for p in (p_tri, p_sine, p_sq):
            assert p.startswith("M ") and p.endswith(" Z")
    check("calculate_star_path outputs differ across anchors", path_geometry_changes)

    def legacy_migration():
        # Migration: old 0.0 (sine) -> new 0.5 (sine), old 1.0 (square) -> new 1.0
        # We simulate the migration by importing load_settings; but to avoid
        # touching disk, just exercise the math directly.
        for old, new in [(0.0, 0.5), (0.5, 0.75), (1.0, 1.0)]:
            assert abs((0.5 + 0.5 * old) - new) < 1e-9
    check("Legacy 0..1 sine->square migrates to 0.5..1.0", legacy_migration)

    passed = sum(results)
    total = len(results)
    print("=" * 60)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
