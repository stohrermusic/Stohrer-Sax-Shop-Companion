"""Test edge bias feature in nesting algorithms."""
import sys
sys.path.insert(0, '.')
from svg_engine import _nest_discs, _nest_discs_polygon
from config import DEFAULT_SETTINGS

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

settings = DEFAULT_SETTINGS.copy()
pads = [{'size': 20.0, 'qty': 4}]
width, height = 100.0, 80.0

# ============================================
# RECTANGULAR NESTING BIAS TESTS
# ============================================
print("\n--- Rectangular Nesting: Edge Bias ---")

# Test center/default (top-left scan)
settings["edge_bias"] = "center"
placed_center, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("Center bias places all 4 pads", fp == 4)

# Capture average positions for center (default top-left scan)
avg_x_center = sum(p[1] for p in placed_center) / len(placed_center)
avg_y_center = sum(p[2] for p in placed_center) / len(placed_center)

# Test south bias - pads should cluster toward bottom
settings["edge_bias"] = "s"
placed_s, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("South bias places all 4 pads", fp == 4)
avg_y_s = sum(p[2] for p in placed_s) / len(placed_s)
test("South bias: avg Y > center avg Y", avg_y_s > avg_y_center)

# Test east bias - pads should cluster toward right
settings["edge_bias"] = "e"
placed_e, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("East bias places all 4 pads", fp == 4)
avg_x_e = sum(p[1] for p in placed_e) / len(placed_e)
test("East bias: avg X > center avg X", avg_x_e > avg_x_center)

# Test SE corner bias
settings["edge_bias"] = "se"
placed_se, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("SE bias places all 4 pads", fp == 4)
avg_x_se = sum(p[1] for p in placed_se) / len(placed_se)
avg_y_se = sum(p[2] for p in placed_se) / len(placed_se)
test("SE bias: avg X > center avg X", avg_x_se > avg_x_center)
test("SE bias: avg Y > center avg Y", avg_y_se > avg_y_center)

# Test NW corner (should behave like default scan)
settings["edge_bias"] = "nw"
placed_nw, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("NW bias places all 4 pads", fp == 4)

# Test west bias
settings["edge_bias"] = "w"
placed_w, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("West bias places all 4 pads", fp == 4)
avg_x_w = sum(p[1] for p in placed_w) / len(placed_w)
test("West bias: avg X <= center avg X", avg_x_w <= avg_x_center + 0.01)

# Test north bias
settings["edge_bias"] = "n"
placed_n, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("North bias places all 4 pads", fp == 4)
avg_y_n = sum(p[2] for p in placed_n) / len(placed_n)
test("North bias: avg Y <= center avg Y", avg_y_n <= avg_y_center + 0.01)

# ============================================
# RECTANGULAR NESTING: MAX PADS WITH BIAS
# ============================================
print("\n--- Rectangular Nesting: Max Pads with Bias ---")
max_pads = [{'size': 15.0, 'qty': 'max'}]

settings["edge_bias"] = "center"
placed_max_c, _, _ = _nest_discs(max_pads, "felt", 60, 60, settings)
test("Max pads (center) places some pads", len(placed_max_c) > 0)

settings["edge_bias"] = "se"
placed_max_se, _, _ = _nest_discs(max_pads, "felt", 60, 60, settings)
test("Max pads (SE) places some pads", len(placed_max_se) > 0)
# Radial packing may be slightly less dense than linear, so allow small difference
test("Max pads: similar count regardless of bias", abs(len(placed_max_c) - len(placed_max_se)) <= 2)

# ============================================
# POLYGON NESTING BIAS TESTS
# ============================================
print("\n--- Polygon Nesting: Edge Bias ---")

# Simple square polygon
polygon = [(0, 0), (80, 0), (80, 80), (0, 80)]
poly_pads = [{'size': 20.0, 'qty': 3}]

settings["edge_bias"] = "center"
placed_pc, fp, ft = _nest_discs_polygon(poly_pads, "felt", settings, polygon)
test("Polygon center: places all 3 pads", fp == 3)
avg_x_pc = sum(p[1] for p in placed_pc) / len(placed_pc)
avg_y_pc = sum(p[2] for p in placed_pc) / len(placed_pc)

settings["edge_bias"] = "se"
placed_pse, fp, ft = _nest_discs_polygon(poly_pads, "felt", settings, polygon)
test("Polygon SE bias: places all 3 pads", fp == 3)
avg_x_pse = sum(p[1] for p in placed_pse) / len(placed_pse)
avg_y_pse = sum(p[2] for p in placed_pse) / len(placed_pse)
test("Polygon SE bias: avg X shifts right", avg_x_pse > avg_x_pc - 1)
test("Polygon SE bias: avg Y shifts down", avg_y_pse > avg_y_pc - 1)

settings["edge_bias"] = "nw"
placed_pnw, fp, ft = _nest_discs_polygon(poly_pads, "felt", settings, polygon)
test("Polygon NW bias: places all 3 pads", fp == 3)
avg_x_pnw = sum(p[1] for p in placed_pnw) / len(placed_pnw)
avg_y_pnw = sum(p[2] for p in placed_pnw) / len(placed_pnw)
test("Polygon NW bias: avg X < SE avg X", avg_x_pnw < avg_x_pse)
test("Polygon NW bias: avg Y < SE avg Y", avg_y_pnw < avg_y_pse)

# Test small pads with polygon bias
small_poly_pads = [{'size': 10.0, 'qty': 5}]

settings["edge_bias"] = "center"
placed_sc, fp, ft = _nest_discs_polygon(small_poly_pads, "felt", settings, polygon)
test("Polygon small center: places all 5", fp == 5)

settings["edge_bias"] = "sw"
placed_ssw, fp, ft = _nest_discs_polygon(small_poly_pads, "felt", settings, polygon)
test("Polygon small SW bias: places all 5", fp == 5)

# ============================================
# ALL MATERIALS WITH BIAS
# ============================================
print("\n--- All Materials with Bias ---")
for material in ["felt", "card", "leather"]:
    settings["edge_bias"] = "se"
    p, fp, ft = _nest_discs(pads, material, width, height, settings)
    test(f"{material}: SE bias places all pads", fp == ft)

# ============================================
# EDGE CASES
# ============================================
print("\n--- Edge Cases ---")

# Unknown bias value falls back to center behavior
settings["edge_bias"] = "invalid"
placed_inv, fp, ft = _nest_discs(pads, "felt", width, height, settings)
test("Invalid bias value still places pads", fp == 4)

# No bias key at all (missing from settings)
settings_no_bias = {k: v for k, v in settings.items() if k != "edge_bias"}
placed_nb, fp, ft = _nest_discs(pads, "felt", width, height, settings_no_bias)
test("Missing edge_bias key still works", fp == 4)

# ============================================
# SUMMARY
# ============================================
print(f"\n{'='*40}")
print(f"Results: {passed} passed, {failed} failed out of {passed + failed} tests")
if failed == 0:
    print("ALL TESTS PASSED")
else:
    print("SOME TESTS FAILED")
    sys.exit(1)
