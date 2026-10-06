"""Every SVG the Pad Maker writes must be well-formed XML with a sane root.

The parity suites regex-match generated SVGs and never parse them, so a
malformed file (an unclosed element, a stray attribute, a bad unit string)
would reach LightBurn before any test noticed. This parses each output with
ElementTree and checks the root: width/height in mm matching the sheet, a
viewBox in the same units, and in compatibility mode no unit suffix on any
child attribute.

Covers all four materials, both export modes, labeled zones, leather
locator marks, darted and plain leather, and a traced-polygon sheet.
"""
import copy
import os
import re
import sys
import tempfile
import traceback
import xml.etree.ElementTree as ET

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config  # noqa: E402
import svg_engine as se  # noqa: E402

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


def settings(**over):
    s = copy.deepcopy(config.DEFAULT_SETTINGS)
    s.update(over)
    return s


PADS = [{'size': 7.0, 'qty': 3}, {'size': 12.5, 'qty': 2}, {'size': 20.0, 'qty': 2}, {'size': 35.5, 'qty': 1}]
SCRAP = [(6, 42), (34, 14), (72, 6), (112, 10), (146, 26), (158, 58), (150, 96), (120, 118), (70, 122), (30, 104), (10, 76)]
W, H = 200.0, 150.0


def _render(material, s, hole=3.0, polygon=None, zones=True):
    placed, zone_list = se.nest_pads_with_zones(PADS, material, W, H, s, polygon=polygon)
    assert placed, f"{material}: nothing placed"
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "out.svg")
        se.generate_svg_from_placed(placed, material, W, H, path, hole, s,
                                    polygon=polygon, zones=zone_list if zones else None)
        text = open(path, encoding="utf-8").read()
    return text, ET.fromstring(text)


def _assert_root(root, text, compat):
    tag = root.tag.split('}')[-1]
    assert tag == "svg", f"root is <{tag}>"
    w, h = root.get("width"), root.get("height")
    assert w == f"{W}mm" or w == f"{W:g}mm", f"width={w!r}"
    assert h == f"{H}mm" or h == f"{H:g}mm", f"height={h!r}"
    vb = [float(v) for v in root.get("viewBox").split()]
    assert vb[2] == W and vb[3] == H, f"viewBox={vb}"
    n_children = sum(1 for _ in root.iter()) - 1
    assert n_children > 0, "empty drawing"
    if compat:
        bad = [m for m in re.findall(r'(\w+)="[0-9.]+mm"', text) if m not in ("width", "height")]
        assert not bad, f"unit suffix on child attributes in compatibility mode: {sorted(set(bad))[:5]}"
    else:
        assert 'r="' in text and 'mm"' in text, "default mode should carry mm suffixes"


def _case(material, compat, **over):
    def fn():
        s = settings(compatibility_mode=compat, **over)
        text, root = _render(material, s)
        _assert_root(root, text, compat)
    return fn


for mat in ("felt", "card", "leather", "exact_size"):
    for compat in (False, True):
        check(f"{mat}, {'compatibility' if compat else 'default'} mode: parses, root is sane",
              _case(mat, compat))

check("leather with locator marks (both styles, dashed): parses",
      _case("leather", False, locator_marks_enabled=True, locator_marks_style="both",
            locator_marks_dashed=True, locator_marks_max_size=40.0))
check("leather with locator marks, compatibility mode: unitless",
      _case("leather", True, locator_marks_enabled=True, locator_marks_style="both",
            locator_marks_max_size=40.0))
check("felt with labeled zones on: parses",
      _case("felt", False, zone_labels_enabled=True))
check("felt with labeled zones, compatibility mode: unitless",
      _case("felt", True, zone_labels_enabled=True))
check("leather with darts off (plain circles): parses",
      _case("leather", False, darts_enabled=False))


def polygon_case():
    s = settings(zone_labels_enabled=True)
    text, root = _render("felt", s, polygon=SCRAP)
    tag = root.tag.split('}')[-1]
    assert tag == "svg"
    assert sum(1 for _ in root.iter()) > 1


check("traced-polygon sheet with zones: parses", polygon_case)


def every_element_known():
    s = settings(locator_marks_enabled=True, locator_marks_style="both", locator_marks_max_size=40.0,
                 zone_labels_enabled=True)
    _, root = _render("leather", s)
    allowed = {"svg", "defs", "circle", "path", "text", "line", "polyline", "rect", "polygon", "g"}
    tags = {el.tag.split('}')[-1] for el in root.iter()}
    assert tags <= allowed, f"unexpected elements: {tags - allowed}"


check("only known element types are emitted", every_element_known)


if __name__ == "__main__":
    print("SVG well-formedness")
    print("=" * 60)
    print("=" * 60)
    print(f"{sum(results)}/{len(results)} passed")
    sys.exit(0 if all(results) else 1)
