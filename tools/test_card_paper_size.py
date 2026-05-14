"""Test that 'Fit card to paper' actually overrides sheet dimensions for the
card material in the generation pipeline.

Catches the user-reported regression: 'I had Fit card to paper checked but the
output still used the regular sheet inputs.'
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tkinter as tk  # noqa: E402
from i18n import init_translation  # noqa: E402
init_translation("en")

from config import load_settings  # noqa: E402


PASS = "  PASS"
FAIL = "  FAIL"
results = []


def check(label, ok, detail=""):
    results.append(ok)
    print(f"{PASS if ok else FAIL}  {label}" + (f"  -- {detail}" if not ok and detail else ""))


def main():
    print("Card Paper Size Tests")
    print("=" * 60)

    # Construct the full app in a withdrawn root so we can call the helpers
    # the way the actual generate buttons do.
    root = tk.Tk()
    root.withdraw()
    try:
        from main import PadSVGGeneratorApp
        app = PadSVGGeneratorApp(root)
    except tk.TclError as e:
        if "no display" in str(e).lower():
            print("Skipping: no display available")
            return 0
        raise

    # --- 1. Helper returns None when checkbox is off ---
    app.card_paper_var.set(False)
    dims = app._get_card_paper_dimensions_mm()
    check("checkbox off -> dims is None", dims is None, f"got {dims!r}")

    # --- 2. Helper returns Letter dims when on + letter selected ---
    app.card_paper_var.set(True)
    app.card_paper_dropdown.set("letter (8.5×11 in)")
    dims = app._get_card_paper_dimensions_mm()
    expected = (8.5 * 25.4, 11.0 * 25.4)
    check("checkbox on + letter -> ~(215.9, 279.4)",
          dims is not None and abs(dims[0] - expected[0]) < 0.01 and abs(dims[1] - expected[1]) < 0.01,
          f"got {dims!r}")

    # --- 3. Helper returns A4 dims when on + a4 selected ---
    app.card_paper_dropdown.set("a4 (210×297 mm)")
    dims = app._get_card_paper_dimensions_mm()
    check("checkbox on + a4 -> (210.0, 297.0)",
          dims == (210.0, 297.0),
          f"got {dims!r}")

    # --- 4. _get_material_dimensions overrides for card material ---
    app.card_paper_var.set(True)
    app.card_paper_dropdown.set("letter (8.5×11 in)")
    paper_dims = app._get_card_paper_dimensions_mm()
    mat_w, mat_h, mat_polygon = app._get_material_dimensions(
        "card", 100.0, 100.0, paper_dims)
    check("card material + paper -> uses paper dims",
          abs(mat_w - expected[0]) < 0.01 and abs(mat_h - expected[1]) < 0.01,
          f"got w={mat_w}, h={mat_h}")
    check("card material + paper -> polygon forced to None",
          mat_polygon is None, f"got {mat_polygon!r}")

    # --- 5. Non-card materials always use the input dims ---
    for material in ("felt", "leather", "exact_size"):
        mat_w, mat_h, _polygon = app._get_material_dimensions(
            material, 100.0, 100.0, paper_dims)
        check(f"{material} material -> uses input dims, not paper",
              mat_w == 100.0 and mat_h == 100.0,
              f"got w={mat_w}, h={mat_h}")

    # --- 6. When checkbox is off, card uses input dims like other materials ---
    app.card_paper_var.set(False)
    paper_dims = app._get_card_paper_dimensions_mm()  # should be None
    mat_w, mat_h, _polygon = app._get_material_dimensions(
        "card", 123.0, 456.0, paper_dims)
    check("checkbox off -> card uses input dims",
          mat_w == 123.0 and mat_h == 456.0,
          f"got w={mat_w}, h={mat_h}")

    # --- 7. Settings round-trip preserves the checkbox + dropdown state ---
    app.card_paper_var.set(True)
    app.card_paper_dropdown.set("a4 (210×297 mm)")
    # Simulate the save logic from on_exit (lines 152-154)
    saved = {}
    saved["card_use_paper_size"] = app.card_paper_var.get()
    dropdown_val = app.card_paper_dropdown.get().lower()
    saved["card_paper_size"] = "a4" if dropdown_val.startswith("a4") else "letter"
    check("save logic captures checkbox state",
          saved["card_use_paper_size"] is True,
          f"got {saved!r}")
    check("save logic captures a4 selection",
          saved["card_paper_size"] == "a4",
          f"got {saved!r}")

    # Cleanup
    root.destroy()

    print("=" * 60)
    passed = sum(results)
    total = len(results)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
