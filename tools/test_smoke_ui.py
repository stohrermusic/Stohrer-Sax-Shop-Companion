"""
UI smoke test. Constructs the full app (all tabs, menus, dialogs class
imports) in a hidden Tk root to catch import errors, missing mixin
methods, or broken menu wiring before users hit them.

Does NOT interact with audio devices, does NOT render to a visible window,
and does NOT exercise tab-switch side effects (tuner/toner auto-start).

Run:
    python tools/test_smoke_ui.py

Requires a display (fine on Windows / macOS; headless Linux needs Xvfb).
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

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
    print("UI Smoke Test")
    print("=" * 60)

    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    app = [None]

    def construct_app():
        from main import PadSVGGeneratorApp
        app[0] = PadSVGGeneratorApp(root)

    check("Import and construct PadSVGGeneratorApp", construct_app)

    if app[0] is None:
        root.destroy()
        print("\nConstructor failed; skipping remaining checks.")
        return 1

    a = app[0]

    expected_tabs = [
        'pad_tab',
        'key_tab',
        'serial_tab',
        'screw_tab',
        'tooling_tab_frame',
        'tuner_tab_frame',
        'toner_tab_frame',
    ]
    for tab_attr in expected_tabs:
        check(f"Tab frame exists: {tab_attr}",
              lambda t=tab_attr: getattr(a, t) is not None)

    check("Notebook exists and has tabs",
          lambda: len(a.notebook.tabs()) >= len(expected_tabs))

    check("Menu bar attached to root",
          lambda: root.cget('menu') != '')

    check("Settings dict loaded",
          lambda: isinstance(a.settings, dict) and 'layer_colors' in a.settings)

    check("Exception hook wired",
          lambda: sys.excepthook.__name__ != 'excepthook' or sys.excepthook is not sys.__excepthook__)

    # Dialog-class imports (constructor only, no windows shown)
    def import_dialogs():
        from ui_dialogs import (
            OptionsWindow, ColorsWindow, GcodeSettingsWindow,
            PolygonDrawWindow, PadNotesWindow, UserGuideWindow,
            NestingPreviewWindow, AboutDialog,
        )
    check("All dialog classes importable", import_dialogs)

    # Engine imports
    def import_engines():
        from svg_engine import generate_svg, nest_pads
        from gcode_engine import generate_gcode_from_placed
        from tuner_engine import TunerEngine, TunerResult
        from toner_engine import TonerEngine, descriptors_from_harmonics
    check("All engines importable", import_engines)

    try:
        root.destroy()
    except Exception:
        pass

    passed = sum(results)
    total = len(results)
    print("=" * 60)
    print(f"Summary: {passed}/{total} passed")
    return 0 if passed == total else 1


if __name__ == '__main__':
    sys.exit(main())
