"""
Pad preview window test. Exercises the live-preview pane attached to the
Sizing Rules dialog:

  - PadPreviewWindow constructs without throwing
  - Toggling the parent's show_preview_var opens/closes the window
  - Render runs without exception for layered + side-by-side modes
  - Polling tick re-runs cleanly when the parent form changes
  - Closing the preview clears the parent toggle / handle

Runs non-interactively in a withdrawn Tk root. Skips on headless Linux.

Run:
    python tools/test_pad_preview.py
"""
import os
import sys
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

# ui_dialogs evaluates _() at class-definition time (UserGuideWindow.
# SECTION_TITLES), so the gettext _ must be installed before importing it.
from i18n import init_translation  # noqa: E402
init_translation("en")

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
    print("Pad Preview Window Test")
    print("=" * 60)

    import copy
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    from config import DEFAULT_SETTINGS, settings_to_sizing_preset
    from ui_dialogs import OptionsWindow, PadPreviewWindow

    class StubApp:
        def open_resonance_window(self):
            pass

    settings = copy.deepcopy(DEFAULT_SETTINGS)
    presets = {"Default": settings_to_sizing_preset(settings)}

    def make_options():
        w = OptionsWindow(
            root, StubApp(), settings, lambda: None, lambda: None,
            sizing_presets=presets,
            sizing_presets_save_callback=lambda: None,
        )
        w.top.withdraw()
        return w

    def construct_preview():
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            try:
                preview.withdraw()
                # Cancel the polling tick so it doesn't fire after destroy.
                preview._cancel_poll()
                # Force a render — must not raise even before <Configure>
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("PadPreviewWindow constructs and renders cleanly", construct_preview)

    def toggle_opens_and_closes():
        opts = make_options()
        try:
            assert opts.preview_window is None
            opts.show_preview_var.set(True)
            opts._toggle_preview_window()
            assert opts.preview_window is not None
            assert opts.preview_window.winfo_exists()
            # Cancel its poll so we don't get stray after-callbacks.
            opts.preview_window._cancel_poll()
            opts.show_preview_var.set(False)
            opts._toggle_preview_window()
            assert opts.preview_window is None
        finally:
            opts.top.destroy()
    check("Toggle var opens and closes the preview", toggle_opens_and_closes)

    def closing_preview_clears_parent_handle():
        opts = make_options()
        try:
            opts.show_preview_var.set(True)
            opts._toggle_preview_window()
            preview = opts.preview_window
            preview._cancel_poll()
            preview._on_close()
            # Parent toggle reset, handle cleared
            assert opts.show_preview_var.get() is False
            assert opts.preview_window is None
        finally:
            opts.top.destroy()
    check("Preview close clears parent toggle var", closing_preview_clears_parent_handle)

    def render_modes_dont_raise():
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                # Layered with all materials
                preview.layout_var.set('layered')
                for mat in preview.MATERIALS:
                    preview.show_vars[mat].set(True)
                preview._render()
                # Side by side with subset
                preview.layout_var.set('side_by_side')
                preview.show_vars['exact_size'].set(False)
                preview._render()
                # No materials selected → message branch
                for mat in preview.MATERIALS:
                    preview.show_vars[mat].set(False)
                preview._render()
                # Pad size very small (skips dart logic, plain disc)
                preview.show_vars['leather'].set(True)
                preview.preview_size_var.set(8.0)
                preview._render()
                # Pad size large enough that no dart is generated
                preview.preview_size_var.set(40.0)
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Render runs cleanly across modes / sizes / material sets", render_modes_dont_raise)

    def options_close_destroys_preview():
        opts = make_options()
        opts.show_preview_var.set(True)
        opts._toggle_preview_window()
        preview = opts.preview_window
        preview._cancel_poll()
        # Destroying the OptionsWindow should tear down the preview too
        # (via the <Destroy> bind).
        opts.top.destroy()
        root.update()
        assert not preview.winfo_exists()
    check("Closing OptionsWindow destroys an open preview", options_close_destroys_preview)

    def render_with_sizing_range_mode():
        # When sizing rules are in range mode, the preview should still
        # render — both for a pad inside a defined range and one outside
        # (which falls back to universal values).
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                opts.sizing_range_mode_var.set("range")
                opts.sizing_ranges = [{
                    "min_size": 12.0, "max_size": 20.0,
                    "felt_offset": 0.85, "card_to_felt_offset": 0.3,
                    "leather_wrap_multiplier": 1.05, "min_hole_size": 14.0,
                    "felt_thickness": 3.5, "felt_thickness_unit": "mm",
                }]
                # In-range pad
                preview.preview_size_var.set(15.0)
                preview._render()
                # Out-of-range pad (falls back to universal)
                preview.preview_size_var.set(30.0)
                preview._render()
                # Below all ranges
                preview.preview_size_var.set(8.0)
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Preview renders cleanly under sizing-range mode (in-range, out-of-range, below)", render_with_sizing_range_mode)

    def render_with_dart_range_mode():
        # Range-mode darts: pads in a defined range get a dart, pads
        # outside get a plain disc. Either branch must render cleanly.
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                opts.dart_range_mode_var.set("range")
                opts.dart_ranges = [{
                    "min_size": 6.0, "max_size": 11.5,
                    "overwrap": 0.4, "wrap_bonus": 0.5,
                    "frequency_multiplier": 1.2,
                    "shape_factor": 0.0,  # triangle, exercises the math edge
                    "engraving_on": True,
                }]
                # In-range pad → dart
                preview.preview_size_var.set(9.0)
                preview._render()
                # Out-of-range pad → plain disc (no dart)
                preview.preview_size_var.set(15.0)
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Preview renders cleanly under dart-range mode (in-range and out-of-range)", render_with_dart_range_mode)

    def render_at_shape_anchors():
        # Triangle / sine / square anchors should each produce a renderable
        # leather dart shape without exception.
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                preview.preview_size_var.set(12.0)
                for sf in (0.0, 0.5, 1.0, 0.25, 0.75):
                    opts.dart_shape_factor_var.set(sf)
                    preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Preview renders across the dart shape spectrum (triangle/sine/square)", render_at_shape_anchors)

    def render_with_darts_disabled():
        # darts_enabled=False should produce plain leather circles even
        # below the threshold — no dart pattern attempt.
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                opts.darts_enabled_var.set(False)
                preview.preview_size_var.set(10.0)  # would normally dart
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Preview renders cleanly with darts disabled", render_with_darts_disabled)

    def parent_form_change_triggers_rerender():
        opts = make_options()
        try:
            preview = PadPreviewWindow(opts)
            preview.withdraw()
            preview._cancel_poll()
            try:
                # Prime the snapshot, render once.
                preview._render()
                preview._last_form_snapshot = opts._capture_form_to_dict()
                # Change a form var; one poll cycle should detect the diff.
                opts.felt_offset_var.set(opts.felt_offset_var.get() + 0.5)
                # Manually run one poll iteration
                snap = opts._capture_form_to_dict()
                assert snap != preview._last_form_snapshot
                preview._last_form_snapshot = snap
                preview._render()
            finally:
                preview.destroy()
        finally:
            opts.top.destroy()
    check("Parent form edit produces a different snapshot", parent_form_change_triggers_rerender)

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
