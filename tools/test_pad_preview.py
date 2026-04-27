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
