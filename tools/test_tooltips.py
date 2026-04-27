"""
Tooltip helper test. Exercises the Tooltip class in ui_dialogs:
  - construction stores text and binds Enter/Leave/etc.
  - update_text changes the stored text and (when shown) the live label
  - show/hide round-trip creates and destroys the popup Toplevel
  - destroy of the parent widget cleans up

Runs non-interactively in a withdrawn Tk root. Skips on headless Linux
(same convention as test_smoke_ui).

Run:
    python tools/test_tooltips.py
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
    print("Tooltip Helper Test")
    print("=" * 60)

    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping: no display available ({e})")
        return 0
    root.withdraw()

    from ui_dialogs import (
        Tooltip, add_tooltip, add_tooltips,
        set_tooltips_enabled, tooltips_enabled,
    )

    label = tk.Label(root, text="hover me")
    label.pack()

    def construct():
        tip = Tooltip(label, "first text", delay_ms=10, wraplength=100)
        assert tip.text == "first text"
        assert tip.delay_ms == 10
        assert tip.wraplength == 100
        # bind data should now contain Enter / Leave handlers
        bindings = label.bind()
        assert "<Enter>" in bindings
        assert "<Leave>" in bindings
    check("Construct Tooltip and bind events", construct)

    def update_text_no_popup():
        tip = Tooltip(label, "before")
        tip.update_text("after")
        assert tip.text == "after"
    check("update_text stores new text when popup not shown", update_text_no_popup)

    def show_and_hide():
        tip = Tooltip(label, "shown text", delay_ms=10)
        tip._show()  # bypass the after-delay path
        assert tip._tip is not None and isinstance(tip._tip, tk.Toplevel)
        # The label inside the popup should hold the tooltip text
        assert tip._label.cget("text") == "shown text"
        tip._hide()
        assert tip._tip is None
        assert tip._label is None
    check("Show creates popup, hide tears it down", show_and_hide)

    def update_text_while_shown():
        tip = Tooltip(label, "before")
        tip._show()
        tip.update_text("after")
        assert tip._label.cget("text") == "after"
        tip._hide()
    check("update_text mutates live popup label", update_text_while_shown)

    def add_tooltip_helper():
        btn = tk.Button(root, text="hi")
        tip = add_tooltip(btn, "convenience")
        assert isinstance(tip, Tooltip)
        assert tip.text == "convenience"
        btn.destroy()
    check("add_tooltip convenience returns Tooltip", add_tooltip_helper)

    def add_tooltips_helper():
        a = tk.Label(root, text="a")
        b = tk.Label(root, text="b")
        c = tk.Label(root, text="c")
        tips = add_tooltips("shared", a, b, c)
        assert len(tips) == 3
        for t, w in zip(tips, (a, b, c)):
            assert t.text == "shared"
            assert t.widget is w
        for w in (a, b, c):
            w.destroy()
    check("add_tooltips attaches one text to many widgets", add_tooltips_helper)

    def global_disable_blocks_show():
        # Disabling globally should make _show a no-op even when text is set.
        try:
            assert tooltips_enabled() is True
            set_tooltips_enabled(False)
            assert tooltips_enabled() is False
            tip = Tooltip(label, "should not appear")
            tip._show()
            assert tip._tip is None, "tooltip popped despite global disable"
        finally:
            set_tooltips_enabled(True)
            assert tooltips_enabled() is True
    check("Global disable blocks Tooltip._show", global_disable_blocks_show)

    def widget_destroy_cleans_up():
        scratch = tk.Label(root, text="scratch")
        tip = Tooltip(scratch, "doomed")
        tip._show()
        scratch.destroy()
        # _on_destroy is bound and should hide the popup
        # It runs synchronously via the <Destroy> event during destroy()
        root.update_idletasks()
        assert tip._tip is None
    check("Destroying parent widget hides tooltip popup", widget_destroy_cleans_up)

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
