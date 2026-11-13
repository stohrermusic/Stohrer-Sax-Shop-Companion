"""What every dialog shares: the dialog background, the tooltip helper, the mouse-wheel binding, the confirmation box, the material names and the G-code file line.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
import logging
import sys



# On macOS, use native system colors for dark/light mode support.
# On Windows/Linux, use our custom cream color for dialogs.
IS_MACOS = sys.platform == 'darwin'
DIALOG_BG = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"

# Camera calibration and laser-control diagnostics. config.setup_logging
# lets this logger write INFO to app.log (the rest of the app logs WARNING
# and up), so what happened on a user's machine can be read afterwards.
mlog = logging.getLogger("ssc.machine")
log = logging.getLogger(__name__)

def material_display_name(material):
    """A pad material key as the Pad Maker's Materials checkboxes show it."""
    if material == 'adhesive_ring':
        return _("Adhesive rings")
    return _(material.replace('_', ' ').capitalize())


# ==========================================
# CROSS-PLATFORM SCROLL HELPER
# ==========================================

def bind_mousewheel(widget, canvas):
    """Bind mousewheel scrolling to a canvas, cross-platform."""
    def _on_mousewheel(event):
        if sys.platform == 'darwin':
            # macOS: delta is usually 1 or -1
            canvas.yview_scroll(int(-1 * event.delta), "units")
        else:
            # Windows: delta is 120 per wheel notch, but precision
            # touchpads send smaller deltas (e.g. ±30) that would
            # truncate to zero — guarantee at least one scroll unit.
            steps = int(-1 * (event.delta / 120))
            if steps == 0 and event.delta != 0:
                steps = -1 if event.delta > 0 else 1
            canvas.yview_scroll(steps, "units")

    def _on_mousewheel_linux(event):
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    if sys.platform == 'linux':
        widget.bind('<Button-4>', _on_mousewheel_linux)
        widget.bind('<Button-5>', _on_mousewheel_linux)
    else:
        widget.bind('<MouseWheel>', _on_mousewheel)

# ==========================================
# TOOLTIP HELPER
# ==========================================

# Module-level on/off so the Feature Set dialog can disable tooltips at
# runtime without us having to walk every existing widget. Tooltip._show
# checks this before constructing its popup; existing bindings stay in
# place but become no-ops while disabled.
_TOOLTIPS_ENABLED = True


def set_tooltips_enabled(enabled):
    """Globally enable or disable tooltip popups. Applies immediately."""
    global _TOOLTIPS_ENABLED
    _TOOLTIPS_ENABLED = bool(enabled)


def tooltips_enabled():
    return _TOOLTIPS_ENABLED


class Tooltip:
    """Lightweight hover-to-explain tooltip for any tk widget.

    Shows a small borderless popup near the cursor after a short hover delay,
    hides on leave / click / focus-out. Safe across Windows/macOS/Linux:
    uses an overrideredirect Toplevel and avoids grabbing focus.
    """

    _BG = "#FFFFE0"  # pale yellow, readable on every platform
    _FG = "#000000"
    _BORDER = "#7A7A7A"
    DELAY_MS = 500
    WRAPLENGTH = 360

    def __init__(self, widget, text, delay_ms=None, wraplength=None):
        self.widget = widget
        self.text = text
        self.delay_ms = self.DELAY_MS if delay_ms is None else delay_ms
        self.wraplength = self.WRAPLENGTH if wraplength is None else wraplength
        self._after_id = None
        self._tip = None
        self._label = None

        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")
        widget.bind("<FocusOut>", self._hide, add="+")
        widget.bind("<Destroy>", self._on_destroy, add="+")

    def update_text(self, text):
        self.text = text
        if self._tip and self._label is not None:
            try:
                self._label.configure(text=text)
            except tk.TclError:
                pass

    def _schedule(self, _event=None):
        self._cancel()
        self._after_id = self.widget.after(self.delay_ms, self._show)

    def _cancel(self):
        if self._after_id is not None:
            try:
                self.widget.after_cancel(self._after_id)
            except (tk.TclError, ValueError):
                pass
            self._after_id = None

    def _show(self):
        if self._tip is not None or not self.text:
            return
        if not _TOOLTIPS_ENABLED:
            return
        try:
            x = self.widget.winfo_pointerx() + 14
            y = self.widget.winfo_pointery() + 18
        except tk.TclError:
            return
        try:
            tip = tk.Toplevel(self.widget)
        except tk.TclError:
            return
        tip.wm_overrideredirect(True)
        # Keep tooltip out of the taskbar / above grabbed dialogs.
        try:
            tip.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        tip.configure(bg=self._BORDER)
        self._label = tk.Label(
            tip, text=self.text,
            bg=self._BG, fg=self._FG,
            justify="left", relief="flat",
            wraplength=self.wraplength,
            padx=6, pady=3,
            font=("Helvetica", 9),
        )
        self._label.pack(padx=1, pady=1)
        tip.update_idletasks()
        # Nudge back on screen if we'd overflow the right/bottom edge.
        try:
            sw = self.widget.winfo_screenwidth()
            sh = self.widget.winfo_screenheight()
            tw = tip.winfo_reqwidth()
            th = tip.winfo_reqheight()
            if x + tw > sw:
                x = max(0, sw - tw - 4)
            if y + th > sh:
                y = max(0, sh - th - 4)
        except tk.TclError:
            pass
        tip.wm_geometry(f"+{x}+{y}")
        self._tip = tip

    def _hide(self, _event=None):
        self._cancel()
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None
            self._label = None

    def _on_destroy(self, _event=None):
        self._hide()


def gcode_file_note():
    """The line every "G-code written to a file" message box ends with.

    A user imported die insert G-code into LightBurn, which drops the
    speeds and powers the file carries, and was confused by what came out
    (2026-11-09). Matt: "is there a dialog when the user generates gcode?
    maybe we can put something in there (this is not for lightburn)", and
    "short message for gcode is best". One string, defined here once, used
    by the Pad Maker and every Tooling tool; never on an SVG or STL
    message, and not on Frame & Cut, which writes no file. The why is in
    the user guide. A function, so it follows the app's language.
    """
    return _("Send this file to the laser as it is. For LightBurn, use Generate SVG.")


def add_tooltip(widget, text, **kwargs):
    """Attach a Tooltip to `widget`. Returns the Tooltip for further tweaking."""
    return Tooltip(widget, text, **kwargs)


def add_tooltips(text, *widgets, **kwargs):
    """Attach the same tooltip text to multiple widgets in one call."""
    return [Tooltip(w, text, **kwargs) for w in widgets]


# ==========================================
# HELPER UTILS
# ==========================================

def get_unique_name(name, existing_keys):
    """
    Returns 'name' if it doesn't exist in existing_keys.
    Otherwise returns 'name (2)', 'name (3)', etc.
    """
    if name not in existing_keys:
        return name
    
    i = 2
    new_name = f"{name} ({i})"
    while new_name in existing_keys:
        i += 1
        new_name = f"{name} ({i})"
    return new_name

# ==========================================
# DIALOG CLASSES
# ==========================================

class ConfirmationDialog(tk.Toplevel):
    def __init__(self, parent, title, message):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.minsize(450, 150)

        self.result = False
        self.dont_show_again = tk.BooleanVar()

        tk.Label(self, text=message, wraplength=430, bg=DIALOG_BG, justify="left").pack(padx=10, pady=10)

        checkbox_frame = tk.Frame(self, bg=DIALOG_BG)
        checkbox_frame.pack(pady=5)
        tk.Checkbutton(checkbox_frame, text=_("Don't show this message again"), variable=self.dont_show_again, bg=DIALOG_BG).pack()

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=10)
        tk.Button(button_frame, text=_("Yes, Proceed"), command=self.on_yes).pack(side="left", padx=10)
        tk.Button(button_frame, text=_("No, Cancel"), command=self.on_no).pack(side="left", padx=10)

        # Let the window size itself to fit the content
        self.update_idletasks()
        self.geometry("")

        self.protocol("WM_DELETE_WINDOW", self.on_no)
        self.wait_window(self)

    def on_yes(self):
        self.result = True
        self.destroy()

    def on_no(self):
        self.result = False
        self.destroy()


# Material colors for preview
_PREVIEW_COLORS = {
    'felt': '#4488CC',
    'card': '#CC8844',
    'leather': '#886633',
    'exact_size': '#66AA66',
    'adhesive_ring': '#B02FB0',
}
