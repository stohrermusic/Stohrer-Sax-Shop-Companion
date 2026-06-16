import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import copy
import math
import random
import json
import sys
import time

from config import (
    DEFAULT_SETTINGS, LIGHTBURN_COLORS, get_resonance_messages,
    ALL_KEY_HEIGHT_FIELDS, save_settings, save_presets,
    SIZING_PRESET_KEYS,
    APP_VERSION, APP_BUILD_DATE,
    get_dart_settings_for_size, get_sizing_for_size,
)
from svg_engine import (
    get_disc_diameter, get_felt_thickness_mm, _wave_value,
    feeds_speeds_label_geometry,
)

# On macOS, use native system colors for dark/light mode support.
# On Windows/Linux, use our custom cream color for dialogs.
IS_MACOS = sys.platform == 'darwin'
DIALOG_BG = "systemWindowBackgroundColor" if IS_MACOS else "#F0EAD6"

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

class OptionsWindow:
    def __init__(self, parent, app, settings, update_callback, save_callback,
                 sizing_presets=None, sizing_presets_save_callback=None):
        self.app = app
        self.settings = settings
        self.update_callback = update_callback
        self.save_callback = save_callback
        # Sizing-rules preset library: {preset_name: {settings keys...}}.
        # Mutated in place; persisted via sizing_presets_save_callback.
        self.sizing_presets = sizing_presets if sizing_presets is not None else {}
        self.sizing_presets_save_callback = sizing_presets_save_callback or (lambda: None)

        self.top = tk.Toplevel(parent)
        self.top.title(_("Sizing Rules"))
        self.top.geometry("500x750")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)

        apply_btn = tk.Button(bottom_button_frame, text=_("Apply"), command=self.on_apply)
        apply_btn.pack(side="left", padx=5)
        cancel_btn = tk.Button(bottom_button_frame, text=_("Cancel"), command=self.on_cancel)
        cancel_btn.pack(side="left", padx=5)
        add_tooltip(apply_btn,
                    _("Apply the values in this dialog to the running app. "
                    "If you have edits not yet captured in a preset, "
                    "you'll be asked to save them as a preset first."))
        add_tooltip(cancel_btn,
                    _("Close without applying. If you have unsaved edits, "
                    "you'll be asked whether to save them as a preset."))
        # Override window-close (X) so it goes through the same dirty-check.
        self.top.protocol("WM_DELETE_WINDOW", self.on_cancel)

        # When the OptionsWindow goes away (any path), tear down the
        # preview Toplevel too so it doesn't stay floating with stale data.
        def _on_top_destroy(event):
            if event.widget is self.top:
                self._close_preview_if_open()
        self.top.bind("<Destroy>", _on_top_destroy, add="+")

        if not IS_MACOS:
            adv_btn = tk.Button(bottom_button_frame, text=_("Advanced"), command=self.app.open_resonance_window)
            adv_btn.pack(side="right", padx=5)
            add_tooltip(adv_btn, _("Hidden corner."))
        revert_btn = tk.Button(bottom_button_frame, text=_("Revert to Defaults"), command=self.revert_to_defaults)
        revert_btn.pack(side="right", padx=5)
        add_tooltip(revert_btn,
                    _("Reset every value in this dialog back to the app's "
                    "factory defaults (use with caution — this clears your "
                    "current settings here)."))

        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True)

        self.canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG, padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, self.canvas)

        # --- Sizing variables ---
        self.unit_var = tk.StringVar(value=self.settings["units"])
        self.felt_offset_var = tk.DoubleVar(value=self.settings["felt_offset"])
        self.card_offset_var = tk.DoubleVar(value=self.settings["card_to_felt_offset"])
        self.leather_mult_var = tk.DoubleVar(value=self.settings["leather_wrap_multiplier"])
        self.min_hole_size_var = tk.DoubleVar(value=self.settings["min_hole_size"])
        self.felt_thickness_var = tk.DoubleVar(value=self.settings["felt_thickness"])
        self.felt_thickness_unit_var = tk.StringVar(value=self.settings["felt_thickness_unit"])
        
        # DART VARS (universal mode)
        self.darts_enabled_var = tk.BooleanVar(value=self.settings.get("darts_enabled", True))
        self.dart_threshold_var = tk.DoubleVar(value=self.settings.get("dart_threshold", 18.0))
        self.dart_overwrap_var = tk.DoubleVar(value=self.settings.get("dart_overwrap", 0.5))
        self.dart_wrap_bonus_var = tk.DoubleVar(value=self.settings.get("dart_wrap_bonus", 0.75))
        self.dart_frequency_multiplier_var = tk.DoubleVar(value=self.settings.get("dart_frequency_multiplier", 1.0))
        self.dart_shape_factor_var = tk.DoubleVar(value=self.settings.get("dart_shape_factor", 0.5))

        # Dart range mode vars
        self.dart_range_mode_var = tk.StringVar(value=self.settings.get("dart_range_mode", "universal"))
        self.dart_ranges = list(self.settings.get("dart_ranges", []))
        self.selected_range_index = None

        # Range editing vars
        self.range_min_var = tk.DoubleVar(value=0.0)
        self.range_max_var = tk.DoubleVar(value=18.0)
        self.range_overwrap_var = tk.DoubleVar(value=0.5)
        self.range_wrap_bonus_var = tk.DoubleVar(value=0.75)
        self.range_freq_mult_var = tk.DoubleVar(value=1.0)
        self.range_shape_factor_var = tk.DoubleVar(value=0.5)
        self.range_engraving_on_var = tk.BooleanVar(value=True)

        self.engraving_on_var = tk.BooleanVar(value=self.settings["engraving_on"])
        self.compatibility_mode_var = tk.BooleanVar(value=self.settings.get("compatibility_mode", False))
        self.engraving_font_size_vars = {}
        self.engraving_loc_vars = {}

        # Dart Engraving Vars (universal mode)
        self.dart_engraving_on_var = tk.BooleanVar(value=self.settings.get("dart_engraving_on", True))

        # Sizing range mode vars
        self.sizing_range_mode_var = tk.StringVar(value=self.settings.get("sizing_range_mode", "universal"))
        self.sizing_ranges = list(self.settings.get("sizing_ranges", []))
        self.sizing_selected_range_index = None
        self.sizing_range_min_var = tk.DoubleVar(value=0.0)
        self.sizing_range_max_var = tk.DoubleVar(value=60.0)
        self.sizing_range_felt_offset_var = tk.DoubleVar(value=0.75)
        self.sizing_range_card_offset_var = tk.DoubleVar(value=0.5)
        self.sizing_range_leather_mult_var = tk.DoubleVar(value=1.0)
        self.sizing_range_min_hole_var = tk.DoubleVar(value=16.5)
        self.sizing_range_felt_thick_var = tk.DoubleVar(value=3.175)
        self.sizing_range_felt_thick_unit_var = tk.StringVar(value="mm")

        # Engraving settings range mode vars
        self.eng_settings_range_mode_var = tk.StringVar(value=self.settings.get("engraving_settings_range_mode", "universal"))
        self.eng_settings_ranges = list(self.settings.get("engraving_settings_ranges", []))
        self.eng_settings_selected_range_index = None
        self.eng_settings_range_min_var = tk.DoubleVar(value=0.0)
        self.eng_settings_range_max_var = tk.DoubleVar(value=60.0)
        self.eng_settings_range_on_var = tk.BooleanVar(value=True)
        self.eng_settings_range_font_vars = {}

        # Engraving placement range mode vars
        self.eng_placement_range_mode_var = tk.StringVar(value=self.settings.get("engraving_placement_range_mode", "universal"))
        self.eng_placement_ranges = list(self.settings.get("engraving_placement_ranges", []))
        self.eng_placement_selected_range_index = None
        self.eng_placement_range_min_var = tk.DoubleVar(value=0.0)
        self.eng_placement_range_max_var = tk.DoubleVar(value=60.0)
        self.eng_placement_range_loc_vars = {}

        # Active preset name (the preset whose values currently sit in the
        # form). Updated on Load and Save Preset. Drives dirty checks and
        # the Save Preset overwrite default.
        self.active_preset_name = None

        # Live preview window state — toggle var + open Toplevel handle.
        self.show_preview_var = tk.BooleanVar(value=False)
        self.preview_window = None

        self.create_option_widgets()
        self._refresh_sizing_preset_combo()

        # Baseline = whatever the form holds right after construction.
        # _is_dirty() compares the current form to this baseline. We reset
        # the baseline after Load and Save Preset so post-load edits are
        # what register as dirty.
        self._baseline = self._capture_form_to_dict()

    # ------------------------------------------------------------------
    # Dirty / baseline tracking
    # ------------------------------------------------------------------

    def _is_dirty(self):
        """True if the form has unsaved edits relative to the baseline."""
        try:
            return self._capture_form_to_dict() != self._baseline
        except tk.TclError:
            # A numeric field is blank or mid-edit; that's an edit.
            return True

    def _form_is_valid(self, action=None):
        """Check every numeric field parses; if not, tell the user nicely.

        Without this, a blanked entry makes DoubleVar.get() raise TclError
        from Apply/Save/close paths — surfacing as the scary global
        'Unexpected Error' dialog and a window that won't respond.
        """
        try:
            self._capture_form_to_dict()
            return True
        except tk.TclError:
            messagebox.showerror(
                _("Invalid value"),
                _("One of the numeric fields is empty or not a number.\n\n"
                  "Please fix it before you {action}.").format(action=action or _("continue")),
                parent=self.top,
            )
            return False

    def _set_baseline_to_current(self):
        """Mark the current form state as the new clean baseline."""
        self._baseline = self._capture_form_to_dict()

    # ------------------------------------------------------------------
    # Live preview toggle
    # ------------------------------------------------------------------

    def _toggle_preview_window(self):
        """Open or close the live PadPreviewWindow per the toggle var."""
        if self.show_preview_var.get():
            if self.preview_window is None or not self.preview_window.winfo_exists():
                self.preview_window = PadPreviewWindow(self)
            else:
                self.preview_window.lift()
        else:
            if self.preview_window is not None and self.preview_window.winfo_exists():
                self.preview_window.destroy()
            self.preview_window = None

    def _close_preview_if_open(self):
        if self.preview_window is not None and self.preview_window.winfo_exists():
            try:
                self.preview_window.destroy()
            except tk.TclError:
                pass
        self.preview_window = None

    def create_option_widgets(self):
        main_frame = self.scrollable_frame

        # Presets live at the top — that's the natural entry point: load a
        # known recipe, tweak, save back. The rest of the form follows.
        self._build_sizing_preset_section(main_frame)

        # Preview toggle. Lives just under the preset row so the user can
        # flip it on, then watch the preview update as they edit values
        # below.
        preview_row = tk.Frame(main_frame, bg=DIALOG_BG)
        preview_row.pack(fill="x", pady=(0, 5))
        preview_cb = tk.Checkbutton(
            preview_row,
            text=_("Show live pad preview"),
            variable=self.show_preview_var, bg=DIALOG_BG,
            command=self._toggle_preview_window,
        )
        preview_cb.pack(side="left")
        add_tooltip(
            preview_cb,
            _("Open a resizable window that draws what the pad will look "
            "like with the current sizing rules. Pick a pad size and "
            "which materials to show; the preview updates live as you "
            "edit settings here."),
        )

        unit_frame = tk.LabelFrame(main_frame, text=_("Sheet Units"), bg=DIALOG_BG, padx=5, pady=5)
        unit_frame.pack(fill="x", pady=5)
        unit_tip = ("Units used for pad-size and sheet-dimension entries "
                    "throughout the app. Internal calculations and output "
                    "files always use millimeters regardless of this setting.")
        unit_in = tk.Radiobutton(unit_frame, text=_("Inches (in)"), variable=self.unit_var, value="in", bg=DIALOG_BG)
        unit_in.pack(side="left", padx=5)
        unit_cm = tk.Radiobutton(unit_frame, text=_("Centimeters (cm)"), variable=self.unit_var, value="cm", bg=DIALOG_BG)
        unit_cm.pack(side="left", padx=5)
        unit_mm = tk.Radiobutton(unit_frame, text=_("Millimeters (mm)"), variable=self.unit_var, value="mm", bg=DIALOG_BG)
        unit_mm.pack(side="left", padx=5)
        add_tooltips(unit_tip, unit_in, unit_cm, unit_mm)

        rules_frame = tk.LabelFrame(main_frame, text=_("Sizing Rules (Advanced)"), bg=DIALOG_BG, padx=5, pady=5)
        rules_frame.pack(fill="x", pady=5)
        rules_frame.columnconfigure(1, weight=1)

        sizing_mode_frame = tk.Frame(rules_frame, bg=DIALOG_BG)
        sizing_mode_frame.grid(row=0, column=0, columnspan=2, sticky='w', pady=2)
        sizing_mode_tip = (
            "Universal: one set of sizing rules for every pad size.\n"
            "Per Size Range: define different rules for different "
            "size bands. Pads outside any defined range fall back to the "
            "universal values."
        )
        sm_uni = tk.Radiobutton(sizing_mode_frame, text=_("Universal"), variable=self.sizing_range_mode_var,
                                value="universal", bg=DIALOG_BG, command=self._toggle_sizing_mode)
        sm_uni.pack(side="left", padx=(0, 10))
        sm_rng = tk.Radiobutton(sizing_mode_frame, text=_("Per Size Range"), variable=self.sizing_range_mode_var,
                                value="range", bg=DIALOG_BG, command=self._toggle_sizing_mode)
        sm_rng.pack(side="left")
        add_tooltips(sizing_mode_tip, sm_uni, sm_rng)

        # === Sizing Universal Sub-Frame ===
        self.sizing_universal_frame = tk.Frame(rules_frame, bg=DIALOG_BG)
        self.sizing_universal_frame.columnconfigure(1, weight=1)
        self._build_sizing_fields(self.sizing_universal_frame,
                                  self.felt_offset_var, self.card_offset_var,
                                  self.leather_mult_var, self.min_hole_size_var,
                                  self.felt_thickness_var, self.felt_thickness_unit_var)

        # === Sizing Range Sub-Frame ===
        self.sizing_range_frame = tk.Frame(rules_frame, bg=DIALOG_BG)
        self.sizing_range_frame.columnconfigure(1, weight=1)

        sr_sel = tk.Frame(self.sizing_range_frame, bg=DIALOG_BG)
        sr_sel.grid(row=0, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Label(sr_sel, text=_("Range:"), bg=DIALOG_BG).pack(side="left")
        self.sizing_range_combo = ttk.Combobox(sr_sel, state="readonly", width=25)
        self.sizing_range_combo.pack(side="left", padx=5)
        self.sizing_range_combo.bind("<<ComboboxSelected>>", self._on_sizing_range_selected)
        add_tooltip(self.sizing_range_combo,
                    _("Pick a defined size range to edit, or use Add Range "
                    "below to create a new one."))

        sr_min_lbl = tk.Label(self.sizing_range_frame, text=_("Min Size (mm):"), bg=DIALOG_BG)
        sr_min_lbl.grid(row=1, column=0, sticky='w', pady=2)
        sr_min_ent = tk.Entry(self.sizing_range_frame, textvariable=self.sizing_range_min_var, width=10)
        sr_min_ent.grid(row=1, column=1, sticky='w', pady=2)
        sr_max_lbl = tk.Label(self.sizing_range_frame, text=_("Max Size (mm):"), bg=DIALOG_BG)
        sr_max_lbl.grid(row=2, column=0, sticky='w', pady=2)
        sr_max_ent = tk.Entry(self.sizing_range_frame, textvariable=self.sizing_range_max_var, width=10)
        sr_max_ent.grid(row=2, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Lower bound (inclusive) of this size range, in millimeters."),
            sr_min_lbl, sr_min_ent,
        )
        add_tooltips(
            _("Upper bound (inclusive) of this size range, in millimeters."),
            sr_max_lbl, sr_max_ent,
        )

        self._build_sizing_fields(self.sizing_range_frame,
                                  self.sizing_range_felt_offset_var, self.sizing_range_card_offset_var,
                                  self.sizing_range_leather_mult_var, self.sizing_range_min_hole_var,
                                  self.sizing_range_felt_thick_var, self.sizing_range_felt_thick_unit_var,
                                  row_start=3)

        sr_btn = tk.Frame(self.sizing_range_frame, bg=DIALOG_BG)
        sr_btn.grid(row=8, column=0, columnspan=2, sticky='ew', pady=5)
        sr_add = tk.Button(sr_btn, text=_("Add Range"), command=self._add_sizing_range)
        sr_add.pack(side="left", padx=2)
        sr_upd = tk.Button(sr_btn, text=_("Update"), command=self._update_sizing_range)
        sr_upd.pack(side="left", padx=2)
        sr_del = tk.Button(sr_btn, text=_("Delete"), command=self._delete_sizing_range)
        sr_del.pack(side="left", padx=2)
        add_tooltip(sr_add, _("Add a new size range using the values entered above."))
        add_tooltip(sr_upd, _("Save the values above into the currently selected range."))
        add_tooltip(sr_del, _("Delete the currently selected range."))

        self._toggle_sizing_mode()

        # --- DART SETTINGS FRAME ---
        darts_frame = tk.LabelFrame(main_frame, text=_("Darts"), bg=DIALOG_BG, padx=5, pady=5)
        darts_frame.pack(fill="x", pady=5)
        darts_frame.columnconfigure(1, weight=1)

        darts_enable_cb = tk.Checkbutton(darts_frame, text=_("Enable Darts"),
                                         variable=self.darts_enabled_var, bg=DIALOG_BG)
        darts_enable_cb.grid(row=0, column=0, columnspan=2, sticky='w', pady=2)
        add_tooltip(
            darts_enable_cb,
            _("When on, leather discs below the threshold size get a dart "
            "cutout so the leather can fold around the felt without bunching.")
        )

        # Mode toggle: Universal vs Range
        mode_frame = tk.Frame(darts_frame, bg=DIALOG_BG)
        mode_frame.grid(row=1, column=0, columnspan=2, sticky='w', pady=2)
        dart_mode_tip = (
            "Universal: one set of dart values applies to all small pads.\n"
            "Per Size Range: define different dart settings for different "
            "size bands. Pads not in any range get no dart pattern (rather "
            "than falling back to universal)."
        )
        dm_uni = tk.Radiobutton(mode_frame, text=_("Universal"), variable=self.dart_range_mode_var,
                                value="universal", bg=DIALOG_BG, command=self._toggle_dart_mode)
        dm_uni.pack(side="left", padx=(0, 10))
        dm_rng = tk.Radiobutton(mode_frame, text=_("Per Size Range"), variable=self.dart_range_mode_var,
                                value="range", bg=DIALOG_BG, command=self._toggle_dart_mode)
        dm_rng.pack(side="left")
        add_tooltips(dart_mode_tip, dm_uni, dm_rng)

        # === UNIVERSAL SUB-FRAME ===
        self.dart_universal_frame = tk.Frame(darts_frame, bg=DIALOG_BG)
        self.dart_universal_frame.columnconfigure(1, weight=1)

        thr_lbl = tk.Label(self.dart_universal_frame, text=_("Use Darts below (mm):"), bg=DIALOG_BG)
        thr_lbl.grid(row=0, column=0, sticky='w', pady=2)
        thr_ent = tk.Entry(self.dart_universal_frame, textvariable=self.dart_threshold_var, width=10)
        thr_ent.grid(row=0, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Pad sizes at or below this value get a darted cut. Larger "
            "pads are cut as a plain wrap with no darts."),
            thr_lbl, thr_ent,
        )

        ow_lbl = tk.Label(self.dart_universal_frame, text=_("Dart Safe Overwrap (Valley) (mm):"), bg=DIALOG_BG)
        ow_lbl.grid(row=1, column=0, sticky='w', pady=2)
        ow_ent = tk.Entry(self.dart_universal_frame, textvariable=self.dart_overwrap_var, width=10)
        ow_ent.grid(row=1, column=1, sticky='w', pady=2)
        add_tooltips(
            _("How far each dart valley overlaps the felt edge. Larger "
            "values leave more material in the valleys (safer wrap, more "
            "bunching). 0.5 mm is a typical default."),
            ow_lbl, ow_ent,
        )

        wb_lbl = tk.Label(self.dart_universal_frame, text=_("Dart Wrap Bonus (Adds to Tip) (mm):"), bg=DIALOG_BG)
        wb_lbl.grid(row=2, column=0, sticky='w', pady=2)
        wb_ent = tk.Entry(self.dart_universal_frame, textvariable=self.dart_wrap_bonus_var, width=10)
        wb_ent.grid(row=2, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Extra material added to each dart tip. Larger values produce "
            "longer, pointier tips with more leather to wrap with."),
            wb_lbl, wb_ent,
        )

        fm_lbl = tk.Label(self.dart_universal_frame, text=_("Dart Frequency Multiplier (1.0=Default):"), bg=DIALOG_BG)
        fm_lbl.grid(row=3, column=0, sticky='w', pady=2)
        fm_ent = tk.Entry(self.dart_universal_frame, textvariable=self.dart_frequency_multiplier_var, width=10)
        fm_ent.grid(row=3, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Scales the number of dart points. 1.0 = default count; "
            ">1.0 = more, finer points; <1.0 = fewer, chunkier points."),
            fm_lbl, fm_ent,
        )

        shape_frame = tk.Frame(self.dart_universal_frame, bg=DIALOG_BG)
        shape_frame.grid(row=4, column=0, columnspan=2, sticky='ew', pady=5)
        shape_tip = (
            "Shape of the dart wave. Triangle (left) = sharp linear ramps "
            "between peaks and valleys. Sine (slider centered) = smooth, "
            "rounded curves. Square (right) = flats at peaks and valleys "
            "with steep transitions. The slider blends smoothly across "
            "the spectrum."
        )
        sh_lbl = tk.Label(shape_frame, text=_("Shape:"), bg=DIALOG_BG)
        sh_lbl.pack(side="left")
        sh_tri = tk.Label(shape_frame, text=_("Triangle"), bg=DIALOG_BG, font=("Arial", 8))
        sh_tri.pack(side="left", padx=(5, 0))
        sh_scale = tk.Scale(shape_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL,
                            variable=self.dart_shape_factor_var, showvalue=0,
                            bg=DIALOG_BG, highlightthickness=0, length=170, resolution=0.01)
        sh_scale.pack(side="left", fill="x", expand=True, padx=4)
        sh_sq = tk.Label(shape_frame, text=_("Square"), bg=DIALOG_BG, font=("Arial", 8))
        sh_sq.pack(side="left")
        add_tooltips(shape_tip, sh_lbl, sh_tri, sh_scale, sh_sq)

        tk.Label(self.dart_universal_frame, text="-------------------------", bg=DIALOG_BG).grid(row=5, column=0, columnspan=2, pady=5)
        dart_eng_cb = tk.Checkbutton(self.dart_universal_frame, text=_("Show Label on Dart Pads"),
                                     variable=self.dart_engraving_on_var, bg=DIALOG_BG)
        dart_eng_cb.grid(row=6, column=0, columnspan=2, sticky='w', pady=2)
        add_tooltip(dart_eng_cb,
                    _("Engrave the size number on darted leather pads. "
                    "Helpful for sorting them out after cutting."))

        # === RANGE SUB-FRAME ===
        self.dart_range_frame = tk.Frame(darts_frame, bg=DIALOG_BG)
        self.dart_range_frame.columnconfigure(1, weight=1)

        # Range selector dropdown
        range_sel_frame = tk.Frame(self.dart_range_frame, bg=DIALOG_BG)
        range_sel_frame.grid(row=0, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Label(range_sel_frame, text=_("Range:"), bg=DIALOG_BG).pack(side="left")
        self.range_combo = ttk.Combobox(range_sel_frame, state="readonly", width=25)
        self.range_combo.pack(side="left", padx=5)
        self.range_combo.bind("<<ComboboxSelected>>", self._on_range_selected)
        add_tooltip(self.range_combo,
                    _("Pick a defined dart size range to edit, or use Add "
                    "Range to define a new one."))

        # Range editing fields
        rmin_lbl = tk.Label(self.dart_range_frame, text=_("Min Size (mm):"), bg=DIALOG_BG)
        rmin_lbl.grid(row=1, column=0, sticky='w', pady=2)
        rmin_ent = tk.Entry(self.dart_range_frame, textvariable=self.range_min_var, width=10)
        rmin_ent.grid(row=1, column=1, sticky='w', pady=2)
        add_tooltips(_("Lower bound (inclusive) of this dart size range."), rmin_lbl, rmin_ent)

        rmax_lbl = tk.Label(self.dart_range_frame, text=_("Max Size (mm):"), bg=DIALOG_BG)
        rmax_lbl.grid(row=2, column=0, sticky='w', pady=2)
        rmax_ent = tk.Entry(self.dart_range_frame, textvariable=self.range_max_var, width=10)
        rmax_ent.grid(row=2, column=1, sticky='w', pady=2)
        add_tooltips(_("Upper bound (inclusive) of this dart size range."), rmax_lbl, rmax_ent)

        row_lbl = tk.Label(self.dart_range_frame, text=_("Overwrap (Valley) (mm):"), bg=DIALOG_BG)
        row_lbl.grid(row=3, column=0, sticky='w', pady=2)
        row_ent = tk.Entry(self.dart_range_frame, textvariable=self.range_overwrap_var, width=10)
        row_ent.grid(row=3, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Overlap at the dart valleys for this range. Larger values "
            "leave more material in the valleys."),
            row_lbl, row_ent,
        )

        rwb_lbl = tk.Label(self.dart_range_frame, text=_("Wrap Bonus (Tip) (mm):"), bg=DIALOG_BG)
        rwb_lbl.grid(row=4, column=0, sticky='w', pady=2)
        rwb_ent = tk.Entry(self.dart_range_frame, textvariable=self.range_wrap_bonus_var, width=10)
        rwb_ent.grid(row=4, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Extra material added to each dart tip for this range."),
            rwb_lbl, rwb_ent,
        )

        rfm_lbl = tk.Label(self.dart_range_frame, text=_("Frequency Multiplier:"), bg=DIALOG_BG)
        rfm_lbl.grid(row=5, column=0, sticky='w', pady=2)
        rfm_ent = tk.Entry(self.dart_range_frame, textvariable=self.range_freq_mult_var, width=10)
        rfm_ent.grid(row=5, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Number of points multiplier for this range. 1.0 = default; "
            ">1.0 = more, finer points; <1.0 = fewer, chunkier points."),
            rfm_lbl, rfm_ent,
        )

        range_shape_frame = tk.Frame(self.dart_range_frame, bg=DIALOG_BG)
        range_shape_frame.grid(row=6, column=0, columnspan=2, sticky='ew', pady=5)
        rsh_lbl = tk.Label(range_shape_frame, text=_("Shape:"), bg=DIALOG_BG)
        rsh_lbl.pack(side="left")
        rsh_tri = tk.Label(range_shape_frame, text=_("Triangle"), bg=DIALOG_BG, font=("Arial", 8))
        rsh_tri.pack(side="left", padx=(5, 0))
        rsh_scale = tk.Scale(range_shape_frame, from_=0.0, to=1.0, orient=tk.HORIZONTAL,
                             variable=self.range_shape_factor_var, showvalue=0,
                             bg=DIALOG_BG, highlightthickness=0, length=170, resolution=0.01)
        rsh_scale.pack(side="left", fill="x", expand=True, padx=4)
        rsh_sq = tk.Label(range_shape_frame, text=_("Square"), bg=DIALOG_BG, font=("Arial", 8))
        rsh_sq.pack(side="left")
        add_tooltips(
            _("Dart wave shape for this range. Triangle = sharp linear "
            "ramps; Sine (slider centered) = smooth curves; Square = "
            "flats with steep transitions."),
            rsh_lbl, rsh_tri, rsh_scale, rsh_sq,
        )

        tk.Label(self.dart_range_frame, text="-------------------------", bg=DIALOG_BG).grid(row=7, column=0, columnspan=2, pady=5)
        rdart_eng_cb = tk.Checkbutton(self.dart_range_frame, text=_("Show Label on Dart Pads"),
                                      variable=self.range_engraving_on_var, bg=DIALOG_BG)
        rdart_eng_cb.grid(row=8, column=0, columnspan=2, sticky='w', pady=2)
        add_tooltip(rdart_eng_cb,
                    _("Engrave the size number on darted leather pads in this range."))

        # Range action buttons
        btn_frame = tk.Frame(self.dart_range_frame, bg=DIALOG_BG)
        btn_frame.grid(row=10, column=0, columnspan=2, sticky='ew', pady=5)
        dr_add = tk.Button(btn_frame, text=_("Add Range"), command=self._add_range)
        dr_add.pack(side="left", padx=2)
        dr_upd = tk.Button(btn_frame, text=_("Update"), command=self._update_range)
        dr_upd.pack(side="left", padx=2)
        dr_del = tk.Button(btn_frame, text=_("Delete"), command=self._delete_range)
        dr_del.pack(side="left", padx=2)
        add_tooltip(dr_add, _("Add a new dart range using the values entered above."))
        add_tooltip(dr_upd, _("Save the values above into the currently selected range."))
        add_tooltip(dr_del, _("Delete the currently selected range."))

        # Show the correct sub-frame
        self._toggle_dart_mode()

        # --- ENGRAVING SETTINGS FRAME ---
        engraving_frame = tk.LabelFrame(main_frame, text=_("Engraving Settings (Standard Pads)"), bg=DIALOG_BG, padx=5, pady=5)
        engraving_frame.pack(fill="x", pady=5)

        es_mode_frame = tk.Frame(engraving_frame, bg=DIALOG_BG)
        es_mode_frame.pack(fill="x", pady=2)
        es_mode_tip = (
            "Universal: same engraving settings for all pad sizes.\n"
            "Per Size Range: define different engraving settings for "
            "different size bands; pads outside any range fall back to "
            "the universal values."
        )
        es_uni = tk.Radiobutton(es_mode_frame, text=_("Universal"), variable=self.eng_settings_range_mode_var,
                                value="universal", bg=DIALOG_BG, command=self._toggle_eng_settings_mode)
        es_uni.pack(side="left", padx=(0, 10))
        es_rng = tk.Radiobutton(es_mode_frame, text=_("Per Size Range"), variable=self.eng_settings_range_mode_var,
                                value="range", bg=DIALOG_BG, command=self._toggle_eng_settings_mode)
        es_rng.pack(side="left")
        add_tooltips(es_mode_tip, es_uni, es_rng)

        # === Engraving Settings Universal ===
        self.eng_settings_universal_frame = tk.Frame(engraving_frame, bg=DIALOG_BG)
        eng_on_cb = tk.Checkbutton(self.eng_settings_universal_frame, text=_("Show Size Label"),
                                   variable=self.engraving_on_var, bg=DIALOG_BG)
        eng_on_cb.pack(anchor='w')
        add_tooltip(eng_on_cb, _("Engrave the pad-size number on each disc for identification."))

        fs_frame = tk.LabelFrame(self.eng_settings_universal_frame, text=_("Font Sizes (mm)"), bg=DIALOG_BG, padx=5, pady=5)
        fs_frame.pack(fill='x', pady=5)
        materials = ['felt', 'card', 'leather', 'exact_size']
        for i, mat in enumerate(materials):
            mat_label = mat.replace('_', ' ').capitalize()
            mat_lbl = tk.Label(fs_frame, text=_(mat_label) + ":", bg=DIALOG_BG)
            mat_lbl.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            fvar = tk.DoubleVar(value=self.settings["engraving_font_size"].get(mat, 2.0))
            self.engraving_font_size_vars[mat] = fvar
            mat_ent = tk.Entry(fs_frame, textvariable=fvar, width=8)
            mat_ent.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            add_tooltips(
                _("Engraving font size in millimeters for {material} discs.").format(material=mat_label.lower()),
                mat_lbl, mat_ent,
            )

        # === Engraving Settings Range ===
        self.eng_settings_range_frame = tk.Frame(engraving_frame, bg=DIALOG_BG)
        self.eng_settings_range_frame.columnconfigure(1, weight=1)

        esr_sel = tk.Frame(self.eng_settings_range_frame, bg=DIALOG_BG)
        esr_sel.grid(row=0, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Label(esr_sel, text=_("Range:"), bg=DIALOG_BG).pack(side="left")
        self.eng_settings_range_combo = ttk.Combobox(esr_sel, state="readonly", width=25)
        self.eng_settings_range_combo.pack(side="left", padx=5)
        self.eng_settings_range_combo.bind("<<ComboboxSelected>>", self._on_eng_settings_range_selected)
        add_tooltip(self.eng_settings_range_combo,
                    _("Pick an engraving-settings range to edit, or use Add "
                    "Range to define a new one."))

        esr_min_lbl = tk.Label(self.eng_settings_range_frame, text=_("Min Size (mm):"), bg=DIALOG_BG)
        esr_min_lbl.grid(row=1, column=0, sticky='w', pady=2)
        esr_min_ent = tk.Entry(self.eng_settings_range_frame, textvariable=self.eng_settings_range_min_var, width=10)
        esr_min_ent.grid(row=1, column=1, sticky='w', pady=2)
        esr_max_lbl = tk.Label(self.eng_settings_range_frame, text=_("Max Size (mm):"), bg=DIALOG_BG)
        esr_max_lbl.grid(row=2, column=0, sticky='w', pady=2)
        esr_max_ent = tk.Entry(self.eng_settings_range_frame, textvariable=self.eng_settings_range_max_var, width=10)
        esr_max_ent.grid(row=2, column=1, sticky='w', pady=2)
        add_tooltips(_("Lower bound (inclusive) of this size range."), esr_min_lbl, esr_min_ent)
        add_tooltips(_("Upper bound (inclusive) of this size range."), esr_max_lbl, esr_max_ent)

        esr_on_cb = tk.Checkbutton(self.eng_settings_range_frame, text=_("Show Size Label"),
                                   variable=self.eng_settings_range_on_var, bg=DIALOG_BG)
        esr_on_cb.grid(row=3, column=0, columnspan=2, sticky='w', pady=2)
        add_tooltip(esr_on_cb, _("Engrave the size number on pads in this size range."))

        esr_fs = tk.LabelFrame(self.eng_settings_range_frame, text=_("Font Sizes (mm)"), bg=DIALOG_BG, padx=5, pady=5)
        esr_fs.grid(row=4, column=0, columnspan=2, sticky='ew', pady=2)
        for i, mat in enumerate(materials):
            mat_label = mat.replace('_', ' ').capitalize()
            mat_lbl = tk.Label(esr_fs, text=f"{mat_label}:", bg=DIALOG_BG)
            mat_lbl.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            fvar = tk.DoubleVar(value=2.0)
            self.eng_settings_range_font_vars[mat] = fvar
            mat_ent = tk.Entry(esr_fs, textvariable=fvar, width=8)
            mat_ent.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            add_tooltips(
                _("Engraving font size for {material} pads in this range.").format(material=mat_label.lower()),
                mat_lbl, mat_ent,
            )

        esr_btn = tk.Frame(self.eng_settings_range_frame, bg=DIALOG_BG)
        esr_btn.grid(row=5, column=0, columnspan=2, sticky='ew', pady=5)
        esr_add = tk.Button(esr_btn, text=_("Add Range"), command=self._add_eng_settings_range)
        esr_add.pack(side="left", padx=2)
        esr_upd = tk.Button(esr_btn, text=_("Update"), command=self._update_eng_settings_range)
        esr_upd.pack(side="left", padx=2)
        esr_del = tk.Button(esr_btn, text=_("Delete"), command=self._delete_eng_settings_range)
        esr_del.pack(side="left", padx=2)
        add_tooltip(esr_add, _("Add a new engraving-settings range using the values above."))
        add_tooltip(esr_upd, _("Save the values above into the currently selected range."))
        add_tooltip(esr_del, _("Delete the currently selected range."))

        self._toggle_eng_settings_mode()

        # --- ENGRAVING PLACEMENT FRAME ---
        engraving_loc_frame = tk.LabelFrame(main_frame, text=_("Engraving Placement"), bg=DIALOG_BG, padx=5, pady=5)
        engraving_loc_frame.pack(fill="x", pady=5)

        ep_mode_frame = tk.Frame(engraving_loc_frame, bg=DIALOG_BG)
        ep_mode_frame.pack(fill="x", pady=2)
        ep_mode_tip = (
            "Universal: same engraving placement for all pad sizes.\n"
            "Per Size Range: define different placement for different "
            "size bands; pads outside any range fall back to the "
            "universal values."
        )
        ep_uni = tk.Radiobutton(ep_mode_frame, text=_("Universal"), variable=self.eng_placement_range_mode_var,
                                value="universal", bg=DIALOG_BG, command=self._toggle_eng_placement_mode)
        ep_uni.pack(side="left", padx=(0, 10))
        ep_rng = tk.Radiobutton(ep_mode_frame, text=_("Per Size Range"), variable=self.eng_placement_range_mode_var,
                                value="range", bg=DIALOG_BG, command=self._toggle_eng_placement_mode)
        ep_rng.pack(side="left")
        add_tooltips(ep_mode_tip, ep_uni, ep_rng)

        # === Engraving Placement Universal ===
        placement_materials = ['leather', 'darted_leather', 'felt', 'card', 'exact_size']
        # Material labels — keep all five as explicit _() literals so pybabel
        # extracts them. The single-word ones share msgids with pad-maker
        # material checkboxes already in the catalog.
        placement_labels = {
            'leather': _("Leather"),
            'darted_leather': _("Darted leather"),
            'felt': _("Felt"),
            'card': _("Card"),
            'exact_size': _("Exact size"),
        }
        placement_help = {
            'leather': "Where the size label sits on plain leather wraps.",
            'darted_leather': "Where the size label sits on darted leather pads.",
            'felt': "Where the size label sits on felt discs.",
            'card': "Where the size label sits on card backing discs.",
            'exact_size': "Where the size label sits on exact-size discs.",
        }
        mode_help = (
            "out = distance measured inward from the outer edge.\n"
            "in = distance measured outward from the center hole.\n"
            "ctr = centered between the two."
        )
        self.eng_placement_universal_frame = tk.Frame(engraving_loc_frame, bg=DIALOG_BG)
        for mat in placement_materials:
            frame = tk.Frame(self.eng_placement_universal_frame, bg=DIALOG_BG)
            frame.pack(fill='x', pady=2)
            label = placement_labels.get(mat, mat.capitalize())
            row_lbl = tk.Label(frame, text=label + ":", bg=DIALOG_BG, width=15, anchor='w')
            row_lbl.pack(side="left")
            loc = self.settings["engraving_location"].get(mat, {"mode": "from_outside", "value": 2.5})
            mode_var = tk.StringVar(value=loc['mode'])
            val_var = tk.DoubleVar(value=loc['value'])
            self.engraving_loc_vars[mat] = {'mode': mode_var, 'value': val_var}
            rb_out = tk.Radiobutton(frame, text=_("out"), variable=mode_var, value="from_outside", bg=DIALOG_BG)
            rb_out.pack(side="left")
            # "int" (not "in") so the msgid doesn't collide with the
            # inches unit "in" used in tooling/felt-thickness pickers.
            rb_in = tk.Radiobutton(frame, text=_("int"), variable=mode_var, value="from_inside", bg=DIALOG_BG)
            rb_in.pack(side="left")
            rb_ctr = tk.Radiobutton(frame, text=_("ctr"), variable=mode_var, value="centered", bg=DIALOG_BG)
            rb_ctr.pack(side="left")
            val_ent = tk.Entry(frame, textvariable=val_var, width=5)
            val_ent.pack(side="left", padx=5)
            mm_lbl = tk.Label(frame, text=_("mm"), bg=DIALOG_BG)
            mm_lbl.pack(side="left")
            add_tooltip(row_lbl, placement_help[mat])
            add_tooltips(mode_help, rb_out, rb_in, rb_ctr)
            add_tooltips(
                _("Distance in millimeters from the chosen reference "
                "(outside, inside, or centered)."),
                val_ent, mm_lbl,
            )

        # === Engraving Placement Range ===
        self.eng_placement_range_frame = tk.Frame(engraving_loc_frame, bg=DIALOG_BG)
        self.eng_placement_range_frame.columnconfigure(1, weight=1)

        epr_sel = tk.Frame(self.eng_placement_range_frame, bg=DIALOG_BG)
        epr_sel.grid(row=0, column=0, columnspan=2, sticky='ew', pady=2)
        tk.Label(epr_sel, text=_("Range:"), bg=DIALOG_BG).pack(side="left")
        self.eng_placement_range_combo = ttk.Combobox(epr_sel, state="readonly", width=25)
        self.eng_placement_range_combo.pack(side="left", padx=5)
        self.eng_placement_range_combo.bind("<<ComboboxSelected>>", self._on_eng_placement_range_selected)
        add_tooltip(self.eng_placement_range_combo,
                    _("Pick an engraving-placement range to edit, or use Add "
                    "Range to define a new one."))

        epr_min_lbl = tk.Label(self.eng_placement_range_frame, text=_("Min Size (mm):"), bg=DIALOG_BG)
        epr_min_lbl.grid(row=1, column=0, sticky='w', pady=2)
        epr_min_ent = tk.Entry(self.eng_placement_range_frame, textvariable=self.eng_placement_range_min_var, width=10)
        epr_min_ent.grid(row=1, column=1, sticky='w', pady=2)
        epr_max_lbl = tk.Label(self.eng_placement_range_frame, text=_("Max Size (mm):"), bg=DIALOG_BG)
        epr_max_lbl.grid(row=2, column=0, sticky='w', pady=2)
        epr_max_ent = tk.Entry(self.eng_placement_range_frame, textvariable=self.eng_placement_range_max_var, width=10)
        epr_max_ent.grid(row=2, column=1, sticky='w', pady=2)
        add_tooltips(_("Lower bound (inclusive) of this size range."), epr_min_lbl, epr_min_ent)
        add_tooltips(_("Upper bound (inclusive) of this size range."), epr_max_lbl, epr_max_ent)

        epr_loc = tk.Frame(self.eng_placement_range_frame, bg=DIALOG_BG)
        epr_loc.grid(row=3, column=0, columnspan=2, sticky='ew', pady=2)
        for mat in placement_materials:
            frame = tk.Frame(epr_loc, bg=DIALOG_BG)
            frame.pack(fill='x', pady=2)
            label = placement_labels.get(mat, mat.capitalize())
            row_lbl = tk.Label(frame, text=label + ":", bg=DIALOG_BG, width=15, anchor='w')
            row_lbl.pack(side="left")
            mode_var = tk.StringVar(value="from_outside")
            val_var = tk.DoubleVar(value=2.5)
            self.eng_placement_range_loc_vars[mat] = {'mode': mode_var, 'value': val_var}
            rb_out = tk.Radiobutton(frame, text=_("out"), variable=mode_var, value="from_outside", bg=DIALOG_BG)
            rb_out.pack(side="left")
            rb_in = tk.Radiobutton(frame, text=_("int"), variable=mode_var, value="from_inside", bg=DIALOG_BG)
            rb_in.pack(side="left")
            rb_ctr = tk.Radiobutton(frame, text=_("ctr"), variable=mode_var, value="centered", bg=DIALOG_BG)
            rb_ctr.pack(side="left")
            val_ent = tk.Entry(frame, textvariable=val_var, width=5)
            val_ent.pack(side="left", padx=5)
            mm_lbl = tk.Label(frame, text=_("mm"), bg=DIALOG_BG)
            mm_lbl.pack(side="left")
            add_tooltip(row_lbl, placement_help[mat] + " (this size range)")
            add_tooltips(mode_help, rb_out, rb_in, rb_ctr)
            add_tooltips(
                _("Distance from the chosen reference, in millimeters."),
                val_ent, mm_lbl,
            )

        epr_btn = tk.Frame(self.eng_placement_range_frame, bg=DIALOG_BG)
        epr_btn.grid(row=4, column=0, columnspan=2, sticky='ew', pady=5)
        epr_add = tk.Button(epr_btn, text=_("Add Range"), command=self._add_eng_placement_range)
        epr_add.pack(side="left", padx=2)
        epr_upd = tk.Button(epr_btn, text=_("Update"), command=self._update_eng_placement_range)
        epr_upd.pack(side="left", padx=2)
        epr_del = tk.Button(epr_btn, text=_("Delete"), command=self._delete_eng_placement_range)
        epr_del.pack(side="left", padx=2)
        add_tooltip(epr_add, _("Add a new engraving-placement range using the values above."))
        add_tooltip(epr_upd, _("Save the values above into the currently selected range."))
        add_tooltip(epr_del, _("Delete the currently selected range."))

        self._toggle_eng_placement_mode()

        export_frame = tk.LabelFrame(main_frame, text=_("Export Settings"), bg=DIALOG_BG, padx=5, pady=5)
        export_frame.pack(fill="x", pady=5)
        compat_cb = tk.Checkbutton(export_frame,
                                   text=_("Enable Inkscape/Compatibility Mode (unitless SVG)"),
                                   variable=self.compatibility_mode_var, bg=DIALOG_BG)
        compat_cb.pack(anchor='w')
        add_tooltip(compat_cb,
                    _("Write SVGs without explicit unit attributes. Turn on "
                    "if Inkscape (or other software) misinterprets the file "
                    "scale. LightBurn does not need this."))

    def _build_sizing_preset_section(self, parent):
        """Top-of-dialog preset controls: pick a preset, load, save, rename, delete."""
        preset_frame = tk.LabelFrame(parent, text=_("Sizing Rules Preset"), bg=DIALOG_BG, padx=5, pady=5)
        preset_frame.pack(fill="x", pady=(0, 5))

        select_row = tk.Frame(preset_frame, bg=DIALOG_BG)
        select_row.pack(fill="x", pady=(0, 4))
        tk.Label(select_row, text=_("Preset:"), bg=DIALOG_BG).pack(side="left")
        self.preset_combo = ttk.Combobox(select_row, state="readonly", width=28)
        self.preset_combo.pack(side="left", padx=5, fill="x", expand=True)
        load_btn = tk.Button(select_row, text=_("Load"), command=self.on_load_sizing_preset)
        load_btn.pack(side="left", padx=2)
        save_btn = tk.Button(select_row, text=_("Save Preset"), command=self.on_save_sizing_preset)
        save_btn.pack(side="left", padx=2)
        add_tooltip(self.preset_combo,
                    _("Saved snapshots of every value in this dialog. Pick "
                    "one and click Load to fill the form below."))
        add_tooltip(load_btn,
                    _("Fill the form below from the selected preset. "
                    "Click Apply at the bottom to commit it to the app."))
        add_tooltip(save_btn,
                    _("Save the current form as a preset — overwrite an "
                    "existing one or create a new one (you'll be asked)."))

        button_row = tk.Frame(preset_frame, bg=DIALOG_BG)
        button_row.pack(fill="x")
        rename_btn = tk.Button(button_row, text=_("Rename"), command=self.on_rename_sizing_preset)
        rename_btn.pack(side="left", padx=2)
        del_btn = tk.Button(button_row, text=_("Delete"), command=self.on_delete_sizing_preset)
        del_btn.pack(side="left", padx=2)
        imp_btn = tk.Button(button_row, text=_("Import..."), command=self.on_import_sizing_presets)
        imp_btn.pack(side="left", padx=2)
        exp_btn = tk.Button(button_row, text=_("Export..."), command=self.on_export_sizing_presets)
        exp_btn.pack(side="left", padx=2)
        add_tooltip(rename_btn,
                    _("Rename the selected preset."))
        add_tooltip(del_btn,
                    _("Remove the selected preset (cannot be undone). At "
                    "least one preset must remain."))
        add_tooltip(imp_btn,
                    _("Add presets from a JSON file shared by another tech."))
        add_tooltip(exp_btn,
                    _("Save the entire preset library to a JSON file you "
                    "can share."))

    # --- Dart Range Management ---

    def _toggle_dart_mode(self):
        """Show/hide universal vs range sub-frames."""
        if self.dart_range_mode_var.get() == "universal":
            self.dart_range_frame.grid_forget()
            self.dart_universal_frame.grid(row=2, column=0, columnspan=2, sticky='ew')
        else:
            self.dart_universal_frame.grid_forget()
            self.dart_range_frame.grid(row=2, column=0, columnspan=2, sticky='ew')
            self._refresh_range_combo()

    def _refresh_range_combo(self):
        """Update the range combobox values from self.dart_ranges.

        Always reloads the editing fields when there's a valid selection so
        that loading a different preset (which keeps the same index but
        changes the underlying range data) doesn't leave the editor showing
        stale values.
        """
        labels = [f"{r['min_size']:.1f} - {r['max_size']:.1f} mm" for r in self.dart_ranges]
        self.range_combo['values'] = labels
        if labels and self.selected_range_index is not None and self.selected_range_index < len(labels):
            self.range_combo.current(self.selected_range_index)
            self._load_range_fields(self.selected_range_index)
        elif labels:
            self.range_combo.current(len(labels) - 1)
            self.selected_range_index = len(labels) - 1
            self._load_range_fields(self.selected_range_index)
        else:
            self.range_combo.set("")
            self.selected_range_index = None

    def _on_range_selected(self, event=None):
        """Populate editing fields from the selected range."""
        idx = self.range_combo.current()
        if idx < 0 or idx >= len(self.dart_ranges):
            return
        self.selected_range_index = idx
        self._load_range_fields(idx)

    def _load_range_fields(self, idx):
        """Load a range's values into the editing fields."""
        r = self.dart_ranges[idx]
        self.range_min_var.set(r.get("min_size", 0.0))
        self.range_max_var.set(r.get("max_size", 18.0))
        self.range_overwrap_var.set(r.get("overwrap", 0.5))
        self.range_wrap_bonus_var.set(r.get("wrap_bonus", 0.75))
        self.range_freq_mult_var.set(r.get("frequency_multiplier", 1.0))
        # 0.5 = Sine, the scale's neutral default (0.0 would be Triangle)
        self.range_shape_factor_var.set(r.get("shape_factor", 0.5))
        self.range_engraving_on_var.set(r.get("engraving_on", True))

    def _read_range_fields(self):
        """Read current editing fields into a range dict."""
        return {
            "min_size": self.range_min_var.get(),
            "max_size": self.range_max_var.get(),
            "overwrap": self.range_overwrap_var.get(),
            "wrap_bonus": self.range_wrap_bonus_var.get(),
            "frequency_multiplier": self.range_freq_mult_var.get(),
            "shape_factor": self.range_shape_factor_var.get(),
            "engraving_on": self.range_engraving_on_var.get(),
        }

    def _add_range(self):
        """Add a new range from the current editing fields."""
        r = self._read_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.dart_ranges.append(r)
        self.dart_ranges.sort(key=lambda x: x["min_size"])
        self.selected_range_index = self.dart_ranges.index(r)
        self._refresh_range_combo()

    def _update_range(self):
        """Update the currently selected range with editing field values."""
        if self.selected_range_index is None or self.selected_range_index >= len(self.dart_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to update."))
            return
        r = self._read_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.dart_ranges[self.selected_range_index] = r
        self.dart_ranges.sort(key=lambda x: x["min_size"])
        self.selected_range_index = self.dart_ranges.index(r)
        self._refresh_range_combo()

    def _delete_range(self):
        """Delete the currently selected range."""
        if self.selected_range_index is None or self.selected_range_index >= len(self.dart_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to delete."))
            return
        del self.dart_ranges[self.selected_range_index]
        self.selected_range_index = None
        self._refresh_range_combo()

    # --- Sizing Fields Helper ---

    def _build_sizing_fields(self, parent, felt_var, card_var, leather_var, hole_var, thick_var, thick_unit_var, row_start=0):
        """Build the common sizing rule fields into a grid frame."""
        felt_lbl = tk.Label(parent, text=_("Felt Diameter Reduction (mm):"), bg=DIALOG_BG)
        felt_lbl.grid(row=row_start, column=0, sticky='w', pady=2)
        felt_ent = tk.Entry(parent, textvariable=felt_var, width=10)
        felt_ent.grid(row=row_start, column=1, sticky='w', pady=2)
        add_tooltips(
            _("How much smaller the felt disc is than the pad-cup size you "
            "enter. 0.75 mm means the felt disc is cut 0.75 mm under the "
            "stated diameter."),
            felt_lbl, felt_ent,
        )

        card_lbl = tk.Label(parent, text=_("Card Additional Reduction (mm):"), bg=DIALOG_BG)
        card_lbl.grid(row=row_start + 1, column=0, sticky='w', pady=2)
        card_ent = tk.Entry(parent, textvariable=card_var, width=10)
        card_ent.grid(row=row_start + 1, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Extra reduction applied on top of the felt offset for the "
            "cardboard backing disc — keeps the card just under the felt."),
            card_lbl, card_ent,
        )

        lea_lbl = tk.Label(parent, text=_("Leather Wrap Multiplier (1.00=default):"), bg=DIALOG_BG)
        lea_lbl.grid(row=row_start + 2, column=0, sticky='w', pady=2)
        lea_ent = tk.Entry(parent, textvariable=leather_var, width=10)
        lea_ent.grid(row=row_start + 2, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Controls how much extra leather is added to wrap around the "
            "felt. 1.00 = standard wrap; raise to add more wrap, lower to "
            "take some away."),
            lea_lbl, lea_ent,
        )

        hole_lbl = tk.Label(parent, text=_("Min. Pad Size for Hole (mm):"), bg=DIALOG_BG)
        hole_lbl.grid(row=row_start + 3, column=0, sticky='w', pady=2)
        hole_ent = tk.Entry(parent, textvariable=hole_var, width=10)
        hole_ent.grid(row=row_start + 3, column=1, sticky='w', pady=2)
        add_tooltips(
            _("Pads smaller than this size skip the center mounting hole "
            "automatically — they're too small for it to be useful."),
            hole_lbl, hole_ent,
        )

        ft_frame = tk.Frame(parent, bg=DIALOG_BG)
        ft_frame.grid(row=row_start + 4, column=0, columnspan=2, sticky='w', pady=2)
        ft_lbl = tk.Label(ft_frame, text=_("Felt Thickness:"), bg=DIALOG_BG)
        ft_lbl.pack(side="left")
        ft_ent = tk.Entry(ft_frame, textvariable=thick_var, width=10)
        ft_ent.pack(side="left", padx=5)
        ft_in = tk.Radiobutton(ft_frame, text=_("in"), variable=thick_unit_var, value="in", bg=DIALOG_BG)
        ft_in.pack(side="left")
        ft_mm = tk.Radiobutton(ft_frame, text=_("mm"), variable=thick_unit_var, value="mm", bg=DIALOG_BG)
        ft_mm.pack(side="left")
        add_tooltips(
            _("Thickness of the felt you're using. Feeds into the "
            "leather-wrap calculation so the wrap accounts for the felt's "
            "side wall."),
            ft_lbl, ft_ent, ft_in, ft_mm,
        )

    # --- Sizing Range Management ---

    def _toggle_sizing_mode(self):
        if self.sizing_range_mode_var.get() == "universal":
            self.sizing_range_frame.grid_forget()
            self.sizing_universal_frame.grid(row=1, column=0, columnspan=2, sticky='ew')
        else:
            self.sizing_universal_frame.grid_forget()
            self.sizing_range_frame.grid(row=1, column=0, columnspan=2, sticky='ew')
            self._refresh_sizing_range_combo()

    def _refresh_sizing_range_combo(self):
        labels = [f"{r['min_size']:.1f} - {r['max_size']:.1f} mm" for r in self.sizing_ranges]
        self.sizing_range_combo['values'] = labels
        if labels and self.sizing_selected_range_index is not None and self.sizing_selected_range_index < len(labels):
            self.sizing_range_combo.current(self.sizing_selected_range_index)
            # Reload fields too: a preset Load may keep the same index but
            # point it at completely different range data.
            self._load_sizing_range_fields(self.sizing_selected_range_index)
        elif labels:
            self.sizing_range_combo.current(len(labels) - 1)
            self.sizing_selected_range_index = len(labels) - 1
            self._load_sizing_range_fields(self.sizing_selected_range_index)
        else:
            self.sizing_range_combo.set("")
            self.sizing_selected_range_index = None

    def _on_sizing_range_selected(self, event=None):
        idx = self.sizing_range_combo.current()
        if 0 <= idx < len(self.sizing_ranges):
            self.sizing_selected_range_index = idx
            self._load_sizing_range_fields(idx)

    def _load_sizing_range_fields(self, idx):
        r = self.sizing_ranges[idx]
        self.sizing_range_min_var.set(r.get("min_size", 0.0))
        self.sizing_range_max_var.set(r.get("max_size", 60.0))
        self.sizing_range_felt_offset_var.set(r.get("felt_offset", 0.75))
        self.sizing_range_card_offset_var.set(r.get("card_to_felt_offset", 0.5))
        self.sizing_range_leather_mult_var.set(r.get("leather_wrap_multiplier", 1.0))
        self.sizing_range_min_hole_var.set(r.get("min_hole_size", 16.5))
        self.sizing_range_felt_thick_var.set(r.get("felt_thickness", 3.175))
        self.sizing_range_felt_thick_unit_var.set(r.get("felt_thickness_unit", "mm"))

    def _read_sizing_range_fields(self):
        return {
            "min_size": self.sizing_range_min_var.get(),
            "max_size": self.sizing_range_max_var.get(),
            "felt_offset": self.sizing_range_felt_offset_var.get(),
            "card_to_felt_offset": self.sizing_range_card_offset_var.get(),
            "leather_wrap_multiplier": self.sizing_range_leather_mult_var.get(),
            "min_hole_size": self.sizing_range_min_hole_var.get(),
            "felt_thickness": self.sizing_range_felt_thick_var.get(),
            "felt_thickness_unit": self.sizing_range_felt_thick_unit_var.get(),
        }

    def _add_sizing_range(self):
        r = self._read_sizing_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.sizing_ranges.append(r)
        self.sizing_ranges.sort(key=lambda x: x["min_size"])
        self.sizing_selected_range_index = self.sizing_ranges.index(r)
        self._refresh_sizing_range_combo()

    def _update_sizing_range(self):
        if self.sizing_selected_range_index is None or self.sizing_selected_range_index >= len(self.sizing_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to update."))
            return
        r = self._read_sizing_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.sizing_ranges[self.sizing_selected_range_index] = r
        self.sizing_ranges.sort(key=lambda x: x["min_size"])
        self.sizing_selected_range_index = self.sizing_ranges.index(r)
        self._refresh_sizing_range_combo()

    def _delete_sizing_range(self):
        if self.sizing_selected_range_index is None or self.sizing_selected_range_index >= len(self.sizing_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to delete."))
            return
        del self.sizing_ranges[self.sizing_selected_range_index]
        self.sizing_selected_range_index = None
        self._refresh_sizing_range_combo()

    # --- Engraving Settings Range Management ---

    def _toggle_eng_settings_mode(self):
        if self.eng_settings_range_mode_var.get() == "universal":
            self.eng_settings_range_frame.pack_forget()
            self.eng_settings_universal_frame.pack(fill="x")
        else:
            self.eng_settings_universal_frame.pack_forget()
            self.eng_settings_range_frame.pack(fill="x")
            self._refresh_eng_settings_range_combo()

    def _refresh_eng_settings_range_combo(self):
        labels = [f"{r['min_size']:.1f} - {r['max_size']:.1f} mm" for r in self.eng_settings_ranges]
        self.eng_settings_range_combo['values'] = labels
        if labels and self.eng_settings_selected_range_index is not None and self.eng_settings_selected_range_index < len(labels):
            self.eng_settings_range_combo.current(self.eng_settings_selected_range_index)
            # Reload fields: preset Load may have replaced the underlying data.
            self._load_eng_settings_range_fields(self.eng_settings_selected_range_index)
        elif labels:
            self.eng_settings_range_combo.current(len(labels) - 1)
            self.eng_settings_selected_range_index = len(labels) - 1
            self._load_eng_settings_range_fields(self.eng_settings_selected_range_index)
        else:
            self.eng_settings_range_combo.set("")
            self.eng_settings_selected_range_index = None

    def _on_eng_settings_range_selected(self, event=None):
        idx = self.eng_settings_range_combo.current()
        if 0 <= idx < len(self.eng_settings_ranges):
            self.eng_settings_selected_range_index = idx
            self._load_eng_settings_range_fields(idx)

    def _load_eng_settings_range_fields(self, idx):
        r = self.eng_settings_ranges[idx]
        self.eng_settings_range_min_var.set(r.get("min_size", 0.0))
        self.eng_settings_range_max_var.set(r.get("max_size", 60.0))
        self.eng_settings_range_on_var.set(r.get("engraving_on", True))
        fs = r.get("engraving_font_size", {})
        for mat, var in self.eng_settings_range_font_vars.items():
            var.set(fs.get(mat, 2.0))

    def _read_eng_settings_range_fields(self):
        return {
            "min_size": self.eng_settings_range_min_var.get(),
            "max_size": self.eng_settings_range_max_var.get(),
            "engraving_on": self.eng_settings_range_on_var.get(),
            "engraving_font_size": {mat: var.get() for mat, var in self.eng_settings_range_font_vars.items()},
        }

    def _add_eng_settings_range(self):
        r = self._read_eng_settings_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.eng_settings_ranges.append(r)
        self.eng_settings_ranges.sort(key=lambda x: x["min_size"])
        self.eng_settings_selected_range_index = self.eng_settings_ranges.index(r)
        self._refresh_eng_settings_range_combo()

    def _update_eng_settings_range(self):
        if self.eng_settings_selected_range_index is None or self.eng_settings_selected_range_index >= len(self.eng_settings_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to update."))
            return
        r = self._read_eng_settings_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.eng_settings_ranges[self.eng_settings_selected_range_index] = r
        self.eng_settings_ranges.sort(key=lambda x: x["min_size"])
        self.eng_settings_selected_range_index = self.eng_settings_ranges.index(r)
        self._refresh_eng_settings_range_combo()

    def _delete_eng_settings_range(self):
        if self.eng_settings_selected_range_index is None or self.eng_settings_selected_range_index >= len(self.eng_settings_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to delete."))
            return
        del self.eng_settings_ranges[self.eng_settings_selected_range_index]
        self.eng_settings_selected_range_index = None
        self._refresh_eng_settings_range_combo()

    # --- Engraving Placement Range Management ---

    def _toggle_eng_placement_mode(self):
        if self.eng_placement_range_mode_var.get() == "universal":
            self.eng_placement_range_frame.pack_forget()
            self.eng_placement_universal_frame.pack(fill="x")
        else:
            self.eng_placement_universal_frame.pack_forget()
            self.eng_placement_range_frame.pack(fill="x")
            self._refresh_eng_placement_range_combo()

    def _refresh_eng_placement_range_combo(self):
        labels = [f"{r['min_size']:.1f} - {r['max_size']:.1f} mm" for r in self.eng_placement_ranges]
        self.eng_placement_range_combo['values'] = labels
        if labels and self.eng_placement_selected_range_index is not None and self.eng_placement_selected_range_index < len(labels):
            self.eng_placement_range_combo.current(self.eng_placement_selected_range_index)
            # Reload fields: preset Load may have replaced the underlying data.
            self._load_eng_placement_range_fields(self.eng_placement_selected_range_index)
        elif labels:
            self.eng_placement_range_combo.current(len(labels) - 1)
            self.eng_placement_selected_range_index = len(labels) - 1
            self._load_eng_placement_range_fields(self.eng_placement_selected_range_index)
        else:
            self.eng_placement_range_combo.set("")
            self.eng_placement_selected_range_index = None

    def _on_eng_placement_range_selected(self, event=None):
        idx = self.eng_placement_range_combo.current()
        if 0 <= idx < len(self.eng_placement_ranges):
            self.eng_placement_selected_range_index = idx
            self._load_eng_placement_range_fields(idx)

    def _load_eng_placement_range_fields(self, idx):
        r = self.eng_placement_ranges[idx]
        self.eng_placement_range_min_var.set(r.get("min_size", 0.0))
        self.eng_placement_range_max_var.set(r.get("max_size", 60.0))
        loc = r.get("engraving_location", {})
        for mat, vars in self.eng_placement_range_loc_vars.items():
            ml = loc.get(mat, {"mode": "from_outside", "value": 2.5})
            vars['mode'].set(ml.get("mode", "from_outside"))
            vars['value'].set(ml.get("value", 2.5))

    def _read_eng_placement_range_fields(self):
        return {
            "min_size": self.eng_placement_range_min_var.get(),
            "max_size": self.eng_placement_range_max_var.get(),
            "engraving_location": {
                mat: {"mode": vars['mode'].get(), "value": vars['value'].get()}
                for mat, vars in self.eng_placement_range_loc_vars.items()
            },
        }

    def _add_eng_placement_range(self):
        r = self._read_eng_placement_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.eng_placement_ranges.append(r)
        self.eng_placement_ranges.sort(key=lambda x: x["min_size"])
        self.eng_placement_selected_range_index = self.eng_placement_ranges.index(r)
        self._refresh_eng_placement_range_combo()

    def _update_eng_placement_range(self):
        if self.eng_placement_selected_range_index is None or self.eng_placement_selected_range_index >= len(self.eng_placement_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to update."))
            return
        r = self._read_eng_placement_range_fields()
        if r["min_size"] >= r["max_size"]:
            messagebox.showerror(_("Invalid Range"), _("Min size must be less than max size."))
            return
        self.eng_placement_ranges[self.eng_placement_selected_range_index] = r
        self.eng_placement_ranges.sort(key=lambda x: x["min_size"])
        self.eng_placement_selected_range_index = self.eng_placement_ranges.index(r)
        self._refresh_eng_placement_range_combo()

    def _delete_eng_placement_range(self):
        if self.eng_placement_selected_range_index is None or self.eng_placement_selected_range_index >= len(self.eng_placement_ranges):
            messagebox.showinfo(_("No Selection"), _("Select a range to delete."))
            return
        del self.eng_placement_ranges[self.eng_placement_selected_range_index]
        self.eng_placement_selected_range_index = None
        self._refresh_eng_placement_range_combo()

    # ------------------------------------------------------------------
    # Apply / Cancel
    # ------------------------------------------------------------------

    def on_apply(self):
        """Apply form values to the live app.

        If the form is clean (matches the loaded preset / initial state),
        commit and close. If dirty, prompt the user to save the changes
        as a preset first — they can either run the save flow (then commit
        and close) or back out to keep editing.
        """
        if not self._form_is_valid(_("apply")):
            return
        if not self._is_dirty():
            self.save_options()
            return

        choice = messagebox.askyesno(
            _("Save changes as a preset?"),
            _("You have edits that aren't captured in any preset.\n\n"
            "Save them as a preset before applying?\n\n"
            "Yes  — open the Save Preset dialog.\n"
            "No   — go back to the sizing dialog without applying."),
            parent=self.top,
        )
        if not choice:
            return  # back to dialog, no commit
        # Run the save flow. If the user cancels it, don't apply either —
        # they're explicitly choosing not to save, so we treat this as
        # "back to dialog" rather than a silent apply of unsaved changes.
        before_baseline = self._baseline
        self.on_save_sizing_preset()
        if self._baseline is before_baseline:
            return  # save was cancelled inside the SaveSizingPresetDialog
        self.save_options()

    def on_cancel(self):
        """Close the dialog. If dirty, prompt to save as preset first."""
        if not self._is_dirty():
            self.top.destroy()
            return

        # Three-way prompt via two stacked questions: simpler than building
        # a custom 3-button dialog, and matches existing app conventions.
        save_first = messagebox.askyesnocancel(
            _("Unsaved changes"),
            _("You have edits that aren't captured in any preset.\n\n"
            "Yes     — save them as a preset, then close.\n"
            "No      — discard edits and close.\n"
            "Cancel  — keep editing."),
            parent=self.top,
        )
        if save_first is None:
            return  # keep editing
        if save_first:
            before_baseline = self._baseline
            self.on_save_sizing_preset()
            if self._baseline is before_baseline:
                return  # user cancelled the save flow → back to dialog
            self.top.destroy()
            return
        # Discard
        self.top.destroy()

    def save_options(self):
        # Sizing
        self.settings["units"] = self.unit_var.get()
        self.settings["felt_offset"] = self.felt_offset_var.get()
        self.settings["card_to_felt_offset"] = self.card_offset_var.get()
        self.settings["leather_wrap_multiplier"] = self.leather_mult_var.get()
        self.settings["min_hole_size"] = self.min_hole_size_var.get()
        self.settings["felt_thickness"] = self.felt_thickness_var.get()
        self.settings["felt_thickness_unit"] = self.felt_thickness_unit_var.get()
        self.settings["sizing_range_mode"] = self.sizing_range_mode_var.get()
        self.settings["sizing_ranges"] = self.sizing_ranges

        # DART SAVE LOGIC
        self.settings["darts_enabled"] = self.darts_enabled_var.get()
        self.settings["dart_range_mode"] = self.dart_range_mode_var.get()
        self.settings["dart_threshold"] = self.dart_threshold_var.get()
        self.settings["dart_overwrap"] = self.dart_overwrap_var.get()
        self.settings["dart_wrap_bonus"] = self.dart_wrap_bonus_var.get()
        self.settings["dart_frequency_multiplier"] = self.dart_frequency_multiplier_var.get()
        self.settings["dart_shape_factor"] = self.dart_shape_factor_var.get()
        self.settings["dart_ranges"] = self.dart_ranges
        
        # Engraving
        self.settings["engraving_on"] = self.engraving_on_var.get()
        for material, var in self.engraving_font_size_vars.items():
            self.settings["engraving_font_size"][material] = var.get()
        self.settings["engraving_settings_range_mode"] = self.eng_settings_range_mode_var.get()
        self.settings["engraving_settings_ranges"] = self.eng_settings_ranges

        for material, vars in self.engraving_loc_vars.items():
            self.settings["engraving_location"][material]['mode'] = vars['mode'].get()
            self.settings["engraving_location"][material]['value'] = vars['value'].get()
        self.settings["engraving_placement_range_mode"] = self.eng_placement_range_mode_var.get()
        self.settings["engraving_placement_ranges"] = self.eng_placement_ranges

        # STAR ENGRAVING SAVE
        self.settings["dart_engraving_on"] = self.dart_engraving_on_var.get()
            
        # Export
        self.settings["compatibility_mode"] = self.compatibility_mode_var.get()

        self.save_callback()
        self.update_callback()
        self.top.destroy()

    def revert_to_defaults(self):
        if messagebox.askyesno(_("Revert to Defaults"), _("Are you sure you want to revert all settings to their original defaults?")):
            # Sizing
            self.unit_var.set(DEFAULT_SETTINGS["units"])
            self.felt_offset_var.set(DEFAULT_SETTINGS["felt_offset"])
            self.card_offset_var.set(DEFAULT_SETTINGS["card_to_felt_offset"])
            self.leather_mult_var.set(DEFAULT_SETTINGS["leather_wrap_multiplier"])
            self.min_hole_size_var.set(DEFAULT_SETTINGS["min_hole_size"])
            self.felt_thickness_var.set(DEFAULT_SETTINGS["felt_thickness"])
            self.felt_thickness_unit_var.set(DEFAULT_SETTINGS["felt_thickness_unit"])
            self.sizing_range_mode_var.set("universal")
            self.sizing_ranges = []
            self._refresh_sizing_range_combo()
            self._toggle_sizing_mode()
            
            # REVERT DART LOGIC
            self.darts_enabled_var.set(DEFAULT_SETTINGS.get("darts_enabled", True))
            self.dart_range_mode_var.set("universal")
            self.dart_threshold_var.set(DEFAULT_SETTINGS.get("dart_threshold", 18.0))
            self.dart_overwrap_var.set(DEFAULT_SETTINGS.get("dart_overwrap", 0.5))
            self.dart_wrap_bonus_var.set(DEFAULT_SETTINGS.get("dart_wrap_bonus", 0.75))
            self.dart_frequency_multiplier_var.set(DEFAULT_SETTINGS.get("dart_frequency_multiplier", 1.0))
            self.dart_shape_factor_var.set(DEFAULT_SETTINGS.get("dart_shape_factor", 0.5))
            self.dart_ranges = []
            self._refresh_range_combo()
            self._toggle_dart_mode()
            
            # Engraving
            self.engraving_on_var.set(DEFAULT_SETTINGS["engraving_on"])
            for material, var in self.engraving_font_size_vars.items():
                 var.set(DEFAULT_SETTINGS["engraving_font_size"][material])
            self.eng_settings_range_mode_var.set("universal")
            self.eng_settings_ranges = []
            self._refresh_eng_settings_range_combo()
            self._toggle_eng_settings_mode()

            for material, vars in self.engraving_loc_vars.items():
                 vars['mode'].set(DEFAULT_SETTINGS["engraving_location"][material]['mode'])
                 vars['value'].set(DEFAULT_SETTINGS["engraving_location"][material]['value'])
            self.eng_placement_range_mode_var.set("universal")
            self.eng_placement_ranges = []
            self._refresh_eng_placement_range_combo()
            self._toggle_eng_placement_mode()
                 
            # Revert Star Engraving
            self.dart_engraving_on_var.set(True)

            # Export
            self.compatibility_mode_var.set(DEFAULT_SETTINGS.get("compatibility_mode", False))

    # ------------------------------------------------------------------
    # Sizing-rules preset support
    # ------------------------------------------------------------------
    def _capture_form_to_dict(self):
        """Snapshot the current form state into a sizing-preset dict."""
        return {
            # Sizing
            "units": self.unit_var.get(),
            "felt_offset": self.felt_offset_var.get(),
            "card_to_felt_offset": self.card_offset_var.get(),
            "leather_wrap_multiplier": self.leather_mult_var.get(),
            "min_hole_size": self.min_hole_size_var.get(),
            "felt_thickness": self.felt_thickness_var.get(),
            "felt_thickness_unit": self.felt_thickness_unit_var.get(),
            "sizing_range_mode": self.sizing_range_mode_var.get(),
            "sizing_ranges": copy.deepcopy(self.sizing_ranges),
            # Darts
            "darts_enabled": self.darts_enabled_var.get(),
            "dart_range_mode": self.dart_range_mode_var.get(),
            "dart_threshold": self.dart_threshold_var.get(),
            "dart_overwrap": self.dart_overwrap_var.get(),
            "dart_wrap_bonus": self.dart_wrap_bonus_var.get(),
            "dart_frequency_multiplier": self.dart_frequency_multiplier_var.get(),
            "dart_shape_factor": self.dart_shape_factor_var.get(),
            "dart_ranges": copy.deepcopy(self.dart_ranges),
            "dart_engraving_on": self.dart_engraving_on_var.get(),
            # Engraving
            "engraving_on": self.engraving_on_var.get(),
            "engraving_font_size": {m: v.get() for m, v in self.engraving_font_size_vars.items()},
            "engraving_settings_range_mode": self.eng_settings_range_mode_var.get(),
            "engraving_settings_ranges": copy.deepcopy(self.eng_settings_ranges),
            "engraving_location": {
                m: {"mode": v["mode"].get(), "value": v["value"].get()}
                for m, v in self.engraving_loc_vars.items()
            },
            "engraving_placement_range_mode": self.eng_placement_range_mode_var.get(),
            "engraving_placement_ranges": copy.deepcopy(self.eng_placement_ranges),
            # Export
            "compatibility_mode": self.compatibility_mode_var.get(),
        }

    def _apply_dict_to_form(self, source):
        """Set every form var from a sizing-preset dict.

        Missing keys fall back to DEFAULT_SETTINGS so older preset files
        keep working when the schema gains new fields.
        """
        d = DEFAULT_SETTINGS

        # Sizing
        self.unit_var.set(source.get("units", d["units"]))
        self.felt_offset_var.set(source.get("felt_offset", d["felt_offset"]))
        self.card_offset_var.set(source.get("card_to_felt_offset", d["card_to_felt_offset"]))
        self.leather_mult_var.set(source.get("leather_wrap_multiplier", d["leather_wrap_multiplier"]))
        self.min_hole_size_var.set(source.get("min_hole_size", d["min_hole_size"]))
        self.felt_thickness_var.set(source.get("felt_thickness", d["felt_thickness"]))
        self.felt_thickness_unit_var.set(source.get("felt_thickness_unit", d["felt_thickness_unit"]))
        self.sizing_range_mode_var.set(source.get("sizing_range_mode", "universal"))
        self.sizing_ranges = copy.deepcopy(list(source.get("sizing_ranges", [])))
        self._refresh_sizing_range_combo()
        self._toggle_sizing_mode()

        # Darts
        self.darts_enabled_var.set(source.get("darts_enabled", d.get("darts_enabled", True)))
        self.dart_range_mode_var.set(source.get("dart_range_mode", "universal"))
        self.dart_threshold_var.set(source.get("dart_threshold", d.get("dart_threshold", 18.0)))
        self.dart_overwrap_var.set(source.get("dart_overwrap", d.get("dart_overwrap", 0.5)))
        self.dart_wrap_bonus_var.set(source.get("dart_wrap_bonus", d.get("dart_wrap_bonus", 0.75)))
        self.dart_frequency_multiplier_var.set(source.get("dart_frequency_multiplier", d.get("dart_frequency_multiplier", 1.0)))
        self.dart_shape_factor_var.set(source.get("dart_shape_factor", d.get("dart_shape_factor", 0.5)))
        self.dart_ranges = copy.deepcopy(list(source.get("dart_ranges", [])))
        self._refresh_range_combo()
        self._toggle_dart_mode()
        self.dart_engraving_on_var.set(source.get("dart_engraving_on", True))

        # Engraving
        self.engraving_on_var.set(source.get("engraving_on", d["engraving_on"]))
        eng_fonts = source.get("engraving_font_size", d["engraving_font_size"])
        for material, var in self.engraving_font_size_vars.items():
            var.set(eng_fonts.get(material, d["engraving_font_size"].get(material, 3.0)))
        self.eng_settings_range_mode_var.set(source.get("engraving_settings_range_mode", "universal"))
        self.eng_settings_ranges = copy.deepcopy(list(source.get("engraving_settings_ranges", [])))
        self._refresh_eng_settings_range_combo()
        self._toggle_eng_settings_mode()

        eng_locs = source.get("engraving_location", d["engraving_location"])
        for material, vars in self.engraving_loc_vars.items():
            loc = eng_locs.get(material, d["engraving_location"].get(material, {"mode": "from_outside", "value": 2.5}))
            vars['mode'].set(loc.get('mode', 'from_outside'))
            vars['value'].set(loc.get('value', 2.5))
        self.eng_placement_range_mode_var.set(source.get("engraving_placement_range_mode", "universal"))
        self.eng_placement_ranges = copy.deepcopy(list(source.get("engraving_placement_ranges", [])))
        self._refresh_eng_placement_range_combo()
        self._toggle_eng_placement_mode()

        # Export
        self.compatibility_mode_var.set(source.get("compatibility_mode", d.get("compatibility_mode", False)))

    def _refresh_sizing_preset_combo(self, select=None):
        names = sorted(self.sizing_presets.keys())
        self.preset_combo['values'] = names
        if select and select in names:
            self.preset_combo.set(select)
        elif not names:
            self.preset_combo.set("")
        elif self.preset_combo.get() not in names:
            # e.g. after a delete: don't leave a vanished name displayed
            self.preset_combo.set("")

    def on_load_sizing_preset(self):
        name = self.preset_combo.get().strip()
        if not name:
            messagebox.showinfo(_("Load Preset"), _("Pick a preset from the dropdown first."), parent=self.top)
            return
        if name not in self.sizing_presets:
            messagebox.showerror(_("Load Preset"), _("Preset '{name}' not found.").format(name=name), parent=self.top)
            return
        # Warn before clobbering unsaved edits.
        if self._is_dirty():
            label = self.active_preset_name or "the current values"
            if not messagebox.askyesno(
                _("Discard unsaved changes?"),
                _("You have unsaved edits to {label}.\n\nLoading '{name}' will discard them. Continue?").format(label=label, name=name),
                parent=self.top,
            ):
                return
        self._apply_dict_to_form(self.sizing_presets[name])
        self.active_preset_name = name
        self._set_baseline_to_current()

    def on_save_sizing_preset(self):
        """Open the Save Preset dialog (overwrite existing or save as new)."""
        if not self._form_is_valid(_("save it as a preset")):
            return
        dlg = SaveSizingPresetDialog(
            self.top,
            existing_names=sorted(self.sizing_presets.keys()),
            default_existing=self.active_preset_name,
        )
        result = dlg.result
        # The child dialog's grab released ours when it closed (Tk has a
        # single global grab) — re-grab so this dialog stays modal.
        if self.top.winfo_exists():
            self.top.grab_set()
        if result is None:
            return  # user cancelled
        target = result["name"]
        self.sizing_presets[target] = self._capture_form_to_dict()
        self.sizing_presets_save_callback()
        self.active_preset_name = target
        self._set_baseline_to_current()
        self._refresh_sizing_preset_combo(select=target)

    def on_rename_sizing_preset(self):
        old = self.preset_combo.get().strip()
        if not old or old not in self.sizing_presets:
            messagebox.showinfo(_("Rename Preset"), _("Pick a preset from the dropdown first."), parent=self.top)
            return
        new = simpledialog.askstring(
            _("Rename Preset"),
            _("New name for '{old}':").format(old=old),
            initialvalue=old,
            parent=self.top,
        )
        # askstring's grab released ours when it closed — re-grab.
        if self.top.winfo_exists():
            self.top.grab_set()
        if new is None:
            return
        new = new.strip()
        if not new:
            messagebox.showwarning(_("Rename Preset"), _("Preset name cannot be empty."), parent=self.top)
            return
        if new == old:
            return
        if new in self.sizing_presets:
            messagebox.showerror(
                _("Rename Preset"),
                _("A preset named '{new}' already exists.").format(new=new),
                parent=self.top,
            )
            return
        self.sizing_presets[new] = self.sizing_presets.pop(old)
        if self.active_preset_name == old:
            self.active_preset_name = new
        self.sizing_presets_save_callback()
        self._refresh_sizing_preset_combo(select=new)

    def on_delete_sizing_preset(self):
        name = self.preset_combo.get().strip()
        if not name or name not in self.sizing_presets:
            messagebox.showinfo(_("Delete Preset"), _("Pick a preset from the dropdown first."), parent=self.top)
            return
        if len(self.sizing_presets) <= 1:
            messagebox.showinfo(
                _("Delete Preset"),
                _("At least one sizing preset must remain. Save another "
                "preset before deleting this one."),
                parent=self.top,
            )
            return
        if not messagebox.askyesno(_("Delete Preset"), _("Delete preset '{name}'?").format(name=name), parent=self.top):
            return
        del self.sizing_presets[name]
        if self.active_preset_name == name:
            self.active_preset_name = None
        self.sizing_presets_save_callback()
        self._refresh_sizing_preset_combo()

    def on_export_sizing_presets(self):
        if not self.sizing_presets:
            messagebox.showinfo(_("Export Presets"), _("There are no sizing presets to export."))
            return
        path = filedialog.asksaveasfilename(
            parent=self.top,
            title=_("Export Sizing Presets"),
            defaultextension=".json",
            filetypes=((_("JSON files"), "*.json"), (_("All files"), "*.*")),
            initialfile="sizing_presets_export.json",
        )
        if not path:
            return
        try:
            with open(path, 'w') as f:
                json.dump(self.sizing_presets, f, indent=2)
            messagebox.showinfo(
                _("Export Successful"),
                _("Exported {count} preset(s).").format(count=len(self.sizing_presets)),
            )
        except Exception as e:
            messagebox.showerror(_("Export Error"), _("Could not export presets:\n{e}").format(e=e))

    def on_import_sizing_presets(self):
        path = filedialog.askopenfilename(
            parent=self.top,
            title=_("Import Sizing Presets"),
            filetypes=((_("JSON files"), "*.json"), (_("All files"), "*.*")),
        )
        if not path:
            return
        try:
            with open(path, 'r') as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror(_("Import Error"), _("Could not read file:\n{e}").format(e=e))
            return
        if not isinstance(data, dict) or not data:
            messagebox.showerror(_("Import Error"), _("File does not contain any sizing presets."))
            return

        added = 0
        renamed = 0
        last_added = None
        for raw_name, preset in data.items():
            if not isinstance(preset, dict):
                continue
            # Drop any keys that aren't part of the schema so an unrelated
            # JSON file can't smuggle stray settings into the dialog.
            clean = {k: copy.deepcopy(v) for k, v in preset.items() if k in SIZING_PRESET_KEYS}
            if not clean:
                continue
            name = raw_name
            while name in self.sizing_presets:
                name += "*"
            if name != raw_name:
                renamed += 1
            self.sizing_presets[name] = clean
            last_added = name
            added += 1

        if added == 0:
            messagebox.showinfo(_("Import"), _("No valid sizing presets found in that file."))
            return

        self.sizing_presets_save_callback()
        self._refresh_sizing_preset_combo(select=last_added)
        msg = f"Imported {added} preset(s)."
        if renamed:
            msg += f"\nRenamed due to conflict: {renamed}."
        messagebox.showinfo(_("Import Successful"), msg)


class PadPreviewWindow(tk.Toplevel):
    """Live, resizable preview of a pad based on the parent OptionsWindow form.

    The user picks a pad size and which materials to show (leather / felt /
    card / exact size). The window draws each material onto a tk.Canvas
    using the same geometry helpers as svg_engine, so what you see here
    matches what the SVG / G-code will cut. A 200 ms polling tick re-reads
    the parent form and re-renders whenever any value has changed.
    """

    # Visually distinct, on-paper-readable colors per material.
    COLORS = {
        'leather':    '#6B4423',
        'felt':       '#A33333',
        'card':       '#C4A484',
        'exact_size': '#444444',
    }
    LABELS = {
        'leather':    'Leather',
        'felt':       'Felt',
        'card':       'Card',
        'exact_size': 'Exact size',
    }
    MATERIALS = ('leather', 'felt', 'card', 'exact_size')
    POLL_MS = 200

    def __init__(self, parent_options):
        super().__init__(parent_options.top)
        self.parent_options = parent_options
        self.title(_("Pad Preview"))
        self.geometry("620x520")
        self.minsize(380, 320)
        self.configure(bg=DIALOG_BG)
        self.resizable(True, True)
        self.transient(parent_options.top)

        # Local controls
        self.preview_size_var = tk.DoubleVar(value=18.0)
        self.show_vars = {
            'leather':    tk.BooleanVar(value=True),
            'felt':       tk.BooleanVar(value=True),
            'card':       tk.BooleanVar(value=True),
            'exact_size': tk.BooleanVar(value=False),
        }
        self.layout_var = tk.StringVar(value='layered')

        # Polling state
        self._last_form_snapshot = None
        self._poll_after_id = None

        self._build_widgets()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        # Trigger a render after the geometry has settled.
        self.after(50, self._render)
        self._schedule_poll()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_widgets(self):
        ctrls = tk.Frame(self, bg=DIALOG_BG, padx=10, pady=8)
        ctrls.pack(fill="x")

        # Pad-size row
        size_row = tk.Frame(ctrls, bg=DIALOG_BG)
        size_row.pack(fill="x", pady=(0, 4))
        tk.Label(size_row, text=_("Preview pad size:"), bg=DIALOG_BG).pack(side="left")
        size_ent = tk.Entry(size_row, textvariable=self.preview_size_var, width=8)
        size_ent.pack(side="left", padx=5)
        tk.Label(size_row, text=_("mm"), bg=DIALOG_BG).pack(side="left")
        add_tooltip(
            size_ent,
            _("Diameter (in mm) of the pad to render. The preview applies "
            "the parent dialog's sizing rules to this size."),
        )
        # Live update on size edit
        self.preview_size_var.trace_add("write", lambda *_a: self._render())

        # Materials row
        mat_row = tk.Frame(ctrls, bg=DIALOG_BG)
        mat_row.pack(fill="x", pady=(0, 4))
        tk.Label(mat_row, text=_("Materials:"), bg=DIALOG_BG).pack(side="left", padx=(0, 5))
        for mat in self.MATERIALS:
            cb = tk.Checkbutton(
                mat_row, text=self.LABELS[mat],
                variable=self.show_vars[mat], bg=DIALOG_BG,
                command=self._render,
            )
            cb.pack(side="left", padx=2)

        # Layout row
        layout_row = tk.Frame(ctrls, bg=DIALOG_BG)
        layout_row.pack(fill="x", pady=(0, 4))
        tk.Label(layout_row, text=_("Layout:"), bg=DIALOG_BG).pack(side="left")
        layered_rb = tk.Radiobutton(
            layout_row, text=_("Layered (concentric)"),
            variable=self.layout_var, value='layered', bg=DIALOG_BG,
            command=self._render,
        )
        layered_rb.pack(side="left", padx=5)
        side_rb = tk.Radiobutton(
            layout_row, text=_("Side by side"),
            variable=self.layout_var, value='side_by_side', bg=DIALOG_BG,
            command=self._render,
        )
        side_rb.pack(side="left", padx=5)
        add_tooltip(
            layered_rb,
            _("Stack each material on the same center, like the layers of "
            "a finished pad. Smaller materials draw on top."),
        )
        add_tooltip(
            side_rb,
            _("Lay the discs out in a row at consistent scale, so you can "
            "compare diameters at a glance."),
        )

        # Canvas
        canvas_frame = tk.Frame(self, bg=DIALOG_BG)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.canvas = tk.Canvas(
            canvas_frame, bg="#FFFFFF",
            highlightthickness=1, highlightbackground="#888888",
        )
        self.canvas.pack(fill="both", expand=True)
        self.canvas.bind("<Configure>", lambda _e: self._render())

    # ------------------------------------------------------------------
    # Polling — re-render whenever the parent form changes.
    # ------------------------------------------------------------------

    def _schedule_poll(self):
        self._poll_after_id = self.after(self.POLL_MS, self._poll)

    def _poll(self):
        if not self.winfo_exists():
            return
        try:
            snap = self.parent_options._capture_form_to_dict()
        except (tk.TclError, ValueError, KeyError, AttributeError):
            snap = None  # mid-edit (empty entry, etc.) — skip this tick
        if snap is not None and snap != self._last_form_snapshot:
            self._last_form_snapshot = snap
            self._render()
        self._poll_after_id = self.after(self.POLL_MS, self._poll)

    def _cancel_poll(self):
        if self._poll_after_id is not None:
            try:
                self.after_cancel(self._poll_after_id)
            except (tk.TclError, ValueError):
                pass
            self._poll_after_id = None

    def _on_close(self):
        self._cancel_poll()
        # Sync the parent toggle so the checkbox reflects the closed state.
        try:
            self.parent_options.show_preview_var.set(False)
            self.parent_options.preview_window = None
        except (tk.TclError, AttributeError):
            pass
        self.destroy()

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------

    def _render(self):
        if not self.winfo_exists():
            return
        self.canvas.delete("all")

        try:
            pad_size = float(self.preview_size_var.get())
        except (tk.TclError, ValueError):
            self._draw_message("Enter a pad size to preview")
            return
        if pad_size <= 0:
            self._draw_message("Pad size must be positive")
            return

        try:
            settings = self.parent_options._capture_form_to_dict()
        except (tk.TclError, ValueError, KeyError, AttributeError):
            self._draw_message("Waiting for valid settings…")
            return

        materials = [m for m in self.MATERIALS if self.show_vars[m].get()]
        if not materials:
            self._draw_message("No materials selected")
            return

        # Compute disc diameters for each requested material.
        diams = {}
        for mat in materials:
            try:
                d = get_disc_diameter(pad_size, mat, settings)
            except Exception:
                d = 0
            if d and d > 0:
                diams[mat] = d
        if not diams:
            self._draw_message("No materials produce a valid disc at this size")
            return

        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)

        if self.layout_var.get() == 'layered':
            self._draw_layered(pad_size, settings, diams, cw, ch)
        else:
            self._draw_side_by_side(pad_size, settings, diams, cw, ch)

    def _draw_message(self, text):
        cw = max(self.canvas.winfo_width(), 1)
        ch = max(self.canvas.winfo_height(), 1)
        self.canvas.create_text(cw / 2, ch / 2, text=text,
                                fill="#888888", font=("Helvetica", 10))

    # --- Layered (concentric) ---

    def _draw_layered(self, pad_size, settings, diams, cw, ch):
        margin = 30
        avail = max(min(cw, ch) - 2 * margin, 50)
        largest_mm = max(diams.values())
        scale = avail / largest_mm  # px per mm
        cx = cw / 2
        cy = ch / 2

        # Draw largest first so smaller materials end up visually "on top".
        for mat in sorted(self.MATERIALS,
                          key=lambda m: -diams.get(m, 0)):
            if mat not in diams:
                continue
            d_mm = diams[mat]
            r_px = (d_mm / 2) * scale
            color = self.COLORS[mat]
            if mat == 'leather':
                self._draw_leather(cx, cy, pad_size, settings, r_px, color, scale)
            else:
                self.canvas.create_oval(
                    cx - r_px, cy - r_px, cx + r_px, cy + r_px,
                    outline=color, width=2, fill='',
                )
            self._draw_center_hole(cx, cy, mat, pad_size, settings, scale, color)

        self._draw_legend(diams)

    # --- Side by side ---

    def _draw_side_by_side(self, pad_size, settings, diams, cw, ch):
        margin = 30
        gap_mm = 5.0
        # Stable left-to-right order: largest material first
        ordered = sorted(diams.items(), key=lambda kv: -kv[1])

        total_w_mm = sum(d for _, d in ordered) + gap_mm * (len(ordered) - 1)
        max_h_mm = max(d for _, d in ordered)
        avail_w = max(cw - 2 * margin, 50)
        avail_h = max(ch - 2 * margin - 26, 50)  # leave room for labels
        scale = min(avail_w / total_w_mm, avail_h / max_h_mm)

        cy = ch / 2
        cur_left_mm = -total_w_mm / 2  # in mm relative to canvas center
        for mat, d_mm in ordered:
            r_mm = d_mm / 2
            cx_mm = cur_left_mm + r_mm
            cx_px = cw / 2 + cx_mm * scale
            r_px = r_mm * scale
            color = self.COLORS[mat]
            if mat == 'leather':
                self._draw_leather(cx_px, cy, pad_size, settings, r_px, color, scale)
            else:
                self.canvas.create_oval(
                    cx_px - r_px, cy - r_px, cx_px + r_px, cy + r_px,
                    outline=color, width=2, fill='',
                )
            self._draw_center_hole(cx_px, cy, mat, pad_size, settings, scale, color)
            self.canvas.create_text(
                cx_px, cy + r_px + 14,
                text=_("{label}\n{mm:.1f} mm").format(label=self.LABELS[mat], mm=d_mm),
                fill=color, justify="center",
                font=("Helvetica", 9),
            )
            cur_left_mm += d_mm + gap_mm

    # --- Leather (with optional dart pattern) ---

    def _draw_leather(self, cx, cy, pad_size, settings, outer_r_px, color, scale):
        dart_cfg = get_dart_settings_for_size(pad_size, settings)
        if not dart_cfg:
            self.canvas.create_oval(
                cx - outer_r_px, cy - outer_r_px, cx + outer_r_px, cy + outer_r_px,
                outline=color, width=2, fill='',
            )
            return

        # Replicate svg_engine's dart geometry exactly.
        sizing = get_sizing_for_size(pad_size, settings)
        felt_thick = get_felt_thickness_mm(settings, sizing)
        overwrap = dart_cfg.get("overwrap", 0.5)
        felt_r_mm = (pad_size - sizing["felt_offset"]) / 2
        inner_r_mm = felt_r_mm + felt_thick + overwrap
        outer_r_mm = get_disc_diameter(pad_size, 'leather', settings) / 2
        if inner_r_mm >= outer_r_mm:
            inner_r_mm = max(outer_r_mm - 0.2, 0.1)

        circumference = 2 * math.pi * inner_r_mm
        freq_mult = dart_cfg.get("frequency_multiplier", 1.0)
        num_points = int((circumference / 3.5) * freq_mult)
        if num_points < 12:
            num_points = 12
        if num_points % 2 != 0:
            num_points += 1

        shape_factor = dart_cfg.get("shape_factor", 0.5)
        steps = max(int(num_points * 8), 64)
        avg_r_px = (outer_r_mm + inner_r_mm) * 0.5 * scale
        amp_px = (outer_r_mm - inner_r_mm) * 0.5 * scale

        coords = []
        for i in range(steps + 1):
            theta = 2 * math.pi * i / steps
            raw = math.cos(num_points * theta)
            shaped = _wave_value(raw, shape_factor)
            r = avg_r_px + amp_px * shaped
            x = cx + r * math.cos(theta)
            y = cy + r * math.sin(theta)
            coords.extend([x, y])
        # Use create_line on a closed loop (polygon would auto-fill its
        # interior, smoothing the perceived shape; we want crisp outlines).
        self.canvas.create_line(*coords, fill=color, width=2, smooth=False)

    # --- Center hole (info-only; matches should_have_center_hole rules) ---

    def _draw_center_hole(self, cx, cy, material, pad_size, settings, scale, color):
        if material == 'exact_size':
            return  # no center hole on exact-size discs
        sizing = get_sizing_for_size(pad_size, settings)
        min_size = sizing.get("min_hole_size", 16.5)
        if pad_size < min_size:
            return
        # Use 3 mm as a conventional preview hole size — the real hole
        # diameter is selected on the Pad Maker tab, not in Sizing Rules,
        # so we just draw an indicator dot that won't be confused for a
        # cut path.
        hole_r_px = (3.0 / 2) * scale
        if hole_r_px < 1.5:
            hole_r_px = 1.5
        self.canvas.create_oval(
            cx - hole_r_px, cy - hole_r_px, cx + hole_r_px, cy + hole_r_px,
            outline=color, width=1, fill='',
        )

    # --- Legend ---

    def _draw_legend(self, diams):
        y = 10
        for mat in self.MATERIALS:
            if mat not in diams:
                continue
            color = self.COLORS[mat]
            self.canvas.create_rectangle(
                10, y + 2, 22, y + 12, fill=color, outline=color,
            )
            text = f"{self.LABELS[mat]} — {diams[mat]:.1f} mm"
            self.canvas.create_text(
                28, y + 7, text=text, fill="#222222",
                anchor="w", font=("Helvetica", 9),
            )
            y += 16


class SaveSizingPresetDialog(tk.Toplevel):
    """Two-mode save dialog: overwrite an existing preset, or save as new.

    On close, sets `self.result` to None (cancelled) or {"name": str}.
    """

    def __init__(self, parent, existing_names, default_existing=None,
                 title=None, intro=None):
        super().__init__(parent)
        self.title(title or _("Save Sizing Preset"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)
        self.result = None

        existing_names = list(existing_names)
        has_existing = bool(existing_names)
        # Default to "new" if there's nothing to overwrite or no active
        # preset; default to "overwrite" otherwise so the common case
        # (load → tweak → re-save) is one click.
        initial = "overwrite" if (has_existing and default_existing in existing_names) else "new"
        self.mode_var = tk.StringVar(value=initial)
        self.existing_var = tk.StringVar(
            value=default_existing if default_existing in existing_names
                  else (existing_names[0] if existing_names else "")
        )
        self.new_name_var = tk.StringVar(value="")

        frame = tk.Frame(self, bg=DIALOG_BG, padx=15, pady=12)
        frame.pack(fill="both", expand=True)

        tk.Label(
            frame,
            text=intro or _("Save the current sizing-rules form as a preset."),
            bg=DIALOG_BG, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        # --- Overwrite row ---
        overwrite_row = tk.Frame(frame, bg=DIALOG_BG)
        overwrite_row.pack(fill="x", pady=2)
        rb_over = tk.Radiobutton(
            overwrite_row, text=_("Overwrite existing:"), variable=self.mode_var,
            value="overwrite", bg=DIALOG_BG,
            command=self._sync_state,
        )
        rb_over.pack(side="left")
        self.existing_combo = ttk.Combobox(
            overwrite_row, textvariable=self.existing_var,
            values=existing_names, state="readonly", width=24,
        )
        self.existing_combo.pack(side="left", padx=8)

        if not has_existing:
            rb_over.configure(state="disabled")
            self.existing_combo.configure(state="disabled")

        # --- New row ---
        new_row = tk.Frame(frame, bg=DIALOG_BG)
        new_row.pack(fill="x", pady=2)
        rb_new = tk.Radiobutton(
            new_row, text=_("Save as new preset:"), variable=self.mode_var,
            value="new", bg=DIALOG_BG,
            command=self._sync_state,
        )
        rb_new.pack(side="left")
        self.new_entry = tk.Entry(new_row, textvariable=self.new_name_var, width=24)
        self.new_entry.pack(side="left", padx=8)

        # --- Buttons ---
        btn_row = tk.Frame(frame, bg=DIALOG_BG)
        btn_row.pack(fill="x", pady=(12, 0))
        save_btn = tk.Button(btn_row, text=_("Save"), command=self._on_save, width=10)
        save_btn.pack(side="right", padx=(5, 0))
        cancel_btn = tk.Button(btn_row, text=_("Cancel"), command=self.destroy, width=10)
        cancel_btn.pack(side="right")

        self.bind("<Return>", lambda _e: self._on_save())
        self.bind("<Escape>", lambda _e: self.destroy())

        # Existing names cached for collision checks
        self._existing = existing_names
        self._sync_state()
        self.update_idletasks()
        self.geometry("")
        if initial == "new":
            self.new_entry.focus_set()
        else:
            self.existing_combo.focus_set()

        self.wait_window(self)

    def _sync_state(self):
        if self.mode_var.get() == "overwrite":
            self.new_entry.configure(state="disabled")
            self.existing_combo.configure(state="readonly" if self._existing else "disabled")
        else:
            self.new_entry.configure(state="normal")
            self.existing_combo.configure(state="disabled")
            self.new_entry.focus_set()

    def _on_save(self):
        if self.mode_var.get() == "overwrite":
            target = self.existing_var.get().strip()
            if not target:
                messagebox.showinfo(_("Save Preset"), _("Pick a preset to overwrite."), parent=self)
                return
            if not messagebox.askyesno(
                _("Overwrite Preset"),
                _("Overwrite '{target}' with the current values?").format(target=target),
                parent=self,
            ):
                return
            self.result = {"name": target}
            self.destroy()
            return

        name = self.new_name_var.get().strip()
        if not name:
            messagebox.showwarning(_("Save Preset"), _("Preset name cannot be empty."), parent=self)
            return
        if name in self._existing:
            messagebox.showerror(
                _("Save Preset"),
                _("A preset named '{name}' already exists.\nUse Overwrite, or choose a different name.").format(name=name),
                parent=self,
            )
            return
        self.result = {"name": name}
        self.destroy()


class LayerColorWindow:
    def __init__(self, parent, settings, save_callback):
        self.settings = settings
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title(_("Layer Color Mapping"))
        self.top.geometry("450x420")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        self.color_map = {name: hex_val for name, hex_val in LIGHTBURN_COLORS}
        color_names = list(self.color_map.keys())
        self.hex_to_name_map = {hex_val: name for name, hex_val in LIGHTBURN_COLORS}

        self.color_vars = {}

        main_frame = tk.Frame(self.top, bg=DIALOG_BG, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        main_frame.columnconfigure(1, weight=1)

        layer_map_keys = [
            'felt_outline', 'felt_center_hole', 'felt_engraving',
            'card_outline', 'card_center_hole', 'card_engraving',
            'leather_outline', 'leather_center_hole', 'leather_engraving',
            'exact_size_outline', 'exact_size_center_hole', 'exact_size_engraving'
        ]
        
        op_word = {"outline": "outer cut",
                   "center_hole": "center hole",
                   "engraving": "engraving"}
        for i, key in enumerate(layer_map_keys):
            label_text = key.replace('_', ' ').capitalize() + ":"
            row_lbl = tk.Label(main_frame, text=label_text, bg=DIALOG_BG)
            row_lbl.grid(row=i, column=0, sticky='w', pady=3)

            var = tk.StringVar()
            current_hex = self.settings["layer_colors"].get(key, "#000000")
            current_name = self.hex_to_name_map.get(current_hex, color_names[0])
            var.set(current_name)

            combo = ttk.Combobox(main_frame, textvariable=var, values=color_names, state="readonly")
            combo.grid(row=i, column=1, sticky='ew', padx=5)
            self.color_vars[key] = var

            # Build a useful description: e.g. "felt outer cut" -> LightBurn layer color
            mat = key.split('_', 1)[0]
            op_key = key.split('_', 1)[1] if '_' in key else key
            op_name = op_word.get(op_key, op_key.replace('_', ' '))
            tip = (f"LightBurn layer color used for the {mat} {op_name} "
                   "lines in the SVG output.")
            add_tooltips(tip, row_lbl, combo)

        button_frame = tk.Frame(main_frame, bg=DIALOG_BG)
        button_frame.grid(row=len(layer_map_keys), column=0, columnspan=2, pady=20)
        save_btn = tk.Button(button_frame, text=_("Save"), command=self.save_colors)
        save_btn.pack(side="left", padx=10)
        cancel_btn = tk.Button(button_frame, text=_("Cancel"), command=self.top.destroy)
        cancel_btn.pack(side="left", padx=10)
        add_tooltip(save_btn, _("Apply the color choices and close."))
        add_tooltip(cancel_btn, _("Close without changing layer colors."))

    def save_colors(self):
        for key, var in self.color_vars.items():
            selected_name = var.get()
            self.settings["layer_colors"][key] = self.color_map[selected_name]
        
        self.save_callback()
        self.top.destroy()

class KeyLayoutWindow:
    def __init__(self, parent, settings, update_callback, save_callback):
        self.settings = settings
        self.update_callback = update_callback 
        self.save_callback = save_callback
        
        self.top = tk.Toplevel(parent)
        self.top.title(_("Key Height Layout Options"))
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()
        self.top.geometry("350x450") 

        # --- Main Layout Frames ---
        bottom_button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        bottom_button_frame.pack(side="bottom", fill="x", pady=10, padx=10)

        kl_save = tk.Button(bottom_button_frame, text=_("Save"), command=self.save_options)
        kl_save.pack(side="left", padx=5)
        kl_cancel = tk.Button(bottom_button_frame, text=_("Cancel"), command=self.top.destroy)
        kl_cancel.pack(side="left", padx=5)
        add_tooltip(kl_save, _("Apply the layout changes and close."))
        add_tooltip(kl_cancel, _("Close without changing the layout."))

        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(side="top", fill="both", expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG, padx=10, pady=10)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, self.canvas)

        self.key_layout_vars = {}
        # Get a copy to modify
        self.key_layout_settings = DEFAULT_SETTINGS["key_layout"].copy()
        self.key_layout_settings.update(self.settings.get("key_layout", {}))

        # --- Create Widgets ---
        info_frame = tk.LabelFrame(self.scrollable_frame, text=_("Horn Info Layout"), bg=DIALOG_BG, padx=5, pady=5)
        info_frame.pack(fill="x", pady=5)

        # Serial Number Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("show_serial", False))
        self.key_layout_vars["show_serial"] = var
        serial_cb = tk.Checkbutton(info_frame, text=_("Show 'Serial' field"), variable=var, bg=DIALOG_BG)
        serial_cb.pack(anchor='w')
        add_tooltip(serial_cb,
                    _("Show a Serial field in the horn-info header for each "
                    "key-height set, so you can record the saxophone's "
                    "serial number alongside the measurements."))

        # Large Notes Checkbox
        var = tk.BooleanVar(value=self.key_layout_settings.get("large_notes", False))
        self.key_layout_vars["large_notes"] = var
        notes_cb = tk.Checkbutton(info_frame, text=_("Use large 'Notes' field"), variable=var, bg=DIALOG_BG)
        notes_cb.pack(anchor='w')
        add_tooltip(notes_cb,
                    _("Use a multi-line Notes field instead of a single line, "
                    "good for jotting setup observations or repair history."))

        keys_frame = tk.LabelFrame(self.scrollable_frame, text=_("Visible Key Heights"), bg=DIALOG_BG, padx=5, pady=5)
        keys_frame.pack(fill="x", pady=5, expand=True)

        for key_name in ALL_KEY_HEIGHT_FIELDS:
            setting_key = f"show_{key_name.replace(' ', '_')}"
            var = tk.BooleanVar(value=self.key_layout_settings.get(setting_key, True))
            self.key_layout_vars[setting_key] = var
            kh_cb = tk.Checkbutton(keys_frame, text=_("Show '{key_name}' field").format(key_name=key_name), variable=var, bg=DIALOG_BG)
            kh_cb.pack(anchor='w')
            add_tooltip(kh_cb,
                        _("Show the '{key_name}' measurement field on key-height sets. Hide fields you never measure to keep the layout compact.").format(key_name=key_name))

    def save_options(self):
        for key, var in self.key_layout_vars.items():
            self.key_layout_settings[key] = var.get()
        
        self.settings["key_layout"] = self.key_layout_settings
        
        self.save_callback()
        self.update_callback() 
        self.top.destroy()
        
class ResonanceWindow(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        self.parent = parent
        
        self.title(_("Resonance Chamber"))
        self.geometry("400x200")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        main_frame = tk.Frame(self, bg=DIALOG_BG)
        main_frame.pack(expand=True)

        res_button = tk.Button(main_frame, text=_("Add Resonance"), command=self.start_resonance, font=("Helvetica", 14, "bold"))
        res_button.pack(pady=20, padx=40, ipadx=10, ipady=10)

    def start_resonance(self):
        self.withdraw()
        ResonanceProgressDialog(self.parent, self.settings, self.save_callback, self.theme_callback)
        self.destroy()

class ResonanceProgressDialog(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        self.parent_app = parent
        
        self.title(_("Optimizing..."))
        self.geometry("300x100")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        
        tk.Label(self, text=_("Applying resonance..."), bg=DIALOG_BG).pack(pady=10)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=5)
        
        self.update_progress(0)

    def update_progress(self, val):
        if not self.winfo_exists():
            return  # closed mid-animation — stop the after() chain
        self.progress['value'] = val
        if val < 100:
            self.after(70, self.update_progress, val + 1)
        else:
            self.after(200, self.finish_resonance)

    def finish_resonance(self):
        if not self.winfo_exists():
            return
        clicks = self.settings.get("resonance_clicks", 0) + 1
        self.settings["resonance_clicks"] = clicks
        
        if clicks >= 100:
            messagebox.showinfo(_("Power Overwhelming"), _("You have become too powerful."))
            self.destroy() 
            UninstallResonanceDialog(self.parent_app, self.settings, self.save_callback, self.theme_callback)
        else:
            self.save_callback()
            messagebox.showinfo(_("Success"), random.choice(get_resonance_messages()))
            self.theme_callback()
            self.destroy()

class UninstallResonanceDialog(tk.Toplevel):
    def __init__(self, parent, settings, save_callback, theme_callback):
        super().__init__(parent)
        self.settings = settings
        self.save_callback = save_callback
        self.theme_callback = theme_callback
        
        self.title(_("Resetting..."))
        self.geometry("300x100")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set() 

        tk.Label(self, text=_("Uninstalling resonance..."), bg=DIALOG_BG).pack(pady=10)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=5)
        self.update_progress(0)

    def update_progress(self, val):
        if not self.winfo_exists():
            return  # closed mid-animation — stop the after() chain
        self.progress['value'] = val
        if val < 100:
            self.after(20, self.update_progress, val + 1)
        else:
            self.after(200, self.finish_uninstall)

    def finish_uninstall(self):
        if not self.winfo_exists():
            return
        self.settings["resonance_clicks"] = 0
        self.save_callback()
        self.theme_callback()
        self.destroy()

class ExportPresetsWindow(tk.Toplevel):
    def __init__(self, parent, presets, title, default_filename, ask_provenance=False):
        super().__init__(parent)
        self.presets = presets
        self._title_text = title  # self.title is the Wm method, not a str
        self.title(title)
        self.default_filename = default_filename
        self.ask_provenance = ask_provenance
        self.geometry("400x500")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text=_("Select sets to export:"), bg=DIALOG_BG, font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text=_("Select All"), command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text=_("Select None"), command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not presets:
             tk.Label(self.scrollable_frame, text=_("No local sets found."), bg=DIALOG_BG).pack(pady=10)
        else:
            # Check if this is a nested dictionary (Key Libraries)
            if any(isinstance(v, dict) for v in presets.values()):
                for lib_name in sorted(self.presets.keys()):
                    tk.Label(self.scrollable_frame, text=f"[{lib_name}]", bg=DIALOG_BG, font=("Helvetica", 10, "bold")).pack(anchor='w', pady=(5,0))
                    for preset_name in sorted(self.presets[lib_name].keys()):
                        var = tk.BooleanVar()
                        full_name = f"{lib_name}::{preset_name}" # Internal delimiter
                        cb = tk.Checkbutton(self.scrollable_frame, text=f"  {preset_name}", variable=var, bg=DIALOG_BG)
                        cb.pack(anchor='w')
                        self.vars[full_name] = var
            else: # Flat dictionary (Pad Presets)
                for name in sorted(self.presets.keys()):
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg=DIALOG_BG)
                    cb.pack(anchor='w')
                    self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        bind_mousewheel(self, self.canvas)

        export_button = tk.Button(self, text=_("Export Selected"), command=self.export_selected, font=("Helvetica", 10, "bold"))
        export_button.pack(pady=10)

    def select_all(self):
        for var in self.vars.values():
            var.set(True)

    def select_none(self):
        for var in self.vars.values():
            var.set(False)

    def export_selected(self):
        to_export = {}
        selected_count = 0
        last_selected_data = None
        
        is_nested = any("::" in k for k in self.vars.keys())

        for name, var in self.vars.items():
            if var.get():
                if is_nested:
                    lib_name, preset_name = name.split("::", 1)
                    preset_data = self.presets[lib_name][preset_name]
                    to_export[f"[{lib_name}] {preset_name}"] = preset_data
                    selected_count += 1
                    last_selected_data = preset_data
                else:
                    to_export[name] = self.presets[name]
                    selected_count += 1
        
        if not to_export:
            messagebox.showwarning(_("No Selection"), _("Please select at least one set to export."))
            return

        initialfile = self.default_filename
        
        if self.ask_provenance:
            user_name = simpledialog.askstring(_("Provenance"), _("Enter your name (for filename):"))
            if not user_name:
                user_name = "Export" # Default if cancelled
            user_name = user_name.replace(" ", "_")
            
            if selected_count == 1 and last_selected_data:
                try:
                    make = last_selected_data.get("make", "UnknownMake").replace(" ", "_")
                    model = last_selected_data.get("model", "UnknownModel").replace(" ", "_")
                    size = last_selected_data.get("size", "UnknownSize").replace(" ", "_")
                    initialfile = f"{make}_{model}_{size}_{user_name}.json"
                except Exception:
                    initialfile = f"key_height_export_{user_name}.json"
            else:
                initialfile = f"key_height_export_{user_name}.json"

        filepath = filedialog.asksaveasfilename(
            title=_("Save {title} As...").format(title=self._title_text),
            defaultextension=".json",
            filetypes=((_("JSON files"), "*.json"), (_("All files"), "*.*")),
            initialfile=initialfile
        )
        
        if not filepath:
            return

        try:
            with open(filepath, 'w') as f:
                json.dump(to_export, f, indent=2)
            messagebox.showinfo(_("Export Successful"), _("Successfully exported {count} sets.").format(count=len(to_export)))
            self.destroy()
        except Exception as e:
            messagebox.showerror(_("Export Error"), _("Could not export presets:\n{e}").format(e=e))

class ImportPresetsWindow(tk.Toplevel):
    def __init__(self, parent, local_presets_lib, imported_presets, file_path, menu_widget, app_instance, preset_type_name="Preset", save_data=None):
        super().__init__(parent)
        self.parent_app = app_instance
        self.local_presets_lib = local_presets_lib 
        self.imported_presets = imported_presets 
        self.file_path = file_path
        self.menu_widget = menu_widget
        self.preset_type_name = preset_type_name
        self.save_data = save_data if save_data is not None else local_presets_lib
        
        self.title(_("Import {preset_type_name}s").format(preset_type_name=preset_type_name))
        self.geometry("450x500")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.vars = {}

        tk.Label(self, text=_("Select {preset_type_name}s to import:").format(preset_type_name=preset_type_name), bg=DIALOG_BG, font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text=_("Select All"), command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text=_("Select None"), command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(list_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG)

        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        if not imported_presets:
             tk.Label(self.scrollable_frame, text=_("No presets found in file."), bg=DIALOG_BG).pack(pady=10)
        else:
            for name in sorted(self.imported_presets.keys()):
                var = tk.BooleanVar(value=True) # Default to selected
                cb = tk.Checkbutton(self.scrollable_frame, text=name, variable=var, bg=DIALOG_BG)
                cb.pack(anchor='w')
                self.vars[name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        bind_mousewheel(self, self.canvas)

        import_button = tk.Button(self, text=_("Import Selected"), command=self.import_selected, font=("Helvetica", 10, "bold"))
        import_button.pack(pady=10)

    def select_all(self):
        for var in self.vars.values():
            var.set(True)

    def select_none(self):
        for var in self.vars.values():
            var.set(False)

    def import_selected(self):
        added_count = 0
        renamed_count = 0
        
        for name, var in self.vars.items():
            if var.get():
                preset_data = self.imported_presets[name]
                new_name = name
                
                # Handle bracketed library names from key set exports
                if new_name.startswith("[") and "] " in new_name:
                    try:
                        new_name = new_name.split("] ", 1)[1]
                    except Exception:
                        pass 
                        
                while new_name in self.local_presets_lib:
                    new_name += "*"
                
                if new_name != name:
                    renamed_count += 1
                
                self.local_presets_lib[new_name] = preset_data
                added_count += 1
        
        if added_count > 0:
            if save_presets(self.save_data, self.file_path):
                # Special refresh
                if self.preset_type_name == "Key Height Set":
                    self.parent_app.update_key_library_dropdown()
                else:
                    self.parent_app.update_pad_library_dropdown()
                
                messagebox.showinfo(_("Import Successful"), 
                                  _("Import complete.\n\nAdded: {added} presets\nRenamed due to conflicts: {renamed} presets").format(added=added_count, renamed=renamed_count))
            else:
                messagebox.showerror(_("Import Error"), _("Could not save new presets to file."))
        else:
            messagebox.showinfo(_("Import Complete"), _("No new presets were imported."))
            
        self.destroy()


class WebImportPresetsWindow(tk.Toplevel):
    """Import dialog for web-fetched presets grouped by library."""
    def __init__(self, parent, web_data, local_presets, file_path, app_instance):
        super().__init__(parent)
        self.web_data = web_data
        self.local_presets = local_presets
        self.file_path = file_path
        self.parent_app = app_instance

        self.title(_("Import Matt's Pad Sets"))
        self.geometry("500x550")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.vars = {}
        self.lib_vars = {}  # library-level toggle vars

        tk.Label(self, text=_("Select pad sets to import:"), bg=DIALOG_BG,
                 font=("Helvetica", 12)).pack(pady=10)

        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=5)
        tk.Button(button_frame, text=_("Select All"), command=self.select_all).pack(side="left", padx=5)
        tk.Button(button_frame, text=_("Select None"), command=self.select_none).pack(side="left", padx=5)

        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.canvas = tk.Canvas(list_frame, bg=DIALOG_BG, highlightthickness=0)
        self.scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg=DIALOG_BG)

        self.scrollable_frame.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        for lib_name in sorted(web_data.keys()):
            presets = web_data[lib_name]
            if not isinstance(presets, dict):
                continue
            lib_var = tk.BooleanVar(value=True)
            self.lib_vars[lib_name] = lib_var
            lib_cb = tk.Checkbutton(self.scrollable_frame, text=f"[{lib_name}]",
                                    variable=lib_var, bg=DIALOG_BG,
                                    font=("Helvetica", 10, "bold"),
                                    command=lambda ln=lib_name: self.toggle_library(ln))
            lib_cb.pack(anchor='w', pady=(5, 0))
            for preset_name in sorted(presets.keys()):
                var = tk.BooleanVar(value=True)
                full_name = f"{lib_name}::{preset_name}"
                cb = tk.Checkbutton(self.scrollable_frame, text=f"  {preset_name}",
                                    variable=var, bg=DIALOG_BG)
                cb.pack(anchor='w')
                self.vars[full_name] = var

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        bind_mousewheel(self, self.canvas)

        import_button = tk.Button(self, text=_("Import Selected"),
                                  command=self.import_selected,
                                  font=("Helvetica", 10, "bold"))
        import_button.pack(pady=10)

    def toggle_library(self, lib_name):
        checked = self.lib_vars[lib_name].get()
        for full_name, var in self.vars.items():
            if full_name.startswith(f"{lib_name}::"):
                var.set(checked)

    def select_all(self):
        for var in self.vars.values():
            var.set(True)
        for var in self.lib_vars.values():
            var.set(True)

    def select_none(self):
        for var in self.vars.values():
            var.set(False)
        for var in self.lib_vars.values():
            var.set(False)

    def import_selected(self):
        added_count = 0
        libs_touched = set()

        for full_name, var in self.vars.items():
            if not var.get():
                continue
            lib_name, preset_name = full_name.split("::", 1)
            preset_data = self.web_data[lib_name][preset_name]

            if lib_name not in self.local_presets:
                self.local_presets[lib_name] = {}
            self.local_presets[lib_name][preset_name] = preset_data
            added_count += 1
            libs_touched.add(lib_name)

        if added_count > 0:
            if save_presets(self.local_presets, self.file_path):
                self.parent_app.update_pad_library_dropdown()
                messagebox.showinfo(_("Import Complete"),
                    _("Imported {added} pad sets into {libs} library/libraries.").format(added=added_count, libs=len(libs_touched)))
            else:
                messagebox.showerror(_("Import Error"), _("Could not save presets to file."))
        else:
            messagebox.showinfo(_("Import Complete"), _("No pad sets were imported."))

        self.destroy()


class ImportTargetWindow(tk.Toplevel):
    def __init__(self, parent, existing_libraries):
        super().__init__(parent)
        self.parent = parent
        self.existing_libraries = existing_libraries
        self.target_library = None

        self.title(_("Select Import Library"))
        self.geometry("350x150")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.mode = tk.StringVar(value="existing")
        
        tk.Label(self, text=_("Where do you want to add these sets?"), bg=DIALOG_BG).pack(pady=10)

        existing_frame = tk.Frame(self, bg=DIALOG_BG)
        existing_frame.pack(fill='x', padx=10)
        tk.Radiobutton(existing_frame, text=_("Add to existing library:"), variable=self.mode, value="existing", bg=DIALOG_BG, command=self.toggle_widgets).pack(side="left")
        self.library_dropdown = ttk.Combobox(existing_frame, values=self.existing_libraries, state="readonly", width=15)
        self.library_dropdown.pack(side="left", padx=5)
        if self.existing_libraries:
            self.library_dropdown.set(self.existing_libraries[0])

        new_frame = tk.Frame(self, bg=DIALOG_BG)
        new_frame.pack(fill='x', padx=10, pady=5)
        tk.Radiobutton(new_frame, text=_("Create new library:"), variable=self.mode, value="new", bg=DIALOG_BG, command=self.toggle_widgets).pack(side="left")
        self.new_lib_entry = tk.Entry(new_frame, width=18)
        self.new_lib_entry.pack(side="left", padx=5)
        
        button_frame = tk.Frame(self, bg=DIALOG_BG)
        button_frame.pack(pady=15)
        tk.Button(button_frame, text=_("Import"), command=self.on_import).pack(side="left", padx=10)
        tk.Button(button_frame, text=_("Cancel"), command=self.on_cancel).pack(side="left", padx=10)
        
        self.toggle_widgets()
        self.wait_window(self)

    def toggle_widgets(self):
        if self.mode.get() == "existing":
            self.library_dropdown.config(state="readonly")
            self.new_lib_entry.config(state="disabled")
        else: # "new"
            self.library_dropdown.config(state="disabled")
            self.new_lib_entry.config(state="normal")
            
    def on_import(self):
        if self.mode.get() == "existing":
            self.target_library = self.library_dropdown.get()
            if not self.target_library:
                messagebox.showwarning(_("No Library"), _("Please select a library."), parent=self)
                return
        else: # "new"
            self.target_library = self.new_lib_entry.get().strip()
            if not self.target_library:
                messagebox.showwarning(_("No Name"), _("Please enter a name for the new library."), parent=self)
                return
        
        self.destroy()
        
    def on_cancel(self):
        self.target_library = None
        self.destroy()

    def get_target_library(self):
        return self.target_library


class PolygonDrawWindow(tk.Toplevel):
    """
    A dialog for drawing a polygon shape on a grid.
    Used for defining irregular leather skin shapes.
    """
    # 32 is comfortable for the canvas + handles camera-captured contours
    # well; previous 8-cap was a leftover from when nesting was slow and
    # high-vertex polygons made the per-pad point-in-polygon checks crawl.
    # Vectorized nesting (v2.40) handles 30+ vertices easily.
    MAX_POINTS = 32
    CANVAS_PX = 450  # Canvas size in pixels
    POINT_RADIUS = 6  # Radius of drawn points in pixels
    CLOSE_THRESHOLD = 15  # Pixels - how close to first point to auto-close
    OVERLAY_REFRESH_MS = 200  # 5 fps for the live camera underlay

    def __init__(self, parent, unit="in", settings=None):
        super().__init__(parent)
        self.unit = unit
        self.polygon_closed = False
        self.result = None  # Will hold the final polygon or None if cancelled

        # Grid size — pick the larger of the user's setting and what's
        # needed to cover the laser bed at 1:1 scale. Absolute-position
        # polygons (camera capture + hand-trace with overlay) extend
        # across the bed, so a too-small grid would clip them. The bed-
        # derived floor auto-upgrades users whose setting predates this
        # change without forcing a migration write.
        self._settings = settings or {}
        if self.unit == "in":
            try:
                self.grid_size = int(
                    self._settings.get("polygon_draw_grid_size_in", 17))
            except (TypeError, ValueError):
                self.grid_size = 17
            mm_per_unit_for_grid = 25.4
        else:
            try:
                self.grid_size = int(
                    self._settings.get("polygon_draw_grid_size_cm", 43))
            except (TypeError, ValueError):
                self.grid_size = 43
            mm_per_unit_for_grid = 10.0
        try:
            bed_max_mm = max(
                float(self._settings.get("laser_bed_x_max", 400)),
                float(self._settings.get("laser_bed_y_max", 415)))
        except (TypeError, ValueError):
            bed_max_mm = 415.0
        self.grid_size = max(
            self.grid_size,
            int(math.ceil(bed_max_mm / mm_per_unit_for_grid)))
        self.grid_size = max(2, self.grid_size)  # sanity floor

        self.points = []  # List of (x, y) in grid units, 0..grid_size each
        # Set when the user captures from the camera. Anchors the grid's
        # local (0, 0) to an absolute machine-mm point so the polygon
        # can be reconstructed in machine coords (for nesting that knows
        # the bed position) AND so a live camera overlay can render at
        # the same machine→grid scale the polygon was captured at.
        # Preserved across vertex edits — the grid frame doesn't move
        # when you tweak a polygon point. Cleared only on Clear or
        # Cancel.
        self._camera_anchor_mm = None
        # Sticky "this polygon has a meaningful machine-coord anchor"
        # flag. Set the first time the camera overlay turns on (overlay
        # always anchors at machine (0, 0), so any clicks the user
        # makes correspond to absolute machine positions in grid units).
        # Once True, never cleared mid-session — even if the user
        # disables the overlay, the polygon they've already traced
        # still carries machine meaning. Used by Frame & Cut to enable
        # the "Try Auto Locate" button (drives the head to the
        # polygon's LB vertex). Plain hand-drawn polygons (never enabled
        # overlay, never captured) keep this False — the button hides.
        self._overlay_used = False

        # Camera-overlay state (populated when the user toggles "Show
        # live camera"). None until enabled.
        self._show_camera_var = None     # BooleanVar — created in _create_widgets
        self._show_camera_chk = None     # Checkbutton widget reference
        self._cap = None
        self._cam_mod = None
        self._calibration = None
        self._PIL_Image = None
        self._PIL_ImageTk = None
        self._overlay_after_id = None
        self._overlay_photo = None        # PhotoImage ref (prevent GC)

        # Title mirrors what the dialog can actually do: shows
        # "Draw / Capture Shape" when camera capture is available,
        # plain "Draw Shape" when only the draw path is wired up.
        # _camera_capture_available() checks both the experimental
        # toggle and calibration presence.
        self.title(_("Draw / Capture Shape")
                    if self._camera_capture_available()
                    else _("Draw Shape"))
        self.geometry("520x680")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        # Calculate pixels per grid unit (1 inch or 1 cm per square)
        self.px_per_unit = self.CANVAS_PX / self.grid_size

        self._create_widgets()
        self._draw_grid()

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.wait_window(self)

    def _create_widgets(self):
        # Instructions
        unit_label = "inches" if self.unit == "in" else "cm"
        instr_text = (
            f"Click anywhere to add a point (max {self.MAX_POINTS} points).\n"
            f"Click near the first point to close. Click a point to remove it."
        )
        tk.Label(self, text=instr_text, bg=DIALOG_BG, justify="center").pack(pady=(10, 5))

        # Grid info - each square = 1 unit (inch or cm)
        tk.Label(self, text=_("Grid: {size}x{size} {unit_label} (1 square = 1 {unit})").format(size=self.grid_size, unit_label=unit_label, unit=self.unit),
                 bg=DIALOG_BG, font=("Helvetica", 9)).pack(pady=(0, 5))

        # Canvas frame
        canvas_frame = tk.Frame(self, bg=DIALOG_BG)
        canvas_frame.pack(padx=10, pady=5)

        self.canvas = tk.Canvas(canvas_frame, width=self.CANVAS_PX, height=self.CANVAS_PX,
                                 bg="white", highlightthickness=1, highlightbackground="gray")
        self.canvas.pack()
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        # Status label
        self.status_var = tk.StringVar(value="Click to add points...")
        tk.Label(self, textvariable=self.status_var, bg=DIALOG_BG, font=("Helvetica", 10)).pack(pady=5)

        # Live-camera overlay toggle. Available as soon as a camera
        # calibration is on disk — the user can either capture a
        # polygon first and then refine its vertices against the live
        # image, OR turn the overlay on with no capture and trace the
        # scrap shape by hand. When there's no capture, the overlay
        # uses the camera's image-bbox machine coords as a fallback
        # anchor so the trace lands at correct grid scale.
        overlay_row = tk.Frame(self, bg=DIALOG_BG)
        overlay_row.pack(pady=(0, 5))
        if self._camera_capture_available():
            self._show_camera_var = tk.BooleanVar(value=False)
            self._show_camera_chk = tk.Checkbutton(
                overlay_row,
                text=_("Show live camera underneath"),
                variable=self._show_camera_var, bg=DIALOG_BG,
                font=("Helvetica", 9),
                command=self._on_camera_overlay_toggle)
            self._show_camera_chk.pack(side="left")

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text=_("Clear"), command=self.on_clear, width=10).pack(side="left", padx=5)
        # "Get from camera" — only shown when OpenCV is available AND a
        # camera calibration file exists on disk.
        if self._camera_capture_available():
            tk.Button(btn_frame, text=_("Get from camera"),
                      command=self.on_get_from_camera,
                      width=15).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Submit"), command=self.on_submit, width=10,
                  font=("Helvetica", 10, "bold")).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Cancel"), command=self.on_cancel, width=10).pack(side="left", padx=5)

    def _camera_capture_available(self):
        """Return True if camera capture is wired up AND calibrated AND
        the user has opted into machine integration AND the platform
        supports it. Gating mirrors the Pad Maker tab: 'Get from camera'
        hides entirely when the experimental machine toggle is off,
        even if a stale calibration file is on disk, and the whole
        machine integration is Windows-only (see PadSVGGeneratorApp.
        _machine_enabled)."""
        import sys
        if sys.platform != 'win32':
            return False
        if not self._settings or not self._settings.get(
                "experimental_machine_menu", False):
            return False
        try:
            import camera_capture
        except ImportError:
            return False
        if not camera_capture.HAS_OPENCV:
            return False
        return camera_capture.load_calibration(
            camera_capture.default_calibration_path()) is not None

    def on_get_from_camera(self):
        """Open CameraCaptureDialog and adopt the returned polygon."""
        try:
            import camera_capture
        except ImportError:
            return
        cam_idx = self._resolve_camera_index()
        if cam_idx is None:
            messagebox.showerror(_("No Camera"),
                                  _("No cameras detected."), parent=self)
            return

        dlg = CameraCaptureDialog(
            self, camera_index=cam_idx,
            calibration_path=camera_capture.default_calibration_path(),
            settings=self._settings)
        if not dlg.result_polygon_mm:
            return

        # CameraCaptureDialog returns the polygon in ABSOLUTE machine-mm
        # (Y-up). Store the polygon at absolute machine grid coords so
        # it appears in the grid view at the same place it sits on the
        # bed (matching what the user saw on the camera screen, and
        # matching where a hand-traced polygon with the live overlay
        # would land). The overlay anchors at machine (0, 0) — grid
        # origin = bed origin — so machine-mm divided by mm_per_unit
        # lands directly on the grid.
        machine_polygon = list(dlg.result_polygon_mm)
        self._camera_anchor_mm = (0.0, 0.0)
        scale_to_unit = 1.0 / 25.4 if self.unit == "in" else 1.0 / 10.0
        captured = [(x * scale_to_unit, y * scale_to_unit)
                     for (x, y) in machine_polygon]
        # Downsample to MAX_POINTS if the contour is denser. Picks every
        # k-th vertex (uniform sampling along the contour).
        if len(captured) > self.MAX_POINTS:
            step = len(captured) / self.MAX_POINTS
            captured = [captured[int(i * step)] for i in range(self.MAX_POINTS)]
        self.points = [(max(0, min(self.grid_size, x)),
                        max(0, min(self.grid_size, y)))
                        for (x, y) in captured]
        self.polygon_closed = True
        self._redraw_polygon()
        # Checkbox is already enabled (overlay works pre- or post-
        # capture); no state change needed here.
        self.status_var.set(
            _("Captured {n} points from camera").format(n=len(self.points)))

    def _resolve_camera_index(self):
        """Same resolver pattern as the dialogs in main.py: persisted
        override first, then Falcon-name heuristic, then last enumerated."""
        try:
            import camera_capture
        except ImportError:
            return None
        if not camera_capture.HAS_OPENCV:
            return None
        if self._settings:
            override = self._settings.get("camera_index_override")
            if override is not None:
                try:
                    return int(override)
                except (TypeError, ValueError):
                    pass
        idx = camera_capture.find_falcon_camera_index()
        if idx is not None:
            return idx
        cams = camera_capture.enumerate_cameras()
        if cams:
            return cams[-1]['index']
        return None

    # ------------------------------------------------------------------
    # Live camera overlay
    # ------------------------------------------------------------------

    def _on_camera_overlay_toggle(self):
        """Checkbox handler — start/stop the live underlay."""
        if self._show_camera_var.get():
            self._start_camera_overlay()
        else:
            self._stop_camera_overlay()

    def _start_camera_overlay(self):
        # No anchor required up front — when there isn't a capture
        # anchor yet, _render_overlay falls back to the camera image's
        # own machine-coord bounding box as an anchor (so the trace
        # lands at correct grid scale even pre-capture).
        if self._cap is not None:
            return  # already running
        # Mark the polygon as machine-coord-anchored. The overlay
        # always renders at machine (0, 0), so any clicks the user
        # makes correspond to absolute machine positions. Setting
        # _camera_anchor_mm here (in addition to _overlay_used)
        # routes the polygon through main.py's _adopt_camera_polygon
        # path on submit — same path camera-captured polygons take —
        # so the camera-polygon safety inset is applied to traced
        # polygons too. A skip-the-anchor implementation would mean
        # only camera captures got the inset; hand-traced shapes
        # would cut right up to the user's click points, which is
        # less safe for irregular leather edges.
        self._overlay_used = True
        if self._camera_anchor_mm is None:
            self._camera_anchor_mm = (0.0, 0.0)
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
            self._calibration = camera_capture.load_calibration(
                camera_capture.default_calibration_path())
            if not self._calibration:
                raise RuntimeError(_("No camera calibration on disk."))
        except (ImportError, RuntimeError) as e:
            messagebox.showerror(_("Camera Overlay Unavailable"),
                                  str(e), parent=self)
            self._show_camera_var.set(False)
            return
        cam_idx = self._resolve_camera_index()
        if cam_idx is None:
            messagebox.showerror(_("No Camera"),
                                  _("No cameras detected."), parent=self)
            self._show_camera_var.set(False)
            return
        self._camera_index = cam_idx
        self.status_var.set(_("Opening camera for overlay..."))
        import threading

        def _worker():
            try:
                cap = camera_capture.open_camera(cam_idx)
                ok, _frame = cap.read()
                if not ok:
                    cap.release()
                    raise RuntimeError(_("Camera no frame"))
                self.after(0, self._camera_overlay_ready, cap)
            except Exception as e:
                self.after(0, self._camera_overlay_failed, e)

        threading.Thread(target=_worker, name='polydraw-overlay',
                          daemon=True).start()

    def _camera_overlay_ready(self, cap):
        # User may have toggled off / cancelled in the meantime.
        if not self._show_camera_var or not self._show_camera_var.get():
            try:
                cap.release()
            except Exception:
                pass
            return
        self._cap = cap
        self.status_var.set(_("Live camera underlay on."))
        self._overlay_refresh_loop()

    def _camera_overlay_failed(self, err):
        self.status_var.set(_("Camera open failed: {e}").format(e=err))
        if self._show_camera_var is not None:
            self._show_camera_var.set(False)

    def _overlay_refresh_loop(self):
        if (self._cap is None
                or not self.winfo_exists()
                or self._show_camera_var is None
                or not self._show_camera_var.get()):
            return
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok = False
        if ok and frame is not None:
            try:
                self._render_overlay(frame)
            except Exception:
                pass  # don't kill the loop on a bad frame
        self._overlay_after_id = self.after(
            self.OVERLAY_REFRESH_MS, self._overlay_refresh_loop)

    def _render_overlay(self, frame):
        """Warp the camera frame into canvas-pixel space using the same
        pixel→machine homography the capture used, plus the grid's
        machine anchor + canvas px-per-mm scale. Result lands UNDER
        the grid + polygon items via tag_lower."""
        import cv2
        import numpy as np
        undist = self._cam_mod.undistort_frame(frame, self._calibration)
        h, w = undist.shape[:2]
        # Map four image corners through pixels_to_mm (which handles
        # schema-1 Y-flip compat internally), then through our
        # machine→canvas-pixel affine, to get the destination corners
        # for the perspective warp.
        src_corners = [(0.0, 0.0), (float(w), 0.0),
                       (float(w), float(h)), (0.0, float(h))]
        machine_corners = self._cam_mod.pixels_to_mm(
            src_corners, self._calibration)
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        scale = self.px_per_unit / mm_per_unit  # canvas-px per machine-mm
        # Anchor priority: explicit capture anchor wins; otherwise
        # fall back to machine (0, 0) — the bed front-left corner.
        # Canvas BL then shows the bed origin, matching the captured-
        # mode convention of "polygon BL at canvas BL" so the camera
        # view is positioned consistently regardless of whether a
        # polygon was captured first.
        #
        # An earlier fallback anchored at the camera image's own bbox
        # bottom-left, which on Matt's calibrated rig (camera sees
        # past the bed edges to ~-133mm in machine X) put the canvas
        # BL way off the bed and made the visible bed content shift
        # into the right half of the canvas — read as "scale off"
        # since the bed appeared smaller than expected.
        if self._camera_anchor_mm is not None:
            ax, ay = self._camera_anchor_mm
        else:
            ax, ay = 0.0, 0.0
        dst_corners = []
        for mx, my in machine_corners:
            cx = (mx - ax) * scale
            # Canvas Y is image-down; grid Y is up — invert here.
            cy = self.CANVAS_PX - (my - ay) * scale
            dst_corners.append((cx, cy))
        src_arr = np.array(src_corners, dtype=np.float32)
        dst_arr = np.array(dst_corners, dtype=np.float32)
        try:
            homography = cv2.getPerspectiveTransform(src_arr, dst_arr)
        except cv2.error:
            return
        warped = cv2.warpPerspective(
            undist, homography,
            (self.CANVAS_PX, self.CANVAS_PX),
            borderMode=cv2.BORDER_CONSTANT,
            borderValue=(255, 255, 255))
        rgb = cv2.cvtColor(warped, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        self._overlay_photo = self._PIL_ImageTk.PhotoImage(img)
        try:
            self.canvas.delete('camera_overlay')
            self.canvas.create_image(
                0, 0, image=self._overlay_photo, anchor='nw',
                tags='camera_overlay')
            # Send to the bottom so grid + polygon items stay visible.
            self.canvas.tag_lower('camera_overlay')
        except tk.TclError:
            pass

    def _stop_camera_overlay(self):
        if self._overlay_after_id is not None:
            try:
                self.after_cancel(self._overlay_after_id)
            except Exception:
                pass
            self._overlay_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        try:
            self.canvas.delete('camera_overlay')
        except tk.TclError:
            pass
        self._overlay_photo = None

    def _draw_grid(self):
        """Draw the grid lines on the canvas."""
        self.canvas.delete("grid")

        # Draw grid lines (1 line per unit)
        # Highlight every 5th line for easier orientation (especially useful for cm grid)
        for i in range(self.grid_size + 1):
            px = i * self.px_per_unit
            # Use reddish color for every 5th line (0, 5, 10, 15, ...)
            if i % 5 == 0:
                line_color = "#CC9999"
            else:
                line_color = "#CCCCCC"
            # Vertical lines
            self.canvas.create_line(px, 0, px, self.CANVAS_PX, fill=line_color, tags="grid")
            # Horizontal lines
            self.canvas.create_line(0, px, self.CANVAS_PX, px, fill=line_color, tags="grid")

        # Draw axis labels - every 5 for inches, every 10 for cm (to avoid crowding)
        label_step = 5 if self.unit == "in" else 10
        for i in range(0, self.grid_size + 1, label_step):
            px = i * self.px_per_unit
            # X-axis labels (bottom) - label value = grid coordinate (1 square = 1 unit)
            self.canvas.create_text(px, self.CANVAS_PX - 5, text=str(i),
                                     font=("Helvetica", 8), anchor="s", tags="grid")
            # Y-axis labels (left) - invert Y so 0 is at bottom
            self.canvas.create_text(5, self.CANVAS_PX - px, text=str(i),
                                     font=("Helvetica", 8), anchor="w", tags="grid")

    def _grid_to_canvas(self, gx, gy):
        """Convert grid coordinates to canvas pixels. Y is inverted (0 at bottom)."""
        cx = gx * self.px_per_unit
        cy = self.CANVAS_PX - (gy * self.px_per_unit)
        return cx, cy

    def _canvas_to_grid(self, cx, cy):
        """Convert canvas pixels to grid coordinates.

        Returns float grid units (no longer snapped to integer
        intersections — a 1-inch grid spacing is too coarse for
        accurate scrap tracing). Camera-captured polygons also produce
        floats, so downstream nesting + machine-coord round-trip
        already handle non-integer points.

        Bounds-clamped to [0, grid_size].
        """
        gx = cx / self.px_per_unit
        gy = (self.CANVAS_PX - cy) / self.px_per_unit
        gx = max(0.0, min(float(self.grid_size), gx))
        gy = max(0.0, min(float(self.grid_size), gy))
        return gx, gy

    def _redraw_polygon(self):
        """Redraw all points and lines."""
        self.canvas.delete("polygon")

        if not self.points:
            return

        # Draw lines between points
        if len(self.points) > 1:
            canvas_points = [self._grid_to_canvas(p[0], p[1]) for p in self.points]
            for i in range(len(canvas_points) - 1):
                x1, y1 = canvas_points[i]
                x2, y2 = canvas_points[i + 1]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

            # If closed, draw line from last to first
            if self.polygon_closed:
                x1, y1 = canvas_points[-1]
                x2, y2 = canvas_points[0]
                self.canvas.create_line(x1, y1, x2, y2, fill="blue", width=2, tags="polygon")

        # Draw points
        for i, (gx, gy) in enumerate(self.points):
            cx, cy = self._grid_to_canvas(gx, gy)
            color = "green" if i == 0 else "red"
            self.canvas.create_oval(
                cx - self.POINT_RADIUS, cy - self.POINT_RADIUS,
                cx + self.POINT_RADIUS, cy + self.POINT_RADIUS,
                fill=color, outline="black", tags="polygon"
            )

    def _update_status(self):
        """Update the status label."""
        if self.polygon_closed:
            self.status_var.set(f"Shape closed ({len(self.points)} points). Click Submit or adjust points.")
        elif len(self.points) >= self.MAX_POINTS:
            self.status_var.set(f"Max points reached ({self.MAX_POINTS}). Click near first point to close.")
        else:
            remaining = self.MAX_POINTS - len(self.points)
            self.status_var.set(f"{len(self.points)} points. {remaining} remaining. Click near first point to close.")

    def on_canvas_click(self, event):
        """Handle click on canvas. The grid's machine anchor (if set
        via camera capture) is preserved across vertex edits — moving
        a point in grid coords still has a valid machine-coord
        translation, so the polygon round-trips to machine on submit."""
        if self.polygon_closed:
            # If already closed, check if clicking on a point to remove it
            clicked_idx = self._get_clicked_point_index(event.x, event.y)
            if clicked_idx is not None:
                self.points.pop(clicked_idx)
                self.polygon_closed = False
                self._redraw_polygon()
                self._update_status()
            return

        # Check if clicking near first point to close the polygon (before point removal)
        if len(self.points) >= 3:
            first_cx, first_cy = self._grid_to_canvas(self.points[0][0], self.points[0][1])
            dist = ((event.x - first_cx) ** 2 + (event.y - first_cy) ** 2) ** 0.5
            if dist < self.CLOSE_THRESHOLD:
                self.polygon_closed = True
                self._redraw_polygon()
                self._update_status()
                return

        # Check if clicking on an existing point (to remove it)
        clicked_idx = self._get_clicked_point_index(event.x, event.y)
        if clicked_idx is not None:
            self.points.pop(clicked_idx)
            self._redraw_polygon()
            self._update_status()
            return

        # Get grid coordinates
        gx, gy = self._canvas_to_grid(event.x, event.y)

        # Check if we can add more points
        if len(self.points) >= self.MAX_POINTS:
            return

        # Check if point already exists at this location
        if (gx, gy) in self.points:
            return

        # Add the point
        self.points.append((gx, gy))
        self._redraw_polygon()
        self._update_status()

    def _get_clicked_point_index(self, cx, cy):
        """Return index of point near click, or None."""
        for i, (gx, gy) in enumerate(self.points):
            pcx, pcy = self._grid_to_canvas(gx, gy)
            dist = ((cx - pcx) ** 2 + (cy - pcy) ** 2) ** 0.5
            if dist <= self.POINT_RADIUS + 4:  # Small tolerance
                return i
        return None

    def on_clear(self):
        """Clear all points AND (if overlay is OFF) the machine
        anchor. If overlay is currently running, preserve the anchor
        so a retrace still has machine reference — the user is
        clearing the polygon but the same camera context still
        applies to the next set of clicks."""
        self.points = []
        self.polygon_closed = False
        if self._cap is None:
            # Overlay isn't running — no machine context applies to
            # the next polygon. Clear both signals.
            self._camera_anchor_mm = None
            self._overlay_used = False
        # Else: overlay is still showing camera content, leave
        # _camera_anchor_mm and _overlay_used as they are.
        self._redraw_polygon()
        self._update_status()

    def get_machine_polygon_mm(self):
        """Return absolute machine-Y-up polygon, reconstructed from the
        current grid points + the camera-set anchor. None if no anchor
        (purely drawn polygon — no machine reference)."""
        if self._camera_anchor_mm is None or not self.points:
            return None
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        ax, ay = self._camera_anchor_mm
        return [(gx * mm_per_unit + ax, gy * mm_per_unit + ay)
                for (gx, gy) in self.points]

    def get_polygon_lb_machine_mm(self):
        """Return the polygon's leftmost-lowest vertex in absolute
        machine coords (Y-up), or None if no machine reference exists.

        Used by Frame & Cut's "Try Auto Locate" — drives the laser
        head to this exact point so the user doesn't have to jog
        manually. Available whenever the polygon was either captured
        from the camera (anchor set) or hand-traced with the live
        overlay on (overlay_used set; the overlay anchors at machine
        (0, 0) so grid coords map 1:1 to machine coords). For purely
        hand-drawn polygons (no overlay, no capture) the grid coords
        are arbitrary — no machine reference — so None is returned
        and the auto-locate button hides.
        """
        if not self.points:
            return None
        if self._camera_anchor_mm is None and not self._overlay_used:
            return None
        mm_per_unit = 25.4 if self.unit == "in" else 10.0
        ax, ay = self._camera_anchor_mm or (0.0, 0.0)
        machine = [(gx * mm_per_unit + ax, gy * mm_per_unit + ay)
                    for (gx, gy) in self.points]
        return min(machine, key=lambda p: (p[0], p[1]))

    def on_submit(self):
        """Submit the polygon."""
        if len(self.points) < 3:
            from tkinter import messagebox
            messagebox.showwarning(_("Not Enough Points"),
                                   _("Please draw at least 3 points to create a shape."),
                                   parent=self)
            return

        if not self.polygon_closed:
            from tkinter import messagebox
            messagebox.showwarning(_("Shape Not Closed"),
                                   _("Please close the shape by clicking near the first point."),
                                   parent=self)
            return

        # Release camera + cancel refresh timer before destroying the
        # dialog. get_machine_polygon_mm() is reconstructed on demand
        # from self.points + self._camera_anchor_mm, both of which
        # survive past destroy(), so the caller can still read the
        # absolute polygon after we close.
        self._stop_camera_overlay()
        # Return points as list of (x, y) tuples in grid units
        self.result = list(self.points)
        self.destroy()

    def on_cancel(self):
        """Cancel and close."""
        # Stop the camera overlay first — releases the cv2 cap handle
        # and cancels the refresh timer. Calling destroy() without
        # this would leak the camera.
        self._stop_camera_overlay()
        self.result = None
        # Drop the anchor so get_machine_polygon_mm doesn't return a
        # stale polygon after the user explicitly cancelled.
        self._camera_anchor_mm = None
        self.destroy()

    def get_polygon(self):
        """Return the polygon points or None if cancelled."""
        return self.result


# ==========================================
# G-CODE SETTINGS DIALOG
# ==========================================

class GcodeSettingsWindow:
    """Dialog for configuring G-code laser settings per material."""

    MATERIALS = [
        ("felt", "Felt"),
        ("card", "Card"),
        ("leather", "Leather"),
        ("acrylic", "Acrylic"),
    ]

    # Non-engraving operations (engraving is handled separately with mode toggle)
    OPERATIONS = [
        ("hole", "Center Hole"),
        ("cut", "Outer Cut"),
    ]

    def __init__(self, parent, settings, save_callback, materials=None,
                 show_tooling_engraving=False,
                 gcode_presets=None, gcode_presets_save_callback=None):
        self.settings = settings
        self.save_callback = save_callback
        # Allow filtering which materials to show
        self.active_materials = materials if materials else self.MATERIALS
        self.show_tooling_engraving = show_tooling_engraving

        # Per-material preset library: {material: {preset_name: data}}.
        # Mutated in place; persisted via gcode_presets_save_callback. None
        # disables the preset UI entirely (for callers that opt out / tests).
        self.gcode_presets = gcode_presets
        self.gcode_presets_save_callback = gcode_presets_save_callback or (lambda: None)
        # Per-material: name of the preset currently loaded (or None),
        # tk widget references for the preset bar, and the baseline snapshot
        # used by dirty-tracking. Populated as material sections are built.
        self.active_preset_name = {}
        self.preset_combos = {}
        self.material_baseline = {}

        self.top = tk.Toplevel(parent)
        title = "Tooling Settings" if show_tooling_engraving else "G-code Laser Settings"
        self.top.title(title)
        self.top.geometry("640x500")
        self.top.configure(bg=DIALOG_BG)
        self.top.transient(parent)
        self.top.grab_set()

        # Get current gcode settings or defaults
        self.gcode_settings = settings.get("gcode_settings", {})

        # Create variable storage
        self.vars = {}  # vars[material][operation]['speed'|'power']

        self._create_widgets()
        # Baselines snapshot the form right after construction, so the
        # initial state of each material is treated as "clean." Also
        # detect whether the current form values match a saved preset
        # so the dropdown shows what's loaded instead of looking empty.
        if self.gcode_presets is not None:
            for mat_key, _label in self.active_materials:
                self.material_baseline[mat_key] = self._capture_material_to_dict(mat_key)
                match = self._detect_active_preset(mat_key)
                if match is not None:
                    self.active_preset_name[mat_key] = match
                    self._refresh_gcode_preset_combo(mat_key, select=match)

    def _create_widgets(self):
        # Header
        header_frame = tk.Frame(self.top, bg=DIALOG_BG)
        header_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(header_frame, text=_("Configure laser speed and power for each material and operation."),
                 bg=DIALOG_BG, wraplength=500, justify="left").pack(anchor="w")

        tk.Label(header_frame, text=_("Order: Engraving → Center Hole → Outer Cut"),
                 bg=DIALOG_BG, font=("Helvetica", 9, "italic")).pack(anchor="w", pady=(5, 0))

        tk.Label(header_frame, text=_("Note: Power uses Grbl's S0-S1000 scale. If power seems wrong, check that "
                 "your machine's $30 setting is 1000 (run \"$30=1000\" in your console)."),
                 bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666", wraplength=420, justify="left").pack(anchor="w", pady=(5, 0))

        # Main content with scrollable frame
        main_canvas_frame = tk.Frame(self.top)
        main_canvas_frame.pack(fill="both", expand=True, padx=10)

        canvas = tk.Canvas(main_canvas_frame, bg=DIALOG_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(main_canvas_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=DIALOG_BG)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        bind_mousewheel(self.top, canvas)

        # Create material sections
        for mat_key, mat_label in self.active_materials:
            self._create_material_section(scrollable_frame, mat_key, mat_label)

        # Tooling engraving settings (shown only for tooling dialog)
        if self.show_tooling_engraving:
            self._create_tooling_engraving_section(scrollable_frame)

        # "Filled" engraving overscan option
        overscan_frame = tk.Frame(scrollable_frame, bg=DIALOG_BG)
        overscan_frame.pack(fill="x", padx=5, pady=(5, 0))

        self.overscan_var = tk.BooleanVar(value=self.settings.get("filled_overscan_enabled", False))
        overscan_cb = tk.Checkbutton(overscan_frame,
                                     text=_('"Filled" engraving overscan optimization'),
                                     variable=self.overscan_var, bg=DIALOG_BG)
        overscan_cb.pack(anchor="w", padx=5)
        overscan_caption = tk.Label(overscan_frame,
                                    text=_("(extends scan lines so laser is at full speed at character edges)"),
                                    bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666")
        overscan_caption.pack(anchor="w", padx=(28, 5))
        add_tooltips(
            _("Extends each filled scan line a few millimeters past the "
            "character edges so the laser head reaches full speed before "
            "burning. Produces cleaner edges on filled engraving at the "
            "cost of slightly more travel time."),
            overscan_cb, overscan_caption,
        )

        # --- Global G-code Settings ---
        global_frame = tk.LabelFrame(scrollable_frame, text=_("Global Settings"), bg=DIALOG_BG, padx=10, pady=10)
        global_frame.pack(fill="x", padx=5, pady=(10, 0))

        # Return-to-home speed
        return_frame = tk.Frame(global_frame, bg=DIALOG_BG)
        return_frame.pack(fill="x")
        ret_lbl = tk.Label(return_frame, text=_("Return-to-home speed:"), bg=DIALOG_BG)
        ret_lbl.pack(side="left")
        self.return_speed_var = tk.IntVar(value=int(self.settings.get("gcode_return_speed", 1000)))
        ret_ent = tk.Entry(return_frame, textvariable=self.return_speed_var, width=8)
        ret_ent.pack(side="left", padx=5)
        ret_unit = tk.Label(return_frame, text=_("mm/min"), bg=DIALOG_BG)
        ret_unit.pack(side="left")
        add_tooltips(
            _("How fast the laser head returns home after a job. Slower "
            "settings can avoid endstop crashes on machines with cheap "
            "limit switches; faster gets you to the next job sooner."),
            ret_lbl, ret_ent, ret_unit,
        )

        # Cut grouping
        grouping_frame = tk.Frame(global_frame, bg=DIALOG_BG)
        grouping_frame.pack(fill="x", pady=(8, 0))
        tk.Label(grouping_frame, text=_("Cut grouping:"), bg=DIALOG_BG).pack(anchor="w")
        self.grouping_var = tk.StringVar(value=self.settings.get("gcode_cut_grouping", "layer"))
        grp_layer = tk.Radiobutton(grouping_frame,
                                   text=_("By layer (all engravings, then all holes, then all cuts)"),
                                   variable=self.grouping_var, value="layer", bg=DIALOG_BG)
        grp_layer.pack(anchor="w", padx=(20, 0))
        grp_pad = tk.Radiobutton(grouping_frame,
                                 text=_("By pad (engrave + hole + cut for each pad, then next)"),
                                 variable=self.grouping_var, value="pad", bg=DIALOG_BG)
        grp_pad.pack(anchor="w", padx=(20, 0))
        add_tooltip(grp_layer,
                    _("Run every engraving, then every hole, then every cut "
                    "across the whole sheet. Good when you want to swap "
                    "settings or check progress between layers."))
        add_tooltip(grp_pad,
                    _("Finish each pad completely (engrave, hole, then cut) "
                    "before moving on to the next. Good when you want each "
                    "finished pad to drop free as it completes."))

        # Buttons
        button_frame = tk.Frame(self.top, bg=DIALOG_BG)
        button_frame.pack(fill="x", padx=10, pady=10)

        apply_btn = tk.Button(button_frame, text=_("Apply"), command=self._on_apply_clicked)
        apply_btn.pack(side="left", padx=5)
        cancel_btn = tk.Button(button_frame, text=_("Cancel"), command=self._on_cancel_clicked)
        cancel_btn.pack(side="left", padx=5)
        reset_btn = tk.Button(button_frame, text=_("Reset to Defaults"), command=self._reset_defaults)
        reset_btn.pack(side="right", padx=5)
        add_tooltip(apply_btn, _("Apply the values in this dialog and close."))
        add_tooltip(cancel_btn, _("Close without applying any changes."))
        add_tooltip(reset_btn,
                    _("Reset every value in this dialog back to factory "
                    "defaults (tuned for the Creality Falcon2 Pro 40W)."))

        # Intercept the window-close X so it goes through the same dirty
        # prompt as Cancel.
        self.top.protocol("WM_DELETE_WINDOW", self._on_cancel_clicked)

    def _on_engraving_mode_changed(self, mat_key, mode):
        """Handle engraving mode checkbox toggle - ensure exactly one is checked."""
        if mode == "line":
            self.vars[mat_key]['engraving_mode'].set("line")
        else:
            self.vars[mat_key]['engraving_mode'].set("filled")

    def _create_material_section(self, parent, mat_key, mat_label):
        """Create a settings section for one material."""
        mat_frame = tk.LabelFrame(parent, text=mat_label, bg=DIALOG_BG, padx=10, pady=10)
        mat_frame.pack(fill="x", pady=5, padx=5)

        self.vars[mat_key] = {}

        # Get current settings for this material
        mat_settings = self.gcode_settings.get(mat_key, {})

        # Preset bar (only when the library was passed in). Packed at the
        # top of the LabelFrame; the actual settings grid lives in a sub-
        # frame below so we can mix pack here with grid in `frame`.
        if self.gcode_presets is not None:
            self._create_material_preset_bar(mat_frame, mat_key, mat_label)

        # Sub-frame that owns the grid layout for the material's settings.
        # Named `frame` (not `grid_frame`) so the existing rows below grid
        # into it without renaming dozens of references.
        frame = tk.Frame(mat_frame, bg=DIALOG_BG)
        frame.pack(fill="x")

        # Header row
        op_hdr = tk.Label(frame, text=_("Operation"), bg=DIALOG_BG, font=("Helvetica", 9, "bold"))
        op_hdr.grid(row=0, column=0, sticky="w", padx=5)
        sp_hdr = tk.Label(frame, text=_("Speed (mm/min)"), bg=DIALOG_BG, font=("Helvetica", 9, "bold"))
        sp_hdr.grid(row=0, column=1, padx=5)
        pw_hdr = tk.Label(frame, text=_("Power (%)"), bg=DIALOG_BG, font=("Helvetica", 9, "bold"))
        pw_hdr.grid(row=0, column=2, padx=5)
        ps_hdr = tk.Label(frame, text=_("Passes"), bg=DIALOG_BG, font=("Helvetica", 9, "bold"))
        ps_hdr.grid(row=0, column=3, padx=5)
        air_hdr = tk.Label(frame, text=_("Air"), bg=DIALOG_BG, font=("Helvetica", 9, "bold"))
        air_hdr.grid(row=0, column=4, padx=5)
        add_tooltip(sp_hdr, _("Feed rate in millimeters per minute."))
        add_tooltip(pw_hdr,
                    _("Laser power, 0–100 %. Mapped onto Grbl's S0–S1000 "
                    "scale, so check that your machine's $30 setting is "
                    "1000 (`$30=1000` in your console) for the percentage "
                    "to mean what it says."))
        add_tooltip(ps_hdr,
                    _("How many times to repeat each stroke in this layer. "
                    "Two or more lower-power passes often cut thicker "
                    "leather or acrylic cleaner than one high-power pass."))
        add_tooltip(air_hdr,
                    _("Per-operation air-assist toggle. Sends M8 (on) before "
                    "the operation and M9 (off) after."))

        # --- Engraving mode section (rows 1-2) ---
        current_mode = mat_settings.get("engraving_mode", "line")
        mode_var = tk.StringVar(value=current_mode)
        self.vars[mat_key]['engraving_mode'] = mode_var

        # Line engraving row
        self.vars[mat_key]['engraving'] = {}
        default_eng_speed = self._get_default(mat_key, "engraving_speed")
        default_eng_power = self._get_default(mat_key, "engraving_power")
        default_eng_passes = self._get_default(mat_key, "engraving_passes", 1)
        current_eng_speed = mat_settings.get("engraving_speed", default_eng_speed)
        current_eng_power = mat_settings.get("engraving_power", default_eng_power)
        current_eng_passes = mat_settings.get("engraving_passes", default_eng_passes)

        eng_speed_var = tk.IntVar(value=int(current_eng_speed))
        eng_power_var = tk.DoubleVar(value=current_eng_power)
        eng_passes_var = tk.IntVar(value=int(current_eng_passes))
        self.vars[mat_key]['engraving']['speed'] = eng_speed_var
        self.vars[mat_key]['engraving']['power'] = eng_power_var
        self.vars[mat_key]['engraving']['passes'] = eng_passes_var

        line_rb = tk.Radiobutton(frame, text=_("Engraving (Line)"), bg=DIALOG_BG,
                                 variable=mode_var, value="line",
                                 command=lambda mk=mat_key: self._on_engraving_mode_changed(mk, "line"))
        line_rb.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        line_speed_ent = tk.Entry(frame, textvariable=eng_speed_var, width=10)
        line_speed_ent.grid(row=1, column=1, padx=5, pady=2)
        line_power_ent = tk.Entry(frame, textvariable=eng_power_var, width=10)
        line_power_ent.grid(row=1, column=2, padx=5, pady=2)
        line_passes_ent = tk.Spinbox(frame, textvariable=eng_passes_var, from_=1, to=10, width=5)
        line_passes_ent.grid(row=1, column=3, padx=5, pady=2)
        add_tooltip(line_rb,
                    _("Single-stroke outline engraving — traces each "
                    "character path once. Faster, simpler text on this material."))
        add_tooltip(line_speed_ent, _("Feed rate for line-engraving on this material."))
        add_tooltip(line_power_ent, _("Laser power for line-engraving on this material."))
        add_tooltip(line_passes_ent,
                    _("How many times to repeat the line-engraving strokes. "
                    "Default 1; raise it if a single pass leaves the text "
                    "too faint on this material."))

        air_eng_var = tk.BooleanVar(value=mat_settings.get("air_assist_engraving", True))
        self.vars[mat_key]['air_assist_engraving'] = air_eng_var
        air_eng_cb = tk.Checkbutton(frame, variable=air_eng_var, bg=DIALOG_BG)
        air_eng_cb.grid(row=1, column=4, padx=5, pady=2)
        add_tooltip(air_eng_cb, _("Air-assist on during line-engraving on this material."))

        # "Filled" engraving row
        self.vars[mat_key]['filled_engraving'] = {}
        default_fill_speed = self._get_default(mat_key, "filled_engraving_speed")
        default_fill_power = self._get_default(mat_key, "filled_engraving_power")
        default_fill_passes = self._get_default(mat_key, "filled_engraving_passes", 1)
        current_fill_speed = mat_settings.get("filled_engraving_speed", default_fill_speed)
        current_fill_power = mat_settings.get("filled_engraving_power", default_fill_power)
        current_fill_passes = mat_settings.get("filled_engraving_passes", default_fill_passes)

        fill_speed_var = tk.IntVar(value=int(current_fill_speed))
        fill_power_var = tk.DoubleVar(value=current_fill_power)
        fill_passes_var = tk.IntVar(value=int(current_fill_passes))
        self.vars[mat_key]['filled_engraving']['speed'] = fill_speed_var
        self.vars[mat_key]['filled_engraving']['power'] = fill_power_var
        self.vars[mat_key]['filled_engraving']['passes'] = fill_passes_var

        fill_rb = tk.Radiobutton(frame, text=_('Engraving ("Filled")'), bg=DIALOG_BG,
                                 variable=mode_var, value="filled",
                                 command=lambda mk=mat_key: self._on_engraving_mode_changed(mk, "filled"))
        fill_rb.grid(row=2, column=0, sticky="w", padx=5, pady=2)
        fill_speed_ent = tk.Entry(frame, textvariable=fill_speed_var, width=10)
        fill_speed_ent.grid(row=2, column=1, padx=5, pady=2)
        fill_power_ent = tk.Entry(frame, textvariable=fill_power_var, width=10)
        fill_power_ent.grid(row=2, column=2, padx=5, pady=2)
        fill_passes_ent = tk.Spinbox(frame, textvariable=fill_passes_var, from_=1, to=10, width=5)
        fill_passes_ent.grid(row=2, column=3, padx=5, pady=2)
        add_tooltip(fill_rb,
                    _("Scan-line raster fill of each character — solid "
                    "filled glyphs. Slower than line-engraving but more "
                    "legible on dark / soft materials."))
        add_tooltip(fill_speed_ent, _("Feed rate for filled engraving on this material."))
        add_tooltip(fill_power_ent, _("Laser power for filled engraving on this material."))
        add_tooltip(fill_passes_ent,
                    _("How many times to repeat the filled-engraving scan. "
                    "Default 1; raise it for darker, more solid coverage."))

        air_fill_var = tk.BooleanVar(value=mat_settings.get("air_assist_filled_engraving", True))
        self.vars[mat_key]['air_assist_filled_engraving'] = air_fill_var
        air_fill_cb = tk.Checkbutton(frame, variable=air_fill_var, bg=DIALOG_BG)
        air_fill_cb.grid(row=2, column=4, padx=5, pady=2)
        add_tooltip(air_fill_cb, _("Air-assist on during filled engraving on this material."))

        # Fill density slider (row 3)
        default_spacing = self._get_default(mat_key, "filled_line_spacing")
        current_spacing = mat_settings.get("filled_line_spacing", default_spacing)
        # Map line spacing to density: less=0.3mm, more=0.08mm
        # Slider value 0-100, density = 0.3 - (slider/100) * 0.22
        density_val = int((0.3 - current_spacing) / 0.22 * 100)
        density_val = max(0, min(100, density_val))
        density_var = tk.IntVar(value=density_val)
        self.vars[mat_key]['fill_density'] = density_var

        density_frame = tk.Frame(frame, bg=DIALOG_BG)
        density_frame.grid(row=3, column=0, columnspan=4, sticky="ew", padx=5, pady=(0, 4))
        density_lbl = tk.Label(density_frame, text=_("Fill density:"), bg=DIALOG_BG, font=("Helvetica", 8))
        density_lbl.pack(side="left", padx=(20, 5))
        density_less = tk.Label(density_frame, text=_("less"), bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666")
        density_less.pack(side="left")
        density_scale = tk.Scale(density_frame, from_=0, to=100, orient="horizontal",
                                 variable=density_var, showvalue=False, length=150,
                                 bg=DIALOG_BG, highlightthickness=0)
        density_scale.pack(side="left", padx=2)
        density_more = tk.Label(density_frame, text=_("more"), bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666")
        density_more.pack(side="left")
        add_tooltips(
            _("Spacing between scan lines for filled engraving. Less = wider "
            "lines (faster, slightly visible gaps). More = tighter lines "
            "(slower, fully solid coverage)."),
            density_lbl, density_less, density_scale, density_more,
        )

        # --- Non-engraving operations (rows 4+) ---
        op_help = {
            "hole": ("center mounting hole", "centre-hole cut"),
            "cut": ("outer disc cut", "outer cut"),
        }
        for i, (op_key, op_label) in enumerate(self.OPERATIONS, start=4):
            self.vars[mat_key][op_key] = {}

            default_speed = self._get_default(mat_key, f"{op_key}_speed")
            default_power = self._get_default(mat_key, f"{op_key}_power")
            default_passes = self._get_default(mat_key, f"{op_key}_passes", 1)
            current_speed = mat_settings.get(f"{op_key}_speed", default_speed)
            current_power = mat_settings.get(f"{op_key}_power", default_power)
            current_passes = mat_settings.get(f"{op_key}_passes", default_passes)

            speed_var = tk.IntVar(value=int(current_speed))
            power_var = tk.DoubleVar(value=current_power)
            passes_var = tk.IntVar(value=int(current_passes))

            self.vars[mat_key][op_key]['speed'] = speed_var
            self.vars[mat_key][op_key]['power'] = power_var
            self.vars[mat_key][op_key]['passes'] = passes_var

            descr, short = op_help.get(op_key, (op_label.lower(), op_label.lower()))
            row_lbl = tk.Label(frame, text=op_label, bg=DIALOG_BG)
            row_lbl.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            speed_ent = tk.Entry(frame, textvariable=speed_var, width=10)
            speed_ent.grid(row=i, column=1, padx=5, pady=2)
            power_ent = tk.Entry(frame, textvariable=power_var, width=10)
            power_ent.grid(row=i, column=2, padx=5, pady=2)
            passes_ent = tk.Spinbox(frame, textvariable=passes_var, from_=1, to=10, width=5)
            passes_ent.grid(row=i, column=3, padx=5, pady=2)
            add_tooltip(row_lbl, _("Settings for the {descr} on this material.").format(descr=descr))
            add_tooltip(speed_ent, _("Feed rate for the {short} on this material.").format(short=short))
            add_tooltip(power_ent, _("Laser power for the {short} on this material.").format(short=short))
            add_tooltip(passes_ent,
                        _("How many times to repeat the {short} on this material. Default 1; raise it for thicker stock where one high-power pass leaves rough edges.").format(short=short))

            air_var = tk.BooleanVar(value=mat_settings.get(f"air_assist_{op_key}", True))
            self.vars[mat_key][f'air_assist_{op_key}'] = air_var
            air_cb = tk.Checkbutton(frame, variable=air_var, bg=DIALOG_BG)
            air_cb.grid(row=i, column=4, padx=5, pady=2)
            add_tooltip(air_cb, _("Air-assist on during the {short} on this material.").format(short=short))

        # Kerf width row
        kerf_row = 4 + len(self.OPERATIONS)
        kerf_lbl = tk.Label(frame, text=_("Kerf width:"), bg=DIALOG_BG)
        kerf_lbl.grid(row=kerf_row, column=0, sticky="w", padx=5, pady=(8, 2))

        default_kerf = self._get_default(mat_key, "kerf_width")
        current_kerf = mat_settings.get("kerf_width", default_kerf)
        kerf_var = tk.DoubleVar(value=current_kerf if current_kerf else 0.0)
        self.vars[mat_key]['kerf_width'] = kerf_var

        kerf_entry = tk.Spinbox(frame, textvariable=kerf_var, from_=0.0, to=1.0,
                                increment=0.05, width=8, format="%.2f")
        kerf_entry.grid(row=kerf_row, column=1, sticky="w", padx=5, pady=(8, 2))
        kerf_unit = tk.Label(frame, text=_("mm"), bg=DIALOG_BG)
        kerf_unit.grid(row=kerf_row, column=2, sticky="w", padx=5, pady=(8, 2))
        add_tooltips(
            _("Full kerf width measured from a calibration cut on this "
            "material (hole ID minus disc OD). The app splits this in half "
            "automatically — outer cuts expand by half-kerf, hole "
            "cuts shrink by half-kerf. Note: LightBurn asks for half-kerf; "
            "this dialog wants the full measured width."),
            kerf_lbl, kerf_entry, kerf_unit,
        )

    def _create_tooling_engraving_section(self, parent):
        """Create the engraving settings section for tooling (die inserts)."""
        tooling = self.settings.get("tooling_settings", {})

        frame = tk.LabelFrame(parent, text=_("Die Engraving"), bg=DIALOG_BG, padx=10, pady=10)
        frame.pack(fill="x", pady=5, padx=5)

        # Engraving mode (filled / line)
        mode_frame = tk.Frame(frame, bg=DIALOG_BG)
        mode_frame.pack(fill="x", pady=(0, 5))

        tk.Label(mode_frame, text=_("Engraving:"), bg=DIALOG_BG).pack(side="left")
        self.tooling_eng_mode_var = tk.StringVar(value=tooling.get("engraving_mode", "filled"))
        die_filled_rb = tk.Radiobutton(mode_frame, text=_("Filled"), variable=self.tooling_eng_mode_var,
                                       value="filled", bg=DIALOG_BG)
        die_filled_rb.pack(side="left", padx=(5, 10))
        die_line_rb = tk.Radiobutton(mode_frame, text=_("Line"), variable=self.tooling_eng_mode_var,
                                     value="line", bg=DIALOG_BG)
        die_line_rb.pack(side="left")
        add_tooltip(die_filled_rb,
                    _("Solid raster fill of the size labels on dies. More "
                    "legible on dark acrylic but slower."))
        add_tooltip(die_line_rb,
                    _("Single-stroke outline of the size labels on dies. "
                    "Fastest; readable on light acrylic."))

        # Font sizes
        font_frame = tk.Frame(frame, bg=DIALOG_BG)
        font_frame.pack(fill="x", pady=(0, 5))

        ring_font_lbl = tk.Label(font_frame, text=_("Ring font size:"), bg=DIALOG_BG)
        ring_font_lbl.pack(side="left")
        self.tooling_ring_font_var = tk.DoubleVar(value=tooling.get("ring_font_size", 3.5))
        ring_font_ent = tk.Entry(font_frame, textvariable=self.tooling_ring_font_var, width=5)
        ring_font_ent.pack(side="left", padx=(2, 5))
        ring_font_unit = tk.Label(font_frame, text=_("mm"), bg=DIALOG_BG)
        ring_font_unit.pack(side="left", padx=(0, 15))
        add_tooltips(
            _("Font size in millimeters for the size label engraved on the "
            "die ring (the outer annulus that stays in the holder)."),
            ring_font_lbl, ring_font_ent, ring_font_unit,
        )

        cutout_font_lbl = tk.Label(font_frame, text=_("Cutout font size:"), bg=DIALOG_BG)
        cutout_font_lbl.pack(side="left")
        self.tooling_cutout_font_var = tk.DoubleVar(value=tooling.get("cutout_font_size", 3.5))
        cutout_font_ent = tk.Entry(font_frame, textvariable=self.tooling_cutout_font_var, width=5)
        cutout_font_ent.pack(side="left", padx=(2, 5))
        cutout_font_unit = tk.Label(font_frame, text=_("mm"), bg=DIALOG_BG)
        cutout_font_unit.pack(side="left")
        add_tooltips(
            _("Font size in millimeters for the actual-size label engraved "
            "on the inner cutout disc (the piece that drops out of the "
            "ring — useful as a pad-cup tool)."),
            cutout_font_lbl, cutout_font_ent, cutout_font_unit,
        )

        # Ring engraving placement
        loc_frame = tk.Frame(frame, bg=DIALOG_BG)
        loc_frame.pack(fill="x", pady=(0, 0))

        ring_loc_lbl = tk.Label(loc_frame, text=_("Ring engraving placement:"), bg=DIALOG_BG)
        ring_loc_lbl.pack(side="left")
        self.tooling_ring_loc_var = tk.StringVar(value=tooling.get("ring_engraving_location", "centered"))
        ring_loc_ctr = tk.Radiobutton(loc_frame, text=_("Centered"), variable=self.tooling_ring_loc_var,
                                      value="centered", bg=DIALOG_BG)
        ring_loc_ctr.pack(side="left", padx=(5, 5))
        ring_loc_out = tk.Radiobutton(loc_frame, text=_("From outside"), variable=self.tooling_ring_loc_var,
                                      value="from_outside", bg=DIALOG_BG)
        ring_loc_out.pack(side="left", padx=(0, 5))
        self.tooling_ring_offset_var = tk.DoubleVar(value=tooling.get("ring_engraving_offset", 0.0))
        ring_loc_ent = tk.Entry(loc_frame, textvariable=self.tooling_ring_offset_var, width=5)
        ring_loc_ent.pack(side="left", padx=(2, 2))
        ring_loc_unit = tk.Label(loc_frame, text=_("mm"), bg=DIALOG_BG)
        ring_loc_unit.pack(side="left")
        add_tooltip(ring_loc_lbl,
                    _("Where the size label sits on the die ring annulus."))
        add_tooltip(ring_loc_ctr,
                    _("Place the label centered between the inner hole and "
                    "the outer edge of the ring."))
        add_tooltip(ring_loc_out,
                    _("Place the label a fixed distance inward from the "
                    "outer edge of the ring (set the distance below)."))
        add_tooltips(
            _("Distance from the ring's outer edge, in millimeters (used "
            "only with the From outside option)."),
            ring_loc_ent, ring_loc_unit,
        )

    def _get_default(self, material, setting_key, fallback=100):
        """Get default value from DEFAULT_SETTINGS."""
        defaults = DEFAULT_SETTINGS.get("gcode_settings", {}).get(material, {})
        return defaults.get(setting_key, fallback)

    # ----- Per-material preset bar -----

    def _create_material_preset_bar(self, parent, mat_key, mat_label):
        """Build the preset dropdown + Load/Save/Rename/Delete row for one material."""
        bar = tk.Frame(parent, bg=DIALOG_BG)
        bar.pack(fill="x", pady=(0, 6))

        tk.Label(bar, text=_("Preset:"), bg=DIALOG_BG).pack(side="left")
        combo = ttk.Combobox(bar, state="readonly", width=22)
        combo.pack(side="left", padx=5, fill="x", expand=True)
        self.preset_combos[mat_key] = combo

        load_btn = tk.Button(bar, text=_("Load"),
                             command=lambda mk=mat_key: self._on_load_gcode_preset(mk))
        load_btn.pack(side="left", padx=2)
        save_btn = tk.Button(bar, text=_("Save"),
                             command=lambda mk=mat_key: self._on_save_gcode_preset(mk))
        save_btn.pack(side="left", padx=2)
        rename_btn = tk.Button(bar, text=_("Rename"),
                               command=lambda mk=mat_key: self._on_rename_gcode_preset(mk))
        rename_btn.pack(side="left", padx=2)
        del_btn = tk.Button(bar, text=_("Delete"),
                            command=lambda mk=mat_key: self._on_delete_gcode_preset(mk))
        del_btn.pack(side="left", padx=2)

        add_tooltip(combo, _("Saved laser settings for {label}. Pick one and click Load.").format(label=mat_label))
        add_tooltip(load_btn,
                    _("Fill the {label} fields below from the selected preset. "
                    "Click Apply at the bottom to commit it to the app.").format(label=mat_label))
        add_tooltip(save_btn,
                    _("Save the current {label} fields as a preset — overwrite "
                    "an existing one or create a new one.").format(label=mat_label))
        add_tooltip(rename_btn, _("Rename the selected {label} preset.").format(label=mat_label))
        add_tooltip(del_btn,
                    _("Delete the selected {label} preset (cannot be undone). "
                    "At least one preset must remain.").format(label=mat_label))

        self.active_preset_name[mat_key] = None
        self._refresh_gcode_preset_combo(mat_key)

    def _refresh_gcode_preset_combo(self, mat_key, select=None):
        """Sync the combobox values from self.gcode_presets[mat_key]."""
        names = sorted(self.gcode_presets.get(mat_key, {}).keys())
        self.preset_combos[mat_key]['values'] = names
        if select and select in names:
            self.preset_combos[mat_key].set(select)
        elif self.active_preset_name.get(mat_key) in names:
            self.preset_combos[mat_key].set(self.active_preset_name[mat_key])
        elif not names:
            self.preset_combos[mat_key].set("")

    # ----- Snapshot helpers -----

    def _capture_material_to_dict(self, mat_key):
        """Read every preset-tracked field for one material from its tk vars.

        Returns a dict with the same shape as a single material entry in
        DEFAULT_SETTINGS["gcode_settings"]. Used for dirty tracking AND
        for save-as-preset.
        """
        v = self.vars[mat_key]
        # Fill density slider maps to filled_line_spacing (same formula as _on_save).
        density_val = v['fill_density'].get()
        line_spacing = round(0.3 - (density_val / 100) * 0.22, 3)
        try:
            data = {
                "engraving_mode": v['engraving_mode'].get(),
                "engraving_speed": int(v['engraving']['speed'].get()),
                "engraving_power": float(v['engraving']['power'].get()),
                "engraving_passes": max(1, int(v['engraving']['passes'].get())),
                "filled_engraving_speed": int(v['filled_engraving']['speed'].get()),
                "filled_engraving_power": float(v['filled_engraving']['power'].get()),
                "filled_engraving_passes": max(1, int(v['filled_engraving']['passes'].get())),
                "filled_line_spacing": line_spacing,
                "hole_speed": int(v['hole']['speed'].get()),
                "hole_power": float(v['hole']['power'].get()),
                "hole_passes": max(1, int(v['hole']['passes'].get())),
                "cut_speed": int(v['cut']['speed'].get()),
                "cut_power": float(v['cut']['power'].get()),
                "cut_passes": max(1, int(v['cut']['passes'].get())),
                "kerf_width": float(v['kerf_width'].get()),
                "air_assist_engraving": bool(v['air_assist_engraving'].get()),
                "air_assist_filled_engraving": bool(v['air_assist_filled_engraving'].get()),
                "air_assist_hole": bool(v['air_assist_hole'].get()),
                "air_assist_cut": bool(v['air_assist_cut'].get()),
            }
        except (tk.TclError, ValueError):
            # Mid-edit invalid state. Return a sentinel that won't equal any
            # real snapshot so dirty-tracking treats this as "still dirty"
            # rather than crashing.
            return {"_invalid": True}
        return data

    def _apply_dict_to_material(self, mat_key, data):
        """Populate this material's tk vars from a preset dict."""
        v = self.vars[mat_key]
        defaults = DEFAULT_SETTINGS.get("gcode_settings", {}).get(mat_key, {})

        def g(key, fallback):
            return data.get(key, defaults.get(key, fallback))

        v['engraving_mode'].set(g("engraving_mode", "line"))
        v['engraving']['speed'].set(int(g("engraving_speed", 1200)))
        v['engraving']['power'].set(g("engraving_power", 8))
        v['engraving']['passes'].set(int(g("engraving_passes", 1)))
        v['filled_engraving']['speed'].set(int(g("filled_engraving_speed", 1200)))
        v['filled_engraving']['power'].set(g("filled_engraving_power", 8))
        v['filled_engraving']['passes'].set(int(g("filled_engraving_passes", 1)))

        spacing = g("filled_line_spacing", 0.15)
        density_val = int((0.3 - spacing) / 0.22 * 100)
        v['fill_density'].set(max(0, min(100, density_val)))

        v['hole']['speed'].set(int(g("hole_speed", 300)))
        v['hole']['power'].set(g("hole_power", 30))
        v['hole']['passes'].set(int(g("hole_passes", 1)))
        v['cut']['speed'].set(int(g("cut_speed", 600)))
        v['cut']['power'].set(g("cut_power", 60))
        v['cut']['passes'].set(int(g("cut_passes", 1)))

        v['kerf_width'].set(g("kerf_width", 0.0))
        v['air_assist_engraving'].set(bool(g("air_assist_engraving", True)))
        v['air_assist_filled_engraving'].set(bool(g("air_assist_filled_engraving", True)))
        v['air_assist_hole'].set(bool(g("air_assist_hole", True)))
        v['air_assist_cut'].set(bool(g("air_assist_cut", True)))

    def _material_is_dirty(self, mat_key):
        baseline = self.material_baseline.get(mat_key)
        if baseline is None:
            return False
        return self._capture_material_to_dict(mat_key) != baseline

    def _detect_active_preset(self, mat_key):
        """Find a saved preset whose data matches this material's current
        form snapshot, so users can see which preset is loaded when the
        dialog opens. Returns the preset name or None."""
        presets = self.gcode_presets.get(mat_key, {})
        if not presets:
            return None
        snapshot = self._capture_material_to_dict(mat_key)
        if snapshot.get("_invalid"):
            return None
        for name, data in presets.items():
            if data == snapshot:
                return name
        return None

    def _dirty_materials(self):
        """List of material keys whose form differs from their baseline."""
        if self.gcode_presets is None:
            return []
        return [mk for mk, _label in self.active_materials if self._material_is_dirty(mk)]

    def _set_material_baseline(self, mat_key):
        self.material_baseline[mat_key] = self._capture_material_to_dict(mat_key)

    # ----- Preset action handlers -----

    def _on_load_gcode_preset(self, mat_key):
        combo = self.preset_combos[mat_key]
        name = combo.get().strip()
        presets = self.gcode_presets.get(mat_key, {})
        if not name:
            messagebox.showinfo(_("Load Preset"),
                                _("Pick a preset from the dropdown first."),
                                parent=self.top)
            return
        if name not in presets:
            messagebox.showerror(_("Load Preset"),
                                 _("Preset '{name}' not found.").format(name=name),
                                 parent=self.top)
            return
        if self._material_is_dirty(mat_key):
            label = self.active_preset_name.get(mat_key) or _("the current values")
            if not messagebox.askyesno(
                _("Discard unsaved changes?"),
                _("You have unsaved edits to {label}.\n\nLoading '{name}' will discard them. Continue?")
                    .format(label=label, name=name),
                parent=self.top,
            ):
                return
        self._apply_dict_to_material(mat_key, presets[name])
        self.active_preset_name[mat_key] = name
        self._set_material_baseline(mat_key)

    def _on_save_gcode_preset(self, mat_key):
        mat_label = dict(self.active_materials).get(mat_key, mat_key)
        existing = sorted(self.gcode_presets.get(mat_key, {}).keys())
        dlg = SaveSizingPresetDialog(
            self.top,
            existing_names=existing,
            default_existing=self.active_preset_name.get(mat_key),
            title=_("Save {label} Preset").format(label=mat_label),
            intro=_("Save the current {label} fields as a preset.").format(label=mat_label),
        )
        result = dlg.result
        if result is None:
            return
        target = result["name"]
        snapshot = self._capture_material_to_dict(mat_key)
        if snapshot.get("_invalid"):
            messagebox.showerror(_("Save Preset"),
                                 _("One or more {label} fields has an invalid value. "
                                   "Fix it before saving as a preset.").format(label=mat_label),
                                 parent=self.top)
            return
        self.gcode_presets.setdefault(mat_key, {})[target] = snapshot
        self.gcode_presets_save_callback()
        self.active_preset_name[mat_key] = target
        self._set_material_baseline(mat_key)
        self._refresh_gcode_preset_combo(mat_key, select=target)

    def _on_rename_gcode_preset(self, mat_key):
        combo = self.preset_combos[mat_key]
        old = combo.get().strip()
        presets = self.gcode_presets.get(mat_key, {})
        if not old or old not in presets:
            messagebox.showinfo(_("Rename Preset"),
                                _("Pick a preset from the dropdown first."),
                                parent=self.top)
            return
        new = simpledialog.askstring(
            _("Rename Preset"),
            _("New name for '{old}':").format(old=old),
            initialvalue=old,
            parent=self.top,
        )
        if new is None:
            return
        new = new.strip()
        if not new:
            messagebox.showwarning(_("Rename Preset"),
                                   _("Preset name cannot be empty."),
                                   parent=self.top)
            return
        if new == old:
            return
        if new in presets:
            messagebox.showerror(_("Rename Preset"),
                                 _("A preset named '{new}' already exists.").format(new=new),
                                 parent=self.top)
            return
        presets[new] = presets.pop(old)
        if self.active_preset_name.get(mat_key) == old:
            self.active_preset_name[mat_key] = new
        self.gcode_presets_save_callback()
        self._refresh_gcode_preset_combo(mat_key, select=new)

    def _on_delete_gcode_preset(self, mat_key):
        combo = self.preset_combos[mat_key]
        name = combo.get().strip()
        presets = self.gcode_presets.get(mat_key, {})
        mat_label = dict(self.active_materials).get(mat_key, mat_key)
        if not name or name not in presets:
            messagebox.showinfo(_("Delete Preset"),
                                _("Pick a preset from the dropdown first."),
                                parent=self.top)
            return
        if len(presets) <= 1:
            messagebox.showinfo(
                _("Delete Preset"),
                _("At least one {label} preset must remain. Save another "
                  "preset before deleting this one.").format(label=mat_label),
                parent=self.top,
            )
            return
        if not messagebox.askyesno(_("Delete Preset"),
                                   _("Delete preset '{name}'?").format(name=name),
                                   parent=self.top):
            return
        del presets[name]
        if self.active_preset_name.get(mat_key) == name:
            self.active_preset_name[mat_key] = None
        self.gcode_presets_save_callback()
        self._refresh_gcode_preset_combo(mat_key)

    def _on_apply_clicked(self):
        """Apply button: warn about dirty materials, then commit + close."""
        if not self._prompt_dirty(context="apply"):
            return
        self._on_save()

    def _on_cancel_clicked(self):
        """Cancel button / window X: warn about dirty materials, then close."""
        if not self._prompt_dirty(context="cancel"):
            return
        self.top.destroy()

    def _prompt_dirty(self, context):
        """Three-way prompt for unsaved per-material edits.

        context="apply" or "cancel". Returns True if the caller should
        proceed (apply settings, or close window). Returns False if the
        user picked "keep editing" or a save-as-preset round-trip was
        cancelled.
        """
        if self.gcode_presets is None:
            return True
        dirty = self._dirty_materials()
        if not dirty:
            return True

        labels = ", ".join(dict(self.active_materials).get(mk, mk) for mk in dirty)
        if context == "apply":
            title = _("Unsaved changes")
            msg = _(
                "You have edits to {labels} that aren't saved as a preset.\n\n"
                "Save them as preset(s)?\n\n"
                "• Yes — save each as a preset, then apply.\n"
                "• No — apply anyway (edits stay in settings but aren't a preset).\n"
                "• Cancel — keep editing."
            ).format(labels=labels)
        else:
            title = _("Unsaved changes")
            msg = _(
                "You have edits to {labels} that aren't saved as a preset.\n\n"
                "Save them as preset(s) before closing?\n\n"
                "• Yes — save each as a preset, then close (changes won't apply).\n"
                "• No — discard edits and close.\n"
                "• Cancel — keep editing."
            ).format(labels=labels)

        choice = messagebox.askyesnocancel(title, msg, parent=self.top)
        if choice is None:
            return False  # keep editing
        if choice is False:
            return True  # proceed without saving as preset
        # Yes: walk each dirty material through the save-preset flow.
        for mk in dirty:
            self._on_save_gcode_preset(mk)
            # If the user cancelled the save dialog the material is still
            # dirty. Re-prompt? Simpler: bail out so they can fix it.
            if self._material_is_dirty(mk):
                return False
        return True

    def _on_save(self):
        """Save settings and close."""
        # Start with existing settings to preserve materials not shown in this dialog
        new_gcode_settings = dict(self.settings.get("gcode_settings", {}))

        for mat_key, _ in self.active_materials:
            new_gcode_settings[mat_key] = {}

            # Engraving mode
            new_gcode_settings[mat_key]["engraving_mode"] = self.vars[mat_key]['engraving_mode'].get()

            # Line engraving speed/power/passes
            try:
                new_gcode_settings[mat_key]["engraving_speed"] = self.vars[mat_key]['engraving']['speed'].get()
                new_gcode_settings[mat_key]["engraving_power"] = self.vars[mat_key]['engraving']['power'].get()
                new_gcode_settings[mat_key]["engraving_passes"] = max(1, int(self.vars[mat_key]['engraving']['passes'].get()))
            except tk.TclError:
                messagebox.showerror(_("Invalid Input"),
                                     _("Invalid value for {mat} line engraving. Please enter valid numbers.").format(mat=mat_key),
                                     parent=self.top)
                return

            # "Filled" engraving speed/power/passes
            try:
                new_gcode_settings[mat_key]["filled_engraving_speed"] = self.vars[mat_key]['filled_engraving']['speed'].get()
                new_gcode_settings[mat_key]["filled_engraving_power"] = self.vars[mat_key]['filled_engraving']['power'].get()
                new_gcode_settings[mat_key]["filled_engraving_passes"] = max(1, int(self.vars[mat_key]['filled_engraving']['passes'].get()))
            except tk.TclError:
                messagebox.showerror(_("Invalid Input"),
                                     _("Invalid value for {mat} filled engraving. Please enter valid numbers.").format(mat=mat_key),
                                     parent=self.top)
                return

            # Fill density slider -> line spacing
            density_val = self.vars[mat_key]['fill_density'].get()
            line_spacing = 0.3 - (density_val / 100) * 0.22
            new_gcode_settings[mat_key]["filled_line_spacing"] = round(line_spacing, 3)

            # Other operations (hole, cut)
            for op_key, _ in self.OPERATIONS:
                try:
                    speed = self.vars[mat_key][op_key]['speed'].get()
                    power = self.vars[mat_key][op_key]['power'].get()
                    passes = max(1, int(self.vars[mat_key][op_key]['passes'].get()))
                    new_gcode_settings[mat_key][f"{op_key}_speed"] = speed
                    new_gcode_settings[mat_key][f"{op_key}_power"] = power
                    new_gcode_settings[mat_key][f"{op_key}_passes"] = passes
                except tk.TclError:
                    messagebox.showerror(_("Invalid Input"),
                                         _("Invalid value for {mat} {op}. Please enter valid numbers.").format(mat=mat_key, op=op_key),
                                         parent=self.top)
                    return

            # Kerf width
            try:
                new_gcode_settings[mat_key]["kerf_width"] = self.vars[mat_key]['kerf_width'].get()
            except tk.TclError:
                messagebox.showerror(_("Invalid Input"),
                                     _("Invalid kerf width for {mat}. Please enter a valid number.").format(mat=mat_key),
                                     parent=self.top)
                return

            # Air assist toggles
            new_gcode_settings[mat_key]["air_assist_engraving"] = self.vars[mat_key]['air_assist_engraving'].get()
            new_gcode_settings[mat_key]["air_assist_filled_engraving"] = self.vars[mat_key]['air_assist_filled_engraving'].get()
            new_gcode_settings[mat_key]["air_assist_hole"] = self.vars[mat_key]['air_assist_hole'].get()
            new_gcode_settings[mat_key]["air_assist_cut"] = self.vars[mat_key]['air_assist_cut'].get()

        self.settings["gcode_settings"] = new_gcode_settings
        self.settings["filled_overscan_enabled"] = self.overscan_var.get()

        # Save tooling engraving settings if shown
        if self.show_tooling_engraving:
            tooling = self.settings.get("tooling_settings", {})
            tooling["engraving_mode"] = self.tooling_eng_mode_var.get()
            try:
                tooling["ring_font_size"] = self.tooling_ring_font_var.get()
                tooling["cutout_font_size"] = self.tooling_cutout_font_var.get()
                tooling["ring_engraving_offset"] = self.tooling_ring_offset_var.get()
            except tk.TclError:
                messagebox.showerror(_("Invalid Input"), _("Font sizes must be valid numbers."), parent=self.top)
                return
            tooling["ring_engraving_location"] = self.tooling_ring_loc_var.get()
            self.settings["tooling_settings"] = tooling

        # Global G-code settings
        try:
            self.settings["gcode_return_speed"] = self.return_speed_var.get()
        except tk.TclError:
            messagebox.showerror(_("Invalid Input"), _("Return speed must be a valid number."), parent=self.top)
            return
        self.settings["gcode_cut_grouping"] = self.grouping_var.get()

        self.save_callback(self.settings)

        self.top.destroy()

    def _reset_defaults(self):
        """Reset all values to defaults."""
        default_gcode = DEFAULT_SETTINGS.get("gcode_settings", {})

        for mat_key, _ in self.active_materials:
            mat_defaults = default_gcode.get(mat_key, {})

            # Reset engraving mode
            self.vars[mat_key]['engraving_mode'].set(mat_defaults.get("engraving_mode", "line"))

            # Reset line engraving
            self.vars[mat_key]['engraving']['speed'].set(mat_defaults.get("engraving_speed", 1200))
            self.vars[mat_key]['engraving']['power'].set(mat_defaults.get("engraving_power", 8))
            self.vars[mat_key]['engraving']['passes'].set(mat_defaults.get("engraving_passes", 1))

            # Reset filled engraving
            self.vars[mat_key]['filled_engraving']['speed'].set(mat_defaults.get("filled_engraving_speed", 1000))
            self.vars[mat_key]['filled_engraving']['power'].set(mat_defaults.get("filled_engraving_power", 12))
            self.vars[mat_key]['filled_engraving']['passes'].set(mat_defaults.get("filled_engraving_passes", 1))

            # Reset fill density
            default_spacing = mat_defaults.get("filled_line_spacing", 0.15)
            density_val = int((0.3 - default_spacing) / 0.22 * 100)
            density_val = max(0, min(100, density_val))
            self.vars[mat_key]['fill_density'].set(density_val)

            # Reset other operations
            for op_key, _ in self.OPERATIONS:
                default_speed = mat_defaults.get(f"{op_key}_speed", 100)
                default_power = mat_defaults.get(f"{op_key}_power", 10)
                default_passes = mat_defaults.get(f"{op_key}_passes", 1)
                self.vars[mat_key][op_key]['speed'].set(default_speed)
                self.vars[mat_key][op_key]['power'].set(default_power)
                self.vars[mat_key][op_key]['passes'].set(default_passes)

            # Reset kerf width
            default_kerf = mat_defaults.get("kerf_width", 0.0)
            self.vars[mat_key]['kerf_width'].set(default_kerf)

            # Reset air assist
            self.vars[mat_key]['air_assist_engraving'].set(mat_defaults.get("air_assist_engraving", True))
            self.vars[mat_key]['air_assist_filled_engraving'].set(mat_defaults.get("air_assist_filled_engraving", True))
            self.vars[mat_key]['air_assist_hole'].set(mat_defaults.get("air_assist_hole", True))
            self.vars[mat_key]['air_assist_cut'].set(mat_defaults.get("air_assist_cut", True))

        # Reset overscan
        self.overscan_var.set(DEFAULT_SETTINGS.get("filled_overscan_enabled", False))

        # Reset tooling engraving settings
        if self.show_tooling_engraving:
            default_tooling = DEFAULT_SETTINGS.get("tooling_settings", {})
            self.tooling_eng_mode_var.set(default_tooling.get("engraving_mode", "filled"))
            self.tooling_ring_font_var.set(default_tooling.get("ring_font_size", 3.5))
            self.tooling_cutout_font_var.set(default_tooling.get("cutout_font_size", 3.5))
            self.tooling_ring_loc_var.set(default_tooling.get("ring_engraving_location", "centered"))
            self.tooling_ring_offset_var.set(default_tooling.get("ring_engraving_offset", 0.0))

        # Reset global settings
        self.return_speed_var.set(DEFAULT_SETTINGS.get("gcode_return_speed", 1000))
        self.grouping_var.set(DEFAULT_SETTINGS.get("gcode_cut_grouping", "layer"))


# ==========================================
# HELP / ABOUT DIALOGS
# ==========================================

class UserGuideWindow(tk.Toplevel):
    """Scrollable window showing the user guide, optionally filtered by section."""

    # Map section names to display titles
    SECTION_TITLES = {
        "pad_generator": _("Pad Maker"),
        "key_heights": _("Key Height Library"),
        "serial_lookup": _("Serial Lookup"),
        "screw_specs": _("Screw Specs"),
        "tooling": _("Tooling"),
        "tuner": _("Tuner"),
        "toner": _("Toner"),
    }

    def __init__(self, parent, section=None):
        super().__init__(parent)
        self._section = section
        if section and section in self.SECTION_TITLES:
            self.title(_("User Guide — {section}").format(section=self.SECTION_TITLES[section]))
        else:
            self.title(_("User Guide"))
        self.geometry("620x700")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)

        # Scrollable text widget
        text_frame = tk.Frame(self, bg=DIALOG_BG)
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side="right", fill="y")

        self.text = tk.Text(text_frame, wrap="word", yscrollcommand=scrollbar.set,
                            font=("Helvetica", 11), padx=15, pady=10)
        self.text.pack(fill="both", expand=True)
        scrollbar.config(command=self.text.yview)

        # Configure tags
        self.text.tag_configure("h1", font=("Helvetica", 16, "bold"), spacing3=6)
        self.text.tag_configure("h2", font=("Helvetica", 13, "bold"), spacing1=12, spacing3=4)
        self.text.tag_configure("body", font=("Helvetica", 11), spacing1=2)
        self.text.tag_configure("bullet", font=("Helvetica", 11), lmargin1=20, lmargin2=35)

        self._insert_content()
        self.text.config(state="disabled")

        # Bottom buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=10)
        if section:
            tk.Button(btn_frame, text=_("Show Full Guide"),
                      command=self._show_all).pack(side="left", padx=(0, 10))
        tk.Button(btn_frame, text=_("Close"), command=self.destroy).pack(side="left")

    def _show_all(self):
        """Reload with all sections visible."""
        self._section = None
        self.title(_("User Guide"))
        self.text.config(state="normal")
        self.text.delete("1.0", "end")
        self._insert_content()
        self.text.config(state="disabled")

    def _h1(self, text):
        self.text.insert("end", text + "\n", "h1")

    def _h2(self, text):
        self.text.insert("end", text + "\n", "h2")

    def _body(self, text):
        self.text.insert("end", text + "\n", "body")

    def _bullet(self, text):
        self.text.insert("end", "  \u2022 " + text + "\n", "bullet")

    def _blank(self):
        self.text.insert("end", "\n")

    def _link(self, display_text, url):
        import webbrowser
        tag = f"link_{id(url)}"
        self.text.tag_configure(tag, font=("Helvetica", 11, "underline"),
                                foreground="#0066CC", lmargin1=20)
        self.text.insert("end", "  " + display_text + "\n", tag)
        self.text.tag_bind(tag, "<Button-1>", lambda e: webbrowser.open(url))
        self.text.tag_bind(tag, "<Enter>",
                           lambda e: self.text.config(cursor="hand2"))
        self.text.tag_bind(tag, "<Leave>",
                           lambda e: self.text.config(cursor=""))

    def _insert_content(self):
        section = self._section

        if section is None:
            # Full guide
            self._h1(_("Stohrer Sax Shop Companion"))
            self._body(_("A tool for saxophone technicians: generate laser-cutting templates, "
                        "record key heights, look up serial numbers, reference screw specs, "
                        "tune, and analyze tone."))
            self._blank()
            self._section_pad_generator()
            self._section_key_heights()
            self._section_serial_lookup()
            self._section_screw_specs()
            self._section_tooling()
            self._section_tuner()
            self._section_toner()
            self._section_import_export()
            self._section_padmaking_guide()
        else:
            dispatch = {
                "pad_generator": self._section_pad_generator,
                "key_heights": self._section_key_heights,
                "serial_lookup": self._section_serial_lookup,
                "screw_specs": self._section_screw_specs,
                "tooling": self._section_tooling,
                "tuner": self._section_tuner,
                "toner": self._section_toner,
            }
            fn = dispatch.get(section)
            if fn:
                fn()
            # Always show the padmaking link and import/export
            self._section_import_export()
            self._section_padmaking_guide()

    def _section_pad_generator(self):
        self._h2(_("Pad Maker (SVG / G-code)"))
        self._body(_("Enter pad sizes in the text area, one per line, in the format "
                    "\"size x quantity\" (e.g. \"42.0 x 3\"). Select one or more materials "
                    "and click Generate to create laser-cutting files."))
        self._bullet(_("Felt: disc diameter = pad size minus felt offset"))
        self._bullet(_("Card: further reduced by card-to-felt offset"))
        self._bullet(_("Leather: enlarged to wrap around felt, with star/dart pattern for small pads"))
        self._bullet(_("Exact Size: no offset applied (SVG only, not available for G-code)"))
        self._bullet(_("\"max\" quantity (e.g. \"18.0 x max\"): fills remaining sheet space with that size. "
                      "Only one \"max\" entry is allowed per sheet."))
        self._blank()

        self._h2(_("SVG vs G-code Output"))
        self._bullet(_("SVG: each operation (engraving, holes, cuts) on its own color layer for "
                      "LightBurn or similar. Use this if your software imports SVG."))
        self._bullet(_("G-code: standalone Grbl with speeds and power baked in. Use this if your "
                      "laser reads G-code directly (e.g. Creality Falcon)."))
        self._body(_("Files are named with your base name plus the material "
                    "(\"my_pad_job_felt.svg\"). One file per selected material."))
        self._blank()

        self._h2(_("Units"))
        self._body(_("Pad sizes and sheet dimensions use the unit set in Options > Sizing Rules "
                    "(inches, mm, or cm). Output files are always in mm."))
        self._blank()

        self._h2(_("Center Hole"))
        self._body(_("Select a center hole size for rivet or screw mounting. Choose None, "
                    "3.0mm, 3.5mm, or enter a custom diameter."))
        self._bullet(_("Pads below the minimum hole size threshold (set in Sizing Rules) "
                      "skip the center hole automatically \u2014 they're too small for it to be useful."))
        self._blank()

        self._h2(_("Sheet Size & Fit to Paper"))
        self._body(_("Enter the width and height of your material sheet. For card stock, "
                    "check \"Fit card to paper\" to use a standard paper size (letter or A4) "
                    "instead of the sheet dimensions."))
        self._blank()

        self._h2(_("Engraving"))
        self._body(_("Each disc is engraved with its pad size number for identification. "
                    "Engraving settings and placement are configured in Options > Sizing Rules."))
        self._bullet(_("Position modes: distance from outside edge, distance from inside "
                      "(center hole), or centered between the two"))
        self._bullet(_("Font size is set per material"))
        self._bullet(_("Both engraving settings (on/off, font sizes) and placement (position modes) "
                      "support Universal or Per Size Range mode, just like sizing rules"))
        self._bullet(_("On small pads, text shifts toward the center to fit; only scales down as "
                      "a last resort."))
        self._bullet(_("If the font exceeds 80 percent of the disc radius, engraving is skipped for that "
                      "pad (warning shown before generating)."))
        self._blank()

        self._h2(_("Pad Presets"))
        self._body(_("Save frequently-used pad lists as presets for quick recall. "
                    "Presets are organized into libraries (e.g. \"My Presets\", \"Customer Jobs\")."))
        self._bullet(_("Save as Preset: saves the current pad list to the selected library"))
        self._bullet(_("Select a preset from the dropdown to load it into the text area"))
        self._bullet(_("Delete Preset: removes the selected preset"))
        self._bullet(_("File > Import/Export Pad Presets: share preset files with colleagues"))
        self._bullet(_("File > Import Matt's Pad Sets: downloads reference pad sets from stohrermusic.com"))
        self._blank()

        self._h2(_("Materials & Sizing Rules"))
        self._body(_("Options > Sizing Rules configures how disc sizes are calculated from "
                    "the pad size you enter:"))
        self._bullet(_("Felt Offset: how much smaller the felt disc is than the pad cup "
                      "(e.g. 0.75mm means the felt disc is 0.75mm smaller in diameter)"))
        self._bullet(_("Card-to-Felt Offset: additional reduction for cardboard backing "
                      "(added on top of the felt offset)"))
        self._bullet(_("Leather Wrap Multiplier: controls how much extra leather is added "
                      "for wrapping around the felt (1.0 = standard wrap)"))
        self._bullet(_("Felt Thickness: the thickness of felt being used, affects leather "
                      "wrap calculation"))
        self._bullet(_("Min. Pad Size for Hole: pads below this size skip the center hole"))
        self._body(_("Sizing rules, engraving settings, engraving placement, and star/dart "
                    "settings each have a Universal/Range toggle. In Universal mode (default), "
                    "one set of values applies to all pad sizes. In Range mode, define multiple "
                    "size ranges with different values for each. Pads not covered by any range "
                    "fall back to the universal values. Star/dart ranges work slightly differently: "
                    "pads not in a range simply get no star pattern."))
        self._blank()
        self._body(_("Darts: for leather pads below a size threshold, a dart "
                    "pattern is added so the leather can fold around the "
                    "felt. Overwrap, wrap bonus, frequency multiplier, and "
                    "shape (triangle / sine / square) are all adjustable."))
        self._blank()
        self._body(_("Sizing-rules presets: at the top of the Sizing Rules dialog you can "
                    "save the entire configuration as a named preset, load presets from the "
                    "dropdown, and export/import preset files to share with other techs."))
        self._bullet(_("Save Preset: snapshots every value in the dialog (sizing, darts, engraving, placement) "
                      "\u2014 overwrite an existing preset or save as a new one"))
        self._bullet(_("Load: fills the dialog from the selected preset \u2014 click Apply to commit"))
        self._bullet(_("Import/Export: JSON files containing one or more presets"))
        self._blank()

        self._h2(_("G-code Settings"))
        self._body(_("Options > G-code Settings configures the laser parameters for each material."))
        self._bullet(_("Speed & Power: set per operation (engraving, center hole, outer cut)"))
        self._bullet(_("Engraving Mode: \"Line\" (single-stroke outline) or \"Filled\" (scan-line raster fill)"))
        self._bullet(_("Fill Density: controls scan line spacing for filled engraving"))
        self._bullet(_("Overscan: extends filled scan lines so the laser is at full speed at character edges"))
        self._bullet(_("Kerf Width: full beam width. The app compensates each cut (expands outer "
                      "cuts, shrinks holes). Note: LightBurn asks for half-kerf; here, enter the "
                      "full measured width."))
        self._bullet(_("Air Assist: per-layer toggle for air assist (M8 on / M9 off)"))
        self._bullet(_("Return Speed: how fast the head returns home after the job (slower avoids endstop issues)"))
        self._bullet(_("Cut Grouping: \"By layer\" does all engravings, then all holes, then all cuts; "
                      "\"By pad\" completes each pad before moving to the next"))
        self._blank()

        self._h2(_("Layer Colors"))
        self._body(_("Options > Layer Colors maps each operation to a LightBurn color layer "
                    "(numbered 00 through 29). This only affects SVG output for use in LightBurn."))
        self._blank()

        self._h2(_("Custom Shapes"))
        self._body(_("\"Draw / Capture Shape\" defines an irregular polygon for leather skins "
                    "or scrap pieces. The nesting algorithm fits circles inside the polygon "
                    "instead of the rectangular sheet."))
        self._bullet(_("Click points on the grid to define the outline. Vertex placement is "
                      "free \u2014 not snapped to grid intersections."))
        self._bullet(_("Grid sized to cover the laser bed (default 17 in / 43 cm; auto-grown "
                      "if your bed is larger)."))
        self._bullet(_("\"Show live camera underneath\" overlays the camera feed at 1:1 scale "
                      "so you can trace your scrap by eye. Requires a saved camera calibration."))
        self._bullet(_("\"Get from camera\" auto-detects the scrap outline. A safety inset "
                      "(default 3 mm, Options > Machine) shrinks the polygon to absorb minor "
                      "camera measurement error at the edges."))
        self._bullet(_("Unload clears the shape and returns to rectangle mode."))
        self._blank()

        self._h2(_("Scrap Mode"))
        self._body(_("Place pads across multiple irregular pieces instead of one sheet."))
        self._bullet(_("One material at a time."))
        self._bullet(_("Set dimensions or draw / capture a shape for each scrap. Generate "
                      "fits what it can and tracks the remainder."))
        self._bullet(_("Between scraps: keep the loaded shape, unload it, or re-capture "
                      "from camera (when calibrated)."))
        self._bullet(_("Files named _scrap1, _scrap2, etc."))
        self._bullet(_("With 75+ pads, an opt-in popup offers \"large-batch optimization\": "
                      "the nester tries multiple disc orderings per scrap and keeps the best "
                      "result. Costs extra compute, fits more pads."))
        self._blank()

        self._h2(_("Edge Bias"))
        self._body(_("The Edge Bias d-pad control lets you tell the nesting algorithm which "
                    "direction to pack circles toward. This is useful for leather skins or "
                    "scrap pieces where some edges are cleaner than others."))
        self._bullet(_("Click an arrow to bias packing toward that edge or corner"))
        self._bullet(_("Click the center dot to return to default behavior (no bias)"))
        self._bullet(_("Cardinal directions (N, S, E, W) scan from that edge inward, "
                      "filling row by row or column by column"))
        self._bullet(_("Corner directions (NW, NE, SW, SE) radiate outward from the corner "
                      "\u2014 small pads nestle into the corner first, larger ones fan out "
                      "behind them in a wedge pattern"))
        self._bullet(_("For custom polygon shapes, positions closer to the biased edge or "
                      "corner are scored more favorably"))
        self._bullet(_("Example: a triangular scrap with two clean edges forming a right angle "
                      "and a rough hypotenuse \u2014 bias toward the corner where the good "
                      "edges meet so pads pack there first"))
        self._bullet(_("The setting is saved and persists between sessions"))
        self._blank()

        self._h2(_("Nesting Preview"))
        self._body(_("Check \"Preview before saving\" to see how your pads will be "
                    "arranged on the sheet before any files are written."))
        self._bullet(_("Preview works with one material at a time \u2014 select a single "
                      "material to use it"))
        self._bullet(_("The preview shows the sheet boundary (or custom polygon) with "
                      "circles at their nested positions, labeled with pad sizes, and "
                      "a material usage percentage"))
        self._bullet(_("Click Save Files to proceed to file generation"))
        self._bullet(_("Click Adjust to go back and change edge bias, sheet dimensions, "
                      "custom polygon shape, or pad sizes/quantities, then generate again"))
        self._bullet(_("Works in scrap mode too \u2014 preview each scrap piece before "
                      "committing. Combine with edge bias and custom polygon shapes "
                      "to optimize irregular scrap pieces."))
        self._blank()

        self._h2(_("Machine Integration (experimental, opt-in)"))
        self._body(_("File > Feature Set > \"Experimental: machine integration\" enables "
                    "direct USB serial control of a Grbl-compatible laser (Creality "
                    "Falcon2 Pro 40W and similar). Off by default; toggle on to use any "
                    "of the machine features below."))
        self._bullet(_("Options > Machine: Home Laser, Test Connection, Clear Errors, "
                      "Reset Falcon, Camera Calibration, Camera-Polygon Inset Margin."))
        self._bullet(_("Camera features (overlay, Get from camera, Frame & Cut auto-locate) "
                      "stay greyed until you've run the one-time camera calibration."))
        self._blank()

        self._h2(_("Camera Calibration (one-time)"))
        self._body(_("Establishes the camera-to-bed mapping so the app can convert what the "
                    "camera sees into machine coords. Two steps:"))
        self._bullet(_("Engrave a ChArUco card on basswood. Place + secure the basswood; "
                      "home the laser; jog to roughly bed center; click Frame to verify "
                      "placement; click Engrave."))
        self._bullet(_("Capture 12 frames with the camera. The first capture(s) MUST be "
                      "with the card still at its engraved position — that anchors camera "
                      "to machine coords. Then move the card around for the remaining "
                      "captures so the math solves lens distortion."))
        self._body(_("Saved to a JSON file in the config directory; reused across sessions. "
                    "Engrave reads your basswood preset (Tooling > Options)."))
        self._blank()

        self._h2(_("Frame & Cut"))
        self._body(_("When a Falcon is connected, a Frame & Cut button appears next to "
                    "Generate SVG / Generate G-code. It generates the G-code in memory, "
                    "lets you position the head, traces the cut outline at low power for "
                    "verification, then streams the cut to the laser."))
        self._bullet(_("Position dialog: Home Laser button (re-home if MPos has drifted); "
                      "jog arrows; Try Auto Locate (drives head to the polygon's bottom-"
                      "left vertex when available and the laser is homed)."))
        self._bullet(_("Framing loops at low power until you click \"Looks Good — Cut!\" — "
                      "lets you jog between passes to fine-tune alignment."))
        self._bullet(_("For tilted scraps: jog to the visible bottom-left corner of the "
                      "material. The G92 work-origin offset bridges the gap between that "
                      "corner and the polygon's bbox origin so both framing and cutting "
                      "land where you expect."))
        self._bullet(_("Pause / Resume / Stop act in real time. Stop is a soft-reset."))
        self._blank()

        self._h2(_("SD Card & Eject"))
        self._body(_("Check \"Eject SD card after G-code export\" below the Generate buttons. "
                    "When generating G-code to a removable drive (USB/SD card), the app "
                    "will automatically eject it when done so you can safely remove it. "
                    "If the destination isn't a removable drive, nothing extra happens. "
                    "(Windows only)"))
        self._blank()

    def _section_key_heights(self):
        self._h2(_("Key Height Library"))
        self._body(_("Record and compare key height measurements for different saxophones. "
                    "Organize sets into libraries, import/export, and share with colleagues."))
        self._bullet(_("Options > Layout Options controls which key fields are visible"))
        self._bullet(_("File > Import Matt's Key Heights downloads reference data from stohrermusic.com"))
        self._blank()

    def _section_serial_lookup(self):
        self._h2(_("Serial Lookup"))
        self._body(_("Look up saxophone serial number ranges by manufacturer to estimate year of production."))
        self._blank()

    def _section_screw_specs(self):
        self._h2(_("Screw Specs"))
        self._body(_("Reference database of screw thread specifications for different saxophone models."))
        self._bullet(_("File > Import Matt's Specs downloads the latest data from stohrermusic.com"))
        self._bullet(_("Import/export to share specs with colleagues"))
        self._blank()

    def _section_tooling(self):
        self._h2(_("Tooling \u2014 Die Inserts"))
        self._body(_("Generate laser-cutting files (SVG or G-code) for acrylic pad die inserts. "
                    "Dies are rings with a fixed outer diameter and an inner hole matching the "
                    "pad size."))
        self._bullet(_("Small dies (pad sizes 7.0\u201339.5mm): 50mm outer diameter"))
        self._bullet(_("Large dies (pad sizes 40.0\u201360.0mm): 70mm outer diameter"))
        self._bullet(_("Enter sizes as individual values (\"7, 8.5, 25\"), ranges (\"15-30\"), "
                      "or use the Full Set buttons. Step size controls the increment for ranges "
                      "(default 0.5mm)."))
        self._bullet(_("\"Engrave size on ring\" labels the die ring with the pad size number"))
        self._bullet(_("\"Engrave size on cutout\" labels the inner disc with its actual physical "
                      "size after kerf deduction"))
        self._blank()

        self._h2(_("Tooling \u2014 Cutout Discs as Pad Cup Tools"))
        self._body(_("When a die ring is laser-cut, the inner cutout falls out as a solid disc "
                    "(pad cup diameter minus kerf). These work as pad cup stiffeners, rim "
                    "rounders, leveling helpers, bending braces. For precise diameter control, "
                    "use the Pad Maker tab's \"Exact Size\" material instead \u2014 note that "
                    "Exact Size outputs SVG only, so import the SVG into your laser software "
                    "and set acrylic speeds and powers there."))
        self._blank()

        self._h2(_("Tooling \u2014 Die Holders"))
        self._body(_("Generate laser-cutting files for the acrylic die holder assembly. "
                    "Each holder is a stack of 85mm discs cemented together permanently:"))
        self._bullet(_("Solid bottom disc"))
        self._bullet(_("Magnet disc (6.5mm hole)"))
        self._bullet(_("Pin discs (3.5mm alignment holes) \u2014 2 for 5-layer, 3 for 6-layer"))
        self._bullet(_("Retaining ring (inner diameter matches the die size class)"))
        self._body(_("Pick the layer count, then the variant: Large (70mm inner), Small "
                    "(50mm inner), or Both \u2014 two complete independent holders, one of "
                    "each size, nested onto the same sheet. Set your sheet width and "
                    "height (in or mm); if the pieces don't fit, generation stops with "
                    "a clear minimum size message."))
        self._blank()

        self._h2(_("Tooling \u2014 Scrap Mode"))
        self._body(_("A full set of dies (107 sizes from 7\u201360mm) takes many sheets of acrylic. "
                    "Check Scrap Mode to spread die generation across multiple sheets:"))
        self._bullet(_("Enter your die sizes and sheet dimensions"))
        self._bullet(_("Generate \u2014 what fits is saved, the rest is tracked"))
        self._bullet(_("Adjust the sheet size if needed and generate again"))
        self._bullet(_("A progress window shows remaining and completed dies"))
        self._bullet(_("Files are named with _scrap1, _scrap2, etc. suffixes"))
        self._blank()

        self._h2(_("Tooling \u2014 Die Organizer"))
        self._body(_("Generate SVGs for a stackable die organizer (230 \u00d7 330 mm). Two parts: "
                    "Upper plate with slots for the dies, Lower base plate."))
        self._bullet(_("Cut three Uppers and one Lower (wood works great; acrylic does too)."))
        self._bullet(_("Align the four 1/8\" corner holes and glue the stack together \u2014 a book "
                      "press lightly clamps it nicely while the glue sets."))
        self._bullet(_("Resize the alignment holes to whatever pin you have."))
        self._bullet(_("SVG only \u2014 open in LightBurn or your laser software to cut."))
        self._blank()

        self._h2(_("Tooling \u2014 Kerf Test"))
        self._body(_("A quick test pattern (three circles at 10/20/30mm) for measuring kerf on any "
                    "material."))
        self._bullet(_("Pick the material, cut the pattern, pop out the three discs."))
        self._bullet(_("Measure hole ID and disc OD with calipers."))
        self._bullet(_("Kerf = hole ID \u2212 disc OD (full width of material vaporized)."))
        self._bullet(_("Enter the full kerf in G-code Settings; the app handles compensation."))
        self._body(_("Defaults to your existing speed/power for that material. Uncheck \"Use "
                    "existing settings\" for a one-off test with custom values."))
        self._blank()

        self._h2(_("Tooling \u2014 Speed & Power Test"))
        self._body(_("Beta. Generates a sheet of small test discs, each cut at a different "
                    "combination of speed/power/passes. Each disc is engraved with its 2-digit "
                    "ID; a legend.txt saved alongside the G-code maps each ID to its parameters. "
                    "Cut the sheet, inspect which discs came free cleanly without scorching, then "
                    "use those values as your starting point for that material."))
        self._bullet(_("Pick the material and click \"Apply material defaults\" to pre-fill the "
                      "speed/power fields with sensible starting values for that material."))
        self._bullet(_("Disc diameter is yours to set. Smaller saves material; bigger helps the "
                      "engraving stay legible. 20 mm is a good default."))
        self._bullet(_("Three sweep rows \u2014 Speed, Power, Passes. Check \"Sweep\" to test "
                      "\"Stops\" evenly-spaced values from Start to End; uncheck to hold constant "
                      "at \"Value\". You can sweep 0, 1, 2, or 3 variables \u2014 a single disc, "
                      "a row, a grid, or a 3-D block."))
        self._bullet(_("Engraving speed/power are editable too. The label engraving has to be "
                      "readable for the test to be useful, so don't leave it at zero \u2014 the "
                      "\"Apply material defaults\" button fills it from the material's existing "
                      "engraving settings."))
        self._bullet(_("\"Air assist\" toggles air on/off for every disc. \"Also test with air off\" "
                      "doubles the matrix so half the discs run with air and half without, "
                      "side by side."))
        self._bullet(_("\"Preview before saving\" pops a layout view so you can sanity-check disc "
                      "count and arrangement before committing to the file dialog."))
        self._bullet(_("G-code only \u2014 no SVG output, since the per-disc speed/power/passes "
                      "would be lost on import. The legend.txt is the source of truth for "
                      "looking up which disc had which settings."))
        self._blank()

        self._h2(_("Tooling \u2014 Pad Press Spacers"))
        self._body(_("3D-printable spacer biscuits for pad pressing. 1.75\" (44.45 mm) square "
                    "biscuits with their thickness engraved on top \u2014 stack them to set pad "
                    "press depth. PLA or PETG both work."))
        self._bullet(_("Half-step set: 3.0 / 3.5 / 4.0 / 4.5 mm (4 of each, 16 total)."))
        self._bullet(_("Quarter-step set: 3.25 / 3.75 / 4.25 mm (4 of each, 12 total)."))
        self._bullet(_("Organizer rack: 7 compartments to hold the whole spacer set."))
        self._body(_("Save the .stl files to disk and slice them in your 3D printer software."))
        self._blank()

        self._h2(_("Tooling \u2014 Settings"))
        self._body(_("Options > Tooling Settings: per-material G-code (acrylic, basswood) "
                    "and die-engraving options."))
        self._bullet(_("Acrylic / basswood G-code: speed, power, passes, kerf, air assist. "
                      "Defaults tuned for the Falcon2 Pro 40W (3 mm acrylic; 3 mm basswood). "
                      "Adjust to your machine."))
        self._bullet(_("Basswood preset feeds the camera-calibration card engrave AND the die "
                      "organizer (when you cut it in LightBurn)."))
        self._bullet(_("Die engraving: filled vs. line mode, font sizes for ring + cutout "
                      "engraving, ring engraving placement."))
        self._blank()

    def _section_tuner(self):
        self._h2(_("Tuner"))
        self._body(_("A 12-wheel chromatic stroboscopic tuner. "
                    "Each wheel shows concentric rings (one per octave) of alternating "
                    "colored and dark segments visible through a wedge-shaped cutout. "
                    "An analog VU meter at the bottom shows the detected fundamental pitch and cents error."))
        self._bullet(_("When the input pitch matches the reference, the pattern freezes (appears stationary)"))
        self._bullet(_("Sharp: pattern drifts right. Flat: pattern drifts left"))
        self._bullet(_("Faster drift = farther from in-tune. Frozen = perfectly in tune"))
        self._bullet(_("Multiple wheels respond simultaneously from harmonics in the sound \u2014 "
                      "this is real FFT analysis of the audio, not simulated"))
        self._bullet(_("Per-pitch-class phase tracking with temporal smoothing keeps each "
                      "wheel's rotation independent and stable"))
        self._blank()

        self._body(_("Rendering:"))
        self._bullet(_("GPU-accelerated via Rust/wgpu when available \u2014 60 to 120 fps"))
        self._bullet(_("Automatic CPU canvas fallback if the GPU path is unavailable "
                      "(older systems, virtual machines, etc.)"))
        self._blank()

        self._body(_("On-screen controls (the slider panel below the wheels):"))
        self._bullet(_("DISP > SENS: how loud a signal needs to be to register"))
        self._bullet(_("DISP > BRIGHT: master brightness for the strobe disc segments"))
        self._bullet(_("DISP > FPS: target frame rate (60 / 90 / 120)"))
        self._bullet(_("PITCH > A=: reference pitch in Hz (default 440)"))
        self._bullet(_("PITCH > KEY: instrument transposition (C / Bb / Eb / F)"))
        self._bullet(_("BIAS > NOTE: per-wheel ring brightness contrast "
                      "(0 = all rings same brightness, 100 = played octave ring brightest)"))
        self._bullet(_("BIAS > OCT.: dominant octave boost (emphasizes the strongest ring)"))
        self._blank()

        self._body(_("Settings (Options > Settings):"))
        self._bullet(_("Input Device: pick the microphone the tuner listens to"))
        self._bullet(_("Backlight Color: color of the strobe disc segments"))
        self._bullet(_("Faceplate Color: background color of the tuner display"))
        self._bullet(_("Show frame rate on screen: overlays a small live FPS counter for diagnostics"))
        self._blank()

        self._h2(_("Microphone"))
        self._body(_("The tuner analyzes audio from your microphone. A quality mic "
                    "makes a significant difference in accuracy, especially for "
                    "low notes where the fundamental frequency may be weak."))
        self._bullet(_("The mic needs to capture the full frequency range of the "
                      "saxophone (down to ~100 Hz for baritone). This requires a "
                      "sample rate of at least 44.1 kHz and a flat low-frequency response."))
        self._bullet(_("Recommended: Audio-Technica AT2020 USB (no audio interface needed)"))
        self._bullet(_("Any condenser mic through an audio interface will also work well"))
        self._bullet(_("Laptop/built-in mics work for basic tuning but may struggle "
                      "with low register notes"))
        self._bullet(_("Bluetooth headset mics do not work \u2014 their sample rate "
                      "is too low (16 kHz) for harmonic analysis"))
        self._bullet(_("Select your mic via Options > Input Device"))
        self._blank()

        self._body(_("The tuner activates automatically when you switch to the Tuner tab "
                    "and stops when you leave it, so there is no CPU or audio usage when "
                    "you are on other tabs."))
        self._blank()

    def _section_toner(self):
        self._h2(_("Toner \u2014 Tone Analyzer"))

        # === WHAT THIS TOOL IS ===
        self._h2(_("What This Tool Does"))
        self._body(_("The Toner is a harmonic analyzer. It detects your fundamental pitch and "
                    "measures the strength of each harmonic up to the 20th. Every reading "
                    "captures the whole chain at once \u2014 you + horn + mouthpiece + reed + "
                    "mic + room \u2014 so a single reading can't isolate any one part. The "
                    "value comes from changing one variable at a time and watching the delta."))
        self._blank()

        self._body(_("This is a tool for relative measurements, not absolute truth about gear. "
                    "In our data, mouthpiece + player + mic routinely dwarf horn-to-horn "
                    "differences \u2014 the same Conn 6M played by two different players on "
                    "two different mouthpieces produced 12 dB of upper-harmonic spread, more "
                    "than the spread across many different horns combined. Compare YOUR setups "
                    "to YOUR other setups, and expect any single reading to be mostly about "
                    "your current chain rather than the horn alone."))
        self._blank()

        # === TWO WAYS TO USE IT ===
        self._h2(_("Two Ways to Use It"))

        self._body(_("1. Live biofeedback while practicing"))
        self._bullet(_("The gauges and spectrum respond in real time. The "
                      "movement is the information \u2014 it shows how your "
                      "air, embouchure, and voicing choices affect the "
                      "sound moment to moment."))
        self._blank()

        self._body(_("2. Tracking changes over time"))
        self._bullet(_("Record a session, change one thing, record another. The comparison tool "
                      "shows what moved and by how much. Works for any variable:"))
        self._bullet(_("Switching mouthpieces \u2014 same horn, reed, mic. Delta is the mouthpiece."))
        self._bullet(_("Day-to-day variation \u2014 same everything. Delta is you."))
        self._bullet(_("Ribbon vs condenser \u2014 same horn, same room. Delta is the recording chain."))
        self._bullet(_("Two horns \u2014 same player, mouthpiece, mic, room. Delta is the horn."))
        self._blank()
        self._body(_("Over many sessions, what stays the same emerges from what drifts. That takes "
                    "time and discipline about what you control. Be skeptical of strong "
                    "conclusions from small samples."))
        self._blank()

        # === GETTING STARTED ===
        self._h2(_("Getting Started"))

        self._body(_("Before your first capture, you need two things: a "
                    "microphone and a preset."))
        self._blank()

        self._body(_("Microphone"))
        self._bullet(_("Mic type and model are set per preset \u2014 you'll "
                      "enter them when creating a preset in File \u2192 "
                      "Presets. Both are required before capturing."))
        self._bullet(_("A condenser mic (e.g. Audio-Technica AT2020 USB) "
                      "gives you the fullest picture \u2014 flat response "
                      "captures upper harmonics accurately, which is "
                      "where most of the interesting differences live."))
        self._bullet(_("Ribbon and dynamic mics can still be used, but they "
                      "attenuate upper harmonics, which skews complexity-based "
                      "comparisons. The mic type is stored with your data so "
                      "you always know what produced it, and the Analyze tool "
                      "warns when comparing across mic types."))
        self._bullet(_("Laptop/built-in mics are not suitable \u2014 they "
                      "roll off both low and high frequencies and add noise"))
        self._bullet(_("Bluetooth headset mics do not work \u2014 sample "
                      "rate is too low (16 kHz) for harmonic analysis"))
        self._blank()

        self._body(_("Mic placement"))
        self._bullet(_("2\u20133 feet (60\u201390 cm) from the bell, slightly off-axis."))
        self._bullet(_("Closer than 1 foot \u2014 proximity effect exaggerates low harmonics."))
        self._bullet(_("Quieter room is better \u2014 background noise masks upper harmonics."))
        self._bullet(_("Mic position matters more than almost anything else. Moving a few inches "
                      "between takes \u2014 same horn, same player, minutes apart \u2014 changed rolloff "
                      "by over 1 dB/harmonic in our tests, enough to shift complexity by 10\u201320%."))
        self._bullet(_("For controlled comparisons, use a mic stand at a fixed position; the "
                      "position change otherwise shows up in your data and can be larger than "
                      "what you're trying to measure. (Or change position deliberately to study "
                      "how much it influences capture.)"))
        self._blank()

        self._body(_("Preset"))
        self._bullet(_("A preset saves your setup details: horn + player + "
                      "mouthpiece + mic. It pre-fills session metadata for "
                      "quick capture start."))
        self._bullet(_("Click Capture, then create or load a preset. "
                      "Six fields are required: Make, Model, Player, "
                      "Mouthpiece, Mic Type, and Mic Model. Optional "
                      "fields like reed, serial, ligature, room, and "
                      "preamp can be enabled in Options \u2192 Settings "
                      "\u2192 Analysis tab."))
        self._bullet(_("The SAX selector sets the transposition and is "
                      "stored with the preset. When you load a preset, "
                      "it updates automatically."))
        self._blank()

        # === CAPTURING ===
        self._h2(_("Capturing"))
        self._bullet(_("Click Capture, select or create a preset, then play."))
        self._bullet(_("The tool auto-detects steady tones: hold a note "
                      "for about a second and it triggers automatically. "
                      "No button-pressing while playing."))
        self._bullet(_("Move to the next note and it captures again. Play "
                      "through the horn's range at your own pace."))
        self._bullet(_("The first ~100ms of each note is automatically "
                      "skipped \u2014 the attack transient doesn't represent "
                      "the sustained tone."))
        self._bullet(_("Capture at least 8 unique notes for a useful "
                      "fingerprint. More is better."))
        self._bullet(_("Click Stop when done. A coverage summary shows "
                      "which notes you hit and where the gaps are."))
        self._bullet(_("Every capture is timestamped. The session date "
                      "is saved automatically, so you can track changes "
                      "over time."))
        self._blank()

        # === THE DISPLAY ===
        self._h2(_("Display"))
        self._bullet(_("Spectrum view: full FFT frequency spectrum with "
                      "harmonics highlighted in amber and the fundamental "
                      "in green"))
        self._bullet(_("Bars view: one bar per harmonic, clean and simple"))
        self._bullet(_("Scale: Linear (default) shows true amplitude ratios. "
                      "dB shows the logarithmic audio scale."))
        self._blank()

        self._h2(_("Live Display"))
        self._body(_("The live display shows two things in real time: the "
                    "intonation gauge (flat/sharp with cents readout, IN TUNE "
                    "lamp within \u00b14 cents) and the spectrum bars "
                    "(complete harmonic shape from H1 through H20)."))
        self._blank()
        self._body(_("Live descriptor gauges (Pure\u2194Complex, Thin\u2194Warm) "
                    "were removed because absolute single-preset readouts "
                    "proved too noisy to trust \u2014 mic position alone "
                    "shifts complexity by 10\u201320% between takes, and "
                    "the mouthpiece dominates the signal much more than the "
                    "horn does. The descriptors still live in the Analyze "
                    "tool, where comparing two presets cancels those "
                    "confounders out."))

        self._h2(_("Spectrum Overlay"))
        self._body(_("You can load a preset as a ghost overlay on the "
                    "live spectrum. From the Analyze tool, view a single "
                    "preset or a group average, then click \"Overlay on "
                    "Spectrum.\" Blue ghost bars appear behind the live "
                    "display, updating per-note as you play."))
        self._blank()
        self._body(_("This lets you eyeball how your live sound compares "
                    "to a stored reference \u2014 useful for quick A/B "
                    "checks. For rigorous comparison, use the Analyze "
                    "tool where both sides are averaged data."))
        self._blank()

        # === ANALYZING ===
        self._h2(_("The Analyze Tool"))
        self._body(_("This is where the tool earns its keep. The Analyze tool "
                    "(File \u2192 Analyze\u2026) shows you what's different "
                    "between two or more presets \u2014 not which one is "
                    "\"better,\" but what changed and where."))
        self._blank()
        self._bullet(_("File \u2192 Analyze opens a picker with filters for horn "
                      "type, player, mouthpiece, and mic type"))
        self._bullet(_("Select one preset: detail view with descriptors and "
                      "harmonic curve"))
        self._bullet(_("Select two presets: side-by-side delta analysis with a "
                      "Difference chart \u2014 a single curve of the "
                      "harmonic-by-harmonic delta that instantly shows where "
                      "the sound diverges"))
        self._bullet(_("Select three or more: spread analysis across the group"))
        self._bullet(_("Toggle between Horn Average and Per-Note to see whether "
                      "differences are across the board or concentrated in "
                      "certain registers"))
        self._bullet(_("Population percentiles: each preset's descriptors are "
                      "ranked against all other presets of the same sax type, "
                      "with low/below avg/mid-range/above avg/high labels"))
        self._bullet(_("\"Overlay on Spectrum\" loads a preset as a blue ghost "
                      "behind the live display for real-time A/B comparison "
                      "while playing"))
        self._bullet(_("Clickable legend labels and chart lines show preset "
                      "details (player, mouthpiece, mic, etc.)"))
        self._bullet(_("Back button on analysis windows returns you to the picker "
                      "to try a different selection"))
        self._blank()
        self._body(_("The table shows mic type and recording quality (rolloff "
                    "rate) for each preset. If mic types differ, the analysis "
                    "notes that harmonic differences may partly reflect the mic "
                    "rather than the horn."))
        self._blank()

        # === PRESET MANAGEMENT EXTRAS ===
        self._h2(_("Preset Management Extras"))
        self._bullet(_("Mutate Preset: duplicates a preset with all fields "
                      "pre-filled (name cleared) so you can change one variable "
                      "(mouthpiece, mic, reed) and save it as a new preset. "
                      "Streamlines A/B testing workflows."))
        self._bullet(_("Sandbox mode: enable in Options \u2192 Settings \u2192 "
                      "Analysis to relax field requirements for non-sax "
                      "instruments, contact mics, effects chains, or "
                      "experimental setups. Sandbox presets are flagged with an "
                      "amber label so you don't confuse them with horn data."))
        self._bullet(_("Mic position field: an optional preset field for "
                      "documenting where the mic was placed (distance, angle, "
                      "axis). Highly recommended \u2014 mic position is one of "
                      "the biggest sources of variation in captures."))
        self._blank()

        # === TONE DESCRIPTORS ===
        self._h2(_("Tone Descriptors"))
        self._body(_("The Analyze tool computes a handful of descriptors from the "
                    "harmonic data. They are most useful for side-by-side comparison, "
                    "where deltas between two presets cancel out mic-position and "
                    "room confounders. Which ones are shown is configurable in "
                    "Options \u2192 Settings \u2192 Analysis tab \u2014 Even/Odd is "
                    "on by default; Rolloff Shape is off by default."))
        self._blank()

        self._body(_("Warmth and brightness are independent"))
        self._body(_("The two main descriptors \u2014 Warmth and Harmonic "
                    "Complexity \u2014 measure independent properties of "
                    "saxophone tone, and they don't have to agree. A horn or "
                    "mouthpiece can be warm and bright, warm and dark, not warm "
                    "and bright, or not warm and dark. All four are real, "
                    "recognizable characters that players name and care about."))
        self._blank()

        self._body(_("The four quadrants:"))
        self._bullet(_("Warm + bright: fat, supported, projecting. Vintage Conns "
                      "played hard, some metal mouthpieces, the classic \"big\" "
                      "tenor sound."))
        self._bullet(_("Warm + dark: round, mellow, rich. Classical large-chamber "
                      "mouthpieces, dark Otto Links."))
        self._bullet(_("Not warm + bright: edgy, cutting, no body. Berg Larsen "
                      "115/0, screaming fusion mouthpieces."))
        self._bullet(_("Not warm + dark: hollow, woody, clarinet-like. The \"pure "
                      "fundamental, suppressed everything else\" character. Some "
                      "vintage large-chamber hard rubber mouthpieces sit here."))
        self._blank()

        self._body(_("Why they're independent: Warmth measures how strongly the "
                    "2nd harmonic (the octave above the fundamental) sits in "
                    "the tone \u2014 that comes from how the chamber and reed "
                    "couple to the fundamental. Brightness measures how much "
                    "energy is in the upper harmonics \u2014 that comes from "
                    "the buzz of the reed against the tip rail. Those are "
                    "physically decoupled mechanisms. You can engineer either "
                    "axis without affecting the other."))
        self._blank()

        self._body(_("The Analyze tool's Character Map plots Warmth against a "
                    "brightness measure so you can see where any selected "
                    "preset sits in the two-dimensional character space. The "
                    "default brightness axis is Harmonic Complexity, because "
                    "it is the most independent of Warmth in our test data. "
                    "The dropdown also offers H4\u2013H5 mean (the raw average "
                    "strength of the 4th and 5th harmonics) and Rolloff "
                    "(inverted), but those measures share a denominator with "
                    "Warmth (both are dB relative to the fundamental) and "
                    "tend to track Warmth rather than provide an independent "
                    "axis."))
        self._blank()

        self._body(_("Character Map is most reliable within a single sax type. Lower-pitched horns "
                    "read warmer and brighter regardless of mouthpiece, so cross-type comparisons "
                    "mostly show you that physics rather than character. The tool warns before "
                    "opening one."))
        self._blank()

        self._body(_("The descriptors in detail"))
        self._bullet(_("Warmth (Thin \u2194 Warm) \u2014 strength of the 2nd "
                      "harmonic relative to the fundamental. A strong H2 "
                      "produces a round body sitting just above the fundamental, "
                      "which is the \"fat\" or \"thick\" quality players "
                      "describe as warmth. Note: warmth is NOT the same as "
                      "darkness. A mouthpiece can be dark in its upper "
                      "harmonics while still having a weak H2 (reads \"not "
                      "warm\") or a strong H2 (reads \"warm\")."))
        self._bullet(_("Some setups \u2014 metal mouthpieces driven hard especially \u2014 produce H2 LOUDER "
                      "than the fundamental, a \"horn-like\" spectrum where the octave dominates. "
                      "Real and recognized. Warmth reads these as moderately warm rather than "
                      "maxed out (perceptually, the H2 \u2248 H1 saturation point); we're still "
                      "working out the best characterization, so this part of the formula may "
                      "evolve."))
        self._blank()

        self._bullet(_("Harmonic Complexity (Pure \u2194 Complex) \u2014 spectral "
                      "flatness, a measure of how evenly distributed the "
                      "harmonic energy is. A pure tone with one dominant "
                      "harmonic reads low; a tone with many strong harmonics "
                      "of comparable amplitude reads high. Tracks the simple-"
                      "vs-rich axis of timbre."))
        self._blank()

        self._bullet(_("Even/Odd Ratio (Odd \u2194 Even) \u2014 the balance "
                      "between even harmonics (H2, H4, H6...) and odd harmonics "
                      "(H3, H5, H7...). Even harmonics produce a round, warm "
                      "quality; odd harmonics produce an edgier, hollower "
                      "quality. Conical bore instruments like saxophone "
                      "produce both, but the ratio varies between horns and "
                      "especially between mouthpieces. Captures something "
                      "Warmth alone doesn't \u2014 the full balance across "
                      "the entire harmonic series, not just H2. In our test "
                      "data, this is the most consistent descriptor across "
                      "varying recording conditions."))
        self._blank()

        self._bullet(_("Rolloff Shape (Smooth \u2194 Peaked) \u2014 how cleanly "
                      "the harmonics roll off from strong to weak. A low "
                      "value means a clean, linear rolloff; a high value "
                      "means bumps or peaks where certain harmonics stick out. "
                      "May correspond to what players describe as \"presence\" "
                      "or \"projection.\" Off by default; turn on in "
                      "Options \u2192 Settings \u2192 Analysis."))
        self._blank()

        self._bullet(_("Evenness (Variable \u2194 Even) \u2014 how uniformly "
                      "the tone character sits across the register. Low "
                      "evenness means the horn sounds different in different "
                      "ranges (e.g. warm in the low end and bright on top). "
                      "High evenness means consistent character throughout. "
                      "Computed only when at least 5 notes are captured."))
        self._blank()

        # === WHERE DIFFERENCES COME FROM ===
        self._h2(_("Where Differences Come From"))
        self._body(_("The analysis text flags which harmonic range moved most. Different parts of "
                    "the saxophone tend to affect different harmonics, though the mapping isn't "
                    "perfectly understood:"))
        self._blank()
        self._bullet(_("H1\u2013H4 (low): research suggests bore-driven more than mouthpiece-driven; "
                      "this range shows the least variation in our player/mouthpiece swaps."))
        self._bullet(_("H7\u2013H13 (upper): where neck swaps show their biggest effect in our data, "
                      "and where mouthpiece changes also show up."))
        self._bullet(_("H3\u2013H12 broadband: mouthpiece and player changes tend to lift or suppress "
                      "the whole upper series rather than a narrow band."))
        self._blank()
        self._body(_("These are observed patterns, not laws of physics. Same-player comparisons "
                    "isolate equipment; cross-player comparisons can't cleanly separate horn from "
                    "player."))
        self._blank()

        self._body(_("Measurement noise: how big does a delta need to be?"))
        self._bullet(_("In our early data, two takes of the same setup ten minutes apart varied by "
                      "1\u20133% on descriptors and ~0.3 dB/H on rolloff. Deltas under that floor are "
                      "likely session-to-session noise; bigger ones are potentially real but "
                      "still need context (same player? same mic? same room?) to interpret. Early "
                      "estimate from one two-take data point \u2014 will tighten as more repeat data "
                      "arrives."))
        self._blank()

        # === REPORTS ===
        self._h2(_("Preset Reports"))
        self._body(_("The report view shows one preset's history: "
                    "descriptors, session-by-session changes (with "
                    "deltas from the previous session), harmonic "
                    "curve, and per-note breakdown. Use it to see "
                    "how your readings on a given setup have drifted "
                    "over time."))
        self._blank()

        # === RECORDING QUALITY ===
        self._h2(_("Recording Quality"))
        self._body(_("The app measures harmonic rolloff \u2014 how quickly "
                    "upper harmonics fade relative to the fundamental. "
                    "This tells you how much of the signal your "
                    "recording setup is capturing:"))
        self._bullet(_("1.0\u20132.0 dB/H: Condenser mic, close placement. "
                      "Full harmonic detail."))
        self._bullet(_("2.0\u20132.5 dB/H: Good mic, further away or "
                      "reflective room."))
        self._bullet(_("2.5+ dB/H: Ribbon, dynamic, built-in, or distant "
                      "mic. Upper harmonics attenuated. The app warns "
                      "you after the first few captures if this is "
                      "detected."))
        self._blank()
        self._body(_("If you want to learn the most about your actual "
                    "setup \u2014 what the horn and mouthpiece are doing "
                    "\u2014 use a condenser mic. If you want to learn "
                    "how your recording chain colors the sound, try "
                    "different mics and compare. Both are valid uses."))
        self._blank()

        # === DATA MANAGEMENT ===
        self._h2(_("Data & Transfer"))
        self._body(_("Presets and sessions are saved automatically. Each "
                    "session records the date, setup details, mic type, "
                    "and mic model alongside the harmonic data."))
        self._blank()
        self._bullet(_("File \u2192 Transfer Data \u2192 Export/Import "
                      "Preset Library: for backup or moving data "
                      "between machines. Exported files are JSON."))
        self._blank()

        self._h2(_("WAV Recording"))
        self._body(_("WAV recording is on by default. Each session writes a WAV alongside the "
                    "harmonic data; first capture asks you to choose a folder."))
        self._blank()
        self._body(_("Why it matters: at the end of each session, the toner re-analyzes the "
                    "recording offline with stricter segment detection than the live pipeline "
                    "manages. Offline extracts about 2\u00d7 the harmonic resolution, so your stored "
                    "fingerprints come from the better measurement (~5 seconds for a 4-minute "
                    "recording). Disable in Options \u2192 Settings \u2192 General if you must (warning "
                    "explains the tradeoff); you can also auto-delete the WAV after analysis."))
        self._blank()
        self._body(_("Other reasons to keep the WAVs:"))
        self._bullet(_("Re-analyze with future, better tools."))
        self._bullet(_("Share recordings for others to analyze or import."))
        self._bullet(_("Listen back and correlate what you hear with what the data shows."))
        self._bullet(_("Keep a record of how a horn sounded on a given day."))
        self._blank()
        self._body(_("Files are named with preset name + session date. The toner activates "
                    "automatically on tab switch and stops when you leave."))
        self._blank()

        # === NERD INFO ===
        self._h2(_("Nerd Info \u2014 How It Works Under the Hood"))

        self._body(_("FFT Pipeline"))
        self._bullet(_("16,384-sample FFT at 44.1 kHz \u2192 2.69 Hz resolution, fine enough for "
                      "individual harmonics on the lowest baritone notes."))
        self._bullet(_("Hann window before the FFT (\u221231.6 dB sidelobes), standard for harmonic "
                      "analysis."))
        self._blank()

        self._body(_("Pitch Detection"))
        self._bullet(_("Peak-picking with sub-harmonic verification: on sax the 2nd or 3rd "
                      "harmonic is often louder than the fundamental, so the detector checks "
                      "whether a strong peak might be H2/H3/H4/H5 of a lower note by looking "
                      "for that candidate's own series. Prevents octave errors."))
        self._bullet(_("Temporal hysteresis prevents frame-to-frame octave jumps."))
        self._blank()

        self._body(_("Harmonic Measurement"))
        self._bullet(_("For each harmonic up to the 20th: peak-pick within \u00b13 bins, then parabolic "
                      "interpolation (CCRMA method) refines frequency and amplitude. Corrects up "
                      "to 1.42 dB of inter-bin scalloping loss."))
        self._bullet(_("Measured in dB relative to the fundamental \u2014 normalizes out volume and "
                      "mic gain."))
        self._bullet(_("Harmonics under \u221260 dB are discarded as noise."))
        self._blank()

        self._body(_("Descriptors \u2014 Formulas"))
        self._bullet(_("Harmonic Complexity uses spectral flatness (the ratio "
                      "of geometric mean to arithmetic mean of harmonic "
                      "amplitudes), scaled by the coverage of significant "
                      "harmonics (those above \u221235 dB relative to the "
                      "fundamental). A pure tone has low flatness; a complex "
                      "tone with many strong harmonics has high flatness."))
        self._bullet(_("Warmth measures the strength of the 2nd harmonic "
                      "(the octave) in dB relative to the fundamental, mapped "
                      "to a 0\u20131 range. A stronger H2 reads warmer. "
                      "Consistent with acoustic research on even/odd harmonic "
                      "ratios in wind instruments."))
        self._bullet(_("Even/Odd Ratio is computed in the linear (not dB) "
                      "amplitude domain as the sum of even harmonic amplitudes "
                      "divided by the sum of all harmonic amplitudes. Higher "
                      "values indicate even-harmonic dominance."))
        self._bullet(_("Rolloff Shape is the standard deviation of residuals "
                      "from a linear fit to the harmonic series in dB. Low "
                      "values mean a clean exponential rolloff; high values "
                      "mean peaks or notches in the harmonic structure."))
        self._bullet(_("Evenness is computed as 1 minus (standard deviation of "
                      "complexity across notes / 0.40), clamped to the 0\u20131 "
                      "range. Requires at least 5 notes captured."))
        self._bullet(_("Rolloff Rate is the slope (dB per harmonic) of the "
                      "linear fit to H1\u2013H12 in dB. Used as both a "
                      "recording-quality indicator and a darkness proxy. "
                      "Sax-type-dependent: typical baritone 1.5\u20132.0, "
                      "tenor 1.1\u20132.0, alto 1.9\u20132.3, soprano 2.8\u20134.0."))
        self._blank()

        self._body(_("Descriptors are never stored \u2014 they are recomputed from raw harmonic data "
                    "using the current formulas, so improvements apply retroactively to every "
                    "historical capture."))
        self._blank()

        self._body(_("What Gets Saved"))
        self._bullet(_("Per capture: raw harmonic amplitudes (dB rel. fundamental), harmonic cents "
                      "deviations, fundamental frequency, spectral centroid, signal level. "
                      "Everything else is derived."))
        self._bullet(_("Captures are averaged per-note first, then across notes with equal weight. "
                      "Prevents register skew \u2014 20 high-note captures and 3 low-note captures "
                      "still represent the whole horn evenly."))
        self._bullet(_("The first ~100 ms of each note is skipped \u2014 attack transients contain "
                      "broadband non-harmonic energy that doesn't represent sustained tone "
                      "(Saldanha & Corso 1964)."))
        self._blank()

        self._body(_("Why Raw Data Matters"))
        self._bullet(_("Spectral centroid is stored as future-proofing \u2014 the most validated "
                      "acoustic correlate of perceived brightness in the literature."))
        self._bullet(_("Harmonics are stored up to H20 even though current gauges use only the "
                      "first few; low-register sax can produce 20+ audible harmonics."))
        self._bullet(_("Today's captures will be fully usable by better analysis tools tomorrow."))
        self._blank()

    def _section_import_export(self):
        self._h2(_("Import / Export & Sharing"))
        self._body(_("Each data tab supports importing and exporting so you can share "
                    "data with colleagues or back up your measurements:"))
        self._bullet(_("Pad Presets: File > Export/Import Pad Presets to share saved pad size lists"))
        self._bullet(_("Key Heights: File > Export/Import Key Heights to share measurement sets"))
        self._bullet(_("Screw Specs: File > Export/Import Screw Specs to share thread data"))
        self._bullet(_("Tone Presets: File > Transfer Data > Export/Import Preset Library"))
        self._bullet(_("All Settings: File > Import Settings from Folder copies all config "
                      "files from another installation"))
        self._body(_("Exported files are standard JSON and can be emailed or shared via any method."))
        self._blank()

    def _section_padmaking_guide(self):
        self._h2(_("Learn to Make Pads"))
        self._body(_("If you somehow got this program and missed the guide that started it all, "
                    "here's the complete how-to on making saxophone pads:"))
        self._link(_("stohrermusic.com/articles/how-to-make-saxophone-pads/"),
                    "https://www.stohrermusic.com/articles/how-to-make-saxophone-pads/")


class AboutDialog(tk.Toplevel):
    """About dialog with feature summary and contact info."""

    def __init__(self, parent):
        import webbrowser
        super().__init__(parent)
        self.title(_("About"))
        self.geometry("440x520")
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        tk.Label(self, text=_("Stohrer Sax Shop Companion"), bg=DIALOG_BG,
                 font=("Helvetica", 14, "bold")).pack(pady=(18, 3))
        tk.Label(self, text=_("by Matt Stohrer"), bg=DIALOG_BG,
                 font=("Helvetica", 11)).pack()
        tk.Label(self, text=_("~Salingeresque Sax Repair Techno-poet~"), bg=DIALOG_BG,
                 font=("Helvetica", 9, "italic")).pack()
        tk.Label(self, text=_("Version {version}  •  Built {build}").format(version=APP_VERSION, build=APP_BUILD_DATE),
                 bg=DIALOG_BG, font=("Helvetica", 9)).pack(pady=(2, 0))

        # Feature summary
        features = _(
            "\u2022  SVG & G-code generation for laser-cut pads\n"
            "\u2022  Key height library with import/export\n"
            "\u2022  Serial number lookup\n"
            "\u2022  OEM screw & rod specifications\n"
            "\u2022  Die insert & die holder tooling\n"
            "\u2022  Chromatic strobe tuner\n"
            "\u2022  Harmonic tone analyzer"
        )
        tk.Label(self, text=features, bg=DIALOG_BG,
                 font=("Helvetica", 10), justify="left",
                 anchor="w").pack(padx=40, pady=(10, 0), fill="x")

        # Separator
        sep = tk.Frame(self, height=1, bg="#CCCCCC")
        sep.pack(fill="x", padx=30, pady=10)

        # Contact
        tk.Label(self, text=_("Questions or feedback:"),
                 bg=DIALOG_BG, font=("Helvetica", 10)).pack()
        tk.Label(self, text="stohrermusic@gmail.com",
                 bg=DIALOG_BG, font=("Helvetica", 10, "bold")).pack()

        # Separator
        sep2 = tk.Frame(self, height=1, bg="#CCCCCC")
        sep2.pack(fill="x", padx=30, pady=10)

        # Donate section
        tk.Label(self, text=_("Like this app? Want to say thanks?"),
                 bg=DIALOG_BG, font=("Helvetica", 10)).pack()
        donate_text = _(
            "PayPal: stohrermusic@gmail.com\n"
            "Venmo: @matthew-stohrer"
        )
        tk.Label(self, text=donate_text, bg=DIALOG_BG,
                 font=("Helvetica", 10), justify="center").pack(pady=(4, 0))

        # YouTube link
        yt_frame = tk.Frame(self, bg=DIALOG_BG)
        yt_frame.pack(pady=(2, 0))
        tk.Label(yt_frame, text=_("Join my "), bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(side="left")
        yt_link = tk.Label(yt_frame, text=_("YouTube channel"),
                           bg=DIALOG_BG, fg="#0066CC",
                           font=("Helvetica", 10, "underline"),
                           cursor="hand2")
        yt_link.pack(side="left")
        yt_link.bind("<Button-1>",
                     lambda e: webbrowser.open("https://www.youtube.com/@StohrerMusic"))
        tk.Label(yt_frame, text=_(" as a paying member"), bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(side="left")

        tk.Label(self, text=_("Or email me for my address to send a six-pack,\n"
                 "some old tools, or whatever!"),
                 bg=DIALOG_BG, font=("Helvetica", 10),
                 justify="center").pack(pady=(4, 0))

        tk.Button(self, text=_("OK"), command=self.destroy, width=10).pack(pady=(12, 15))


class PadNotesWindow(tk.Toplevel):
    """Small modal dialog for viewing/editing notes on a pad preset."""

    def __init__(self, parent, preset_name, notes_text=""):
        super().__init__(parent)
        self.title(_("Preset Notes — {name}").format(name=preset_name))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.resizable(False, False)

        self.result = None  # Will be the new notes text if saved, None if cancelled

        tk.Label(self, text=_('Notes for "{name}":').format(name=preset_name), bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(padx=15, pady=(15, 5), anchor="w")

        self.notes_text = tk.Text(self, height=6, width=45, font=("Helvetica", 10), wrap="word")
        self.notes_text.pack(padx=15, pady=5)
        self.notes_text.insert("1.0", notes_text)

        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(5, 15))
        tk.Button(btn_frame, text=_("Save"), command=self.on_save, width=10).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Cancel"), command=self.on_cancel, width=10).pack(side="left", padx=5)

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
        self.notes_text.focus_set()
        self.wait_window(self)

    def on_save(self):
        self.result = self.notes_text.get("1.0", tk.END).strip()
        self.destroy()

    def on_cancel(self):
        self.result = None
        self.destroy()


# Material colors for preview
_PREVIEW_COLORS = {
    'felt': '#4488CC',
    'card': '#CC8844',
    'leather': '#886633',
    'exact_size': '#66AA66',
}


class NestingPreviewWindow(tk.Toplevel):
    """Preview window showing the nested pad layout before file generation.

    Shows the sheet boundary (or polygon) with circles at their nested
    positions. User can proceed to save or go back to adjust settings.

    Result:
        self.result = "save" if user clicks Save, None if Adjust/close.
    """

    def __init__(self, parent, placements, width_mm, height_mm, polygon=None):
        """
        Args:
            parent: Parent tk widget
            placements: dict of {material: [(pad_size, cx, cy, r), ...]}
            width_mm: Sheet width in mm
            height_mm: Sheet height in mm
            polygon: Optional list of (x, y) polygon points in mm
        """
        super().__init__(parent)
        self.title(_("Nesting Preview"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.result = None

        # First-run tutorial (persisted in settings via parent app)
        app = getattr(parent, 'settings', None) if not isinstance(parent, dict) else None
        if app is None:
            # Try to get settings from the root app
            try:
                app = parent.master.settings if hasattr(parent, 'master') else None
            except Exception:
                app = None

        # Use a simple attribute check — the actual persistence happens
        # via the seen_preview_tutorial setting
        show_tutorial = False
        try:
            if hasattr(parent, 'settings') and not parent.settings.get("seen_preview_tutorial"):
                show_tutorial = True
                parent.settings["seen_preview_tutorial"] = True
                save_settings(parent.settings)
        except Exception:
            pass

        if show_tutorial:
            messagebox.showinfo(_("Nesting Preview"),
                _("This preview shows how your pads will be arranged "
                "on the sheet before files are generated.\n\n"
                "If the layout looks good, click Save to write the files.\n\n"
                "If not, click Adjust to go back and change:\n"
                "  \u2022 Edge bias (pack toward a different edge)\n"
                "  \u2022 Sheet dimensions\n"
                "  \u2022 Custom polygon shape\n"
                "  \u2022 Pad sizes or quantities\n\n"
                "Preview works with one material at a time."),
                parent=self)

        self._placements = placements
        self._width_mm = width_mm
        self._height_mm = height_mm
        self._polygon = polygon
        self._materials = list(placements.keys())

        # Material selector (if multiple materials)
        if len(self._materials) > 1:
            sel_frame = tk.Frame(self, bg=DIALOG_BG)
            sel_frame.pack(pady=(10, 0))
            tk.Label(sel_frame, text=_("Material:"), bg=DIALOG_BG,
                     font=("Helvetica", 10)).pack(side="left", padx=(0, 5))
            self._mat_var = tk.StringVar(value=self._materials[0])
            mat_combo = ttk.Combobox(
                sel_frame, textvariable=self._mat_var,
                values=self._materials, state="readonly", width=12)
            mat_combo.pack(side="left")
            mat_combo.bind("<<ComboboxSelected>>", lambda e: self._draw())
        else:
            self._mat_var = tk.StringVar(value=self._materials[0])

        # Canvas
        self._canvas = tk.Canvas(self, bg="white", highlightthickness=1,
                                  highlightbackground="#999999")
        self._canvas.pack(fill="both", expand=True, padx=10, pady=(10, 5))
        self._canvas.bind("<Configure>", self._on_resize)

        # Info label (updated per material)
        self._info_label = tk.Label(self, text="", bg=DIALOG_BG,
                                     font=("Helvetica", 9))
        self._info_label.pack(pady=(0, 5))

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 10))
        tk.Button(btn_frame, text=_("Save Files"), command=self._on_save,
                  font=("Helvetica", 10, "bold"), width=12).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Adjust"), command=self._on_adjust,
                  font=("Helvetica", 10), width=12).pack(side="left", padx=5)

        # Size the window
        self.geometry("700x550")
        self.minsize(400, 300)

        self.protocol("WM_DELETE_WINDOW", self._on_adjust)
        self.wait_window(self)

    def _on_save(self):
        self.result = "save"
        self.destroy()

    def _on_adjust(self):
        self.result = None
        self.destroy()

    def _on_resize(self, event=None):
        self._draw()

    def _draw(self):
        cv = self._canvas
        cv.delete("all")

        cw = cv.winfo_width()
        ch = cv.winfo_height()
        if cw < 20 or ch < 20:
            return

        margin = 20
        draw_w = cw - 2 * margin
        draw_h = ch - 2 * margin

        # Scale to fit sheet in canvas, maintaining aspect ratio.
        # When a polygon is loaded the nest packs in the polygon's own
        # coordinate space (its bounding box), which may not match the
        # user's typed width/height fields — typically a polygon drawn
        # at e.g. 0..200mm with a 38x76 sheet typed in. Use the polygon's
        # actual bbox for canvas scaling so the contents land in view.
        if self._polygon:
            poly_xs = [p[0] for p in self._polygon]
            poly_ys = [p[1] for p in self._polygon]
            sheet_w = max(poly_xs) - min(poly_xs)
            sheet_h = max(poly_ys) - min(poly_ys)
        else:
            sheet_w = self._width_mm
            sheet_h = self._height_mm
        scale_x = draw_w / sheet_w if sheet_w > 0 else 1
        scale_y = draw_h / sheet_h if sheet_h > 0 else 1
        scale = min(scale_x, scale_y)

        # Center the sheet in the canvas
        offset_x = margin + (draw_w - sheet_w * scale) / 2
        offset_y = margin + (draw_h - sheet_h * scale) / 2

        # Draw sheet boundary
        if self._polygon:
            pts = []
            for px, py in self._polygon:
                pts.extend([offset_x + px * scale, offset_y + py * scale])
            cv.create_polygon(pts, fill="#F8F8F0", outline="#888888", width=2)
        else:
            cv.create_rectangle(
                offset_x, offset_y,
                offset_x + sheet_w * scale,
                offset_y + sheet_h * scale,
                fill="#F8F8F0", outline="#888888", width=2)

        # Draw pads for selected material only
        material = self._mat_var.get()
        placed = self._placements.get(material, [])
        color = _PREVIEW_COLORS.get(material, '#888888')
        fill = self._lighten(color, 0.7)

        for pad_size, cx, cy, r in placed:
            sx = offset_x + cx * scale
            sy = offset_y + cy * scale
            sr = r * scale

            cv.create_oval(sx - sr, sy - sr, sx + sr, sy + sr,
                           fill=fill, outline=color, width=1)

            # Label with pad size (skip if too small to read)
            if sr > 12:
                font_size = max(6, min(11, int(sr * 0.5)))
                cv.create_text(sx, sy, text=f"{pad_size:.1f}",
                               fill=color, font=("Helvetica", font_size))

        # Usage percentage for this material
        sheet_area = sheet_w * sheet_h
        if self._polygon:
            pts = self._polygon
            n = len(pts)
            area = 0
            for i in range(n):
                j = (i + 1) % n
                area += pts[i][0] * pts[j][1]
                area -= pts[j][0] * pts[i][1]
            sheet_area = abs(area) / 2

        pad_area = sum(3.14159 * r * r for _, _, _, r in placed)
        usage_pct = pad_area / sheet_area * 100 if sheet_area > 0 else 0

        cv.create_text(cw - margin, ch - 5,
                       text=_("{pct:.0f}% used").format(pct=usage_pct),
                       fill="#666666", font=("Helvetica", 9),
                       anchor="se")

        # Update info label
        self._info_label.configure(
            text=_("{count} pads ({material}) on {w:.0f} x {h:.0f} mm{shape}").format(count=len(placed), material=material, w=sheet_w, h=sheet_h, shape=_(' (custom shape)') if self._polygon else ''))

    @staticmethod
    def _lighten(hex_color, factor):
        """Lighten a hex color toward white."""
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        r = int(r + (255 - r) * factor)
        g = int(g + (255 - g) * factor)
        b = int(b + (255 - b) * factor)
        return f"#{r:02x}{g:02x}{b:02x}"


class SpeedPowerTestPreview(tk.Toplevel):
    """Preview window for the Tooling-tab Speed & Power Test.

    Mirrors NestingPreviewWindow's shape (modal canvas, Save/Adjust buttons)
    but renders the test-disc grid: each disc shows its 2-char numeric ID,
    and when "Also test with air off" is enabled the two air states are
    color-coded so the user can spot at a glance which discs are which.

    Result:
        self.result = "save" if user clicks Save, None if Adjust/close.
    """

    AIR_ON_COLOR = "#3a78c2"   # blue outline
    AIR_OFF_COLOR = "#cc7a3a"  # orange outline
    NEUTRAL_COLOR = "#555555"  # dark gray when air doesn't vary

    def __init__(self, parent, test_pieces, sheet_w_mm, sheet_h_mm,
                  eng_speed, eng_power, material_display):
        super().__init__(parent)
        self.title(_("Speed & Power Test Preview"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self._pieces = test_pieces
        self._sheet_w = sheet_w_mm
        self._sheet_h = sheet_h_mm
        self._eng_speed = eng_speed
        self._eng_power = eng_power
        self._material_display = material_display

        air_states = {p.get('air_assist', True) for p in test_pieces}
        self._air_varies = len(air_states) > 1

        # Summary line(s) at top
        sample = test_pieces[0] if test_pieces else None
        if sample is not None:
            speeds = sorted({p['speed'] for p in test_pieces})
            powers = sorted({p['power'] for p in test_pieces})
            passes_vals = sorted({p['passes'] for p in test_pieces})

            def _rng(vals):
                if len(vals) == 1:
                    return str(vals[0])
                return f"{vals[0]}-{vals[-1]}"

            summary_text = _("{n} discs - Speed {sp} mm/min, Power {pw}%, "
                              "Passes {ps}").format(
                n=len(test_pieces),
                sp=_rng(speeds),
                pw=_rng(powers),
                ps=_rng(passes_vals))
        else:
            summary_text = ""

        tk.Label(self, text=summary_text, bg=DIALOG_BG,
                 font=("Helvetica", 10, "bold")).pack(pady=(10, 2))

        eng_line = _("Material: {m}   Sheet: {w:.0f} x {h:.0f} mm   "
                      "Engraving: {es} mm/min @ {ep}%").format(
            m=material_display, w=sheet_w_mm, h=sheet_h_mm,
            es=int(eng_speed), ep=int(eng_power))
        tk.Label(self, text=eng_line, bg=DIALOG_BG,
                 font=("Helvetica", 9), fg="#555555").pack(pady=(0, 4))

        if self._air_varies:
            legend_frame = tk.Frame(self, bg=DIALOG_BG)
            legend_frame.pack(pady=(0, 4))
            tk.Label(legend_frame, text="●", bg=DIALOG_BG,
                     fg=self.AIR_ON_COLOR,
                     font=("Helvetica", 14)).pack(side="left")
            tk.Label(legend_frame, text=_("air ON  "), bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(side="left")
            tk.Label(legend_frame, text="●", bg=DIALOG_BG,
                     fg=self.AIR_OFF_COLOR,
                     font=("Helvetica", 14)).pack(side="left")
            tk.Label(legend_frame, text=_("air OFF"), bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(side="left")

        # Canvas
        self._canvas = tk.Canvas(self, bg="white", highlightthickness=1,
                                  highlightbackground="#999999")
        self._canvas.pack(fill="both", expand=True, padx=10, pady=(2, 5))
        self._canvas.bind("<Configure>", lambda e: self._draw())

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 10))
        tk.Button(btn_frame, text=_("Save G-code"), command=self._on_save,
                  font=("Helvetica", 10, "bold"), width=12).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Adjust"), command=self._on_adjust,
                  font=("Helvetica", 10), width=12).pack(side="left", padx=5)

        self.geometry("720x600")
        self.minsize(420, 320)
        self.protocol("WM_DELETE_WINDOW", self._on_adjust)
        self.wait_window(self)

    def _on_save(self):
        self.result = "save"
        self.destroy()

    def _on_adjust(self):
        self.result = None
        self.destroy()

    def _draw(self):
        cv = self._canvas
        cv.delete("all")

        cw = cv.winfo_width()
        ch = cv.winfo_height()
        if cw < 20 or ch < 20:
            return

        margin = 20
        draw_w = cw - 2 * margin
        draw_h = ch - 2 * margin
        scale_x = draw_w / self._sheet_w if self._sheet_w > 0 else 1
        scale_y = draw_h / self._sheet_h if self._sheet_h > 0 else 1
        scale = min(scale_x, scale_y)

        offset_x = margin + (draw_w - self._sheet_w * scale) / 2
        offset_y = margin + (draw_h - self._sheet_h * scale) / 2

        # Sheet outline
        cv.create_rectangle(
            offset_x, offset_y,
            offset_x + self._sheet_w * scale,
            offset_y + self._sheet_h * scale,
            fill="#F8F8F0", outline="#888888", width=2)

        for piece in self._pieces:
            cx = offset_x + piece['cx'] * scale
            cy = offset_y + piece['cy'] * scale
            sr = (piece['diameter'] / 2.0) * scale
            inner = piece.get('inner_diameter', 0) or 0
            if self._air_varies:
                color = (self.AIR_ON_COLOR
                          if piece.get('air_assist', True)
                          else self.AIR_OFF_COLOR)
            else:
                color = self.NEUTRAL_COLOR
            fill = _PREVIEW_COLORS.get('felt', '#dddddd')  # subtle light fill
            fill = self._lighten_hex(color, 0.85)
            cv.create_oval(cx - sr, cy - sr, cx + sr, cy + sr,
                           fill=fill, outline=color, width=1)
            # Center hole (washer / shim): punch it out of the fill in white
            # so the ring reads as solid material, matching the cut output.
            if inner > 0:
                ir = (inner / 2.0) * scale
                cv.create_oval(cx - ir, cy - ir, cx + ir, cy + ir,
                               fill="white", outline=color, width=1)
            # ID label — centered for a solid disc, dropped into the lower ring
            # for a washer (mirrors feeds_speeds_label_geometry / the G-code).
            label_dy_mm, _font_mm = feeds_speeds_label_geometry(
                piece['diameter'], inner)
            label_y = cy + label_dy_mm * scale
            if sr > 8:
                font_size = max(7, min(12, int(sr * 0.6)))
                if inner > 0:
                    # Keep the number inside the ring band on screen too.
                    ring_px = sr - (inner / 2.0) * scale
                    font_size = max(6, min(font_size, int(ring_px * 0.7)))
                cv.create_text(cx, label_y, text=piece['id'],
                               fill=color, font=("Helvetica", font_size, "bold"))

    @staticmethod
    def _lighten_hex(hex_color, factor):
        h = hex_color.lstrip('#')
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        r = int(r + (255 - r) * factor)
        g = int(g + (255 - g) * factor)
        b = int(b + (255 - b) * factor)
        return f"#{r:02x}{g:02x}{b:02x}"


class CameraCalibrationDialog(tk.Toplevel):
    """One-time camera calibration wizard.

    Shows the live camera feed at ~10 fps with detected ChArUco corners
    overlaid in green. User moves the calibration card to several
    positions on the bed; each frame where the board is detected gets
    captured. Once enough good frames are collected, the user clicks
    Calibrate; the dialog runs the calibration math and saves the result.

    Result:
        self.result = "saved" if calibration was successfully saved, else None.
    """

    LIVE_REFRESH_MS = 100   # ~10 fps preview
    TARGET_FRAMES = 12      # frames to collect for a robust calibration
    PREVIEW_W = 640
    PREVIEW_H = 480

    def __init__(self, parent, camera_index, calibration_path,
                  falcon_port=None, settings=None):
        super().__init__(parent)
        self.title(_("Camera Calibration"))
        self.configure(bg=DIALOG_BG)
        # NOT transient(parent) — on Windows that strips the maximize
        # button. The dialog has a lot of content (instructions +
        # preview + jog cluster + buttons) that benefits from being
        # maximizable on small screens.
        self.grab_set()
        self.resizable(True, True)
        self.result = None

        self._camera_index = camera_index
        self._calibration_path = calibration_path
        self._falcon_port = falcon_port
        self._settings = settings or {}
        self._cap = None
        self._board = None
        self._detector_imports_ok = False
        self._captures = []       # list of (charuco_corners, charuco_ids)
        self._last_frame_shape = None
        self._image_size = (640, 480)
        self._latest_photo = None  # keep a reference so PIL doesn't GC it
        self._refresh_after_id = None
        # Phase state: 'engrave' or 'capture'. Decided after the
        # persisted-offset check below — if SSC saved an offset from
        # a previous (possibly 2-hour) engrave, default to capture
        # phase so the user isn't asked to re-engrave a card that's
        # already on the bed.
        self._phase = None  # set below after offset restore
        # Capture-phase sub-state: 'reference' (card untouched, take
        # 1+ reference captures) → 'intrinsics' (card moved for
        # different poses). Tracked so calibrate_from_frames can pool
        # all reference captures for the machine-mm homography.
        self._capture_substate = 'reference'
        self._reference_count = 0
        # Has the laser been homed this dialog session? Required
        # before Engrave so MPos is meaningful — otherwise the
        # calibration's machine reference is whatever bogus value
        # Grbl booted with.
        self._engrave_homed = False
        # Linear-flow flags driving button enablement: Frame →
        # _framing_done → enables Engrave; Engrave → engrave offset
        # set → enables Next; Next → capture phase.
        self._framing_done = False
        # Center of the last framing pass, captured by the framing
        # provider on each iteration. Engrave uses THIS (not current
        # MPos) so Stopping mid-trace doesn't move the engrave's
        # center to wherever the head happened to be when Stop hit.
        self._framing_center = None

        # Where the card was engraved on the bed, in machine-mm.
        # Either captured this session (set by _on_engrave_clicked
        # after a successful engrave) or restored from a previously-
        # persisted value (below).
        self._captured_engrave_offset = None

        # If the previous engrave's offset was saved (persisted to
        # settings on engrave success), restore it. Then the dialog
        # opens straight in capture mode — the user shouldn't have
        # to remember a "Skip — already engraved" step if the system
        # already knows the card was engraved at a specific offset.
        saved = self._settings.get(
            'camera_calibration_engrave_offset_mm')
        if (saved and isinstance(saved, (list, tuple))
                and len(saved) == 2):
            try:
                self._captured_engrave_offset = (
                    float(saved[0]), float(saved[1]))
            except (TypeError, ValueError):
                pass

        # Phase pick: capture if we have a saved offset (most users
        # will, after their first engrave), else engrave. Without a
        # Falcon connected, can't engrave anyway → capture.
        if self._captured_engrave_offset is not None or not falcon_port:
            self._phase = 'capture'
        else:
            self._phase = 'engrave'

        # Block system/display sleep for the lifetime of the dialog —
        # the engrave can run 45-60 minutes and Windows would
        # otherwise sleep the system and kill the USB serial mid-job.
        # Released in _on_close.
        try:
            import sleep_lock
            sleep_lock.prevent_sleep()
        except Exception:
            pass

        # Try to import deps. If they're missing we'll show an error and close.
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV is not available")
            self._board = camera_capture.make_charuco_board()
            self._detector_imports_ok = True
        except ImportError as e:
            tk.Label(self, text=_("Camera calibration requires opencv-python "
                                    "and Pillow:\n\n    pip install opencv-python Pillow\n\n"
                                    "Error: {e}").format(e=e),
                     bg=DIALOG_BG, fg="#a30000", padx=20, pady=20).pack()
            tk.Button(self, text=_("Close"),
                      command=self.destroy).pack(pady=(0, 15))
            return

        # ---- UI ----
        # Horizontal split layout: title bar at top, status at
        # bottom, then a body with the camera preview on the LEFT
        # (expands to fill) and the controls (instructions + buttons
        # + engrave-jog cluster) on the RIGHT (fixed width). Makes
        # better use of widescreen monitors and keeps everything
        # visible without scrolling.
        tk.Label(self, text=_("Camera Calibration"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold")
                 ).pack(side='top', pady=(10, 2))

        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(side='bottom', pady=(0, 8))

        body = tk.Frame(self, bg=DIALOG_BG)
        body.pack(side='top', fill='both', expand=True,
                   padx=10, pady=(0, 8))

        # LEFT: camera preview, expands to fill available space.
        preview_frame = tk.Frame(body, bg='black')
        preview_frame.pack(side='left', fill='both', expand=True,
                            padx=(0, 8))
        self._preview_label = tk.Label(
            preview_frame, bg="#000000",
            width=self.PREVIEW_W, height=self.PREVIEW_H)
        self._preview_label.pack(fill='both', expand=True)

        # RIGHT: control panel (instructions on top, jog/buttons
        # below). Fixed width so the preview gets the rest.
        right_panel = tk.Frame(body, bg=DIALOG_BG, width=340)
        right_panel.pack(side='left', fill='y')
        right_panel.pack_propagate(False)  # honor the width=340

        # Instructions live in the right panel, narrower wraplength.
        self._instructions_var = tk.StringVar(value="")
        self._instructions_label = tk.Label(
            right_panel, textvariable=self._instructions_var,
            bg=DIALOG_BG, wraplength=325, justify="left",
            font=("Helvetica", 9))
        self._instructions_label.pack(side='top', anchor='w',
                                        padx=4, pady=(0, 8))

        # Bottom-row buttons in the right panel. Engrave jog cluster
        # also lives here (added by _enter_engrave_phase).
        self._btn_frame = tk.Frame(right_panel, bg=DIALOG_BG)
        self._btn_frame.pack(side='bottom', pady=(0, 8), fill='x')

        # Three-button linear flow: Frame → Engrave → Next. Each
        # enables the next as its prerequisite completes. Cancel is
        # always available.
        self._frame_btn = tk.Button(
            self._btn_frame, text=_("Frame"),
            command=self._on_frame_clicked,
            font=("Helvetica", 10), width=24)
        self._engrave_btn = tk.Button(
            self._btn_frame, text=_("Engrave"),
            command=self._on_engrave_clicked,
            font=("Helvetica", 10), width=24,
            state="disabled")  # enabled after framing done
        self._next_btn = tk.Button(
            self._btn_frame, text=_("Next →"),
            command=self._on_next_clicked,
            font=("Helvetica", 10, "bold"), width=24,
            bg="#4caf50", fg="white", activebackground="#388e3c",
            state="disabled")  # enabled after engrave done
        # "Engrave new card" — shown in CAPTURE phase to let the user
        # start over with a fresh piece of basswood (can't re-engrave
        # the same wood — each engrave is permanent, so a new card
        # means new material). Skip — already engraved button was
        # removed — it always used the (50,50) default which is wrong
        # for users who jogged before engraving.
        self._reengrave_btn = tk.Button(
            self._btn_frame, text=_("Engrave new card"),
            command=self._on_reengrave_clicked,
            font=("Helvetica", 9), width=16)

        # Capture-phase buttons. Label updates with substate.
        self._capture_btn = tk.Button(
            self._btn_frame, text=_("Capture reference"),
            command=self._on_capture_clicked,
            font=("Helvetica", 10, "bold"), width=20, state="disabled")
        # "Done with references — move card now" advances from the
        # reference substate (multiple captures of the untouched
        # card, pooled for the machine-mm homography) to the
        # intrinsics substate (card moved, used for distortion).
        self._done_refs_btn = tk.Button(
            self._btn_frame, text=_("Done with references"),
            command=self._on_done_refs_clicked,
            font=("Helvetica", 9), width=18, state="disabled")
        # Retake-last drops the most recent capture so the user can
        # redo a bad one (detection sketchy, lighting bad, etc.)
        # without losing everything else.
        self._retake_ref_btn = tk.Button(
            self._btn_frame, text=_("Retake last"),
            command=self._on_retake_last_clicked,
            font=("Helvetica", 9), width=12, state="disabled")
        self._reset_btn = tk.Button(
            self._btn_frame, text=_("Reset all"),
            command=self._on_reset_captures_clicked,
            font=("Helvetica", 9), width=10, state="disabled")
        self._calibrate_btn = tk.Button(
            self._btn_frame, text=_("Calibrate & Save"),
            command=self._on_calibrate_clicked,
            font=("Helvetica", 10), width=18, state="disabled")

        # Reconnect Camera — recovery if the USB camera dies mid-session
        # (the long engrave saturates USB and the cam can stall or drop;
        # cv2.VideoCapture has a stale handle after that, so we need
        # an explicit release+reopen to come back).
        self._reconnect_btn = tk.Button(
            self._btn_frame, text=_("Reconnect camera"),
            command=self._on_reconnect_camera_clicked,
            font=("Helvetica", 9), width=18)
        # Switch camera — fallback for when the auto-resolver picked
        # the wrong device (laptop webcam vs overhead Falcon cam).
        # Persists the chosen index so other camera-using dialogs also
        # default correctly afterward.
        self._switch_cam_btn = tk.Button(
            self._btn_frame, text=_("Switch camera"),
            command=self._on_switch_camera_clicked,
            font=("Helvetica", 9), width=18)

        # Manual-offset entry — escape hatch when an engrave succeeded
        # but the persisted offset was lost (e.g. older settings file
        # missing the DEFAULT_SETTINGS key, manual config edit, etc.)
        # and the card is still sitting on the bed at known coords.
        # Avoids requiring another 2-hour re-engrave to recover.
        self._manual_offset_btn = tk.Button(
            self._btn_frame, text=_("Enter offset manually..."),
            command=self._on_manual_offset_clicked,
            font=("Helvetica", 9), width=22)

        # Always-visible Cancel
        self._cancel_btn = tk.Button(
            self._btn_frame, text=_("Cancel"),
            command=self._on_close, width=10)

        # Wide default: preview width + right panel + padding.
        # Tall enough for the right panel's stacked controls.
        self.geometry(f"{self.PREVIEW_W + 380}x{max(self.PREVIEW_H, 600)}")
        self.minsize(900, 600)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Enter the initial phase. If no Falcon connected, skip
        # straight to capture phase (user runs engrave externally).
        if self._phase == 'engrave':
            self._enter_engrave_phase()
        else:
            self._enter_capture_phase()
        self.wait_window(self)

    def _enter_engrave_phase(self):
        """Show the engrave-phase UI: live camera preview + jog buttons
        + MPos display + "Engrave Centered Here" button. User homes the
        laser, jogs into the camera view, then clicks Engrave — card
        engraves around the head's current position."""
        self._phase = 'engrave'
        self._instructions_var.set(_(
            "Step 1 of 2 — engrave the calibration card.\n\n"
            "(Dimensions below are for the Falcon2 Pro 40W. Adjust "
            "for your machine. The engrave uses your basswood G-code "
            "preset — tune it under Tooling > Options.)\n\n"
            "1. Place a 12×12-inch basswood blank near the center "
            "of the bed and secure it. It cannot move between "
            "engraving and the Step 2 captures.\n"
            "2. Click Home Laser — head parks at the home corner.\n"
            "3. JOG THE HEAD TO ROUGHLY BED CENTER (X=200, Y=200, "
            "or X=~8\", Y=~8\"). The card is ~11\" square and "
            "engraves CENTERED on the head, so at the home corner "
            "the trace runs off the bed. Use 50mm step for fast "
            "travel, then 10mm or 1mm to fine-tune.\n"
            "4. Click Frame — low-power outline traces where the "
            "card will engrave (~11×11 inches). The outline follows "
            "your jog buttons in real time. Adjust until the trace "
            "sits comfortably on the basswood (~½\" margin all "
            "around). Click Stop on the framing window.\n"
            "5. Click Engrave to commit. (Takes a while.)"))
        self._status_var.set(_("Ready. Home the laser and jog into position."))
        self._preview_label.config(image="", text=_("Camera opening..."),
                                    fg="#888888", bg="#000000")
        for w in (self._capture_btn, self._done_refs_btn,
                   self._retake_ref_btn, self._reset_btn,
                   self._reengrave_btn, self._calibrate_btn,
                   self._next_btn):
            w.pack_forget()
        self._engrave_btn.config(text=_("Engrave"))
        # Build jog cluster + home + MPos display (once)
        if not hasattr(self, '_engrave_jog_built'):
            self._build_engrave_controls()
            self._engrave_jog_built = True
        # side='bottom' so the jog cluster pins above the bottom
        # button row (status_var) regardless of preview size. Without
        # this, on small screens the jog cluster falls off-screen
        # below the preview.
        self._engrave_jog_frame.pack(side='bottom', pady=(0, 8))
        # Stack vertically in the narrow (340px) right panel.
        self._frame_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._engrave_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._next_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reconnect_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._switch_cam_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._manual_offset_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._cancel_btn.pack(side="top", fill='x', padx=5, pady=2)
        # Reset to disabled, then enable based on current flow flags.
        # Without the reset, going from capture phase BACK to engrave
        # phase via Engrave-new-card would leave stale enabled state.
        self._engrave_btn.config(state="disabled")
        self._next_btn.config(state="disabled")
        if self._captured_engrave_offset is not None:
            self._engrave_btn.config(state="normal")
            self._next_btn.config(state="normal")
        elif getattr(self, '_framing_done', False):
            self._engrave_btn.config(state="normal")
        # Open the camera early so the user can see the head as they jog.
        if self._cap is None and self._detector_imports_ok:
            self.after(50, self._open_camera_and_start)
        self._start_mpos_polling()

    def _build_engrave_controls(self):
        """Build the home/jog/MPos cluster used during engrave-phase
        positioning. Parented to the right-panel control area so it
        sits with the other controls, not floating in the dialog."""
        self._engrave_jog_frame = tk.LabelFrame(
            self._btn_frame.master, text=_("Position the head"),
            bg=DIALOG_BG, padx=8, pady=4)

        row = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        row.pack(anchor='w')
        tk.Button(row, text=_("Home Laser ($H)"),
                   command=self._on_engrave_home).pack(side='left', padx=4)
        self._mpos_var = tk.StringVar(value=_("MPos: --"))
        tk.Label(row, textvariable=self._mpos_var, bg=DIALOG_BG,
                  font=("Courier", 10)).pack(side='left', padx=10)

        step_row = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        step_row.pack(anchor='w', pady=(4, 2))
        tk.Label(step_row, text=_("Step (mm):"), bg=DIALOG_BG,
                  font=("Helvetica", 9)).pack(side='left')
        self._engrave_jog_step = tk.DoubleVar(value=10.0)
        for step in (1.0, 10.0, 50.0):
            tk.Radiobutton(
                step_row, text=str(step),
                variable=self._engrave_jog_step, value=step,
                bg=DIALOG_BG, font=("Helvetica", 9),
                highlightthickness=0).pack(side='left')

        btn_grid = tk.Frame(self._engrave_jog_frame, bg=DIALOG_BG)
        btn_grid.pack(pady=2)
        tk.Button(btn_grid, text="↑", width=3,
                   command=lambda: self._engrave_jog(0, 1)
                   ).grid(row=0, column=1, padx=2, pady=1)
        tk.Button(btn_grid, text="←", width=3,
                   command=lambda: self._engrave_jog(-1, 0)
                   ).grid(row=1, column=0, padx=2, pady=1)
        tk.Button(btn_grid, text="→", width=3,
                   command=lambda: self._engrave_jog(1, 0)
                   ).grid(row=1, column=2, padx=2, pady=1)
        tk.Button(btn_grid, text="↓", width=3,
                   command=lambda: self._engrave_jog(0, -1)
                   ).grid(row=2, column=1, padx=2, pady=1)

    def _start_mpos_polling(self):
        """Poll Falcon's MPos every ~300ms while in engrave phase."""
        if self._phase != 'engrave' or not self._falcon_port:
            return
        self._mpos_after_id = self.after(300, self._poll_mpos)

    def _poll_mpos(self):
        if self._phase != 'engrave' or not self._falcon_port:
            return
        # Quick connect/disconnect to read status — the persistent
        # sender pattern would tie up the port and conflict with the
        # engrave/jog actions which also need to connect.
        try:
            import falcon_sender
            s = falcon_sender.FalconSender(port=self._falcon_port)
            s.connect()
            st = s.get_status(timeout=0.2)
            s.disconnect()
            if st and st.get('mpos'):
                self._mpos_var.set(
                    _("MPos: X={x:7.2f}  Y={y:7.2f}").format(
                        x=st['mpos'][0], y=st['mpos'][1]))
        except Exception:
            pass
        if self._phase == 'engrave':
            self._mpos_after_id = self.after(800, self._poll_mpos)

    def _on_engrave_home(self):
        """Send $H standalone, with a small "Homing..." indicator."""
        if not self._falcon_port:
            return
        try:
            import falcon_sender
        except ImportError:
            return
        if not messagebox.askyesno(
                _("Home Laser?"),
                _("Send the laser home? The head will travel to the "
                  "home corner at full speed."), parent=self):
            return
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        self._status_var.set(_("Homing..."))
        self.update_idletasks()
        import threading
        result = {}

        def _worker():
            try:
                ok, m = sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = m
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        t = threading.Thread(target=_worker, daemon=True)
        t.start()
        # after()-poll until done. wait_variable pumps the Tk event
        # loop without the WM_DELETE_WINDOW destroy-root-mid-loop
        # hazard of `while t.is_alive(): self.update()`.
        done_var = tk.BooleanVar(value=False)

        def _check():
            if t.is_alive():
                self.after(50, _check)
            else:
                done_var.set(True)

        self.after(50, _check)
        self.wait_variable(done_var)
        try:
            sender.disconnect()
        except Exception:
            pass
        if result.get('ok'):
            self._engrave_homed = True
            self._status_var.set(_("Homed. Now jog into the camera view."))
        else:
            self._status_var.set(
                _("Home failed: {m}").format(m=result.get('msg', '?')))

    def _engrave_jog(self, dx_sign, dy_sign):
        """Send a relative jog via Grbl $J=."""
        if not self._falcon_port:
            return
        try:
            import falcon_sender
            s = falcon_sender.FalconSender(port=self._falcon_port)
            s.connect()
            try:
                step = float(self._engrave_jog_step.get())
            except Exception:
                step = 10.0
            s.jog(x=dx_sign * step, y=dy_sign * step,
                   feed=3000, relative=True)
            s.disconnect()
        except Exception:
            pass

    def _enter_capture_phase(self):
        """Switch to live-preview + ChArUco capture mode."""
        self._phase = 'capture'
        self._instructions_var.set(_(
            "Step 2 of 2 — calibration captures.\n\n"
            "FIRST capture(s): board must stay at its engraved "
            "position. This anchors camera to machine coords; any "
            "shift here offsets every future scrap capture by that "
            "amount.\n\n"
            "Click 'Done with references' after 1+ high quality "
            "reference shots, then move the card to different spots "
            "between the remaining captures so the math can solve "
            "lens distortion. Watch the green-corner overlay; click "
            "Capture when the board is fully or mostly detected. "
            "Play with lighting if necessary.\n\n"
            "Target: {n} good captures total."
        ).format(n=self.TARGET_FRAMES))
        # Stop MPos polling and hide engrave-phase controls.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None
        for w in (self._frame_btn, self._engrave_btn, self._next_btn,
                   self._reconnect_btn, self._manual_offset_btn):
            w.pack_forget()
        if hasattr(self, '_engrave_jog_frame'):
            self._engrave_jog_frame.pack_forget()
        self._status_var.set(_("Opening camera ..."))
        self._capture_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._done_refs_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._retake_ref_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reset_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reengrave_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._calibrate_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._reconnect_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._switch_cam_btn.pack(side="top", fill='x', padx=5, pady=2)
        self._cancel_btn.pack(side="top", fill='x', padx=5, pady=2)
        # Always-visible Reconnect — the camera can die during the long
        # engrave or at dialog-transition. Without an explicit reopen
        # path the user has to cancel and lose any in-memory captures.
        # Kick off camera open + live refresh (may already be open
        # from engrave phase — _open_camera_and_start handles that).
        self.after(50, self._open_camera_and_start)

    def _engrave_geometry(self):
        """Return (cols, rows, sq, board_w, board_h, border, label_h)
        — the card-layout constants the engrave + framing flows both
        need."""
        import camera_capture as _cam
        cols, rows = _cam.CHARUCO_COLS, _cam.CHARUCO_ROWS
        sq = _cam.CHARUCO_SQUARE_MM
        return (cols, rows, sq, cols * sq, rows * sq, 10.0, 8.0)

    def _engrave_check_position_safe(self):
        """Read current MPos and verify the engrave area would fit on
        the bed. Returns True if safe, else False (with an error
        already shown). Avoids ALARM:3 from soft-limit trips when
        the user clicks Frame with the head still at home corner."""
        try:
            import falcon_sender
        except ImportError:
            return True  # can't check; let it run and fail loudly
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            st = sender.get_status(timeout=0.5)
        except Exception:
            try:
                sender.disconnect()
            except Exception:
                pass
            return True
        try:
            sender.disconnect()
        except Exception:
            pass
        if not st or not st.get('mpos'):
            return True
        hx, hy = st['mpos'][0], st['mpos'][1]
        cols, rows, sq, board_w, board_h, border, label_h = \
            self._engrave_geometry()
        ox = hx - border - board_w / 2.0
        oy = hy - border - label_h - board_h / 2.0
        fx = ox + board_w + 2 * border
        fy = oy + board_h + 2 * border + label_h
        bed_x_max = float(
            self._settings.get("laser_bed_x_max", 400.0))
        bed_y_max = float(
            self._settings.get("laser_bed_y_max", 415.0))
        if (ox < 0 or oy < 0
                or fx > bed_x_max + 1.0 or fy > bed_y_max + 1.0):
            messagebox.showerror(
                _("Card Would Run Off Bed"),
                _("With head at machine (X={x:.1f}, Y={y:.1f}), the "
                  "card would engrave at ({ox:.0f},{oy:.0f}) to "
                  "({fx:.0f},{fy:.0f}) — outside the {bx:.0f}×{by:.0f}"
                  " bed.\n\n"
                  "Jog the head further from the home corner so the "
                  "whole card area fits on the bed (and on basswood).\n\n"
                  "Need head at roughly X={midx:.0f}-{maxx:.0f}, "
                  "Y={midy:.0f}-{maxy:.0f}.").format(
                    x=hx, y=hy, ox=ox, oy=oy, fx=fx, fy=fy,
                    bx=bed_x_max, by=bed_y_max,
                    midx=border + board_w / 2.0,
                    maxx=bed_x_max - border - board_w / 2.0,
                    midy=border + label_h + board_h / 2.0,
                    maxy=bed_y_max - border - board_h / 2.0,
                ), parent=self)
            return False
        return True

    def _engrave_check_ready(self):
        """Common pre-flight for Frame / Engrave: Falcon connected,
        laser homed. Returns True if all good, else False (with the
        appropriate error/prompt already shown). Sets self._engrave_homed
        based on user choice."""
        if not self._falcon_port:
            messagebox.showerror(
                _("Falcon Not Detected"),
                _("Can't operate — no Falcon connection."), parent=self)
            return False
        if not self._engrave_homed:
            choice = messagebox.askyesnocancel(
                _("Home Laser First?"),
                _("The laser needs to be homed so machine coordinates "
                  "have a known reference. Without homing, the "
                  "engrave's position is bogus.\n\n"
                  "Yes — home now (head moves to home corner).\n"
                  "No — I already homed this session (LightBurn / "
                  "physical button / before opening SSC).\n"
                  "Cancel — back out."),
                parent=self)
            if choice is None:
                return False
            if choice:
                self._on_engrave_home()
                if not self._engrave_homed:
                    return False
            else:
                self._engrave_homed = True
        return True

    def _on_frame_clicked(self):
        """Run a continuous framing loop that re-reads MPos each
        iteration — the trace FOLLOWS the head as the user jogs
        with the in-dialog buttons. User clicks Stop or the X to
        end framing, then can verify alignment and click Engrave."""
        if not self._engrave_check_ready():
            return
        if not self._engrave_check_position_safe():
            return
        try:
            import falcon_sender
            from gcode_engine import generate_framing_gcode
        except ImportError as e:
            messagebox.showerror(_("Missing Dependency"),
                                  _("Frame needs pyserial: {e}"
                                    ).format(e=e), parent=self)
            return
        cols, rows, sq, board_w, board_h, border, label_h = \
            self._engrave_geometry()

        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return

        def _framing_provider():
            st = sender.get_status(timeout=0.3)
            if not st or not st.get('mpos'):
                # Fall back to a default near home rather than crash
                hx, hy = 100.0, 100.0
            else:
                hx, hy = st['mpos'][0], st['mpos'][1]
            ox = hx - border - board_w / 2.0
            oy = hy - border - label_h - board_h / 2.0
            fx = ox + board_w + 2 * border
            fy = oy + board_h + 2 * border + label_h
            return generate_framing_gcode(
                ox, oy, fx, fy,
                power_s=self._settings.get("laser_framing_power_s", 10),
                feed=self._settings.get("laser_framing_feed", 2000),
                return_to_origin=False,  # absolute coords; don't send
                                          # head to machine (0, 0) (=home
                                          # corner) between iterations —
                                          # that loses user's jog edits.
                # Park at the head's original MPos so the next iter
                # produces the same bbox (and so the engrave-phase
                # MPos read after Done matches what the user jogged
                # to). The card geometry is slightly asymmetric in Y
                # (label adds ~4mm offset) — parking at hx, hy keeps
                # MPos invariant across iterations.
                park_xy=(hx, hy),
            )

        # Stop MPos polling while the streamer owns the port.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None

        try:
            FalconRunDialog(
                self, sender, _framing_provider(),
                title=_("Framing — jog to position, click Done when "
                         "alignment looks right"),
                loop=True, gcode_provider=_framing_provider,
                show_pause_resume=False, show_cut_button=False,
                stop_needs_confirm=False)
        finally:
            try:
                sender.disconnect()
            except Exception:
                pass
            self._start_mpos_polling()
            # Recycle the camera handle. Empirically the cv2.VideoCapture
            # frame stream stalls (frozen preview) when a long-lived
            # modal FalconRunDialog closes — probably a Windows
            # DirectShow / USB-focus quirk we can't suppress at the
            # cv2 layer. Cheaper to just reopen than diagnose.
            self._auto_recycle_camera()
        # Framing window closed (Stop or X). Treat as "user is done
        # positioning" — enable the Engrave button. User may run
        # Frame again to re-verify or jog more; that's fine.
        self._framing_done = True
        if self._phase == 'engrave':
            try:
                self._engrave_btn.config(state="normal")
                self._status_var.set(
                    _("Framing done. Click Engrave when ready."))
            except tk.TclError:
                pass

    def _on_next_clicked(self):
        """Engrave-phase Next button: advance to capture phase. Only
        active after the engrave has completed and saved an offset."""
        if self._captured_engrave_offset is None:
            messagebox.showerror(
                _("Engrave first"),
                _("No engrave offset on file yet. Click Engrave (and "
                  "wait for it to finish) before moving to captures."),
                parent=self)
            return
        self._enter_capture_phase()

    def _on_engrave_clicked(self):
        """Commit to the engrave at the head's current position.
        Confirmation popup ("are you sure?"). If yes, ~45-60 minute
        engrave runs. If no, returns to the engrave-phase dialog
        where the user can Frame again or jog."""
        if not self._engrave_check_ready():
            return
        if not self._engrave_check_position_safe():
            return
        try:
            import falcon_sender
            from gcode_engine import generate_calibration_card_gcode
            import camera_capture as _cam
        except ImportError as e:
            messagebox.showerror(
                _("Missing Dependency"),
                _("Engrave needs pyserial:\n\n{e}").format(e=e),
                parent=self)
            return

        # Read the head's current MPos to center the engrave around it.
        sender = falcon_sender.FalconSender(port=self._falcon_port)
        try:
            sender.connect()
            sender.unlock()
            status = sender.get_status(timeout=0.5)
        except Exception as e:
            messagebox.showerror(_("Connection Failed"),
                                  str(e), parent=self)
            try:
                sender.disconnect()
            except Exception:
                pass
            return
        if not status or not status.get('mpos'):
            messagebox.showerror(
                _("No machine position"),
                _("Couldn't read MPos. Home the laser first."),
                parent=self)
            sender.disconnect()
            return
        head_x, head_y = status['mpos'][0], status['mpos'][1]

        # Compute offset so the BOARD's center lands at (head_x, head_y).
        # Card layout: bottom-left at (offset_x, offset_y), board area
        # at (offset_x + border, offset_y + border + label_h) to
        # (offset_x + border + board_w, offset_y + border + label_h + board_h).
        border = 10.0
        label_h = 8.0
        cols, rows = _cam.CHARUCO_COLS, _cam.CHARUCO_ROWS
        sq = _cam.CHARUCO_SQUARE_MM
        board_w = cols * sq
        board_h = rows * sq
        offset_x = head_x - border - board_w / 2.0
        offset_y = head_y - border - label_h - board_h / 2.0

        card_far_x = offset_x + board_w + 2 * border
        card_far_y = offset_y + board_h + 2 * border + label_h
        if not messagebox.askyesno(
                _("Start Engrave?"),
                _("Are you sure? Card will engrave centered on machine "
                  "(X={x:.0f}, Y={y:.0f}), covering ({ox:.0f}, "
                  "{oy:.0f}) to ({fx:.0f}, {fy:.0f}).\n\n"
                  "Takes a while. The cover must stay closed "
                  "for the laser to fire.\n\n"
                  "If you haven't framed yet to verify placement, "
                  "click No and run Frame first.").format(
                    x=head_x, y=head_y,
                    ox=offset_x, oy=offset_y,
                    fx=card_far_x, fy=card_far_y),
                parent=self):
            sender.disconnect()
            return

        # Generate G-code to a temp file (no $H — head's already homed
        # and at the user-chosen jog position; G-code uses absolute
        # coords from there). Read as a list of lines for the streamer.
        import tempfile
        import os as _os
        with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False,
                                          mode='w') as tmp:
            tmp_path = tmp.name
        try:
            generate_calibration_card_gcode(
                tmp_path,
                cols=cols, rows=rows,
                square_mm=sq, marker_mm=_cam.CHARUCO_MARKER_MM,
                settings=self._settings,
                offset_x_mm=offset_x, offset_y_mm=offset_y,
                home_first=False)  # already homed
            with open(tmp_path, 'r') as f:
                gcode_lines = f.read().splitlines()
        finally:
            try:
                _os.remove(tmp_path)
            except Exception:
                pass

        # Stop polling MPos while the stream owns the port.
        if hasattr(self, '_mpos_after_id') and self._mpos_after_id:
            self.after_cancel(self._mpos_after_id)
            self._mpos_after_id = None

        try:
            run_dlg = FalconRunDialog(
                self, sender, gcode_lines,
                title=_("Engraving Calibration Card"))
            # Run dialog closed (regardless of success/stop/error).
            # Recycle the camera before the next phase needs it — see
            # _auto_recycle_camera for the rationale.
            self._auto_recycle_camera()
            if run_dlg._final_reason != "complete":
                # Engrave stopped or errored mid-flight — the head
                # is now at some weird mid-trace MPos that's no
                # longer the planned card center. Force a re-frame
                # (with fresh basswood, presumably) before allowing
                # another engrave.
                self._framing_done = False
                try:
                    self._engrave_btn.config(state="disabled")
                    self._status_var.set(
                        _("Engrave stopped. Put fresh basswood, "
                          "re-Frame, then Engrave."))
                except tk.TclError:
                    pass
                self._start_mpos_polling()
                return
        finally:
            try:
                sender.disconnect()
            except Exception:
                pass

        # Engrave done — remember the offset for the calibration solver
        # AND persist it to settings so a closed/reopened dialog or
        # crashed-then-relaunched SSC can pick up where we left off
        # (the engrave is the most expensive step; losing the offset
        # to a UI hiccup wastes ~2 hours of basswood).
        self._captured_engrave_offset = (offset_x, offset_y)
        try:
            self._settings['camera_calibration_engrave_offset_mm'] = [
                offset_x, offset_y]
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            pass  # never block on settings save
        # Activate Next — but DON'T auto-transition. User clicks Next
        # explicitly. Prevents losing the offset to a dialog-close
        # accident, and lets the user verify the engrave looks right
        # before moving on. Disable Frame + Engrave so user can't
        # accidentally re-engrave at a different MPos (would mismatch
        # the saved offset).
        try:
            self._frame_btn.config(state="disabled")
            self._engrave_btn.config(state="disabled")
            self._next_btn.config(state="normal")
            self._status_var.set(
                _("Engrave complete. Click Next when ready for "
                  "calibration captures."))
        except tk.TclError:
            pass
        messagebox.showinfo(
            _("Card Engraved — DO NOT TOUCH"),
            _("Card engraved successfully.\n\n"
              "⚠ DO NOT MOVE THE CARD. ⚠\n\n"
              "Click Next to start calibration captures. The first "
              "capture becomes the position reference linking the "
              "camera to machine coords — if the card moves before "
              "then, the calibration is off."), parent=self)

    def _on_reengrave_clicked(self):
        """Capture-phase button: discard the saved engrave offset and
        return to engrave phase so the user can engrave a NEW card
        on fresh basswood. Can't re-engrave the same wood — each
        engrave is permanent.

        Also clears any captures already taken in this session
        (different engrave = different machine-coord reference, so
        old captures would be tied to the wrong position).
        """
        if not messagebox.askyesno(
                _("Engrave a New Card?"),
                _("Discard the existing calibration session and start "
                  "over with a fresh basswood blank?\n\n"
                  "You CAN'T re-engrave the same wood — each engrave "
                  "is permanent. Use a new piece of basswood.\n\n"
                  "Current captures ({n}) will be discarded."
                  ).format(n=len(self._captures)),
                parent=self):
            return
        self._captures = []
        self._reference_count = 0
        self._capture_substate = 'reference'
        self._captured_engrave_offset = None
        # Force fresh framing for the new card — new wood / new
        # position warrants re-positioning, so disable Engrave until
        # they re-frame.
        self._framing_done = False
        # Clear the persisted offset so the next dialog open lands
        # in engrave phase again (not auto-skipped to capture).
        try:
            self._settings.pop(
                'camera_calibration_engrave_offset_mm', None)
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            pass
        self._enter_engrave_phase()

    # ------------------------------------------------------------------
    def _open_camera_and_start(self):
        """Open the camera on a background thread so the dialog stays
        responsive during the 1-3 second cv2.VideoCapture warm-up."""
        import threading
        self._status_var.set(
            _("Opening camera..."))

        def _worker():
            try:
                cap = self._cam_mod.open_camera(self._camera_index)
                ok, frame = cap.read()
                if not ok or frame is None:
                    cap.release()
                    raise RuntimeError(_("Camera opened but returned no frame."))
                self.after(0, self._camera_ready, cap, frame)
            except Exception as e:
                self.after(0, self._camera_failed, e)

        threading.Thread(target=_worker, name='camcal-open',
                          daemon=True).start()

    def _camera_ready(self, cap, first_frame):
        self._cap = cap
        h, w = first_frame.shape[:2]
        self._image_size = (w, h)
        self._last_frame_shape = (h, w)
        self._status_var.set(
            _("Camera live. Captured 0 / {n} good frames.").format(n=self.TARGET_FRAMES))
        self._refresh_loop()

    def _camera_failed(self, err):
        self._status_var.set(
            _("Could not open camera index {i}: {e}").format(
                i=self._camera_index, e=err))

    def _refresh_loop(self):
        if not self.winfo_exists() or self._cap is None:
            return
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok, frame = False, None
        if ok and frame is not None:
            self._failed_reads = 0
            self._latest_frame = frame
            try:
                self._render_frame_with_overlay(frame)
            except Exception:
                pass  # don't kill the loop on a bad frame
        else:
            # Camera disconnect / USB hung. Surface it after a few
            # consecutive failures so the user knows to click
            # Reconnect rather than staring at a frozen preview.
            self._failed_reads = getattr(self, '_failed_reads', 0) + 1
            if self._failed_reads == 30:  # ~3s at 10fps
                try:
                    self._status_var.set(_(
                        "Camera lost — click Reconnect Camera to recover."))
                except tk.TclError:
                    pass
        self._refresh_after_id = self.after(self.LIVE_REFRESH_MS,
                                              self._refresh_loop)

    def _on_manual_offset_clicked(self):
        """Engrave-phase escape hatch. If the user knows the machine-mm
        coords where their engraved card sits (e.g. recovered from a
        backup, or printed at engrave time), they can punch in the
        offset and skip straight to capture phase. Avoids losing 2
        hours of basswood to a state-persistence hiccup."""
        from tkinter import simpledialog
        prompt = _(
            "Enter the engraved card's bottom-left machine-mm coords.\n\n"
            "Format: X,Y (e.g. 66.025,99.025)\n\n"
            "This is the offset that was persisted to settings after a "
            "successful engrave. Use this if the persisted value got "
            "lost but the card is still on the bed at known coords.")
        val = simpledialog.askstring(
            _("Manual Engrave Offset"), prompt, parent=self)
        if not val:
            return
        try:
            parts = val.replace(';', ',').split(',')
            if len(parts) != 2:
                raise ValueError("need two numbers separated by a comma")
            ox = float(parts[0].strip())
            oy = float(parts[1].strip())
        except (ValueError, TypeError) as e:
            messagebox.showerror(
                _("Bad Offset"),
                _("Couldn't parse '{v}' as X,Y mm:\n{e}").format(v=val, e=e),
                parent=self)
            return
        self._captured_engrave_offset = (ox, oy)
        try:
            self._settings['camera_calibration_engrave_offset_mm'] = [ox, oy]
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            pass
        # Re-enter engrave phase so Next becomes enabled (the offset
        # check happens at phase-enter time), then user can click Next
        # to advance to capture.
        self._enter_engrave_phase()
        self._status_var.set(
            _("Offset set to ({x:.3f}, {y:.3f}). Click Next to capture.").format(
                x=ox, y=oy))

    def _on_reconnect_camera_clicked(self):
        """Release the (likely-dead) cv2.VideoCapture handle and open
        a fresh one. Used when the USB camera died mid-session and the
        existing handle is stale. Doesn't touch captures already taken
        so the user keeps their progress."""
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self._failed_reads = 0
        self._status_var.set(_("Reopening camera..."))
        self.after(50, self._open_camera_and_start)

    def _on_switch_camera_clicked(self):
        """Cycle to the next enumerated camera; persist the choice."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        try:
            self._settings['camera_index_override'] = int(self._camera_index)
            from config import save_settings
            save_settings(self._settings)
        except Exception:
            pass
        self._status_var.set(
            _("Switched to camera index {i}. Reopening...").format(
                i=self._camera_index))
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self._failed_reads = 0
        self.after(50, self._open_camera_and_start)

    def _auto_recycle_camera(self):
        """Called after every FalconRunDialog close. The cv2 frame
        stream consistently stalls when a long-lived modal dialog
        closes (preview shows the last good frame forever, with no
        error from cap.read()). Reopen the camera to get fresh frames
        flowing again. No-op if the camera wasn't open."""
        if self._cap is None:
            return
        self._stop_refresh()
        try:
            self._cap.release()
        except Exception:
            pass
        self._cap = None
        self._failed_reads = 0
        self.after(50, self._open_camera_and_start)

    def _render_frame_with_overlay(self, frame):
        import cv2
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        corners, ids = self._cam_mod.detect_charuco(gray, self._board)

        annotated = frame.copy()
        detected_count = 0 if corners is None else len(corners)
        if corners is not None and ids is not None and len(corners) >= 4:
            # Draw green dots at detected ChArUco corners
            for pt in corners:
                x, y = int(pt[0][0]), int(pt[0][1])
                cv2.circle(annotated, (x, y), 4, (0, 255, 0), -1)
            self._capture_btn.config(state="normal")
        else:
            self._capture_btn.config(state="disabled")

        # OpenCV is BGR; PIL/tk want RGB.
        rgb = cv2.cvtColor(annotated, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        # Cap preview size while preserving aspect ratio
        img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
        self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
        self._preview_label.config(image=self._latest_photo,
                                    width=img.width, height=img.height)

        n_caps = len(self._captures)
        if corners is not None:
            self._status_var.set(
                _("Detected {d} corners — captured {n} / {target} good frames.").format(
                    d=detected_count, n=n_caps, target=self.TARGET_FRAMES))
        else:
            self._status_var.set(
                _("Move the card into view — captured {n} / {target} good frames.").format(
                    n=n_caps, target=self.TARGET_FRAMES))

    def _on_capture_clicked(self):
        if not hasattr(self, '_latest_frame'):
            return
        import cv2
        gray = cv2.cvtColor(self._latest_frame, cv2.COLOR_BGR2GRAY)
        corners, ids = self._cam_mod.detect_charuco(gray, self._board)
        if corners is None or ids is None or len(corners) < 4:
            return
        self._captures.append((corners, ids))
        is_ref = (self._capture_substate == 'reference')
        if is_ref:
            self._reference_count += 1
        n = len(self._captures)
        target = self.TARGET_FRAMES
        if n >= self._cam_mod.CALIB_MIN_FRAMES:
            self._calibrate_btn.config(state="normal")
        if n >= 1:
            self._retake_ref_btn.config(state="normal")
            self._reset_btn.config(state="normal")
            self._done_refs_btn.config(
                state="normal" if is_ref else "disabled")
        if is_ref:
            hint = _("REFERENCE captured ({nref} of card untouched). "
                      "Take more references for higher accuracy, "
                      "Retake last if detection was bad, or click "
                      "'Done with references' when ready to MOVE "
                      "the card.").format(nref=self._reference_count)
        elif n >= target:
            hint = _("Click Calibrate & Save when ready.")
        else:
            hint = _("MOVE the card to a new position and capture again.")
        self._status_var.set(
            _("Captured {n} / {t} good frames. {hint}").format(
                n=n, t=target, hint=hint))

    def _on_done_refs_clicked(self):
        """Switch from reference substate to intrinsics substate.
        Subsequent captures will be 'card moved' poses for the
        distortion solve."""
        if self._reference_count < 1:
            return
        self._capture_substate = 'intrinsics'
        self._capture_btn.config(text=_("Capture (card moved)"))
        self._done_refs_btn.config(state="disabled")
        self._status_var.set(
            _("References done ({n} pooled). Now MOVE the card to "
              "different positions / tilts and capture {m} more "
              "frames for lens-distortion calibration.").format(
                n=self._reference_count,
                m=max(0, self.TARGET_FRAMES - len(self._captures))))

    def _on_retake_last_clicked(self):
        """Drop the most recent capture so the user can redo it."""
        if not self._captures:
            return
        # Decrement reference count if the removed capture was a reference
        if (self._capture_substate == 'reference'
                and self._reference_count > 0):
            self._reference_count -= 1
        elif (self._capture_substate == 'intrinsics'
              and len(self._captures) - 1 < self._reference_count):
            # Defensive: shouldn't happen, but if it did, fix the count
            self._reference_count = len(self._captures) - 1
        self._captures.pop()
        n = len(self._captures)
        self._calibrate_btn.config(
            state="normal" if n >= self._cam_mod.CALIB_MIN_FRAMES
            else "disabled")
        if n == 0:
            self._retake_ref_btn.config(state="disabled")
            self._reset_btn.config(state="disabled")
            self._done_refs_btn.config(state="disabled")
        self._status_var.set(
            _("Removed last capture. {n} / {t} remaining.").format(
                n=n, t=self.TARGET_FRAMES))

    def _on_reset_captures_clicked(self):
        """Discard all captures and restart. Card needs to be back at
        the engraved position for the next first reference capture."""
        if not messagebox.askyesno(
                _("Reset all captures?"),
                _("Discard all {n} captures and start over.\n\n"
                  "If you've moved the card, move it BACK to its "
                  "engraved position before retaking references.\n\n"
                  "Continue?").format(n=len(self._captures)),
                parent=self):
            return
        self._captures = []
        self._reference_count = 0
        self._capture_substate = 'reference'
        self._capture_btn.config(text=_("Capture reference"))
        self._retake_ref_btn.config(state="disabled")
        self._reset_btn.config(state="disabled")
        self._done_refs_btn.config(state="disabled")
        self._calibrate_btn.config(state="disabled")
        self._status_var.set(
            _("Captures cleared. Place the card in its engraved "
              "position and click Capture reference."))

    def _on_calibrate_clicked(self):
        if len(self._captures) < self._cam_mod.CALIB_MIN_FRAMES:
            return
        self._stop_refresh()
        self._status_var.set(_("Computing calibration ..."))
        self.update()
        try:
            eng_off = getattr(self, '_captured_engrave_offset', None)
            kwargs = {}
            if eng_off:
                kwargs['card_offset_x_mm'] = eng_off[0]
                kwargs['card_offset_y_mm'] = eng_off[1]
            # Pool ALL reference captures for the machine-mm homography.
            # The first N captures (where N = _reference_count) were
            # taken with the card untouched at its engraved position.
            ref_count = max(1, getattr(self, '_reference_count', 1))
            kwargs['reference_indices'] = tuple(range(ref_count))
            calib = self._cam_mod.calibrate_from_frames(
                self._captures, self._image_size, self._board, **kwargs)
            self._cam_mod.save_calibration(calib, self._calibration_path)
            # Calibration committed. Clear the in-progress recovery
            # hint (engrave offset in settings) — it was useful for
            # restoring state during the capture phase, but now the
            # full calibration is on disk. Next dialog open should
            # start in engrave phase for a FRESH recalibration, not
            # auto-skip to capture against a stale offset.
            #
            # Also persist the working camera index so future opens
            # (live preview, scrap capture) skip the enumerate /
            # find-Falcon guess and go straight to the right camera.
            try:
                self._settings.pop(
                    'camera_calibration_engrave_offset_mm', None)
                self._settings['camera_index_override'] = int(
                    self._camera_index)
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                pass
            self.result = "saved"
            messagebox.showinfo(
                _("Calibration Saved"),
                _("Camera calibration complete.\n\n"
                  "RMS reprojection error: {rms:.2f} px (lower is better; "
                  "< 1 px is excellent)\n"
                  "Frames used: {n}\n\n"
                  "Saved to:\n{p}").format(
                    rms=calib['rms_reprojection_error_px'],
                    n=calib['frame_count'], p=self._calibration_path),
                parent=self)
            self._on_close()
        except Exception as e:
            messagebox.showerror(_("Calibration Failed"),
                                  _("Calibration did not converge:\n\n{e}\n\n"
                                    "Try capturing more frames at different "
                                    "positions and tilts.").format(e=e),
                                  parent=self)
            # Resume the refresh loop so user can capture more
            self._refresh_loop()

    def _stop_refresh(self):
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
            self._refresh_after_id = None

    def _on_close(self):
        self._stop_refresh()
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        try:
            import sleep_lock
            sleep_lock.allow_sleep()
        except Exception:
            pass
        self.destroy()


class CameraCaptureDialog(tk.Toplevel):
    """Live camera-capture dialog for snapping a scrap polygon from the bed.

    Shows the live camera feed with the detected scrap contour overlaid
    in green. User can adjust the threshold sensitivity, then click Use
    to accept the polygon (returned in mm coords) or Retry to keep
    looking. Requires a saved camera calibration.

    Result:
        self.result_polygon_mm = list of (x_mm, y_mm) on success, None on cancel.
    """

    LIVE_REFRESH_MS = 100
    PREVIEW_W = 640
    PREVIEW_H = 480

    def __init__(self, parent, camera_index, calibration_path, settings=None):
        super().__init__(parent)
        self.title(_("Capture Scrap Outline"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.result_polygon_mm = None

        self._camera_index = camera_index
        self._cap = None
        self._calibration = None
        self._latest_undistorted = None
        self._latest_polygon_px = None
        self._latest_polygon_mm = None
        self._latest_photo = None
        self._refresh_after_id = None
        # Optional settings dict — when present, the detection-bias
        # slider persists its value across captures via
        # 'camera_detection_threshold_bias'. When None (e.g. when the
        # dialog was spawned from inside PolygonDrawWindow without a
        # settings handle), the slider still works but its value is
        # lost on close.
        self._settings = settings
        try:
            self._initial_bias = int(
                settings.get('camera_detection_threshold_bias', 0)
                if settings else 0)
        except (TypeError, ValueError):
            self._initial_bias = 0

        # Dependency check + calibration load
        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
            self._calibration = camera_capture.load_calibration(calibration_path)
            if not self._calibration:
                raise RuntimeError(_("No camera calibration found. Run "
                                       "Options > Camera Calibration first."))
        except (ImportError, RuntimeError) as e:
            tk.Label(self, text=str(e), bg=DIALOG_BG, fg="#a30000",
                     padx=20, pady=20, wraplength=400).pack()
            tk.Button(self, text=_("Close"),
                      command=self.destroy).pack(pady=(0, 15))
            return

        # ---- UI ----
        tk.Label(self, text=_("Capture Scrap Outline"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold")
                 ).pack(pady=(10, 2))
        tk.Label(self, text=_(
            "Place the scrap on the laser bed (close the cover). The green "
            "outline shows what will be captured. Click Use when it looks "
            "right, or Retry to re-detect."),
            bg=DIALOG_BG, wraplength=620, justify="left",
            font=("Helvetica", 9)).pack(padx=15, pady=(0, 6))

        # Preview label shows a prominent "loading" message until the
        # first camera frame arrives, then displays live frames.
        self._preview_label = tk.Label(
            self, bg="#222222", fg="#cccccc",
            text=_("Waking camera..."),
            font=("Helvetica", 14),
            width=self.PREVIEW_W, height=self.PREVIEW_H)
        self._preview_label.pack(padx=15, pady=(0, 8))

        self._status_var = tk.StringVar(value=_("Opening camera ..."))
        tk.Label(self, textvariable=self._status_var, bg=DIALOG_BG,
                 font=("Helvetica", 10)).pack(pady=(0, 4))

        # Detection sensitivity slider — biases the auto-threshold
        # used to separate scrap from bed. Center = use OpenCV's Otsu
        # auto-pick (works for most lighting). Slide LEFT (more
        # sensitive) when the scrap barely contrasts with the bed,
        # RIGHT (less sensitive) when honeycomb shadows are getting
        # picked up as material. Updates the live overlay on every
        # next frame.
        bias_frame = tk.Frame(self, bg=DIALOG_BG)
        bias_frame.pack(fill='x', padx=20, pady=(0, 6))
        tk.Label(bias_frame, text=_("Detection sensitivity:"),
                 bg=DIALOG_BG, font=("Helvetica", 9)
                 ).pack(side='left', padx=(0, 6))
        tk.Label(bias_frame, text=_("more"), bg=DIALOG_BG,
                 fg='#666', font=("Helvetica", 8)).pack(side='left')
        self._bias_var = tk.IntVar(value=self._initial_bias)
        # Slider value maps directly to threshold_bias:
        #   -80 (left)  = subtract 80 from Otsu → lower threshold →
        #                 more pixels classified as scrap = MORE sensitive
        #   +80 (right) = add 80 → fewer pixels = LESS sensitive
        tk.Scale(bias_frame, from_=-80, to=80, orient='horizontal',
                 variable=self._bias_var, bg=DIALOG_BG,
                 length=300, showvalue=False,
                 command=self._on_bias_changed).pack(
                     side='left', padx=4, fill='x', expand=True)
        tk.Label(bias_frame, text=_("less"), bg=DIALOG_BG,
                 fg='#666', font=("Helvetica", 8)).pack(side='left')

        # Invert toggle — flips the threshold polarity. Default off
        # (THRESH_BINARY) assumes scrap is BRIGHTER than the bed (the
        # common case: leather on a dark honeycomb). Turn on for the
        # opposite (e.g. dark felt on a light scrap board) — without
        # inversion, dark-on-light produces no contour because the
        # largest bright region is the bed itself. Persisted so users
        # who routinely shoot one polarity don't re-toggle each time.
        invert_frame = tk.Frame(self, bg=DIALOG_BG)
        invert_frame.pack(fill='x', padx=20, pady=(0, 6))
        try:
            initial_invert = bool(
                self._settings.get('camera_detection_invert', False)
                if self._settings else False)
        except (AttributeError, TypeError):
            initial_invert = False
        self._invert_var = tk.BooleanVar(value=initial_invert)
        tk.Checkbutton(
            invert_frame, bg=DIALOG_BG,
            text=_("Invert colors (use when scrap is DARKER than the bed)"),
            variable=self._invert_var,
            command=self._on_invert_changed).pack(side='left')

        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 12))
        self._use_btn = tk.Button(
            btn_frame, text=_("Use"), command=self._on_use_clicked,
            font=("Helvetica", 10, "bold"), width=12, state="disabled")
        self._use_btn.pack(side="left", padx=5)
        # Switch camera — fallback for when the auto-resolver picked the
        # wrong device (e.g. integrated laptop webcam). Cycles through
        # enumerated cameras and persists the choice.
        tk.Button(btn_frame, text=_("Switch camera"),
                  command=self._on_switch_camera, width=14
                  ).pack(side="left", padx=5)
        tk.Button(btn_frame, text=_("Cancel"),
                  command=self._on_close, width=10
                  ).pack(side="left", padx=5)

        self.geometry(f"{self.PREVIEW_W + 50}x{self.PREVIEW_H + 240}")
        self.minsize(660, 660)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        self.after(50, self._open_camera_and_start)
        self.wait_window(self)

    def _on_switch_camera(self):
        """Cycle to the next enumerated camera; persist the choice."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        if self._settings is not None:
            self._settings['camera_index_override'] = int(self._camera_index)
            try:
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                pass
        self._status_var.set(
            _("Switched to camera index {i}. Reopening...").format(
                i=self._camera_index))
        # Release current and reopen.
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self.after(50, self._open_camera_and_start)

    def _open_camera_and_start(self):
        """Background-thread camera open so the dialog stays responsive
        during cv2.VideoCapture's 1-3 second initialization."""
        import threading
        self._status_var.set(
            _("Opening camera..."))

        def _worker():
            try:
                cap = self._cam_mod.open_camera(self._camera_index)
                ok, _frame = cap.read()
                if not ok:
                    cap.release()
                    raise RuntimeError(_("Camera opened but returned no frame."))
                self.after(0, self._camera_ready, cap)
            except Exception as e:
                self.after(0, self._camera_failed, e)

        threading.Thread(target=_worker, name='camcap-open',
                          daemon=True).start()

    def _camera_ready(self, cap):
        self._cap = cap
        self._refresh_loop()

    def _camera_failed(self, err):
        self._status_var.set(_("Could not open camera: {e}").format(e=err))

    def _refresh_loop(self):
        if not self.winfo_exists() or self._cap is None:
            return
        ok, frame = self._cap.read()
        if ok and frame is not None:
            self._render_frame_with_overlay(frame)
        self._refresh_after_id = self.after(self.LIVE_REFRESH_MS,
                                              self._refresh_loop)

    def _on_bias_changed(self, _value):
        """Persist the slider value to settings (if available). The
        next live-refresh tick will use the new bias automatically —
        no need to force a re-render here."""
        if self._settings is not None:
            try:
                self._settings['camera_detection_threshold_bias'] = int(
                    self._bias_var.get())
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                pass  # never block the live preview on save failure

    def _on_invert_changed(self):
        """Persist the invert flag — same pattern as the bias slider."""
        if self._settings is not None:
            try:
                self._settings['camera_detection_invert'] = bool(
                    self._invert_var.get())
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                pass

    def _render_frame_with_overlay(self, frame):
        import cv2
        # Undistort + detect contour
        undistorted = self._cam_mod.undistort_frame(frame, self._calibration)
        try:
            bias = int(self._bias_var.get())
        except (AttributeError, tk.TclError, ValueError):
            bias = 0
        try:
            invert = bool(self._invert_var.get())
        except (AttributeError, tk.TclError, ValueError):
            invert = False
        polygon_px = self._cam_mod.detect_scrap_contour(
            undistorted, threshold_bias=bias, invert=invert)
        self._latest_undistorted = undistorted
        self._latest_polygon_px = polygon_px
        if polygon_px:
            self._latest_polygon_mm = self._cam_mod.pixels_to_mm(
                polygon_px, self._calibration)
        else:
            self._latest_polygon_mm = None

        # Overlay the polygon
        annotated = undistorted.copy()
        if polygon_px and len(polygon_px) >= 3:
            import numpy as np
            pts = np.array(polygon_px, dtype=np.int32).reshape(-1, 1, 2)
            cv2.polylines(annotated, [pts], isClosed=True,
                          color=(0, 255, 0), thickness=2)
            self._use_btn.config(state="normal")
            n_pts = len(polygon_px)
            # Compute bbox extent in mm for the status line
            xs_mm = [p[0] for p in self._latest_polygon_mm]
            ys_mm = [p[1] for p in self._latest_polygon_mm]
            w_mm = max(xs_mm) - min(xs_mm)
            h_mm = max(ys_mm) - min(ys_mm)
            self._status_var.set(_(
                "Detected: {n}-vertex polygon, {w:.0f} × {h:.0f} mm bbox").format(
                n=n_pts, w=w_mm, h=h_mm))
        else:
            self._use_btn.config(state="disabled")
            self._status_var.set(_("No scrap detected — place a piece on the bed."))

        # OpenCV BGR → PIL RGB
        rgb = cv2.cvtColor(annotated, cv2.COLOR_BGR2RGB)
        img = self._PIL_Image.fromarray(rgb)
        img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
        self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
        self._preview_label.config(image=self._latest_photo,
                                    width=img.width, height=img.height)

    def _on_use_clicked(self):
        if self._latest_polygon_mm:
            # Return the polygon in ABSOLUTE machine-Y-up coords —
            # do NOT normalize here. Caller (main.py) handles both the
            # normalization and the Y-up→storage-convention conversion
            # so it can also derive the polygon's machine-mm offset
            # (needed for Frame & Cut auto mode). Pre-normalizing here
            # would zero the offset and lose the scrap's bed position.
            self.result_polygon_mm = list(self._latest_polygon_mm)
        self._on_close()

    def _on_close(self):
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self.destroy()


# Grbl 1.1 error / alarm code translations — picked from the official
# Grbl error/alarm tables. Not exhaustive; covers the codes a user is
# most likely to hit in a Frame & Cut workflow. Returns None if we
# don't have a translation (caller shows just the raw text).
_GRBL_ERROR_TEXT = {
    '1': "G-code letter missing or malformed",
    '2': "Numeric value invalid",
    '3': "'$' command not recognized",
    '5': "Homing cycle is not enabled in Grbl ($22=0)",
    '7': "EEPROM read failed — settings restored",
    '8': "'$' command only valid when idle",
    '9': "G-code locked out during alarm or jog",
    '10': "Soft limits cannot be enabled without homing",
    '15': "Travel exceeds machine soft limits",
    '17': "Setting disabled",
    '20': "Unsupported G-code command",
    '22': "Feed rate has not been set",
    '24': "Two G-code commands tried to use the same axis",
    '33': "Motion target invalid",
}
_GRBL_ALARM_TEXT = {
    '1': "Hard limit triggered — machine position lost",
    '2': "G-code motion would exceed soft limits",
    '3': "Reset while in motion — position lost",
    '8': "Homing fail — pull-off didn't clear limit switch",
    '9': "Homing fail — limit switch not contacted in cycle",
}


def _decode_grbl_error(msg):
    """Translate 'error:N' or 'ALARM:N' to a plain-English description."""
    if not msg:
        return None
    parts = msg.strip().split(':', 1)
    if len(parts) != 2:
        return None
    tag, code = parts[0].lower(), parts[1].strip()
    if tag == 'error':
        return _GRBL_ERROR_TEXT.get(code)
    if tag == 'alarm':
        return _GRBL_ALARM_TEXT.get(code)
    return None


class FalconRunDialog(tk.Toplevel):
    """Live-progress dialog while G-code streams to the Falcon.

    Owns a FalconSender for the duration of the job. Callbacks from the
    sender's worker thread are marshalled to the Tk main thread via
    self.after(0, ...). Buttons send real-time commands via the sender.

    Window-close ⨯ behaves as Stop (with confirm if mid-job).
    """

    def __init__(self, parent, sender, gcode_lines, title=None,
                  on_finished=None, loop=False, gcode_provider=None,
                  show_pause_resume=True, show_cut_button=True,
                  stop_needs_confirm=True, done_button_label=None,
                  require_jog_first=False,
                  auto_locate_target=None, is_homed=False):
        super().__init__(parent)
        self.title(title or _("Sending to Falcon"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        # Resizable so a long error message (Grbl code + decoded
        # description + failing line) doesn't get clipped.
        self.resizable(True, True)

        self._sender = sender
        self._gcode_lines = gcode_lines  # for error context
        self._total = len(gcode_lines)
        self._sent = 0
        self._start_time = time.monotonic()
        self._finished = False
        self._on_finished = on_finished
        self._final_reason = None
        self._last_error = None  # preserved across _on_done overwrite
        # Loop mode: restart the stream after each successful completion
        # until the user clicks "Looks Good — Cut!" or "Stop". Used
        # for the framing pass so the user has time to verify
        # alignment (or jog the head) across multiple traces.
        #
        # When ``gcode_provider`` is set, each loop iteration calls
        # it to regenerate fresh G-code instead of re-streaming the
        # same lines. Used by the calibration dialog's "Frame &
        # Engrave Here" flow so the framing trace TRACKS the head as
        # the user jogs between passes (provider re-reads MPos and
        # generates framing centered on the new position).
        self._loop = loop
        self._gcode_provider = gcode_provider
        self._cut_requested = False
        self._pass_count = 0
        # Hide pause/resume in contexts where pause-jog-resume can
        # corrupt state (e.g. calibration framing — absolute coords
        # mean the next G-code line after resume hops back to the
        # pre-jog absolute position regardless of how the user jogged).
        self._show_pause_resume = show_pause_resume
        # The green "Looks Good — Cut!" button only makes sense when
        # the run dialog is followed by a CUT in the same workflow.
        # Calibration framing has Frame and Engrave as separate
        # buttons in the parent dialog, so the in-run Cut button
        # would be misleading there.
        self._show_cut_button = show_cut_button
        # Whether Stop should pop a "are you sure? this can ruin
        # material" confirmation. False for low-power framing where
        # stopping is harmless; True for actual cuts where mid-job
        # abort can burn weird marks into the material.
        self._stop_needs_confirm = stop_needs_confirm
        # Label for the post-completion button (default "Close").
        # Callers that hand off to a follow-up step (e.g. the seed
        # dialog that drives the head to scrap position then yields
        # to a framing loop) override this to "Start Frame →" so the
        # button text describes what happens next, not just that this
        # dialog goes away.
        self._done_button_label = done_button_label or _("Close")
        # MANUAL-mode safety gate: the post-stream done button stays
        # disabled until the user clicks at least one jog button. Used
        # by the "jog to material position" step so the user can't
        # accidentally proceed with the head still at the home corner
        # (where framing/cutting could destroy material). When False
        # (the default), the done button enables immediately on
        # stream completion.
        self._require_jog_first = require_jog_first
        self._has_jogged = not require_jog_first
        # Auto-locate target: if set, show a "Try Auto Locate" button
        # next to Home Laser that drives the head to this machine coord
        # (the polygon's LB vertex). Locked until the laser has been
        # homed in this session — relying on absolute machine coords
        # before homing means MPos drift could send the head to the
        # wrong place. is_homed reflects the parent app's "homed in this
        # session" tracking; the dialog's own Home button updates it
        # when homing succeeds.
        self._auto_locate_target = auto_locate_target
        self._is_homed = is_homed

        # Hook up callbacks. Marshal each to the Tk thread.
        sender.on_status = lambda s: self.after(0, self._on_status, s)
        sender.on_progress = lambda i, n: self.after(0, self._on_progress, i, n)
        sender.on_error = lambda m: self.after(0, self._on_error, m)
        sender.on_alarm = lambda m: self.after(0, self._on_alarm, m)
        sender.on_done = lambda r: self.after(0, self._on_done, r)

        # ---- UI ----
        # Jog-only mode: caller passed an empty gcode list, meaning
        # this dialog is purely a "position the head" step (no actual
        # stream to run). Hide the progress bar / line counter /
        # elapsed timer in that case — they all read 0 forever and
        # confuse users into thinking the dialog is broken.
        self._is_jog_only_mode = (self._total == 0)

        # Wrap long titles so the dialog doesn't stretch the full
        # screen width to fit a one-line instructional sentence.
        tk.Label(self, text=title or _("Sending to Falcon"),
                 bg=DIALOG_BG, font=("Helvetica", 12, "bold"),
                 wraplength=480, justify="center"
                 ).pack(pady=(12, 4), padx=12)

        # Initial state label. For jog-only mode, "Connecting…" is
        # misleading (it transitions to "Complete ✓" the moment the
        # empty stream returns, which Matt observed as confusing). Use
        # a state that describes what the user should DO instead.
        if self._is_jog_only_mode:
            initial_state = _("Position the head, then click {label}.").format(
                label=self._done_button_label)
        else:
            initial_state = _("Connecting ...")
        self._state_var = tk.StringVar(value=initial_state)
        tk.Label(self, textvariable=self._state_var, bg=DIALOG_BG,
                 font=("Helvetica", 11), wraplength=480, justify="center"
                 ).pack(pady=(0, 2), padx=12)

        self._pos_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._pos_var, bg=DIALOG_BG,
                 fg="#555555", font=("Courier", 10)).pack(pady=(0, 6))

        if not self._is_jog_only_mode:
            # Progress bar + line counter + elapsed time only when there's
            # actually something streaming.
            from tkinter import ttk
            self._progress = ttk.Progressbar(self, length=400, mode='determinate',
                                              maximum=self._total)
            self._progress.pack(padx=20, pady=(0, 4))

            self._progress_text_var = tk.StringVar(
                value=_("Line 0 / {n}").format(n=self._total))
            tk.Label(self, textvariable=self._progress_text_var, bg=DIALOG_BG,
                     font=("Helvetica", 9)).pack(pady=(0, 4))

            self._timing_var = tk.StringVar(value=_("Elapsed 0:00"))
            tk.Label(self, textvariable=self._timing_var, bg=DIALOG_BG,
                     fg="#555555", font=("Helvetica", 9)).pack(pady=(0, 10))
        else:
            # Stubs so methods that touch these don't AttributeError.
            self._progress = None
            self._progress_text_var = None
            self._timing_var = None

        # Buttons
        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=(0, 8))
        self._pause_btn = tk.Button(btn_frame, text=_("Pause"),
                                      command=self._on_pause_clicked,
                                      width=10)
        self._resume_btn = tk.Button(btn_frame, text=_("Resume"),
                                       command=self._on_resume_clicked,
                                       width=10, state="disabled")
        if self._show_pause_resume:
            self._pause_btn.pack(side="left", padx=4)
            self._resume_btn.pack(side="left", padx=4)
        # In framing context (loop=True, no cut button), the user's
        # only exit is "Done" — graceful (let the current pass finish
        # so the trailing G0 to bbox-center runs, then close the
        # dialog). In all other contexts the button is "Stop" — abort
        # via soft-reset.
        stop_text = (_("Done") if (loop and not show_cut_button)
                      else _("Stop"))
        stop_bg = "#4caf50" if (loop and not show_cut_button) else None
        stop_fg = "white" if stop_bg else None
        stop_kwargs = {'bg': stop_bg, 'fg': stop_fg,
                        'activebackground': '#388e3c'} if stop_bg else {}
        self._stop_btn = tk.Button(btn_frame, text=stop_text,
                                     command=self._on_stop_clicked,
                                     width=10,
                                     font=("Helvetica", 10, "bold"),
                                     **stop_kwargs)
        self._stop_btn.pack(side="left", padx=4)
        # In jog-only mode (position-the-head dialog), the user needs
        # an explicit Cancel to bail out without proceeding to the
        # next step. Without it, only the X-out (which we now treat
        # as cancel) is available, and the "Start Frame →" button is
        # the only visible action — Matt observed that X-ing out
        # auto-started framing because there was no clearer escape.
        if self._is_jog_only_mode:
            self._cancel_btn = tk.Button(
                btn_frame, text=_("Cancel"),
                command=self._on_cancel_jog_clicked,
                width=10, font=("Helvetica", 10))
            self._cancel_btn.pack(side="left", padx=4)
        # Loop mode shows a prominent "Looks Good — Cut!" button so
        # the user can break out of the repeating framing pass once
        # alignment looks right.
        if self._loop and self._show_cut_button:
            self._cut_btn = tk.Button(
                btn_frame, text=_("Looks Good — Cut!"),
                command=self._on_cut_clicked,
                font=("Helvetica", 10, "bold"),
                bg="#4caf50", fg="white", activebackground="#388e3c")
            self._cut_btn.pack(side="left", padx=8)

        # Jog cluster — useful in G92 mode (the jog moves the cut's
        # reference point) and harmless in auto-framing mode (the cut
        # uses absolute coords so jog only repositions the head for
        # the user's convenience). Sends Grbl $J= via sender.jog(),
        # which Grbl can process even mid-stream.
        jog_frame = tk.LabelFrame(self, text=_("Nudge head"),
                                    bg=DIALOG_BG, padx=8, pady=4)
        jog_frame.pack(pady=(0, 14))
        # Default to 5 mm: on a 400 mm bed, 1 mm is barely
        # perceptible. Users start coarse-jogging into position then
        # switch to 0.5/1 for fine adjustment.
        self._jog_step_var = tk.DoubleVar(value=5.0)
        step_row = tk.Frame(jog_frame, bg=DIALOG_BG)
        step_row.pack(anchor='w')
        tk.Label(step_row, text=_("Step (mm):"), bg=DIALOG_BG,
                  font=("Helvetica", 9)).pack(side='left')
        for step in (0.5, 1.0, 5.0, 25.0):
            tk.Radiobutton(
                step_row, text=str(step), variable=self._jog_step_var,
                value=step, bg=DIALOG_BG, font=("Helvetica", 9),
                highlightthickness=0).pack(side='left')
        btn_grid = tk.Frame(jog_frame, bg=DIALOG_BG)
        btn_grid.pack(pady=2)
        tk.Button(btn_grid, text="↑", width=3,
                   command=lambda: self._on_jog_clicked(0, 1)
                   ).grid(row=0, column=1, padx=2, pady=1)
        tk.Button(btn_grid, text="←", width=3,
                   command=lambda: self._on_jog_clicked(-1, 0)
                   ).grid(row=1, column=0, padx=2, pady=1)
        tk.Button(btn_grid, text="→", width=3,
                   command=lambda: self._on_jog_clicked(1, 0)
                   ).grid(row=1, column=2, padx=2, pady=1)
        tk.Button(btn_grid, text="↓", width=3,
                   command=lambda: self._on_jog_clicked(0, -1)
                   ).grid(row=2, column=1, padx=2, pady=1)
        # Home Laser button — handy when MPos has drifted (lid-open
        # positioning, prior frame stopped mid-trace, etc.) so the user
        # can re-zero without backing out of the dialog. Mid-stream
        # homing aborts the run, so we confirm in that case.
        action_row = tk.Frame(jog_frame, bg=DIALOG_BG)
        action_row.pack(pady=(6, 0))
        self._home_btn = tk.Button(
            action_row, text=_("Home Laser ($H)"),
            command=self._on_home_clicked,
            font=("Helvetica", 9))
        self._home_btn.pack(side='left', padx=2)
        # Try Auto Locate — only shown when caller passed an auto-locate
        # target (a polygon machine coord). Disabled until the laser
        # has been homed in this session.
        self._auto_locate_btn = None
        if self._auto_locate_target is not None:
            self._auto_locate_btn = tk.Button(
                action_row, text=_("Try Auto Locate"),
                command=self._on_auto_locate_clicked,
                font=("Helvetica", 9),
                state=("normal" if self._is_homed else "disabled"))
            self._auto_locate_btn.pack(side='left', padx=2)

        self.protocol("WM_DELETE_WINDOW", self._on_close_window)

        # Tick the timing display once a second
        self._timing_tick()

        # Kick off the stream
        try:
            sender.start_stream(gcode_lines)
            self._state_var.set(_("Running ..."))
        except Exception as e:
            self._state_var.set(_("Failed to start: {e}").format(e=e))
            self._stop_btn.config(text=_("Close"),
                                    command=self._on_close_window)
            self._finished = True

        self.wait_window(self)

    # ------------------------------------------------------------------
    # Tk-thread callbacks (marshalled from FalconSender worker)
    # ------------------------------------------------------------------

    def _on_status(self, status):
        state = status.get('state', '?')
        labels = {
            'Idle': _("Idle"),
            'Run': _("Running"),
            'Hold': _("Paused (feed hold)"),
            'Alarm': _("ALARM"),
            'Door': _("Safety door open"),
            'Jog': _("Jogging"),
            'Home': _("Homing"),
            'Sleep': _("Sleeping"),
        }
        self._state_var.set(labels.get(state, state))
        mpos = status.get('mpos')
        if mpos:
            self._pos_var.set(
                f"X {mpos[0]:7.2f}  Y {mpos[1]:7.2f}  Z {mpos[2]:7.2f}")

    def _on_progress(self, sent, total):
        self._sent = sent
        # Jog-only mode has no progress widgets — skip the updates.
        if self._progress is None or self._progress_text_var is None:
            return
        try:
            self._progress['value'] = sent
        except Exception:
            pass
        pct = (100.0 * sent / total) if total else 0
        self._progress_text_var.set(
            _("Line {s} / {t}  ({pct:.0f}%)").format(s=sent, t=total, pct=pct))

    def _on_error(self, msg):
        self._last_error = msg
        self._state_var.set(_("Error: {m}").format(m=msg))

    def _on_alarm(self, msg):
        self._last_error = msg
        self._state_var.set(_("ALARM: {m}").format(m=msg))

    def _on_done(self, reason):
        # Loop mode: if the pass completed cleanly and the user hasn't
        # clicked "Looks Good — Cut!" yet, restart the stream after
        # making sure Grbl is back in Idle. If we restart while a
        # jog is still finishing (Grbl in JOG state), the next
        # iteration's M3 hits error:9 (G-code locked out during jog).
        if (self._loop and reason == "complete"
                and not self._cut_requested):
            self._pass_count += 1
            self._sent = 0
            self._progress['value'] = 0
            self._state_var.set(
                _("Pass {n} complete — waiting for Idle...").format(
                    n=self._pass_count))
            self._wait_idle_retries = 0
            self.after(300, self._wait_idle_then_restart)
            return
        self._finished = True
        self._final_reason = reason
        # If the user explicitly requested the loop end (clicked Cut
        # or Done in framing context) AND the pass completed cleanly,
        # auto-close the dialog so the caller's next step (the actual
        # cut, or the calibration engrave) starts immediately. Without
        # this, the user has to click Close manually to release
        # wait_window — which looks like "Cut did nothing" because
        # the cut doesn't kick off until this dialog destroys.
        if self._loop and self._cut_requested and reason == "complete":
            if self._on_finished:
                try:
                    self._on_finished(reason)
                except Exception:
                    pass
            self.destroy()
            return
        if reason == "complete":
            self._state_var.set(_("Complete ✓"))
        elif reason == "stopped":
            self._state_var.set(_("Stopped"))
        elif reason == "error":
            # Keep the Grbl error/alarm visible — without this, the
            # specific code (error:5, ALARM:9, etc.) flashes by and the
            # user only sees a generic "Stopped on error".
            err = self._last_error or _("unknown error")
            decoded = _decode_grbl_error(err)
            failing_line = (
                self._gcode_lines[self._sent]
                if 0 <= self._sent < len(self._gcode_lines) else "")
            txt = _("Stopped on error: {e}").format(e=err)
            if decoded:
                txt += _("  — {d}").format(d=decoded)
            if failing_line:
                txt += _("\nFailing line {n}: {l}").format(
                    n=self._sent + 1, l=failing_line)
            self._state_var.set(txt)
        # Switch buttons: only "Close" remains active. Wrap in
        # TclError guards: the worker thread's on_done can fire after
        # the dialog has been destroyed (user closed the window during
        # the small window between stream end and dialog teardown),
        # and configuring a destroyed widget raises "invalid command
        # name" exceptions that surface as unhandled-exception popups.
        try:
            self._pause_btn.config(state="disabled")
        except tk.TclError:
            pass
        try:
            self._resume_btn.config(state="disabled")
        except tk.TclError:
            pass
        try:
            # Post-stream rewire: button now means "advance to the
            # next step" (e.g. Start Frame →) — NOT the same as
            # X-out / Cancel. Use a dedicated handler that destroys
            # without clobbering _final_reason. Routing this through
            # _on_close_window broke Start Frame because that handler
            # marks the dialog as cancelled in jog-only mode (the
            # X-out / Cancel path), which makes main.py return
            # without proceeding to the framing dialog.
            self._stop_btn.config(
                text=self._done_button_label,
                command=self._on_advance_clicked,
                state="normal" if self._has_jogged else "disabled")
            if not self._has_jogged:
                self._state_var.set(
                    _("Jog the head to your material's bottom-left "
                      "corner, then click {label}.").format(
                        label=self._done_button_label))
        except tk.TclError:
            pass
        if self._on_finished:
            try:
                self._on_finished(reason)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _on_pause_clicked(self):
        try:
            self._sender.pause()
        except Exception:
            pass
        self._pause_btn.config(state="disabled")
        self._resume_btn.config(state="normal")

    def _on_resume_clicked(self):
        try:
            self._sender.resume()
        except Exception:
            pass
        self._pause_btn.config(state="normal")
        self._resume_btn.config(state="disabled")

    def _wait_idle_then_restart(self):
        """Poll Grbl's state until it's Idle (any pending jog has
        finished), then restart the stream. Sending the next iter's
        M3 while Grbl is still in JOG state would trigger error:9
        (G-code locked out during alarm or jog)."""
        if self._finished or self._cut_requested:
            return
        # Guard against the dialog being destroyed while a deferred
        # call is in flight — accessing tk widgets after destroy
        # raises TclError.
        try:
            if not self.winfo_exists():
                return
        except tk.TclError:
            return
        try:
            st = self._sender.get_status(timeout=0.2)
        except Exception:
            st = None
        state = (st or {}).get('state')
        if state in ('Idle', None):
            # Settle delay before restart: gives any in-flight jog its
            # ok-byte time to land in the OS serial buffer so
            # start_stream's reset_input_buffer can flush it (without
            # this, a jog issued mid-settle can have its ack arrive
            # after the new worker has started reading, where it's
            # misattributed to a stream-line ack). Empirically 150ms
            # covers Falcon USB latency comfortably.
            self.after(150, self._restart_stream)
            return
        # Still busy (Jog / Run / Hold). Wait a bit and re-check.
        self._wait_idle_retries += 1
        if self._wait_idle_retries > 50:  # ~10 seconds
            try:
                self._state_var.set(
                    _("Timed out waiting for Idle (state={s})").format(s=state))
            except tk.TclError:
                pass
            self._finished = True
            self._final_reason = "error"
            return
        try:
            self.after(200, self._wait_idle_then_restart)
        except tk.TclError:
            pass

    def _restart_stream(self):
        """Re-stream gcode for the next loop iteration. If a
        ``gcode_provider`` callback is set, calls it to get fresh
        gcode (used by the calibration dialog so the framing trace
        follows the head as the user jogs between passes). Bails
        silently if the user has since stopped."""
        if self._finished or self._cut_requested:
            return
        if self._gcode_provider is not None:
            try:
                self._gcode_lines = self._gcode_provider()
                self._total = len(self._gcode_lines)
                self._progress['maximum'] = max(1, self._total)
            except Exception as e:
                self._state_var.set(_("Regenerate failed: {e}").format(e=e))
                self._finished = True
                self._final_reason = "error"
                return
        try:
            self._sender.start_stream(self._gcode_lines)
        except Exception as e:
            self._state_var.set(_("Restart failed: {e}").format(e=e))
            self._finished = True
            self._final_reason = "error"

    def _is_door_open(self):
        """True if Grbl currently reports the safety door is open.

        Reads the cached status (set by the on_status callbacks at
        STATUS_POLL_HZ during streaming) — no synchronous get_status
        call, so this stays non-blocking and doesn't race the
        streamer for the serial port. Cache is at most ~250ms stale
        during active framing, which is fine for a "is the door open
        right now?" confirmation.
        """
        st = self._sender.status or {}
        if st.get('state') == 'Door':
            return True
        # Grbl 1.1 Pn: field reports active input pins. 'D' = door.
        pins = st.get('pins') or ''
        if 'D' in pins:
            return True
        return False

    def _confirm_door_closed(self):
        """If the door is open, pop a confirmation. Returns True if it
        is safe to proceed (door closed OR user explicitly accepted
        the warning), False if the user backed out.

        Used at every "leaving the framing dialog into something that
        fires the laser" boundary — without this, an open lid silently
        wedges the next stream because Grbl blocks motion in Door
        state (no error returned, just an unresponsive controller)."""
        if not self._is_door_open():
            return True
        return messagebox.askyesno(
            _("Lid Appears Open"),
            _("The laser is reporting that the safety lid is open. "
              "Grbl blocks motion in Door state, so the next step "
              "won't fire until the lid is closed.\n\n"
              "Close the lid, then click Yes to continue. Click No "
              "to stay in framing."),
            parent=self)

    def _on_cut_clicked(self):
        """User clicked 'Looks Good — Cut!' during loop mode. Flag the
        request; the current pass will be allowed to finish, then the
        loop exits cleanly with reason='complete'."""
        if not self._confirm_door_closed():
            return
        self._cut_requested = True
        try:
            self._cut_btn.config(state='disabled',
                                  text=_("Finishing pass..."))
        except Exception:
            pass
        self._state_var.set(
            _("Cut requested — finishing current framing pass..."))

    def _on_jog_clicked(self, dx_sign, dy_sign):
        """Send a relative jog via Grbl's ``$J=`` real-time command.
        Works during streaming — Grbl handles jog requests outside
        the planner queue. Useful for nudging the head while a
        framing pass is running so the user can fine-tune alignment.
        """
        try:
            step = float(self._jog_step_var.get())
        except Exception:
            step = 1.0
        try:
            self._sender.jog(x=dx_sign * step, y=dy_sign * step,
                              feed=2000, relative=True)
        except Exception:
            pass
        # If this dialog is gating its done-button on a jog click
        # (MANUAL-mode position confirmation), unlock it now.
        if not self._has_jogged:
            self._has_jogged = True
            if self._finished:
                try:
                    self._stop_btn.config(state="normal")
                    self._state_var.set(
                        _("Head positioned. Click {label} when ready.").format(
                            label=self._done_button_label))
                except tk.TclError:
                    pass

    def _on_home_clicked(self):
        """Home Laser ($H). Reliable way to recover from drifted MPos —
        lid-open positioning, prior frame aborted mid-trace, the head
        bumped into a stop, etc. Runs the home on a worker thread so
        the Tk loop stays responsive (Grbl's $H blocks until the cycle
        finishes — typically 15-30 seconds on a Falcon)."""
        if not self._finished:
            if not messagebox.askyesno(
                    _("Home Laser?"),
                    _("Homing will abort the run in progress. Continue?"),
                    parent=self):
                return
            try:
                self._sender.stop()
            except Exception:
                pass
        import threading
        try:
            self._home_btn.config(state="disabled", text=_("Homing..."))
        except tk.TclError:
            return
        self._state_var.set(_("Homing..."))
        result = {}

        def worker():
            try:
                ok, msg = self._sender.home(timeout_s=60.0)
                result['ok'] = ok
                result['msg'] = msg
            except Exception as exc:
                result['ok'] = False
                result['msg'] = str(exc)

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        def check():
            if t.is_alive():
                try:
                    self.after(200, check)
                except tk.TclError:
                    pass
                return
            try:
                self._home_btn.config(state="normal",
                                       text=_("Home Laser ($H)"))
                if result.get('ok'):
                    self._is_homed = True
                    # Tell the parent app so other dialogs in this
                    # session can also see "homed."
                    parent = self.master
                    if hasattr(parent, '_falcon_homed_this_session'):
                        parent._falcon_homed_this_session = True
                    if self._auto_locate_btn is not None:
                        self._auto_locate_btn.config(state="normal")
                    self._state_var.set(_("Homing complete. MPos is now (0, 0)."))
                else:
                    self._state_var.set(_("Homing failed: {m}").format(
                        m=result.get('msg', '?')))
            except tk.TclError:
                pass

        try:
            self.after(200, check)
        except tk.TclError:
            pass

    def _on_auto_locate_clicked(self):
        """Drive the head to the polygon's LB vertex via an absolute
        G0. Available only after the laser has been homed in this
        session — without homing, MPos drift could put the head far
        from the actual target. Counts as a jog for the require-jog
        gate so users can auto-locate and immediately click Start
        Frame without manually nudging."""
        if self._auto_locate_target is None:
            return
        target_x, target_y = self._auto_locate_target
        # Send via the same channel as jog so it serializes with any
        # in-flight stream activity. Use $J for absolute jog — it's
        # cancellable + tracked separately from stream lines.
        try:
            self._sender.jog(x=target_x, y=target_y,
                              feed=3000, relative=False)
        except Exception as e:
            self._state_var.set(_("Auto-locate failed: {e}").format(e=e))
            return
        self._state_var.set(_("Driving to polygon BL at MPos "
                              "({x:.1f}, {y:.1f}) …").format(
                                x=target_x, y=target_y))
        if not self._has_jogged:
            self._has_jogged = True
            if self._finished:
                try:
                    self._stop_btn.config(state="normal")
                except tk.TclError:
                    pass

    def _on_cancel_jog_clicked(self):
        """Cancel button in the jog-only dialog. Same effect as the
        window X — sets _final_reason to "cancelled" and closes, so
        main.py knows not to proceed to framing."""
        self._final_reason = "cancelled"
        self.destroy()

    def _on_advance_clicked(self):
        """Post-stream 'advance to next step' button (e.g.
        'Start Frame →' on the jog-to-position dialog).
        _final_reason is already 'complete' from _on_done — leave it
        alone and just destroy so the caller proceeds. Routing this
        through _on_close_window would cancel via the jog-only mode
        path."""
        self.destroy()

    def _on_stop_clicked(self):
        if self._finished:
            # Post-stream "Stop" button click is the "advance to next
            # step" path (e.g. the "Start Frame →" button on the jog-
            # to-position dialog). _final_reason is already "complete"
            # from _on_done — leave it alone and just destroy so the
            # caller proceeds to the next step.
            self.destroy()
            return
        # Graceful "Done" path (framing context, no cut button): let
        # the current pass finish — its trailing G0 moves the head
        # to the bbox center — then exit the loop without a soft-
        # reset. Same mechanism as the Cut button uses.
        graceful = self._loop and not self._show_cut_button
        if graceful:
            if not self._confirm_door_closed():
                return
            self._cut_requested = True
            try:
                self._stop_btn.config(state='disabled',
                                       text=_("Finishing pass..."))
            except tk.TclError:
                pass
            self._state_var.set(
                _("Finishing current pass + centering head..."))
            return
        # Aggressive Stop path (cut context): confirm + soft-reset.
        if self._stop_needs_confirm:
            if not messagebox.askyesno(
                    _("Stop?"),
                    _("Stop the job? This will soft-reset the laser "
                      "and clear the planner buffer immediately. "
                      "Material in progress may be ruined."),
                    parent=self):
                return
        try:
            self._sender.stop()
        except Exception:
            pass
        # Sender's on_done will set self._finished and update buttons.

    def _on_close_window(self):
        if not self._finished:
            if not messagebox.askyesno(
                    _("Stop and close?"),
                    _("A job is still running. Stop it and close the window?"),
                    parent=self):
                return
            try:
                self._sender.stop()
            except Exception:
                pass
        # In jog-only mode, the stream completed instantly (empty
        # gcode list), which set _final_reason = "complete" via
        # _on_done. If the user X's out NOW (or clicks Cancel), we
        # want main.py to treat that as a cancellation, not as
        # "Start Frame was clicked." Override the reason so the
        # caller's `_final_reason != "complete"` check fires.
        if self._is_jog_only_mode:
            self._final_reason = "cancelled"
        self.destroy()

    # ------------------------------------------------------------------
    # Timing tick
    # ------------------------------------------------------------------

    def _timing_tick(self):
        if not self.winfo_exists():
            return
        # Jog-only mode skips the timer widget entirely.
        if self._timing_var is None:
            return
        elapsed = int(time.monotonic() - self._start_time)
        em, es = divmod(elapsed, 60)
        if self._finished:
            self._timing_var.set(_("Elapsed {em}:{es:02d}").format(em=em, es=es))
        elif self._sent > 0 and self._total > 0:
            # Simple linear ETA from progress so far
            est_total = elapsed * (self._total / self._sent)
            remaining = max(0, int(est_total - elapsed))
            rm, rs = divmod(remaining, 60)
            self._timing_var.set(
                _("Elapsed {em}:{es:02d}  ETA ~{rm}:{rs:02d}").format(
                    em=em, es=es, rm=rm, rs=rs))
        else:
            self._timing_var.set(_("Elapsed {em}:{es:02d}").format(em=em, es=es))
        if not self._finished:
            self.after(500, self._timing_tick)
        else:
            # One last tick after done to settle the elapsed display
            self.after(1000, self._timing_tick)


class LiveCameraWindow(tk.Toplevel):
    """Non-modal live camera preview, intended to run ALONGSIDE another
    dialog (typically FalconRunDialog while a Frame & Cut job is in flight)
    so the user can watch the laser head moving on the bed.

    Refresh rate is intentionally slow (5 fps) to keep CPU/disk overhead
    minimal while a cut is happening.
    """

    REFRESH_MS = 200      # 5 fps
    PREVIEW_W = 480
    PREVIEW_H = 360

    def __init__(self, parent, camera_index, title=None, settings=None):
        super().__init__(parent)
        self.title(title or _("Live camera"))
        self.configure(bg=DIALOG_BG)
        # Non-modal — explicitly DON'T grab_set or transient-parent, so
        # the FalconRunDialog stays interactive on top of us.
        self._camera_index = camera_index
        self._settings = settings
        self._cap = None
        self._latest_photo = None
        self._refresh_after_id = None
        self._closed = False

        try:
            import camera_capture
            from PIL import Image, ImageTk
            self._cam_mod = camera_capture
            self._PIL_Image = Image
            self._PIL_ImageTk = ImageTk
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
        except ImportError:
            tk.Label(self, text=_("Camera preview unavailable"),
                     bg=DIALOG_BG, padx=15, pady=15).pack()
            return

        # Preview label — match the CameraCaptureDialog pattern (no
        # text in the initial label so width/height are interpreted
        # as the placeholder pixel size, not 480 characters wide).
        self._preview_label = tk.Label(
            self, bg="#222222",
            width=self.PREVIEW_W, height=self.PREVIEW_H)
        self._preview_label.pack(padx=8, pady=(8, 4))

        # Switch camera button below the preview — single-line bar.
        ctrl_row = tk.Frame(self, bg=DIALOG_BG)
        ctrl_row.pack(fill='x', padx=8, pady=(0, 8))
        tk.Button(ctrl_row, text=_("Switch camera"),
                   command=self._on_switch_camera,
                   font=("Helvetica", 9)).pack(side='right')

        self.geometry(f"{self.PREVIEW_W + 30}x{self.PREVIEW_H + 60}")
        # Position to the right of the parent so it doesn't overlap the
        # FalconRunDialog that will open after us.
        try:
            px = parent.winfo_x() + parent.winfo_width() + 20
            py = parent.winfo_y()
            self.geometry(f"+{px}+{py}")
        except Exception:
            pass

        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(50, self._open_camera_async)

    def _on_switch_camera(self):
        """Cycle to the next enumerated camera; persist the choice."""
        try:
            cams = self._cam_mod.enumerate_cameras()
        except Exception:
            cams = []
        if len(cams) < 2:
            from tkinter import messagebox
            messagebox.showinfo(
                _("Only One Camera"),
                _("Only one camera is detected."), parent=self)
            return
        indices = [c['index'] for c in cams]
        try:
            pos = indices.index(self._camera_index)
        except ValueError:
            pos = -1
        self._camera_index = indices[(pos + 1) % len(indices)]
        if self._settings is not None:
            self._settings['camera_index_override'] = int(self._camera_index)
            try:
                from config import save_settings
                save_settings(self._settings)
            except Exception:
                pass
        # Cancel pending refresh, release current cap, reopen.
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        self.after(50, self._open_camera_async)

    def _open_camera_async(self):
        import threading

        def _worker():
            try:
                cap = self._cam_mod.open_camera(self._camera_index)
                self.after(0, self._camera_ready, cap)
            except Exception:
                pass

        threading.Thread(target=_worker, name='livecam-open',
                          daemon=True).start()

    def _camera_ready(self, cap):
        if self._closed:
            try:
                cap.release()
            except Exception:
                pass
            return
        self._cap = cap
        self._refresh_loop()

    def _refresh_loop(self):
        if self._closed or self._cap is None or not self.winfo_exists():
            return
        import cv2
        try:
            ok, frame = self._cap.read()
        except Exception:
            ok = False
        if ok and frame is not None:
            rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = self._PIL_Image.fromarray(rgb)
            img.thumbnail((self.PREVIEW_W, self.PREVIEW_H))
            self._latest_photo = self._PIL_ImageTk.PhotoImage(img)
            self._preview_label.config(
                image=self._latest_photo,
                width=img.width, height=img.height)
        self._refresh_after_id = self.after(self.REFRESH_MS, self._refresh_loop)

    def close(self):
        """Programmatic close — called by the parent when the job is done."""
        self._closed = True
        if self._refresh_after_id is not None:
            try:
                self.after_cancel(self._refresh_after_id)
            except Exception:
                pass
            self._refresh_after_id = None
        if self._cap is not None:
            try:
                self._cap.release()
            except Exception:
                pass
            self._cap = None
        try:
            self.destroy()
        except Exception:
            pass
