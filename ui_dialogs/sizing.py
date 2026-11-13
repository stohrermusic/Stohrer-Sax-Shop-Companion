"""Sizing Rules and its companions: OptionsWindow, the live pad preview, the preset save dialog, Layer Colors.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import copy
import math
import json


from config import (
    DEFAULT_SETTINGS, LIGHTBURN_COLORS, SIZING_PRESET_KEYS, normalize_sizing_preset,
    get_dart_settings_for_size, get_sizing_for_size,
)
from svg_engine import (
    get_disc_diameter, get_felt_thickness_mm, _wave_value,
    dart_wave_raw, DART_TIP_ANGLE,
    locator_mark_geometry, locator_circle_points,
    adhesive_ring_geometry,
)

from .common import DIALOG_BG, IS_MACOS, add_tooltip, add_tooltips, bind_mousewheel, material_display_name
import ui_dialogs as _pkg   # the dialogs suites stand in for are looked up through the package at call time


def grey_value_when_mid(mode_var, *widgets):
    """Disable a placement row's mm entry and unit label while it is set to
    "mid" (mode "centered"), whose position ignores the value, so the number
    doesn't look like it's in use (Matt, 2026-10-28). Follows the mode var,
    so a preset load or a range switch updates it too."""
    def sync(*_args):
        state = "disabled" if mode_var.get() == "centered" else "normal"
        for w in widgets:
            try:
                w.config(state=state)
            except tk.TclError:
                pass  # dialog already gone
    mode_var.trace_add("write", sync)
    sync()

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
        self._dialog_w, self._dialog_h = 500, 750
        self.top.geometry(f"{self._dialog_w}x{self._dialog_h}")
        # Aqua's fonts are wider than Windows'; at 500 px the preset bar's
        # "Save Preset" was clipped on the Mac tour screenshot (2026-10-06).
        # Widen to the content once laid out. after_idle can fire before the
        # window is mapped on macOS and X11 (winfo_width() is still 1 then),
        # so the fit works from the requested size, and runs again on <Map>.
        self.top.after_idle(self._fit_dialog_width)
        self.top.bind("<Map>", lambda e: self._fit_dialog_width(), add="+")
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
        revert_btn = tk.Button(bottom_button_frame, text=_("Revert to Program Defaults"), command=self.revert_to_defaults)
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

        # Labeled zones
        self.zone_labels_enabled_var = tk.BooleanVar(
            value=self.settings.get("zone_labels_enabled", False))
        self.darted_out_from_valley_var = tk.BooleanVar(
            value=self.settings.get("darted_out_from_valley", False))
        self.zone_label_min_var = tk.DoubleVar(
            value=self.settings.get("zone_label_min_size", 7.0))
        self.zone_label_max_var = tk.DoubleVar(
            value=self.settings.get("zone_label_max_size", 12.5))

        # Leather locator marks (lesser-used section)
        self.locator_enabled_var = tk.BooleanVar(
            value=self.settings.get("locator_marks_enabled", False))
        self.locator_min_var = tk.DoubleVar(
            value=self.settings.get("locator_marks_min_size", 7.0))
        self.locator_max_var = tk.DoubleVar(
            value=self.settings.get("locator_marks_max_size", 16.0))
        self.locator_style_var = tk.StringVar(
            value=self.settings.get("locator_marks_style", "lines"))
        self.locator_dashed_var = tk.BooleanVar(
            value=self.settings.get("locator_marks_dashed", False))

        # Adhesive rings (lesser-used section, beta)
        self.ring_enabled_var = tk.BooleanVar(
            value=self.settings.get("adhesive_ring_enabled", False))
        self.ring_trim_var = tk.DoubleVar(
            value=self.settings.get("adhesive_ring_trim_mm", 0.0))
        self.engraving_font_size_vars = {}
        self.engraving_loc_vars = {}
        # (mode_var, mm entry) for every placement row, universal and range
        self.placement_value_entries = []

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

        # Detect whether the starting values match a saved preset, so the
        # dropdown names what's loaded instead of looking empty.
        match = self._detect_active_preset()
        if match is not None:
            self.active_preset_name = match
            self._refresh_sizing_preset_combo(select=match)

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

        rules_frame = tk.LabelFrame(main_frame, text=_("Sizing Rules"), bg=DIALOG_BG, padx=5, pady=5)
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
            mat_lbl = tk.Label(fs_frame, text=material_display_name(mat) + ":", bg=DIALOG_BG)
            mat_lbl.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            fvar = tk.DoubleVar(value=self.settings["engraving_font_size"].get(mat, 2.0))
            self.engraving_font_size_vars[mat] = fvar
            mat_ent = tk.Entry(fs_frame, textvariable=fvar, width=8)
            mat_ent.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            add_tooltips(
                _("Engraving font size in millimeters for {material} discs.").format(
                    material=material_display_name(mat).lower()),
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
            mat_lbl = tk.Label(esr_fs, text=material_display_name(mat) + ":", bg=DIALOG_BG)
            mat_lbl.grid(row=i, column=0, sticky='w', padx=5, pady=2)
            fvar = tk.DoubleVar(value=2.0)
            self.eng_settings_range_font_vars[mat] = fvar
            mat_ent = tk.Entry(esr_fs, textvariable=fvar, width=8)
            mat_ent.grid(row=i, column=1, sticky='w', padx=5, pady=2)
            add_tooltips(
                _("Engraving font size for {material} pads in this range.").format(
                    material=material_display_name(mat).lower()),
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

        # Darted leather: "out" from the dart valley rather than the cut edge.
        valley_cb = tk.Checkbutton(
            engraving_loc_frame,
            text=_("Darted leather: measure \"out\" from the dart valley"),
            variable=self.darted_out_from_valley_var, bg=DIALOG_BG)
        valley_cb.pack(anchor="w", pady=(0, 4))
        add_tooltip(valley_cb, _(
            "On darted leather pads the dart valley sits inside the cut edge, "
            "and deeper on bigger pads. With this on, the darted leather \"out\" "
            "value is the gap between the valley and the top of the number "
            "(0 = touching), so one value keeps the label just inside the "
            "darts at every size. Off: measured from the cut edge to the "
            "middle of the number, as before."))

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
            "out = distance measured inward from the outer edge "
            "(for darted leather with that box ticked: the gap between the "
            "dart valley and the top of the number).\n"
            "mid = halfway between the center hole (or the center, if "
            "the pad has no hole) and the outer edge. The mm value is "
            "not used.\n"
            "int = distance measured outward from the center hole "
            "(from the center, if the pad has no hole)."
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
            # Order out / mid / int: outer edge to centre (Matt, 2026-10-28).
            rb_out = tk.Radiobutton(frame, text=_("out"), variable=mode_var, value="from_outside", bg=DIALOG_BG)
            rb_out.pack(side="left")
            rb_ctr = tk.Radiobutton(frame, text=_("mid"), variable=mode_var, value="centered", bg=DIALOG_BG)
            rb_ctr.pack(side="left")
            # "int" (not "in") so the msgid doesn't collide with the
            # inches unit "in" used in tooling/felt-thickness pickers.
            rb_in = tk.Radiobutton(frame, text=_("int"), variable=mode_var, value="from_inside", bg=DIALOG_BG)
            rb_in.pack(side="left")
            val_ent = tk.Entry(frame, textvariable=val_var, width=5)
            val_ent.pack(side="left", padx=5)
            mm_lbl = tk.Label(frame, text=_("mm"), bg=DIALOG_BG)
            mm_lbl.pack(side="left")
            grey_value_when_mid(mode_var, val_ent, mm_lbl)
            self.placement_value_entries.append((mode_var, val_ent))
            add_tooltip(row_lbl, placement_help[mat])
            add_tooltips(mode_help, rb_out, rb_in, rb_ctr)
            add_tooltips(
                _("Distance in millimeters from the chosen reference "
                "(out or int; mid does not use it)."),
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
            rb_ctr = tk.Radiobutton(frame, text=_("mid"), variable=mode_var, value="centered", bg=DIALOG_BG)
            rb_ctr.pack(side="left")
            rb_in = tk.Radiobutton(frame, text=_("int"), variable=mode_var, value="from_inside", bg=DIALOG_BG)
            rb_in.pack(side="left")
            val_ent = tk.Entry(frame, textvariable=val_var, width=5)
            val_ent.pack(side="left", padx=5)
            mm_lbl = tk.Label(frame, text=_("mm"), bg=DIALOG_BG)
            mm_lbl.pack(side="left")
            grey_value_when_mid(mode_var, val_ent, mm_lbl)
            self.placement_value_entries.append((mode_var, val_ent))
            add_tooltip(row_lbl, placement_help[mat] + " (this size range)")
            add_tooltips(mode_help, rb_out, rb_in, rb_ctr)
            add_tooltips(
                _("Distance from the chosen reference, in millimeters "
                "(out or int; mid does not use it)."),
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

        # --- Lesser-used settings (collapsed) ---------------------------
        # Labeled zones, leather locator marks and export compatibility live
        # behind one header so the dialog stays about pad geometry. The
        # section opens itself when anything inside is non-default, so an
        # active setting is never out of sight.
        self.lesser_open = False
        lesser_frame = tk.Frame(main_frame, bg=DIALOG_BG)
        lesser_frame.pack(fill="x", pady=5)
        self.lesser_btn = tk.Button(lesser_frame, text="", anchor="w", relief="flat",
                                    bg=DIALOG_BG, command=self._toggle_lesser_used)
        self.lesser_btn.pack(fill="x")
        add_tooltip(self.lesser_btn,
                    _("Settings most people never touch: labeled zones, leather "
                      "locator marks, adhesive rings and SVG export compatibility."))
        self.lesser_body = tk.Frame(lesser_frame, bg=DIALOG_BG)

        # --- Labeled Zones ---------------------------------------------
        zone_frame = tk.LabelFrame(self.lesser_body, text=_("Labeled Zones"), bg=DIALOG_BG, padx=5, pady=5)
        zone_frame.pack(fill="x", pady=5)

        zone_cb = tk.Checkbutton(
            zone_frame, text=_("Group same-size pads into labeled zones"),
            variable=self.zone_labels_enabled_var, bg=DIALOG_BG,
            command=self._toggle_zone_fields)
        zone_cb.pack(anchor='w')
        add_tooltip(zone_cb,
                    _("Cut small pads in bordered blocks, one per size, with "
                      "the size engraved beside them. Small discs look alike "
                      "and their own number is often too small to read, so "
                      "the label goes on the sheet instead of the part."))

        zone_row = tk.Frame(zone_frame, bg=DIALOG_BG)
        zone_row.pack(anchor='w', pady=(4, 0))
        self.zone_range_label = tk.Label(zone_row, text=_("Applies to pad sizes"), bg=DIALOG_BG)
        self.zone_range_label.pack(side="left")
        self.zone_min_entry = tk.Entry(zone_row, textvariable=self.zone_label_min_var, width=6)
        self.zone_min_entry.pack(side="left", padx=4)
        self.zone_to_label = tk.Label(zone_row, text=_("to"), bg=DIALOG_BG)
        self.zone_to_label.pack(side="left")
        self.zone_max_entry = tk.Entry(zone_row, textvariable=self.zone_label_max_var, width=6)
        self.zone_max_entry.pack(side="left", padx=4)
        self.zone_mm_label = tk.Label(zone_row, text=_("mm"), bg=DIALOG_BG)
        self.zone_mm_label.pack(side="left")
        add_tooltips(_("Only pads in this range are grouped into zones. "
                       "Larger pads nest normally at full density."),
                     self.zone_range_label, self.zone_min_entry,
                     self.zone_to_label, self.zone_max_entry)

        tk.Label(zone_frame,
                 text=_("Zones use a grid, so they may increase material wastage."),
                 bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666",
                 justify="left").pack(anchor='w', pady=(3, 0))

        self._toggle_zone_fields()

        locator_frame = tk.LabelFrame(self.lesser_body, text=_("Leather Locator Marks"),
                                      bg=DIALOG_BG, padx=5, pady=5)
        locator_frame.pack(fill="x", pady=5)
        loc_cb = tk.Checkbutton(
            locator_frame, text=_("Engrave felt locator marks on leather pads"),
            variable=self.locator_enabled_var, bg=DIALOG_BG,
            command=self._toggle_locator_fields)
        loc_cb.pack(anchor='w')
        add_tooltip(loc_cb,
                    _("Marks on the leather showing where the felt goes, for "
                      "centering it on pads that have no center hole. They land "
                      "on the side facing up on the laser, so cut the leather "
                      "flesh-side (fuzzy side) up when using them."))

        loc_row = tk.Frame(locator_frame, bg=DIALOG_BG)
        loc_row.pack(anchor='w', pady=(4, 0))
        loc_range_label = tk.Label(loc_row, text=_("Applies to pad sizes"), bg=DIALOG_BG)
        loc_range_label.pack(side="left")
        self.locator_min_entry = tk.Entry(loc_row, textvariable=self.locator_min_var, width=6)
        self.locator_min_entry.pack(side="left", padx=4)
        loc_to_label = tk.Label(loc_row, text=_("to"), bg=DIALOG_BG)
        loc_to_label.pack(side="left")
        self.locator_max_entry = tk.Entry(loc_row, textvariable=self.locator_max_var, width=6)
        self.locator_max_entry.pack(side="left", padx=4)
        tk.Label(loc_row, text=_("mm"), bg=DIALOG_BG).pack(side="left")
        add_tooltips(_("Only leather pads in this range get marks. The default "
                       "covers the sizes that get no center hole under the "
                       "default sizing rules."),
                     loc_range_label, self.locator_min_entry, loc_to_label, self.locator_max_entry)

        style_row = tk.Frame(locator_frame, bg=DIALOG_BG)
        style_row.pack(anchor='w', pady=(4, 0))
        tk.Label(style_row, text=_("Style:"), bg=DIALOG_BG).pack(side="left")
        self.locator_style_radios = []
        for label, value, tip in (
                (_("Four lines"), "lines",
                 _("Four lines from near the edge in to where the felt edge "
                   "should be. Centered when the gaps look equal all round.")),
                (_("Circle"), "circle",
                 _("A circle the size of the felt disc.")),
                (_("Both"), "both",
                 _("The circle and the four lines together."))):
            rb = tk.Radiobutton(style_row, text=label, value=value,
                                variable=self.locator_style_var, bg=DIALOG_BG)
            rb.pack(side="left", padx=(6, 0))
            add_tooltip(rb, tip)
            self.locator_style_radios.append(rb)
        self.locator_dashed_cb = tk.Checkbutton(style_row, text=_("Dashed"),
                                                variable=self.locator_dashed_var, bg=DIALOG_BG)
        self.locator_dashed_cb.pack(side="left", padx=(12, 0))
        add_tooltip(self.locator_dashed_cb,
                    _("Draw the marks as short dashes instead of solid lines."))

        tk.Label(locator_frame,
                 text=_("SVG output gives the marks their own layer (Options > Layer Colors). "
                        "G-code engraves them at the leather engraving settings, so use "
                        "SVG if they need their own power."),
                 bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666",
                 justify="left", wraplength=420).pack(anchor='w', pady=(3, 0))

        # --- Adhesive rings (beta) -------------------------------------
        # An opt-in fifth Pad Maker material: a ring of double-sided tape
        # per pad. Its rules (on/off and the width trim) live here; the
        # material checkbox, its layer color, its G-code section and its
        # framing power only appear while this is ticked.
        ring_frame = tk.LabelFrame(self.lesser_body, text=_("Adhesive Rings (beta)"),
                                   bg=DIALOG_BG, padx=5, pady=5)
        ring_frame.pack(fill="x", pady=5)
        self.ring_frame = ring_frame
        self.ring_cb = tk.Checkbutton(
            ring_frame, text=_("Add adhesive rings as a material"),
            variable=self.ring_enabled_var, bg=DIALOG_BG,
            command=self._toggle_ring_fields)
        self.ring_cb.pack(anchor='w')
        add_tooltip(self.ring_cb,
                    _("Rings of double-sided tape, one per pad, for sticking a "
                      "silicone wrap to the back of the card. Adds an Adhesive "
                      "rings material to the Pad Maker, with its own layer "
                      "color and G-code settings. Beta."))

        # "Wrap clearance" (Matt's name, 2026-11-05) is the stored
        # adhesive_ring_trim_mm, unchanged: mm of ring width per side,
        # positive = the ring's inner edge reaches past the edge of the
        # wrap (extra ring showing), negative = it stays hidden under the
        # wrap. A counter in 0.1 steps; typing is still allowed.
        ring_row = tk.Frame(ring_frame, bg=DIALOG_BG)
        ring_row.pack(anchor='w', pady=(4, 0))
        self.ring_clearance_label = tk.Label(ring_row, text=_("Wrap clearance:"), bg=DIALOG_BG)
        self.ring_clearance_label.pack(side="left")
        # Tk reformats and clamps the variable when a spinbox is created on
        # it (a stored 0.25 would become 0.2 just by opening the dialog), so
        # the stored value is put back afterwards, exactly as it was saved.
        stored_clearance = self.ring_trim_var.get()
        self.ring_clearance_spin = tk.Spinbox(
            ring_row, textvariable=self.ring_trim_var, width=6,
            from_=self.RING_CLEARANCE_MIN, to=self.RING_CLEARANCE_MAX,
            increment=self.RING_CLEARANCE_STEP, format="%.1f")
        self.ring_trim_var.set(stored_clearance)
        self.ring_clearance_spin.pack(side="left", padx=4)
        # Typed text isn't checked by Tk; pull it back into range when the
        # field is left. (Apply clamps too, see _ring_clearance_value.)
        self.ring_clearance_spin.bind("<FocusOut>", lambda _e: self._clamp_ring_clearance(), add="+")
        self.ring_clearance_spin.bind("<Return>", lambda _e: self._clamp_ring_clearance(), add="+")
        self.ring_clearance_unit = tk.Label(ring_row, text=_("mm"), bg=DIALOG_BG)
        self.ring_clearance_unit.pack(side="left")
        self.ring_clearance_tips = add_tooltips(
            _("How far the ring's inner edge reaches past the edge of the "
              "wrap. Positive leaves extra ring showing. Negative keeps the "
              "ring's edge hidden under the wrap."),
            self.ring_clearance_label, self.ring_clearance_spin, self.ring_clearance_unit)

        tk.Label(ring_frame,
                 text=_("The outer edge matches the card disc. Each ring's size is "
                        "cut in its hole, or just under the ring when the hole is "
                        "too small to read it."),
                 bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666",
                 justify="left", wraplength=420).pack(anchor='w', pady=(3, 0))
        self._toggle_ring_fields()

        export_frame = tk.LabelFrame(self.lesser_body, text=_("Export Settings"), bg=DIALOG_BG, padx=5, pady=5)
        export_frame.pack(fill="x", pady=5)
        compat_cb = tk.Checkbutton(export_frame,
                                   text=_("Enable Inkscape/Compatibility Mode (unitless SVG)"),
                                   variable=self.compatibility_mode_var, bg=DIALOG_BG)
        compat_cb.pack(anchor='w')
        add_tooltip(compat_cb,
                    _("Write SVGs without explicit unit attributes. Turn on "
                    "if Inkscape (or other software) misinterprets the file "
                    "scale. LightBurn does not need this."))

        self.locator_widgets = [self.locator_min_entry, self.locator_max_entry,
                                self.locator_dashed_cb] + self.locator_style_radios
        self._toggle_locator_fields()

        if (self.zone_labels_enabled_var.get() or self.locator_enabled_var.get()
                or self.compatibility_mode_var.get() or self.ring_enabled_var.get()):
            self._toggle_lesser_used()
        else:
            self._refresh_lesser_header()

    def _fit_dialog_width(self):
        """Widen the dialog so the scrollable content's widest row fits."""
        try:
            self.top.update_idletasks()
            need = self.scrollable_frame.winfo_reqwidth() + self.scrollbar.winfo_reqwidth() + 8
            if need > self._dialog_w:
                self._dialog_w = int(need)
                h = self.top.winfo_height()
                self.top.geometry(f"{self._dialog_w}x{h if h > 1 else self._dialog_h}")
        except tk.TclError:
            pass

    def _toggle_lesser_used(self):
        """Expand or collapse the lesser-used settings section."""
        self.lesser_open = not self.lesser_open
        if self.lesser_open:
            self.lesser_body.pack(fill="x", pady=(2, 0))
        else:
            self.lesser_body.pack_forget()
        self._refresh_lesser_header()

    def _refresh_lesser_header(self):
        arrow = "\u25bc" if self.lesser_open else "\u25b6"
        self.lesser_btn.config(text=f"{arrow} " + _("Lesser-used settings"))

    # Wrap clearance counter (the stored adhesive_ring_trim_mm), in mm.
    RING_CLEARANCE_MIN = -5.0
    RING_CLEARANCE_MAX = 5.0
    RING_CLEARANCE_STEP = 0.1

    def _toggle_ring_fields(self):
        """Grey the wrap clearance counter when adhesive rings are off."""
        state = "normal" if self.ring_enabled_var.get() else "disabled"
        self.ring_clearance_spin.config(state=state)

    def _ring_clearance_value(self):
        """The wrap clearance as it will be stored: what the counter shows,
        held to its range. Raises tk.TclError while the field isn't a
        number, like every other numeric field (see _form_is_valid)."""
        value = self.ring_trim_var.get()
        return max(self.RING_CLEARANCE_MIN, min(self.RING_CLEARANCE_MAX, value))

    def _clamp_ring_clearance(self):
        """Show a typed out-of-range clearance as the value that will be used."""
        try:
            value = self.ring_trim_var.get()
        except tk.TclError:
            return      # not a number yet; Apply says so
        clamped = max(self.RING_CLEARANCE_MIN, min(self.RING_CLEARANCE_MAX, value))
        if clamped != value:
            self.ring_trim_var.set(clamped)

    def _toggle_locator_fields(self):
        """Grey the locator options when the marks are off."""
        state = "normal" if self.locator_enabled_var.get() else "disabled"
        for w in self.locator_widgets:
            w.config(state=state)

    def _toggle_zone_fields(self):
        """Grey the zone size range when zones are off."""
        state = "normal" if self.zone_labels_enabled_var.get() else "disabled"
        for w in (self.zone_min_entry, self.zone_max_entry):
            w.config(state=state)

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
        self.settings["darted_out_from_valley"] = self.darted_out_from_valley_var.get()

        # STAR ENGRAVING SAVE
        self.settings["dart_engraving_on"] = self.dart_engraving_on_var.get()
            
        # Labeled zones
        self.settings["zone_labels_enabled"] = self.zone_labels_enabled_var.get()
        self.settings["zone_label_min_size"] = self.zone_label_min_var.get()
        self.settings["zone_label_max_size"] = self.zone_label_max_var.get()

        # Export
        self.settings["compatibility_mode"] = self.compatibility_mode_var.get()

        # Leather locator marks
        self.settings["locator_marks_enabled"] = self.locator_enabled_var.get()
        self.settings["locator_marks_min_size"] = self.locator_min_var.get()
        self.settings["locator_marks_max_size"] = self.locator_max_var.get()
        self.settings["locator_marks_style"] = self.locator_style_var.get()
        self.settings["locator_marks_dashed"] = self.locator_dashed_var.get()

        # Adhesive rings
        self.settings["adhesive_ring_enabled"] = self.ring_enabled_var.get()
        self.settings["adhesive_ring_trim_mm"] = self._ring_clearance_value()

        self.save_callback()
        self.update_callback()
        self.top.destroy()

    def revert_to_defaults(self):
        if messagebox.askyesno(_("Revert to Program Defaults"),
                               _("Reset every value in this dialog to the program's original defaults? "
                                 "Your saved presets are not touched.")):
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
            self.darted_out_from_valley_var.set(DEFAULT_SETTINGS["darted_out_from_valley"])

            # Labeled zones
            self.zone_labels_enabled_var.set(DEFAULT_SETTINGS["zone_labels_enabled"])
            self.zone_label_min_var.set(DEFAULT_SETTINGS["zone_label_min_size"])
            self.zone_label_max_var.set(DEFAULT_SETTINGS["zone_label_max_size"])
            self._toggle_zone_fields()

            # Export
            self.compatibility_mode_var.set(DEFAULT_SETTINGS.get("compatibility_mode", False))

            # Leather locator marks
            self.locator_enabled_var.set(DEFAULT_SETTINGS["locator_marks_enabled"])
            self.locator_min_var.set(DEFAULT_SETTINGS["locator_marks_min_size"])
            self.locator_max_var.set(DEFAULT_SETTINGS["locator_marks_max_size"])
            self.locator_style_var.set(DEFAULT_SETTINGS["locator_marks_style"])
            self.locator_dashed_var.set(DEFAULT_SETTINGS["locator_marks_dashed"])
            self._toggle_locator_fields()

            # Adhesive rings
            self.ring_enabled_var.set(DEFAULT_SETTINGS["adhesive_ring_enabled"])
            self.ring_trim_var.set(DEFAULT_SETTINGS["adhesive_ring_trim_mm"])
            self._toggle_ring_fields()

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
            "darted_out_from_valley": self.darted_out_from_valley_var.get(),
            # Labeled zones
            "zone_labels_enabled": self.zone_labels_enabled_var.get(),
            "zone_label_min_size": self.zone_label_min_var.get(),
            "zone_label_max_size": self.zone_label_max_var.get(),
            # Export
            "compatibility_mode": self.compatibility_mode_var.get(),
            # Leather locator marks
            "locator_marks_enabled": self.locator_enabled_var.get(),
            "locator_marks_min_size": self.locator_min_var.get(),
            "locator_marks_max_size": self.locator_max_var.get(),
            "locator_marks_style": self.locator_style_var.get(),
            "locator_marks_dashed": self.locator_dashed_var.get(),
            # Adhesive rings
            "adhesive_ring_enabled": self.ring_enabled_var.get(),
            "adhesive_ring_trim_mm": self._ring_clearance_value(),
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
        self.darted_out_from_valley_var.set(
            source.get("darted_out_from_valley", d["darted_out_from_valley"]))

        # Labeled zones
        self.zone_labels_enabled_var.set(source.get("zone_labels_enabled", d["zone_labels_enabled"]))
        self.zone_label_min_var.set(source.get("zone_label_min_size", d["zone_label_min_size"]))
        self.zone_label_max_var.set(source.get("zone_label_max_size", d["zone_label_max_size"]))
        self._toggle_zone_fields()

        # Export
        self.compatibility_mode_var.set(source.get("compatibility_mode", d.get("compatibility_mode", False)))

        # Leather locator marks
        self.locator_enabled_var.set(source.get("locator_marks_enabled", d["locator_marks_enabled"]))
        self.locator_min_var.set(source.get("locator_marks_min_size", d["locator_marks_min_size"]))
        self.locator_max_var.set(source.get("locator_marks_max_size", d["locator_marks_max_size"]))
        self.locator_style_var.set(source.get("locator_marks_style", d["locator_marks_style"]))
        self.locator_dashed_var.set(source.get("locator_marks_dashed", d["locator_marks_dashed"]))
        self._toggle_locator_fields()

        # Adhesive rings
        self.ring_enabled_var.set(source.get("adhesive_ring_enabled", d["adhesive_ring_enabled"]))
        self.ring_trim_var.set(source.get("adhesive_ring_trim_mm", d["adhesive_ring_trim_mm"]))
        self._toggle_ring_fields()

    def _detect_active_preset(self):
        """Find a saved preset whose values match the form's current snapshot.

        Applied settings always come from a preset — Apply refuses to commit
        edits that aren't captured in one — so on open this names the preset
        the user is actually working in. Returns the name, or None when the
        values match nothing saved (e.g. a hand-edited config).
        """
        if not self.sizing_presets:
            return None
        try:
            snapshot = self._capture_form_to_dict()
        except tk.TclError:
            return None
        # Sorted so an exact tie between two identical presets resolves the
        # same way every open, and matches the dropdown's own ordering.
        for name in sorted(self.sizing_presets):
            # Normalize first: a preset saved before a key was added to
            # SIZING_PRESET_KEYS is missing it, and would never match the
            # snapshot on exact equality.
            if normalize_sizing_preset(self.sizing_presets[name]) == snapshot:
                return name
        return None

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
        # Normalize so an older preset fills its missing keys with defaults
        # rather than leaving whatever was in the form from before.
        self._apply_dict_to_form(normalize_sizing_preset(self.sizing_presets[name]))
        self.active_preset_name = name
        self._set_baseline_to_current()

    def on_save_sizing_preset(self):
        """Open the Save Preset dialog (overwrite existing or save as new)."""
        if not self._form_is_valid(_("save it as a preset")):
            return
        dlg = _pkg.SaveSizingPresetDialog(
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
            from config import read_json
            data = read_json(path)
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
        'adhesive_ring': '#B02FB0',
    }
    LABELS = {
        'leather':    _('Leather'),
        'felt':       _('Felt'),
        'card':       _('Card'),
        'exact_size': _('Exact size'),
        'adhesive_ring': _('Adhesive rings'),
    }
    MATERIALS = ('leather', 'felt', 'card', 'exact_size', 'adhesive_ring')
    # Adhesive rings are offered only while the parent form has the
    # feature ticked (the checkbox comes and goes with it, live).
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
            'adhesive_ring': tk.BooleanVar(value=True),
        }
        self._mat_checkbuttons = {}
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
            self._mat_checkbuttons[mat] = cb
            if mat != 'adhesive_ring':
                cb.pack(side="left", padx=2)
        self._sync_ring_checkbox(self._ring_enabled_in_form())

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

    def _ring_enabled_in_form(self):
        """Whether the parent form has adhesive rings ticked right now."""
        try:
            return bool(self.parent_options.ring_enabled_var.get())
        except (tk.TclError, AttributeError):
            return False

    def _sync_ring_checkbox(self, enabled):
        """Show the Adhesive rings checkbox only while the feature is on."""
        cb = self._mat_checkbuttons.get('adhesive_ring')
        if cb is None:
            return
        shown = bool(cb.winfo_manager())
        if enabled and not shown:
            cb.pack(side="left", padx=2)
        elif shown and not enabled:
            cb.pack_forget()

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

        ring_on = bool(settings.get("adhesive_ring_enabled", False))
        self._sync_ring_checkbox(ring_on)
        materials = [m for m in self.MATERIALS if self.show_vars[m].get()
                     and (m != 'adhesive_ring' or ring_on)]
        if not materials:
            self._draw_message("No materials selected")
            return

        # Compute disc diameters for each requested material.
        diams = {}
        for mat in materials:
            try:
                if mat == 'adhesive_ring':
                    # The ring's cut, not its nesting footprint.
                    d = adhesive_ring_geometry(pad_size, settings)["outer_r"] * 2
                else:
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
            elif mat == 'adhesive_ring':
                self._draw_adhesive_ring(cx, cy, pad_size, settings, scale, color)
                continue   # a ring has its own hole, not the pad's center hole
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
        if 'adhesive_ring' in diams:
            # A number under a small ring hangs below the row of discs; leave
            # it (and the caption under it) room, or it runs off the canvas.
            ring = adhesive_ring_geometry(pad_size, settings)
            max_h_mm = max(max_h_mm, 2 * ring["box"][3])
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
            label_drop = 0.0
            if mat == 'leather':
                self._draw_leather(cx_px, cy, pad_size, settings, r_px, color, scale)
                self._draw_center_hole(cx_px, cy, mat, pad_size, settings, scale, color)
            elif mat == 'adhesive_ring':
                geom = self._draw_adhesive_ring(cx_px, cy, pad_size, settings, scale, color)
                # Keep the caption clear of a number that sits under the ring.
                if geom["placement"] == "outside":
                    label_drop = (geom["box"][3] - geom["outer_r"]) * scale + 12
            else:
                self.canvas.create_oval(
                    cx_px - r_px, cy - r_px, cx_px + r_px, cy + r_px,
                    outline=color, width=2, fill='',
                )
                self._draw_center_hole(cx_px, cy, mat, pad_size, settings, scale, color)
            self.canvas.create_text(
                cx_px, cy + r_px + 14 + label_drop,
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
            self._draw_locator(cx, cy, pad_size, settings, outer_r_px / scale, scale)
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
            # Canvas is Y-down like the SVG: a tip at 12 o'clock.
            raw = dart_wave_raw(num_points, theta, DART_TIP_ANGLE)
            shaped = _wave_value(raw, shape_factor)
            r = avg_r_px + amp_px * shaped
            x = cx + r * math.cos(theta)
            y = cy + r * math.sin(theta)
            coords.extend([x, y])
        # Use create_line on a closed loop (polygon would auto-fill its
        # interior, smoothing the perceived shape; we want crisp outlines).
        self.canvas.create_line(*coords, fill=color, width=2, smooth=False)
        self._draw_locator(cx, cy, pad_size, settings, outer_r_mm, scale)

    # --- Adhesive ring (same geometry as both engines) ---

    def _draw_adhesive_ring(self, cx, cy, pad_size, settings, scale, color):
        """Outer cut, inner cut and the number where it will really go
        (in the hole, or under the ring). Returns the geometry drawn."""
        geom = adhesive_ring_geometry(pad_size, settings)
        ro = geom["outer_r"] * scale
        self.canvas.create_oval(cx - ro, cy - ro, cx + ro, cy + ro,
                                outline=color, width=2, fill='', tags=("ring",))
        if geom["inner_r"] > 0:
            ri = geom["inner_r"] * scale
            self.canvas.create_oval(cx - ri, cy - ri, cx + ri, cy + ri,
                                    outline=color, width=2, fill='', tags=("ring", "ring_inner"))
        for stroke in geom["strokes"]:
            pts = [c for x, y in stroke for c in (cx + x * scale, cy + y * scale)]
            self.canvas.create_line(*pts, fill=color, width=1, tags=("ring_number",))
        return geom

    # --- Leather locator marks (same geometry as both engines) ---

    LOCATOR_COLOR = '#1A9E9E'

    def _draw_locator(self, cx, cy, pad_size, settings, outer_r_mm, scale):
        loc = locator_mark_geometry(pad_size, outer_r_mm, settings)
        if loc is None:
            return
        for lr in loc["circles"]:
            if loc["dashed"]:
                for arc in locator_circle_points(lr, True):
                    pts = [c for x, y in arc for c in (cx + x * scale, cy + y * scale)]
                    self.canvas.create_line(*pts, fill=self.LOCATOR_COLOR, width=1, tags=("locator",))
            else:
                rp = lr * scale
                self.canvas.create_oval(cx - rp, cy - rp, cx + rp, cy + rp,
                                        outline=self.LOCATOR_COLOR, width=1, tags=("locator",))
        for (x0, y0), (x1, y1) in loc["segments"]:
            self.canvas.create_line(cx + x0 * scale, cy + y0 * scale,
                                    cx + x1 * scale, cy + y1 * scale,
                                    fill=self.LOCATOR_COLOR, width=1, tags=("locator",))

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
        self.top.geometry("450x450")
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
            'leather_outline', 'leather_center_hole', 'leather_engraving', 'leather_locator',
            'exact_size_outline', 'exact_size_center_hole', 'exact_size_engraving'
        ]
        # Adhesive rings: one layer for the ring and its number, listed
        # only while the feature is on (Sizing Rules > Lesser-used settings).
        if self.settings.get("adhesive_ring_enabled", False):
            layer_map_keys.append('adhesive_ring_cut')

        # Row labels in the user's language: the material as the Materials
        # checkboxes name it, the step as a translated phrase (until
        # 2026-11-10 the label was the key with its underscores swapped for
        # spaces, "Leather locator", in every language).
        step_label = {"outline": _("{material} outline"),
                      "center_hole": _("{material} center hole"),
                      "engraving": _("{material} engraving"),
                      "locator": _("{material} locator marks"),
                      "cut": _("{material} cut")}
        for i, key in enumerate(layer_map_keys):
            mat = next((m for m in ('exact_size', 'adhesive_ring', 'felt', 'card', 'leather')
                        if key.startswith(m + '_')), key.split('_', 1)[0])
            step = key[len(mat) + 1:]
            phrase = step_label.get(step, "{material} " + step.replace('_', ' '))
            label_text = phrase.format(material=material_display_name(mat)) + ":"
            row_lbl = tk.Label(main_frame, text=label_text, bg=DIALOG_BG)
            row_lbl.grid(row=i, column=0, sticky='w', pady=3)

            var = tk.StringVar()
            current_hex = self.settings["layer_colors"].get(key, "#000000")
            current_name = self.hex_to_name_map.get(current_hex, color_names[0])
            var.set(current_name)

            combo = ttk.Combobox(main_frame, textvariable=var, values=color_names, state="readonly")
            combo.grid(row=i, column=1, sticky='ew', padx=5)
            self.color_vars[key] = var

            tip = _("LightBurn layer color for this step's lines in the SVG output.")
            if key == 'adhesive_ring_cut':
                tip = _("LightBurn layer color used for the adhesive rings in the SVG output: "
                        "the cuts and the size number share it.")
            add_tooltips(tip, row_lbl, combo)

        button_frame = tk.Frame(main_frame, bg=DIALOG_BG)
        button_frame.grid(row=len(layer_map_keys), column=0, columnspan=2, pady=20)
        save_btn = tk.Button(button_frame, text=_("Save"), command=self.save_colors)
        save_btn.pack(side="left", padx=10)
        cancel_btn = tk.Button(button_frame, text=_("Cancel"), command=self.top.destroy)
        cancel_btn.pack(side="left", padx=10)
        add_tooltip(save_btn, _("Apply the color choices and close."))
        add_tooltip(cancel_btn, _("Close without changing layer colors."))

        # An extra row (adhesive rings) must not push the buttons off the
        # fixed-size window.
        self.top.update_idletasks()
        need_h = main_frame.winfo_reqheight() + 10
        if need_h > 450:
            self.top.geometry(f"450x{need_h}")

    def save_colors(self):
        for key, var in self.color_vars.items():
            selected_name = var.get()
            self.settings["layer_colors"][key] = self.color_map[selected_name]
        
        self.save_callback()
        self.top.destroy()
