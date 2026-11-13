"""The per-material G-code settings dialog (Pad Maker and Tooling).

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog


from config import (
    DEFAULT_SETTINGS, GCODE_CUT_ONLY_MATERIALS,
)

from .common import DIALOG_BG, add_tooltip, add_tooltips, bind_mousewheel
import ui_dialogs as _pkg   # the dialogs suites stand in for are looked up through the package at call time


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

        if mat_key == "cork":
            # Packed after `frame`, so it sits under the rows built below.
            # Only the cut may be called tested (see the cork defaults in
            # config.py); the kerf and the engraving are the user's to find.
            tk.Label(
                mat_frame,
                text=_("The default cut was tested on 3/32 in (2.4 mm) sheet cork on a "
                       "40 W laser. Kerf is 0 and engraving is not set (0) until you "
                       "measure them: Tooling > Kerf Test and Speed & Power Test."),
                bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666",
                justify="left", wraplength=520).pack(anchor="w", pady=(4, 0))

        self._add_header_row(frame)

        if mat_key in GCODE_CUT_ONLY_MATERIALS:
            # Adhesive rings: one cut row and the kerf. No engraving and no
            # center hole exist for this material, so those rows (and their
            # tk vars) are simply not built; _capture_material_to_dict and
            # friends branch on the same tuple.
            self._add_operation_row(
                frame, mat_key, mat_settings, "cut", _("Cut (ring and number)"), 1,
                descr="ring cuts and the size number", short="ring cut")
            self._add_kerf_row(frame, mat_key, mat_settings, 2)
            note = tk.Label(
                mat_frame,
                text=_("The number is traced at these same settings. The defaults "
                       "are a starting point: dial them in with Tooling > Speed & "
                       "Power Test."),
                bg=DIALOG_BG, font=("Helvetica", 8), fg="#666666",
                justify="left", wraplength=520)
            note.pack(anchor="w", pady=(4, 0))
            return

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
            descr, short = op_help.get(op_key, (op_label.lower(), op_label.lower()))
            self._add_operation_row(frame, mat_key, mat_settings, op_key, op_label, i,
                                    descr=descr, short=short)

        self._add_kerf_row(frame, mat_key, mat_settings, 4 + len(self.OPERATIONS))

    def _add_header_row(self, frame):
        """The Operation / Speed / Power / Passes / Air column headers."""
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

    def _add_operation_row(self, frame, mat_key, mat_settings, op_key, op_label, row,
                           descr, short):
        """One speed / power / passes / air row (center hole, outer cut, or
        the single cut row of a cut-only material)."""
        i = row
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

    def _add_kerf_row(self, frame, mat_key, mat_settings, kerf_row):
        """The kerf width spinbox for one material."""
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
        if mat_key in GCODE_CUT_ONLY_MATERIALS:
            try:
                return {
                    "cut_speed": int(v['cut']['speed'].get()),
                    "cut_power": float(v['cut']['power'].get()),
                    "cut_passes": max(1, int(v['cut']['passes'].get())),
                    "kerf_width": float(v['kerf_width'].get()),
                    "air_assist_cut": bool(v['air_assist_cut'].get()),
                }
            except (tk.TclError, ValueError):
                return {"_invalid": True}
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

        if mat_key in GCODE_CUT_ONLY_MATERIALS:
            v['cut']['speed'].set(int(g("cut_speed", 600)))
            v['cut']['power'].set(g("cut_power", 60))
            v['cut']['passes'].set(int(g("cut_passes", 1)))
            v['kerf_width'].set(g("kerf_width", 0.0))
            v['air_assist_cut'].set(bool(g("air_assist_cut", True)))
            return

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
        dlg = _pkg.SaveSizingPresetDialog(
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

        for mat_key, _label in self.active_materials:
            new_gcode_settings[mat_key] = {}

            if mat_key in GCODE_CUT_ONLY_MATERIALS:
                # Cut, kerf and air only: this material has no engraving
                # or center-hole settings to save.
                try:
                    new_gcode_settings[mat_key] = {
                        "cut_speed": self.vars[mat_key]['cut']['speed'].get(),
                        "cut_power": self.vars[mat_key]['cut']['power'].get(),
                        "cut_passes": max(1, int(self.vars[mat_key]['cut']['passes'].get())),
                        "kerf_width": self.vars[mat_key]['kerf_width'].get(),
                        "air_assist_cut": self.vars[mat_key]['air_assist_cut'].get(),
                    }
                except tk.TclError:
                    messagebox.showerror(_("Invalid Input"),
                                         _("Invalid value for {mat} {op}. Please enter valid numbers.").format(mat=mat_key.replace('_', ' '), op="cut"),
                                         parent=self.top)
                    return
                continue

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
            for op_key, _label in self.OPERATIONS:
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

        for mat_key, _label in self.active_materials:
            mat_defaults = default_gcode.get(mat_key, {})

            if mat_key in GCODE_CUT_ONLY_MATERIALS:
                self.vars[mat_key]['cut']['speed'].set(mat_defaults.get("cut_speed", 100))
                self.vars[mat_key]['cut']['power'].set(mat_defaults.get("cut_power", 10))
                self.vars[mat_key]['cut']['passes'].set(mat_defaults.get("cut_passes", 1))
                self.vars[mat_key]['kerf_width'].set(mat_defaults.get("kerf_width", 0.0))
                self.vars[mat_key]['air_assist_cut'].set(mat_defaults.get("air_assist_cut", True))
                continue

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
            for op_key, _label in self.OPERATIONS:
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
