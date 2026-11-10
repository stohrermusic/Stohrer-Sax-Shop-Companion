"""
Tooling tab mixin for Stohrer Sax Shop Companion.

Provides die insert and die holder generation for laser-cutting acrylic tooling.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import logging
import os
import sys
import tempfile
import webbrowser
from collections import Counter

from svg_engine import (
    _nest_discs, try_nest_partial, can_all_pads_fit,
    generate_die_svg, generate_die_svg_from_placed,
    generate_holder_svg, generate_kerf_test_svg, check_holder_sheet,
    generate_die_organizer_svg, _holder_pieces_for, KERF_TEST_DIAMETERS,
    CrescentKerfError, OrganizerDrawingError, OrganizerKerfError, die_organizer_min_sheet,
    _min_feeds_speeds_sheet, _grid_pack_discs,
    build_feeds_speeds_matrix, feeds_speeds_ring_fits_label,
    CRESCENT_CURVE_STANDARD, CRESCENT_MAX_PERCENT, CRESCENT_MAX_TONEHOLE_MM,
    CRESCENT_MAX_THICKNESS_MM, CRESCENT_MAX_QTY, tonehole_crescent_geometry, tonehole_crescent_outline,
    generate_tonehole_crescent_svg, generate_tonehole_crescent_stl,
)
from gcode_engine import (
    generate_die_gcode_from_placed, generate_holder_gcode, generate_kerf_test_gcode,
    generate_feeds_speeds_test_gcode,
    generate_tonehole_crescent_gcode, generate_die_organizer_gcode, NoCutSettingsError,
)
from config import engraving_is_set, TOOLING_GCODE_MATERIALS
from ui_dialogs import GcodeSettingsWindow, SpeedPowerTestPreview, add_tooltip, gcode_file_note

IS_MACOS = sys.platform == 'darwin'

# The materials the Tooling dropdowns offer, by their English names. Lower-
# cased, each is its settings key. The dropdowns show translated names;
# _tooling_material_key maps back.
_TOOLING_MATERIALS = ("Felt", "Card", "Leather", "Acrylic", "Basswood", "Cork")
# What the Tonehole Crescent and Die Organizer material choices offer, first
# one the default: crescents are usually cork, the organizer usually wood.
_CRESCENT_MATERIALS = ("cork", "acrylic", "basswood", "leather", "felt", "card")
_ORGANIZER_MATERIALS = ("basswood", "acrylic", "cork")


class ToolingTabMixin:
    """Mixin class that adds the Tooling tab to the main application."""

    def _init_tooling_state(self):
        """Initialize tooling-specific state. Called from __init__."""
        self.tooling_scrap_session = {
            'active': False,
            'original_pads': [],
            'remaining_pads': [],
            'scrap_count': 0,
            'save_dir': '',
        }
        self.tooling_scrap_window = None
        # (button, pack options) for each section's Frame & Cut button.
        # main._refresh_machine_ui_state shows and hides them.
        self._tooling_frame_cut_btns = []

    def _add_tooling_frame_cut_button(self, parent, command, **pack_options):
        """A Frame & Cut button beside a section's Generate G-code button:
        the same job, sent straight to the laser instead of to a file.

        Created but not packed. main._refresh_machine_ui_state packs it
        when machine integration is switched on and a laser was detected,
        and takes it away again otherwise. Unlike the Pad Maker's button
        it does not wait for a camera calibration: a Tooling sheet is a
        rectangle the head is jogged to by hand.
        """
        label = _("Frame & Cut")
        # As wide as its neighbours, and never narrower than its own text:
        # the German label is 20 characters and a fixed 15 would clip it.
        btn = tk.Button(parent, text=label, width=max(15, len(label)), command=command)
        btn.tooltip = add_tooltip(btn, _("Send this job straight to the laser: position "
                                         "the head, frame the area, then cut."))
        self._tooling_frame_cut_btns.append((btn, pack_options))
        return btn

    def _build_phil_noy_credit(self, parent_frame, bg, lead_text):
        """Phil Noy credit block — inserted at the top of Die Inserts and
        Die Holders sections. `lead_text` is the section-specific opening
        phrase (e.g. "Die inserts are for the").
        """
        credit_frame = tk.Frame(parent_frame, bg=bg)
        credit_frame.pack(fill='x', pady=(0, 6))

        line1 = tk.Frame(credit_frame, bg=bg)
        line1.pack(anchor='w')
        tk.Label(line1, text=lead_text + " ",
                 bg=bg, font=("Helvetica", 9)).pack(side='left')
        noy_link = tk.Label(line1, text="Phil Noy",
                            bg=bg, fg="#0066CC",
                            font=("Helvetica", 9, "underline"),
                            cursor="hand2")
        noy_link.pack(side='left')
        noy_link.bind("<Button-1>", lambda e: webbrowser.open(
            "https://noysaxophonesupplies.com"))
        tk.Label(line1, text=_(" method of making saxophone pads."),
                 bg=bg, font=("Helvetica", 9)).pack(side='left')

        tk.Label(credit_frame,
                 text=_('They are engraved "DESIGNED BY PHIL NOY" because Phil '
                        'gave away this knowledge for free.'),
                 bg=bg, font=("Helvetica", 9),
                 justify="left").pack(anchor='w')
        tk.Label(credit_frame,
                 text=_("Please pay him the respect of including this in your tooling."),
                 bg=bg, font=("Helvetica", 9),
                 justify="left").pack(anchor='w', pady=(0, 4))

    def create_tooling_tab(self, parent):
        """Build the Tooling tab UI with accordion-style tool selector."""
        tooling = self.settings.get("tooling_settings", {})
        bg = self.root.cget('bg') if IS_MACOS else self.default_bg

        # Tool selector buttons at top, three to a row so longer labels
        # like "Speed & Power Test" stay readable in the 640 px window.
        selector_frame = tk.Frame(parent, bg=bg)
        selector_frame.pack(fill='x', padx=10, pady=(10, 5))

        self._tooling_sections = {}
        self._tooling_buttons = {}

        tool_names = [
            ("die_inserts", _("Die Inserts")),
            ("die_holders", _("Die Holders")),
            ("die_organizer", _("Die Organizer")),
            ("kerf_test", _("Kerf Test")),
            ("feeds_speeds_tester", _("Speed & Power Test")),
            ("pad_press_spacers", _("Pad Press Spacers")),
            ("tonehole_crescent", _("Tonehole Crescent")),
        ]
        # Calibration Card was deliberately removed from this list —
        # it's a workflow step in Pad Maker > Options > Machine >
        # Camera Calibration (which engraves the card AND captures
        # frames in one wizard), not a standalone tool. The engine
        # entry point `generate_calibration_card_gcode` lives on in
        # gcode_engine.py and is invoked by the calibration dialog.
        per_row = 3  # four 22-character buttons overflow the default window
        # One grid, its three columns in one uniform group, every button
        # stuck to both sides of its cell: "one width" is a property of the
        # layout, not of seven separate requests. Tk sizes the columns of a
        # uniform group alike, to the widest cell any of them holds, so
        # whatever a port asks for one button (Aqua gave two widths to the
        # seven when they were packed into row frames, CI 2026-11-09, a gap
        # that depended on the labels) every button is stretched to the
        # same cell. Row 3's single button sits in column 0 of the same
        # group. The character width below is the cells' minimum.
        for c in range(per_row):
            selector_frame.grid_columnconfigure(c, uniform="tooling-sections")
        for idx, (key, label) in enumerate(tool_names):
            r, c = divmod(idx, per_row)
            btn = tk.Button(selector_frame, text=label,
                            command=lambda k=key: self._show_tooling_section(k))
            btn.grid(row=r, column=c, sticky="ew", padx=(0, 5), pady=(4 if r else 0, 0))
            self._tooling_buttons[key] = btn
        # One width for all seven: 22 characters, as always, when every
        # label fits that (English on Windows), else as many as the longest
        # label needs. At a fixed 22, "Geschwindigkeits- und Leistungstest"
        # (German) and "Espaciadores de prensa de zapatillas" (Spanish) were
        # cut off at both ends, and so was English "Pad Press Spacers" with
        # Linux's wider font (CI, 2026-11-09). Measured, not assumed: a
        # button's requested width is computed when it is configured, on
        # every Tk port, so the labels' natural widths are read first and
        # the character width is raised until the button asks for at least
        # the widest of them. The window is sized to these rows by
        # main._fit_window_to_tabs, so they are never squeezed either.
        buttons = list(self._tooling_buttons.values())
        widest = max(b.winfo_reqwidth() for b in buttons)
        width = 22
        probe = buttons[0]
        probe.config(width=width)
        while probe.winfo_reqwidth() < widest and width < 80:
            width += 1
            probe.config(width=width)
        for b in buttons:
            b.config(width=width)

        # Content area (each section is a frame, only one visible at a time)
        self._tooling_content = tk.Frame(parent, bg=bg)
        self._tooling_content.pack(fill='both', expand=True, padx=10, pady=(0, 5))

        # ========================================
        # DIE INSERTS SECTION
        # ========================================
        die_frame = tk.Frame(self._tooling_content, bg=bg)

        self._build_phil_noy_credit(die_frame, bg,
                                     _("Die inserts are for the"))

        # Size input
        size_input_frame = tk.Frame(die_frame, bg=bg)
        size_input_frame.pack(fill='x', pady=(0, 5))

        tk.Label(size_input_frame, text=_("Enter sizes (mm):"), bg=bg).pack(anchor='w')
        self.die_size_entry = tk.Text(size_input_frame, height=2, width=50, font=("Courier", 10))
        self.die_size_entry.pack(fill='x', pady=(2, 2))

        hint_label = tk.Label(size_input_frame, text=_('e.g. "7, 8.5, 15-30" or "7-39.5"'),
                              bg=bg, fg="gray", font=("Helvetica", 8))
        hint_label.pack(anchor='w')

        # Quick-fill buttons and step size
        buttons_frame = tk.Frame(die_frame, bg=bg)
        buttons_frame.pack(fill='x', pady=(0, 5))

        tk.Button(buttons_frame, text=_("Full Set (S)"), width=12,
                  command=lambda: self._fill_die_sizes("7-39.5")).pack(side='left', padx=(0, 5))
        tk.Button(buttons_frame, text=_("Full Set (L)"), width=12,
                  command=lambda: self._fill_die_sizes("40-60")).pack(side='left', padx=(0, 5))
        tk.Button(buttons_frame, text=_("Full Set (All)"), width=12,
                  command=lambda: self._fill_die_sizes("7-60")).pack(side='left', padx=(0, 15))

        tk.Label(buttons_frame, text=_("Step:"), bg=bg).pack(side='left')
        self.die_step_var = tk.StringVar(value=tooling.get("step_size", "0.5"))
        step_entry = tk.Entry(buttons_frame, textvariable=self.die_step_var, width=5)
        step_entry.pack(side='left', padx=(2, 2))
        tk.Label(buttons_frame, text=_("mm"), bg=bg).pack(side='left')

        # Engraving options
        eng_frame = tk.Frame(die_frame, bg=bg)
        eng_frame.pack(fill='x', pady=(5, 5))

        self.die_engrave_ring_var = tk.BooleanVar(value=tooling.get("engrave_ring", True))
        tk.Checkbutton(eng_frame, text=_("Engrave size on ring"), variable=self.die_engrave_ring_var,
                       bg=bg).pack(side='left', padx=(0, 15))

        self.die_engrave_cutout_var = tk.BooleanVar(value=tooling.get("engrave_cutout", True))
        tk.Checkbutton(eng_frame, text=_("Engrave size on cutout"), variable=self.die_engrave_cutout_var,
                       bg=bg).pack(side='left', padx=(0, 15))

        # Sheet size
        sheet_frame = tk.LabelFrame(die_frame, text=_("Sheet Size"), bg=bg, padx=5, pady=5)
        sheet_frame.pack(fill='x', pady=(5, 5))

        size_row = tk.Frame(sheet_frame, bg=bg)
        size_row.pack(fill='x')

        tk.Label(size_row, text=_("Width:"), bg=bg).pack(side='left')
        self.die_width_var = tk.StringVar(value=tooling.get("sheet_width", "12"))
        tk.Entry(size_row, textvariable=self.die_width_var, width=8).pack(side='left', padx=(2, 10))

        tk.Label(size_row, text=_("Height:"), bg=bg).pack(side='left')
        self.die_height_var = tk.StringVar(value=tooling.get("sheet_height", "12"))
        tk.Entry(size_row, textvariable=self.die_height_var, width=8).pack(side='left', padx=(2, 10))

        self.die_units_var = tk.StringVar(value=self.settings.get("units", "in"))
        tk.Radiobutton(size_row, text=_("in"), variable=self.die_units_var,
                       value="in", bg=bg).pack(side='left')
        tk.Radiobutton(size_row, text=_("mm"), variable=self.die_units_var,
                       value="mm", bg=bg).pack(side='left', padx=(5, 0))

        # Scrap mode
        scrap_row = tk.Frame(die_frame, bg=bg)
        scrap_row.pack(fill='x', pady=(5, 5))

        self.die_scrap_var = tk.BooleanVar(value=False)
        tk.Checkbutton(scrap_row, text=_("Scrap Mode"), variable=self.die_scrap_var,
                       bg=bg, command=self._toggle_die_scrap_mode).pack(side='left')

        self.die_scrap_clear_btn = tk.Button(scrap_row, text=_("Clear"), width=5,
                                              command=self._on_clear_die_scrap_clicked)
        # Hidden initially
        self.die_scrap_status_var = tk.StringVar(value="")
        self.die_scrap_status_label = tk.Label(scrap_row, textvariable=self.die_scrap_status_var,
                                                bg=bg, font=("Helvetica", 9))

        # Output filename
        name_frame = tk.Frame(die_frame, bg=bg)
        name_frame.pack(fill='x', pady=(5, 5))

        tk.Label(name_frame, text=_("Output filename:"), bg=bg).pack(side='left')
        self.die_filename_var = tk.StringVar(value="my_dies")
        tk.Entry(name_frame, textvariable=self.die_filename_var, width=25).pack(side='left', padx=(5, 0))

        # Generate buttons
        gen_frame = tk.Frame(die_frame, bg=bg)
        gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(gen_frame, text=_("Generate SVG"), width=15,
                  command=self._on_generate_die_svg).pack(side='left', padx=(0, 10))
        tk.Button(gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_die_gcode).pack(side='left')
        self._add_tooling_frame_cut_button(
            gen_frame, self._on_frame_cut_die, side='left', padx=(10, 0))

        self._tooling_sections['die_inserts'] = die_frame

        # ========================================
        # DIE HOLDERS SECTION
        # ========================================
        holder_frame = tk.Frame(self._tooling_content, bg=bg)

        self._build_phil_noy_credit(holder_frame, bg,
                                     _("Die holders are for the"))

        # Layer description (live-updated from layer count radio)
        self.holder_info_var = tk.StringVar()
        holder_info = tk.Label(holder_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                               textvariable=self.holder_info_var,
                               justify="left")
        holder_info.pack(anchor='w', pady=(0, 8))

        # Layer count
        layer_frame = tk.Frame(holder_frame, bg=bg)
        layer_frame.pack(fill='x', pady=(0, 5))
        tk.Label(layer_frame, text=_("Layers:"), bg=bg).pack(side='left')
        self.holder_layer_count_var = tk.IntVar(value=int(tooling.get("holder_layer_count", 6)))
        tk.Radiobutton(layer_frame, text=_("5-layer (2\u00d7 pin)"), variable=self.holder_layer_count_var,
                       value=5, bg=bg, command=self._update_holder_info).pack(side='left', padx=(5, 10))
        tk.Radiobutton(layer_frame, text=_("6-layer (3\u00d7 pin)"), variable=self.holder_layer_count_var,
                       value=6, bg=bg, command=self._update_holder_info).pack(side='left')

        # Variant
        variant_frame = tk.Frame(holder_frame, bg=bg)
        variant_frame.pack(fill='x', pady=(0, 5))
        tk.Label(variant_frame, text=_("Generate:"), bg=bg).pack(side='left')
        self.holder_variant_var = tk.StringVar(value="large")
        tk.Radiobutton(variant_frame, text=_("Large (40\u201360mm pads)"), variable=self.holder_variant_var,
                       value="large", bg=bg).pack(side='left', padx=(5, 10))
        tk.Radiobutton(variant_frame, text=_("Small (7\u201339.5mm pads)"), variable=self.holder_variant_var,
                       value="small", bg=bg).pack(side='left', padx=(0, 10))
        tk.Radiobutton(variant_frame, text=_("Both"), variable=self.holder_variant_var,
                       value="both", bg=bg).pack(side='left')

        # Sheet size (mirrors the Die Inserts pattern)
        holder_sheet_frame = tk.LabelFrame(holder_frame, text=_("Sheet Size"), bg=bg, padx=5, pady=5)
        holder_sheet_frame.pack(fill='x', pady=(5, 5))

        holder_size_row = tk.Frame(holder_sheet_frame, bg=bg)
        holder_size_row.pack(fill='x')

        tk.Label(holder_size_row, text=_("Width:"), bg=bg).pack(side='left')
        self.holder_width_var = tk.StringVar(value=tooling.get("holder_sheet_width", "12"))
        tk.Entry(holder_size_row, textvariable=self.holder_width_var, width=8).pack(side='left', padx=(2, 10))

        tk.Label(holder_size_row, text=_("Height:"), bg=bg).pack(side='left')
        self.holder_height_var = tk.StringVar(value=tooling.get("holder_sheet_height", "12"))
        tk.Entry(holder_size_row, textvariable=self.holder_height_var, width=8).pack(side='left', padx=(2, 10))

        self.holder_units_var = tk.StringVar(value=self.settings.get("units", "in"))
        tk.Radiobutton(holder_size_row, text=_("in"), variable=self.holder_units_var,
                       value="in", bg=bg).pack(side='left')
        tk.Radiobutton(holder_size_row, text=_("mm"), variable=self.holder_units_var,
                       value="mm", bg=bg).pack(side='left', padx=(5, 0))

        holder_name_frame = tk.Frame(holder_frame, bg=bg)
        holder_name_frame.pack(fill='x', pady=(0, 5))

        tk.Label(holder_name_frame, text=_("Output filename:"), bg=bg).pack(side='left')
        self.holder_filename_var = tk.StringVar(value="die_holder")
        tk.Entry(holder_name_frame, textvariable=self.holder_filename_var, width=25).pack(side='left', padx=(5, 0))

        holder_gen_frame = tk.Frame(holder_frame, bg=bg)
        holder_gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(holder_gen_frame, text=_("Generate SVG"), width=15,
                  command=self._on_generate_holder_svg).pack(side='left', padx=(0, 10))
        tk.Button(holder_gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_holder_gcode).pack(side='left')
        self._add_tooling_frame_cut_button(
            holder_gen_frame, self._on_frame_cut_holder, side='left', padx=(10, 0))

        self._tooling_sections['die_holders'] = holder_frame
        self._update_holder_info()

        # ========================================
        # DIE ORGANIZER SECTION
        # ========================================
        organizer_frame = tk.Frame(self._tooling_content, bg=bg)

        # Description + Matt's instructions
        organizer_info = tk.Label(organizer_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                                  text=_("Plate: 230 × 330 mm. Generate G-code to cut it with this "
                                         "app, or open the SVG in LightBurn or your laser software."),
                                  justify="left", wraplength=540)
        organizer_info.pack(anchor='w', pady=(0, 4))

        instructions_text = _(
            "Cut three of the upper and one of the lower. Use the locating "
            "holes in the corners to align the layers and glue them together "
            "(a book press lightly clamps it nicely while the glue sets). "
            "Locating holes are 1/8\" — resize for whatever pin you have. "
            "Wood works great; acrylic will too."
        )
        instructions_label = tk.Label(organizer_frame, bg=bg, font=("Helvetica", 9),
                                      text=instructions_text, justify="left",
                                      wraplength=520)
        instructions_label.pack(anchor='w', pady=(0, 8))

        # Variant
        organizer_variant_frame = tk.Frame(organizer_frame, bg=bg)
        organizer_variant_frame.pack(fill='x', pady=(0, 5))
        tk.Label(organizer_variant_frame, text=_("Generate:"), bg=bg).pack(side='left')
        self.organizer_variant_var = tk.StringVar(value="upper")
        tk.Radiobutton(organizer_variant_frame, text=_("Upper (slotted)"),
                       variable=self.organizer_variant_var,
                       value="upper", bg=bg).pack(side='left', padx=(5, 10))
        tk.Radiobutton(organizer_variant_frame, text=_("Lower (base)"),
                       variable=self.organizer_variant_var,
                       value="lower", bg=bg).pack(side='left')

        # Material and sheet: the G-code needs both (the SVG is the bare
        # drawing and takes neither).
        organizer_mat_row = tk.Frame(organizer_frame, bg=bg)
        organizer_mat_row.pack(fill='x', pady=(0, 5))
        tk.Label(organizer_mat_row, text=_("Material:"), bg=bg).pack(side='left')
        self.organizer_material_var = tk.StringVar(value=self._tooling_material_label(
            tooling.get("organizer_material"), _ORGANIZER_MATERIALS))
        organizer_mat_combo = ttk.Combobox(
            organizer_mat_row, textvariable=self.organizer_material_var,
            values=[self._tooling_material_label(k) for k in _ORGANIZER_MATERIALS],
            state="readonly", width=12)
        organizer_mat_combo.pack(side='left', padx=(5, 5))
        organizer_mat_note = tk.Label(organizer_mat_row, text=_("(G-code only)"), bg=bg)
        organizer_mat_note.pack(side='left')
        add_tooltip(organizer_mat_combo,
                    _("What the plate is cut from. The G-code takes its cut speed, power, "
                      "passes and kerf from this material in Options > Settings."))

        organizer_sheet_frame = tk.LabelFrame(organizer_frame, text=_("Sheet Size"), bg=bg, padx=5, pady=5)
        organizer_sheet_frame.pack(fill='x', pady=(0, 5))
        organizer_size_row = tk.Frame(organizer_sheet_frame, bg=bg)
        organizer_size_row.pack(fill='x')
        org_units = tooling.get("organizer_sheet_units") or (
            "in" if self.settings.get("units", "in") == "in" else "mm")
        tk.Label(organizer_size_row, text=_("Width:"), bg=bg).pack(side='left')
        self.organizer_width_var = tk.StringVar(
            value=tooling.get("organizer_sheet_width") or ("12" if org_units == "in" else "305"))
        tk.Entry(organizer_size_row, textvariable=self.organizer_width_var, width=8).pack(side='left', padx=(2, 10))
        tk.Label(organizer_size_row, text=_("Height:"), bg=bg).pack(side='left')
        self.organizer_height_var = tk.StringVar(
            value=tooling.get("organizer_sheet_height") or ("18" if org_units == "in" else "457"))
        tk.Entry(organizer_size_row, textvariable=self.organizer_height_var, width=8).pack(side='left', padx=(2, 10))
        self.organizer_units_var = tk.StringVar(value=org_units)
        tk.Radiobutton(organizer_size_row, text=_("in"), variable=self.organizer_units_var,
                       value="in", bg=bg).pack(side='left')
        tk.Radiobutton(organizer_size_row, text=_("mm"), variable=self.organizer_units_var,
                       value="mm", bg=bg).pack(side='left', padx=(5, 0))

        # Output filename
        organizer_name_frame = tk.Frame(organizer_frame, bg=bg)
        organizer_name_frame.pack(fill='x', pady=(0, 5))
        tk.Label(organizer_name_frame, text=_("Output filename:"), bg=bg).pack(side='left')
        self.organizer_filename_var = tk.StringVar(value="die_organizer")
        tk.Entry(organizer_name_frame, textvariable=self.organizer_filename_var,
                 width=25).pack(side='left', padx=(5, 0))

        organizer_gen_frame = tk.Frame(organizer_frame, bg=bg)
        organizer_gen_frame.pack(fill='x', pady=(5, 0))
        tk.Button(organizer_gen_frame, text=_("Generate SVG"), width=15,
                  command=self._on_generate_organizer_svg).pack(side='left', padx=(0, 10))
        tk.Button(organizer_gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_organizer_gcode).pack(side='left')
        self._add_tooling_frame_cut_button(
            organizer_gen_frame, self._on_frame_cut_organizer, side='left', padx=(10, 0))

        self._tooling_sections['die_organizer'] = organizer_frame

        # ========================================
        # KERF TEST SECTION
        # ========================================
        kerf_frame = tk.Frame(self._tooling_content, bg=bg)

        # Sheet size info
        kerf_info = tk.Label(kerf_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                             text=_("Pattern size: 110 \u00d7 54 mm (4.3 \u00d7 2.1 in)"),
                             justify="left")
        kerf_info.pack(anchor='w', pady=(0, 8))

        # Material dropdown
        mat_row = tk.Frame(kerf_frame, bg=bg)
        mat_row.pack(fill='x', pady=(0, 5))

        tk.Label(mat_row, text=_("Material:"), bg=bg).pack(side='left')
        self.kerf_material_var = tk.StringVar(value="Acrylic")
        kerf_materials = [_("Felt"), _("Card"), _("Leather"), _("Acrylic"), _("Cork")]
        kerf_dropdown = ttk.Combobox(mat_row, textvariable=self.kerf_material_var,
                                     values=kerf_materials, state="readonly", width=12)
        kerf_dropdown.pack(side='left', padx=(5, 0))

        # Use existing settings checkbox
        settings_row = tk.Frame(kerf_frame, bg=bg)
        settings_row.pack(fill='x', pady=(0, 5))

        self.kerf_use_existing_var = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_row, text=_("Use existing G-code settings for this material"),
                       variable=self.kerf_use_existing_var, bg=bg,
                       command=self._toggle_kerf_custom_settings).pack(anchor='w')

        # Custom settings (hidden by default)
        self.kerf_custom_frame = tk.Frame(kerf_frame, bg=bg)

        custom_row = tk.Frame(self.kerf_custom_frame, bg=bg)
        custom_row.pack(fill='x')

        tk.Label(custom_row, text=_("Cut speed:"), bg=bg).pack(side='left')
        self.kerf_speed_var = tk.StringVar(value="180")
        tk.Entry(custom_row, textvariable=self.kerf_speed_var, width=8).pack(side='left', padx=(2, 5))
        tk.Label(custom_row, text=_("mm/min"), bg=bg).pack(side='left', padx=(0, 15))

        tk.Label(custom_row, text=_("Cut power:"), bg=bg).pack(side='left')
        self.kerf_power_var = tk.StringVar(value="100")
        tk.Entry(custom_row, textvariable=self.kerf_power_var, width=8).pack(side='left', padx=(2, 5))
        tk.Label(custom_row, text="%", bg=bg).pack(side='left')

        # Generate buttons
        kerf_gen_frame = tk.Frame(kerf_frame, bg=bg)
        kerf_gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(kerf_gen_frame, text=_("Generate SVG"), width=15,
                  command=self._on_generate_kerf_svg).pack(side='left', padx=(0, 10))
        tk.Button(kerf_gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_kerf_gcode).pack(side='left')
        self._add_tooling_frame_cut_button(
            kerf_gen_frame, self._on_frame_cut_kerf, side='left', padx=(10, 0))

        self._tooling_sections['kerf_test'] = kerf_frame

        # ========================================
        # SPEED & POWER TEST SECTION
        # ========================================
        fs_frame = tk.Frame(self._tooling_content, bg=bg)
        fs_settings = tooling.get("feeds_speeds_tester", {})

        tk.Label(fs_frame, text=_("BETA"), bg=bg, fg="#cc6600",
                 font=("Helvetica", 9, "bold")).pack(anchor='w')

        fs_info = tk.Label(
            fs_frame, bg=bg, font=("Helvetica", 9), fg="gray",
            text=_("Generate a sheet of small test discs, each cut at a different "
                   "combination of speed/power/passes. Each disc is engraved with "
                   "its ID; a legend.txt saved alongside maps IDs to parameters. "
                   "Inspect which discs cut cleanly to find your settings."),
            justify="left", wraplength=540)
        fs_info.pack(anchor='w', pady=(0, 8))

        # Material picker + apply-defaults
        fs_mat_row = tk.Frame(fs_frame, bg=bg)
        fs_mat_row.pack(fill='x', pady=(0, 5))
        tk.Label(fs_mat_row, text=_("Material:"), bg=bg).pack(side='left')
        self.fs_material_var = tk.StringVar(value=fs_settings.get("material", "Felt"))
        fs_mat_combo = ttk.Combobox(
            fs_mat_row, textvariable=self.fs_material_var,
            values=[_("Felt"), _("Card"), _("Leather"), _("Acrylic"),
                    _("Basswood"), _("Cork")],
            state="readonly", width=12)
        fs_mat_combo.pack(side='left', padx=(5, 10))
        tk.Button(fs_mat_row, text=_("Apply material defaults"),
                  command=self._apply_feeds_speeds_material_defaults).pack(side='left')

        # Disc diameter + optional center hole (for washers / shims)
        fs_dia_row = tk.Frame(fs_frame, bg=bg)
        fs_dia_row.pack(fill='x', pady=(0, 5))
        tk.Label(fs_dia_row, text=_("Disc diameter:"), bg=bg).pack(side='left')
        self.fs_diameter_var = tk.StringVar(
            value=str(fs_settings.get("disc_diameter_mm", 20.0)))
        tk.Entry(fs_dia_row, textvariable=self.fs_diameter_var, width=8
                 ).pack(side='left', padx=(2, 2))
        tk.Label(fs_dia_row, text=_("mm"), bg=bg).pack(side='left', padx=(0, 15))
        tk.Label(fs_dia_row, text=_("Hole diameter:"), bg=bg).pack(side='left')
        self.fs_inner_diameter_var = tk.StringVar(
            value=str(fs_settings.get("inner_diameter_mm", 0.0)))
        tk.Entry(fs_dia_row, textvariable=self.fs_inner_diameter_var, width=8
                 ).pack(side='left', padx=(2, 2))
        tk.Label(fs_dia_row, text=_("mm  (0 = solid disc)"), bg=bg).pack(side='left')
        self.fs_diameter_var.trace_add(
            "write", lambda *a: self._update_feeds_speeds_preview())
        self.fs_inner_diameter_var.trace_add(
            "write", lambda *a: self._update_feeds_speeds_preview())

        # Sweep rows (Speed, Power, Passes)
        sweep_frame = tk.LabelFrame(fs_frame, text=_("Sweep Variables"),
                                    bg=bg, padx=5, pady=5)
        sweep_frame.pack(fill='x', pady=(5, 5))

        sweep_hint = tk.Label(
            sweep_frame, bg=bg, fg="gray", font=("Helvetica", 8),
            text=_('"Sweep" checked → test "Stops" evenly-spaced values from Start to End. '
                   'Unchecked → held constant at "Value".'))
        sweep_hint.pack(anchor='w', pady=(0, 3))

        self.fs_sweep_vars = {}
        for var_name, label in [("speed", _("Speed (mm/min)")),
                                 ("power", _("Power (%)")),
                                 ("passes", _("Passes"))]:
            row = tk.Frame(sweep_frame, bg=bg)
            row.pack(fill='x', pady=(2, 2))

            sweep_var = tk.BooleanVar(value=fs_settings.get(f"{var_name}_sweep", True))
            tk.Checkbutton(row, text=label, variable=sweep_var, bg=bg, width=18,
                           anchor='w',
                           command=self._update_feeds_speeds_preview).pack(side='left')

            val_var = tk.StringVar(value=str(fs_settings.get(f"{var_name}_value", 0)))
            tk.Label(row, text=_("Value:"), bg=bg).pack(side='left', padx=(10, 2))
            tk.Entry(row, textvariable=val_var, width=6).pack(side='left')

            start_var = tk.StringVar(value=str(fs_settings.get(f"{var_name}_start", 0)))
            end_var = tk.StringVar(value=str(fs_settings.get(f"{var_name}_end", 0)))
            # Stops is always an integer — coerce in case a previous save
            # left it as a float (e.g. "4.0") so the entry displays "4".
            try:
                stops_initial = int(float(fs_settings.get(f"{var_name}_stops", 4)))
            except (ValueError, TypeError):
                stops_initial = 4
            stops_var = tk.StringVar(value=str(stops_initial))

            tk.Label(row, text=_("Start:"), bg=bg).pack(side='left', padx=(10, 2))
            tk.Entry(row, textvariable=start_var, width=6).pack(side='left')
            tk.Label(row, text=_("End:"), bg=bg).pack(side='left', padx=(5, 2))
            tk.Entry(row, textvariable=end_var, width=6).pack(side='left')
            tk.Label(row, text=_("Stops:"), bg=bg).pack(side='left', padx=(5, 2))
            tk.Entry(row, textvariable=stops_var, width=4).pack(side='left')

            for entry_var in (val_var, start_var, end_var, stops_var):
                entry_var.trace_add(
                    "write", lambda *a: self._update_feeds_speeds_preview())

            self.fs_sweep_vars[var_name] = {
                'sweep': sweep_var, 'value': val_var,
                'start': start_var, 'end': end_var, 'stops': stops_var,
            }

        # Sheet size
        fs_sheet_frame = tk.LabelFrame(fs_frame, text=_("Sheet Size"),
                                       bg=bg, padx=5, pady=5)
        fs_sheet_frame.pack(fill='x', pady=(5, 5))

        fs_sheet_row = tk.Frame(fs_sheet_frame, bg=bg)
        fs_sheet_row.pack(fill='x')
        tk.Label(fs_sheet_row, text=_("Width:"), bg=bg).pack(side='left')
        self.fs_sheet_w_var = tk.StringVar(value=str(fs_settings.get("sheet_w", "4")))
        tk.Entry(fs_sheet_row, textvariable=self.fs_sheet_w_var, width=8
                 ).pack(side='left', padx=(2, 10))
        tk.Label(fs_sheet_row, text=_("Height:"), bg=bg).pack(side='left')
        self.fs_sheet_h_var = tk.StringVar(value=str(fs_settings.get("sheet_h", "6")))
        tk.Entry(fs_sheet_row, textvariable=self.fs_sheet_h_var, width=8
                 ).pack(side='left', padx=(2, 10))
        self.fs_sheet_unit_var = tk.StringVar(value=fs_settings.get("sheet_unit", "in"))
        tk.Radiobutton(fs_sheet_row, text=_("in"), variable=self.fs_sheet_unit_var,
                       value="in", bg=bg,
                       command=self._update_feeds_speeds_preview).pack(side='left')
        tk.Radiobutton(fs_sheet_row, text=_("mm"), variable=self.fs_sheet_unit_var,
                       value="mm", bg=bg,
                       command=self._update_feeds_speeds_preview
                       ).pack(side='left', padx=(5, 0))
        self.fs_sheet_w_var.trace_add(
            "write", lambda *a: self._update_feeds_speeds_preview())
        self.fs_sheet_h_var.trace_add(
            "write", lambda *a: self._update_feeds_speeds_preview())

        # Engraving + air assist
        fs_opt_row = tk.Frame(fs_frame, bg=bg)
        fs_opt_row.pack(fill='x', pady=(0, 2))
        self.fs_engraving_var = tk.BooleanVar(
            value=fs_settings.get("engraving_on", True))
        tk.Checkbutton(fs_opt_row, text=_("Engrave ID on each disc"),
                       variable=self.fs_engraving_var, bg=bg
                       ).pack(side='left', padx=(0, 15))
        self.fs_air_assist_var = tk.BooleanVar(
            value=fs_settings.get("air_assist", True))
        tk.Checkbutton(fs_opt_row, text=_("Air assist"),
                       variable=self.fs_air_assist_var, bg=bg
                       ).pack(side='left', padx=(0, 15))
        # "Also test with air off" doubles the matrix so the user can see
        # the effect of air on edge quality side-by-side on one sheet.
        self.fs_also_no_air_var = tk.BooleanVar(
            value=fs_settings.get("also_test_no_air", False))
        tk.Checkbutton(fs_opt_row,
                       text=_("Also test with air off (doubles disc count)"),
                       variable=self.fs_also_no_air_var, bg=bg,
                       command=self._update_feeds_speeds_preview).pack(side='left')

        # Engraving feed/power — needed because the labels themselves have
        # to engrave well or the user can't identify their best disc.
        # Pre-filled from material defaults via "Apply material defaults".
        fs_eng_row = tk.Frame(fs_frame, bg=bg)
        fs_eng_row.pack(fill='x', pady=(0, 5))
        tk.Label(fs_eng_row, text=_("Engraving speed:"), bg=bg).pack(side='left')
        self.fs_eng_speed_var = tk.StringVar(
            value=str(fs_settings.get("eng_speed_value", 1200)))
        tk.Entry(fs_eng_row, textvariable=self.fs_eng_speed_var, width=6
                 ).pack(side='left', padx=(2, 2))
        tk.Label(fs_eng_row, text=_("mm/min"), bg=bg).pack(side='left', padx=(0, 15))
        tk.Label(fs_eng_row, text=_("Engraving power:"), bg=bg).pack(side='left')
        self.fs_eng_power_var = tk.StringVar(
            value=str(fs_settings.get("eng_power_value", 10)))
        tk.Entry(fs_eng_row, textvariable=self.fs_eng_power_var, width=6
                 ).pack(side='left', padx=(2, 2))
        tk.Label(fs_eng_row, text="%", bg=bg).pack(side='left')

        # Output filename
        fs_name_row = tk.Frame(fs_frame, bg=bg)
        fs_name_row.pack(fill='x', pady=(0, 5))
        tk.Label(fs_name_row, text=_("Output filename:"), bg=bg).pack(side='left')
        self.fs_filename_var = tk.StringVar(
            value=fs_settings.get("filename", "speed_power_test"))
        tk.Entry(fs_name_row, textvariable=self.fs_filename_var, width=25
                 ).pack(side='left', padx=(5, 0))

        # Live preview
        self.fs_preview_var = tk.StringVar(value="")
        tk.Label(fs_frame, textvariable=self.fs_preview_var, bg=bg,
                 fg="#444", font=("Helvetica", 9, "italic")
                 ).pack(anchor='w', pady=(2, 5))

        # Generate buttons + preview-before-saving toggle
        fs_gen_frame = tk.Frame(fs_frame, bg=bg)
        fs_gen_frame.pack(fill='x', pady=(5, 0))
        # G-code only — an SVG would just show plain circles since the
        # per-disc speed/power/passes only exist in G-code layer params.
        tk.Button(fs_gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_feeds_speeds_gcode
                  ).pack(side='left', padx=(0, 15))
        self.fs_show_preview_var = tk.BooleanVar(
            value=fs_settings.get("show_preview", True))
        fs_preview_chk = tk.Checkbutton(
            fs_gen_frame, text=_("Preview before saving"),
            variable=self.fs_show_preview_var, bg=bg)
        fs_preview_chk.pack(side='left')
        # Between Generate G-code and the preview tick, which applies to both.
        self._add_tooling_frame_cut_button(
            fs_gen_frame, self._on_frame_cut_feeds_speeds,
            side='left', padx=(0, 15), before=fs_preview_chk)

        self._tooling_sections['feeds_speeds_tester'] = fs_frame

        # ========================================
        # PAD PRESS SPACERS SECTION
        # ========================================
        spacers_frame = tk.Frame(self._tooling_content, bg=bg)

        spacers_info = tk.Label(
            spacers_frame, bg=bg, font=("Helvetica", 9), fg="gray",
            text=_("3D-printable spacer biscuits for pad pressing. "
                   "1.75″ (44.45 mm) square. Each biscuit has its thickness "
                   "engraved on top. PLA or PETG works well; the spacers "
                   "stack to set pad press depth."),
            justify="left", wraplength=520)
        spacers_info.pack(anchor='w', pady=(0, 8))

        # Half-step plate
        half_row = tk.Frame(spacers_frame, bg=bg)
        half_row.pack(fill='x', pady=(0, 4))
        tk.Label(half_row, bg=bg, font=("Helvetica", 10),
                 text=_("Half-step set — 3.0 / 3.5 / 4.0 / 4.5 mm (4 of each, 16 total)"),
                 justify="left", anchor='w').pack(side='left')
        tk.Button(half_row, text=_("Save STL…"), width=12,
                  command=lambda: self._save_pad_spacer_stl("pad_spacers_all.stl")
                  ).pack(side='right', padx=(8, 0))

        # Quarter-step plate
        qtr_row = tk.Frame(spacers_frame, bg=bg)
        qtr_row.pack(fill='x', pady=(0, 4))
        tk.Label(qtr_row, bg=bg, font=("Helvetica", 10),
                 text=_("Quarter-step set — 3.25 / 3.75 / 4.25 mm (4 of each, 12 total)"),
                 justify="left", anchor='w').pack(side='left')
        tk.Button(qtr_row, text=_("Save STL…"), width=12,
                  command=lambda: self._save_pad_spacer_stl("pad_spacers_qtr_all.stl")
                  ).pack(side='right', padx=(8, 0))

        # Organizer
        org_row = tk.Frame(spacers_frame, bg=bg)
        org_row.pack(fill='x', pady=(0, 4))
        tk.Label(org_row, bg=bg, font=("Helvetica", 10),
                 text=_("Organizer rack — 7 compartments for all 28 spacers (3.0 – 4.5 mm)"),
                 justify="left", anchor='w').pack(side='left')
        tk.Button(org_row, text=_("Save STL…"), width=12,
                  command=lambda: self._save_pad_spacer_stl("pad_spacer_organizer.stl")
                  ).pack(side='right', padx=(8, 0))

        self._tooling_sections['pad_press_spacers'] = spacers_frame

        # ========================================
        # TONEHOLE CRESCENT SECTION
        # ========================================
        crescent_frame = tk.Frame(self._tooling_content, bg=bg)

        tk.Label(crescent_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                 text=_("The outer edge is the tonehole ID exactly. G-code allows for the "
                        "material's kerf; for SVG and STL allow for kerf or print tolerance yourself."),
                 justify="left", wraplength=540).pack(anchor='w', pady=(0, 8))

        cr_size_row = tk.Frame(crescent_frame, bg=bg)
        cr_size_row.pack(fill='x', pady=(0, 5))
        tk.Label(cr_size_row, text=_("Tonehole ID:"), bg=bg).pack(side='left')
        self.crescent_tonehole_var = tk.StringVar(value=str(tooling.get("crescent_tonehole_d", "40")))
        tk.Entry(cr_size_row, textvariable=self.crescent_tonehole_var, width=8).pack(side='left', padx=(2, 2))
        tk.Label(cr_size_row, text=_("mm"), bg=bg).pack(side='left', padx=(0, 15))
        tk.Label(cr_size_row, text=_("Reduce open area by:"), bg=bg).pack(side='left')
        self.crescent_percent_var = tk.StringVar(value=str(tooling.get("crescent_percent", "10")))
        tk.Entry(cr_size_row, textvariable=self.crescent_percent_var, width=6).pack(side='left', padx=(2, 2))
        tk.Label(cr_size_row, text="%", bg=bg).pack(side='left')

        cr_curve_row = tk.Frame(crescent_frame, bg=bg)
        cr_curve_row.pack(fill='x', pady=(0, 5))
        tk.Label(cr_curve_row, text=_("Inner edge:"), bg=bg).pack(side='left')
        tk.Label(cr_curve_row, text=_("Flatter"), bg=bg, font=("Helvetica", 8)).pack(side='left', padx=(6, 0))
        try:
            curve_initial = float(tooling.get("crescent_curve", CRESCENT_CURVE_STANDARD))
        except (ValueError, TypeError):
            curve_initial = CRESCENT_CURVE_STANDARD
        self.crescent_curve_var = tk.DoubleVar(value=curve_initial)
        tk.Scale(cr_curve_row, from_=0, to=100, orient='horizontal', variable=self.crescent_curve_var,
                 showvalue=False, length=200, bg=bg, highlightthickness=0,
                 command=lambda v: self._update_crescent_preview()).pack(side='left', padx=4)
        tk.Label(cr_curve_row, text=_("Rounder"), bg=bg, font=("Helvetica", 8)).pack(side='left', padx=(0, 10))
        tk.Button(cr_curve_row, text=_("Match tonehole curve"),
                  command=lambda: (self.crescent_curve_var.set(CRESCENT_CURVE_STANDARD),
                                   self._update_crescent_preview())).pack(side='left')

        # Preview: shape only (it depends on the percentage and curve, not
        # on the ID), mouthpiece side up, the way it sits in the horn.
        cr_preview_row = tk.Frame(crescent_frame, bg=bg)
        cr_preview_row.pack(fill='x', pady=(2, 5))
        self.crescent_canvas = tk.Canvas(cr_preview_row, width=180, height=180, bg=bg,
                                         highlightthickness=0)
        self.crescent_canvas.pack(side='left')
        self.crescent_readout_var = tk.StringVar(value="")
        tk.Label(cr_preview_row, textvariable=self.crescent_readout_var, bg=bg,
                 fg="#444", font=("Helvetica", 9, "italic"), justify="left", wraplength=340
                 ).pack(side='left', anchor='n', padx=(10, 0), pady=(30, 0))
        for var in (self.crescent_tonehole_var, self.crescent_percent_var):
            var.trace_add("write", lambda *a: self._update_crescent_preview())

        # Several at once: each list entry is a crescent with its own
        # quantity. With an empty list, Generate makes the one above.
        cr_list_row = tk.Frame(crescent_frame, bg=bg)
        cr_list_row.pack(fill='x', pady=(0, 3))
        tk.Label(cr_list_row, text=_("Quantity:"), bg=bg).pack(side='left')
        self.crescent_qty_var = tk.StringVar(value=str(tooling.get("crescent_qty", "1")))
        tk.Entry(cr_list_row, textvariable=self.crescent_qty_var, width=4).pack(side='left', padx=(2, 10))
        tk.Button(cr_list_row, text=_("Add to list"),
                  command=self._on_add_crescent).pack(side='left', padx=(0, 5))
        tk.Button(cr_list_row, text=_("Remove"),
                  command=self._on_remove_crescent).pack(side='left')
        self.crescent_listbox = tk.Listbox(crescent_frame, height=4, width=64,
                                           activestyle='none', exportselection=False)
        self.crescent_listbox.pack(anchor='w', pady=(0, 2))
        self.crescent_list_hint_var = tk.StringVar(value="")
        tk.Label(crescent_frame, textvariable=self.crescent_list_hint_var, bg=bg, fg="gray",
                 font=("Helvetica", 8), justify="left", wraplength=540).pack(anchor='w', pady=(0, 5))
        self.crescent_list = []
        saved_list = tooling.get("crescent_list", [])
        for raw in saved_list if isinstance(saved_list, list) else []:
            entry = self._clean_crescent_entry(raw)
            if entry:
                self.crescent_list.append(entry)
        self._refresh_crescent_list()

        cr_stl_row = tk.Frame(crescent_frame, bg=bg)
        cr_stl_row.pack(fill='x', pady=(0, 5))
        tk.Label(cr_stl_row, text=_("Print thickness:"), bg=bg).pack(side='left')
        self.crescent_thickness_var = tk.StringVar(value=str(tooling.get("crescent_thickness", "2.0")))
        tk.Entry(cr_stl_row, textvariable=self.crescent_thickness_var, width=6).pack(side='left', padx=(2, 2))
        tk.Label(cr_stl_row, text=_("mm  (STL only)"), bg=bg).pack(side='left', padx=(0, 15))
        # The G-code needs to know what it is cutting: speed, power and
        # the kerf it compensates for all come from the material.
        tk.Label(cr_stl_row, text=_("Material:"), bg=bg).pack(side='left')
        self.crescent_material_var = tk.StringVar(value=self._tooling_material_label(
            tooling.get("crescent_material"), _CRESCENT_MATERIALS))
        cr_mat_combo = ttk.Combobox(
            cr_stl_row, textvariable=self.crescent_material_var,
            values=[self._tooling_material_label(k) for k in _CRESCENT_MATERIALS],
            state="readonly", width=12)
        cr_mat_combo.pack(side='left', padx=(5, 5))
        cr_mat_note = tk.Label(cr_stl_row, text=_("(G-code only)"), bg=bg)
        cr_mat_note.pack(side='left')
        add_tooltip(cr_mat_combo,
                    _("What the crescents are cut from. The G-code takes its cut speed, power, "
                      "passes and kerf from this material, and cuts half a kerf outside the "
                      "outline so the part comes out at size."))

        cr_gen_frame = tk.Frame(crescent_frame, bg=bg)
        cr_gen_frame.pack(fill='x', pady=(5, 0))
        tk.Button(cr_gen_frame, text=_("Generate SVG"), width=15,
                  command=self._on_generate_crescent_svg).pack(side='left', padx=(0, 10))
        tk.Button(cr_gen_frame, text=_("Generate STL"), width=15,
                  command=self._on_generate_crescent_stl).pack(side='left', padx=(0, 10))
        tk.Button(cr_gen_frame, text=_("Generate G-code"), width=15,
                  command=self._on_generate_crescent_gcode).pack(side='left')
        self._add_tooling_frame_cut_button(
            cr_gen_frame, self._on_frame_cut_crescent, side='left', padx=(10, 0))

        self._tooling_sections['tonehole_crescent'] = crescent_frame
        self._update_crescent_preview()

        # Calibration Card UI was removed from the Tooling tab —
        # see the comment on the tool_names list. The integrated
        # Camera Calibration wizard (Pad Maker > Options > Machine)
        # owns the full engrave+capture workflow.

        # Nothing shown initially — user clicks a button
        self._active_tooling_section = None

    def _show_tooling_section(self, key):
        """Show the selected tooling section, hide others. Toggle off if already active."""
        # Toggle: clicking the active button hides it
        if self._active_tooling_section == key:
            self._tooling_sections[key].pack_forget()
            self._tooling_buttons[key].config(relief="raised")
            self._active_tooling_section = None
            return

        for section_key, frame in self._tooling_sections.items():
            if section_key == key:
                frame.pack(fill='both', expand=True, pady=(5, 0))
            else:
                frame.pack_forget()

        for btn_key, btn in self._tooling_buttons.items():
            if btn_key == key:
                btn.config(relief="sunken")
            else:
                btn.config(relief="raised")

        self._active_tooling_section = key

    # ========================================
    # SIZE PARSING
    # ========================================

    def _fill_die_sizes(self, range_text):
        """Fill the die size entry with a range string."""
        self.die_size_entry.delete("1.0", tk.END)
        self.die_size_entry.insert("1.0", range_text)

    def _parse_die_sizes(self):
        """Parse die size input text into a sorted list of float sizes."""
        text = self.die_size_entry.get("1.0", tk.END).strip()
        if not text:
            return []

        try:
            step = float(self.die_step_var.get())
        except (ValueError, TypeError):
            step = 0.5

        if step <= 0:
            step = 0.5

        sizes = set()
        for part in text.split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part:
                parts = part.split('-', 1)
                try:
                    lo = float(parts[0].strip())
                    hi = float(parts[1].strip())
                except ValueError:
                    continue
                if lo > hi:
                    lo, hi = hi, lo
                current = lo
                while current <= hi + 0.001:
                    sizes.add(round(current, 2))
                    current += step
            else:
                try:
                    sizes.add(float(part))
                except ValueError:
                    continue

        return sorted(sizes)

    def _sizes_to_pads(self, sizes):
        """Convert a list of sizes to pad dicts for nesting."""
        return [{'size': s, 'qty': 1} for s in sizes]

    # ========================================
    # SHEET SIZE HELPERS
    # ========================================

    def _get_die_sheet_mm(self):
        """Parse sheet width/height and convert to mm."""
        try:
            w = float(self.die_width_var.get())
            h = float(self.die_height_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"), _("Sheet width and height must be valid numbers."))
            return None, None

        if w <= 0 or h <= 0:
            messagebox.showerror(_("Invalid Input"), _("Sheet dimensions must be positive."))
            return None, None

        if self.die_units_var.get() == "in":
            w *= 25.4
            h *= 25.4

        return w, h

    # ========================================
    # TOOLING SETTINGS SYNC
    # ========================================

    def _open_tooling_gcode_settings(self):
        """Open settings dialog showing tooling G-code presets (acrylic
        + basswood) and die engraving options. Acrylic is the typical
        die-insert / die-holder material; basswood covers the camera-
        calibration card engrave AND the die organizer SVG (when run
        in LightBurn etc.) AND any future basswood prototyping. The
        calibration-card engrave reads from gcode_settings.basswood,
        so users can tune machine feeds here once."""
        from config import save_settings, save_presets, GCODE_PRESET_FILE
        tooling_materials = [(key, self._tooling_material_label(key))
                             for key in TOOLING_GCODE_MATERIALS]
        GcodeSettingsWindow(self.root, self.settings, lambda s: save_settings(s),
                            materials=tooling_materials, show_tooling_engraving=True,
                            gcode_presets=self.gcode_presets,
                            gcode_presets_save_callback=lambda: save_presets(self.gcode_presets, GCODE_PRESET_FILE))

    def _open_camera_calibration(self):
        """Open the one-time ChArUco camera-calibration wizard.

        Requires opencv-python and an engraved calibration card on the bed.
        Saves the resulting calibration JSON to the platform config dir;
        the "Get from camera" buttons (Pad Maker scrap mode / polygon draw)
        become available once the calibration file exists.
        """
        from ui_dialogs import CameraCalibrationDialog
        try:
            import camera_capture
            if not camera_capture.HAS_OPENCV:
                raise ImportError("OpenCV not loaded")
        except ImportError:
            messagebox.showerror(
                _("OpenCV Required"),
                _("Camera calibration needs OpenCV (opencv-python):\n\n"
                  "    pip install opencv-python Pillow"))
            return

        # If a pre-v3.0 calibration exists, warn the user that it's
        # incompatible with the new integrated workflow and will be
        # replaced by this run.
        cal_path = camera_capture.default_calibration_path()
        if camera_capture.is_legacy_calibration(cal_path):
            if not messagebox.askyesno(
                    _("Old Calibration Detected"),
                    _("An older calibration file is on disk that uses "
                      "the pre-v3.0 board-frame format. It will be "
                      "OVERWRITTEN by this new calibration.\n\n"
                      "Make sure you've engraved a fresh card via "
                      "Tooling > Engrave Calibration Card (which now "
                      "homes the laser and engraves at a known machine "
                      "position).\n\n"
                      "Continue?")):
                return

        # Camera index via the shared resolver (override → Falcon
        # heuristic → last enumerated); see PadSVGGeneratorApp.
        cam_idx = self._resolve_camera_index()
        if cam_idx is None:
            messagebox.showerror(_("No Camera"),
                                  _("No cameras were detected. Plug in "
                                    "the laser camera and try again."))
            return

        # The app's own live preview must not hold the camera this
        # dialog is about to open (main._release_live_camera).
        release = getattr(self, '_release_live_camera', None)
        if release is not None:
            release()
        CameraCalibrationDialog(
            self.root,
            camera_index=cam_idx,
            calibration_path=camera_capture.default_calibration_path(),
            falcon_port=getattr(self, 'falcon_port', None),
            settings=self.settings,
        )
        # Live-update the machine UI: if this was the user's first
        # successful calibration, the other Machine menu items un-grey
        # and Frame & Cut becomes packable. No restart required.
        try:
            self._refresh_machine_ui_state()
        except Exception:
            logging.getLogger(__name__).exception(
                "could not refresh the Machine menu after camera calibration")

    def _update_tooling_settings(self):
        """Sync tooling UI state back to settings dict."""
        # Merge tab-level controls into existing tooling_settings
        # (engraving mode/font/placement are saved by the settings dialog)
        tooling = self.settings.get("tooling_settings", {})
        tooling["sheet_width"] = self.die_width_var.get()
        tooling["sheet_height"] = self.die_height_var.get()
        tooling["engrave_ring"] = self.die_engrave_ring_var.get()
        tooling["engrave_cutout"] = self.die_engrave_cutout_var.get()
        tooling["step_size"] = self.die_step_var.get()
        if hasattr(self, 'holder_layer_count_var'):
            tooling["holder_layer_count"] = int(self.holder_layer_count_var.get())
            tooling["holder_sheet_width"] = self.holder_width_var.get()
            tooling["holder_sheet_height"] = self.holder_height_var.get()
        if hasattr(self, 'organizer_material_var'):
            tooling["organizer_material"] = self._organizer_material_key()
            tooling["organizer_sheet_width"] = self.organizer_width_var.get()
            tooling["organizer_sheet_height"] = self.organizer_height_var.get()
            tooling["organizer_sheet_units"] = self.organizer_units_var.get()
        if hasattr(self, 'crescent_material_var'):
            tooling["crescent_material"] = self._crescent_material_key()
        if hasattr(self, 'crescent_curve_var'):
            tooling["crescent_tonehole_d"] = self.crescent_tonehole_var.get()
            tooling["crescent_percent"] = self.crescent_percent_var.get()
            tooling["crescent_curve"] = round(float(self.crescent_curve_var.get()), 1)
            tooling["crescent_thickness"] = self.crescent_thickness_var.get()
            tooling["crescent_qty"] = self.crescent_qty_var.get()
            tooling["crescent_list"] = [dict(e) for e in self.crescent_list]
        self.settings["tooling_settings"] = tooling
        # Persist the Feeds & Speeds Tester state too (if the section is built)
        if hasattr(self, 'fs_sweep_vars'):
            self._persist_feeds_speeds_settings()

    def _update_holder_info(self):
        """Refresh the holder description label when layer count changes."""
        n = int(self.holder_layer_count_var.get())
        pin = n - 3
        self.holder_info_var.set(
            _("{n}-layer holder: base, magnet (6.5mm), {pin}× pin (3.5mm), retaining ring.").format(n=n, pin=pin)
        )

    def _get_holder_sheet_mm(self):
        """Parse holder sheet width/height and convert to mm."""
        try:
            w = float(self.holder_width_var.get())
            h = float(self.holder_height_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"), _("Sheet width and height must be valid numbers."))
            return None, None

        if w <= 0 or h <= 0:
            messagebox.showerror(_("Invalid Input"), _("Sheet dimensions must be positive."))
            return None, None

        if self.holder_units_var.get() == "in":
            w *= 25.4
            h *= 25.4

        return w, h

    def _get_die_settings(self):
        """Return settings dict with current tooling UI state applied."""
        self._update_tooling_settings()
        # Sync engraving mode from tooling settings to acrylic gcode settings
        tooling = self.settings.get("tooling_settings", {})
        gcode_settings = self.settings.get("gcode_settings", {})
        acrylic = gcode_settings.get("acrylic", {})
        acrylic["engraving_mode"] = tooling.get("engraving_mode", "filled")
        gcode_settings["acrylic"] = acrylic
        self.settings["gcode_settings"] = gcode_settings
        return self.settings

    # ========================================
    # DIE INSERT GENERATION
    # ========================================

    def _on_generate_die_svg(self):
        """Generate SVG for die inserts."""
        if self.die_scrap_var.get():
            self._generate_die_svg_scrap()
            return

        try:
            sizes = self._parse_die_sizes()
            if not sizes:
                messagebox.showerror(_("Error"), _("No valid die sizes entered."))
                return

            width_mm, height_mm = self._get_die_sheet_mm()
            if width_mm is None:
                return

            pads = self._sizes_to_pads(sizes)
            settings = self._get_die_settings()

            if not can_all_pads_fit(pads, 'die_ring', width_mm, height_mm, settings):
                messagebox.showerror(_("Nesting Error"),
                    _("Could not fit all {n} dies on the specified sheet.\n"
                      "Try a larger sheet or fewer sizes, or use Scrap Mode.").format(n=len(sizes)))
                return

            base = self.die_filename_var.get().strip() or "my_dies"
            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Insert SVG"),
                defaultextension=".svg",
                filetypes=[(_("SVG files"), "*.svg")],
                initialfile=f"{base}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            placed = generate_die_svg(pads, width_mm, height_mm, save_path, settings)
            self._record_tooling_export("svg", save_path, tool='die_inserts',
                                        description=self._die_description(placed), material='acrylic',
                                        count=len(placed), sheet_mm=(width_mm, height_mm))

            messagebox.showinfo(_("Done"),
                _("Generated {n} die rings.\n\nSaved to: {path}").format(n=len(placed), path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _validate_die_job(self):
        """The Die Inserts form, checked: (pads, width_mm, height_mm,
        settings), or None once the error has been shown. Generate G-code
        and Frame & Cut both start here, so they refuse the same input in
        the same words."""
        sizes = self._parse_die_sizes()
        if not sizes:
            messagebox.showerror(_("Error"), _("No valid die sizes entered."))
            return None

        width_mm, height_mm = self._get_die_sheet_mm()
        if width_mm is None:
            return None

        pads = self._sizes_to_pads(sizes)
        settings = self._get_die_settings()

        if not can_all_pads_fit(pads, 'die_ring', width_mm, height_mm, settings):
            messagebox.showerror(_("Nesting Error"),
                _("Could not fit all {n} dies on the specified sheet.\n"
                  "Try a larger sheet or fewer sizes, or use Scrap Mode.").format(n=len(sizes)))
            return None
        return pads, width_mm, height_mm, settings

    def _on_generate_die_gcode(self):
        """Generate G-code for die inserts."""
        if self.die_scrap_var.get():
            self._generate_die_gcode_scrap()
            return

        try:
            job = self._validate_die_job()
            if job is None:
                return
            pads, width_mm, height_mm, settings = job

            base = self.die_filename_var.get().strip() or "my_dies"
            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Insert G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)

            placed, _unused_a, _unused_b = _nest_discs(pads, 'die_ring', width_mm, height_mm, settings)
            generate_die_gcode_from_placed(placed, width_mm, height_mm, save_path, settings)
            self._record_tooling_export("gcode", save_path, tool='die_inserts',
                                        description=self._die_description(placed), material='acrylic',
                                        count=len(placed), sheet_mm=(width_mm, height_mm))

            messagebox.showinfo(_("Done"),
                _("Generated G-code for {n} die rings.\n\nSaved to: {path}").format(n=len(placed), path=save_path)
                + "\n\n" + gcode_file_note())

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    # ========================================
    # JOB HISTORY
    # ========================================
    # Matt, 2026-11-09: "tooling runs go into job history if easy". Every
    # Tooling output (a file written, a Frame & Cut run) is logged with the
    # Pad Maker's jobs, as an entry of the same shape with a `tool` key and
    # a one-line description in the user's terms where the pad list goes
    # (main._record_tooling_job). The dialog shows the tool's name on those
    # rows and will not load them into the Pad Maker's form. Nothing here
    # changes what is written or streamed: each line below runs after the
    # file exists or the cut dialog has closed, and the recorder swallows
    # its own errors.

    @staticmethod
    def _tooling_history(tool, description, material, count, sheet_mm=None, base=None, scrap_num=None):
        """The history fields of a Tooling job: merged into a Frame & Cut
        job dict, or handed to _record_tooling_export after a write."""
        w, h = sheet_mm if sheet_mm else (None, None)
        return {'tool': tool, 'description': description, 'material': material, 'count': count,
                'sheet_w_mm': w, 'sheet_h_mm': h, 'base': base, 'scrap_num': scrap_num}

    def _record_tooling_export(self, output, save_path, **fields):
        """One line after a Tooling file is written: `output` is "svg",
        "gcode" or "stl"; the file's own name is the entry's base."""
        fields.setdefault('base', os.path.splitext(os.path.basename(save_path))[0])
        self._record_tooling_job(self._tooling_history(**fields), output,
                                 save_dir=os.path.dirname(save_path))

    @staticmethod
    def _die_description(placed):
        """The die sizes on the sheet, in mm, with a count where a size repeats."""
        counts = Counter(float(p[0]) for p in placed)
        return ", ".join(f"{s:g}" + (f" ×{n}" if n > 1 else "") for s, n in sorted(counts.items())) + " mm"

    @staticmethod
    def _holder_description(variant, layer_count):
        variant_label = {"large": _("Large (40–60mm)"), "small": _("Small (7–39.5mm)"),
                         "both": _("Both (Small + Large)")}.get(variant, str(variant))
        layer_label = {5: _("5-layer (2× pin)"), 6: _("6-layer (3× pin)")}.get(layer_count, str(layer_count))
        return f"{variant_label}, {layer_label}"

    @staticmethod
    def _organizer_description(variant):
        return {"upper": _("Upper (slotted)"), "lower": _("Lower (base)")}.get(variant, str(variant))

    @staticmethod
    def _kerf_description(job):
        """The cut settings the pattern was cut with: that is what a kerf test measures."""
        return f"{job['cut_speed']:g} mm/min, {job['cut_power']:g}%"

    @staticmethod
    def _feeds_speeds_description(test_pieces):
        """The matrix: disc size, then the speeds and powers swept."""
        def values(key):
            return ", ".join(f"{v:g}" for v in sorted({p[key] for p in test_pieces}))
        diameter = test_pieces[0]['diameter'] if test_pieces else 0
        return (_("Disc diameter: {d:.1f} mm").format(d=diameter)
                + f"; {values('speed')} mm/min; {values('power')}%")

    @staticmethod
    def _crescent_description(entries):
        """The list as the tab shows it, one entry after another."""
        rows = []
        for n, e in enumerate(entries, 1):
            g = tonehole_crescent_geometry(e['tonehole_d'], e['percent'], e['curve'])
            rows.append(_("{n}.  {d} mm, {pct}%, inner edge radius {r} mm, × {qty}").format(
                n=n, d=f"{e['tonehole_d']:g}", pct=f"{e['percent']:g}",
                r=f"{g['inner_radius']:.1f}", qty=e['qty']))
        return "; ".join(rows)

    # ========================================
    # FRAME & CUT (straight to the laser)
    # ========================================

    @staticmethod
    def _gcode_text_from(write):
        """Run `write(path)`, a generator call that saves a .gcode file,
        and return the text it wrote.

        The path is a temp file that is gone again before this returns.
        That is how Frame & Cut streams exactly what the file export would
        have saved (the same generator call, read back) and leaves nothing
        in the user's folders. Read with the encoding the generators write
        with, the platform default, as the Pad Maker's Frame & Cut does.
        """
        with tempfile.NamedTemporaryFile(suffix='.gcode', delete=False, mode='w') as tmp:
            path = tmp.name
        try:
            write(path)
            with open(path, 'r') as f:
                return f.read()
        finally:
            try:
                os.remove(path)
            except OSError:
                pass

    def _on_frame_cut_die(self):
        """Frame & Cut for die inserts: the Generate G-code job, sent to
        the laser. Acrylic, like the file it would have written. In Scrap
        Mode it is one sheet per click."""
        if self.die_scrap_var.get():
            self.on_frame_and_cut(tooling_job=self._die_scrap_machine_job)
            return

        def prepare():
            job = self._validate_die_job()
            if job is None:
                return None
            pads, width_mm, height_mm, settings = job
            placed, _unused_a, _unused_b = _nest_discs(pads, 'die_ring', width_mm, height_mm, settings)
            text = self._gcode_text_from(
                lambda path: generate_die_gcode_from_placed(
                    placed, width_mm, height_mm, path, settings))
            return {'gcode_text': text, 'material': 'acrylic',
                    'label': 'die inserts', 'count': len(placed),
                    **self._tooling_history('die_inserts', self._die_description(placed), 'acrylic',
                                            len(placed), sheet_mm=(width_mm, height_mm),
                                            base=self.die_filename_var.get().strip() or None)}
        self.on_frame_and_cut(tooling_job=prepare)

    def _die_scrap_machine_job(self):
        """The Frame & Cut job for one sheet of a Die Scrap Mode session.

        Preparing it counts nothing. The sheet is counted, and the dies on
        it come off Remaining, only when its cut has streamed to the end:
        a run cancelled at the jog step or during framing, or a cut that
        was stopped or errored, leaves the session exactly as it was, so
        the same dies are offered again (the Pad Maker's rule, see
        main._frame_cut_scrap_advance).
        """
        begun = self._die_scrap_begin(None)
        if begun is None:
            return None
        placed, remaining, width_mm, height_mm, settings = begun
        text = self._gcode_text_from(
            lambda path: generate_die_gcode_from_placed(
                placed, width_mm, height_mm, path, settings))

        def after_cut(reason):
            if reason != "complete":
                return
            session = self.tooling_scrap_session
            session['scrap_count'] += 1
            scrap_num = session['scrap_count']
            session['remaining_pads'] = remaining
            self._update_die_scrap_status()
            self._update_die_scrap_window()

            placed_count = len(placed)
            remaining_count = self._count_die_remaining()
            if remaining_count == 0:
                from config import save_settings
                save_settings(self.settings)
                messagebox.showinfo(_("Session Complete!"),
                    _("Cut {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "All dies cut! Session complete.").format(
                          placed_count=placed_count, scrap_num=scrap_num))
            else:
                messagebox.showinfo(_("Sheet Cut"),
                    _("Cut {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "{remaining_count} dies remaining.\n"
                      "Load the next sheet, adjust the sheet size if needed "
                      "and click Frame & Cut again.").format(
                          placed_count=placed_count, scrap_num=scrap_num,
                          remaining_count=remaining_count))

        return {'gcode_text': text, 'material': 'acrylic',
                'label': 'die inserts', 'count': len(placed),
                'after_cut': after_cut,
                **self._tooling_history('die_inserts', self._die_description(placed), 'acrylic',
                                        len(placed), sheet_mm=(width_mm, height_mm),
                                        base=self.die_filename_var.get().strip() or None,
                                        scrap_num=self.tooling_scrap_session['scrap_count'] + 1)}

    # ========================================
    # DIE SCRAP MODE
    # ========================================

    def _toggle_die_scrap_mode(self):
        """Handle die scrap mode checkbox toggle."""
        if not self.die_scrap_var.get():
            if self.tooling_scrap_session['active']:
                remaining = self._count_die_remaining()
                if remaining > 0:
                    if not messagebox.askyesno(_("Clear Session?"),
                        _("You have {remaining} dies remaining.\n"
                          "Disabling scrap mode will clear the session.\n\nContinue?").format(remaining=remaining)):
                        self.die_scrap_var.set(True)
                        return
                self._clear_die_scrap_session()
        self._update_die_scrap_status()

    def _clear_die_scrap_session(self):
        """Reset die scrap session state."""
        self.tooling_scrap_session = {
            'active': False,
            'original_pads': [],
            'remaining_pads': [],
            'scrap_count': 0,
            'save_dir': '',
        }
        self._close_die_scrap_window()
        self._update_die_scrap_status()

    def _on_clear_die_scrap_clicked(self):
        """Handle Clear button for die scrap mode."""
        if self.tooling_scrap_session['active']:
            remaining = self._count_die_remaining()
            if remaining > 0:
                if not messagebox.askyesno(_("Clear Session?"),
                    _("You have {remaining} dies remaining.\n\nClear session?").format(remaining=remaining)):
                    return
        self._clear_die_scrap_session()
        self.die_scrap_var.set(False)
        self._update_die_scrap_status()

    def _count_die_remaining(self):
        """Count total remaining dies in session."""
        return sum(p['qty'] for p in self.tooling_scrap_session['remaining_pads'])

    def _update_die_scrap_status(self):
        """Update scrap mode status display."""
        if self.die_scrap_var.get():
            self.die_scrap_clear_btn.pack(side='left', padx=(10, 0))
            if self.tooling_scrap_session['active']:
                remaining = self._count_die_remaining()
                if remaining == 0:
                    self.die_scrap_status_var.set("Done!")
                    self.die_scrap_status_label.config(fg="green")
                else:
                    count = self.tooling_scrap_session['scrap_count']
                    self.die_scrap_status_var.set(f"{remaining} left ({count} sheets)")
                    self.die_scrap_status_label.config(fg="blue")
                self.die_scrap_status_label.pack(side='left', padx=(10, 0))
            else:
                self.die_scrap_status_label.pack_forget()
        else:
            self.die_scrap_status_label.pack_forget()
            self.die_scrap_clear_btn.pack_forget()

    def _start_die_scrap_session(self, pads, save_dir):
        """Initialize a new die scrap session."""
        fixed_pads = [{'size': p['size'], 'qty': p['qty']} for p in pads]
        self.tooling_scrap_session = {
            'active': True,
            'original_pads': [p.copy() for p in fixed_pads],
            'remaining_pads': fixed_pads,
            'scrap_count': 0,
            'save_dir': save_dir,
        }
        self._open_die_scrap_window()

    def _die_scrap_begin(self, folder_title):
        """Die Scrap Mode's front half, shared by Generate SVG, Generate
        G-code and Frame & Cut so the session can't drift between them:
        start the session or carry on with it, then nest what fits on
        this sheet.

        folder_title: the title of the folder prompt for the two file
        exports, asked when a session starts. None for Frame & Cut, which
        writes nothing and starts the session with no folder; a file
        export joining such a session asks for its folder then, once.

        Returns (placed, remaining, width_mm, height_mm, settings), or
        None when there is nothing to do (every way out has shown its own
        message). Nothing is counted here: the caller counts the sheet
        once its file is written or its cut has finished.
        """
        width_mm, height_mm = self._get_die_sheet_mm()
        if width_mm is None:
            return None

        settings = self._get_die_settings()

        if not self.tooling_scrap_session['active']:
            # Starting new session
            sizes = self._parse_die_sizes()
            if not sizes:
                messagebox.showerror(_("Error"), _("No valid die sizes entered."))
                return None
            pads = self._sizes_to_pads(sizes)

            save_dir = ''
            if folder_title:
                save_dir = filedialog.askdirectory(
                    title=folder_title,
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return None
                self.settings["last_output_dir"] = save_dir
            self._start_die_scrap_session(pads, save_dir)
        else:
            if folder_title and not self.tooling_scrap_session.get('save_dir'):
                save_dir = filedialog.askdirectory(
                    title=folder_title,
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return None
                self.settings["last_output_dir"] = save_dir
                self.tooling_scrap_session['save_dir'] = save_dir
            pads = self.tooling_scrap_session['remaining_pads']

        if not pads:
            messagebox.showinfo(_("Session Complete"), _("All dies have been placed!"))
            return None

        placed, remaining, any_placed = try_nest_partial(
            pads, 'die_ring', width_mm, height_mm, settings)

        if not any_placed:
            min_size = min(p['size'] for p in pads)
            messagebox.showwarning(_("No Dies Fit"),
                _("No dies could be placed on this sheet.\n\n"
                  "Smallest remaining die: {min_size}mm\n"
                  "Try a larger sheet.").format(min_size=min_size))
            return None
        return placed, remaining, width_mm, height_mm, settings

    def _generate_die_svg_scrap(self):
        """Handle SVG generation in die scrap mode."""
        try:
            begun = self._die_scrap_begin(_("Select Folder to Save SVGs"))
            if begun is None:
                return
            placed, remaining, width_mm, height_mm, settings = begun

            self.tooling_scrap_session['scrap_count'] += 1
            scrap_num = self.tooling_scrap_session['scrap_count']
            save_dir = self.tooling_scrap_session['save_dir']
            base = self.die_filename_var.get().strip() or "my_dies"
            filename = os.path.join(save_dir, f"{base}_scrap{scrap_num}.svg")

            generate_die_svg_from_placed(placed, width_mm, height_mm, filename, settings)
            self._record_tooling_export("svg", filename, tool='die_inserts',
                                        description=self._die_description(placed), material='acrylic',
                                        count=len(placed), sheet_mm=(width_mm, height_mm), scrap_num=scrap_num)

            self.tooling_scrap_session['remaining_pads'] = remaining
            self._update_die_scrap_status()
            self._update_die_scrap_window()

            placed_count = len(placed)
            remaining_count = self._count_die_remaining()

            if remaining_count == 0:
                from config import save_settings
                save_settings(self.settings)
                messagebox.showinfo(_("Session Complete!"),
                    _("Placed {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "All dies placed! Session complete.\n"
                      "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir))
            else:
                messagebox.showinfo(_("Sheet Generated"),
                    _("Placed {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "{remaining_count} dies remaining.\n"
                      "Adjust sheet size if needed and generate again.").format(placed_count=placed_count, scrap_num=scrap_num, remaining_count=remaining_count))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _generate_die_gcode_scrap(self):
        """Handle G-code generation in die scrap mode."""
        try:
            begun = self._die_scrap_begin(_("Select Folder to Save G-code"))
            if begun is None:
                return
            placed, remaining, width_mm, height_mm, settings = begun

            self.tooling_scrap_session['scrap_count'] += 1
            scrap_num = self.tooling_scrap_session['scrap_count']
            save_dir = self.tooling_scrap_session['save_dir']
            base = self.die_filename_var.get().strip() or "my_dies"
            filename = os.path.join(save_dir, f"{base}_scrap{scrap_num}.gcode")

            generate_die_gcode_from_placed(placed, width_mm, height_mm, filename, settings)
            self._record_tooling_export("gcode", filename, tool='die_inserts',
                                        description=self._die_description(placed), material='acrylic',
                                        count=len(placed), sheet_mm=(width_mm, height_mm), scrap_num=scrap_num)

            self.tooling_scrap_session['remaining_pads'] = remaining
            self._update_die_scrap_status()
            self._update_die_scrap_window()

            placed_count = len(placed)
            remaining_count = self._count_die_remaining()

            if remaining_count == 0:
                from config import save_settings
                save_settings(self.settings)
                messagebox.showinfo(_("Session Complete!"),
                    _("Generated G-code for {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "All dies placed! Session complete.\n"
                      "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir)
                    + "\n\n" + gcode_file_note())
            else:
                messagebox.showinfo(_("Sheet Generated"),
                    _("Generated G-code for {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "{remaining_count} dies remaining.\n"
                      "Adjust sheet size if needed and generate again.").format(placed_count=placed_count, scrap_num=scrap_num, remaining_count=remaining_count)
                    + "\n\n" + gcode_file_note())

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    # ========================================
    # DIE SCRAP PROGRESS WINDOW
    # ========================================

    def _open_die_scrap_window(self):
        """Open floating progress window for die scrap mode."""
        if self.tooling_scrap_window is not None:
            try:
                self.tooling_scrap_window.destroy()
            except tk.TclError:
                pass

        self.tooling_scrap_window = tk.Toplevel(self.root)
        self.tooling_scrap_window.title(_("Die Scrap Mode Progress"))
        self.tooling_scrap_window.geometry("320x300")
        theme_bg = self._get_theme_color()
        self.tooling_scrap_window.configure(bg=theme_bg)
        self.tooling_scrap_window.resizable(True, True)

        self.root.update_idletasks()
        x = self.root.winfo_x() + self.root.winfo_width() + 10
        y = self.root.winfo_y()
        self.tooling_scrap_window.geometry(f"+{x}+{y}")

        header_frame = tk.Frame(self.tooling_scrap_window, bg=theme_bg)
        header_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.die_scrap_header = tk.Label(header_frame, text="", bg=theme_bg,
                                          font=("Helvetica", 10, "bold"))
        self.die_scrap_header.pack(anchor="w")

        columns_frame = tk.Frame(self.tooling_scrap_window, bg=theme_bg)
        columns_frame.pack(fill="both", expand=True, padx=10, pady=5)

        remaining_frame = tk.Frame(columns_frame, bg=theme_bg)
        remaining_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        tk.Label(remaining_frame, text=_("Remaining"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.die_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.die_remaining_listbox.pack(fill="both", expand=True)

        done_frame = tk.Frame(columns_frame, bg=theme_bg)
        done_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        tk.Label(done_frame, text=_("Done"), bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="green").pack(anchor="w")
        self.die_done_listbox = tk.Listbox(done_frame, font=("Courier", 10), height=10, width=12)
        self.die_done_listbox.pack(fill="both", expand=True)

        self.die_scrap_footer = tk.Label(self.tooling_scrap_window, text="", bg=theme_bg,
                                          font=("Helvetica", 9))
        self.die_scrap_footer.pack(pady=(0, 5))

        tk.Button(self.tooling_scrap_window, text=_("Close"),
                  command=self._close_die_scrap_window).pack(pady=(0, 10))

        self.tooling_scrap_window.protocol("WM_DELETE_WINDOW", self._close_die_scrap_window)
        self._update_die_scrap_window()

    def _update_die_scrap_window(self):
        """Update die scrap progress window contents."""
        if self.tooling_scrap_window is None:
            if self.tooling_scrap_session.get('active', False):
                self._open_die_scrap_window()
            return
        try:
            remaining_pads = self.tooling_scrap_session.get('remaining_pads', [])
            original_pads = self.tooling_scrap_session.get('original_pads', [])
            scraps = self.tooling_scrap_session.get('scrap_count', 0)

            remaining_by_size = {p['size']: p['qty'] for p in remaining_pads}
            done_pads = []
            for orig in original_pads:
                size = orig['size']
                orig_qty = orig['qty']
                remaining_qty = remaining_by_size.get(size, 0)
                done_qty = orig_qty - remaining_qty
                if done_qty > 0:
                    done_pads.append({'size': size, 'qty': done_qty})

            total_remaining = sum(p['qty'] for p in remaining_pads)
            total_done = sum(p['qty'] for p in done_pads)
            total_original = sum(p['qty'] for p in original_pads)

            if total_remaining == 0:
                self.die_scrap_header.config(text=_("All done!"), fg="green")
            else:
                self.die_scrap_header.config(
                    text=_("{done} / {total} dies complete").format(done=total_done, total=total_original), fg="blue")

            self.die_remaining_listbox.delete(0, tk.END)
            if remaining_pads:
                for pad in sorted(remaining_pads, key=lambda p: -p['size']):
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.die_remaining_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.die_remaining_listbox.insert(tk.END, " (none)")

            self.die_done_listbox.delete(0, tk.END)
            if done_pads:
                for pad in sorted(done_pads, key=lambda p: -p['size']):
                    size_str = f"{pad['size']:.1f}".rstrip('0').rstrip('.')
                    self.die_done_listbox.insert(tk.END, f" {pad['qty']} x {size_str}")
            else:
                self.die_done_listbox.insert(tk.END, " (none)")

            self.die_scrap_footer.config(text=_("die_ring | {scraps} sheet(s) used").format(scraps=scraps))

        except tk.TclError:
            self.tooling_scrap_window = None

    def _close_die_scrap_window(self):
        """Close the die scrap progress window."""
        if self.tooling_scrap_window is not None:
            try:
                self.tooling_scrap_window.destroy()
            except tk.TclError:
                pass
            self.tooling_scrap_window = None

    # ========================================
    # KERF TEST GENERATION
    # ========================================

    def _toggle_kerf_custom_settings(self):
        """Show/hide custom speed/power fields based on checkbox."""
        if self.kerf_use_existing_var.get():
            self.kerf_custom_frame.pack_forget()
        else:
            self.kerf_custom_frame.pack(fill='x', pady=(0, 5))

    @staticmethod
    def _tooling_material_key(display):
        """Settings key ('felt', 'acrylic', ...) for a material as a Tooling
        dropdown shows it.

        The dropdowns list translated names, and the key used to be the
        shown name lower-cased, which is only right in English: with the
        app in German, "Filz" looked up gcode_settings['filz'], found
        nothing, and the kerf test ran at the fallback 180 mm/min and 100 %
        instead of the user's felt settings (found 2026-11-09 while adding
        Frame & Cut here, where that goes straight to the laser). Matches
        the translated name first, then the English one; anything else
        falls back to the old lower-casing.
        """
        shown = (display or "").strip()
        for english in _TOOLING_MATERIALS:
            if shown == _(english) or shown.lower() == english.lower():
                return english.lower()
        return shown.lower()

    @staticmethod
    def _tooling_material_label(key, choices=None):
        """The name a Tooling dropdown shows for a material key, in the
        app's language. With `choices`, a key that is missing or not one of
        them gives the first choice's name (the default for that dropdown)."""
        if choices is not None and key not in choices:
            key = choices[0]
        for english in _TOOLING_MATERIALS:
            if english.lower() == key:
                return _(english)
        return str(key).replace('_', ' ').title()

    @staticmethod
    def _material_settings_place(material_key):
        """Where a material's speeds and powers are edited, as the user
        would follow it: the Tooling dialog for acrylic, basswood and cork,
        the Pad Maker's for felt, card and leather."""
        if material_key in TOOLING_GCODE_MATERIALS:
            return _("Tooling > Options > Settings")
        return _("Pad Maker > Options > G-code Settings")

    def _engraving_not_set_notice(self, material_key, what):
        """The one line said when engraving was left out because the
        material has no engraving speed and power (0 means not set: cork
        ships that way). Nothing is ever skipped silently: this goes in
        the Done box for a file and is asked before a Frame & Cut run.
        `what` is 'numbers' (crescent list), 'labels' (kerf test) or 'ids'
        (Speed & Power, whose values come from its own Engraving fields).
        """
        if what == 'ids':
            return _("The engraving speed and power are 0, so the disc IDs will not be "
                     "engraved. Fill in the Engraving fields to label the discs.")
        material = self._tooling_material_label(material_key)
        where = self._material_settings_place(material_key)
        if what == 'numbers':
            return _("{material} has no engraving speed and power set ({where}), so the "
                     "crescents will not be numbered.").format(material=material, where=where)
        return _("{material} has no engraving speed and power set ({where}), so the pattern's "
                 "title and sizes will not be engraved.").format(material=material, where=where)

    def _confirm_engraving_skipped(self, notice):
        """Before a Frame & Cut run that leaves engraving out: say so, and
        let the user back out. True to go on."""
        return bool(messagebox.askokcancel(_("Engraving Not Set"), notice))

    def _get_kerf_material_key(self):
        """Map display name to settings key."""
        return self._tooling_material_key(self.kerf_material_var.get())

    def _show_kerf_instructions(self, material_name):
        """Show kerf test measurement instructions with a don't-show-again option."""
        if self.settings.get("seen_kerf_test_tutorial", False):
            return

        from config import save_settings

        mat_key = self._tooling_material_key(material_name)
        # Map to settings location
        if mat_key == "acrylic":
            settings_path = _("Tooling > Options > Settings > Acrylic > Kerf Width")
        elif mat_key in TOOLING_GCODE_MATERIALS:
            settings_path = _("Tooling > Options > Settings > {material_name} > Kerf Width").format(material_name=material_name)
        else:
            settings_path = _("Pad Maker > Options > G-code Settings > {material_name} > Kerf Width").format(material_name=material_name)

        dlg = tk.Toplevel(self.root)
        dlg.title(_("Kerf Test \u2014 {material_name}").format(material_name=material_name))
        dlg.geometry("420x310")
        theme_bg = self._get_theme_color()
        dlg.configure(bg=theme_bg)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text=_("Kerf Test \u2014 {material_name}").format(material_name=material_name), bg=theme_bg,
                 font=("Helvetica", 12, "bold")).pack(pady=(15, 10))

        steps = [
            _("1.  Cut the pattern on your material"),
            _("2.  Pop out the 3 discs (10mm, 20mm, 30mm)"),
            _("3.  For each circle, measure the hole ID\n"
              "     and disc OD with calipers"),
            _("4.  Kerf = hole ID \u2212 disc OD"),
            _("5.  Average the 3 results for best accuracy"),
            _("6.  Enter the full kerf value in:\n     {settings_path}").format(settings_path=settings_path),
        ]
        for step in steps:
            tk.Label(dlg, text=step, bg=theme_bg, font=("Helvetica", 10),
                     justify="left", anchor="w").pack(fill="x", padx=20, pady=1)

        dont_show_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dlg, text=_("Don't show this again"), variable=dont_show_var,
                       bg=theme_bg).pack(pady=(10, 5))

        def on_ok():
            if dont_show_var.get():
                self.settings["seen_kerf_test_tutorial"] = True
                save_settings(self.settings)
            dlg.destroy()

        tk.Button(dlg, text=_("OK"), command=on_ok, width=10).pack(pady=(5, 15))

    def _on_generate_kerf_svg(self):
        """Generate SVG kerf test pattern."""
        try:
            material_name = self.kerf_material_var.get()
            save_path = filedialog.asksaveasfilename(
                title=_("Save Kerf Test SVG"),
                defaultextension=".svg",
                filetypes=[(_("SVG files"), "*.svg")],
                initialfile=f"kerf_test_{material_name.lower()}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            generate_kerf_test_svg(material_name, save_path, self.settings)
            # The SVG carries no cut settings (LightBurn's are used), so its
            # entry has no description beyond the material.
            self._record_tooling_export("svg", save_path, tool='kerf_test', description="",
                                        material=self._get_kerf_material_key(), count=len(KERF_TEST_DIAMETERS))
            self._show_kerf_instructions(material_name)

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _kerf_gcode_job(self):
        """The Kerf Test form's G-code inputs, checked: a dict of what
        generate_kerf_test_gcode takes, or None once the error has been
        shown. Generate G-code and Frame & Cut both start here."""
        material_name = self.kerf_material_var.get()
        mat_key = self._get_kerf_material_key()

        # Get cut speed/power
        if self.kerf_use_existing_var.get():
            gcode_settings = self.settings.get("gcode_settings", {})
            mat_settings = gcode_settings.get(mat_key, {})
            cut_speed = mat_settings.get("cut_speed", 180)
            cut_power = mat_settings.get("cut_power", 100)
            eng_mode = mat_settings.get("engraving_mode", "line")
            if eng_mode == "filled":
                eng_speed = mat_settings.get("filled_engraving_speed", 1000)
                eng_power = mat_settings.get("filled_engraving_power", 15)
            else:
                eng_speed = mat_settings.get("engraving_speed", 1000)
                eng_power = mat_settings.get("engraving_power", 15)
            filled_spacing = mat_settings.get("filled_line_spacing", 0.15)
        else:
            try:
                cut_speed = float(self.kerf_speed_var.get())
                cut_power = float(self.kerf_power_var.get())
            except (ValueError, TypeError):
                messagebox.showerror(_("Invalid Input"), _("Speed and power must be valid numbers."))
                return None
            eng_mode = "line"
            eng_speed = 1000
            eng_power = 15
            filled_spacing = 0.15
        return {
            'material_name': material_name, 'mat_key': mat_key,
            'cut_speed': cut_speed, 'cut_power': cut_power,
            'eng_speed': eng_speed, 'eng_power': eng_power,
            'eng_mode': eng_mode, 'filled_spacing': filled_spacing,
            # The material has no engraving settings (cork): the circles
            # are cut, the title and sizes are left off, and the user is told.
            'labels_skipped': not engraving_is_set(eng_speed, eng_power),
        }

    def _write_kerf_gcode(self, job, path):
        """Write the kerf test G-code for a _kerf_gcode_job() to `path`."""
        generate_kerf_test_gcode(job['material_name'], path, self.settings,
                                 job['cut_speed'], job['cut_power'],
                                 job['eng_speed'], job['eng_power'],
                                 engraving_mode=job['eng_mode'],
                                 filled_line_spacing=job['filled_spacing'])

    def _on_generate_kerf_gcode(self):
        """Generate G-code kerf test pattern."""
        try:
            job = self._kerf_gcode_job()
            if job is None:
                return
            material_name = job['material_name']

            save_path = filedialog.asksaveasfilename(
                title=_("Save Kerf Test G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                # Named from what the dropdown shows, as the SVG is, so the
                # file name a translated app suggests stays what it was.
                initialfile=f"kerf_test_{material_name.lower()}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            self._write_kerf_gcode(job, save_path)
            self._record_tooling_export("gcode", save_path, tool='kerf_test',
                                        description=self._kerf_description(job), material=job['mat_key'],
                                        count=len(KERF_TEST_DIAMETERS))
            # A Done box of its own (2026-11-09): the measuring steps below
            # stop showing once "Don't show this again" is ticked, and then
            # this export said nothing at all, not even that it had worked.
            parts = [_("Saved to: {d}").format(d=save_path)]
            if job['labels_skipped']:
                parts.append(self._engraving_not_set_notice(job['mat_key'], 'labels'))
            parts.append(gcode_file_note())
            messagebox.showinfo(_("Done"), "\n\n".join(parts))
            self._show_kerf_instructions(material_name)

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_frame_cut_kerf(self):
        """Frame & Cut for the kerf test: the Generate G-code job, sent to
        the laser. Frames at the selected material's framing power."""
        def prepare():
            job = self._kerf_gcode_job()
            if job is None:
                return None
            if job['labels_skipped'] and not self._confirm_engraving_skipped(
                    self._engraving_not_set_notice(job['mat_key'], 'labels')):
                return None
            text = self._gcode_text_from(lambda path: self._write_kerf_gcode(job, path))

            def after_cut(reason):
                # The measuring steps belong after a pattern that was cut.
                if reason == "complete":
                    self._show_kerf_instructions(job['material_name'])
            return {'gcode_text': text, 'material': job['mat_key'],
                    'label': 'kerf test', 'count': len(KERF_TEST_DIAMETERS),
                    'after_cut': after_cut,
                    **self._tooling_history('kerf_test', self._kerf_description(job), job['mat_key'],
                                            len(KERF_TEST_DIAMETERS))}
        self.on_frame_and_cut(tooling_job=prepare)

    # ========================================
    # FEEDS & SPEEDS TESTER
    # ========================================

    def _get_feeds_speeds_material_key(self):
        """Map display name to settings key (mirrors kerf-test pattern)."""
        return self._tooling_material_key(self.fs_material_var.get())

    def _get_feeds_speeds_sheet_mm(self):
        """Parse the feeds-speeds sheet W/H into mm. Returns (w, h) or (None, None)."""
        try:
            w = float(self.fs_sheet_w_var.get())
            h = float(self.fs_sheet_h_var.get())
        except (ValueError, TypeError):
            return None, None
        if w <= 0 or h <= 0:
            return None, None
        if self.fs_sheet_unit_var.get() == "in":
            w *= 25.4
            h *= 25.4
        return w, h

    def _apply_feeds_speeds_material_defaults(self):
        """Pre-fill cut sweep AND engraving fields from the material's gcode_settings."""
        mat_key = self._get_feeds_speeds_material_key()
        mat_settings = self.settings.get("gcode_settings", {}).get(mat_key, {})
        cut_speed = mat_settings.get("cut_speed", 180)
        cut_power = mat_settings.get("cut_power", 60)

        self.fs_sweep_vars['speed']['value'].set(str(int(round(cut_speed))))
        self.fs_sweep_vars['speed']['start'].set(str(int(round(cut_speed * 0.6))))
        self.fs_sweep_vars['speed']['end'].set(str(int(round(cut_speed * 1.4))))

        self.fs_sweep_vars['power']['value'].set(str(int(round(cut_power))))
        # Power can't exceed 100%; floor at 1%.
        self.fs_sweep_vars['power']['start'].set(
            str(max(1, int(round(cut_power * 0.6)))))
        self.fs_sweep_vars['power']['end'].set(
            str(min(100, int(round(cut_power * 1.4)))))

        # Engraving — pick the right pair (line vs filled) based on the
        # material's currently-selected engraving_mode.
        eng_mode = mat_settings.get("engraving_mode", "line")
        if eng_mode == "filled":
            eng_speed = mat_settings.get("filled_engraving_speed", 1200)
            eng_power = mat_settings.get("filled_engraving_power", 10)
        else:
            eng_speed = mat_settings.get("engraving_speed", 1200)
            eng_power = mat_settings.get("engraving_power", 10)
        self.fs_eng_speed_var.set(str(int(round(eng_speed))))
        # A material with no engraving power (cork ships at 0, "not set")
        # pre-fills as 0, not as a made-up 1 %: the test then runs without
        # disc IDs and says so, or the user types a value of their own.
        if eng_power <= 0:
            self.fs_eng_power_var.set("0")
        else:
            self.fs_eng_power_var.set(str(max(1, min(100, int(round(eng_power))))))

    def _feeds_speeds_sweep_count(self, var_name):
        """How many discs this variable contributes: 1 if not swept, else stops (clamped 2-10)."""
        v = self.fs_sweep_vars[var_name]
        if not v['sweep'].get():
            return 1
        try:
            stops = int(v['stops'].get())
        except (ValueError, TypeError):
            return 1
        return max(2, min(10, stops))

    def _feeds_speeds_sweep_cfg(self, var_name):
        """Read one sweep row's tk vars into the dict shape expected by build_feeds_speeds_matrix."""
        v = self.fs_sweep_vars[var_name]
        try:
            stops = int(v['stops'].get())
        except (ValueError, TypeError):
            stops = 4
        return {
            'sweep': bool(v['sweep'].get()),
            'value': v['value'].get(),
            'start': v['start'].get(),
            'end': v['end'].get(),
            'stops': stops,
        }

    def _build_feeds_speeds_matrix(self):
        """Read tk vars and delegate to engine helper. See build_feeds_speeds_matrix."""
        return build_feeds_speeds_matrix(
            self._feeds_speeds_sweep_cfg('speed'),
            self._feeds_speeds_sweep_cfg('power'),
            self._feeds_speeds_sweep_cfg('passes'),
        )

    def _update_feeds_speeds_preview(self):
        """Refresh the 'N discs ≈ W × H mm' label below the inputs."""
        if not hasattr(self, 'fs_preview_var'):
            return
        try:
            n_speed = self._feeds_speeds_sweep_count('speed')
            n_power = self._feeds_speeds_sweep_count('power')
            n_passes = self._feeds_speeds_sweep_count('passes')
            air_factor = 2 if (hasattr(self, 'fs_also_no_air_var')
                                and self.fs_also_no_air_var.get()) else 1
            n_total = n_speed * n_power * n_passes * air_factor
            diameter = float(self.fs_diameter_var.get())
            if diameter <= 0:
                self.fs_preview_var.set("")
                return
            min_w, min_h = _min_feeds_speeds_sheet(
                diameter, n_speed, n_power, n_passes * air_factor)
            sheet_w, sheet_h = self._get_feeds_speeds_sheet_mm()
            if sheet_w is None:
                fit_note = ""
            elif min_w <= sheet_w + 1e-6 and min_h <= sheet_h + 1e-6:
                fit_note = _("  ✓ fits sheet")
            else:
                # Sheet size is a target, not a hard cap — generation grows it
                # to fit rather than erroring.
                fit_note = _("  → sheet will grow to fit")
            self.fs_preview_var.set(
                _("{n} discs — need at least {w:.0f} × {h:.0f} mm{note}").format(
                    n=n_total, w=min_w, h=min_h, note=fit_note))
        except (ValueError, TypeError):
            self.fs_preview_var.set("")

    def _persist_feeds_speeds_settings(self):
        """Sync the feeds-speeds tab state into self.settings['tooling_settings']."""
        tooling = self.settings.get("tooling_settings", {})
        fs = tooling.get("feeds_speeds_tester", {})
        fs["material"] = self.fs_material_var.get()
        try:
            fs["disc_diameter_mm"] = float(self.fs_diameter_var.get())
        except (ValueError, TypeError):
            pass
        try:
            fs["inner_diameter_mm"] = float(self.fs_inner_diameter_var.get())
        except (ValueError, TypeError):
            pass
        for var_name in ("speed", "power", "passes"):
            v = self.fs_sweep_vars[var_name]
            fs[f"{var_name}_sweep"] = bool(v['sweep'].get())
            for sub in ("value", "start", "end"):
                try:
                    fs[f"{var_name}_{sub}"] = float(v[sub].get())
                except (ValueError, TypeError):
                    pass
            # Stops is a step count — always an integer, never a decimal.
            try:
                fs[f"{var_name}_stops"] = int(float(v['stops'].get()))
            except (ValueError, TypeError):
                pass
        fs["engraving_on"] = bool(self.fs_engraving_var.get())
        try:
            fs["eng_speed_value"] = float(self.fs_eng_speed_var.get())
        except (ValueError, TypeError):
            pass
        try:
            fs["eng_power_value"] = float(self.fs_eng_power_var.get())
        except (ValueError, TypeError):
            pass
        fs["air_assist"] = bool(self.fs_air_assist_var.get())
        fs["also_test_no_air"] = bool(self.fs_also_no_air_var.get())
        fs["show_preview"] = bool(self.fs_show_preview_var.get())
        fs["sheet_w"] = self.fs_sheet_w_var.get()
        fs["sheet_h"] = self.fs_sheet_h_var.get()
        fs["sheet_unit"] = self.fs_sheet_unit_var.get()
        fs["filename"] = self.fs_filename_var.get()
        tooling["feeds_speeds_tester"] = fs
        self.settings["tooling_settings"] = tooling

    def _prepare_feeds_speeds_pieces(self):
        """
        Validate inputs and build (test_pieces, sheet_w_mm, sheet_h_mm,
        eng_speed, eng_power).

        Returns the tuple on success, or None and shows an error messagebox.
        """
        sheet_w, sheet_h = self._get_feeds_speeds_sheet_mm()
        if sheet_w is None:
            messagebox.showerror(_("Invalid Input"),
                                  _("Sheet width and height must be positive numbers."))
            return None
        try:
            diameter = float(self.fs_diameter_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"),
                                  _("Disc diameter must be a valid number."))
            return None
        if diameter <= 0:
            messagebox.showerror(_("Invalid Input"),
                                  _("Disc diameter must be greater than zero."))
            return None

        try:
            inner_diameter = float(self.fs_inner_diameter_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"),
                                  _("Hole diameter must be a valid number (use 0 for a solid disc)."))
            return None
        if inner_diameter < 0:
            messagebox.showerror(_("Invalid Input"),
                                  _("Hole diameter can't be negative (use 0 for a solid disc)."))
            return None
        if inner_diameter > 0 and inner_diameter >= diameter:
            messagebox.showerror(_("Invalid Input"),
                                  _("Hole diameter must be smaller than the disc diameter."))
            return None

        try:
            eng_speed = float(self.fs_eng_speed_var.get())
            eng_power = float(self.fs_eng_power_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"),
                                  _("Engraving speed and power must be valid numbers."))
            return None
        # 0 is allowed and means "not set" (cork pre-fills that way): the
        # IDs are then left off and the user is told. Refusing here would
        # only push them to type a guess. Negative, or over 100 %, is a typo.
        if eng_speed < 0 or eng_power < 0 or eng_power > 100:
            messagebox.showerror(_("Invalid Input"),
                                  _("Engraving speed must be > 0 and power must be 1-100%."))
            return None

        try:
            triples, cols, rows, num_blocks = self._build_feeds_speeds_matrix()
        except ValueError as e:
            messagebox.showerror(_("Invalid Input"), str(e))
            return None

        # When "Also test with air off" is checked, every triple is run
        # twice (once with air on, once with air off). The air-on set comes
        # first so its piece IDs are lower-numbered, which aligns with the
        # main "Air assist" checkbox being the primary state.
        air_default = bool(self.fs_air_assist_var.get())
        also_no_air = bool(self.fs_also_no_air_var.get())
        if also_no_air:
            air_states = [True, False]
        else:
            air_states = [air_default]
        # Pair each (triple, air_state) in placement order so positions
        # align with the grid packer's (block, row, col) iteration.
        expanded = [(triple, air) for air in air_states for triple in triples]
        total_blocks = num_blocks * len(air_states)

        try:
            positions, total_w, total_h = _grid_pack_discs(
                diameter, cols, rows, total_blocks, sheet_w, sheet_h)
        except ValueError:
            # The sheet size is a target, not a hard cap — who knows what
            # people will want to run. If the matrix doesn't fit, grow the
            # sheet (each axis only as far as needed) and carry on. The layout
            # preview shows the enlarged sheet before the file dialog, so the
            # user still sees the real size before committing.
            min_w, min_h = _min_feeds_speeds_sheet(
                diameter, cols, rows, total_blocks)
            sheet_w = max(sheet_w, min_w)
            sheet_h = max(sheet_h, min_h)
            positions, total_w, total_h = _grid_pack_discs(
                diameter, cols, rows, total_blocks, sheet_w, sheet_h)

        material_key = self._get_feeds_speeds_material_key()
        engraving_on = bool(self.fs_engraving_var.get())

        # Warn (don't block) when a hole leaves the ring too thin to hold a
        # legible ID — the label would shrink to the floor and may run into
        # the cut. The legend + grid position still identify the disc.
        if (engraving_on and inner_diameter > 0 and engraving_is_set(eng_speed, eng_power)
                and not feeds_speeds_ring_fits_label(diameter, inner_diameter)):
            if not messagebox.askyesno(
                    _("Thin Ring"),
                    _("With a {hole:.1f} mm hole in a {dia:.1f} mm disc, the ring is "
                      "too narrow to engrave a clear ID — the number may be hard to "
                      "read or run into the cut.\n\nGenerate anyway?").format(
                          hole=inner_diameter, dia=diameter)):
                return None

        test_pieces = []
        for idx, ((cx, cy), ((speed, power, passes), air_state)) in enumerate(
                zip(positions, expanded), start=1):
            test_pieces.append({
                'id': f"{idx:02d}",
                'cx': cx,
                'cy': cy,
                'diameter': diameter,
                'inner_diameter': inner_diameter,
                'speed': speed,
                'power': power,
                'passes': passes,
                'air_assist': air_state,
                'material': material_key,
                'engraving_on': engraving_on,
            })
        return test_pieces, sheet_w, sheet_h, eng_speed, eng_power

    def _write_feeds_speeds_legend(self, save_path, test_pieces, diameter,
                                    material_display, air_assist,
                                    eng_speed, eng_power, legend_path=None):
        """Write the legend.txt mapping IDs to (speed, power, passes, air).

        The file export derives the legend's name from the G-code's
        (<name>_legend.txt beside it). Frame & Cut has no G-code file, so
        it passes legend_path, the place the user chose, and the legend is
        written exactly there.
        """
        if legend_path is None:
            base, _ext = os.path.splitext(save_path)
            legend_path = base + "_legend.txt"
            header = _("Speed & Power Test — {fname}").format(fname=os.path.basename(save_path))
        else:
            # A Frame & Cut run has no G-code file to name: say what the run
            # was and when (it used to name the legend file itself).
            from datetime import datetime
            header = _("Speed & Power Test — sent to the laser with Frame & Cut, {when}").format(
                when=datetime.now().strftime("%Y-%m-%d %H:%M"))

        # Show "Air assist" as a per-piece value if the pieces vary; otherwise
        # collapse to the single global state.
        air_states = {p.get('air_assist', air_assist) for p in test_pieces}
        air_varies = len(air_states) > 1

        lines = []
        lines.append(header)
        lines.append(_("Material: {m}").format(m=material_display))
        lines.append(_("Disc diameter: {d:.1f} mm").format(d=diameter))
        inner = test_pieces[0].get('inner_diameter', 0) if test_pieces else 0
        if inner and inner > 0:
            lines.append(_("Hole diameter: {d:.1f} mm (washer / shim)").format(d=inner))
        lines.append(_("Engraving (labels): {s:.0f} mm/min @ {p:.0f}%").format(
            s=eng_speed, p=eng_power))
        if air_varies:
            lines.append(_("Air assist: varies per disc (see Air column)"))
        else:
            lines.append(_("Air assist: {state}").format(
                state=_("ON") if air_assist else _("OFF")))
        lines.append("")
        if air_varies:
            lines.append(f"{'ID':<5}{'Speed (mm/min)':>16}{'Power (%)':>12}"
                          f"{'Passes':>10}{'Air':>6}")
            for p in test_pieces:
                air_str = "ON" if p.get('air_assist', air_assist) else "OFF"
                lines.append(f"{p['id']:<5}{p['speed']:>16}{p['power']:>12}"
                              f"{p['passes']:>10}{air_str:>6}")
        else:
            lines.append(f"{'ID':<5}{'Speed (mm/min)':>16}{'Power (%)':>12}"
                          f"{'Passes':>10}")
            for p in test_pieces:
                lines.append(f"{p['id']:<5}{p['speed']:>16}{p['power']:>12}"
                              f"{p['passes']:>10}")
        with open(legend_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
        return legend_path

    def _write_feeds_speeds_gcode(self, prepared, path):
        """Write the Speed & Power test G-code for what
        _prepare_feeds_speeds_pieces() returned to `path`. The one call
        both the file export and Frame & Cut make, so what is streamed is
        what would have been saved."""
        test_pieces, sheet_w, sheet_h, eng_speed, eng_power = prepared
        generate_feeds_speeds_test_gcode(
            test_pieces, sheet_w, sheet_h, path, self.settings,
            air_assist=bool(self.fs_air_assist_var.get()),
            eng_speed_override=eng_speed,
            eng_power_override=eng_power)

    def _feeds_speeds_ids_skipped(self, prepared):
        """True when the disc IDs were asked for but cannot be engraved:
        the Engraving speed or power is 0 (not set)."""
        _pieces, _w, _h, eng_speed, eng_power = prepared
        return bool(self.fs_engraving_var.get()) and not engraving_is_set(eng_speed, eng_power)

    def _on_frame_cut_feeds_speeds(self):
        """Frame & Cut for the Speed & Power test: the Generate G-code job,
        sent to the laser. Frames at the selected material's framing power.

        The legend is the only record of which disc was cut at which
        settings, so it is asked for and written before anything moves;
        backing out of that prompt backs out of the run."""
        def prepare():
            prepared = self._prepare_feeds_speeds_pieces()
            if prepared is None:
                return None
            test_pieces, sheet_w, sheet_h, eng_speed, eng_power = prepared

            if self.fs_show_preview_var.get():
                preview = SpeedPowerTestPreview(
                    self.root, test_pieces, sheet_w, sheet_h,
                    eng_speed=eng_speed, eng_power=eng_power,
                    material_display=self.fs_material_var.get(),
                    proceed_label=_("Continue →"))
                if preview.result != "save":
                    return None

            if self._feeds_speeds_ids_skipped(prepared) and not self._confirm_engraving_skipped(
                    self._engraving_not_set_notice(None, 'ids')):
                return None

            base = self.fs_filename_var.get().strip() or "speed_power_test"
            mat_name = self.fs_material_var.get().lower()
            legend_path = filedialog.asksaveasfilename(
                title=_("Save Speed & Power Test Legend"),
                defaultextension=".txt",
                filetypes=[(_("Text files"), "*.txt")],
                initialfile=f"{base}_{mat_name}_legend.txt",
                initialdir=self.settings.get("last_output_dir", ""))
            if not legend_path:
                return None

            self.settings["last_output_dir"] = os.path.dirname(legend_path)
            self._persist_feeds_speeds_settings()
            text = self._gcode_text_from(
                lambda path: self._write_feeds_speeds_gcode(prepared, path))
            self._write_feeds_speeds_legend(
                legend_path, test_pieces,
                diameter=test_pieces[0]['diameter'],
                material_display=self.fs_material_var.get(),
                air_assist=bool(self.fs_air_assist_var.get()),
                eng_speed=eng_speed, eng_power=eng_power,
                legend_path=legend_path)
            return {'gcode_text': text,
                    'material': self._get_feeds_speeds_material_key(),
                    'label': 'speed & power test', 'count': len(test_pieces),
                    **self._tooling_history('feeds_speeds_tester', self._feeds_speeds_description(test_pieces),
                                            self._get_feeds_speeds_material_key(), len(test_pieces),
                                            sheet_mm=(sheet_w, sheet_h), base=base)}
        self.on_frame_and_cut(tooling_job=prepare)

    def _on_generate_feeds_speeds_gcode(self):
        """Generate the Speed & Power test G-code and the matching legend.txt."""
        try:
            prepared = self._prepare_feeds_speeds_pieces()
            if prepared is None:
                return
            test_pieces, sheet_w, sheet_h, eng_speed, eng_power = prepared

            # Show the preview first — gated by the "Preview before saving"
            # checkbox so a user who's already dialed-in their inputs can
            # skip straight to the file dialog.
            if self.fs_show_preview_var.get():
                preview = SpeedPowerTestPreview(
                    self.root, test_pieces, sheet_w, sheet_h,
                    eng_speed=eng_speed, eng_power=eng_power,
                    material_display=self.fs_material_var.get())
                if preview.result != "save":
                    return

            base = self.fs_filename_var.get().strip() or "speed_power_test"
            # Named from what the dropdown shows, so the file name a
            # translated app suggests stays what it was.
            mat_name = self.fs_material_var.get().lower()
            save_path = filedialog.asksaveasfilename(
                title=_("Save Speed & Power Test G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}_{mat_name}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            self._persist_feeds_speeds_settings()
            self._write_feeds_speeds_gcode(prepared, save_path)
            legend_path = self._write_feeds_speeds_legend(
                save_path, test_pieces,
                diameter=test_pieces[0]['diameter'],
                material_display=self.fs_material_var.get(),
                air_assist=bool(self.fs_air_assist_var.get()),
                eng_speed=eng_speed, eng_power=eng_power)
            self._record_tooling_export("gcode", save_path, tool='feeds_speeds_tester',
                                        description=self._feeds_speeds_description(test_pieces),
                                        material=self._get_feeds_speeds_material_key(),
                                        count=len(test_pieces), sheet_mm=(sheet_w, sheet_h))
            parts = [_("Wrote {n} test discs to:\n{p}\n\nLegend saved to:\n{l}").format(
                n=len(test_pieces), p=save_path, l=legend_path)]
            if self._feeds_speeds_ids_skipped(prepared):
                parts.append(self._engraving_not_set_notice(None, 'ids'))
            parts.append(gcode_file_note())
            messagebox.showinfo(_("G-code Generated"), "\n\n".join(parts))

        except Exception as e:
            messagebox.showerror(_("Error"),
                                  _("Something went wrong:\n\n{e}").format(e=e))

    # ========================================
    # DIE HOLDER GENERATION
    # ========================================

    def _on_generate_holder_svg(self):
        """Generate SVG for die holder pieces."""
        try:
            variant = self.holder_variant_var.get()
            layer_count = int(self.holder_layer_count_var.get())
            base = self.holder_filename_var.get().strip() or "die_holder"

            sheet_w, sheet_h = self._get_holder_sheet_mm()
            if sheet_w is None:
                return
            try:
                check_holder_sheet(variant, layer_count, sheet_w, sheet_h)
            except ValueError as ve:
                messagebox.showerror(_("Sheet too small"), str(ve))
                return

            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Holder SVG"),
                defaultextension=".svg",
                filetypes=[(_("SVG files"), "*.svg")],
                initialfile=f"{base}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            settings = self._get_die_settings()
            try:
                generate_holder_svg(variant, save_path, settings,
                                    layer_count=layer_count,
                                    sheet_width_mm=sheet_w, sheet_height_mm=sheet_h)
            except ValueError as ve:
                messagebox.showerror(_("Sheet too small"), str(ve))
                return
            self._record_tooling_export("svg", save_path, tool='die_holders',
                                        description=self._holder_description(variant, layer_count),
                                        material='acrylic', count=len(_holder_pieces_for(variant, layer_count)),
                                        sheet_mm=(sheet_w, sheet_h))

            variant_label = {"large": _("Large (40\u201360mm)"),
                             "small": _("Small (7\u201339.5mm)"),
                             "both": _("Both (Small + Large)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {layer_count}-layer {variant_label} die holder.\n\nSaved to: {path}").format(layer_count=layer_count, variant_label=variant_label, path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _validate_holder_job(self):
        """The Die Holders form, checked: (variant, layer_count, sheet_w,
        sheet_h) in mm, or None once the error has been shown. Generate
        G-code and Frame & Cut both start here."""
        variant = self.holder_variant_var.get()
        layer_count = int(self.holder_layer_count_var.get())

        sheet_w, sheet_h = self._get_holder_sheet_mm()
        if sheet_w is None:
            return None
        try:
            check_holder_sheet(variant, layer_count, sheet_w, sheet_h)
        except ValueError as ve:
            messagebox.showerror(_("Sheet too small"), str(ve))
            return None
        return variant, layer_count, sheet_w, sheet_h

    def _on_frame_cut_holder(self):
        """Frame & Cut for die holders: the Generate G-code job, sent to
        the laser. Acrylic, like the file it would have written."""
        def prepare():
            job = self._validate_holder_job()
            if job is None:
                return None
            variant, layer_count, sheet_w, sheet_h = job
            settings = self._get_die_settings()
            try:
                text = self._gcode_text_from(
                    lambda path: generate_holder_gcode(
                        variant, path, settings, layer_count=layer_count,
                        sheet_width_mm=sheet_w, sheet_height_mm=sheet_h))
            except ValueError as ve:
                messagebox.showerror(_("Sheet too small"), str(ve))
                return None
            return {'gcode_text': text, 'material': 'acrylic',
                    'label': 'die holder',
                    'count': len(_holder_pieces_for(variant, layer_count)),
                    **self._tooling_history('die_holders', self._holder_description(variant, layer_count),
                                            'acrylic', len(_holder_pieces_for(variant, layer_count)),
                                            sheet_mm=(sheet_w, sheet_h),
                                            base=self.holder_filename_var.get().strip() or None)}
        self.on_frame_and_cut(tooling_job=prepare)

    def _on_generate_holder_gcode(self):
        """Generate G-code for die holder pieces."""
        try:
            job = self._validate_holder_job()
            if job is None:
                return
            variant, layer_count, sheet_w, sheet_h = job
            base = self.holder_filename_var.get().strip() or "die_holder"

            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Holder G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            settings = self._get_die_settings()
            try:
                generate_holder_gcode(variant, save_path, settings,
                                      layer_count=layer_count,
                                      sheet_width_mm=sheet_w, sheet_height_mm=sheet_h)
            except ValueError as ve:
                messagebox.showerror(_("Sheet too small"), str(ve))
                return
            self._record_tooling_export("gcode", save_path, tool='die_holders',
                                        description=self._holder_description(variant, layer_count),
                                        material='acrylic', count=len(_holder_pieces_for(variant, layer_count)),
                                        sheet_mm=(sheet_w, sheet_h))

            variant_label = {"large": _("Large (40\u201360mm)"),
                             "small": _("Small (7\u201339.5mm)"),
                             "both": _("Both (Small + Large)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {layer_count}-layer {variant_label} die holder G-code.\n\nSaved to: {path}").format(layer_count=layer_count, variant_label=variant_label, path=save_path)
                + "\n\n" + gcode_file_note())

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_generate_organizer_svg(self):
        """Generate SVG for the die organizer (upper or lower plate)."""
        try:
            variant = self.organizer_variant_var.get()
            base = self.organizer_filename_var.get().strip() or "die_organizer"

            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Organizer SVG"),
                defaultextension=".svg",
                filetypes=[(_("SVG files"), "*.svg")],
                initialfile=f"{base}_{variant}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            generate_die_organizer_svg(variant, save_path, self.settings)
            # The drawing as Matt drew it: no material, no sheet (LightBurn's).
            self._record_tooling_export("svg", save_path, tool='die_organizer',
                                        description=self._organizer_description(variant), material=None, count=1)

            variant_label = {"upper": _("Upper (slotted)"), "lower": _("Lower (base)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {variant_label} die organizer.\n\nSaved to: {path}").format(variant_label=variant_label, path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _organizer_material_key(self):
        """Settings key of the material picked for the organizer's G-code."""
        key = self._tooling_material_key(self.organizer_material_var.get())
        return key if key in _ORGANIZER_MATERIALS else _ORGANIZER_MATERIALS[0]

    def _organizer_gcode_job(self):
        """The Die Organizer form's G-code inputs, checked: (variant,
        material, sheet_w, sheet_h) in mm, or None once the error has been
        shown. Generate G-code and Frame & Cut both start here."""
        variant = self.organizer_variant_var.get()
        material = self._organizer_material_key()
        try:
            sheet_w = float(self.organizer_width_var.get())
            sheet_h = float(self.organizer_height_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"), _("Sheet width and height must be valid numbers."))
            return None
        if sheet_w <= 0 or sheet_h <= 0:
            messagebox.showerror(_("Invalid Input"), _("Sheet dimensions must be positive."))
            return None
        if self.organizer_units_var.get() == "in":
            sheet_w *= 25.4
            sheet_h *= 25.4
        need_w, need_h = die_organizer_min_sheet(variant)
        if sheet_w < need_w - 1e-6 or sheet_h < need_h - 1e-6:
            messagebox.showerror(
                _("Sheet too small"),
                _("The organizer plate needs a sheet at least {w} × {h} mm ({w_in} × {h_in} in), "
                  "upright as drawn.").format(
                      w=f"{need_w:.0f}", h=f"{need_h:.0f}",
                      w_in=f"{need_w / 25.4:.1f}", h_in=f"{need_h / 25.4:.1f}"))
            return None
        return variant, material, sheet_w, sheet_h

    def _write_organizer_gcode(self, job, path):
        """Write the organizer G-code for an _organizer_gcode_job() to
        `path`. Returns the generator's info, or None once a refusal (no
        cut settings, a slot narrower than the kerf, a drawing it cannot
        read) has been shown."""
        variant, material, sheet_w, sheet_h = job
        try:
            return generate_die_organizer_gcode(variant, material, path, self.settings,
                                                sheet_width_mm=sheet_w, sheet_height_mm=sheet_h)
        except OrganizerKerfError as e:
            messagebox.showerror(_("Invalid Input"), self._organizer_kerf_message(material, e))
        except OrganizerDrawingError as e:
            messagebox.showerror(_("Error"), str(e))
        except NoCutSettingsError:
            # Only this case gets this sentence. Any other ValueError (the
            # engine's own sheet check, a settings value that is not a
            # number) goes up to the handler's "Something went wrong" box as
            # what it is; until 2026-11-09 it was reported as missing cut
            # settings.
            messagebox.showerror(_("Invalid Input"), self._no_cut_settings_message(material))
        return None

    def _organizer_kerf_message(self, material, e):
        """The organizer's kerf refusal in the user's language (the engine's
        words are English), the way the crescent's kerf refusals have one."""
        name = self._tooling_material_label(material)
        where = self._material_settings_place(material)
        if e.opening is None:
            return _("The {material} kerf is negative ({kerf} mm). Check it in {where}.").format(
                material=name, kerf=f"{e.kerf_mm:g}", where=where)
        return _("The {material} kerf ({kerf} mm) is too wide for an opening of {opening} mm in the "
                 "organizer plate. Check it in {where}.").format(
                     material=name, kerf=f"{e.kerf_mm:g}", opening=e.opening, where=where)

    def _no_cut_settings_message(self, material_key):
        return _("{material} has no cut speed and power set ({where}).").format(
            material=self._tooling_material_label(material_key),
            where=self._material_settings_place(material_key))

    def _on_generate_organizer_gcode(self):
        """Generate G-code for the die organizer (upper or lower plate)."""
        try:
            job = self._organizer_gcode_job()
            if job is None:
                return
            variant = job[0]
            base = self.organizer_filename_var.get().strip() or "die_organizer"
            save_path = filedialog.asksaveasfilename(
                title=_("Save Die Organizer G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}_{variant}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            info = self._write_organizer_gcode(job, save_path)
            if info is None:
                try:
                    os.remove(save_path)        # nothing usable was written
                except OSError:
                    pass
                return
            self._record_tooling_export("gcode", save_path, tool='die_organizer',
                                        description=self._organizer_description(variant), material=job[1],
                                        count=sum(info['counts'].values()), sheet_mm=(job[2], job[3]))
            variant_label = {"upper": _("Upper (slotted)"), "lower": _("Lower (base)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {variant_label} die organizer.\n\nSaved to: {path}").format(variant_label=variant_label, path=save_path)
                + "\n\n" + gcode_file_note())

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_frame_cut_organizer(self):
        """Frame & Cut for the die organizer: the Generate G-code job, sent
        to the laser. Frames at the chosen material's framing power."""
        def prepare():
            job = self._organizer_gcode_job()
            if job is None:
                return None
            got = {}

            def write(path):
                got['info'] = self._write_organizer_gcode(job, path)
            text = self._gcode_text_from(write)
            if got['info'] is None:
                return None
            return {'gcode_text': text, 'material': job[1],
                    'label': f'die organizer {job[0]}',
                    'count': sum(got['info']['counts'].values()),
                    **self._tooling_history('die_organizer', self._organizer_description(job[0]), job[1],
                                            sum(got['info']['counts'].values()), sheet_mm=(job[2], job[3]),
                                            base=self.organizer_filename_var.get().strip() or None)}
        self.on_frame_and_cut(tooling_job=prepare)

    def _save_pad_spacer_stl(self, stl_filename):
        """Copy a bundled pad-press-spacer STL to a user-chosen location."""
        import shutil
        import sys
        try:
            if getattr(sys, 'frozen', False):
                base = sys._MEIPASS
            else:
                base = os.path.dirname(os.path.abspath(__file__))
            src = os.path.join(base, 'pad_press_spacers', stl_filename)

            if not os.path.isfile(src):
                messagebox.showerror(_("Error"),
                    _("Could not find bundled STL: {path}").format(path=src))
                return

            save_path = filedialog.asksaveasfilename(
                title=_("Save Pad Press Spacer STL"),
                defaultextension=".stl",
                filetypes=[(_("STL files"), "*.stl")],
                initialfile=stl_filename,
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            shutil.copyfile(src, save_path)
            self.settings["last_output_dir"] = os.path.dirname(save_path)

            messagebox.showinfo(_("Done"),
                _("Saved {filename} to:\n\n{path}").format(filename=stl_filename, path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    # ========================================
    # TONEHOLE CRESCENT
    # ========================================

    def _read_crescent_inputs(self):
        """Parse the crescent fields into (tonehole_d, percent, curve).

        Raises ValueError carrying a message ready to show the user.
        """
        try:
            d = float(self.crescent_tonehole_var.get())
            pct = float(self.crescent_percent_var.get())
        except (ValueError, TypeError):
            raise ValueError(_("Enter the tonehole ID and the percentage as numbers.")) from None
        if not 0 < d <= CRESCENT_MAX_TONEHOLE_MM:
            raise ValueError(_("Tonehole ID must be more than 0 and at most {max} mm.").format(
                max=f"{CRESCENT_MAX_TONEHOLE_MM:g}"))
        if not 0 < pct <= CRESCENT_MAX_PERCENT:
            raise ValueError(_("The percentage must be more than 0 and at most {max}.").format(
                max=f"{CRESCENT_MAX_PERCENT:g}"))
        return d, pct, float(self.crescent_curve_var.get())

    def _update_crescent_preview(self):
        """Redraw the crescent preview and the measurements under it."""
        cv = self.crescent_canvas
        cv.delete("all")
        r_px = 70
        cx, cy = int(cv.cget('width')) / 2, int(cv.cget('height')) / 2
        cv.create_oval(cx - r_px, cy - r_px, cx + r_px, cy + r_px, outline="#888888", dash=(3, 3))
        try:
            geom = tonehole_crescent_geometry(*self._read_crescent_inputs())
        except ValueError as e:
            self.crescent_readout_var.set(str(e))
            return
        scale = r_px / geom['radius']
        outer, inner = tonehole_crescent_outline(geom, 64)
        pts = []
        for x, y in outer + inner[-2:0:-1]:
            pts += [cx + x * scale, cy + y * scale]
        cv.create_polygon(pts, fill="#E8A080", outline="#993C1D")
        self.crescent_readout_var.set(
            _("Widest point {w} mm, inner edge radius {r} mm").format(
                w=f"{geom['width']:.2f}", r=f"{geom['inner_radius']:.1f}")
            + "\n"
            + _("Open area left {pct}%, same as a {d} mm round hole").format(
                pct=f"{geom['open_percent']:g}", d=f"{geom['equivalent_d']:.1f}"))

    def _read_crescent_qty(self):
        try:
            qty = int(self.crescent_qty_var.get())
        except (ValueError, TypeError):
            qty = 0
        if not 1 <= qty <= CRESCENT_MAX_QTY:
            raise ValueError(_("Quantity must be a whole number from 1 to {max}.").format(max=CRESCENT_MAX_QTY))
        return qty

    @staticmethod
    def _clean_crescent_entry(raw):
        """A saved list entry, normalised, or None if it no longer makes a crescent."""
        try:
            entry = {'tonehole_d': float(raw['tonehole_d']), 'percent': float(raw['percent']),
                     'curve': float(raw.get('curve', CRESCENT_CURVE_STANDARD)), 'qty': int(raw.get('qty', 1))}
            tonehole_crescent_geometry(entry['tonehole_d'], entry['percent'], entry['curve'])
        except (KeyError, TypeError, ValueError, AttributeError):
            return None
        return entry if 1 <= entry['qty'] <= CRESCENT_MAX_QTY else None

    def _on_add_crescent(self):
        try:
            d, pct, curve = self._read_crescent_inputs()
            qty = self._read_crescent_qty()
        except ValueError as e:
            messagebox.showerror(_("Invalid Input"), str(e))
            return
        self.crescent_list.append({'tonehole_d': d, 'percent': pct, 'curve': curve, 'qty': qty})
        self._refresh_crescent_list()

    def _on_remove_crescent(self):
        sel = self.crescent_listbox.curselection()
        if sel:
            del self.crescent_list[sel[0]]
            self._refresh_crescent_list()

    def _refresh_crescent_list(self):
        """Redraw the list and the line saying what Generate will make."""
        lb = self.crescent_listbox
        lb.delete(0, tk.END)
        for n, e in enumerate(self.crescent_list, 1):
            g = tonehole_crescent_geometry(e['tonehole_d'], e['percent'], e['curve'])
            lb.insert(tk.END, _("{n}.  {d} mm, {pct}%, inner edge radius {r} mm, × {qty}").format(
                n=n, d=f"{e['tonehole_d']:g}", pct=f"{e['percent']:g}",
                r=f"{g['inner_radius']:.1f}", qty=e['qty']))
        if not self.crescent_list:
            hint = _("The list is empty, so Generate makes the crescent above.")
        else:
            total = sum(e['qty'] for e in self.crescent_list)
            hint = ngettext("Generate makes the list: {n} crescent.",
                            "Generate makes the list: {n} crescents.", total).format(n=total)
            if len(self.crescent_list) > 1:
                hint += " " + _("The SVG engraves each one's list number on the scrap beside it.")
        self.crescent_list_hint_var.set(hint)

    def _on_generate_crescent_svg(self):
        self._generate_crescent("svg")

    def _on_generate_crescent_stl(self):
        self._generate_crescent("stl")

    def _crescent_entries(self):
        """What Generate makes: the list, or with an empty list the one
        crescent in the fields. Raises ValueError for a bad field."""
        if self.crescent_list:
            return [dict(e) for e in self.crescent_list]
        d, pct, curve = self._read_crescent_inputs()
        return [{'tonehole_d': d, 'percent': pct, 'curve': curve,
                 'qty': self._read_crescent_qty()}]

    def _crescent_material_key(self):
        """Settings key of the material picked for the crescents' G-code."""
        key = self._tooling_material_key(self.crescent_material_var.get())
        return key if key in _CRESCENT_MATERIALS else _CRESCENT_MATERIALS[0]

    def _crescent_initial_name(self, entries, ext):
        if len(entries) == 1:
            return f"crescent_{entries[0]['tonehole_d']:g}mm_{entries[0]['percent']:g}pct{ext}"
        return f"tonehole_crescents{ext}"

    def _write_crescent_gcode(self, entries, material, path):
        """Write the crescents' G-code to `path`. Returns the generator's
        info, or None once a refusal has been shown: no cut settings for
        the material, or a crescent whose horns are closer together than
        the material's kerf (the beam would bridge the gap)."""
        try:
            return generate_tonehole_crescent_gcode(entries, material, path, self.settings)
        except CrescentKerfError as e:
            name = self._tooling_material_label(material)
            if e.reason == 'gap':
                message = _("The opening between this crescent's horns ({gap} mm) is narrower than "
                            "the {material} kerf ({kerf} mm), so the laser cannot cut it. Use a "
                            "smaller percentage or a flatter inner edge.").format(
                                gap=f"{e.size_mm:.2f}", material=name, kerf=f"{e.kerf_mm:g}")
            else:
                message = _("The {material} kerf ({kerf} mm) is too large for this crescent. "
                            "Check it in {where}.").format(
                                material=name, kerf=f"{e.kerf_mm:g}",
                                where=self._material_settings_place(material))
            messagebox.showerror(_("Invalid Input"), message)
        except NoCutSettingsError:
            # Only the no-cut-settings case; see _write_organizer_gcode.
            messagebox.showerror(_("Invalid Input"), self._no_cut_settings_message(material))
        return None

    def _on_generate_crescent_gcode(self):
        """Generate G-code for the list (or the crescent in the fields),
        cut from the chosen material with its kerf compensated."""
        try:
            entries = self._crescent_entries()
        except ValueError as e:
            messagebox.showerror(_("Invalid Input"), str(e))
            return
        try:
            material = self._crescent_material_key()
            save_path = filedialog.asksaveasfilename(
                title=_("Save Tonehole Crescent"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=self._crescent_initial_name(entries, ".gcode"),
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            info = self._write_crescent_gcode(entries, material, save_path)
            if info is None:
                try:
                    os.remove(save_path)        # nothing usable was written
                except OSError:
                    pass
                return
            self._record_tooling_export("gcode", save_path, tool='tonehole_crescent',
                                        description=self._crescent_description(entries), material=material,
                                        count=len(info['parts']))
            parts = [_("Saved to: {d}").format(d=save_path)]
            if info['numbers_skipped']:
                parts.append(self._engraving_not_set_notice(material, 'numbers'))
            parts.append(gcode_file_note())
            messagebox.showinfo(_("Done"), "\n\n".join(parts))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_frame_cut_crescent(self):
        """Frame & Cut for tonehole crescents: the Generate G-code job,
        sent to the laser. Frames at the chosen material's framing power."""
        def prepare():
            try:
                entries = self._crescent_entries()
            except ValueError as e:
                messagebox.showerror(_("Invalid Input"), str(e))
                return None
            material = self._crescent_material_key()
            got = {}

            def write(path):
                got['info'] = self._write_crescent_gcode(entries, material, path)
            text = self._gcode_text_from(write)
            info = got['info']
            if info is None:
                return None
            if info['numbers_skipped'] and not self._confirm_engraving_skipped(
                    self._engraving_not_set_notice(material, 'numbers')):
                return None
            return {'gcode_text': text, 'material': material,
                    'label': 'tonehole crescent', 'count': len(info['parts']),
                    **self._tooling_history('tonehole_crescent', self._crescent_description(entries),
                                            material, len(info['parts']))}
        self.on_frame_and_cut(tooling_job=prepare)

    def _generate_crescent(self, kind):
        """Save the list (or, with an empty list, the crescent in the fields)
        as SVG cut paths ("svg") or a printable STL ("stl")."""
        try:
            entries = self._crescent_entries()
            thickness = None
            if kind == "stl":
                try:
                    thickness = float(self.crescent_thickness_var.get())
                except (ValueError, TypeError):
                    thickness = 0.0
                if not 0 < thickness <= CRESCENT_MAX_THICKNESS_MM:
                    raise ValueError(_("Print thickness must be more than 0 and at most {max} mm.").format(
                        max=f"{CRESCENT_MAX_THICKNESS_MM:g}"))
        except ValueError as e:
            messagebox.showerror(_("Invalid Input"), str(e))
            return

        try:
            ext = "." + kind
            initial = self._crescent_initial_name(entries, ext)
            save_path = filedialog.asksaveasfilename(
                title=_("Save Tonehole Crescent"),
                defaultextension=ext,
                filetypes=[(_("SVG files") if kind == "svg" else _("STL files"), "*" + ext)],
                initialfile=initial,
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            if kind == "svg":
                generate_tonehole_crescent_svg(entries, save_path, self.settings)
            else:
                generate_tonehole_crescent_stl(entries, thickness, save_path)
            # Neither knows a material: the SVG is cut with LightBurn's
            # settings, the STL is printed.
            self._record_tooling_export(kind, save_path, tool='tonehole_crescent',
                                        description=self._crescent_description(entries), material=None,
                                        count=sum(e['qty'] for e in entries))
            messagebox.showinfo(_("Done"), _("Saved to: {d}").format(d=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))
