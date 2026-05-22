"""
Tooling tab mixin for Stohrer Sax Shop Companion.

Provides die insert and die holder generation for laser-cutting acrylic tooling.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import webbrowser

from svg_engine import (
    _nest_discs, try_nest_partial, can_all_pads_fit,
    generate_die_svg, generate_die_svg_from_placed,
    generate_holder_svg, generate_kerf_test_svg,
    generate_die_organizer_svg,
    _min_feeds_speeds_sheet, _grid_pack_discs,
    build_feeds_speeds_matrix,
)
from gcode_engine import (
    generate_die_gcode_from_placed, generate_holder_gcode, generate_kerf_test_gcode,
    generate_feeds_speeds_test_gcode,
)
from ui_dialogs import GcodeSettingsWindow, SpeedPowerTestPreview

IS_MACOS = sys.platform == 'darwin'


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

        # Tool selector buttons at top, laid out in two rows so longer
        # labels like "Feeds & Speeds Tester" stay readable.
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
        ]
        row_frames = [tk.Frame(selector_frame, bg=bg), tk.Frame(selector_frame, bg=bg)]
        row_frames[0].pack(fill='x')
        row_frames[1].pack(fill='x', pady=(4, 0))
        per_row = (len(tool_names) + 1) // 2  # split evenly, first row gets extras
        for idx, (key, label) in enumerate(tool_names):
            row = row_frames[0 if idx < per_row else 1]
            btn = tk.Button(row, text=label, width=22,
                            command=lambda k=key: self._show_tooling_section(k))
            btn.pack(side='left', padx=(0, 5))
            self._tooling_buttons[key] = btn

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

        self._tooling_sections['die_holders'] = holder_frame
        self._update_holder_info()

        # ========================================
        # DIE ORGANIZER SECTION
        # ========================================
        organizer_frame = tk.Frame(self._tooling_content, bg=bg)

        # Description + Matt's instructions
        organizer_info = tk.Label(organizer_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                                  text=_("Plate: 230 × 330 mm. Open the SVG in LightBurn or your "
                                         "laser software to cut."),
                                  justify="left")
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
                  command=self._on_generate_organizer_svg).pack(side='left')

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
        kerf_materials = [_("Felt"), _("Card"), _("Leather"), _("Acrylic")]
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
            values=[_("Felt"), _("Card"), _("Leather"), _("Acrylic")],
            state="readonly", width=12)
        fs_mat_combo.pack(side='left', padx=(5, 10))
        tk.Button(fs_mat_row, text=_("Apply material defaults"),
                  command=self._apply_feeds_speeds_material_defaults).pack(side='left')

        # Disc diameter
        fs_dia_row = tk.Frame(fs_frame, bg=bg)
        fs_dia_row.pack(fill='x', pady=(0, 5))
        tk.Label(fs_dia_row, text=_("Disc diameter:"), bg=bg).pack(side='left')
        self.fs_diameter_var = tk.StringVar(
            value=str(fs_settings.get("disc_diameter_mm", 20.0)))
        tk.Entry(fs_dia_row, textvariable=self.fs_diameter_var, width=8
                 ).pack(side='left', padx=(2, 2))
        tk.Label(fs_dia_row, text=_("mm"), bg=bg).pack(side='left')
        self.fs_diameter_var.trace_add(
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
        tk.Checkbutton(fs_gen_frame, text=_("Preview before saving"),
                       variable=self.fs_show_preview_var, bg=bg
                       ).pack(side='left')

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
        """Open settings dialog showing acrylic G-code settings + die engraving options."""
        from config import save_settings
        acrylic_materials = [("acrylic", _("Acrylic"))]
        GcodeSettingsWindow(self.root, self.settings, lambda s: save_settings(s),
                            materials=acrylic_materials, show_tooling_engraving=True)

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

            messagebox.showinfo(_("Done"),
                _("Generated {n} die rings.\n\nSaved to: {path}").format(n=len(placed), path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_generate_die_gcode(self):
        """Generate G-code for die inserts."""
        if self.die_scrap_var.get():
            self._generate_die_gcode_scrap()
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
                title=_("Save Die Insert G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)

            placed, _, _ = _nest_discs(pads, 'die_ring', width_mm, height_mm, settings)
            generate_die_gcode_from_placed(placed, width_mm, height_mm, save_path, settings)

            messagebox.showinfo(_("Done"),
                _("Generated G-code for {n} die rings.\n\nSaved to: {path}").format(n=len(placed), path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

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

    def _generate_die_svg_scrap(self):
        """Handle SVG generation in die scrap mode."""
        try:
            width_mm, height_mm = self._get_die_sheet_mm()
            if width_mm is None:
                return

            settings = self._get_die_settings()

            if not self.tooling_scrap_session['active']:
                # Starting new session
                sizes = self._parse_die_sizes()
                if not sizes:
                    messagebox.showerror(_("Error"), _("No valid die sizes entered."))
                    return
                pads = self._sizes_to_pads(sizes)

                save_dir = filedialog.askdirectory(
                    title=_("Select Folder to Save SVGs"),
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir
                self._start_die_scrap_session(pads, save_dir)
            else:
                pads = self.tooling_scrap_session['remaining_pads']

            if not pads:
                messagebox.showinfo(_("Session Complete"), _("All dies have been placed!"))
                return

            placed, remaining, any_placed = try_nest_partial(
                pads, 'die_ring', width_mm, height_mm, settings)

            if not any_placed:
                min_size = min(p['size'] for p in pads)
                messagebox.showwarning(_("No Dies Fit"),
                    _("No dies could be placed on this sheet.\n\n"
                      "Smallest remaining die: {min_size}mm\n"
                      "Try a larger sheet.").format(min_size=min_size))
                return

            self.tooling_scrap_session['scrap_count'] += 1
            scrap_num = self.tooling_scrap_session['scrap_count']
            save_dir = self.tooling_scrap_session['save_dir']
            base = self.die_filename_var.get().strip() or "my_dies"
            filename = os.path.join(save_dir, f"{base}_scrap{scrap_num}.svg")

            generate_die_svg_from_placed(placed, width_mm, height_mm, filename, settings)

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
            width_mm, height_mm = self._get_die_sheet_mm()
            if width_mm is None:
                return

            settings = self._get_die_settings()

            if not self.tooling_scrap_session['active']:
                sizes = self._parse_die_sizes()
                if not sizes:
                    messagebox.showerror(_("Error"), _("No valid die sizes entered."))
                    return
                pads = self._sizes_to_pads(sizes)

                save_dir = filedialog.askdirectory(
                    title=_("Select Folder to Save G-code"),
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir
                self._start_die_scrap_session(pads, save_dir)
            else:
                pads = self.tooling_scrap_session['remaining_pads']

            if not pads:
                messagebox.showinfo(_("Session Complete"), _("All dies have been placed!"))
                return

            placed, remaining, any_placed = try_nest_partial(
                pads, 'die_ring', width_mm, height_mm, settings)

            if not any_placed:
                min_size = min(p['size'] for p in pads)
                messagebox.showwarning(_("No Dies Fit"),
                    _("No dies could be placed on this sheet.\n\n"
                      "Smallest remaining die: {min_size}mm\n"
                      "Try a larger sheet.").format(min_size=min_size))
                return

            self.tooling_scrap_session['scrap_count'] += 1
            scrap_num = self.tooling_scrap_session['scrap_count']
            save_dir = self.tooling_scrap_session['save_dir']
            base = self.die_filename_var.get().strip() or "my_dies"
            filename = os.path.join(save_dir, f"{base}_scrap{scrap_num}.gcode")

            generate_die_gcode_from_placed(placed, width_mm, height_mm, filename, settings)

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
                      "Files saved to: {save_dir}").format(placed_count=placed_count, scrap_num=scrap_num, save_dir=save_dir))
            else:
                messagebox.showinfo(_("Sheet Generated"),
                    _("Generated G-code for {placed_count} dies on sheet #{scrap_num}.\n\n"
                      "{remaining_count} dies remaining.\n"
                      "Adjust sheet size if needed and generate again.").format(placed_count=placed_count, scrap_num=scrap_num, remaining_count=remaining_count))

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

    def _get_kerf_material_key(self):
        """Map display name to settings key."""
        return self.kerf_material_var.get().lower()

    def _show_kerf_instructions(self, material_name):
        """Show kerf test measurement instructions with a don't-show-again option."""
        if self.settings.get("seen_kerf_test_tutorial", False):
            return

        from config import save_settings

        mat_key = material_name.lower()
        # Map to settings location
        if mat_key == "acrylic":
            settings_path = _("Tooling > Options > Settings > Acrylic > Kerf Width")
        else:
            settings_path = _("Pad Generator > Options > G-code Settings > {material_name} > Kerf Width").format(material_name=material_name)

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
            self._show_kerf_instructions(material_name)

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_generate_kerf_gcode(self):
        """Generate G-code kerf test pattern."""
        try:
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
                    return
                eng_mode = "line"
                eng_speed = 1000
                eng_power = 15
                filled_spacing = 0.15

            save_path = filedialog.asksaveasfilename(
                title=_("Save Kerf Test G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"kerf_test_{mat_key}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            generate_kerf_test_gcode(material_name, save_path, self.settings,
                                     cut_speed, cut_power, eng_speed, eng_power,
                                     engraving_mode=eng_mode,
                                     filled_line_spacing=filled_spacing)
            self._show_kerf_instructions(material_name)

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    # ========================================
    # FEEDS & SPEEDS TESTER
    # ========================================

    def _get_feeds_speeds_material_key(self):
        """Map display name to settings key (mirrors kerf-test pattern)."""
        return self.fs_material_var.get().lower()

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
                fit_note = _("  ✗ does NOT fit current sheet")
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
            eng_speed = float(self.fs_eng_speed_var.get())
            eng_power = float(self.fs_eng_power_var.get())
        except (ValueError, TypeError):
            messagebox.showerror(_("Invalid Input"),
                                  _("Engraving speed and power must be valid numbers."))
            return None
        if eng_speed <= 0 or eng_power <= 0 or eng_power > 100:
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
        except ValueError as e:
            messagebox.showerror(_("Sheet Too Small"), str(e))
            return None

        material_key = self._get_feeds_speeds_material_key()
        engraving_on = bool(self.fs_engraving_var.get())

        test_pieces = []
        for idx, ((cx, cy), ((speed, power, passes), air_state)) in enumerate(
                zip(positions, expanded), start=1):
            test_pieces.append({
                'id': f"{idx:02d}",
                'cx': cx,
                'cy': cy,
                'diameter': diameter,
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
                                    eng_speed, eng_power):
        """Write the legend.txt mapping IDs to (speed, power, passes, air)."""
        base, _ext = os.path.splitext(save_path)
        legend_path = base + "_legend.txt"

        # Show "Air assist" as a per-piece value if the pieces vary; otherwise
        # collapse to the single global state.
        air_states = {p.get('air_assist', air_assist) for p in test_pieces}
        air_varies = len(air_states) > 1

        lines = []
        lines.append(_("Speed & Power Test — {fname}").format(fname=os.path.basename(save_path)))
        lines.append(_("Material: {m}").format(m=material_display))
        lines.append(_("Disc diameter: {d:.1f} mm").format(d=diameter))
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
            mat_key = self._get_feeds_speeds_material_key()
            save_path = filedialog.asksaveasfilename(
                title=_("Save Speed & Power Test G-code"),
                defaultextension=".gcode",
                filetypes=[(_("G-code files"), "*.gcode")],
                initialfile=f"{base}_{mat_key}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            self._persist_feeds_speeds_settings()
            generate_feeds_speeds_test_gcode(
                test_pieces, sheet_w, sheet_h, save_path, self.settings,
                air_assist=bool(self.fs_air_assist_var.get()),
                eng_speed_override=eng_speed,
                eng_power_override=eng_power)
            legend_path = self._write_feeds_speeds_legend(
                save_path, test_pieces,
                diameter=test_pieces[0]['diameter'],
                material_display=self.fs_material_var.get(),
                air_assist=bool(self.fs_air_assist_var.get()),
                eng_speed=eng_speed, eng_power=eng_power)
            messagebox.showinfo(
                _("G-code Generated"),
                _("Wrote {n} test discs to:\n{p}\n\nLegend saved to:\n{l}").format(
                    n=len(test_pieces), p=save_path, l=legend_path))

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

            variant_label = {"large": _("Large (40\u201360mm)"),
                             "small": _("Small (7\u201339.5mm)"),
                             "both": _("Both (Small + Large)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {layer_count}-layer {variant_label} die holder.\n\nSaved to: {path}").format(layer_count=layer_count, variant_label=variant_label, path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

    def _on_generate_holder_gcode(self):
        """Generate G-code for die holder pieces."""
        try:
            variant = self.holder_variant_var.get()
            layer_count = int(self.holder_layer_count_var.get())
            base = self.holder_filename_var.get().strip() or "die_holder"

            sheet_w, sheet_h = self._get_holder_sheet_mm()
            if sheet_w is None:
                return

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

            variant_label = {"large": _("Large (40\u201360mm)"),
                             "small": _("Small (7\u201339.5mm)"),
                             "both": _("Both (Small + Large)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {layer_count}-layer {variant_label} die holder G-code.\n\nSaved to: {path}").format(layer_count=layer_count, variant_label=variant_label, path=save_path))

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

            variant_label = {"upper": _("Upper (slotted)"), "lower": _("Lower (base)")}[variant]
            messagebox.showinfo(_("Done"),
                _("Generated {variant_label} die organizer.\n\nSaved to: {path}").format(variant_label=variant_label, path=save_path))

        except Exception as e:
            messagebox.showerror(_("Error"), _("Something went wrong:\n\n{e}").format(e=e))

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
