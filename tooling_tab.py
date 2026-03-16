"""
Tooling tab mixin for Stohrer Sax Shop Companion.

Provides die insert and die holder generation for laser-cutting acrylic tooling.
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys

from svg_engine import (
    _nest_discs, try_nest_partial, compute_remaining_pads, can_all_pads_fit,
    generate_die_svg, generate_die_svg_from_placed,
    generate_holder_svg, generate_kerf_test_svg
)
from gcode_engine import (
    generate_die_gcode_from_placed, generate_holder_gcode, generate_kerf_test_gcode
)
from ui_dialogs import GcodeSettingsWindow

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

    def create_tooling_tab(self, parent):
        """Build the Tooling tab UI with accordion-style tool selector."""
        tooling = self.settings.get("tooling_settings", {})
        bg = self.root.cget('bg') if IS_MACOS else self.default_bg

        # Tool selector buttons at top
        selector_frame = tk.Frame(parent, bg=bg)
        selector_frame.pack(fill='x', padx=10, pady=(10, 5))

        self._tooling_sections = {}
        self._tooling_buttons = {}

        tool_names = [
            ("die_inserts", "Die Inserts"),
            ("die_holders", "Die Holders"),
            ("kerf_test", "Kerf Test"),
        ]
        for key, label in tool_names:
            btn = tk.Button(selector_frame, text=label, width=15,
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

        # Size input
        size_input_frame = tk.Frame(die_frame, bg=bg)
        size_input_frame.pack(fill='x', pady=(0, 5))

        tk.Label(size_input_frame, text="Enter sizes (mm):", bg=bg).pack(anchor='w')
        self.die_size_entry = tk.Text(size_input_frame, height=2, width=50, font=("Courier", 10))
        self.die_size_entry.pack(fill='x', pady=(2, 2))

        hint_label = tk.Label(size_input_frame, text='e.g. "7, 8.5, 15-30" or "7-39.5"',
                              bg=bg, fg="gray", font=("Helvetica", 8))
        hint_label.pack(anchor='w')

        # Quick-fill buttons and step size
        buttons_frame = tk.Frame(die_frame, bg=bg)
        buttons_frame.pack(fill='x', pady=(0, 5))

        tk.Button(buttons_frame, text="Full Set (S)", width=12,
                  command=lambda: self._fill_die_sizes("7-39.5")).pack(side='left', padx=(0, 5))
        tk.Button(buttons_frame, text="Full Set (L)", width=12,
                  command=lambda: self._fill_die_sizes("40-60")).pack(side='left', padx=(0, 5))
        tk.Button(buttons_frame, text="Full Set (All)", width=12,
                  command=lambda: self._fill_die_sizes("7-60")).pack(side='left', padx=(0, 15))

        tk.Label(buttons_frame, text="Step:", bg=bg).pack(side='left')
        self.die_step_var = tk.StringVar(value=tooling.get("step_size", "0.5"))
        step_entry = tk.Entry(buttons_frame, textvariable=self.die_step_var, width=5)
        step_entry.pack(side='left', padx=(2, 2))
        tk.Label(buttons_frame, text="mm", bg=bg).pack(side='left')

        # Engraving options
        eng_frame = tk.Frame(die_frame, bg=bg)
        eng_frame.pack(fill='x', pady=(5, 5))

        self.die_engrave_ring_var = tk.BooleanVar(value=tooling.get("engrave_ring", True))
        tk.Checkbutton(eng_frame, text="Engrave size on ring", variable=self.die_engrave_ring_var,
                       bg=bg).pack(side='left', padx=(0, 15))

        self.die_engrave_cutout_var = tk.BooleanVar(value=tooling.get("engrave_cutout", True))
        tk.Checkbutton(eng_frame, text="Engrave size on cutout", variable=self.die_engrave_cutout_var,
                       bg=bg).pack(side='left', padx=(0, 15))

        # Sheet size
        sheet_frame = tk.LabelFrame(die_frame, text="Sheet Size", bg=bg, padx=5, pady=5)
        sheet_frame.pack(fill='x', pady=(5, 5))

        size_row = tk.Frame(sheet_frame, bg=bg)
        size_row.pack(fill='x')

        tk.Label(size_row, text="Width:", bg=bg).pack(side='left')
        self.die_width_var = tk.StringVar(value=tooling.get("sheet_width", "12"))
        tk.Entry(size_row, textvariable=self.die_width_var, width=8).pack(side='left', padx=(2, 10))

        tk.Label(size_row, text="Height:", bg=bg).pack(side='left')
        self.die_height_var = tk.StringVar(value=tooling.get("sheet_height", "12"))
        tk.Entry(size_row, textvariable=self.die_height_var, width=8).pack(side='left', padx=(2, 10))

        self.die_units_var = tk.StringVar(value=self.settings.get("units", "in"))
        tk.Radiobutton(size_row, text="in", variable=self.die_units_var,
                       value="in", bg=bg).pack(side='left')
        tk.Radiobutton(size_row, text="mm", variable=self.die_units_var,
                       value="mm", bg=bg).pack(side='left', padx=(5, 0))

        # Scrap mode
        scrap_row = tk.Frame(die_frame, bg=bg)
        scrap_row.pack(fill='x', pady=(5, 5))

        self.die_scrap_var = tk.BooleanVar(value=False)
        tk.Checkbutton(scrap_row, text="Scrap Mode", variable=self.die_scrap_var,
                       bg=bg, command=self._toggle_die_scrap_mode).pack(side='left')

        self.die_scrap_clear_btn = tk.Button(scrap_row, text="Clear", width=5,
                                              command=self._on_clear_die_scrap_clicked)
        # Hidden initially
        self.die_scrap_status_var = tk.StringVar(value="")
        self.die_scrap_status_label = tk.Label(scrap_row, textvariable=self.die_scrap_status_var,
                                                bg=bg, font=("Helvetica", 9))

        # Output filename
        name_frame = tk.Frame(die_frame, bg=bg)
        name_frame.pack(fill='x', pady=(5, 5))

        tk.Label(name_frame, text="Output filename:", bg=bg).pack(side='left')
        self.die_filename_var = tk.StringVar(value="my_dies")
        tk.Entry(name_frame, textvariable=self.die_filename_var, width=25).pack(side='left', padx=(5, 0))

        # Generate buttons
        gen_frame = tk.Frame(die_frame, bg=bg)
        gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(gen_frame, text="Generate SVG", width=15,
                  command=self._on_generate_die_svg).pack(side='left', padx=(0, 10))
        tk.Button(gen_frame, text="Generate G-code", width=15,
                  command=self._on_generate_die_gcode).pack(side='left')

        self._tooling_sections['die_inserts'] = die_frame

        # ========================================
        # DIE HOLDERS SECTION
        # ========================================
        holder_frame = tk.Frame(self._tooling_content, bg=bg)

        # Sheet size info
        holder_info = tk.Label(holder_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                               text="6-layer holder: base, magnet (6.5mm), 3\u00d7 pin (3.5mm), ring.\n"
                                    "Required sheet: 275 \u00d7 185 mm (10.8 \u00d7 7.3 in).",
                               justify="left")
        holder_info.pack(anchor='w', pady=(0, 8))

        variant_frame = tk.Frame(holder_frame, bg=bg)
        variant_frame.pack(fill='x', pady=(0, 5))

        tk.Label(variant_frame, text="Generate:", bg=bg).pack(side='left')
        self.holder_variant_var = tk.StringVar(value="large")
        tk.Radiobutton(variant_frame, text="Large (40\u201360mm pads)", variable=self.holder_variant_var,
                       value="large", bg=bg).pack(side='left', padx=(5, 10))
        tk.Radiobutton(variant_frame, text="Small (7\u201339.5mm pads)", variable=self.holder_variant_var,
                       value="small", bg=bg).pack(side='left')

        holder_name_frame = tk.Frame(holder_frame, bg=bg)
        holder_name_frame.pack(fill='x', pady=(0, 5))

        tk.Label(holder_name_frame, text="Output filename:", bg=bg).pack(side='left')
        self.holder_filename_var = tk.StringVar(value="die_holder")
        tk.Entry(holder_name_frame, textvariable=self.holder_filename_var, width=25).pack(side='left', padx=(5, 0))

        holder_gen_frame = tk.Frame(holder_frame, bg=bg)
        holder_gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(holder_gen_frame, text="Generate SVG", width=15,
                  command=self._on_generate_holder_svg).pack(side='left', padx=(0, 10))
        tk.Button(holder_gen_frame, text="Generate G-code", width=15,
                  command=self._on_generate_holder_gcode).pack(side='left')

        self._tooling_sections['die_holders'] = holder_frame

        # ========================================
        # KERF TEST SECTION
        # ========================================
        kerf_frame = tk.Frame(self._tooling_content, bg=bg)

        # Sheet size info
        kerf_info = tk.Label(kerf_frame, bg=bg, font=("Helvetica", 9), fg="gray",
                             text="Pattern size: 110 \u00d7 54 mm (4.3 \u00d7 2.1 in)",
                             justify="left")
        kerf_info.pack(anchor='w', pady=(0, 8))

        # Material dropdown
        mat_row = tk.Frame(kerf_frame, bg=bg)
        mat_row.pack(fill='x', pady=(0, 5))

        tk.Label(mat_row, text="Material:", bg=bg).pack(side='left')
        self.kerf_material_var = tk.StringVar(value="Acrylic")
        kerf_materials = ["Felt", "Card", "Leather", "Acrylic"]
        kerf_dropdown = ttk.Combobox(mat_row, textvariable=self.kerf_material_var,
                                     values=kerf_materials, state="readonly", width=12)
        kerf_dropdown.pack(side='left', padx=(5, 0))

        # Use existing settings checkbox
        settings_row = tk.Frame(kerf_frame, bg=bg)
        settings_row.pack(fill='x', pady=(0, 5))

        self.kerf_use_existing_var = tk.BooleanVar(value=True)
        tk.Checkbutton(settings_row, text="Use existing G-code settings for this material",
                       variable=self.kerf_use_existing_var, bg=bg,
                       command=self._toggle_kerf_custom_settings).pack(anchor='w')

        # Custom settings (hidden by default)
        self.kerf_custom_frame = tk.Frame(kerf_frame, bg=bg)

        custom_row = tk.Frame(self.kerf_custom_frame, bg=bg)
        custom_row.pack(fill='x')

        tk.Label(custom_row, text="Cut speed:", bg=bg).pack(side='left')
        self.kerf_speed_var = tk.StringVar(value="180")
        tk.Entry(custom_row, textvariable=self.kerf_speed_var, width=8).pack(side='left', padx=(2, 5))
        tk.Label(custom_row, text="mm/min", bg=bg).pack(side='left', padx=(0, 15))

        tk.Label(custom_row, text="Cut power:", bg=bg).pack(side='left')
        self.kerf_power_var = tk.StringVar(value="100")
        tk.Entry(custom_row, textvariable=self.kerf_power_var, width=8).pack(side='left', padx=(2, 5))
        tk.Label(custom_row, text="%", bg=bg).pack(side='left')

        # Generate buttons
        kerf_gen_frame = tk.Frame(kerf_frame, bg=bg)
        kerf_gen_frame.pack(fill='x', pady=(5, 0))

        tk.Button(kerf_gen_frame, text="Generate SVG", width=15,
                  command=self._on_generate_kerf_svg).pack(side='left', padx=(0, 10))
        tk.Button(kerf_gen_frame, text="Generate G-code", width=15,
                  command=self._on_generate_kerf_gcode).pack(side='left')

        self._tooling_sections['kerf_test'] = kerf_frame

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
            messagebox.showerror("Invalid Input", "Sheet width and height must be valid numbers.")
            return None, None

        if w <= 0 or h <= 0:
            messagebox.showerror("Invalid Input", "Sheet dimensions must be positive.")
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
        acrylic_materials = [("acrylic", "Acrylic")]
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
        self.settings["tooling_settings"] = tooling

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
                messagebox.showerror("Error", "No valid die sizes entered.")
                return

            width_mm, height_mm = self._get_die_sheet_mm()
            if width_mm is None:
                return

            pads = self._sizes_to_pads(sizes)
            settings = self._get_die_settings()

            if not can_all_pads_fit(pads, 'die_ring', width_mm, height_mm, settings):
                messagebox.showerror("Nesting Error",
                    f"Could not fit all {len(sizes)} dies on the specified sheet.\n"
                    "Try a larger sheet or fewer sizes, or use Scrap Mode.")
                return

            base = self.die_filename_var.get().strip() or "my_dies"
            save_path = filedialog.asksaveasfilename(
                title="Save Die Insert SVG",
                defaultextension=".svg",
                filetypes=[("SVG files", "*.svg")],
                initialfile=f"{base}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            placed = generate_die_svg(pads, width_mm, height_mm, save_path, settings)

            messagebox.showinfo("Done",
                f"Generated {len(placed)} die rings.\n\nSaved to: {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    def _on_generate_die_gcode(self):
        """Generate G-code for die inserts."""
        if self.die_scrap_var.get():
            self._generate_die_gcode_scrap()
            return

        try:
            sizes = self._parse_die_sizes()
            if not sizes:
                messagebox.showerror("Error", "No valid die sizes entered.")
                return

            width_mm, height_mm = self._get_die_sheet_mm()
            if width_mm is None:
                return

            pads = self._sizes_to_pads(sizes)
            settings = self._get_die_settings()

            if not can_all_pads_fit(pads, 'die_ring', width_mm, height_mm, settings):
                messagebox.showerror("Nesting Error",
                    f"Could not fit all {len(sizes)} dies on the specified sheet.\n"
                    "Try a larger sheet or fewer sizes, or use Scrap Mode.")
                return

            base = self.die_filename_var.get().strip() or "my_dies"
            save_path = filedialog.asksaveasfilename(
                title="Save Die Insert G-code",
                defaultextension=".gcode",
                filetypes=[("G-code files", "*.gcode")],
                initialfile=f"{base}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)

            placed, _, _ = _nest_discs(pads, 'die_ring', width_mm, height_mm, settings)
            generate_die_gcode_from_placed(placed, width_mm, height_mm, save_path, settings)

            messagebox.showinfo("Done",
                f"Generated G-code for {len(placed)} die rings.\n\nSaved to: {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    # ========================================
    # DIE SCRAP MODE
    # ========================================

    def _toggle_die_scrap_mode(self):
        """Handle die scrap mode checkbox toggle."""
        if not self.die_scrap_var.get():
            if self.tooling_scrap_session['active']:
                remaining = self._count_die_remaining()
                if remaining > 0:
                    if not messagebox.askyesno("Clear Session?",
                        f"You have {remaining} dies remaining.\n"
                        "Disabling scrap mode will clear the session.\n\nContinue?"):
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
                if not messagebox.askyesno("Clear Session?",
                    f"You have {remaining} dies remaining.\n\nClear session?"):
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
                    messagebox.showerror("Error", "No valid die sizes entered.")
                    return
                pads = self._sizes_to_pads(sizes)

                save_dir = filedialog.askdirectory(
                    title="Select Folder to Save SVGs",
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir
                self._start_die_scrap_session(pads, save_dir)
            else:
                pads = self.tooling_scrap_session['remaining_pads']

            if not pads:
                messagebox.showinfo("Session Complete", "All dies have been placed!")
                return

            placed, remaining, any_placed = try_nest_partial(
                pads, 'die_ring', width_mm, height_mm, settings)

            if not any_placed:
                min_size = min(p['size'] for p in pads)
                messagebox.showwarning("No Dies Fit",
                    f"No dies could be placed on this sheet.\n\n"
                    f"Smallest remaining die: {min_size}mm\n"
                    f"Try a larger sheet.")
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
                messagebox.showinfo("Session Complete!",
                    f"Placed {placed_count} dies on sheet #{scrap_num}.\n\n"
                    f"All dies placed! Session complete.\n"
                    f"Files saved to: {save_dir}")
            else:
                messagebox.showinfo("Sheet Generated",
                    f"Placed {placed_count} dies on sheet #{scrap_num}.\n\n"
                    f"{remaining_count} dies remaining.\n"
                    f"Adjust sheet size if needed and generate again.")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

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
                    messagebox.showerror("Error", "No valid die sizes entered.")
                    return
                pads = self._sizes_to_pads(sizes)

                save_dir = filedialog.askdirectory(
                    title="Select Folder to Save G-code",
                    initialdir=self.settings.get("last_output_dir", ""))
                if not save_dir:
                    return
                self.settings["last_output_dir"] = save_dir
                self._start_die_scrap_session(pads, save_dir)
            else:
                pads = self.tooling_scrap_session['remaining_pads']

            if not pads:
                messagebox.showinfo("Session Complete", "All dies have been placed!")
                return

            placed, remaining, any_placed = try_nest_partial(
                pads, 'die_ring', width_mm, height_mm, settings)

            if not any_placed:
                min_size = min(p['size'] for p in pads)
                messagebox.showwarning("No Dies Fit",
                    f"No dies could be placed on this sheet.\n\n"
                    f"Smallest remaining die: {min_size}mm\n"
                    f"Try a larger sheet.")
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
                messagebox.showinfo("Session Complete!",
                    f"Generated G-code for {placed_count} dies on sheet #{scrap_num}.\n\n"
                    f"All dies placed! Session complete.\n"
                    f"Files saved to: {save_dir}")
            else:
                messagebox.showinfo("Sheet Generated",
                    f"Generated G-code for {placed_count} dies on sheet #{scrap_num}.\n\n"
                    f"{remaining_count} dies remaining.\n"
                    f"Adjust sheet size if needed and generate again.")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

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
        self.tooling_scrap_window.title("Die Scrap Mode Progress")
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
        tk.Label(remaining_frame, text="Remaining", bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="blue").pack(anchor="w")
        self.die_remaining_listbox = tk.Listbox(remaining_frame, font=("Courier", 10), height=10, width=12)
        self.die_remaining_listbox.pack(fill="both", expand=True)

        done_frame = tk.Frame(columns_frame, bg=theme_bg)
        done_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        tk.Label(done_frame, text="Done", bg=theme_bg,
                 font=("Helvetica", 9, "bold"), fg="green").pack(anchor="w")
        self.die_done_listbox = tk.Listbox(done_frame, font=("Courier", 10), height=10, width=12)
        self.die_done_listbox.pack(fill="both", expand=True)

        self.die_scrap_footer = tk.Label(self.tooling_scrap_window, text="", bg=theme_bg,
                                          font=("Helvetica", 9))
        self.die_scrap_footer.pack(pady=(0, 5))

        tk.Button(self.tooling_scrap_window, text="Close",
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
                self.die_scrap_header.config(text="All done!", fg="green")
            else:
                self.die_scrap_header.config(
                    text=f"{total_done} / {total_original} dies complete", fg="blue")

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

            self.die_scrap_footer.config(text=f"die_ring | {scraps} sheet(s) used")

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
            settings_path = "Tooling > Options > Settings > Acrylic > Kerf Width"
        else:
            settings_path = f"Pad Generator > Options > G-code Settings > {material_name} > Kerf Width"

        dlg = tk.Toplevel(self.root)
        dlg.title(f"Kerf Test \u2014 {material_name}")
        dlg.geometry("420x310")
        theme_bg = self._get_theme_color()
        dlg.configure(bg=theme_bg)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text=f"Kerf Test \u2014 {material_name}", bg=theme_bg,
                 font=("Helvetica", 12, "bold")).pack(pady=(15, 10))

        steps = [
            "1.  Cut the pattern on your material",
            "2.  Pop out the 3 discs (10mm, 20mm, 30mm)",
            "3.  For each circle, measure the hole ID\n"
            "     and disc OD with calipers",
            "4.  Kerf = hole ID \u2212 disc OD",
            "5.  Average the 3 results for best accuracy",
            f"6.  Enter the full kerf value in:\n     {settings_path}",
        ]
        for step in steps:
            tk.Label(dlg, text=step, bg=theme_bg, font=("Helvetica", 10),
                     justify="left", anchor="w").pack(fill="x", padx=20, pady=1)

        dont_show_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dlg, text="Don't show this again", variable=dont_show_var,
                       bg=theme_bg).pack(pady=(10, 5))

        def on_ok():
            if dont_show_var.get():
                self.settings["seen_kerf_test_tutorial"] = True
                save_settings(self.settings)
            dlg.destroy()

        tk.Button(dlg, text="OK", command=on_ok, width=10).pack(pady=(5, 15))

    def _on_generate_kerf_svg(self):
        """Generate SVG kerf test pattern."""
        try:
            material_name = self.kerf_material_var.get()
            save_path = filedialog.asksaveasfilename(
                title="Save Kerf Test SVG",
                defaultextension=".svg",
                filetypes=[("SVG files", "*.svg")],
                initialfile=f"kerf_test_{material_name.lower()}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            generate_kerf_test_svg(material_name, save_path, self.settings)
            self._show_kerf_instructions(material_name)

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

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
                    messagebox.showerror("Invalid Input", "Speed and power must be valid numbers.")
                    return
                eng_mode = "line"
                eng_speed = 1000
                eng_power = 15
                filled_spacing = 0.15

            save_path = filedialog.asksaveasfilename(
                title="Save Kerf Test G-code",
                defaultextension=".gcode",
                filetypes=[("G-code files", "*.gcode")],
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
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    # ========================================
    # DIE HOLDER GENERATION
    # ========================================

    def _on_generate_holder_svg(self):
        """Generate SVG for die holder pieces."""
        try:
            variant = self.holder_variant_var.get()
            base = self.holder_filename_var.get().strip() or "die_holder"

            save_path = filedialog.asksaveasfilename(
                title="Save Die Holder SVG",
                defaultextension=".svg",
                filetypes=[("SVG files", "*.svg")],
                initialfile=f"{base}.svg",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            settings = self._get_die_settings()
            generate_holder_svg(variant, save_path, settings)

            variant_label = {"large": "Large (40\u201360mm)", "small": "Small (7\u201339.5mm)"}[variant]
            messagebox.showinfo("Done",
                f"Generated {variant_label} die holder.\n\nSaved to: {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")

    def _on_generate_holder_gcode(self):
        """Generate G-code for die holder pieces."""
        try:
            variant = self.holder_variant_var.get()
            base = self.holder_filename_var.get().strip() or "die_holder"

            save_path = filedialog.asksaveasfilename(
                title="Save Die Holder G-code",
                defaultextension=".gcode",
                filetypes=[("G-code files", "*.gcode")],
                initialfile=f"{base}.gcode",
                initialdir=self.settings.get("last_output_dir", ""))
            if not save_path:
                return

            self.settings["last_output_dir"] = os.path.dirname(save_path)
            settings = self._get_die_settings()
            generate_holder_gcode(variant, save_path, settings)

            variant_label = {"large": "Large (40\u201360mm)", "small": "Small (7\u201339.5mm)"}[variant]
            messagebox.showinfo("Done",
                f"Generated {variant_label} die holder G-code.\n\nSaved to: {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n\n{e}")
