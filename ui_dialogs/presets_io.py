"""Importing and exporting presets, the key layout, pad notes.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import json


from config import (
    DEFAULT_SETTINGS, ALL_KEY_HEIGHT_FIELDS, save_presets,
)

from .common import DIALOG_BG, add_tooltip, bind_mousewheel

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
