"""The resonance theme windows.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import messagebox, ttk
import random


from config import (
    get_resonance_messages,
)

from .common import DIALOG_BG
        
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
