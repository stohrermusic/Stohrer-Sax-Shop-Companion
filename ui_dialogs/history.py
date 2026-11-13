"""Job History.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import messagebox
import datetime



from .common import DIALOG_BG, add_tooltip, material_display_name


def tooling_tool_name(tool):
    """The Tooling section's title for a Job History entry's `tool` key
    (the keys are the section keys in tooling_tab.create_tooling_tab)."""
    return {
        'die_inserts': _("Die Inserts"),
        'die_holders': _("Die Holders"),
        'die_organizer': _("Die Organizer"),
        'kerf_test': _("Kerf Test"),
        'feeds_speeds_tester': _("Speed & Power Test"),
        'tonehole_crescent': _("Tonehole Crescent"),
    }.get(tool) or str(tool)


def tooling_job_not_loadable():
    """Why a Tooling row has no Load: said in the detail pane and on a
    double-click, the one place the sentence is written."""
    return _("This is a Tooling run. Only Pad Maker jobs can be loaded into the form.")


class JobHistoryWindow(tk.Toplevel):
    """Browse past jobs that reached an output stage: the Pad Maker's
    batches and, since 2026-11-09, the Tooling tab's outputs (entries with
    a `tool` key: shown with the tool's name, never loaded into the form).

    `jobs` is the list from config.load_job_history() (newest first);
    `save_history` is config.save_job_history, called after a delete so
    the trimmed list hits disk. On close, `self.result` holds the job
    the user chose to reload, or None.
    """

    # Column widths are measured from the content at refresh time rather
    # than hardcoded: translated headers and values vary a lot in length
    # ("Pads" is "Zapatillas" in Spanish), and a fixed width would either
    # truncate them or leave the Sheet column ragged.
    _MAX_COL = 30      # safety cap so one corrupt entry can't span the pane
    _COL_GAP = 2

    def __init__(self, parent, jobs, save_history):
        super().__init__(parent)
        self.title(_("Job History"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.grab_set()
        self.geometry("820x560")
        self.minsize(640, 420)

        self.jobs = list(jobs)
        self._save_history = save_history
        self.result = None

        tk.Label(self, text=_("Past Jobs"), bg=DIALOG_BG,
                 font=("Helvetica", 14, "bold")).pack(pady=(12, 0))
        tk.Label(self,
                 text=_("Every batch that was written to file or sent to the "
                        "laser, newest first.\n"
                        "A * on the output means that laser run ended early."),
                 bg=DIALOG_BG, font=("Helvetica", 9),
                 justify="center").pack(pady=(2, 8))

        # --- Job list ---
        list_frame = tk.Frame(self, bg=DIALOG_BG)
        list_frame.pack(fill="both", expand=True, padx=12)

        self.header_label = tk.Label(list_frame, text="", bg=DIALOG_BG,
                                     anchor="w", font=("Courier", 9, "bold"))
        self.header_label.pack(fill="x")

        inner = tk.Frame(list_frame, bg=DIALOG_BG)
        inner.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(inner, font=("Courier", 9),
                                  activestyle="dotbox", exportselection=False)
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(inner, command=self.listbox.yview)
        scroll.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scroll.set)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._on_select())
        self.listbox.bind("<Double-Button-1>", lambda e: self.on_load())

        # --- Detail pane ---
        detail_frame = tk.Frame(self, bg=DIALOG_BG)
        detail_frame.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        self.detail = tk.Text(detail_frame, height=9, wrap="word",
                              font=("Helvetica", 9), bg=DIALOG_BG,
                              relief="flat", padx=6, pady=4, state="disabled")
        d_scroll = tk.Scrollbar(detail_frame, command=self.detail.yview)
        self.detail.configure(yscrollcommand=d_scroll.set)
        d_scroll.pack(side="right", fill="y")
        self.detail.pack(side="left", fill="both", expand=True)

        # --- Buttons ---
        btns = tk.Frame(self, bg=DIALOG_BG)
        btns.pack(fill="x", padx=12, pady=12)
        self.load_btn = tk.Button(btns, text=_("Load into Pad Maker"),
                                  command=self.on_load, state="disabled",
                                  font=("Helvetica", 10, "bold"))
        self.load_btn.pack(side="left")
        self.delete_btn = tk.Button(btns, text=_("Delete Entry"),
                                    command=self.on_delete, state="disabled")
        self.delete_btn.pack(side="left", padx=6)
        tk.Button(btns, text=_("Clear History"),
                  command=self.on_clear).pack(side="left")
        tk.Button(btns, text=_("Close"), command=self.destroy,
                  width=10).pack(side="right")

        add_tooltip(self.load_btn,
                    _("Put this job's pad list, materials, sheet size and "
                      "filename back into the form so you can run it again. "
                      "Sizing rules and custom shapes aren't changed."))
        add_tooltip(self.delete_btn, _("Remove just this job from the history."))

        self._refresh_list()
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.wait_window(self)

    # ---------- formatting ----------

    @staticmethod
    def _when(job):
        """'2026-08-01 14:22' from the stored ISO timestamp."""
        raw = job.get("timestamp") or ""
        try:
            return datetime.datetime.fromisoformat(raw).strftime("%Y-%m-%d %H:%M")
        except (ValueError, TypeError):
            return raw[:16]

    @staticmethod
    def _output_label(job):
        """Short output description, flagged when a laser run was cut short."""
        out = job.get("output")
        if out == "laser":
            if job.get("status") not in (None, "complete"):
                return _("Laser*")
            return _("Laser")
        if out == "gcode":
            return _("G-code")
        if out == "svg":
            return _("SVG")
        if out == "stl":
            return "STL"        # a tonehole crescent print; the format's name is the same in every language
        return out or "?"

    @staticmethod
    def _materials_label(job):
        # Same display transform as the Materials checkboxes on the tab,
        # so a Spanish user reads "Cuero" here and "Cuero" there.
        mats = job.get("materials") or []
        if not mats:
            return "-"
        return "+".join(material_display_name(m) for m in mats)

    def _materials_short(self, job):
        """Material names for the list column, or a count once the join
        gets unwieldy ('felt+card+leather+exact size' crowds out the rest
        of the row, and the detail pane spells it out anyway)."""
        mats = job.get("materials") or []
        if len(mats) > 2:
            return ngettext("{n} material", "{n} materials",
                            len(mats)).format(n=len(mats))
        return self._materials_label(job)

    @staticmethod
    def _pad_count(job):
        """Discs on the sheet, falling back to what was asked for.

        placed_count is the truth for a 'max' fill or a scrap subset;
        older/partial entries may only have the requested total.
        """
        n = job.get("placed_count")
        if not isinstance(n, int) or n <= 0:
            n = job.get("requested_count")
        return n if isinstance(n, int) else 0

    def _sheet_label(self, job):
        w, h, u = job.get("sheet_w"), job.get("sheet_h"), job.get("units")
        if w and h:
            return f"{w} x {h} {u or ''}".strip()
        w_mm, h_mm = job.get("sheet_w_mm"), job.get("sheet_h_mm")
        if isinstance(w_mm, (int, float)) and isinstance(h_mm, (int, float)):
            return f"{w_mm:g} x {h_mm:g} mm"
        return "-"

    @staticmethod
    def _headers():
        return [_("When"), _("Output"), _("Material"), _("Pads"), _("Sheet")]

    def _cells(self, job):
        """The five column values for one job, unpadded. A Tooling row
        names its tool where a pad row counts its pads."""
        if job.get("tool"):
            pad_txt = tooling_tool_name(job["tool"])
        else:
            pads = self._pad_count(job)
            pad_txt = ngettext("{n} pad", "{n} pads", pads).format(n=pads)
        if job.get("scrap_num"):
            pad_txt += f" #{job['scrap_num']}"
        return [
            self._when(job),
            self._output_label(job),
            self._materials_short(job),
            pad_txt,
            self._sheet_label(job),
        ]

    def _column_widths(self, jobs):
        """Width per column: the widest of header and values, capped."""
        rows = [self._headers()] + [self._cells(j) for j in jobs]
        return [min(max(len(r[i]) for r in rows) + self._COL_GAP, self._MAX_COL)
                for i in range(len(self._headers()))]

    def _join(self, cells, widths):
        """Pad each cell to its column width; the last one runs free."""
        out = []
        for i, cell in enumerate(cells):
            cell = str(cell)
            if i == len(cells) - 1:
                out.append(cell)
                break
            if len(cell) >= widths[i]:
                cell = cell[:widths[i] - 1]
            out.append(cell.ljust(widths[i]))
        return "".join(out)

    def _row_text(self, job, widths=None):
        cells = self._cells(job)
        if widths is None:
            widths = self._column_widths([job])
        return self._join(cells, widths)

    # ---------- list / detail ----------

    def _refresh_list(self, keep_index=None):
        self.listbox.delete(0, tk.END)
        widths = self._column_widths(self.jobs)
        self.header_label.config(text=self._join(self._headers(), widths))

        if not self.jobs:
            self.listbox.insert(tk.END, _("  (no jobs yet — generate or cut "
                                          "something and it'll show up here)"))
            self.listbox.config(state="disabled")
            self._set_detail("")
            self.load_btn.config(state="disabled")
            self.delete_btn.config(state="disabled")
            return

        self.listbox.config(state="normal")
        for job in self.jobs:
            self.listbox.insert(tk.END, self._row_text(job, widths))

        if keep_index is not None and self.jobs:
            idx = min(keep_index, len(self.jobs) - 1)
            self.listbox.selection_set(idx)
            self.listbox.see(idx)
        self._on_select()

    def _selected_job(self):
        if not self.jobs:
            return None
        sel = self.listbox.curselection()
        if not sel or sel[0] >= len(self.jobs):
            return None
        return self.jobs[sel[0]]

    def _set_detail(self, text):
        self.detail.config(state="normal")
        self.detail.delete("1.0", tk.END)
        self.detail.insert("1.0", text)
        self.detail.config(state="disabled")

    def _on_select(self):
        job = self._selected_job()
        state = "normal" if job else "disabled"
        # A Tooling row has nothing to put back into the Pad Maker's form.
        self.load_btn.config(state="disabled" if (not job or job.get("tool")) else "normal")
        self.delete_btn.config(state=state)
        self._set_detail(self._detail_text(job) if job else "")

    def _detail_text(self, job):
        lines = []
        out = job.get("output")
        tooling = bool(job.get("tool"))
        if out == "laser":
            what = _("Sent to the laser")
            status = job.get("status")
            if status and status != "complete":
                what += _(" — run ended early ({status})").format(status=status)
        elif out == "gcode":
            what = _("G-code files written")
        elif out == "svg":
            what = _("SVG files written")
        elif out == "stl":
            what = _("STL file written")
        else:
            what = str(out)
        lines.append(f"{self._when(job)}  —  {what}")

        if tooling:
            lines.append(tooling_tool_name(job["tool"]))

        if job.get("scrap_num"):
            lines.append(_("Scrap piece #{n} of a scrap-mode session").format(
                n=job["scrap_num"]))

        lines.append(_("Material: {m}").format(m=self._materials_label(job)))

        pads = self._pad_count(job)
        if tooling:
            pad_line = ngettext("{n} piece on the sheet", "{n} pieces on the sheet",
                                pads).format(n=pads)
        else:
            pad_line = ngettext("{n} pad on the sheet", "{n} pads on the sheet",
                                pads).format(n=pads)
        if job.get("has_max"):
            pad_line += _(" (a 'max' line filled the leftover space)")
        lines.append(pad_line)

        sheet = _("Sheet: {s}").format(s=self._sheet_label(job))
        if job.get("card_paper"):
            sheet += _(" (card used the paper size setting)")
        lines.append(sheet)

        hole = job.get("hole_option")
        if hole == "Custom":
            lines.append(_("Center hole: {v}mm (custom)").format(
                v=job.get("custom_hole", "?")))
        elif hole:
            lines.append(_("Center hole: {v}").format(v=hole))

        if job.get("base"):
            lines.append(_("Filename: {b}").format(b=job["base"]))
        if job.get("save_dir"):
            lines.append(_("Saved to: {d}").format(d=job["save_dir"]))
        if job.get("preset_name"):
            lines.append(_("From preset: [{lib}] {name}").format(
                lib=job.get("preset_library", "?"), name=job["preset_name"]))
        if job.get("polygon_vertices"):
            lines.append(ngettext(
                "Custom shape: {n} point (not restored on load)",
                "Custom shape: {n} points (not restored on load)",
                job["polygon_vertices"]).format(n=job["polygon_vertices"]))

        pad_text = job.get("pad_text")
        if pad_text and tooling:
            # The tool's own description of the job (die sizes, holder
            # variant, the kerf test's cut settings, ...), as it stands.
            lines.append("")
            lines.append(pad_text)
        elif pad_text:
            lines.append("")
            # One comma-joined run rather than one line per size — a
            # 20-size tenor set would otherwise push everything else out
            # of the pane. The Text widget word-wraps it.
            entries = [ln.strip() for ln in pad_text.splitlines() if ln.strip()]
            lines.append(_("Pad list: {list}").format(list=", ".join(entries)))

        if tooling:
            lines.append("")
            lines.append(tooling_job_not_loadable())

        return "\n".join(lines)

    # ---------- actions ----------

    def on_load(self):
        job = self._selected_job()
        if not job:
            return
        if job.get("tool"):
            # The button is disabled for these rows; this is the double-click.
            messagebox.showinfo(_("Load Job"), tooling_job_not_loadable())
            return
        if not messagebox.askyesno(
                _("Load Job"),
                _("Replace what's currently in the Pad Maker form with this "
                  "job?\n\n{when} — {mats}\n\nYour sizing rules and G-code "
                  "settings stay as they are.").format(
                      when=self._when(job),
                      mats=self._materials_label(job))):
            return
        self.result = job
        self.destroy()

    def on_delete(self):
        job = self._selected_job()
        if not job:
            return
        idx = self.listbox.curselection()[0]
        if not messagebox.askyesno(
                _("Delete Entry"),
                _("Remove this job from the history?\n\n"
                  "{row}").format(row=self._row_text(job).strip())):
            return
        del self.jobs[idx]
        self._save_history(self.jobs)
        self._refresh_list(keep_index=idx)

    def on_clear(self):
        if not self.jobs:
            return
        if not messagebox.askyesno(
                _("Clear History"),
                _("Delete all {n} jobs from the history?\n\n"
                  "This only clears the log — your presets and files are "
                  "untouched.").format(n=len(self.jobs))):
            return
        self.jobs = []
        self._save_history(self.jobs)
        self._refresh_list()
