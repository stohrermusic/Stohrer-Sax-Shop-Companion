"""The User Guide, the Quick Start Guide, the Welcome window and About.

Cut from ui_dialogs.py on 2026-11-10, a move only: every body is the old text. Imported names a dialog here
constructs that suites stand in for on the package (ui_dialogs.X = Fake) are reached through the package
at call time, never bound at import; see ui_dialogs/__init__.py."""
import tkinter as tk
from tkinter import ttk
import copy
import math


from config import (
    DEFAULT_SETTINGS, APP_VERSION, APP_BUILD_DATE,
)
from svg_engine import (
    get_disc_diameter, adhesive_ring_geometry,
)

from .common import DIALOG_BG, IS_MACOS, _PREVIEW_COLORS
from .sizing import PadPreviewWindow

# Help > About links here; GitHub sends it to the newest release's page.
LATEST_RELEASE_URL = (
    "https://github.com/stohrermusic/Stohrer-Sax-Shop-Companion/releases/latest")


# ==========================================
# HELP / ABOUT DIALOGS
# ==========================================

class _FigureCanvasHost:
    """Lends PadPreviewWindow's disc drawing to a User Guide figure: those
    methods need only `self.canvas` (and LOCATOR_COLOR), so the guide's
    pictures are the preview's own geometry, not a second renderer."""
    LOCATOR_COLOR = PadPreviewWindow.LOCATOR_COLOR
    _draw_leather = PadPreviewWindow._draw_leather
    _draw_locator = PadPreviewWindow._draw_locator
    _draw_adhesive_ring = PadPreviewWindow._draw_adhesive_ring

    def __init__(self, canvas):
        self.canvas = canvas


class UserGuideWindow(tk.Toplevel):
    """The User Guide (Help > User Guide; Help > Welcome and the Quick
    Start Guide link here too).

    Matt, 2026-11-09: "let's give the user guide a bit of an upgrade- keep
    it concise, but have it be a little more of a guide, vs just blocks of
    description. have a TOC that is clickable on the right, and have it
    open to full screen. put example graphics where appropriate."

    So: the WHOLE guide, always, with a clickable contents list down the
    right side (sections and sub-sections; a click scrolls to the place;
    the entry for the part being read is highlighted as the user scrolls),
    a search box (Ctrl+F / Cmd+F), and drawn figures. `section` is the
    tab's key to open AT: Help > User Guide scrolls to the section for the
    tab the user is on. The old one-tab filtering and its "Show Full Guide"
    button are gone, so search always covers everything.

    Opens maximized: a normal window filling the screen with its title bar
    (state 'zoomed'; macOS Aqua, which has no such state, gets the largest
    size Tk allows, the visible frame: see _maximize), not kiosk full
    screen. NOT transient: a transient window loses its maximize
    button on Windows (the toner's Analyze window learned that first). On a
    wide screen the text does not run edge to edge: the reading column stays
    TEXT_MAX_PX wide (a right margin on every line, refitted on resize).

    Figures are small canvases drawn from the app's own geometry helpers
    (the sizing rules, the dart wave, the nester, the ring and crescent
    geometry, the die engraving's arc text), so they are always right, need
    no image files and carry no baked-in English; every label in them goes
    through _(), and every colour is explicit (macOS dark mode).
    """

    SECTION_KEYS = ("pad_generator", "laser", "tooling", "key_heights", "serial_lookup",
                    "screw_specs", "tuner", "toner", "sharing")
    # Kept for callers and tests that map a tab to its section.
    SECTION_TITLES = {
        "pad_generator": _("Pad Maker"),
        "key_heights": _("Key Height Library"),
        "serial_lookup": _("Serial Lookup"),
        "screw_specs": _("Screw Specs"),
        "tooling": _("Tooling"),
        "tuner": _("Tuner"),
        "toner": _("Analyzer"),
    }
    TEXT_MAX_PX = 760       # the reading column
    TEXT_PADX = 15
    FIGURE_W = 520          # px, every figure canvas
    TOC_WIDTH_PX = 250
    FIGURE_BG = "#FFFFFF"
    FIGURE_INK = "#222222"
    FIGURE_FAINT = "#9A9A9A"
    FIGURE_SHEET = "#F8F8F0"
    FIGURE_ZONE = "#2E7D32"
    FIGURE_CORK = "#C9A26B"
    FIGURE_NAMES = ("pad_discs", "nest_zone", "adhesive_rings", "die_ring", "kerf", "crescent", "three_ways")

    def __init__(self, parent, section=None):
        super().__init__(parent)
        self.opened_at = section            # the tab's section this guide opened at (None: the top)
        self.title(_("User Guide"))
        self.configure(bg=DIALOG_BG)
        self._maximize()

        # Search (2026-11-09): the Welcome window tells people to search
        # the guide, so the guide can be searched. Typing highlights every
        # match and jumps to the first; Enter or Next goes to the next and
        # wraps round. The whole guide is always loaded, so it always
        # searches the whole guide.
        search_row = tk.Frame(self, bg=DIALOG_BG)
        search_row.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(search_row, text=_("Search:"), bg=DIALOG_BG).pack(side="left")
        self._search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_row, textvariable=self._search_var, width=28)
        self.search_entry.pack(side="left", padx=(5, 5))
        self._search_next_btn = tk.Button(search_row, text=_("Next"), command=self.search_next)
        self._search_next_btn.pack(side="left")
        self._search_status_var = tk.StringVar(value="")
        tk.Label(search_row, textvariable=self._search_status_var, bg=DIALOG_BG,
                 fg="#555555").pack(side="left", padx=(8, 0))
        self._search_hits = []          # (start, end) Text indexes, in reading order
        self._search_at = -1
        self._search_var.trace_add("write", lambda *a: self._search_update())
        self.search_entry.bind("<Return>", lambda e: self.search_next())
        self.search_entry.bind("<KP_Enter>", lambda e: self.search_next())
        self.bind("<Command-f>" if IS_MACOS else "<Control-f>", lambda e: self.focus_search())

        body = tk.Frame(self, bg=DIALOG_BG)
        body.pack(fill="both", expand=True, padx=10, pady=(6, 0))

        # The contents list, down the right side.
        toc_frame = tk.Frame(body, bg=DIALOG_BG, width=self.TOC_WIDTH_PX)
        toc_frame.pack(side="right", fill="y", padx=(10, 0), pady=(4, 0))
        toc_frame.pack_propagate(False)
        tk.Label(toc_frame, text=_("Contents"), bg=DIALOG_BG, anchor="w",
                 font=("Helvetica", 11, "bold")).pack(fill="x", pady=(4, 4))
        toc_box = tk.Frame(toc_frame, bg=DIALOG_BG)
        toc_box.pack(fill="both", expand=True)
        toc_scroll = tk.Scrollbar(toc_box)
        toc_scroll.pack(side="right", fill="y")
        self.toc = ttk.Treeview(toc_box, show="tree", selectmode="browse",
                                yscrollcommand=toc_scroll.set)
        self.toc.pack(side="left", fill="both", expand=True)
        toc_scroll.config(command=self.toc.yview)
        # A click, not <<TreeviewSelect>>: selection_set fires that event too,
        # after any guard flag has been reset, so the scroll sync and the
        # click handler would select each other for ever (found 2026-11-09).
        self.toc.bind("<ButtonRelease-1>", self._on_toc_click)
        self.toc.bind("<Return>", self._on_toc_click)

        # The text, with its own scrollbar, on the left.
        text_frame = tk.Frame(body, bg=DIALOG_BG)
        text_frame.pack(side="left", fill="both", expand=True)
        self._scrollbar = tk.Scrollbar(text_frame)
        self._scrollbar.pack(side="right", fill="y")
        self.text = tk.Text(text_frame, wrap="word", yscrollcommand=self._on_text_scroll,
                            font=("Helvetica", 11), padx=self.TEXT_PADX, pady=10,
                            highlightthickness=0)
        self.text.pack(fill="both", expand=True)
        self._scrollbar.config(command=self.text.yview)

        self.text.tag_configure("h1", font=("Helvetica", 17, "bold"), spacing1=18, spacing3=6)
        self.text.tag_configure("h2", font=("Helvetica", 13, "bold"), spacing1=12, spacing3=4)
        self.text.tag_configure("body", font=("Helvetica", 11), spacing1=2)
        self.text.tag_configure("bullet", font=("Helvetica", 11), lmargin1=20, lmargin2=35)
        self.text.tag_configure("step", font=("Helvetica", 11), lmargin1=20, lmargin2=38)
        # The reading column: a right margin on every line, set in _fit_column.
        self.text.tag_configure("col", rmargin=0)
        self.text.tag_lower("col")
        # Explicit colours, so a match shows the same in macOS dark mode.
        self.text.tag_configure("search_hit", background="#FFE98A", foreground="#000000")
        self.text.tag_configure("search_current", background="#FF9F1C", foreground="#000000")
        self.text.tag_raise("search_current")

        self._marks = []            # (mark name, contents item) in reading order
        self._toc_items = {}        # section key -> contents item
        self._figures = []          # (name, canvas), in reading order
        self._cur_section_item = None
        self._current_item = None
        self._fold_head = None      # the folded part's heading mark (_fold_start), or None
        self.fold_open = True
        self._insert_content()
        self.text.tag_add("col", "1.0", "end")
        self.text.config(state="disabled")
        self.text.bind("<Configure>", self._fit_column)

        btn_frame = tk.Frame(self, bg=DIALOG_BG)
        btn_frame.pack(pady=8)
        tk.Button(btn_frame, text=_("Close"), command=self.destroy).pack()

        # Open at the tab's section: after the first layout, so the heading
        # really lands at the top (no timer; an idle task, run now).
        self.update_idletasks()
        if section in self._toc_items:
            self.scroll_to(section)
        else:
            self._sync_toc()

    # ---- window -----------------------------------------------------------

    def _maximize(self):
        """A normal maximized window. Where the window manager has no
        'zoomed' state (macOS Aqua), the window is made as large as Tk says
        a window may be: `wm maxsize`, with no size set by the script, is
        the screens' visible frame (the screen less the menu bar and the
        Dock) less the window frame's own decorations, so nothing is
        clipped. Asked for at +0+0 (Tk's y=0 is the top of the screen);
        AppKit moves the title bar down below the menu bar, which puts the
        frame exactly on the visible frame. A screen-sized geometry was
        asked for before (2026-11-09): CI's Macs lose 122 to 138 px of
        height to the menu bar, title bar and Dock, and the window manager
        cut the window down to what fitted. A reading that is not credible
        (0, or more than the screen) gets a screen-sized geometry instead."""
        try:
            self.state("zoomed")
            if self.state() == "zoomed":
                return
        except tk.TclError:
            pass
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        try:
            max_w, max_h = (int(v) for v in self.wm_maxsize())
        except (tk.TclError, TypeError, ValueError):
            max_w = max_h = 0
        if 0 < max_w <= sw and 0 < max_h <= sh:
            self.geometry(f"{max_w}x{max_h}+0+0")
        else:
            self.geometry(f"{sw - 40}x{sh - 80}+20+20")

    def _fit_column(self, event=None):
        """Keep the text to TEXT_MAX_PX: the rest of the widget's width is
        right margin."""
        width = self.text.winfo_width()
        rmargin = max(0, width - 2 * self.TEXT_PADX - self.TEXT_MAX_PX)
        self.text.tag_configure("col", rmargin=rmargin)

    # ---- contents list ------------------------------------------------------

    def scroll_to(self, key):
        """Put a section's heading at the top of the view (Help > User
        Guide opens here for the tab the user is on)."""
        item = self._toc_items.get(key)
        if item is None:
            return False
        self._go_to_item(item)
        return True

    def _go_to_item(self, item):
        mark = next((m for m, it in self._marks if it == item), None)
        if mark is None:
            return
        if mark == self._fold_head or self._in_fold(mark):
            self.set_fold(True)
        self._current_item = item
        self._select_toc(item)
        self.text.yview(mark)
        self._sync_toc()

    def _on_toc_click(self, event=None):
        """The user clicked (or pressed Return on) a contents entry."""
        item = None
        if event is not None and getattr(event, "type", None) == tk.EventType.ButtonRelease:
            item = self.toc.identify_row(event.y) or None
        if item is None:
            sel = self.toc.selection()
            item = sel[0] if sel else None
        if item:
            self._go_to_item(item)

    def _select_toc(self, item):
        self.toc.selection_set(item)
        self.toc.see(item)

    def _on_text_scroll(self, first, last):
        self._scrollbar.set(first, last)
        self._sync_toc()

    def _heading_at_top(self):
        """The contents item for what is being read: the last heading at or
        above the top of the view (a heading a line below the top still
        belongs to the part before it). At the very bottom, where a short
        last section can never reach the top, the selected item stays while
        its heading is on screen."""
        if not self._marks:
            return None
        top = self.text.index("@0,0")
        current = None
        for mark, item in self._marks:
            if self.text.compare(mark, "<=", top):
                current = item
            else:
                break
        if current is None:
            current = self._marks[0][1]         # the title and introduction: the first section's entry
        try:
            at_bottom = self.text.yview()[1] >= 0.999
        except tk.TclError:
            at_bottom = False
        if at_bottom and self._current_item is not None:
            cur_mark = next((m for m, it in self._marks if it == self._current_item), None)
            if cur_mark is not None and self.text.bbox(cur_mark) is not None:
                return self._current_item
        return current

    def _sync_toc(self):
        item = self._heading_at_top()
        if item is None or item == self._current_item:
            return
        self._current_item = item
        self._select_toc(item)

    # ---- search ----------------------------------------------------------

    def focus_search(self):
        """Put the cursor in the search box (Ctrl+F / Cmd+F, and the
        Welcome window's "Search the User Guide" link)."""
        self.search_entry.focus_set()
        self.search_entry.select_range(0, "end")
        return "break"

    def _search_update(self):
        """The search text changed: find every match afresh, highlight
        them all and go to the first. Case-insensitive, plain text."""
        query = self._search_var.get().strip()
        self.text.tag_remove("search_hit", "1.0", "end")
        self.text.tag_remove("search_current", "1.0", "end")
        self._search_hits, self._search_at = [], -1
        if not query:
            self._search_status_var.set("")
            return
        count = tk.IntVar()
        pos = "1.0"
        while True:
            pos = self.text.search(query, pos, stopindex="end", nocase=True, count=count, elide=True)
            if not pos or count.get() == 0:
                break
            end = f"{pos}+{count.get()}c"
            self._search_hits.append((pos, end))
            self.text.tag_add("search_hit", pos, end)
            pos = end
        if not self._search_hits:
            self._search_status_var.set(_("No matches"))
            return
        self._search_go(0)

    def search_next(self):
        """Go to the next match; after the last one, back to the first."""
        if not self._search_hits:
            if self._search_var.get().strip():
                self._search_update()
            return "break"
        self._search_go((self._search_at + 1) % len(self._search_hits))
        return "break"

    def _search_go(self, index):
        self.text.tag_remove("search_current", "1.0", "end")
        self._search_at = index
        start, end = self._search_hits[index]
        if self._in_fold(start):
            self.set_fold(True)
        self.text.tag_add("search_current", start, end)
        self.text.see(start)
        self._search_status_var.set(_("{n} of {total}").format(
            n=index + 1, total=len(self._search_hits)))

    # ---- the text's building blocks ------------------------------------------

    def _section(self, key, title):
        """A section heading: a top entry in the contents list."""
        mark = f"sec_{key}"
        self.text.mark_set(mark, "end-1c")
        self.text.mark_gravity(mark, "left")
        self.text.insert("end", title + "\n", ("h1", "col"))
        item = self.toc.insert("", "end", text=title, open=True)
        self._marks.append((mark, item))
        self._toc_items[key] = item
        self._cur_section_item = item

    def _sub(self, title):
        """A sub-section heading: a contents entry under the current section."""
        mark = f"sub_{len(self._marks)}"
        self.text.mark_set(mark, "end-1c")
        self.text.mark_gravity(mark, "left")
        self.text.insert("end", title + "\n", ("h2", "col"))
        parent = self._cur_section_item if self._cur_section_item is not None else ""
        item = self.toc.insert(parent, "end", text=title)
        self._marks.append((mark, item))

    # A part of the guide folded away until its heading is clicked (Matt,
    # 2026-11-11: "make the harmonics section click-to-expand"). One fold;
    # the text is all there, hidden with the tag's elide, so the search
    # finds it (and opens it) and the contents list reaches into it.
    FOLD_CLOSED, FOLD_OPEN = "\u25b8", "\u25be"

    def _fold_start(self, title):
        """A sub-heading that opens and closes everything up to _fold_stop."""
        mark = f"sub_{len(self._marks)}"
        self.text.mark_set(mark, "end-1c")
        self.text.mark_gravity(mark, "left")
        self.text.insert("end", f"{self.FOLD_CLOSED} {title}\n", ("h2", "col", "fold_head"))
        parent = self._cur_section_item if self._cur_section_item is not None else ""
        item = self.toc.insert(parent, "end", text=title)
        self._marks.append((mark, item))
        self.text.mark_set("fold_start", "end-1c")
        self.text.mark_gravity("fold_start", "left")
        self._fold_head = mark

    def _fold_stop(self):
        self.text.mark_set("fold_end", "end-1c")
        self.text.mark_gravity("fold_end", "left")
        self.text.tag_add("fold_body", "fold_start", "fold_end")
        self.text.tag_configure("fold_body", elide=True)
        self.fold_open = False
        self.text.tag_bind("fold_head", "<Button-1>", lambda e: self.set_fold(not self.fold_open))
        self.text.tag_bind("fold_head", "<Enter>", lambda e: self.text.config(cursor="hand2"))
        self.text.tag_bind("fold_head", "<Leave>", lambda e: self.text.config(cursor=""))

    def _in_fold(self, index):
        if self._fold_head is None:
            return False
        return self.text.compare(index, ">=", "fold_start") and self.text.compare(index, "<", "fold_end")

    def set_fold(self, open_):
        """Open or close the folded part; the heading's marker follows."""
        if self._fold_head is None or open_ == self.fold_open:
            return
        self.fold_open = open_
        self.text.tag_configure("fold_body", elide=not open_)
        state = str(self.text.cget("state"))
        self.text.config(state="normal")
        self.text.delete(self._fold_head, f"{self._fold_head}+1c")
        self.text.insert(self._fold_head, self.FOLD_OPEN if open_ else self.FOLD_CLOSED,
                         ("h2", "col", "fold_head"))
        self.text.config(state=state)

    def _title(self, text):
        """The guide's own title: a heading with no contents entry."""
        self.text.insert("end", text + "\n", ("h1", "col"))

    def _body(self, text):
        self.text.insert("end", text + "\n", ("body", "col"))

    def _bullet(self, text):
        self.text.insert("end", "  • " + text + "\n", ("bullet", "col"))

    def _steps(self, items):
        """Numbered steps; the numbers are added here, so each step is one
        plain translatable string."""
        for n, text in enumerate(items, 1):
            self.text.insert("end", f"  {n}. {text}\n", ("step", "col"))

    def _blank(self):
        self.text.insert("end", "\n", ("col",))

    def _link(self, display_text, url):
        import webbrowser
        tag = f"link_{id(url)}"
        self.text.tag_configure(tag, font=("Helvetica", 11, "underline"),
                                foreground="#0066CC", lmargin1=20)
        self.text.insert("end", "  " + display_text + "\n", (tag, "col"))
        self.text.tag_bind(tag, "<Button-1>", lambda e: webbrowser.open(url))
        self.text.tag_bind(tag, "<Enter>",
                           lambda e: self.text.config(cursor="hand2"))
        self.text.tag_bind(tag, "<Leave>",
                           lambda e: self.text.config(cursor=""))

    # ---- figures ---------------------------------------------------------------

    FIGURE_HEIGHTS = {"pad_discs": 230, "nest_zone": 250, "adhesive_rings": 210, "die_ring": 260,
                      "kerf": 190, "crescent": 230, "three_ways": 170}

    def _figure(self, name):
        """A drawn figure, from the app's own geometry, in a canvas of its
        own inside the text. Fixed size, so it is drawn at once with no
        layout to wait for."""
        height = self.FIGURE_HEIGHTS[name]
        canvas = tk.Canvas(self.text, width=self.FIGURE_W, height=height, bg=self.FIGURE_BG,
                           highlightthickness=1, highlightbackground="#BBBBBB")
        self.text.window_create("end", window=canvas, padx=20, pady=8)
        self.text.insert("end", "\n", ("col",))
        getattr(self, f"_fig_{name}")(canvas, self.FIGURE_W, height)
        self._figures.append((name, canvas))
        return canvas

    @staticmethod
    def _figure_settings(**overrides):
        settings = copy.deepcopy(DEFAULT_SETTINGS)
        settings.update(overrides)
        return settings

    def _fig_pad_discs(self, cv, W, H):
        """One pad's felt, card and leather at true relative size: the
        leather with its darts, locator marks and number."""
        settings = self._figure_settings(locator_marks_enabled=True)
        pad = 14.0
        host = _FigureCanvasHost(cv)
        diams = [(m, get_disc_diameter(pad, m, settings)) for m in ("felt", "card", "leather")]
        gap_mm = 6.0
        total_w = sum(d for _m, d in diams) + gap_mm * 2
        scale = min((W - 60) / total_w, (H - 70) / max(d for _m, d in diams))
        cy = (H - 30) / 2
        x_mm = -total_w / 2
        for mat, d in diams:
            r_mm = d / 2
            cx = W / 2 + (x_mm + r_mm) * scale
            color = PadPreviewWindow.COLORS[mat]
            if mat == "leather":
                host._draw_leather(cx, cy, pad, settings, r_mm * scale, color, scale)
                cv.create_text(cx, cy - r_mm * scale * 0.45, text=f"{pad:g}", fill=color,
                               font=("Helvetica", max(8, int(r_mm * scale * 0.28)), "bold"))
            else:
                cv.create_oval(cx - r_mm * scale, cy - r_mm * scale, cx + r_mm * scale, cy + r_mm * scale,
                               outline=color, width=2, fill="")
            cv.create_text(cx, cy + r_mm * scale + 16,
                           text=_("{label}\n{mm:.1f} mm").format(label=PadPreviewWindow.LABELS[mat], mm=d),
                           fill=color, justify="center", font=("Helvetica", 9))
            x_mm += d + gap_mm
        cv.create_text(W / 2, 6, anchor="n", width=W - 16, justify="center",
                       text=_("One {size} mm pad: felt, card and leather at true relative size").format(size=f"{pad:g}"),
                       fill=self.FIGURE_INK, font=("Helvetica", 9))

    def _fig_nest_zone(self, cv, W, H):
        """A small nest on a traced scrap, with one labeled-zone block."""
        from svg_engine import nest_pads_with_zones
        # A corner Edge Bias, so the block and the loose pads are seen starting
        # from the same corner (they follow the bias together since 2026-11-11;
        # before, the block was always top-left and the default "center" put
        # the pads in the middle, which Matt read as a broken nest).
        settings = self._figure_settings(zone_labels_enabled=True, edge_bias="nw")
        polygon = [(0.0, 0.0), (140.0, 0.0), (140.0, 60.0), (95.0, 100.0), (0.0, 95.0)]
        pads = [{'size': 9.0, 'qty': 6}, {'size': 24.0, 'qty': 3}, {'size': 30.0, 'qty': 2}]
        placed, zones = nest_pads_with_zones(pads, 'felt', 140.0, 100.0, settings, polygon=polygon)
        margin = 24
        scale = min((W - 2 * margin) / 140.0, (H - 2 * margin) / 100.0)
        ox = (W - 140.0 * scale) / 2
        oy = (H - 100.0 * scale) / 2
        pts = [c for x, y in polygon for c in (ox + x * scale, oy + y * scale)]
        cv.create_polygon(pts, fill=self.FIGURE_SHEET, outline="#888888", width=2)
        color = _PREVIEW_COLORS['felt']
        for zone in zones:
            x0, y0 = ox + zone['x'] * scale, oy + zone['y'] * scale
            cv.create_rectangle(x0, y0, x0 + zone['w'] * scale, y0 + zone['h'] * scale,
                                outline=self.FIGURE_ZONE, width=1, dash=(4, 2))
            # The label where the engraving puts it: centred in the strip
            # inside the top border, above the discs (_render_svg_zones).
            cv.create_text(x0 + zone['w'] * scale / 2, y0 + (zone['border'] + zone['font'] / 2) * scale,
                           text=str(zone.get('label', '')), fill=self.FIGURE_ZONE,
                           font=("Helvetica", max(7, int(zone['font'] * scale)), "bold"))
        for size, cx, cy, r in placed:
            x, y, rp = ox + cx * scale, oy + cy * scale, r * scale
            cv.create_oval(x - rp, y - rp, x + rp, y + rp, outline=color, width=1, fill="#DCE8F5")
            if rp > 11:
                cv.create_text(x, y, text=f"{size:g}", fill=color, font=("Helvetica", 8))
        cv.create_text(W / 2, H - 6, anchor="s", width=W - 16, justify="center",
                       fill=self.FIGURE_INK, font=("Helvetica", 9),
                       text=_("A traced scrap: the 9 mm pads in their labeled block, the rest nested as usual"))

    def _fig_adhesive_rings(self, cv, W, H):
        """Two adhesive rings: the number in the hole, and under a small ring."""
        settings = self._figure_settings(adhesive_ring_enabled=True)
        host = _FigureCanvasHost(cv)
        color = PadPreviewWindow.COLORS['adhesive_ring']
        sizes = (20.0, 11.0)
        geoms = [adhesive_ring_geometry(s, settings) for s in sizes]
        widest = max(g["outer_r"] for g in geoms) * 2
        tallest = max(g["box"][3] for g in geoms) * 2
        scale = min((W - 80) / (widest * 2 + 20), (H - 60) / tallest)
        for i, (size, geom) in enumerate(zip(sizes, geoms)):
            cx = W * (0.3 + 0.4 * i)
            cy = (H - 26) / 2 - 6
            host._draw_adhesive_ring(cx, cy, size, settings, scale, color)
            caption = _("{size} mm pad: number in the hole") if geom["placement"] == "inside" \
                else _("{size} mm pad: number under the ring")
            cv.create_text(cx, H - 6, anchor="s", width=W * 0.42, justify="center",
                           text=caption.format(size=f"{size:g}"), fill=color, font=("Helvetica", 9))

    def _fig_die_ring(self, cv, W, H):
        """A die insert for a 30 mm pad beside a large holder ring, with the
        engravings the app puts on them: the size and NOY, and DESIGNED BY
        PHIL NOY (the credit; see the Tooling tab)."""
        from gcode_engine import get_text_strokes, get_text_strokes_arc
        from svg_engine import HOLDER_OUTER_R, HOLDER_LARGE_INNER_R
        pad = 30.0
        die_r = get_disc_diameter(pad, 'die_ring', self._figure_settings()) / 2
        die_inner = pad / 2
        scale = min((H - 60) / (2 * HOLDER_OUTER_R), (W - 60) / (2 * die_r + 2 * HOLDER_OUTER_R + 14))
        cy = (H - 24) / 2
        cx_die = 30 + die_r * scale
        cx_holder = W - 30 - HOLDER_OUTER_R * scale
        ink = self.FIGURE_INK
        engraving = "#00A000"

        def ring(cx, r_out, r_in):
            cv.create_oval(cx - r_out * scale, cy - r_out * scale, cx + r_out * scale, cy + r_out * scale,
                           outline="#CC0000", width=2, fill="#EFEFEF")
            cv.create_oval(cx - r_in * scale, cy - r_in * scale, cx + r_in * scale, cy + r_in * scale,
                           outline="#0000CC", width=2, fill=self.FIGURE_BG)

        def strokes(cx, stroke_list):
            # Y-up mm (the engraving helpers' frame) to the canvas's Y-down pixels.
            for stroke in stroke_list:
                pts = [c for x, y in stroke for c in (cx + x * scale, cy - y * scale)]
                if len(pts) >= 4:
                    cv.create_line(*pts, fill=engraving, width=1)

        # The die: its size at the top of the ring, NOY arced at the bottom.
        ring(cx_die, die_r, die_inner)
        ring_w = die_r - die_inner
        font = min(3.5, ring_w * 0.9)
        mid_r = (die_r + die_inner) / 2
        strokes(cx_die, get_text_strokes(f"{pad:g}", font, 0.0, mid_r))
        strokes(cx_die, get_text_strokes_arc("NOY", font, 0.0, 0.0, mid_r, -math.pi / 2.0, side='bottom'))
        cv.create_text(cx_die, H - 6, anchor="s", width=2 * die_r * scale + 50, justify="center",
                       text=_("Die insert, {size} mm pad").format(size=f"{pad:g}"),
                       fill=ink, font=("Helvetica", 9))
        # The holder's retaining ring (large): the credit arced along the top.
        ring(cx_holder, HOLDER_OUTER_R, HOLDER_LARGE_INNER_R)
        hw = HOLDER_OUTER_R - HOLDER_LARGE_INNER_R
        hfont = min(3.5, hw * 0.9)
        strokes(cx_holder, get_text_strokes_arc("DESIGNED BY PHIL NOY", hfont, 0.0, 0.0,
                                                (HOLDER_OUTER_R + HOLDER_LARGE_INNER_R) / 2,
                                                math.pi / 2.0, side='top'))
        cv.create_text(cx_holder, H - 6, anchor="s", width=2 * HOLDER_OUTER_R * scale + 50, justify="center",
                       text=_("Holder retaining ring"), fill=ink, font=("Helvetica", 9))

    def _fig_kerf(self, cv, W, H):
        """A kerf test disc in the hole it came out of, the gap drawn wide so
        it can be seen: kerf = hole minus disc."""
        from svg_engine import KERF_TEST_DIAMETERS
        hole_d = KERF_TEST_DIAMETERS[1]
        shown_gap = 1.0                     # mm, exaggerated for the picture
        scale = (H - 70) / hole_d
        cx, cy = W * 0.26, (H - 20) / 2
        rh = hole_d / 2 * scale
        rd = (hole_d / 2 - shown_gap) * scale
        cv.create_rectangle(cx - rh - 30, cy - rh - 14, cx + rh + 30, cy + rh + 14,
                            fill=self.FIGURE_SHEET, outline="#888888")
        cv.create_oval(cx - rh, cy - rh, cx + rh, cy + rh, fill="#FFFFFF", outline="#CC0000", width=1)
        cv.create_oval(cx - rd, cy - rd, cx + rd, cy + rd, fill="#EFEFEF", outline="#CC0000", width=1)
        # The gap, marked on the right of the disc; the notes wrap to the room left.
        tx, tw = W * 0.52, W * 0.46
        cv.create_line(cx + rd, cy, cx + rh, cy, fill=self.FIGURE_INK, width=2)
        cv.create_line(cx + rh + 4, cy, tx - 6, cy, fill=self.FIGURE_INK, width=1)
        ink = self.FIGURE_INK
        cv.create_text(tx, cy - 28, anchor="w", fill=ink, font=("Helvetica", 9), width=tw,
                       text=_("hole: measure its diameter"))
        cv.create_text(tx, cy - 10, anchor="w", fill=ink, font=("Helvetica", 9), width=tw,
                       text=_("disc: measure its diameter"))
        cv.create_text(tx, cy + 10, anchor="w", fill=ink, font=("Helvetica", 9, "bold"), width=tw,
                       text=_("kerf = hole − disc"))
        cv.create_text(tx, cy + 30, anchor="w", fill=self.FIGURE_FAINT, font=("Helvetica", 8), width=tw,
                       text=_("(the gap is drawn wider than it really is)"))
        cv.create_text(cx, H - 6, anchor="s", width=2 * (rh + 40), justify="center",
                       text=_("{d} mm circle of the Kerf Test").format(d=f"{hole_d:g}"),
                       fill=ink, font=("Helvetica", 9))

    def _fig_crescent(self, cv, W, H):
        """A tonehole crescent in its tonehole, mouthpiece side up."""
        from svg_engine import tonehole_crescent_geometry, tonehole_crescent_outline
        tonehole, percent = 43.0, 20.0
        geom = tonehole_crescent_geometry(tonehole, percent)
        outer, inner = tonehole_crescent_outline(geom)
        R = geom['radius']
        scale = (H - 70) / (2 * R)
        cx, cy = W / 2, 28 + R * scale
        cv.create_oval(cx - R * scale, cy - R * scale, cx + R * scale, cy + R * scale,
                       outline="#888888", width=2, fill="#FFFFFF")
        outline = outer + list(reversed(inner))[1:-1]
        pts = [c for x, y in outline for c in (cx + x * scale, cy + y * scale)]
        cv.create_polygon(pts, fill=self.FIGURE_CORK, outline="#6B4A1E", width=2)
        cv.create_text(cx, 12, text=_("mouthpiece side"), fill=self.FIGURE_INK, font=("Helvetica", 9))
        cv.create_text(cx, H - 6, anchor="s", width=W - 16, justify="center", fill=self.FIGURE_INK,
                       font=("Helvetica", 9),
                       text=_("{d} mm tonehole, {p}% of its area covered").format(d=f"{tonehole:g}", p=f"{percent:g}"))

    def _fig_three_ways(self, cv, W, H):
        """The three ways a job gets to the laser, one lane each."""
        rows = [
            (_("Generate G-code"), _("SD card"), _("Laser")),
            (_("Generate SVG"), _("LightBurn"), _("Laser")),
            (_("Frame & Cut"), _("USB cable"), _("Laser")),
        ]
        fills = ("#E3EEF9", "#FDF1DC", "#E6F4E6")
        box_w, box_h, gap = 150, 34, 26
        xs = (12, 12 + box_w + gap, 12 + 2 * (box_w + gap))
        for i, (row, fill) in enumerate(zip(rows, fills)):
            y = 12 + i * (box_h + 14)
            for j, label in enumerate(row):
                x = xs[j]
                cv.create_rectangle(x, y, x + box_w, y + box_h, fill=fill, outline="#555555")
                # Wrapped inside the box: a translated label can be long.
                cv.create_text(x + box_w / 2, y + box_h / 2, text=label, fill=self.FIGURE_INK,
                               font=("Helvetica", 9), width=box_w - 6, justify="center")
                if j < 2:
                    cv.create_line(x + box_w + 4, y + box_h / 2, x + box_w + gap - 4, y + box_h / 2,
                                   fill=self.FIGURE_INK, width=2, arrow="last")
        cv.create_text(W / 2, H - 6, anchor="s", width=W - 16, justify="center",
                       text=_("Three ways to the laser."),
                       fill=self.FIGURE_INK, font=("Helvetica", 9))

    # ---- the content -----------------------------------------------------------

    def _insert_content(self):
        self._title(_("Stohrer Sax Shop Companion"))
        self._body(_("A tool for saxophone technicians: generate laser-cutting templates, "
                    "record key heights, look up serial numbers, reference screw specs, "
                    "tune, and analyze tone."))
        # The credit at the top of the guide (Matt, 2026-11-11: "put a shoutout
        # to phil and his method at the top of the user guide, and link to his
        # website"). An addition: the Tooling section's credit stays as it is.
        self._body(_("The pad-making method this app is built around is Phil Noy's, which he "
                     "shares freely. Thank you, Phil."))
        self._link(_("noysaxophonesupplies.com"), "https://www.noysaxophonesupplies.com")
        # The pad-making guide, at the top since 2026-11-11 (it was a last
        # section of its own, "Learn to make pads"; Matt: "put a link
        # somewhere at the top").
        self._body(_("To make pads, read the guide:"))
        self._link(_("How to make saxophone pads"),
                   "https://www.stohrermusic.com/articles/how-to-make-saxophone-pads/")
        self._intro()
        self._blank()
        self._section_pad_maker()
        self._section_laser()
        self._section_tooling()
        self._section_key_heights()
        self._section_serial_lookup()
        self._section_screw_specs()
        self._section_tuner()
        self._section_toner()
        self._section_sharing()

    def _intro(self):
        # Optional: one line for the top of the full guide (in _insert_content),
        # not one of the nine sections. Drop it if the window says this itself.
        self._body(_("Type in the search box (Ctrl+F, or Cmd+F on a Mac) to find a word, "
                  "or click a heading in the contents list on the right. "
                  "Help > Welcome has a link to the Quick Start Guide, the basics on one page."))

    def _section_pad_maker(self):
        self._section("pad_generator", _("Pad Maker"))
        self._body(_("Makes the laser files for the felt, card and leather of each pad "
                  "from a list of pad sizes."))

        self._sub(_("Start here"))
        self._steps([
            _("Type the pad sizes, one per line, as size x quantity: 42.0 x 3."),
            _("Tick the materials to cut."),
            _("Enter the sheet size and a base filename."),
            _("Click Generate G-code, Generate SVG or Frame & Cut (see Getting a job to the laser)."),
        ])
        self._figure("pad_discs")
        self._bullet(_("You get one file per material, named with the base name and the "
                    "material: my_job_felt.gcode."))
        self._bullet(_("\"max\" as the quantity (18.0 x max) fills the rest of the sheet with "
                    "that size. Only one \"max\" line per sheet."))
        self._bullet(_("Exact size cuts the disc at the size typed, with no offset. SVG only."))
        self._bullet(_("Sizes and sheet are in the unit set in Sizing Rules (in, mm or cm). "
                    "Files are always in mm."))

        self._sub(_("Sizing Rules"))
        self._body(_("Options > Sizing Rules... decides how big each disc is for the pad "
                  "size you type."))
        self._bullet(_("Felt Diameter Reduction: the felt disc is this much smaller than the pad."))
        self._bullet(_("Card Additional Reduction: the card is smaller again, by this much."))
        self._bullet(_("Leather Wrap Multiplier and Felt Thickness: how far the leather wraps "
                    "over the back. 1.00 is the standard wrap."))
        self._bullet(_("Min. Pad Size for Hole: smaller pads get no center hole."))
        self._bullet(_("Universal uses one set of values for every size. Per Size Range sets "
                    "values for size ranges; a size outside every range uses the universal values."))
        self._bullet(_("Darts: leather pads in the size ranges you set get darts so the leather folds flat "
                    "over the back. Shape, frequency, overwrap and sizes are adjustable. "
                    "A size outside every dart range gets no darts."))
        self._bullet(_("\"Show live pad preview\" draws one pad's felt, card and leather with "
                    "the values in the dialog as you change them."))

        self._sub(_("Sizing presets"))
        self._bullet(_("The top of Sizing Rules saves the whole dialog as a named preset: "
                    "Save Preset, Load, Rename, Delete. The dropdown names the preset in use."))
        self._bullet(_("Apply asks you to save your changes as a preset first, so what is "
                    "applied is always a saved preset."))

        self._sub(_("Center Hole"))
        self._body(_("Pick None, 3.0 mm, 3.5 mm or a custom diameter for rivet or screw "
                  "mounting. Pads under Min. Pad Size for Hole are cut without one."))

        self._sub(_("Sheets and shapes"))
        self._bullet(_("Enter the sheet width and height. For card, \"Fit card to paper\" "
                    "uses letter or A4 instead."))
        self._bullet(_("\"Draw / Capture Shape...\" outlines a skin or scrap: click its corners "
                    "on the grid (points are not snapped). Pads nest inside the shape. "
                    "Unload goes back to the rectangle."))
        self._bullet(_("Edge Bias: click an arrow to pack the pads toward that edge or corner, "
                    "say where a skin has a clean edge. The center dot packs from the middle outward."))
        self._bullet(_("\"Preview before saving\" shows the nest and the material used before "
                    "anything is written. Save Files writes; Adjust goes back. One material at a time."))

        self._sub(_("Scrap Mode"))
        self._body(_("Spreads one pad list over several scraps."))
        self._steps([
            _("Tick Scrap Mode and one material."),
            _("Enter this scrap's size, or draw or capture its shape."),
            _("Generate or Frame & Cut. What fits is cut; the rest waits for the next scrap."),
            _("Between scraps keep the shape, unload it, or Re-capture from camera."),
        ])
        self._bullet(_("The progress window shows what is left, what is on this scrap and what "
                    "is done. A scrap counts once its file is written or its cut has finished."))
        self._bullet(_("You can edit the pad list during a session; it counts from the next scrap."))
        self._bullet(_("Files are named _scrap1, _scrap2 and so on."))

        self._sub(_("Labeled Zones"))
        self._body(_("Small discs look alike once cut. Labeled Zones cuts each size in its own "
                  "block with the size engraved on the border, on the waste. Sizing Rules > "
                  "Lesser-used settings; off by default; set the size range."))
        self._figure("nest_zone")
        self._bullet(_("Zones use more material than a free nest."))
        self._bullet(_("A size that cannot get its block is left uncut and reported, never cut unlabeled."))
        self._bullet(_("A \"max\" size inside the zone range is refused. A \"max\" size outside "
                    "it fills in around the blocks."))
        self._bullet(_("Zones are off in Scrap Mode."))

        self._sub(_("Leather Locator Marks"))
        self._body(_("Engraves where the felt sits on a leather pad (four lines, a circle, or "
                  "both), so a pad with no center hole can be centered by eye. Sizing Rules > "
                  "Lesser-used settings; set the size range."))
        self._bullet(_("The marks belong on the side the felt touches: cut the leather flesh "
                    "(fuzzy) side up when you use them."))
        self._bullet(_("SVG puts them on their own layer ({layer} in Layer Colors) for "
                    "their own power in LightBurn. G-code engraves them at the leather "
                    "engraving settings.").format(
                        layer=_("{material} locator marks").format(material=_("Leather"))))

        self._sub(_("Adhesive Rings (beta)"))
        self._body(_("A fifth material: a ring of double-sided tape per pad, for sticking a "
                  "silicone wrap to the back of the card. Switch it on in Sizing Rules > "
                  "Lesser-used settings, then tick Adhesive rings."))
        self._figure("adhesive_rings")
        self._bullet(_("The outer edge is the card disc; the ring is as wide as the leather wrap "
                    "for that size. Wrap clearance (mm, plus or minus) makes it wider or narrower."))
        self._bullet(_("The pad size is cut as single-line digits in the hole, on the piece that "
                    "drops out, or just under the ring when the hole is too small."))
        self._bullet(_("Pads of 9.5 mm and under are solid discs, with the number under them."))
        self._bullet(_("SVG: one layer for ring and number. G-code cuts the number, the inside, "
                    "then the outside, at the ring's cut settings."))
        self._bullet(_("The laser defaults are a starting point, not tested values. Dial them in "
                    "on your tape with Tooling > Speed & Power Test."))

        self._sub(_("Engraving"))
        self._body(_("Each disc is engraved with its pad size. Font sizes and placement are in "
                  "Sizing Rules, Universal or Per Size Range."))
        self._bullet(_("out: from the cut edge. mid: halfway between hole and edge. int: from "
                    "the center hole."))
        self._bullet(_("Darted leather: measure \"out\" from the dart valley puts the number just "
                    "inside the darts. A dart tip always sits over the number."))
        self._bullet(_("A number too big for its disc moves toward the center, then shrinks. If "
                    "the font is over 80 percent of the disc radius it is left off, with a "
                    "warning first."))

        self._sub(_("G-code Settings"))
        self._body(_("Options > G-code Settings... holds speed, power and passes for each "
                  "material and each step: engraving, center hole, outer cut."))
        self._bullet(_("Engraving mode: Line (single stroke) or Filled. Fill density and overscan "
                    "apply to Filled."))
        self._bullet(_("Kerf width: enter the full measured kerf (LightBurn asks for half). When "
                       "outputting G-code or using Frame & Cut, the app compensates every cut, in "
                       "the appropriate direction, automatically. Measure it with Tooling > Kerf "
                       "Test."))
        self._bullet(_("Air assist per step, return-to-home speed, and Cut grouping: by layer "
                    "(all engravings, all holes, all cuts) or by pad."))
        self._bullet(_("Each material has its own presets: Load, Save, Rename, Delete."))
        self._bullet(_("Test speeds and powers on the material you will really cut."))

        self._sub(_("Layer Colors"))
        self._body(_("Options > Layer Colors... sets the LightBurn color layer for each step. "
                  "SVG only. Lesser-used settings > Export Settings has a compatibility mode "
                  "(unitless SVG) for older software."))

        self._sub(_("Job History"))
        self._body(_("File > Job History... lists the last 300 jobs that were written or sent "
                  "to the laser, newest first."))
        self._bullet(_("\"Load into Pad Maker\" puts a job's pad list, materials, sheet, center "
                    "hole and filename back in the form. Sizing rules, G-code settings and "
                    "shapes are left as they are."))
        self._bullet(_("A stopped or errored laser run is marked *. Scrap Mode logs each scrap."))
        self._bullet(_("Tooling jobs are listed by tool and cannot be loaded."))

        self._sub(_("Pad presets"))
        self._bullet(_("Save as Preset keeps the pad list in a library; pick it from the dropdown "
                    "to load it. Delete Preset removes it."))
        self._bullet(_("File > Import Matt's Pad Sets downloads Matt's sets from stohrermusic.com."))

    def _section_laser(self):
        self._section("laser", _("Getting a job to the laser"))

        self._sub(_("Three ways"))
        self._figure("three_ways")
        self._bullet(_("SD card: save the G-code to the card and run it on the laser."))
        self._bullet(_("LightBurn: Generate SVG and set speed and power there."))
        self._bullet(_("Frame & Cut: the app sends it over USB."))

        self._sub(_("G-code or SVG?"))
        self._body(_("A G-code file carries its own speeds and powers, the ones set in this app. "
                  "Send it to the laser as it is, on an SD card or with Frame & Cut."))
        self._body(_("LightBurn ignores those values when it imports G-code, so for LightBurn use "
                     "Generate SVG and set speed and power there. You can set a fixed layer color "
                     "for your SVG outputs, and then set LightBurn to always treat that color as "
                     "the same type of layer, saving you time."))

        self._sub(_("Machine integration"))
        self._steps([
            _("File > Feature Set..., under Experimental / In Progress: tick \"Direct machine "
              "control (GRBL)\". Restart."),
            _("Connect the laser by USB, before or after starting the app. Frame & Cut "
              "appears within seconds."),
        ])
        self._bullet(_("Options > Machine (Pad Maker): Home Laser, Test Connection, Clear Errors, "
                    "Reset Falcon, Camera Calibration, Camera-Polygon Inset Margin, Framing "
                    "Power (dim, per material)."))

        self._sub(_("Frame & Cut"))
        self._steps([
            _("Click Frame & Cut, check the preview, Continue →."),
            _("Jog to the material's visible bottom-left corner, even on a tilted scrap."),
            _("Click Start Frame →. It asks \"Is the lid closed?\": framing turns the laser "
              "on, low power is still power, and an open lid alarms the machine."),
            _("It traces the outline at low power; jog between passes if needed."),
            _("Click \"Looks Good — Cut!\"."),
        ])
        self._bullet(_("Framing is a live laser: keep your hands out of the bed."))
        self._bullet(_("Stop is a soft reset. No calibration needed."))

        self._sub(_("Camera"))
        # Matt's own words (2026-11-11), lightly tidied.
        self._body(_("Ultimate workflow for GRBL lasers: place material on the bed, capture its "
                     "shape with the camera, auto-locate the laser, cut! You can go from entering "
                     "pad numbers to cutting in under a minute, and the precision lets you make "
                     "the most of your material."))
        self._body(_("Optional. Options > Machine > Camera Calibration..., done once:"))
        self._steps([
            _("Secure basswood, home, jog to mid-bed, Frame, then Engrave."),
            _("Capture the card untouched first, click Done with references, then move it around "
              "the bed for the rest of the 12 captures (at least 5 moved), so the camera knows "
              "where in space things on the bed are."),
        ])
        self._bullet(_("This unlocks \"Get from camera\", the shape dialog's live overlay, Try "
                    "Auto Locate, Re-capture from camera and \"Live camera preview\"."))
        self._bullet(_("Live camera preview closes while another window needs the camera "
                    "and returns with the next Frame & Cut."))
        self._bullet(_("\"Camera #1 could not be opened; showing camera #2.\": the saved camera is "
                    "gone. Switch camera saves the other one."))

        self._sub(_("SD card"))
        self._body(_("Windows: \"Eject SD card after G-code export\" ejects it after writing."))
        self._bullet(_("For the Falcon, the SD card must be empty except for the one G-code file "
                    "it will run, so put one material's file on it at a time."))

    def _section_tooling(self):
        self._section("tooling", _("Tooling"))
        self._body(_("Laser files for the die holders and die inserts you assemble pads in, "
                  "plus a few shop tools. Pick a tool with its button, fill it in, then "
                  "Generate SVG, Generate G-code or Frame & Cut."))
        self._body(_("The die inserts and die holders follow Phil Noy's pad-making method, which "
                  "he shares freely: every die insert is engraved \"NOY\" and every holder "
                  "ring \"DESIGNED BY PHIL NOY\"."))
        self._link(_("noysaxophonesupplies.com"), "https://www.noysaxophonesupplies.com")
        self._figure("die_ring")

        self._sub(_("Die Inserts"))
        self._body(_("An acrylic ring for each pad size; the pad is assembled in its hole."))
        self._bullet(_("Pads 7.0 to 39.5 mm get a 50 mm ring, 40.0 to 60.0 mm a 70 mm ring."))
        self._bullet(_("Enter sizes (7, 8.5, 25) or ranges (15-30, by Step), or click Full Set "
                    "(S), Full Set (L) or Full Set (All)."))
        self._bullet(_("\"Engrave size on ring\" numbers the ring. \"Engrave size on cutout\" "
                    "numbers the disc that drops out with its real size."))
        self._bullet(_("Dies nest in rows from the sheet's bottom-left corner, largest first."))
        self._bullet(_("The cutout discs make handy pad cup tools. For an exact diameter, use the "
                    "Pad Maker's Exact size (SVG only)."))
        self._bullet(_("Scrap Mode: tick it to spread a set over several sheets: each Generate "
                       "fills one sheet, a window tracks what is left, and files are named "
                       "_scrap1, _scrap2."))

        self._sub(_("Die Holders"))
        self._body(_("A stack of 85 mm acrylic discs cemented together for good: a solid bottom, "
                  "a magnet disc (6.5 mm hole), two or three pin discs (3.5 mm holes) and a "
                  "retaining ring."))
        self._bullet(_("Pick 5 or 6 layers, and Small (50 mm dies), Large (70 mm) or Both, two "
                    "complete holders on one sheet."))
        self._bullet(_("If the pieces don't fit the sheet, it says the smallest sheet that works."))

        self._sub(_("Die Organizer"))
        self._body(_("A stacking rack for dies, 230 x 330 mm. Cut three upper plates and one "
                  "lower, line up the four 1/8 in corner holes on pins and glue the stack; a "
                  "book press clamps it nicely."))
        self._bullet(_("Pick basswood or acrylic. G-code cuts every slot and hole before the "
                       "plate's edge; the sheet must be at least 234 x 334 mm."))
        self._bullet(_("The SVG is the bare drawing, for LightBurn."))

        self._sub(_("Kerf Test"))
        self._body(_("Kerf is the full width the beam removes. Measure it for each material:"))
        self._figure("kerf")
        self._steps([
            _("Pick the material and cut the test: three discs, 10, 20 and 30 mm."),
            _("Pop the discs out and measure each hole and disc with calipers."),
            _("Kerf = hole diameter minus disc diameter."),
            _("Enter that full kerf in G-code Settings (LightBurn asks for half). The app "
              "compensates every cut."),
        ])
        self._bullet(_("It cuts at the material's settings; untick \"Use existing G-code settings "
                    "for this material\" to try others."))

        self._sub(_("Speed & Power Test"))
        self._body(_("Beta. Cuts small discs, each at its own speed, power and passes and engraved "
                  "with an ID; a legend file lists them. Keep the settings of the discs that "
                  "came free cleanly without scorching."))
        self._bullet(_("\"Apply material defaults\" fills the fields from the material."))
        self._bullet(_("Tick Sweep on Speed, Power or Passes to test Stops values from Start to "
                    "End; untick to hold one value."))
        self._bullet(_("\"Also test with air off\" runs half the discs without air."))
        self._bullet(_("Keep the engraving readable: the IDs are the point. A hole diameter above "
                    "0 makes washers, for shim stock."))
        self._bullet(_("G-code only: an SVG would lose each disc's settings."))
        self._bullet(_("Test on the material you will really cut."))

        self._sub(_("Pad Press Spacers"))
        self._body(_("3D-printable 1.75 in square spacers with their thickness on top; put them "
                     "at press corners to set press depth. A half-step set (3.0 to 4.5 mm), a "
                     "quarter-step set (3.25 to 4.25 mm) and a rack. Print with 100% infill."))

        self._sub(_("Tonehole Crescent"))
        self._body(_("A crescent glued into the mouthpiece side of a tonehole covers part of it "
                  "and flattens that note. Enter the tonehole ID and the percent of its open "
                  "area to cover."))
        self._figure("crescent")
        self._bullet(_("Inner edge: \"Match tonehole curve\" is the classic crescent; flatter is "
                    "shorter and wider, rounder longer and thinner. The area stays the same."))
        self._bullet(_("For several, set each up with a quantity and click Add to list. Generate "
                    "makes the list, or the fields when it is empty. The SVG numbers each part "
                    "on the scrap."))
        self._bullet(_("G-code cuts the chosen material (cork by default) half a kerf outside the "
                    "outline, so the part comes out at size."))
        self._bullet(_("The SVG has no kerf compensation; set it in LightBurn. Generate STL to "
                       "output an STL at the shape and thickness entered."))

        self._sub(_("Frame & Cut"))
        self._body(_("Every tool that writes G-code has Frame & Cut beside Generate G-code. It "
                  "sends the same job the file would hold, the same way as on the Pad Maker."))
        self._bullet(_("Jog to the sheet's bottom-left corner."))
        self._bullet(_("Speed & Power Test asks where to save its legend before the laser moves."))
        self._bullet(_("In Scrap Mode each click cuts one sheet; a sheet counts only once its cut "
                    "has finished."))

        self._sub(_("Settings"))
        self._body(_("Options > Settings...: speed, power, passes, kerf and air for acrylic, "
                  "basswood and cork, and the die engraving options (line or filled, font "
                  "sizes, ring placement)."))
        self._bullet(_("Die inserts and holders use the acrylic settings. Defaults suit a Falcon2 "
                    "Pro 40W on 3 mm acrylic and basswood; adjust for your laser."))
        self._bullet(_("The basswood settings also engrave the camera calibration card."))
        self._bullet(_("Cork: the default cut was tested at 1000 mm/min and 30 % on 3/32 in cork "
                    "with a 40 W laser. The kerf is not measured (0) and the engraving is not "
                    "set (0 and 0)."))
        self._bullet(_("While cork's engraving is not set, nothing is engraved on cork, and the "
                    "app says so before it cuts."))

    def _section_key_heights(self):
        self._section("key_heights", _("Key Height Library"))
        self._body(_("Record and compare key heights for different saxophones, in libraries you "
                     "organize. Entries and libraries are exportable, and you can share with "
                     "other SSC users if you want."))
        self._bullet(_("Options > Layout Options... chooses which keys are shown."))
        self._bullet(_("File > Import Matt's Key Heights downloads Matt's sets from stohrermusic.com."))

    def _section_serial_lookup(self):
        self._section("serial_lookup", _("Serial Lookup"))
        self._body(_("Pick the maker and type the serial number to estimate the year it was made."))

    def _section_screw_specs(self):
        self._section("screw_specs", _("Screw Specs"))
        self._body(_("Thread specifications for screws and rods, by saxophone model."))
        self._bullet(_("File > Import Matt's Specs downloads the latest from stohrermusic.com. "
                       "Also exportable if you desire."))

    def _section_tuner(self):
        self._section("tuner", _("Tuner"))
        self._body(_("A 12-wheel chromatic strobe tuner: one wheel per note, one ring per octave. "
                     "It runs while the tab is open."))
        self._bullet(_("Frozen pattern: in tune. Drifting right: sharp. Left: flat. The faster "
                    "the drift, the further off."))
        self._bullet(_("Several wheels move at once: those are the harmonics in your sound."))
        self._bullet(_("The VU meter at the bottom shows the note and how many cents off it is."))

        self._sub(_("Controls"))
        self._bullet(_("DISP: SENS (how loud a sound must be), BRIGHT, FPS (60, 90 or 120)."))
        self._bullet(_("PITCH: A= reference (440 Hz by default), KEY transposition (C, Bb, Eb, F)."))
        self._bullet(_("BIAS: NOTE brightens the played octave's ring; OCT. boosts the strongest ring."))
        self._bullet(_("Options > Settings...: backlight and faceplate colors, frame-rate readout. "
                    "Options > Input Device... picks the microphone."))

        self._sub(_("Microphone"))
        self._bullet(_("A good mic is a must. A USB condenser such as the Audio-Technica AT2020 "
                    "USB works well."))
        self._bullet(_("Built-in laptop mics manage basic tuning but struggle low down and leave "
                    "out harmonics. Bluetooth headset mics do not work: their sample rate is "
                    "too low."))

        self._sub(_("How it works"))
        self._bullet(_("Each wheel turns at the difference between the sound and its note, from "
                    "real analysis of the audio, not a simulation."))
        self._bullet(_("It draws on the graphics card where it can (60 to 120 fps) and falls back "
                    "to a slower drawing mode where it can't."))

    def _section_toner(self):
        self._section("toner", _("Analyzer"))
        self._body(_("The Analyzer listens to you play and charts your intonation: each written "
                  "note as played, in cents. Its harmonic side, behind the HARMONICS switch, "
                  "is experimental and measures the overtones."))

        self._sub(_("An intonation run"))
        # The required fields as the preset form names them, in the user's language.
        needed = ", ".join([_("Horn Type"), _("Make"), _("Model"), _("Player"), _("Mouthpiece")])
        self._steps([
            _("Open the Preset Library (the button under the chart, or File > Preset Library...): "
              "Load a preset, Capture to start on it, or Record first to play now and choose the "
              "preset when you stop. A preset needs {fields}.").format(fields=needed),
            _("Click Capture and play. Each note held long enough is kept, and the chart draws "
              "as you go."),
            _("Click Stop. The run is saved with the preset, dated. A run recorded first asks where "
              "it goes: an existing preset, a new one, or discard; Keep for now holds it until you "
              "decide, and nothing leaves the tab, loads or quits without asking."),
        ])
        self._bullet(_("Unload on the control strip puts the preset away; Capture then opens the "
                    "Preset Library."))
        self._bullet(_("Free mode, the default, takes the notes as you play them. Guided mode walks "
                    "you up a chromatic scale, 5 seconds a note, with a large note name to play."))
        self._bullet(_("The readout counts the frames a held note has given and turns green once "
                    "there are enough to keep, about a second. A note let go sooner is named as "
                    "too short."))
        self._bullet(_("Several takes of one note are one point: their median, with a bar from the "
                    "lowest to the highest."))
        self._bullet(_("The dashed line marks the register change. The SAX selector sets the written "
                    "names and is saved with the preset."))
        self._bullet(_("Key heights: a preset can link to a set in the Key Heights library, or enter "
                    "the horn's heights from the preset form (saved there as a new set). Each run "
                    "keeps a copy of the heights it was recorded with. The chart shows nothing of "
                    "them: name the preset so you know which horn it is."))

        self._sub(_("Compare and export"))
        self._bullet(_("The Library box above the list shows one library's presets, or All Libraries; "
                    "pick All Libraries to compare presets from different libraries."))
        self._bullet(_("In the Preset Library, select two or more rows and click Compare: presets, or "
                    "single sessions of a preset with several, each run in its own colour and marker. "
                    "A preset row is all its sessions together; a session row is that session alone."))
        self._bullet(_("Click a run's line, a point or its legend entry to bring it to the front; the "
                    "others go grey. Click it again, or an empty spot, to let it go."))
        self._bullet(_("Export chart... in the Preset Library writes the chart as a PNG or a PDF."))
        self._bullet(_("Low notes on the left or on the right: Options > Settings... > General."))
        self._bullet(_("The TUNER switch shows a needle meter with the note and an IN TUNE lamp beside "
                    "the chart. The chart stays in view in every arrangement, so you can watch it "
                    "fill in."))

        self._fold_start(_("Harmonics (experimental and SUPER mic-sensitive)"))
        self._body(_("The HARMONICS switch on the control strip shows the harmonic analyzer beside the "
                  "chart: it finds the note you play and measures each harmonic up to the 20th. "
                  "Its presets also need Mic Type and Mic Model."))
        self._body(_("Every reading is the whole chain at once: player, horn, mouthpiece, reed, "
                  "mic and room. So it measures change, not gear. Change one thing at a time "
                  "and compare. Mouthpiece, player and mic routinely outweigh horn-to-horn "
                  "differences; compare your setups with your other setups."))

        self._sub(_("Two ways to use it"))
        self._bullet(_("Live: watch the spectrum while you play to see how air, embouchure and "
                    "voicing change the sound."))
        self._bullet(_("Over time: record a session, change one thing, record another, and "
                    "compare. The thing you changed is the difference: a mouthpiece, a mic, "
                    "a horn, or you on another day."))
        self._bullet(_("Be skeptical of strong conclusions from small samples."))

        self._sub(_("Before you capture"))
        # The required fields as the preset form names them, in the user's language.
        required = ", ".join([_("Make"), _("Model"), _("Player"), _("Mouthpiece"), _("Mic Type"), _("Mic Model")])
        self._bullet(_("You need a preset: File > Presets.... {fields} are required. More fields "
                    "can be switched on in Options > Settings... > Analysis.").format(fields=required))
        self._bullet(_("The SAX selector sets the transposition and is saved with the preset."))
        self._bullet(_("A condenser mic (such as an AT2020 USB) gives the fullest picture. Ribbon "
                    "and dynamic mics dull the upper harmonics; the Analyze tool warns when mic "
                    "types differ."))
        self._bullet(_("Laptop mics are not suitable. Bluetooth headset mics do not work."))

        self._sub(_("Mic placement"))
        self._bullet(_("2 to 3 feet (60 to 90 cm) from the bell, slightly off-axis, in a quiet room. "
                    "Closer than a foot exaggerates the low harmonics."))
        self._bullet(_("Mic position matters more than almost anything: a few inches between "
                    "takes can shift complexity 10 to 20 percent. Use a stand at a fixed "
                    "spot, and note it in the {field} field.").format(field=_("Mic Position")))

        self._sub(_("Capturing"))
        self._steps([
            _("Click Capture and pick or create a preset."),
            _("Play. Hold a note about a second and it is captured; move on to the next."),
            _("Click Stop. A summary shows which notes you covered and where the gaps are."),
        ])
        self._bullet(_("Capture at least 8 different notes; more is better."))
        self._bullet(_("Guided mode walks you up a chromatic scale, 5 seconds per note. "
                    "File > Presets... > Import Audio File... analyzes a WAV."))
        self._bullet(_("Sessions are dated and saved automatically."))

        self._sub(_("Display"))
        self._bullet(_("Spectrum or Bars view, with Linear (true ratios) or dB scale."))
        self._bullet(_("The intonation gauge shows cents flat or sharp; the IN TUNE lamp lights "
                    "within 4 cents."))
        self._bullet(_("Overlay on Spectrum (from the Analyze tool) shows a stored preset as blue "
                    "ghost bars behind the live spectrum, note by note."))

        self._sub(_("Analyze"))
        self._body(_("File > Analyze... shows what differs between presets: not which is better, "
                  "but what changed and where."))
        self._bullet(_("Filter the picker by horn type, player, mouthpiece and mic type."))
        self._bullet(_("One preset: its descriptors and harmonic curve. Two: side by side with a "
                    "Difference chart. Three or more: the spread across the group."))
        self._bullet(_("Horn Average or Per-Note shows whether a difference is everywhere or in "
                    "one register."))
        self._bullet(_("Percentiles rank each preset against all presets of the same sax type."))
        self._bullet(_("Click a legend label or chart line for that preset's details. Back returns "
                    "to the picker."))
        self._bullet(_("The table shows each preset's mic type and rolloff, and notes when mics "
                    "differ."))

        self._sub(_("Descriptors"))
        self._body(_("Most useful as differences between two presets, where mic and room effects "
                  "cancel. Choose which are shown in Options > Settings... > Analysis."))
        self._bullet(_("Warmth (Thin to Warm): the strength of the 2nd harmonic. Not the same as "
                    "dark: a sound can be warm and bright, or thin and dark."))
        self._bullet(_("Some setups, metal mouthpieces played hard especially, make the 2nd "
                    "harmonic louder than the fundamental. Warmth reads these as moderately "
                    "warm; this part of the formula may change."))
        self._bullet(_("Harmonic Complexity (Pure to Complex): how evenly the energy spreads over "
                    "the harmonics."))
        self._bullet(_("Even/Odd Ratio: even against odd harmonics, round against edgy. The most "
                    "consistent across recording conditions. On by default."))
        self._bullet(_("Rolloff Shape (Smooth to Peaked): bumps where harmonics stick out. Off by "
                    "default."))
        self._bullet(_("Evenness (Variable to Even): how alike the character is across the range. "
                    "Needs at least 5 notes."))

        self._sub(_("Character Map"))
        self._body(_("Plots Warmth against Harmonic Complexity, so each preset sits in one of four "
                  "corners: warm and bright, warm and dark, thin and bright, thin and dark. All "
                  "four are real characters."))
        self._bullet(_("Warmth comes from how the chamber and reed couple to the fundamental, "
                    "brightness from the reed's buzz on the tip rail, so they move independently."))
        self._bullet(_("H4-H5 mean and Rolloff are offered as other axes, but they tend to follow Warmth."))
        self._bullet(_("Compare within one sax type: lower horns read warmer and brighter whatever "
                    "the mouthpiece. It warns before mixing types."))

        self._sub(_("Reading the differences"))
        self._bullet(_("H1 to H4: mostly the bore; the least change in player and mouthpiece swaps."))
        self._bullet(_("H7 to H13: where neck and mouthpiece swaps show most."))
        self._bullet(_("H3 to H12 all together: usually mouthpiece or player."))
        self._bullet(_("Two takes of one setup ten minutes apart differed by 1 to 3 percent on "
                    "descriptors and about 0.3 dB/H on rolloff. Smaller differences are probably "
                    "noise."))
        self._bullet(_("These are observed patterns, not laws. Same-player comparisons isolate "
                    "equipment best."))

        self._sub(_("Recording quality"))
        self._body(_("Rolloff, how fast the upper harmonics fade, shows how much your recording "
                  "catches:"))
        self._bullet(_("1.0 to 2.0 dB/H: condenser, close. Full detail."))
        self._bullet(_("2.0 to 2.5 dB/H: good mic, further away or a lively room."))
        self._bullet(_("Over 2.5 dB/H: ribbon, dynamic, built-in or distant mic. The app warns "
                    "after the first few captures."))

        self._sub(_("Presets and data"))
        self._bullet(_("Mutate Preset copies a preset with its fields filled in, so you can change "
                    "one thing and save it as new."))
        self._bullet(_("Sandbox mode (Options > Settings... > Analysis) relaxes the required "
                    "fields for other instruments and experiments. Sandbox presets are marked "
                    "in amber."))
        self._bullet(_("A preset's report shows its sessions over time, with the change from each "
                    "to the next and a per-note breakdown."))
        self._bullet(_("File > Transfer Data exports or imports the preset library."))

        self._sub(_("WAV recording"))
        self._body(_("On by default; the first capture asks for a folder. At the end of a session "
                  "the recording is analyzed again, more strictly than live, and that better "
                  "result is what is stored."))
        self._bullet(_("Turn it off, or delete each WAV after analysis, in Options > Settings... > General."))
        self._bullet(_("Keep the WAVs to listen back, share, or re-analyze with better tools later."))

        self._sub(_("How it works"))
        self._bullet(_("Harmonics are measured in dB against the fundamental, so volume and mic "
                    "gain drop out. Below -60 dB counts as noise."))
        self._bullet(_("Only the frames whose whole analysis window lies inside the note count, "
                    "so the attack and the release are left out."))
        self._bullet(_("The detector checks whether a loud peak is a harmonic of a lower note, "
                    "which keeps it from reading the wrong octave."))
        self._bullet(_("Notes are averaged first, then the notes together, so many captures of one "
                    "register don't outweigh the rest."))
        self._bullet(_("Only raw measurements are stored. Descriptors are worked out fresh each "
                    "time, so improvements apply to old captures too."))
        self._fold_stop()

    def _section_sharing(self):
        self._section("sharing", _("Sharing data"))
        self._body(_("Each data tab can export to a file and import one, to share with colleagues "
                  "or keep a backup:"))
        self._bullet(_("Pad presets: File > Export Pad Presets... and Import Pad Presets...."))
        self._bullet(_("Key heights: File > Export Key Sets... and Import Key Sets...."))
        self._bullet(_("Screw specs: File > Export Screw Specs... and Import Screw Specs...."))
        self._bullet(_("Analyzer presets: File > Transfer Data."))
        self._bullet(_("Everything: File > Import Settings from Folder... copies the settings of "
                    "another installation."))


class QuickStartWindow(tk.Toplevel):
    """The quick start guide: one small page of plain text for someone new
    to the app or not computer-minded, so they don't get the three ways of
    cutting backwards (Matt, 2026-11-09: "maybe some very basic instruction
    would be handy so people don't get things backwards", and "whatever
    guide we write should be SUPER concise").

    Opened from the Welcome window's link only (Help > Welcome, or the
    Welcome a new version shows once), never by itself and not from the
    Help menu (Matt, 2026-11-11). No images, no steps, no scrolling: the window
    is as big as its text. Not modal. Every line is its own translatable
    string, because the wording is Matt's to change. The words are his
    own since 2026-11-11 (he asked for the text and sent it back
    rewritten); only the tab's and the button's names were set to match
    the app.

    show_frame_cut: the direct laser control line is only for a build
    that can drive a laser (main.machine_supported()), and the line
    above the ways counts them.
    """

    WRAP = 460

    @staticmethod
    def lines(show_frame_cut=True):
        """The page, top to bottom: (kind, text), kind 'h' (a heading),
        'p' (a line) or 'link' (the line that opens the User Guide)."""
        page = [
            ('p', _("Welcome. This is a complex, many-featured program.")),
            ('link', _("See Help > User Guide for the full rundown.")),
            ('h', _("Quick start")),
        ]
        if show_frame_cut:
            page.append(('p', _("Pad Maker outputs pads ready to cut, per material. Three ways to do it:")))
        else:
            page.append(('p', _("Pad Maker outputs pads ready to cut, per material. Two ways to do it:")))
        page += [
            ('p', _("SD card: Enter pads, click Generate G-code, save it to the SD card, run it on the laser.")),
            ('p', _("LightBurn: Enter pads, click Generate SVG, save them, open them in LightBurn, "
                    "set speed and power for each color, export to SD or direct laser control.")),
        ]
        if show_frame_cut:
            page.append(('p', _("Direct laser control: Sax Shop Companion controls your laser. Enter your "
                                "pads, click Frame & Cut, frame, cut. Switch it on in File > Feature Set. "
                                "Can be manual, or set it up to be camera-assisted for a very quick and "
                                "automated workflow.")))
        page += [
            ('h', _("Before you cut!")),
            ('p', _("Speeds and powers are in Options > G-code Settings (Pad Maker) and "
                    "Options > Settings (Tooling). Defaults are for a 40 W diode laser.")),
            ('p', _("Test first. Don't start a fire.")),
        ]
        return page

    LINK = "#0066CC"

    def __init__(self, parent, show_frame_cut=True, on_user_guide=None):
        super().__init__(parent)
        self.title(_("Quick Start Guide"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.resizable(False, False)
        self._on_user_guide = on_user_guide

        body = tk.Frame(self, bg=DIALOG_BG)
        body.pack(padx=22, pady=(14, 6))
        self.line_labels = []
        page = self.lines(show_frame_cut)
        self.guide_link = None
        for kind, text in page:
            if kind == 'h':
                label = tk.Label(body, text=text, bg=DIALOG_BG, font=("Helvetica", 12, "bold"),
                                 justify="left", anchor="w", wraplength=self.WRAP)
                label.pack(anchor="w", pady=(10, 2))
            elif kind == 'link' and on_user_guide is not None:
                # The line that names the User Guide opens it (2026-11-09,
                # the guide upgrade; the second line since 2026-11-11).
                label = tk.Label(body, text=text, bg=DIALOG_BG, fg=self.LINK, cursor="hand2",
                                 font=("Helvetica", 10, "underline"), justify="left", anchor="w",
                                 wraplength=self.WRAP)
                label.pack(anchor="w", pady=1)
                label.bind("<Button-1>", lambda e: self._open_user_guide())
                self.guide_link = label
            else:
                label = tk.Label(body, text=text, bg=DIALOG_BG, font=("Helvetica", 10),
                                 justify="left", anchor="w", wraplength=self.WRAP)
                label.pack(anchor="w", pady=1)
            self.line_labels.append(label)
        tk.Button(self, text=_("Close"), width=10, command=self.destroy).pack(pady=(8, 14))

    def _open_user_guide(self):
        if self._on_user_guide is not None:
            self._on_user_guide()


class WelcomeWindow(tk.Toplevel):
    """The Welcome window: shown once for each version on its first launch
    (main.maybe_show_welcome), and any time from Help > Welcome.

    Matt, 2026-11-09: "each install shows a quick welcome, a version
    number, and thank you for installing, be sure to search the user guide
    with any questions, and says click here for the quick start guide."

    A plain, small window in the app's ordinary dialog look, and it holds
    exactly this: the app name and version, the thank-you, two links
    (search the User Guide, the quick start guide) and Close. A first
    version was styled after an old dial-up sign-on screen, with three
    panels that lit up; Matt took that back the same day: "let's have the
    welcome screen be simple. if I have a detailed ui request I will give
    it." So: no styling of its own, no pictures, no animation, no timers.
    Do not add any.

    Not modal: it takes no grab, so the main window can be used with it
    open.
    """

    LINK = "#0066CC"
    WRAP = 380

    def __init__(self, parent, version, on_search_guide=None, on_quick_start=None):
        super().__init__(parent)
        self.title(_("Welcome"))
        self.configure(bg=DIALOG_BG)
        self.transient(parent)
        self.resizable(False, False)

        self._on_search_guide = on_search_guide
        self._on_quick_start = on_quick_start

        self.name_label = tk.Label(self, text=_("Stohrer Sax Shop Companion"), bg=DIALOG_BG,
                                   font=("Helvetica", 14, "bold"), wraplength=self.WRAP)
        self.name_label.pack(padx=28, pady=(18, 2))
        self.version_label = tk.Label(self, text=_("Version {version}").format(version=version),
                                      bg=DIALOG_BG, font=("Helvetica", 9), wraplength=self.WRAP)
        self.version_label.pack()
        self.thanks_label = tk.Label(self, text=_("Thank you for installing."), bg=DIALOG_BG,
                                     font=("Helvetica", 10), wraplength=self.WRAP)
        self.thanks_label.pack(padx=28, pady=(10, 10))
        # The two links in Matt's words and his order (2026-11-11): the quick
        # start first, then the User Guide.
        self.quick_start_link = self._link(_("Click here for the quick start or"), self._quick_start)
        self.guide_link = self._link(_("Be a BOSS and click HERE for the User Guide"), self._search_guide)
        self.close_button = tk.Button(self, text=_("Close"), width=10, command=self.destroy)
        self.close_button.pack(pady=(14, 16))

    def _link(self, text, command):
        """A link in the app's usual link look (as in the About dialog)."""
        label = tk.Label(self, text=text, bg=DIALOG_BG, fg=self.LINK,
                         font=("Helvetica", 10, "underline"), cursor="hand2", wraplength=self.WRAP)
        label.pack(padx=28, pady=2)
        label.bind("<Button-1>", lambda e: command())
        return label

    def _search_guide(self):
        if self._on_search_guide is not None:
            self._on_search_guide()

    def _quick_start(self):
        if self._on_quick_start is not None:
            self._on_quick_start()


class AboutDialog(tk.Toplevel):
    """About dialog with feature summary and contact info."""

    def __init__(self, parent):
        import webbrowser
        super().__init__(parent)
        self.title(_("About"))
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
        # GitHub redirects /releases/latest to the newest release's page.
        latest = tk.Label(self, text=_("Check for a newer version"),
                          bg=DIALOG_BG, fg="#0066CC",
                          font=("Helvetica", 9, "underline"), cursor="hand2")
        latest.pack()
        latest.bind("<Button-1>", lambda e: webbrowser.open(LATEST_RELEASE_URL))

        # Feature summary
        features = _(
            "\u2022  SVG & G-code generation for laser-cut pads\n"
            "\u2022  Key height library with import/export\n"
            "\u2022  Serial number lookup\n"
            "\u2022  OEM screw & rod specifications\n"
            "\u2022  Die insert & die holder tooling\n"
            "\u2022  Chromatic strobe tuner\n"
            "\u2022  Intonation chart and harmonic analyzer"
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
        # Height from the content, not a fixed number: the window doesn't
        # resize, so every added line (like the newer-version link) or a
        # longer translation would otherwise push OK off the bottom.
        self.update_idletasks()
        self.geometry(f"440x{max(520, self.winfo_reqheight())}")
