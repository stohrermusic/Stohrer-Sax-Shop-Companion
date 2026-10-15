"""v2.86 fixes outside the calibration dialog (that one is
test_calibration_dialog).

  - Frame & Cut: "Try Auto Locate", the G92 offset and the polygon
    dialog's machine anchor all name the same corner. Auto Locate used the
    leftmost vertex while the cut used the one nearest the bounding box's
    bottom-left; on a scrap tilted 2 degrees CCW those are 150 mm apart.
  - Auto Locate jogs in machine coordinates ($J=...G53...), so a G92 left
    by the previous Frame & Cut can't shift it.
  - The calibration card clears G92 before it moves.
  - The calibration file follows SAXSHOP_CONFIG_DIR like every other
    settings file (tests could overwrite a user's real calibration).
  - Help > About links to the latest release, and sizes itself so OK
    stays on screen.
  - Tonehole Crescent: no "Mouthpiece side" caption over the preview, and
    the description is just the exact-size/kerf sentence.

Isolated profile (SAXSHOP_CONFIG_DIR). GUI cases self-skip without a
display.
"""
import math
import os
import sys
import tempfile
import traceback
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PROFILE = tempfile.mkdtemp(prefix="ssc-2-86-")
os.environ["SAXSHOP_CONFIG_DIR"] = _PROFILE

from i18n import init_translation  # noqa: E402
init_translation("en")

results = []


def check(name, fn):
    try:
        if fn() is False:
            raise AssertionError("check returned False")
        print(f"  PASS  {name}")
        results.append(True)
    except Exception as e:
        print(f"  FAIL  {name}: {e}")
        traceback.print_exc()
        results.append(False)


def rotated_rect(w, h, deg, cx=200.0, cy=150.0):
    """A w x h rectangle turned `deg` counter-clockwise about (cx, cy),
    machine Y-up, corners in order BL, BR, TR, TL before rotation."""
    a = math.radians(deg)
    pts = []
    for x, y in ((-w / 2, -h / 2), (w / 2, -h / 2), (w / 2, h / 2), (-w / 2, h / 2)):
        pts.append((cx + x * math.cos(a) - y * math.sin(a),
                    cy + x * math.sin(a) + y * math.cos(a)))
    return pts


def main():
    from svg_engine import polygon_anchor_vertex
    print("Engine")
    print("=" * 60)
    scrap = rotated_rect(250, 150, 2.0)

    def anchor_is_the_bottom_left_corner():
        v = polygon_anchor_vertex(scrap)
        assert v == scrap[0], (v, scrap[0])  # the BL corner, turned 2 deg
        leftmost = min(scrap, key=lambda p: (p[0], p[1]))
        assert leftmost == scrap[3], "fixture: leftmost should be the TL corner"
        gap = math.dist(v, leftmost)
        assert gap > 140, gap  # what the old rule cost: ~150 mm
    check("anchor vertex = corner nearest the bbox bottom-left (the old "
          "leftmost rule picks the top-left corner, ~150 mm away)",
          anchor_is_the_bottom_left_corner)

    def axis_aligned_unchanged():
        rect = [(10, 20), (110, 20), (110, 90), (10, 90)]
        assert polygon_anchor_vertex(rect) == (10, 20)
    check("axis-aligned scrap: still its bottom-left corner", axis_aligned_unchanged)

    def polygon_dialog_anchor():
        import ui_dialogs as ui
        fake = types.SimpleNamespace(
            points=[(x / 10.0, y / 10.0) for x, y in scrap], unit="cm",
            _camera_anchor_mm=(0.0, 0.0), _overlay_used=True)
        got = ui.PolygonDrawWindow.get_polygon_lb_machine_mm(fake)
        want = polygon_anchor_vertex(scrap)
        assert math.dist(got, want) < 1e-6, (got, want)
    check("polygon dialog's machine anchor uses the same corner",
          polygon_dialog_anchor)

    def card_clears_g92():
        import gcode_engine
        p = os.path.join(_PROFILE, "card.gcode")
        gcode_engine.generate_calibration_card_gcode(
            p, cols=10, rows=10, offset_x_mm=60, offset_y_mm=60)
        lines = open(p, encoding="utf-8").read().splitlines()
        first_move = next(i for i, ln in enumerate(lines)
                          if ln.startswith(("G0", "G1")) and ("X" in ln or "Y" in ln))
        assert "G92.1" in lines[:first_move], lines[:first_move + 1]
    check("calibration card clears G92 before its first move", card_clears_g92)

    def machine_jog():
        import falcon_sender as fs
        if not fs.HAS_PYSERIAL:
            print("    (pyserial missing; skipped)")
            return

        class W:
            is_open = True

            def __init__(self):
                self.out = []

            def write(self, b):
                self.out.append(b)
        s = fs.FalconSender(port="COM-TEST")
        s._serial = W()
        s.jog(x=123.5, y=45.25, feed=3000, relative=False, machine=True)
        s.jog(x=1, feed=3000, relative=True)
        assert s._serial.out[0] == b"$J=G90G53X123.5Y45.25F3000\n", s._serial.out[0]
        assert s._serial.out[1] == b"$J=G91X1F3000\n", s._serial.out[1]
        import inspect
        import ui_dialogs as ui
        src = inspect.getsource(ui.FalconRunDialog._on_auto_locate_clicked)
        assert "machine=True" in src, "Auto Locate doesn't jog in machine coordinates"
    check("Auto Locate jogs in machine coordinates (G53); relative jogs unchanged",
          machine_jog)

    def calibration_path_isolated():
        import camera_capture
        p = camera_capture.default_calibration_path()
        assert os.path.dirname(p) == os.path.normpath(_PROFILE) or \
            os.path.normcase(os.path.dirname(p)) == os.path.normcase(_PROFILE), p
    check("calibration file follows SAXSHOP_CONFIG_DIR", calibration_path_isolated)

    print()
    print("App")
    print("=" * 60)
    import tkinter as tk
    try:
        root = tk.Tk()
    except tk.TclError as e:
        print(f"Skipping the app cases: no display available ({e})")
        return 0 if all(results) else 1
    root.withdraw()
    try:
        import main as main_mod
        import ui_dialogs as ui
        app = main_mod.PadSVGGeneratorApp(root)
        root.update()

        def auto_locate_matches_g92():
            app._adopt_camera_polygon(list(scrap))
            target = app._custom_polygon_lb_machine
            gx, gy = app._frame_cut_anchor()
            xmin = min(p[0] for p in scrap)
            ymin = min(p[1] for p in scrap)
            # Auto Locate puts the head on `target`; the G92 says the head
            # is at (gx, gy) relative to the outline's bbox. Same point:
            assert abs((target[0] - xmin) - gx) < 1e-6 and \
                abs((target[1] - ymin) - gy) < 1e-6, (target, (gx, gy), (xmin, ymin))
        check("Frame & Cut: Auto Locate target is the vertex the G92 assumes",
              auto_locate_matches_g92)

        def about():
            # The dialog is transient to the withdrawn test root, so Tk
            # never maps it and reports 1x1; record what it asks for.
            set_sizes = []
            real_geometry = ui.AboutDialog.geometry

            def record(self, new=None):
                if new is not None:
                    set_sizes.append(new)
                return real_geometry(self, new)
            ui.AboutDialog.geometry = record
            try:
                dlg = ui.AboutDialog(root)
            finally:
                ui.AboutDialog.geometry = real_geometry
            try:
                dlg.update_idletasks()
                texts = [w.cget("text") for w in dlg.winfo_children()
                         if isinstance(w, tk.Label)]
                assert "Check for a newer version" in texts, texts
                assert ui.LATEST_RELEASE_URL.endswith("/releases/latest")
                assert set_sizes, "About never sized itself"
                h = int(set_sizes[-1].split("+")[0].split("x")[1])
                assert h >= dlg.winfo_reqheight(), (set_sizes, dlg.winfo_reqheight())
            finally:
                dlg.destroy()
        check("About links to the latest release and fits its content", about)

        def crescent():
            app._update_crescent_preview()
            cv = app.crescent_canvas
            texts = [cv.itemcget(i, "text") for i in cv.find_all()
                     if cv.type(i) == "text"]
            assert "Mouthpiece side" not in texts, texts

            def labels(w):
                for c in w.winfo_children():
                    if isinstance(c, tk.Label):
                        yield c.cget("text")
                    yield from labels(c)
            all_text = list(labels(app.root))
            assert not any("A crescent set into the mouthpiece side" in t
                           for t in all_text), "old first sentence still shown"
            assert any(t.startswith("The outer edge is the tonehole ID exactly")
                       for t in all_text), "kerf sentence missing"
        check("Tonehole Crescent: no 'Mouthpiece side' caption; description "
              "is the kerf sentence", crescent)
    finally:
        try:
            root.destroy()
        except tk.TclError:
            pass

    passed = sum(results)
    print()
    print(f"{passed}/{len(results)} passed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
