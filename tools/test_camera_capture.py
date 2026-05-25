"""
Engine-level tests for camera_capture.py — no actual camera required.

Covers:
  - OpenCV-availability gating (module imports without OpenCV)
  - ChArUco board generation produces a valid PNG of expected pixel size
  - ChArUco detection round-trips on the rendered card itself
  - Calibration JSON save/load round-trip
  - Synthetic full pipeline: render board -> 'photograph' it via warp ->
    detect -> calibrate -> undistort -> verify low reprojection error
  - Scrap contour detection on a synthetic image with a known shape
  - pixels_to_mm honors the homography
  - default_calibration_path returns a platform-appropriate path
"""

import sys
import os
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import camera_capture as cam

passed = 0
failed = 0


def check(name, condition):
    global passed, failed
    if condition:
        print(f"  PASS: {name}")
        passed += 1
    else:
        print(f"  FAIL: {name}")
        failed += 1


if not cam.HAS_OPENCV:
    print("\n!! OpenCV not available - skipping camera_capture tests")
    sys.exit(0)

import cv2
import numpy as np


# ============================================================
print("\n=== Module-level constants & gating ===")
# ============================================================

check("HAS_OPENCV is True when cv2 is importable", cam.HAS_OPENCV is True)
check("CHARUCO_COLS x ROWS defaults are reasonable",
      cam.CHARUCO_COLS >= 4 and cam.CHARUCO_ROWS >= 3)
check("CHARUCO_SQUARE_MM positive", cam.CHARUCO_SQUARE_MM > 0)
check("CHARUCO_MARKER_MM smaller than square",
      cam.CHARUCO_MARKER_MM < cam.CHARUCO_SQUARE_MM)
check("CALIB_MIN_FRAMES at least 4 (geometric minimum)",
      cam.CALIB_MIN_FRAMES >= 4)


# ============================================================
print("\n=== make_charuco_board ===")
# ============================================================

board = cam.make_charuco_board()
check("make_charuco_board returns an object",
      board is not None)
check("board.getChessboardSize matches defaults",
      tuple(board.getChessboardSize()) == (cam.CHARUCO_COLS, cam.CHARUCO_ROWS))


# ============================================================
print("\n=== render_charuco_card_png ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    png_path = os.path.join(tmpdir, "card.png")
    cam.render_charuco_card_png(png_path, dpi=300)
    check("PNG file created", os.path.exists(png_path))

    img = cv2.imread(png_path)
    check("PNG is readable by cv2", img is not None)

    # At 300 DPI, 6 cols × 25mm + 2× 10mm border ~= 2006 px wide
    # (slight slop because px_per_square is rounded to whole pixels).
    expected_w = int(cam.CHARUCO_COLS * cam.CHARUCO_SQUARE_MM * 300 / 25.4
                     + 2 * 10.0 * 300 / 25.4)
    h, w = img.shape[:2]
    check(f"PNG width ~{expected_w} px (got {w})",
          abs(w - expected_w) <= 4)

    # Detect ChArUco corners back from the rendered image itself.
    # Should find all interior corners since there's no distortion / noise.
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    corners, ids = cam.detect_charuco(gray, board)
    check("detect_charuco finds corners on the rendered card",
          corners is not None and len(corners) >= 4)
    interior_corners = (cam.CHARUCO_COLS - 1) * (cam.CHARUCO_ROWS - 1)
    check(f"detected all {interior_corners} interior corners",
          corners is not None and len(corners) == interior_corners)


# ============================================================
print("\n=== detect_charuco returns None when nothing visible ===")
# ============================================================

blank = np.full((480, 640), 200, dtype=np.uint8)
c, i = cam.detect_charuco(blank, board)
check("blank image -> (None, None)", c is None and i is None)


# ============================================================
print("\n=== Calibration save/load round-trip ===")
# ============================================================

fake_calib = {
    'opencv_version': '4.13.0',
    'camera_matrix': [[1000.0, 0.0, 320.0],
                       [0.0, 1000.0, 240.0],
                       [0.0, 0.0, 1.0]],
    'dist_coeffs': [[-0.3, 0.1, 0.0, 0.0, 0.0]],
    'rms_reprojection_error_px': 0.42,
    'image_size': [640, 480],
    'frame_count': 12,
    'homography_px_to_machine_mm': [[0.5, 0.0, -100.0],
                              [0.0, 0.5, -75.0],
                              [0.0, 0.0, 1.0]],
    'board': {'cols': 6, 'rows': 4, 'square_mm': 25.0,
              'marker_mm': 18.0, 'dict': 'DICT_4X4_50'},
}

with tempfile.TemporaryDirectory() as tmpdir:
    path = os.path.join(tmpdir, "calib.json")
    cam.save_calibration(fake_calib, path)
    check("calibration file written", os.path.exists(path))

    loaded = cam.load_calibration(path)
    check("loaded matches saved (camera_matrix)",
          loaded['camera_matrix'] == fake_calib['camera_matrix'])
    check("loaded matches saved (homography)",
          loaded['homography_px_to_machine_mm'] == fake_calib['homography_px_to_machine_mm'])

    # load_calibration returns None for missing file
    missing = cam.load_calibration(os.path.join(tmpdir, "nope.json"))
    check("load_calibration returns None for missing file", missing is None)


# ============================================================
print("\n=== Synthetic full pipeline ===")
# ============================================================
# Render a ChArUco board sized so the natural pixels-per-square is high
# enough for marker detection, paste it into several 640x480 "frames" at
# different positions (no scaling), detect the corners, run calibration,
# verify the reprojection error is small, and confirm undistort runs.
# We avoid warpAffine because it blurs the ArUco markers below the
# detector's threshold at small sizes.

with tempfile.TemporaryDirectory() as tmpdir:
    # Render the card at a size that fits in 640x480 at full scale.
    # At DPI=35 with the 8x6 default board (200x150mm + 5mm border on
    # each side), the rendered card is ~290x220 px, leaving room to
    # paste it at several positions inside the canvas.
    png_path = os.path.join(tmpdir, "card.png")
    # dpi=20 keeps the rendered 10×10×27mm card under ~270px so it
    # fits inside the 480x640 synthetic test canvas with room for
    # multiple positions.
    cam.render_charuco_card_png(png_path, dpi=20, border_mm=5)
    card_img = cv2.imread(png_path)
    card_gray = cv2.cvtColor(card_img, cv2.COLOR_BGR2GRAY)
    ch, cw = card_gray.shape[:2]
    check(f"synthetic card fits in 640x480 ({cw}x{ch})",
          cw < 640 and ch < 480)

    image_size = (640, 480)
    detections = []
    # Translate the card to several positions inside the frame. Position
    # variation alone is enough for the synthetic case since the card
    # has no perspective distortion to solve for.
    max_x = 640 - cw
    max_y = 480 - ch
    positions = [
        (0, 0), (max_x, 0), (0, max_y), (max_x, max_y),
        (max_x // 2, 0), (max_x // 2, max_y),
        (0, max_y // 2), (max_x, max_y // 2),
    ]
    for tx, ty in positions:
        if tx + cw > 640 or ty + ch > 480:
            continue
        canvas = np.full((480, 640), 128, dtype=np.uint8)
        canvas[ty:ty + ch, tx:tx + cw] = card_gray
        c, ids = cam.detect_charuco(canvas, board)
        if c is not None and len(c) >= 4:
            detections.append((c, ids))

    check(f"detected charuco in >= {cam.CALIB_MIN_FRAMES} synthetic frames",
          len(detections) >= cam.CALIB_MIN_FRAMES)

    if len(detections) >= cam.CALIB_MIN_FRAMES:
        calib = cam.calibrate_from_frames(detections, image_size, board)
        check("calibration produced a camera_matrix",
              'camera_matrix' in calib
              and len(calib['camera_matrix']) == 3)
        check("calibration produced dist_coeffs",
              'dist_coeffs' in calib)
        check("rms_reprojection_error_px is finite",
              np.isfinite(calib['rms_reprojection_error_px']))
        # No real distortion, so reprojection error should be tiny
        check("rms_reprojection_error_px < 1 px on undistorted synthetic",
              calib['rms_reprojection_error_px'] < 1.0)
        check("homography_px_to_machine_mm shape is 3x3",
              len(calib['homography_px_to_machine_mm']) == 3
              and len(calib['homography_px_to_machine_mm'][0]) == 3)

        # Undistort one of the synthetic frames using the calibration —
        # should run without error and preserve shape.
        sample = np.full((480, 640), 128, dtype=np.uint8)
        sample[100:100 + ch, 100:100 + cw] = card_gray
        und = cam.undistort_frame(sample, calib)
        check("undistort_frame returns same-shape image",
              und.shape == sample.shape)


# ============================================================
print("\n=== detect_scrap_contour on synthetic image ===")
# ============================================================

# Build an image with a black background and a bright triangle.
img = np.full((480, 640), 30, dtype=np.uint8)
triangle = np.array([[200, 100], [500, 200], [300, 400]], dtype=np.int32)
cv2.fillPoly(img, [triangle], 220)

poly = cam.detect_scrap_contour(img)
check("contour detected on synthetic triangle", poly is not None)
check("triangle approximated as ~3-vertex polygon",
      poly is not None and 3 <= len(poly) <= 5)

# Tiny shape below threshold should yield None
img_tiny = np.full((480, 640), 30, dtype=np.uint8)
tiny = np.array([[10, 10], [15, 10], [12, 15]], dtype=np.int32)
cv2.fillPoly(img_tiny, [tiny], 220)
check("contour below min_area_frac -> None",
      cam.detect_scrap_contour(img_tiny) is None)


# ============================================================
print("\n=== pixels_to_mm uses the homography ===")
# ============================================================

# Use an identity homography: pixel coords == mm coords. Mark as
# schema 2 so the legacy Y-flip correction doesn't fire.
identity_calib = {
    'calibration_schema_version': 2,
    'camera_matrix': [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]],
    'dist_coeffs': [[0.0, 0.0, 0.0, 0.0, 0.0]],
    'homography_px_to_machine_mm': [[1.0, 0.0, 0.0],
                              [0.0, 1.0, 0.0],
                              [0.0, 0.0, 1.0]],
}
mm = cam.pixels_to_mm([(100, 200), (300, 400)], identity_calib)
check("identity homography preserves coords",
      mm == [(100.0, 200.0), (300.0, 400.0)])

# Scale-by-0.5 homography: pixels halved into mm
scale_calib = dict(identity_calib)
scale_calib['homography_px_to_machine_mm'] = [[0.5, 0.0, 0.0],
                                        [0.0, 0.5, 0.0],
                                        [0.0, 0.0, 1.0]]
mm = cam.pixels_to_mm([(100, 200), (300, 400)], scale_calib)
check("scale-0.5 homography halves the coords",
      mm == [(50.0, 100.0), (150.0, 200.0)])

# Legacy calibration: schema_version absent → pixels_to_mm applies a
# Y-flip correction to undo the broken _board_corner_to_machine_mm
# convention used during schema-1 fits. With a known card geometry
# (offset_y=99, border=10, label_h=8, board_h=250) the constant K is
# 2*99 + 2*10 + 8 + 250 = 476. So a homography output of (x, 142)
# should be flipped to (x, 334).
legacy_calib = {
    'camera_matrix': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
    'dist_coeffs': [[0]],
    'homography_px_to_machine_mm': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
    'card_engrave_offset_mm': [66.0, 99.0],
    'card_border_mm': 10.0,
    'card_label_height_mm': 8.0,
    'board': {'rows': 10, 'square_mm': 25.0},
    # NOTE: no calibration_schema_version → treated as schema 1
}
mm = cam.pixels_to_mm([(100.0, 142.0), (200.0, 342.0)], legacy_calib)
check("legacy calibration Y-flips around card center",
      mm == [(100.0, 334.0), (200.0, 134.0)])

# Empty polygon round-trips
check("pixels_to_mm([]) returns []", cam.pixels_to_mm([], identity_calib) == [])
check("pixels_to_mm with None calibration returns input unchanged",
      cam.pixels_to_mm([(1, 2)], None) == [(1, 2)])


# ============================================================
print("\n=== Integrated calibration helpers ===")
# ============================================================

# Board-corner to machine-mm: board frame is Y-DOWN (OpenCV convention,
# verified by detecting a freshly-rendered board image), so corner
# (0, 0) is the TOP-LEFT of the board area which lands at machine
# (offset_x + border, offset_y + border + board_h) — the back-left of
# the engraved card area.
m_xy = cam._board_corner_to_machine_mm(
    (0.0, 0.0), offset_x=50.0, offset_y=50.0,
    board_h_mm=250.0, border_mm=10.0, label_height_mm=8.0)
check("board (0,0) -> machine (offset_x + border, offset_y + border + board_h)",
      m_xy == (60.0, 310.0))

# Bottom-right of the board area in board frame is (board_w, board_h);
# after Y-flip it lands at the FRONT of the board area, just above
# the label strip: machine (offset_x + border + bw, offset_y + border).
m_xy_far = cam._board_corner_to_machine_mm(
    (200.0, 250.0), offset_x=50.0, offset_y=50.0,
    board_h_mm=250.0, border_mm=10.0, label_height_mm=8.0)
check("board (200, board_h) -> machine (260, 60)",
      m_xy_far == (260.0, 60.0))

# Constants exist and are sane defaults.
check("CARD_ENGRAVE_OFFSET_X_MM is 50",
      cam.CARD_ENGRAVE_OFFSET_X_MM == 50.0)
check("CARD_ENGRAVE_OFFSET_Y_MM is 50",
      cam.CARD_ENGRAVE_OFFSET_Y_MM == 50.0)


# ============================================================
print("\n=== Legacy calibration detection ===")
# ============================================================

with tempfile.TemporaryDirectory() as tmpdir:
    legacy_cal = {
        'camera_matrix': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
        'dist_coeffs': [[0]],
        'homography_px_to_mm': [[1, 0, 0], [0, 1, 0], [0, 0, 1]],
        # Note: no homography_px_to_machine_mm
    }
    p = os.path.join(tmpdir, "legacy.json")
    import json
    with open(p, 'w') as f:
        json.dump(legacy_cal, f)
    check("is_legacy_calibration True on old format",
          cam.is_legacy_calibration(p))
    check("load_calibration returns None on legacy format",
          cam.load_calibration(p) is None)

    new_cal = dict(legacy_cal)
    del new_cal['homography_px_to_mm']
    new_cal['homography_px_to_machine_mm'] = [[1, 0, 0], [0, 1, 0], [0, 0, 1]]
    p2 = os.path.join(tmpdir, "new.json")
    with open(p2, 'w') as f:
        json.dump(new_cal, f)
    check("is_legacy_calibration False on new format",
          not cam.is_legacy_calibration(p2))
    check("load_calibration loads new format successfully",
          cam.load_calibration(p2) is not None)
    check("is_legacy_calibration False on missing file",
          not cam.is_legacy_calibration(
              os.path.join(tmpdir, "nope.json")))


# ============================================================
print("\n=== default_calibration_path ===")
# ============================================================

path = cam.default_calibration_path()
check("default_calibration_path returns a string", isinstance(path, str))
check("path ends in camera_calibration.json",
      path.endswith('camera_calibration.json'))
check("path includes StohrerSaxShopCompanion folder",
      'StohrerSaxShopCompanion' in path)


# ============================================================
print("\n=== Dot calibration: detect_dot_centers on synthetic image ===")
# ============================================================
#
# Render a synthetic image of an 8x8 grid of 6mm dots + 1x 10mm marker
# at known pixel positions. detect_dot_centers should find them all.
# Then match_dots_to_grid + fit_dot_homography should round-trip the
# known machine coords back to machine coords with sub-mm error.

# Layout: 8x8 grid at 25mm spacing in local coords, marker at (200, 200).
# Engrave offset: (50, 50) machine mm. So grid dots are at machine
# (50 + col*25, 50 + row*25) for col, row in 0..7.
# Render at 2 px/mm scale → grid spans (100, 100) to (450, 450) in pixels.
# 600x600 image gives margin.
PX_PER_MM = 2.0
ENGRAVE_OFFSET = (50.0, 50.0)
IMG_W, IMG_H = 600, 600


def _machine_to_pixel(mx, my):
    # Identity-ish mapping: (mx - 0) * px_per_mm = px (just shifted).
    return (int(round(mx * PX_PER_MM)), int(round(my * PX_PER_MM)))


synth = np.full((IMG_H, IMG_W), 230, dtype=np.uint8)  # light background
# Grid dots
expected_grid_machine = []
for col in range(cam.DOT_GRID_COLS):
    for row in range(cam.DOT_GRID_ROWS):
        mx = ENGRAVE_OFFSET[0] + col * cam.DOT_SPACING_MM
        my = ENGRAVE_OFFSET[1] + row * cam.DOT_SPACING_MM
        px, py = _machine_to_pixel(mx, my)
        radius_px = int(round(cam.DOT_DIAMETER_MM / 2.0 * PX_PER_MM))
        cv2.circle(synth, (px, py), radius_px, 30, -1)
        expected_grid_machine.append((mx, my))
# Marker
marker_mx = ENGRAVE_OFFSET[0] + cam.DOT_MARKER_LOCAL_XY[0]
marker_my = ENGRAVE_OFFSET[1] + cam.DOT_MARKER_LOCAL_XY[1]
m_px, m_py = _machine_to_pixel(marker_mx, marker_my)
marker_radius_px = int(round(cam.DOT_MARKER_DIAMETER_MM / 2.0 * PX_PER_MM))
cv2.circle(synth, (m_px, m_py), marker_radius_px, 30, -1)

detected = cam.detect_dot_centers(synth, min_area_px=10, max_area_px=500)
check(f"detected at least 65 dots ({len(detected)} found)",
      len(detected) >= 65)

# Build a synthetic "old calibration" matching the identity-ish mapping.
# pixel_to_machine: mx = px / PX_PER_MM, my = py / PX_PER_MM.
# Homography for that: [[1/PX_PER_MM, 0, 0], [0, 1/PX_PER_MM, 0], [0, 0, 1]]
synthetic_old_cal = {
    'calibration_schema_version': 2,
    'camera_matrix': [[600.0, 0.0, IMG_W / 2.0],
                       [0.0, 600.0, IMG_H / 2.0],
                       [0.0, 0.0, 1.0]],
    'dist_coeffs': [[0.0, 0.0, 0.0, 0.0, 0.0]],
    'homography_px_to_machine_mm': [
        [1.0 / PX_PER_MM, 0.0, 0.0],
        [0.0, 1.0 / PX_PER_MM, 0.0],
        [0.0, 0.0, 1.0],
    ],
}

match_result = cam.match_dots_to_grid(
    detected, synthetic_old_cal, ENGRAVE_OFFSET)
check(f"match_dots_to_grid no error: {match_result.get('error')}",
      match_result.get('error') is None)
check(f"matched at least 60 grid dots ({len(match_result['matched_grid'])})",
      len(match_result['matched_grid']) >= 60)
check("marker_pixel identified",
      match_result.get('marker_pixel') is not None)

# Fit the homography and verify round-trip accuracy.
new_homography = cam.fit_dot_homography(match_result['matched_grid'])
check("fit_dot_homography returns 3x3", len(new_homography) == 3 and len(new_homography[0]) == 3)

# Round-trip: pick a few detected dots, apply new homography, compare to known machine coords.
new_cal = cam.build_dot_calibration(
    synthetic_old_cal, new_homography, ENGRAVE_OFFSET,
    len(match_result['matched_grid']))
errors = []
for (px, py), _cell, _local, (true_mx, true_my) in match_result['matched_grid']:
    pred = cam.pixels_to_mm([(px, py)], new_cal)[0]
    err = ((pred[0] - true_mx) ** 2 + (pred[1] - true_my) ** 2) ** 0.5
    errors.append(err)
mean_err = sum(errors) / len(errors)
max_err = max(errors)
check(f"round-trip mean error < 0.5mm (got {mean_err:.3f})", mean_err < 0.5)
check(f"round-trip max error < 1.5mm (got {max_err:.3f})", max_err < 1.5)


# ============================================================
print("\n=== Dot calibration: build_dot_calibration preserves intrinsics ===")
# ============================================================

# Take an existing calibration, build a dot-cal result, check that
# camera_matrix and dist_coeffs are unchanged (intrinsics preserved
# from the prior ChArUco fit — that's the whole point of dot-cal).
test_old = {
    'calibration_schema_version': 2,
    'camera_matrix': [[600.0, 0.0, 320.0], [0.0, 600.0, 240.0], [0.0, 0.0, 1.0]],
    'dist_coeffs': [[-0.3, 0.1, 0.0, 0.0, 0.0]],
    'rms_reprojection_error_px': 0.42,
    'homography_px_to_machine_mm': [[0.5, 0.0, -100.0],
                              [0.0, 0.5, -75.0],
                              [0.0, 0.0, 1.0]],
    'image_size': [640, 480],
    'frame_count': 12,
}
new_h = [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]]
test_new = cam.build_dot_calibration(test_old, new_h, (60.0, 70.0), 50)

check("intrinsics preserved: camera_matrix",
      test_new['camera_matrix'] == test_old['camera_matrix'])
check("intrinsics preserved: dist_coeffs",
      test_new['dist_coeffs'] == test_old['dist_coeffs'])
check("intrinsics preserved: rms_reprojection_error_px",
      test_new['rms_reprojection_error_px']
      == test_old['rms_reprojection_error_px'])
check("homography overwritten",
      test_new['homography_px_to_machine_mm'] == new_h)
check("schema bumped to 3",
      test_new['calibration_schema_version'] == 3)
check("card_engrave_offset_mm overwritten",
      test_new['card_engrave_offset_mm'] == [60.0, 70.0])
check("calibration_method set to dot_grid",
      test_new.get('calibration_method') == 'dot_grid')


# ============================================================
print("\n=== Dot calibration: compare_calibrations math ===")
# ============================================================

# Build two calibrations and a matched_grid with synthetic data.
# Old maps pixel (100, 100) → machine (50, 50). Truth is (60, 60). Error = sqrt(200) ~= 14.14mm.
# New maps pixel (100, 100) → machine (60, 60). Error = 0.
fake_matched = [
    ((100.0, 100.0), (0, 0), (0.0, 0.0), (60.0, 60.0)),
]
cal_off = {
    'calibration_schema_version': 2,
    'camera_matrix': [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]],
    'dist_coeffs': [[0.0, 0.0, 0.0, 0.0, 0.0]],
    'homography_px_to_machine_mm': [[0.5, 0.0, 0.0], [0.0, 0.5, 0.0], [0.0, 0.0, 1.0]],
}
cal_perfect = {
    'calibration_schema_version': 2,
    'camera_matrix': [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]],
    'dist_coeffs': [[0.0, 0.0, 0.0, 0.0, 0.0]],
    'homography_px_to_machine_mm': [[0.6, 0.0, 0.0], [0.0, 0.6, 0.0], [0.0, 0.0, 1.0]],
}
comp = cam.compare_calibrations(cal_off, cal_perfect, fake_matched)
check(f"compare: old error ~= 14.14 (got {comp['old_avg_mm']:.2f})",
      abs(comp['old_avg_mm'] - 14.142) < 0.1)
check(f"compare: new error ~= 0 (got {comp['new_avg_mm']:.4f})",
      comp['new_avg_mm'] < 0.001)
check("compare: n_dots matches input", comp['n_dots'] == 1)


# ============================================================
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
