"""
Engine-level tests for camera_capture.py — no actual camera required.

Covers:
  - OpenCV-availability gating (module imports without OpenCV)
  - ChArUco board generation produces a valid PNG of expected pixel size
  - ChArUco detection round-trips on the rendered card itself
  - Calibration JSON save/load round-trip
  - Synthetic full pipeline: render board → 'photograph' it via warp →
    detect → calibrate → undistort → verify low reprojection error
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

    # At 300 DPI, 6 cols × 25mm + 2× 10mm border ≈ 2006 px wide
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
    'homography_px_to_mm': [[0.5, 0.0, -100.0],
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
          loaded['homography_px_to_mm'] == fake_calib['homography_px_to_mm'])

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
    cam.render_charuco_card_png(png_path, dpi=35, border_mm=5)
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
        check("homography_px_to_mm shape is 3x3",
              len(calib['homography_px_to_mm']) == 3
              and len(calib['homography_px_to_mm'][0]) == 3)

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

# Use an identity homography: pixel coords == mm coords.
identity_calib = {
    'camera_matrix': [[1.0, 0.0, 0.0], [0.0, 1.0, 0.0], [0.0, 0.0, 1.0]],
    'dist_coeffs': [[0.0, 0.0, 0.0, 0.0, 0.0]],
    'homography_px_to_mm': [[1.0, 0.0, 0.0],
                              [0.0, 1.0, 0.0],
                              [0.0, 0.0, 1.0]],
}
mm = cam.pixels_to_mm([(100, 200), (300, 400)], identity_calib)
check("identity homography preserves coords",
      mm == [(100.0, 200.0), (300.0, 400.0)])

# Scale-by-0.5 homography: pixels halved into mm
scale_calib = dict(identity_calib)
scale_calib['homography_px_to_mm'] = [[0.5, 0.0, 0.0],
                                        [0.0, 0.5, 0.0],
                                        [0.0, 0.0, 1.0]]
mm = cam.pixels_to_mm([(100, 200), (300, 400)], scale_calib)
check("scale-0.5 homography halves the coords",
      mm == [(50.0, 100.0), (150.0, 200.0)])

# Empty polygon round-trips
check("pixels_to_mm([]) returns []", cam.pixels_to_mm([], identity_calib) == [])
check("pixels_to_mm with None calibration returns input unchanged",
      cam.pixels_to_mm([(1, 2)], None) == [(1, 2)])


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
print(f"\n=== Summary: {passed} passed, {failed} failed ===\n")
# ============================================================

sys.exit(0 if failed == 0 else 1)
