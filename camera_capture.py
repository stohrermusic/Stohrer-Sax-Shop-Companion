"""Camera-based scrap shape capture using a ChArUco calibration card.

OpenCV is an optional dependency. The rest of the app degrades gracefully
when it's not installed (the "Get from camera" buttons stay hidden until
both OpenCV imports cleanly AND a calibration file exists).

Workflow:
    1. User generates a ChArUco calibration card via the Tooling tab and
       prints / engraves it on a sheet of paper at known scale.
    2. One-time calibration: user places the card on the laser bed at
       several poses; ``calibrate_from_frames`` solves for the camera
       intrinsics + distortion coefficients and the bed-plane homography.
    3. Per-use capture: ``capture_scrap_polygon`` snaps a frame, undistorts
       it, finds the largest contour, simplifies it, and maps the contour
       through the homography to produce a polygon in bed-millimeter coords
       that the Pad Maker can consume.

Camera notes (validated on Creality Falcon2 Pro 40W):
    - Camera appears as a standard USB webcam under Windows.
    - Native resolution 640x480; higher resolutions are not negotiated.
    - The orange/red safety glass on the closed cover acts as a built-in
      ND filter, so auto-exposure works reliably with the cover closed.
      No manual exposure tuning required for normal operation.
    - Severe barrel distortion at the edges — calibration is mandatory.
"""

from __future__ import annotations

import json
import os
import platform
import sys

try:
    import cv2
    import numpy as np
    HAS_OPENCV = True
except ImportError:  # pragma: no cover
    cv2 = None
    np = None
    HAS_OPENCV = False


# =============================================================================
# Constants
# =============================================================================

# ChArUco card defaults. 6 columns x 4 rows of 25mm squares gives a card
# that's 150 x 100 mm — fits Letter / A4 with plenty of margin and shows
# enough markers (15) for OpenCV to solve a robust calibration even when
# part of the board is occluded.
CHARUCO_DICT_NAME = "DICT_4X4_50"
CHARUCO_COLS = 8
CHARUCO_ROWS = 6
CHARUCO_SQUARE_MM = 25.0
CHARUCO_MARKER_MM = 18.0  # ~0.72 * square_mm — OpenCV's recommended ratio

# Default engraving recipe for the calibration card on basswood
# (Creality Falcon2 Pro 40W). User can override in the Tooling tab.
CALIB_DEFAULT_ENG_SPEED = 6000  # mm/min
CALIB_DEFAULT_ENG_POWER = 25    # percent
CALIB_DEFAULT_ENG_PASSES = 1

# Camera probe range. The Falcon shows up at index 1 on Matt's machine
# (integrated laptop camera is index 0); 6 covers most realistic setups.
CAMERA_PROBE_INDICES = range(6)

# Contour-detection defaults. Tuned for the Falcon's 640x480 image with
# a leather scrap on the honeycomb bed.
SCRAP_MIN_AREA_FRAC = 0.01     # ignore contours smaller than 1% of frame
SCRAP_APPROX_EPS_FRAC = 0.005  # polygon approximation tolerance, 0.5% of perimeter

# Minimum frames worth of ChArUco detections to compute a reliable calibration.
CALIB_MIN_FRAMES = 6
CALIB_RECOMMENDED_FRAMES = 12


# =============================================================================
# OpenCV availability gate
# =============================================================================

def _require_opencv():
    if not HAS_OPENCV:
        raise RuntimeError(
            "OpenCV (opencv-python) is not installed; camera features unavailable."
        )


# =============================================================================
# Camera enumeration & open
# =============================================================================

def enumerate_cameras(max_index=6):
    """Probe a range of camera indices and return what's available.

    Returns a list of dicts: ``[{'index': int, 'width': int, 'height': int}, ...]``.
    Each dict represents a camera that opened AND returned at least one frame.
    """
    if not HAS_OPENCV:
        return []
    cams = []
    for i in range(max_index):
        cap = cv2.VideoCapture(i)
        if not cap.isOpened():
            cap.release()
            continue
        # Some cameras need a few reads to warm up before they return a frame.
        ok = False
        for _ in range(3):
            ok, _frame = cap.read()
            if ok:
                break
        w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
        h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
        cap.release()
        if ok:
            cams.append({'index': i, 'width': w, 'height': h})
    return cams


def find_falcon_camera_index():
    """Best-effort identification of the Falcon camera on Windows.

    Queries ``Get-PnpDevice`` for camera devices and matches "Falcon" by
    name. Returns the FIRST OpenCV index that successfully opens after
    seeing such a device in PnP. None if not found or off-Windows.

    OpenCV's webcam indices don't expose device names, so we can't map
    PnP name → OpenCV index directly. This function relies on the order
    being stable: if the Falcon is present in PnP, the Falcon's index is
    typically the LAST detected one (the integrated laptop cam being
    index 0).
    """
    if not HAS_OPENCV or platform.system() != 'Windows':
        return None
    try:
        import subprocess
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command',
             "Get-PnpDevice -Class Camera -PresentOnly | "
             "Select-Object -ExpandProperty FriendlyName"],
            capture_output=True, text=True, timeout=5,
        )
        names = [ln.strip() for ln in result.stdout.splitlines() if ln.strip()]
    except Exception:
        return None

    has_falcon = any('falcon' in n.lower() for n in names)
    if not has_falcon:
        return None

    cams = enumerate_cameras()
    if not cams:
        return None
    # If there are multiple cameras and the Falcon is present, assume it
    # is the highest-indexed one (integrated laptop cam is usually 0).
    return cams[-1]['index']


def open_camera(index, request_width=None, request_height=None):
    """Open a camera by index. Returns a VideoCapture (caller releases)."""
    _require_opencv()
    cap = cv2.VideoCapture(index)
    if not cap.isOpened():
        raise RuntimeError(f"Could not open camera index {index}")
    if request_width:
        cap.set(cv2.CAP_PROP_FRAME_WIDTH, request_width)
    if request_height:
        cap.set(cv2.CAP_PROP_FRAME_HEIGHT, request_height)
    return cap


def capture_frame(cap, warmup_frames=5):
    """Read ``warmup_frames`` discarded frames (lets auto-exposure settle),
    then return one good frame. Returns None on failure.
    """
    _require_opencv()
    for _ in range(warmup_frames):
        cap.read()
    ok, frame = cap.read()
    if not ok or frame is None:
        return None
    return frame


# =============================================================================
# ChArUco board generation
# =============================================================================

def _get_charuco_dict():
    _require_opencv()
    return cv2.aruco.getPredefinedDictionary(
        getattr(cv2.aruco, CHARUCO_DICT_NAME)
    )


def make_charuco_board(cols=CHARUCO_COLS, rows=CHARUCO_ROWS,
                        square_mm=CHARUCO_SQUARE_MM,
                        marker_mm=CHARUCO_MARKER_MM):
    """Construct a ChArUco board (OpenCV native object).

    The unit passed for ``square_mm`` / ``marker_mm`` flows through OpenCV
    unchanged — we use millimeters everywhere so calibration output is
    directly in mm.
    """
    dictionary = _get_charuco_dict()
    board = cv2.aruco.CharucoBoard(
        (cols, rows), square_mm, marker_mm, dictionary
    )
    return board


def make_calibration_card_strokes(cols=CHARUCO_COLS, rows=CHARUCO_ROWS,
                                    square_mm=CHARUCO_SQUARE_MM,
                                    marker_mm=CHARUCO_MARKER_MM,
                                    line_spacing_mm=0.15,
                                    origin_x_mm=0.0, origin_y_mm=0.0):
    """Convert the ChArUco board to horizontal scan-line strokes in mm.

    Each stroke is a 2-point list ``[(x0_mm, y_mm), (x1_mm, y_mm)]`` that
    can be passed to gcode_engine's ``generate_gcode_layer`` for raster
    engraving. The image is rendered at a DPI matched to ``line_spacing_mm``
    so every pixel row corresponds to exactly one engraved line.
    """
    _require_opencv()
    board = make_charuco_board(cols, rows, square_mm, marker_mm)

    # DPI chosen so each pixel row equals `line_spacing_mm` in width.
    px_per_mm = 1.0 / line_spacing_mm
    img_w = int(round(cols * square_mm * px_per_mm))
    img_h = int(round(rows * square_mm * px_per_mm))
    img = board.generateImage((img_w, img_h), marginSize=0)

    # Convert to a dark-pixel mask. `generateImage` returns 8-bit grayscale.
    mask = img < 128

    strokes = []
    mm_per_px_x = (cols * square_mm) / img_w
    mm_per_px_y = (rows * square_mm) / img_h
    for y in range(img_h):
        row = mask[y]
        if not row.any():
            continue
        # Find runs of dark pixels via diff on a padded boolean row.
        padded = np.concatenate(([False], row, [False])).astype(np.int8)
        diffs = np.diff(padded)
        starts = np.where(diffs == 1)[0]
        ends = np.where(diffs == -1)[0]
        y_mm = origin_y_mm + (y + 0.5) * mm_per_px_y  # center of pixel row
        for s, e in zip(starts, ends):
            x0 = origin_x_mm + s * mm_per_px_x
            x1 = origin_x_mm + e * mm_per_px_x
            strokes.append([(x0, y_mm), (x1, y_mm)])
    return strokes


def render_charuco_card_png(filename, cols=CHARUCO_COLS, rows=CHARUCO_ROWS,
                             square_mm=CHARUCO_SQUARE_MM,
                             marker_mm=CHARUCO_MARKER_MM,
                             dpi=300, border_mm=10.0):
    """Render the calibration card as a PNG sized for the requested DPI.

    Print at 100% scale to get accurate square sizes.

    A white border around the printable area gives the cutter / printer
    room and helps OpenCV detect the outer markers reliably.
    """
    _require_opencv()
    board = make_charuco_board(cols, rows, square_mm, marker_mm)

    px_per_mm = dpi / 25.4
    # Round to whole pixels per square so every cell ends up the same size
    # and OpenCV's generateImage doesn't fail an internal ROI assertion.
    px_per_square = int(round(square_mm * px_per_mm))
    inner_w_px = cols * px_per_square
    inner_h_px = rows * px_per_square
    border_px = int(round(border_mm * px_per_mm))

    inner = board.generateImage((inner_w_px, inner_h_px), marginSize=0)
    # Wrap in a white border so the printer has bleed and the outer
    # markers aren't trimmed.
    full_w = inner_w_px + 2 * border_px
    full_h = inner_h_px + 2 * border_px
    img = np.full((full_h, full_w), 255, dtype=np.uint8)
    img[border_px:border_px + inner_h_px,
        border_px:border_px + inner_w_px] = inner

    cv2.imwrite(filename, img)
    return filename


# =============================================================================
# ChArUco detection & calibration
# =============================================================================

def detect_charuco(frame, board=None):
    """Detect ChArUco corners/IDs in a frame.

    Returns ``(charuco_corners, charuco_ids)`` arrays or ``(None, None)``
    if detection failed.
    """
    _require_opencv()
    if board is None:
        board = make_charuco_board()
    detector = cv2.aruco.CharucoDetector(board)
    charuco_corners, charuco_ids, _marker_corners, _marker_ids = (
        detector.detectBoard(frame)
    )
    if charuco_corners is None or len(charuco_corners) < 4:
        return None, None
    return charuco_corners, charuco_ids


def calibrate_from_frames(detections, image_size, board=None):
    """Compute camera intrinsics + bed-plane homography from multiple captures.

    ``detections`` is a list of ``(charuco_corners, charuco_ids)`` tuples,
    one per usable frame. ``image_size`` is ``(width, height)`` in pixels.

    Returns a dict with the calibration data, suitable for ``save_calibration``.
    """
    _require_opencv()
    if board is None:
        board = make_charuco_board()
    if len(detections) < CALIB_MIN_FRAMES:
        raise ValueError(
            f"Need at least {CALIB_MIN_FRAMES} detected frames; got {len(detections)}"
        )

    # OpenCV 4.7+ removed cv2.aruco.calibrateCameraCharuco. Use the new
    # board.matchImagePoints API to convert each frame's detected ChArUco
    # corners into 3D-2D correspondences, then feed those to the standard
    # cv2.calibrateCamera.
    all_obj_points = []
    all_img_points = []
    for charuco_corners, charuco_ids in detections:
        obj_pts, img_pts = board.matchImagePoints(charuco_corners, charuco_ids)
        if obj_pts is None or len(obj_pts) < 4:
            continue
        all_obj_points.append(obj_pts)
        all_img_points.append(img_pts)
    if len(all_obj_points) < CALIB_MIN_FRAMES:
        raise ValueError(
            f"Need at least {CALIB_MIN_FRAMES} usable frames after "
            f"matchImagePoints; got {len(all_obj_points)}"
        )

    rms, camera_matrix, dist_coeffs, _rvecs, _tvecs = cv2.calibrateCamera(
        all_obj_points, all_img_points, image_size, None, None
    )

    # Bed-plane homography: assume the last detection was the "card flat on
    # the bed" pose. Map detected ChArUco corner pixel positions to their
    # known mm positions on the card to get the H pixel→mm transform.
    bed_corners_px, bed_ids = detections[-1]
    board_corners_mm = board.getChessboardCorners()  # Nx3, mm coords
    # Pick out only the corners we actually detected, in the same order.
    src_pts = bed_corners_px.reshape(-1, 2).astype(np.float32)
    dst_pts = np.array(
        [board_corners_mm[int(i[0]), :2] for i in bed_ids],
        dtype=np.float32,
    )
    if len(src_pts) < 4:
        raise ValueError("Last detection has too few corners for homography")

    # Undistort the source points so the homography lives in the undistorted
    # image plane, matching what undistort_frame produces at use time.
    undist_src = cv2.undistortPoints(
        src_pts.reshape(-1, 1, 2), camera_matrix, dist_coeffs, P=camera_matrix
    ).reshape(-1, 2)
    homography, _ = cv2.findHomography(undist_src, dst_pts, cv2.RANSAC, 3.0)

    return {
        'opencv_version': cv2.__version__,
        'camera_matrix': camera_matrix.tolist(),
        'dist_coeffs': dist_coeffs.tolist(),
        'rms_reprojection_error_px': float(rms),
        'image_size': list(image_size),
        'frame_count': len(detections),
        'homography_px_to_mm': homography.tolist(),
        'board': {
            'cols': CHARUCO_COLS,
            'rows': CHARUCO_ROWS,
            'square_mm': CHARUCO_SQUARE_MM,
            'marker_mm': CHARUCO_MARKER_MM,
            'dict': CHARUCO_DICT_NAME,
        },
    }


def save_calibration(calibration, path):
    """Save calibration dict as JSON."""
    os.makedirs(os.path.dirname(path) or '.', exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(calibration, f, indent=2)


def load_calibration(path):
    """Load a previously-saved calibration JSON. Returns None if not present."""
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def _calibration_arrays(calibration):
    if not HAS_OPENCV or not calibration:
        return None, None, None
    return (
        np.array(calibration['camera_matrix'], dtype=np.float64),
        np.array(calibration['dist_coeffs'], dtype=np.float64),
        np.array(calibration['homography_px_to_mm'], dtype=np.float64),
    )


# =============================================================================
# Undistortion + per-use capture
# =============================================================================

def undistort_frame(frame, calibration):
    """Apply calibration's undistortion to a frame."""
    _require_opencv()
    if not calibration:
        return frame
    camera_matrix, dist_coeffs, _h = _calibration_arrays(calibration)
    return cv2.undistort(frame, camera_matrix, dist_coeffs, None, camera_matrix)


def detect_scrap_contour(frame, min_area_frac=SCRAP_MIN_AREA_FRAC,
                          epsilon_frac=SCRAP_APPROX_EPS_FRAC):
    """Find the largest contour in the frame and approximate it as a polygon.

    Returns a list of (x, y) pixel-coordinate tuples or ``None`` if no
    contour above ``min_area_frac`` was found. Uses Otsu thresholding,
    which works well when the scrap is significantly brighter or darker
    than the bed (typical case for leather on the Falcon's honeycomb).
    """
    _require_opencv()
    if frame is None:
        return None
    gray = (
        cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if len(frame.shape) == 3 else frame
    )
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _thresh_val, thresh = cv2.threshold(
        blurred, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU
    )
    contours, _hier = cv2.findContours(
        thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE
    )
    if not contours:
        return None

    img_area = frame.shape[0] * frame.shape[1]
    largest = max(contours, key=cv2.contourArea)
    if cv2.contourArea(largest) < img_area * min_area_frac:
        return None

    epsilon = epsilon_frac * cv2.arcLength(largest, True)
    approx = cv2.approxPolyDP(largest, epsilon, True)
    return [(float(p[0][0]), float(p[0][1])) for p in approx]


def pixels_to_mm(polygon_px, calibration):
    """Apply the bed-plane homography to convert a pixel-coord polygon
    into bed-millimeter coordinates.
    """
    _require_opencv()
    if not calibration or not polygon_px:
        return polygon_px
    _cm, _dc, homography = _calibration_arrays(calibration)
    pts = np.array(polygon_px, dtype=np.float64).reshape(-1, 1, 2)
    transformed = cv2.perspectiveTransform(pts, homography)
    return [(float(p[0][0]), float(p[0][1])) for p in transformed]


def capture_scrap_polygon(cap, calibration, min_area_frac=SCRAP_MIN_AREA_FRAC,
                           epsilon_frac=SCRAP_APPROX_EPS_FRAC,
                           warmup_frames=5):
    """End-to-end capture: snap → undistort → find largest contour → mm.

    Returns the polygon as a list of ``(x_mm, y_mm)`` tuples or None.
    """
    _require_opencv()
    frame = capture_frame(cap, warmup_frames=warmup_frames)
    if frame is None:
        return None
    undistorted = undistort_frame(frame, calibration)
    polygon_px = detect_scrap_contour(
        undistorted, min_area_frac=min_area_frac, epsilon_frac=epsilon_frac
    )
    if not polygon_px:
        return None
    return pixels_to_mm(polygon_px, calibration)


# =============================================================================
# Public utility: where to store calibration on disk
# =============================================================================

def default_calibration_path():
    """Return platform-appropriate path for the calibration file."""
    # Mirror config.py's locations to avoid pulling that module here.
    if sys.platform == 'win32':
        base = os.path.join(os.environ.get('APPDATA', '.'),
                             'StohrerSaxShopCompanion')
    elif sys.platform == 'darwin':
        base = os.path.join(os.path.expanduser('~'),
                             'Library', 'Application Support',
                             'StohrerSaxShopCompanion')
    else:
        base = os.path.join(
            os.environ.get('XDG_CONFIG_HOME',
                            os.path.join(os.path.expanduser('~'), '.config')),
            'StohrerSaxShopCompanion'
        )
    return os.path.join(base, 'camera_calibration.json')
