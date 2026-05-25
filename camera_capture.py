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
    # Silence OpenCV's DSHOW probe warnings ("backend is generally
    # available but can't be used to capture by index N"). They fire
    # for every empty camera slot we probe in enumerate_cameras and
    # clutter the console with apparent errors that aren't.
    try:
        cv2.setLogLevel(cv2.LOG_LEVEL_ERROR)
    except (AttributeError, Exception):
        try:
            cv2.utils.logging.setLogLevel(
                cv2.utils.logging.LOG_LEVEL_ERROR)
        except Exception:
            pass
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
# 10×10 board × 25mm squares = 250×250mm board. With 10mm borders +
# 8mm label strip the engraved card is ~278×278mm ≈ 10.9" square.
# On a 12×12-inch (305mm) basswood blank that leaves ~½" margin all
# around — generous enough that small positioning errors don't put
# the trace off-material. Square layout = equal corner spread in
# X and Y. 50 markers needed (half of 100 squares); DICT_4X4_50
# has exactly 50 IDs.
CHARUCO_COLS = 10
CHARUCO_ROWS = 10
CHARUCO_SQUARE_MM = 25.0
CHARUCO_MARKER_MM = 18.0  # ~0.72 * square_mm — OpenCV's recommended ratio

# Default machine-coord position for the card's bottom-left corner
# when SSC engraves it. The integrated calibration dialog now lets
# the user JOG the head into the camera's view and engrave around
# the head's current position, so these defaults are only used as a
# fallback when the user doesn't pick a position.
CARD_ENGRAVE_OFFSET_X_MM = 50.0
CARD_ENGRAVE_OFFSET_Y_MM = 50.0

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
    use_dshow = platform.system() == "Windows"
    cams = []
    for i in range(max_index):
        cap = (cv2.VideoCapture(i, cv2.CAP_DSHOW)
               if use_dshow else cv2.VideoCapture(i))
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
    """Open a camera by index. Returns a VideoCapture (caller releases).

    Forces the DirectShow backend on Windows. Without it, OpenCV
    defaults to Media Foundation (CAP_MSMF) which adds 2-3 seconds of
    initialization latency on every open — DirectShow opens USB webcams
    in well under a second. Other platforms get the default backend.
    """
    _require_opencv()
    if platform.system() == "Windows":
        cap = cv2.VideoCapture(index, cv2.CAP_DSHOW)
    else:
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
                                    origin_x_mm=0.0, origin_y_mm=0.0,
                                    orientation_label="FRONT OF MACHINE",
                                    label_height_mm=8.0):
    """Convert the ChArUco board to horizontal scan-line strokes in mm.

    Each stroke is a 2-point list ``[(x0_mm, y_mm), (x1_mm, y_mm)]`` that
    can be passed to gcode_engine's ``generate_gcode_layer`` for raster
    engraving. The image is rendered at a DPI matched to ``line_spacing_mm``
    so every pixel row corresponds to exactly one engraved line.

    ``orientation_label`` is rendered below the ChArUco board (via cv2.putText)
    so the user can place the card in the same orientation every time. The
    text is engraved on the bottom edge — lay the card with this text
    closest to the front of the machine for the captured-polygon coordinate
    system to align with the user's mental model of the bed. Pass an empty
    string to skip the label.
    """
    _require_opencv()
    board = make_charuco_board(cols, rows, square_mm, marker_mm)

    # DPI chosen so each pixel row equals `line_spacing_mm` in width.
    px_per_mm = 1.0 / line_spacing_mm
    board_w_px = int(round(cols * square_mm * px_per_mm))
    board_h_px = int(round(rows * square_mm * px_per_mm))
    board_img = board.generateImage((board_w_px, board_h_px), marginSize=0)

    # Compose a final image that has the ChArUco board on top and an
    # orientation-label strip below (white background, black text).
    label_h_px = int(round(label_height_mm * px_per_mm)) if orientation_label else 0
    img_w = board_w_px
    img_h = board_h_px + label_h_px
    img = np.full((img_h, img_w), 255, dtype=np.uint8)
    img[:board_h_px, :] = board_img

    if orientation_label and label_h_px > 0:
        font = cv2.FONT_HERSHEY_DUPLEX
        # Pick a font scale that makes the text height ~60% of the label strip
        target_text_h_px = int(label_h_px * 0.6)
        # The font face Hershey-Duplex at scale=1 produces text ~22px tall;
        # scale linearly from there.
        font_scale = max(0.4, target_text_h_px / 22.0)
        thickness = max(1, int(round(font_scale * 1.5)))
        (text_w, text_h), _baseline = cv2.getTextSize(
            orientation_label, font, font_scale, thickness)
        text_x = max(0, (img_w - text_w) // 2)
        text_y = board_h_px + (label_h_px + text_h) // 2
        cv2.putText(img, orientation_label, (text_x, text_y), font,
                     font_scale, color=0, thickness=thickness,
                     lineType=cv2.LINE_AA)

    # Convert to a dark-pixel mask. Scan-line every row.
    mask = img < 128

    strokes = []
    total_h_mm = (img_h / px_per_mm)
    mm_per_px_x = (cols * square_mm) / img_w
    mm_per_px_y = total_h_mm / img_h
    for y in range(img_h):
        row = mask[y]
        if not row.any():
            continue
        padded = np.concatenate(([False], row, [False])).astype(np.int8)
        diffs = np.diff(padded)
        starts = np.where(diffs == 1)[0]
        ends = np.where(diffs == -1)[0]
        y_mm = origin_y_mm + (y + 0.5) * mm_per_px_y
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


def _board_corner_to_machine_mm(board_xy, offset_x, offset_y,
                                  board_h_mm,
                                  border_mm=10.0, label_height_mm=8.0):
    """Convert a ChArUco board-frame corner position to machine-mm,
    given the engrave offset and card geometry.

    Card layout (assuming engrave G-code uses make_calibration_card_strokes
    with default border + label):

        ┌────────────────────────┐ ← back of machine (high Y)
        │ border                 │
        │  ┌──────────────────┐  │
        │  │  ChArUco board   │  │
        │  │  (image y=0 at   │  │
        │  │   top, y=board_h │  │
        │  │   at bottom)     │  │
        │  └──────────────────┘  │
        │ FRONT OF MACHINE label │
        └────────────────────────┘ ← front of machine (low Y, machine origin nearby)
        ↑ offset_x               ↑ offset_x + sheet_w

    OpenCV's CharucoBoard.getChessboardCorners() returns points in a
    Y-DOWN frame matching the rendered image (verified empirically: a
    corner at board (25, 25) lands at image pixel (100, 100) when
    rendered at 4 px/mm — top-left area). The engraved card places
    image-top at the back of the machine (high machine Y), so board Y
    must be FLIPPED to machine Y:

      board (0, 0)              → image top-left      → machine high Y, low X
      board (board_w, board_h)  → image bottom-right  → machine low Y,  high X

    The board area in machine coords spans
    ``[offset_y + border, offset_y + border + board_h]``.
    """
    bx, by = float(board_xy[0]), float(board_xy[1])
    machine_x = offset_x + border_mm + bx
    machine_y = offset_y + border_mm + (board_h_mm - by)
    # label_height_mm is part of the function signature for source-of-
    # truth bookkeeping but the board-area math doesn't depend on it
    # (the label sits BELOW the board area in machine Y, at
    # [offset_y + border - label_h, offset_y + border]).
    _ = label_height_mm
    return (machine_x, machine_y)


def calibrate_from_frames(detections, image_size, board=None,
                           reference_indices=None,
                           card_offset_x_mm=None,
                           card_offset_y_mm=None,
                           card_border_mm=10.0,
                           card_label_height_mm=8.0):
    """Compute camera intrinsics + pixel→machine-mm homography from
    multiple captures.

    Integrated calibration:
      - All captures feed ``cv2.calibrateCamera`` for intrinsics +
        distortion (the more poses the better).
      - One or more REFERENCE captures (``reference_indices``, default
        ``(0,)`` — just the first) are taken with the card untouched
        at its engraved position. Each detected ChArUco corner in
        those frames has a KNOWN machine-mm position. Pooling all
        reference corners and fitting one homography averages out
        detection noise → more accurate pixel→machine-mm mapping.

    ``card_offset_x_mm`` / ``card_offset_y_mm`` default to the
    constants the engrave G-code used (``CARD_ENGRAVE_OFFSET_*_MM``).

    Returns a calibration dict suitable for ``save_calibration``.
    """
    _require_opencv()
    if board is None:
        board = make_charuco_board()
    if card_offset_x_mm is None:
        card_offset_x_mm = CARD_ENGRAVE_OFFSET_X_MM
    if card_offset_y_mm is None:
        card_offset_y_mm = CARD_ENGRAVE_OFFSET_Y_MM
    if reference_indices is None:
        reference_indices = (0,)
    if len(detections) < CALIB_MIN_FRAMES:
        raise ValueError(
            f"Need at least {CALIB_MIN_FRAMES} detected frames; got {len(detections)}"
        )
    for ref_idx in reference_indices:
        if not (0 <= ref_idx < len(detections)):
            raise ValueError(
                f"reference index {ref_idx} out of range")

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

    # pixel → machine-mm homography. Pool corners from every reference
    # frame (card untouched between them). Each detected ChArUco
    # corner has a known machine-mm position; using N×corners across
    # M reference frames gives the findHomography RANSAC fit more
    # data and averages out per-frame detection noise.
    board_corners_mm = board.getChessboardCorners()  # Nx3 in board frame
    rows = board.getChessboardSize()[1]
    board_h_mm = rows * (board.getSquareLength() if hasattr(board, 'getSquareLength')
                          else CHARUCO_SQUARE_MM)
    src_pts_list = []
    machine_dst_pts = []
    for ref_idx in reference_indices:
        ref_corners_px, ref_ids = detections[ref_idx]
        src_pts_list.append(
            ref_corners_px.reshape(-1, 2).astype(np.float32))
        for i in ref_ids:
            bxy = board_corners_mm[int(i[0]), :2]
            mxy = _board_corner_to_machine_mm(
                bxy, card_offset_x_mm, card_offset_y_mm,
                board_h_mm=board_h_mm,
                border_mm=card_border_mm,
                label_height_mm=card_label_height_mm)
            machine_dst_pts.append(mxy)
    src_pts = np.vstack(src_pts_list)
    dst_pts = np.array(machine_dst_pts, dtype=np.float32)
    if len(src_pts) < 4:
        raise ValueError(
            "Reference frames have too few corners for homography")

    # Undistort source pixels so the homography lives in the undistorted
    # image plane (matching what undistort_frame produces at use time).
    undist_src = cv2.undistortPoints(
        src_pts.reshape(-1, 1, 2), camera_matrix, dist_coeffs, P=camera_matrix
    ).reshape(-1, 2)
    homography, _ = cv2.findHomography(undist_src, dst_pts, cv2.RANSAC, 3.0)

    return {
        'opencv_version': cv2.__version__,
        # Bumped to 2 when _board_corner_to_machine_mm started using
        # the correct Y-DOWN board frame (was treating it as Y-UP,
        # which fit a vertically-mirrored homography). pixels_to_mm
        # checks this on load and applies a Y-flip correction for
        # legacy (v1 / missing) calibrations so users don't have to
        # re-engrave + re-capture a card.
        'calibration_schema_version': 2,
        'camera_matrix': camera_matrix.tolist(),
        'dist_coeffs': dist_coeffs.tolist(),
        'rms_reprojection_error_px': float(rms),
        'image_size': list(image_size),
        'frame_count': len(detections),
        'reference_frame_count': len(reference_indices),
        # Replaces the old `homography_px_to_mm` (board-frame). The
        # name change makes it clear this homography returns machine
        # coords directly, no further transform needed.
        'homography_px_to_machine_mm': homography.tolist(),
        'card_engrave_offset_mm': [card_offset_x_mm, card_offset_y_mm],
        'card_border_mm': card_border_mm,
        'card_label_height_mm': card_label_height_mm,
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
    """Load a previously-saved calibration JSON. Returns None if not
    present. Returns None and prints a hint if it's an OLD-FORMAT
    calibration (board-frame homography from the pre-integrated
    workflow) — old calibrations are incompatible and the user must
    recalibrate."""
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        cal = json.load(f)
    if 'homography_px_to_machine_mm' not in cal:
        # Old format (had homography_px_to_mm + machine_origin_transform)
        # — the geometry meaning differs, can't be silently migrated.
        return None
    return cal


def is_legacy_calibration(path):
    """True if a calibration JSON exists at ``path`` but uses the
    pre-integrated format (old board-frame homography). UI can use
    this to show a "please recalibrate" prompt."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, encoding='utf-8') as f:
            cal = json.load(f)
    except Exception:
        return False
    return ('homography_px_to_mm' in cal
            and 'homography_px_to_machine_mm' not in cal)


def _calibration_arrays(calibration):
    if not HAS_OPENCV or not calibration:
        return None, None, None
    return (
        np.array(calibration['camera_matrix'], dtype=np.float64),
        np.array(calibration['dist_coeffs'], dtype=np.float64),
        np.array(calibration['homography_px_to_machine_mm'],
                  dtype=np.float64),
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
                          epsilon_frac=SCRAP_APPROX_EPS_FRAC,
                          threshold_bias=0, invert=False):
    """Find the largest contour in the frame and approximate it as a polygon.

    Returns a list of (x, y) pixel-coordinate tuples or ``None`` if no
    contour above ``min_area_frac`` was found. Uses Otsu thresholding,
    which works well when the scrap is significantly brighter or darker
    than the bed (typical case for leather on the Falcon's honeycomb).

    ``threshold_bias`` shifts Otsu's auto-picked threshold by N units
    on the 0–255 grayscale scale. Range typically [-80, 80]. Default 0
    = use Otsu as-is. Positive values raise the threshold (fewer
    pixels classified as scrap, less sensitive to dim material).
    Negative values lower it (more aggressive — catches material that
    barely contrasts with the bed). Tune by eye in low-contrast
    conditions where Otsu's bimodal-histogram assumption breaks down
    (uneven lighting, scrap color similar to bed, etc.).

    ``invert`` selects the threshold polarity. Default False uses
    ``THRESH_BINARY`` — "scrap brighter than bed" (leather on dark
    honeycomb). True uses ``THRESH_BINARY_INV`` — "scrap darker than
    bed" (dark felt on a light scrap board). Without inversion, dark
    scrap on a light surface produces no contour ("bed" is the largest
    region, not the scrap).
    """
    _require_opencv()
    if frame is None:
        return None
    gray = (
        cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        if len(frame.shape) == 3 else frame
    )
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    base_mode = cv2.THRESH_BINARY_INV if invert else cv2.THRESH_BINARY
    otsu_val, thresh = cv2.threshold(
        blurred, 0, 255, base_mode + cv2.THRESH_OTSU
    )
    if threshold_bias:
        # Re-threshold with the biased value (keeps Otsu as the baseline
        # so the slider centers naturally on the "automatic" pick).
        biased = max(0, min(255, int(round(otsu_val + threshold_bias))))
        _val, thresh = cv2.threshold(
            blurred, biased, 255, base_mode
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


def inset_polygon_mm(polygon_mm, inset_mm, resolution_per_mm=10):
    """Shrink a polygon inward by ``inset_mm`` on every edge.

    Used as a safety margin for camera-captured polygons: the camera's
    measurement accuracy is worst near the edges of its view, so we
    inset the captured shape by a few mm before nesting. Any pad that
    would have landed near the edge is now pulled inboard, preventing
    "circle hangs off the edge of the leather because the camera was
    off by 2mm" failures.

    Approach: rasterize the polygon at ``resolution_per_mm`` px/mm,
    erode by a circular kernel of size ``inset_mm``, re-extract the
    largest contour, convert back to mm. This handles concave
    polygons cleanly without the per-vertex math of a true geometric
    offset.

    Returns the inset polygon as a list of ``(x, y)`` mm tuples. If
    ``inset_mm <= 0`` the input is returned unchanged. If the inset
    would eliminate the polygon (too aggressive), the original is
    returned with no inset applied.
    """
    _require_opencv()
    if not polygon_mm or inset_mm <= 0:
        return polygon_mm

    xs = [p[0] for p in polygon_mm]
    ys = [p[1] for p in polygon_mm]
    xmin, ymin = min(xs), min(ys)
    xmax, ymax = max(xs), max(ys)
    w_mm = xmax - xmin
    h_mm = ymax - ymin

    res = resolution_per_mm
    margin_mm = inset_mm + 2.0
    margin_px = int(round(margin_mm * res))
    img_w = int(round(w_mm * res)) + 2 * margin_px
    img_h = int(round(h_mm * res)) + 2 * margin_px

    img = np.zeros((img_h, img_w), dtype=np.uint8)
    poly_px = np.array(
        [((p[0] - xmin) * res + margin_px,
          (p[1] - ymin) * res + margin_px) for p in polygon_mm],
        dtype=np.int32)
    cv2.fillPoly(img, [poly_px], 255)

    kernel_diameter = max(1, int(round(inset_mm * res * 2)) | 1)
    kernel = cv2.getStructuringElement(
        cv2.MORPH_ELLIPSE, (kernel_diameter, kernel_diameter))
    eroded = cv2.erode(img, kernel)

    contours, _hier = cv2.findContours(
        eroded, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return polygon_mm  # eroded away — too aggressive, fall back
    largest = max(contours, key=cv2.contourArea)
    if cv2.contourArea(largest) < 100:
        return polygon_mm
    epsilon = 0.005 * cv2.arcLength(largest, True)
    approx = cv2.approxPolyDP(largest, epsilon, True)
    result = []
    for pt in approx:
        px, py = pt[0]
        x = (px - margin_px) / res + xmin
        y = (py - margin_px) / res + ymin
        result.append((float(x), float(y)))
    return result


def pixels_to_mm(polygon_px, calibration):
    """Apply the pixel→machine-mm homography to convert a pixel-coord
    polygon into MACHINE-MM coordinates.

    The integrated calibration baked the camera-to-machine transform
    into the homography itself (using the engrave card's known machine
    position as the geometric reference), so this returns machine
    coords directly — no separate auto-framing step needed.

    Schema-1 (legacy) calibrations were fit with a vertically-flipped
    set of destination points because ``_board_corner_to_machine_mm``
    treated OpenCV's Y-DOWN board frame as Y-UP. The homography returns
    machine coords that are MIRRORED in Y around the card's center line.
    We detect those by the missing ``calibration_schema_version`` field
    and undo the flip post-hoc so the user doesn't have to re-engrave +
    re-capture a card.
    """
    _require_opencv()
    if not calibration or not polygon_px:
        return polygon_px
    _cm, _dc, homography = _calibration_arrays(calibration)
    pts = np.array(polygon_px, dtype=np.float64).reshape(-1, 1, 2)
    transformed = cv2.perspectiveTransform(pts, homography)
    out = [(float(p[0][0]), float(p[0][1])) for p in transformed]

    if int(calibration.get('calibration_schema_version', 1)) < 2:
        # Recover the right Y values by mirroring around the card's
        # vertical center line in the schema-1 coord system. In the
        # broken function:
        #   broken_y = offset_y + border + label_h + by
        # In the correct one:
        #   correct_y = offset_y + border + board_h - by
        # Adding them:  broken_y + correct_y = 2*offset_y + 2*border
        #                                       + label_h + board_h
        # So: correct_y = K - broken_y, where K is the constant above.
        try:
            off = calibration.get('card_engrave_offset_mm') or [0.0, 0.0]
            offset_y = float(off[1])
            border = float(calibration.get('card_border_mm', 10.0))
            label_h = float(calibration.get('card_label_height_mm', 8.0))
            board_info = calibration.get('board') or {}
            board_h = (float(board_info.get('rows', CHARUCO_ROWS))
                       * float(board_info.get('square_mm', CHARUCO_SQUARE_MM)))
            k = 2 * offset_y + 2 * border + label_h + board_h
            out = [(x, k - y) for (x, y) in out]
        except (TypeError, ValueError, KeyError):
            pass
    return out



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
