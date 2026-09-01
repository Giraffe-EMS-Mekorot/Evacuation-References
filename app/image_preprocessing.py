"""Conditional image preprocessing (2026-09) - deskew, contrast enhancement,
denoising, and mild sharpening - applied to a certificate page BEFORE it's
downscaled/encoded and sent to Claude (see image_utils.downscale_and_encode
and extractor.py's _extract_page), to counter common real-world scan/photo
quality problems: crooked photos, washed-out or shadowed lighting, and
light motion blur/grain.

**Every correction is conditional, never blanket.** Each one is gated
behind its own "is this actually needed" check (see preprocess_image's own
body, and each correction's threshold constants above it) computed from
the image itself - an already-good image is returned
byte-for-byte unmodified (well, pixel-for-pixel: no correction function is
even called), never "improved" into something worse. This is the whole
point, not an optimization: per the spec this implements, blanket
processing risks damaging already-good scans for no benefit.

Every metric is computed on a small, resolution-normalized grayscale copy
(see _analysis_copy) - fast, and keeps the threshold constants below
meaningful regardless of whether the real image came from a 300 DPI PDF
render (~2500x3500px) or a phone photo of arbitrary resolution. The actual
correction, when one fires, is always applied to the FULL-resolution image,
never the small analysis copy - preprocessing must not lower the effective
resolution ceiling image_utils.downscale_and_encode's own adaptive ladder
already carefully maximizes.

**Deblurring vs. denoising - a real distinction, not a technicality.** True
motion-blur *correction* (blind deconvolution) is a hard, failure-prone
problem in its own right and risks inventing detail that was never there -
out of scope here. What this module actually does for a soft/blurry image
is a mild unsharp-mask sharpening pass (crisps up existing edges, invents
nothing) - a real but partial mitigation, not a fix. Said honestly in the
notes text (see _describe_corrections) so nobody reads "טיפול בטשטוש" as a
stronger claim than what actually happened.
"""
from typing import List, Optional, Tuple

import cv2
import numpy as np
from PIL import Image

# --- Analysis copy ----------------------------------------------------------
# All "is this needed" metrics are computed against a copy resized to this
# long-side dimension - small enough to keep every check fast (well under
# 100ms combined - see the timing note in the PR description), and fixes a
# single, consistent scale for the threshold constants below regardless of
# the real image's own resolution.
_ANALYSIS_MAX_DIMENSION = 1200


def _analysis_copy(gray_full: np.ndarray) -> np.ndarray:
    height, width = gray_full.shape[:2]
    longest = max(height, width)
    if longest <= _ANALYSIS_MAX_DIMENSION:
        return gray_full
    scale = _ANALYSIS_MAX_DIMENSION / longest
    size = (max(1, round(width * scale)), max(1, round(height * scale)))
    return cv2.resize(gray_full, size, interpolation=cv2.INTER_AREA)


# --- Deskew -------------------------------------------------------------
# Below this, a detected tilt is imperceptible and not worth the
# interpolation softening a rotation introduces - matches the spec's "only
# correct when actually needed".
_DESKEW_APPLY_DEGREES = 0.6
# A correction at or above this is what actually gets mentioned in
# record["notes"] (see _describe_corrections) - "significant", per the
# spec, not just "technically applied".
_DESKEW_NOTE_DEGREES = 2.0
# A "detected" angle beyond this is far more likely a bad detection (e.g. a
# sparse or table-free page with no reliable line signal) than a real
# photo taken this crooked - skipped rather than risking a wrong rotation
# that makes a fine image worse.
_DESKEW_MAX_CORRECTABLE_DEGREES = 20.0


def _estimate_skew_angle(gray_small: np.ndarray) -> Optional[float]:
    """Estimates rotation angle (degrees; positive = counter-clockwise tilt
    to correct) via the dominant angle of near-horizontal lines found by a
    Hough transform on Canny edges - standard document-deskew technique,
    and more robust than a text-blob-bounding-box approach (minAreaRect on
    thresholded text) for these certificates specifically, since most carry
    real table/ruling lines to lock onto even when handwritten text is
    sparse or irregular. Returns None if too few line segments were found
    to trust an estimate (a blank-ish or very sparse page) - "no signal"
    must never be treated as "confirmed straight".
    """
    edges = cv2.Canny(gray_small, 50, 150, apertureSize=3)
    lines = cv2.HoughLinesP(
        edges, 1, np.pi / 360, threshold=80, minLineLength=gray_small.shape[1] // 4, maxLineGap=10
    )
    if lines is None or len(lines) < 3:
        return None

    angles = []
    # .reshape(-1, 4) normalizes across cv2 versions - HoughLinesP has
    # returned (N, 1, 4) historically but (N, 4) directly as of at least
    # opencv 5.0 (confirmed against the installed version); indexing
    # lines[:, 0] to get 4-tuples only works for the former and silently
    # raises (caught by preprocess_image's blanket try/except, which is why
    # this needs an explicit test to catch, not just "no exception surfaced").
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        angle = np.degrees(np.arctan2(y2 - y1, x2 - x1))
        # Keep only near-horizontal segments (within 45 of horizontal) -
        # a near-vertical line (e.g. a table's side border) reflects the
        # same tilt but the arctan2 wrap makes averaging it in directly
        # wrong; simplest robust fix is to only use the horizontal set,
        # which every one of these certificates has plenty of (row rulings,
        # underlines, table borders).
        if abs(angle) <= 45:
            angles.append(angle)
    if len(angles) < 3:
        return None
    return float(np.median(angles))


def _deskew(image_bgr: np.ndarray, angle: float) -> np.ndarray:
    """Rotates `image_bgr` to undo a detected tilt of `angle` degrees (as
    returned by _estimate_skew_angle - passed straight through, unnegated)
    around its center, expanding the canvas so no content is cropped at the
    corners (the standard "rotate without cropping" recipe: recompute the
    bounding box for the rotated rectangle, then shift the rotation
    matrix's translation to re-center the result in it). New corners are
    filled white, not black - these are document scans on a light
    background, and a black border reads as more of a corruption to the
    model than a white one blending into the page.

    Verified empirically, not just reasoned about: passing angle straight
    through to cv2.getRotationMatrix2D (NOT negated) is what actually
    drives the residual detected tilt to ~0 on a real test certificate - an
    earlier version of this function negated it, which silently DOUBLED a
    real ~3.4-degree tilt into ~6.5 degrees instead of correcting it. See
    the PR description for the before/after numbers this caught.
    """
    height, width = image_bgr.shape[:2]
    center = (width / 2, height / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)

    cos, sin = abs(matrix[0, 0]), abs(matrix[0, 1])
    new_width = int(height * sin + width * cos)
    new_height = int(height * cos + width * sin)
    matrix[0, 2] += (new_width / 2) - center[0]
    matrix[1, 2] += (new_height / 2) - center[1]

    return cv2.warpAffine(
        image_bgr, matrix, (new_width, new_height), flags=cv2.INTER_CUBIC, borderValue=(255, 255, 255)
    )


# --- Contrast enhancement (CLAHE) ----------------------------------------
# A percentile-range metric (p99 - p1, robust to a few outlier pixels),
# NOT plain standard deviation - measured directly against real certificate
# scans while building this feature and rejected: a typical certificate is
# mostly blank white page with sparse text, so whole-image std is dominated
# by that blank majority and sits low (~20) even on a perfectly clear scan,
# barely distinguishable from a genuinely washed-out one (~7-14) - too
# narrow a margin to threshold on reliably. The 1st-to-99th-percentile
# INTENSITY RANGE instead directly captures what "washed out" really means
# (true blacks and true whites both pulled toward middle gray) regardless
# of how much of the page is blank: measured ~138 on a clear scan, ~48 on a
# synthetically washed-out version of the same page - a wide, reliable gap.
_LOW_CONTRAST_RANGE = 100.0
_VERY_LOW_CONTRAST_RANGE = 65.0
# A near-blank page (p1 already near-white) has a low RANGE too (nothing to
# spread), but for the opposite reason - there's no real dark content to
# recover, not a washed-out version of real content. Only enhance when the
# darker end is actually dark enough to be real content worth recovering.
_MAX_P1_FOR_CLAHE = 200.0


def _contrast_range(gray_small: np.ndarray) -> Tuple[float, float]:
    """Returns (p99-p1 intensity range, p1) - see _LOW_CONTRAST_RANGE above
    for why this replaced a plain standard-deviation metric.
    """
    p1, p99 = np.percentile(gray_small, [1, 99])
    return float(p99 - p1), float(p1)


def _apply_clahe(image_bgr: np.ndarray) -> np.ndarray:
    """Contrast-Limited Adaptive Histogram Equalization on the L (lightness)
    channel of LAB color space only - boosts local contrast without
    distorting color balance the way running CLAHE per-RGB-channel would.
    clipLimit=2.0 is a deliberately modest cap (higher values amplify noise
    and produce a harsh, over-processed look) - per the spec's "without
    exaggerating".
    """
    lab = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2LAB)
    l_channel, a_channel, b_channel = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    l_enhanced = clahe.apply(l_channel)
    enhanced_lab = cv2.merge((l_enhanced, a_channel, b_channel))
    return cv2.cvtColor(enhanced_lab, cv2.COLOR_LAB2BGR)


# --- Denoising ------------------------------------------------------------
# Noise proxy: how much the image changes under a small median blur, MASKED
# to only "flat" (low-gradient) pixels - measured directly while building
# this feature and revised for a real problem: unmasked, sharp text edges
# survive a median filter about as poorly as actual sensor noise does (a
# clean, dense-text certificate scored ~11-12 on the unmasked version - not
# meaningfully lower than a copy with real gaussian noise deliberately
# added, ~11), because the metric couldn't tell "this pixel changed because
# there's a text edge here" from "this pixel changed because there's
# noise". Restricting the residual to pixels where the Laplacian is already
# small (no real edge nearby) fixes that: same clean scan then measured
# ~0.5, the noise-added version ~1.0 - about 2x, a real and reliable gap.
_NOISE_FLAT_GRADIENT_THRESHOLD = 4.0
_NOISE_APPLY_THRESHOLD = 0.7
_NOISE_NOTE_THRESHOLD = 0.85


def _noise_level(gray_small: np.ndarray) -> float:
    denoised = cv2.medianBlur(gray_small, 3)
    residual = gray_small.astype(np.int16) - denoised.astype(np.int16)
    gradient = cv2.Laplacian(gray_small, cv2.CV_64F)
    flat_mask = np.abs(gradient) < _NOISE_FLAT_GRADIENT_THRESHOLD
    if flat_mask.sum() < 500:  # not enough flat area to trust an estimate
        return 0.0
    return float(np.std(residual[flat_mask]))


def _apply_denoise(image_bgr: np.ndarray) -> np.ndarray:
    """A mild edge-preserving bilateral filter - smooths flat/noisy regions
    while leaving real edges (text strokes, digits) largely intact, per the
    spec's "without harming small digits". Chosen over
    cv2.fastNlMeansDenoisingColored specifically for speed: measured ~5.3s
    for fastNlMeansDenoisingColored against a real ~2500x3500 certificate
    scan versus ~0.1s for this bilateral filter at a visually comparable
    mild strength - see the PR description's timing numbers. d=7 (neighborhood
    diameter) and sigmaColor/sigmaSpace=50 are deliberately modest, not the
    stronger values (d=9, sigma=75) common in general photo-denoising code.
    """
    return cv2.bilateralFilter(image_bgr, d=7, sigmaColor=50, sigmaSpace=50)


# --- Mild sharpening (partial blur mitigation) -----------------------------
# Laplacian variance on the analysis copy - a standard blur proxy (a crisp
# image has strong, varied second-derivative response at edges; a blurry
# one is smoother, so lower variance). Below this, worth a mild sharpening
# pass; this is NOT a blur/deblur measurement in any absolute sense, only a
# relative "soft vs. crisp" signal calibrated for the 1200px analysis copy.
_BLUR_APPLY_VARIANCE = 150.0
_BLUR_NOTE_VARIANCE = 80.0


def _laplacian_variance(gray_small: np.ndarray) -> float:
    return float(cv2.Laplacian(gray_small, cv2.CV_64F).var())


def _apply_unsharp(image_bgr: np.ndarray) -> np.ndarray:
    """A mild unsharp mask (subtract a blurred copy, weighted, from the
    original) - crisps up existing edges. Deliberately NOT deconvolution -
    see this module's own docstring on why true deblurring is out of scope.
    A gentle 1.3x/-0.3x blend, not the more aggressive 1.5/-0.5 sometimes
    seen in general photo-sharpening code - again, small digits are the
    thing most at risk of a sharpening pass turning soft-but-legible into a
    haloed, harder-to-read mess.
    """
    blurred = cv2.GaussianBlur(image_bgr, (0, 0), sigmaX=2.0)
    return cv2.addWeighted(image_bgr, 1.3, blurred, -0.3, 0)


def _pil_to_bgr(image: Image.Image) -> np.ndarray:
    if image.mode != "RGB":
        image = image.convert("RGB")
    return cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)


def _bgr_to_pil(image_bgr: np.ndarray) -> Image.Image:
    return Image.fromarray(cv2.cvtColor(image_bgr, cv2.COLOR_BGR2RGB))


def _describe_corrections(
    skew_angle: Optional[float], contrast_range: Optional[float], noise_level: Optional[float], blur_variance: Optional[float]
) -> List[str]:
    """Turns whichever corrections crossed their "significant" (NOTE)
    threshold into short Hebrew fragments for record["notes"] (see
    extractor.py's _flag_preprocessing) - only the ones actually significant
    enough to mention, per the spec; a correction that fired at the lower
    "apply" threshold but stayed under its "note" threshold is applied
    silently, same as this whole module's "fix the small stuff quietly,
    flag the real problems" design.
    """
    fragments = []
    if skew_angle is not None and abs(skew_angle) >= _DESKEW_NOTE_DEGREES:
        fragments.append(f"יושרה זווית ({abs(skew_angle):.1f}°)")
    if contrast_range is not None and contrast_range < _VERY_LOW_CONTRAST_RANGE:
        fragments.append("שופרה ניגודיות")
    if noise_level is not None and noise_level >= _NOISE_NOTE_THRESHOLD:
        fragments.append("הופחת רעש")
    if blur_variance is not None and blur_variance < _BLUR_NOTE_VARIANCE:
        fragments.append("חודדה תמונה מטושטשת")
    return fragments


def preprocess_image(image: Image.Image) -> Tuple[Image.Image, List[str]]:
    """Runs deskew -> denoise -> contrast enhancement -> mild sharpening on
    `image`, each step conditional on its own "is this needed" check (see
    the module docstring), and returns (possibly-corrected image, list of
    short Hebrew fragments describing whichever corrections were
    significant enough to flag - see _describe_corrections). An empty list
    means either nothing needed correcting, or only minor/below-the-"note"-
    threshold correction happened - both cases where the caller
    (extractor.py) should add no note.

    Order matters: deskew first (geometric correction before anything else,
    and its own interpolation is the only step that itself slightly softens
    the image, so doing it first means every later step reacts to the truly
    final geometry) -> denoise (before contrast enhancement, so CLAHE isn't
    locally amplifying noise) -> CLAHE -> sharpen last (after denoising, so
    it isn't re-sharpening noise speckles back into visibility).

    Never raises - any unexpected failure in a single correction step falls
    back to the image as it was going into that step (see the try/except
    around each stage), since a bug in this quality-improvement feature
    must never be why a whole certificate fails to extract at all.
    """
    image_bgr = _pil_to_bgr(image)
    gray_full = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
    gray_small = _analysis_copy(gray_full)

    applied_skew_angle = None
    applied_contrast_range = None
    applied_noise_level = None
    applied_blur_variance = None

    try:
        skew_angle = _estimate_skew_angle(gray_small)
        if skew_angle is not None and _DESKEW_APPLY_DEGREES <= abs(skew_angle) <= _DESKEW_MAX_CORRECTABLE_DEGREES:
            image_bgr = _deskew(image_bgr, skew_angle)
            applied_skew_angle = skew_angle
            # Geometry changed - every later metric must be recomputed
            # against the corrected image, not the original tilted one.
            gray_full = cv2.cvtColor(image_bgr, cv2.COLOR_BGR2GRAY)
            gray_small = _analysis_copy(gray_full)
    except Exception:
        pass  # deskew is best-effort - never let a bad estimate break extraction

    try:
        noise_level = _noise_level(gray_small)
        if noise_level >= _NOISE_APPLY_THRESHOLD:
            image_bgr = _apply_denoise(image_bgr)
            applied_noise_level = noise_level
    except Exception:
        pass

    try:
        contrast_range, p1 = _contrast_range(gray_small)
        if contrast_range < _LOW_CONTRAST_RANGE and p1 < _MAX_P1_FOR_CLAHE:
            image_bgr = _apply_clahe(image_bgr)
            applied_contrast_range = contrast_range
    except Exception:
        pass

    try:
        blur_variance = _laplacian_variance(gray_small)
        if blur_variance < _BLUR_APPLY_VARIANCE:
            image_bgr = _apply_unsharp(image_bgr)
            applied_blur_variance = blur_variance
    except Exception:
        pass

    notes = _describe_corrections(applied_skew_angle, applied_contrast_range, applied_noise_level, applied_blur_variance)
    return _bgr_to_pil(image_bgr), notes
