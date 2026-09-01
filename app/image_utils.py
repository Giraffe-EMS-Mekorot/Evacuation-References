"""Shared image loading/downscaling/encoding for the Anthropic vision API.

Split out of extractor.py (2026-09) so app.examples_library can encode a
few-shot example image exactly the same way a real certificate page is
encoded, without examples_library importing from extractor.py (which
itself needs to import FROM examples_library to prepend few-shot messages -
that would be a circular import). Neither this module nor fields.py imports
from extractor.py or examples_library.py, so both of those can safely
import from here.
"""
import io

from PIL import Image

# Adaptive downscaling ladder: try the highest resolution first and only step
# down as far as actually needed to clear MAX_IMAGE_BYTES, instead of
# shrinking every page to one fixed size regardless of how compressible it
# is. 3000 is the top of the "2600-3000px" range from the original spec -
# starting at the ceiling, not the middle, actually maximizes resolution
# per-page as intended; a smaller/plainer scan may well clear the limit on
# the first try and keep its full 3000px, while a noisier one steps down
# only as needed.
_ADAPTIVE_DIMENSIONS = [3000, 2600, 2200, 1800]
_JPEG_QUALITY = 90
# Render PDF pages at 300 DPI before downscaling - high enough that a
# standard page (A4/Letter) actually renders past 3000px on its long side,
# so the adaptive ladder above has real headroom to pick from; at the ~200
# DPI this used to render at, a standard page tops out under 2400px and the
# "try 3000 first" step would never have anything to bite on.
PDF_RENDER_SCALE = 300 / 72
# Safety net matching Anthropic's documented per-image limit. In practice a
# page that's stepped all the way down to 1800px at quality 90 should never
# get close to this - this guards the case where it somehow still does (e.g.
# a pathologically noisy scan) rather than sending it and hoping.
MAX_IMAGE_BYTES = 5 * 1024 * 1024


def downscale_and_encode(image: Image.Image) -> bytes:
    """Encodes `image` as a JPEG, trying _ADAPTIVE_DIMENSIONS from highest to
    lowest and stopping at the first one that clears MAX_IMAGE_BYTES - keeps
    each specific page at the best resolution *it* can afford, rather than
    shrinking every page to the same fixed size up front regardless of how
    compressible it actually turns out to be. Never upscales past the
    image's own native size. If every step is still oversized, returns the
    smallest (last) attempt anyway - the caller checks the size and turns
    that into a clear per-page failure instead of sending it.
    """
    if image.mode not in ("RGB", "L"):
        image = image.convert("RGB")
    width, height = image.size
    longest = max(width, height)

    encoded = b""
    for max_dimension in _ADAPTIVE_DIMENSIONS:
        if longest > max_dimension:
            ratio = max_dimension / longest
            resized = image.resize(
                (max(1, round(width * ratio)), max(1, round(height * ratio))), Image.LANCZOS
            )
        else:
            resized = image
        buffer = io.BytesIO()
        resized.save(buffer, format="JPEG", quality=_JPEG_QUALITY)
        encoded = buffer.getvalue()
        if len(encoded) <= MAX_IMAGE_BYTES:
            return encoded
    return encoded


def encode_image_block(jpeg_bytes: bytes) -> dict:
    import base64

    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": "image/jpeg",
            "data": base64.standard_b64encode(jpeg_bytes).decode("utf-8"),
        },
    }
