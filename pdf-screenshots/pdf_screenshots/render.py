"""PDF page rendering and JPEG output using pypdfium2 and Pillow.

Windows-only tool. pypdfium2 ships as a self-contained pip wheel with no
external binaries (unlike pdf2image/poppler) and is permissively licensed
(Apache 2.0 / BSD-3).
"""

from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image

# When True, the rendered page keeps its native aspect ratio and the output is
# letterboxed/fitted into the requested frame. When False (the default), the
# page is stretched to exactly the requested width x height, which slightly
# distorts a standard Letter/A4 page but produces a consistent, screenshot-like
# frame for every page. Flip this constant to change the behavior without
# editing any logic below.
PRESERVE_ASPECT = False


def get_page_count(pdf_path: Path) -> int:
    """Return the number of pages in the PDF."""
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        return len(pdf)
    finally:
        pdf.close()


def _render_page_image(page, width: int, height: int) -> Image.Image:
    """Render a single pdfium page to a Pillow RGB image at width x height."""
    page_width_pt, page_height_pt = page.get_size()

    # Choose a render scale so the rendered bitmap is at least as large as the
    # requested output in both dimensions before we downscale. Rendering larger
    # than needed and then downsampling with LANCZOS keeps text sharp.
    render_scale = max(width / page_width_pt, height / page_height_pt) * 1.0

    bitmap = page.render(scale=render_scale)
    pil_image = bitmap.to_pil()

    # JPEG cannot encode an alpha channel; pypdfium2 bitmaps may carry one.
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")

    if PRESERVE_ASPECT:
        # Fit the page within the requested frame without distortion, padding
        # the remainder with white so every output image is exactly the same
        # size.
        fitted = pil_image.copy()
        fitted.thumbnail((width, height), Image.LANCZOS)
        canvas = Image.new("RGB", (width, height), (255, 255, 255))
        offset = ((width - fitted.width) // 2, (height - fitted.height) // 2)
        canvas.paste(fitted, offset)
        return canvas

    # Stretch to the exact requested dimensions (intentional slight distortion).
    return pil_image.resize((width, height), Image.LANCZOS)


def render_pdf_to_jpegs(
    pdf_path: Path,
    output_dir: Path,
    width: int,
    height: int,
    quality: int,
    progress_callback=None,
):
    """Render every page of ``pdf_path`` to a JPEG in ``output_dir``.

    Files are named ``page_01.jpg``, ``page_02.jpg``, ... zero-padded based on
    the total page count.

    ``progress_callback`` is called as ``callback(page_number, total, filename)``
    after each page is saved.

    Returns the list of written Path objects.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(pdf_path))
    written = []
    try:
        n_pages = len(pdf)
        pad = len(str(n_pages))

        for i in range(n_pages):
            page = pdf[i]
            image = _render_page_image(page, width, height)

            filename = f"page_{str(i + 1).zfill(pad)}.jpg"
            out_path = output_dir / filename

            _save_with_retry(image, out_path, quality)
            written.append(out_path)

            if progress_callback is not None:
                progress_callback(i + 1, n_pages, filename)
    finally:
        pdf.close()

    return written


def _save_with_retry(image: Image.Image, out_path: Path, quality: int):
    """Save the image as JPEG, retrying on PermissionError.

    A PermissionError typically means the target file is open in another
    program. Rather than aborting a partially finished run, prompt the user to
    close it and retry.
    """
    while True:
        try:
            image.save(out_path, format="JPEG", quality=quality, optimize=True)
            return
        except PermissionError:
            input(
                f"Cannot write {out_path.name} — the file may be open in "
                f"another program. Close it and press Enter to retry."
            )
