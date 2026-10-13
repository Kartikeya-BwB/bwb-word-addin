"""Thumbnails: renders the library with LibreOffice and pdftoppm, then writes 480 px palette PNGs."""
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

from .catalogue import SLIDES
from .sources import SLIDES_OUT, THUMB_WIDTH


def render(pptx: Path, preview: Path = None) -> int:
    """Renders every slide with LibreOffice, then writes the palette PNG thumbnails. Returns their total size."""
    for tool in ("soffice", "pdftoppm"):
        if not shutil.which(tool):
            sys.exit(f"{tool} is not installed; use --no-thumbs to skip the thumbnails")
    thumbs = SLIDES_OUT / "thumbs"
    thumbs.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="bwb_slides_") as tmp:
        tmp = Path(tmp)
        subprocess.run(["soffice", f"-env:UserInstallation={(tmp / 'profile').as_uri()}", "--headless",
                        "--convert-to", "pdf", "--outdir", str(tmp), str(pptx)], check=True, capture_output=True,
                       timeout=600)
        pdf = tmp / f"{pptx.stem}.pdf"
        if not pdf.exists():
            sys.exit("LibreOffice did not produce a PDF")
        subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(THUMB_WIDTH), "-scale-to-y", "-1", str(pdf),
                        str(tmp / "page")], check=True)
        pages = sorted(tmp.glob("page-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))
        if len(pages) != len(SLIDES):
            sys.exit(f"Rendered {len(pages)} pages for {len(SLIDES)} slides")
        keep = {f"{s.id}.png" for s in SLIDES}
        for old in thumbs.glob("*.png"):
            if old.name not in keep:
                old.unlink()
        total = 0
        for spec, page in zip(SLIDES, pages):
            with Image.open(page) as im:
                small = im.convert("RGB").quantize(colors=128, method=Image.Quantize.MEDIANCUT,
                                                   dither=Image.Dither.FLOYDSTEINBERG)
            path = thumbs / f"{spec.id}.png"
            small.save(path, optimize=True)
            total += path.stat().st_size
        if preview:
            preview.mkdir(parents=True, exist_ok=True)
            for old in preview.glob("slide-*.png"):
                old.unlink()
            subprocess.run(["pdftoppm", "-png", "-r", "144", str(pdf), str(preview / "slide")], check=True)
    return total
