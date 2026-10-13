"""Builds the BwB PowerPoint template, its slide library, the starter deck and the library index the add-in reads.

Usage:
  python3 scripts/make-slide-library.py [--no-thumbs] [--preview DIR]

Writes, relative to source/:
  public/templates/bwb-slides.pptx       a 16:9 BwB slide master (theme, seven layouts) and the library slides
  public/templates/bwb-deck.pptx         the same master with a blank title slide and a blank title and content
                                         slide: the starting point for a new BwB deck
  public/assets/slides/index.json        the template, the starter deck (starterDeck) and one entry per slide: id,
                                         title, category, description, tags, slideId, thumb
  public/assets/slides/thumbs/<id>.png   480 px thumbnails, rendered with LibreOffice and pdftoppm (palette PNG)

The code is in scripts/slide_library/: sources (the grid, palette and disclaimer, read from the add-in's
TypeScript), drawing (DrawingML helpers), assets, master (theme, master and layouts), elements and the slides_*
modules (the library slides), chart, catalogue, starter, build (assembly and the index) and thumbs.

Positions come from src/brand/slideGrid.ts and colours from src/brand/tokens.ts, read when the script runs, so the
template keeps to the grid the add-in's alignment tools use. Photos come from public/assets/images, the logos from
public/assets/brand and the icons from lucide-static (rendered with cairosvg, or with @resvg/resvg-js when cairosvg is
not installed). Icons go in as SVG with a PNG fallback, so PowerPoint keeps them sharp and recolourable.

Fonts are named, never embedded: Candara and Aptos must be installed for the thumbnails to be faithful.

Every slide has a fixed sldId and a p14:creationId derived from its library id, so the add-in's slide ids
("<sldId>#<creationId>") stay the same from one build to the next. Both files are written with fixed timestamps,
the workbook behind the native chart included, so rebuilding without changes gives the same bytes.
"""
import argparse
import sys
from pathlib import Path

sys.dont_write_bytecode = True  # keep __pycache__ out of scripts/slide_library

from slide_library.build import build, check_specs, write_index
from slide_library.catalogue import SLIDES
from slide_library.sources import DECK_OUT, PPTX_OUT, ROOT, SLIDES_OUT
from slide_library.starter import STARTER
from slide_library.thumbs import render


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-thumbs", action="store_true", help="skip the thumbnails (they need LibreOffice and pdftoppm)")
    ap.add_argument("--preview", type=Path, help="also write a 1920 px render of every slide to this folder")
    args = ap.parse_args()
    check_specs()
    PPTX_OUT.parent.mkdir(parents=True, exist_ok=True)
    SLIDES_OUT.mkdir(parents=True, exist_ok=True)
    PPTX_OUT.write_bytes(build(SLIDES, "BwB slide template and library"))
    DECK_OUT.write_bytes(build(STARTER, ""))
    index = write_index()
    print(f"{len(SLIDES)} slides, {PPTX_OUT.stat().st_size / 1e3:.0f} KB: {PPTX_OUT.relative_to(ROOT)}")
    print(f"{len(STARTER)} slides, {DECK_OUT.stat().st_size / 1e3:.0f} KB: {DECK_OUT.relative_to(ROOT)}")
    print(f"Index: {index.relative_to(ROOT)}")
    if not args.no_thumbs:
        total = render(PPTX_OUT, args.preview)
        print(f"Thumbnails: {total / 1e3:.0f} KB in {(SLIDES_OUT / 'thumbs').relative_to(ROOT)}")


if __name__ == "__main__":
    main()
