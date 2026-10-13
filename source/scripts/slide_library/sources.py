"""Where the inputs and outputs live, and what the add-in's TypeScript says: the slide grid
(src/brand/slideGrid.ts), the palette (src/brand/tokens.ts) and the standard disclaimer (src/docx/sections.ts).
They are read when the script runs, so the template keeps to the add-in's grid, colours and wording.
"""
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
PPTX_OUT = ROOT / "public" / "templates" / "bwb-slides.pptx"
DECK_OUT = ROOT / "public" / "templates" / "bwb-deck.pptx"
SLIDES_OUT = ROOT / "public" / "assets" / "slides"
IMAGES = ROOT / "public" / "assets" / "images"
BRAND = ROOT / "public" / "assets" / "brand"
LUCIDE = ROOT / "node_modules" / "lucide-static" / "icons"
FIXED_TIME = (2026, 10, 13, 9, 0, 0)
FIXED_ISO = "2026-10-13T09:00:00Z"
THUMB_WIDTH = 480


class Grid:
    """The BwB slide grid in points (src/brand/slideGrid.ts)."""

    def __init__(self, ts: str):
        num = r"([\d.]+)"

        def find(pattern: str, what: str) -> list:
            m = re.search(pattern, ts)
            if not m:
                sys.exit(f"Could not read {what} from src/brand/slideGrid.ts")
            return [float(v) for v in m.groups()]

        def area(name: str) -> tuple:
            return tuple(find(rf"export const {name}: Area = \{{ left: {num}, top: {num}, width: {num}, height: {num} \}}",
                              name))

        self.slide, self.title, self.content = area("SLIDE"), area("TITLE_AREA"), area("CONTENT_AREA")
        self.source, self.footer = area("SOURCE_AREA"), area("FOOTER_AREA")
        left, right, top, bottom = find(rf"export const MARGINS = \{{ left: {num}, right: {num}, top: {num}, "
                                        rf"bottom: {num} \}}", "MARGINS")
        self.margin_left, self.margin_right, self.margin_top, self.margin_bottom = left, right, top, bottom
        self.strip = find(rf"export const STRIP_HEIGHT = {num}", "STRIP_HEIGHT")[0]
        self.columns = int(find(rf"export const COLUMNS = {num}", "COLUMNS")[0])
        self.gutter = find(rf"export const GUTTER = {num}", "GUTTER")[0]
        self.width, self.height = self.slide[2], self.slide[3]
        safe = self.width - left - right
        self.col = (safe - (self.columns - 1) * self.gutter) / self.columns

    def x(self, column: int) -> float:
        """Left edge of a column, counting from 0."""
        return self.margin_left + column * (self.col + self.gutter)

    def w(self, columns: int) -> float:
        """Width of a span of whole columns and the gutters between them."""
        return columns * self.col + (columns - 1) * self.gutter


def read_colours() -> dict:
    ts = (ROOT / "src" / "brand" / "tokens.ts").read_text(encoding="utf-8")
    block = re.search(r"export const COLOURS = \{(.*?)\} as const", ts, re.S)
    if not block:
        sys.exit("Could not read COLOURS from src/brand/tokens.ts")
    return {k: v.upper() for k, v in re.findall(r"(\w+): '([0-9A-Fa-f]{6})'", block.group(1))}


G = Grid((ROOT / "src" / "brand" / "slideGrid.ts").read_text(encoding="utf-8"))
C = read_colours()
NAVY, GREEN, LIME, CHARCOAL = C["navy"], C["green"], C["lime"], C["charcoal"]
GREEN50, GREY, OFFWHITE, WHITE, BLACK, GREY20 = (C["green50"], C["lightGrey"], C["offWhite"], C["white"], C["black"],
                                                 C["charcoal20"])
HEAD, BODY = "Candara", "Aptos"
CT = G.content[1]                    # content top
CB = G.content[1] + G.content[3]     # content bottom
CH = G.content[3]
RIGHT = G.width - G.margin_right


def read_disclaimer() -> tuple:
    """The standard disclaimer and the FCA note, from src/docx/sections.ts, so slides and documents match."""
    ts = (ROOT / "src" / "docx" / "sections.ts").read_text(encoding="utf-8")
    block = re.search(r"export const DEFAULT_DISCLAIMER: readonly string\[\] = \[(.*?)\n\]", ts, re.S)
    note = re.search(r"export const DISCLAIMER_NOTE = '((?:[^'\\]|\\.)*)'", ts)
    if not block or not note:
        sys.exit("Could not read DEFAULT_DISCLAIMER and DISCLAIMER_NOTE from src/docx/sections.ts")

    def unescape(text: str) -> str:
        text = re.sub(r"\\u\{([0-9A-Fa-f]+)\}", lambda m: chr(int(m.group(1), 16)), text)
        return text.replace("\\'", "'")

    paras = [unescape(t) for t in re.findall(r"'((?:[^'\\]|\\.)*)'", block.group(1))]
    return paras, unescape(note.group(1))
