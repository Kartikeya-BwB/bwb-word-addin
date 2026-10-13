"""Brand assets for the slides: logos trimmed to their visible edges, library photos cropped to their frames,
Lucide icons recoloured, and the native chart, each added to a slide part with its relationship.
"""
import io
import re
import subprocess
from pathlib import Path

from PIL import Image
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part

from .sources import IMAGES, LUCIDE, ROOT


def trimmed_png(path: Path) -> tuple:
    with Image.open(path) as im:
        im = im.convert("RGBA")
        im = im.crop(im.getchannel("A").getbbox())
        buf = io.BytesIO()
        im.save(buf, "PNG", optimize=True)
        return buf.getvalue(), im.size[0] / im.size[1]


def photo(image_id: str, aspect: float, cx: float, cy: float, scale: float, width_px: int) -> bytes:
    """Crops a library photo to the frame's aspect ratio around (cx, cy), given as fractions of the photo.

    scale is the crop's width as a fraction of the photo's width (capped by the photo's height)."""
    with Image.open(IMAGES / f"{image_id}.jpg") as im:
        w, h = im.size
        cw = min(w * scale, h * aspect, w)
        ch = cw / aspect
        left = min(max(cx * w - cw / 2, 0), w - cw)
        top = min(max(cy * h - ch / 2, 0), h - ch)
        crop = im.crop((round(left), round(top), round(left + cw), round(top + ch)))
        crop = crop.convert("RGB").resize((width_px, round(width_px / aspect)), Image.LANCZOS)
        buf = io.BytesIO()
        crop.save(buf, "JPEG", quality=82, optimize=True)
        return buf.getvalue()


def lucide_svg(name: str, colour: str) -> bytes:
    svg = (LUCIDE / f"{name}.svg").read_text(encoding="utf-8")
    svg = re.sub(r'\s*class="[^"]*"', "", svg).replace("currentColor", f"#{colour}")
    return svg.encode("utf-8")


def svg_to_png(svg: bytes, width: int) -> bytes:
    try:
        import cairosvg
        return cairosvg.svg2png(bytestring=svg, output_width=width)
    except ImportError:
        script = ("const {Resvg}=require('@resvg/resvg-js');let s='';process.stdin.on('data',d=>s+=d);"
                  "process.stdin.on('end',()=>{const r=new Resvg(s,{fitTo:{mode:'width',value:+process.argv[1]}});"
                  "process.stdout.write(r.render().asPng())})")
        return subprocess.run(["node", "-e", script, str(width)], input=svg, cwd=ROOT, capture_output=True,
                              check=True).stdout


class Assets:
    """Adds pictures, icons and charts to one slide part and hands back their relationship ids."""

    def __init__(self, part, package):
        self.part, self.package = part, package

    def image(self, data: bytes) -> str:
        return self.part.get_or_add_image_part(io.BytesIO(data))[1]

    def svg(self, svg: bytes, png_width: int) -> tuple:
        """An SVG picture with a PNG fallback: the ids of the PNG and the SVG."""
        png_rid = self.image(svg_to_png(svg, png_width))
        svg_part = Part(self.package.next_image_partname("svg"), "image/svg+xml", self.package, svg)
        return png_rid, self.part.relate_to(svg_part, RT.IMAGE)

    def icon(self, name: str, colour: str) -> tuple:
        return self.svg(lucide_svg(name, colour), 160)

    def chart(self, chart_type, data) -> tuple:
        """A native chart: python-pptx writes the chart part and the workbook behind it, for Edit Data."""
        rid = self.part.add_chart_part(chart_type, data)
        return rid, self.part.related_part(rid).chart
