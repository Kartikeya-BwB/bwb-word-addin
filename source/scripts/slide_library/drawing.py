"""DrawingML helpers: runs, paragraphs and list styles, table cells, stable creation ids, and the Canvas that
collects the shapes of a slide, layout or master. Everything is in points; emu() converts.
"""
import zlib
from xml.sax.saxutils import escape

from .sources import BLACK, BODY, CHARCOAL, G, GREEN, LIME, NAVY


NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
CREATION_EXT = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
SVG_EXT = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
SLIDENUM_FIELD = "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}"
NO_STYLE_TABLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def emu(pt: float) -> int:
    return int(round(pt * 12700))


def esc(text: str) -> str:
    return escape(text, {'"': "&quot;"})


def solid(colour: str) -> str:
    return f'<a:solidFill><a:srgbClr val="{colour}"/></a:solidFill>'


def xfrm(box: tuple, tag: str = "a:xfrm", flip_v: bool = False) -> str:
    x, y, w, h = box
    flip = ' flipV="1"' if flip_v else ""
    return f'<{tag}{flip}><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></{tag}>'


def rpr(size=None, colour=None, font=None, bold=None, italic=None, tag="a:rPr") -> str:
    attrs = ['lang="en-GB"']
    if size is not None:
        attrs.append(f'sz="{int(round(size * 100))}"')
    if bold is not None:
        attrs.append(f'b="{1 if bold else 0}"')
    if italic is not None:
        attrs.append(f'i="{1 if italic else 0}"')
    inner = solid(colour) if colour else ""
    if font:
        inner += f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/>'
    return f'<{tag} {" ".join(attrs)}>{inner}</{tag}>'


def runs(text: str, size=None, colour=None, font=None, bold=None, italic=None) -> str:
    """Runs for one paragraph. **Bold** marks bold words and a vertical tab is a line break."""
    out = []
    for li, line in enumerate(text.split("\v")):
        if li:
            out.append(f'<a:br>{rpr(size, colour, font, bold, italic)}</a:br>')
        for i, part in enumerate(line.split("**")):
            if part:
                strong = True if i % 2 else bold
                out.append(f'<a:r>{rpr(size, colour, font, strong, italic)}<a:t>{esc(part)}</a:t></a:r>')
    return "".join(out)


# bullet: colour, character, font, size in per cent, left margin and hanging indent in points. The glyphs, fonts and
# colours are the ones the BwB Word template uses for its lists.
BULLETS = {
    "square": (GREEN, "\u25A0", "Arial", 80, 16, 16),
    "circle": (NAVY, "\u25CF", "Arial", 70, 32, 14),
    "dash": (CHARCOAL, "\u2013", "Arial", 100, 46, 14),
}


def bullet_xml(bullet: str) -> tuple:
    """The margin attributes and bullet elements of a paragraph or list level."""
    if not bullet:
        return 'marL="0" indent="0"', "<a:buNone/>"
    colour, char, font, pct, marl, hang = BULLETS[bullet]
    return (f'marL="{emu(marl)}" indent="{emu(-hang)}"',
            f'<a:buClr><a:srgbClr val="{colour}"/></a:buClr><a:buSzPct val="{pct * 1000}"/>'
            f'<a:buFont typeface="{font}" panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
            f'<a:buChar char="{char}"/>')


def P(text: str = "", size: float = 14, colour: str = BLACK, font: str = BODY, bold: bool = False,
      italic: bool = False, align: str = "l", after: float = 0, before: float = 0, line: float = 0,
      bullet: str = "") -> str:
    """One paragraph with explicit formatting, for text boxes and shapes."""
    attrs, bu = bullet_xml(bullet)
    spacing = f'<a:lnSpc><a:spcPct val="{int(line * 1000)}"/></a:lnSpc>' if line else ""
    spacing += f'<a:spcBef><a:spcPts val="{int(before * 100)}"/></a:spcBef>'
    spacing += f'<a:spcAft><a:spcPts val="{int(after * 100)}"/></a:spcAft>'
    ppr = f'<a:pPr {attrs} algn="{align}">{spacing}{bu}</a:pPr>'
    end = rpr(size, colour, font, tag="a:endParaRPr")
    return f"<a:p>{ppr}{runs(text, size, colour, font, bold or None, italic or None)}{end}</a:p>"


def PH(text: str, level: int = 0) -> str:
    """One paragraph for a placeholder: the formatting comes from the layout and the master."""
    ppr = f'<a:pPr lvl="{level}"/>' if level else ""
    return f'<a:p>{ppr}{runs(text)}<a:endParaRPr lang="en-GB"/></a:p>'


def body_pr(anchor: str = "t", inset: tuple = (0, 0, 0, 0), vert: str = "", autofit: str = "<a:noAutofit/>") -> str:
    l, t, r, b = inset
    rotate = f' vert="{vert}"' if vert else ""
    return (f'<a:bodyPr{rotate} wrap="square" lIns="{emu(l)}" tIns="{emu(t)}" rIns="{emu(r)}" bIns="{emu(b)}" '
            f'rtlCol="0" anchor="{anchor}" anchorCtr="0">{autofit}</a:bodyPr>')


def lvl_ppr(level: int, size: float, colour: str, font: str = "+mn-lt", bold: bool = False, italic: bool = False,
            align: str = "l", line: float = 0, after: float = 0, bullet: str = "") -> str:
    """A list style level (lvlNpPr) for layout placeholders and the master's text styles."""
    attrs, bu = bullet_xml(bullet)
    spacing = f'<a:lnSpc><a:spcPct val="{int(line * 1000)}"/></a:lnSpc>' if line else ""
    spacing += f'<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="{int(after * 100)}"/></a:spcAft>'
    style = ' b="1"' if bold else ' b="0"'
    style += ' i="1"' if italic else ""
    ea, cs = ("+mj-ea", "+mj-cs") if font.startswith("+mj") else ("+mn-ea", "+mn-cs") if font.startswith("+") else (
        "+mn-ea", font)
    rp = (f'<a:defRPr lang="en-GB" sz="{int(size * 100)}"{style} kern="1200">{solid(colour)}'
          f'<a:latin typeface="{font}"/><a:ea typeface="{ea}"/><a:cs typeface="{cs}"/></a:defRPr>')
    return (f'<a:lvl{level}pPr {attrs} algn="{align}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" '
            f'hangingPunct="1">{spacing}{bu}{rp}</a:lvl{level}pPr>')


class Canvas:
    """Collects the shapes of one slide, layout or master and numbers them from 2."""

    def __init__(self):
        self.parts = []
        self.last = 1

    def _id(self) -> int:
        self.last += 1
        return self.last

    def shape(self, name: str, box: tuple, *, geom: str = "rect", adj: dict = None, fill: str = None,
              line: str = None, line_w: float = 0.75, text: str = None, anchor: str = "t",
              inset: tuple = (0, 0, 0, 0), txbox: bool = False, vert: str = "") -> int:
        sid = self._id()
        tb = ' txBox="1"' if txbox else ""
        guides = "".join(f'<a:gd name="{k}" fmla="val {v}"/>' for k, v in (adj or {}).items())
        ln = f'<a:ln w="{emu(line_w)}">{solid(line)}</a:ln>' if line else "<a:ln><a:noFill/></a:ln>"
        sp_pr = (f'<p:spPr>{xfrm(box)}<a:prstGeom prst="{geom}"><a:avLst>{guides}</a:avLst></a:prstGeom>'
                 f'{solid(fill) if fill else "<a:noFill/>"}{ln}</p:spPr>')
        tx = f"<p:txBody>{body_pr(anchor, inset, vert)}<a:lstStyle/>{text}</p:txBody>" if text is not None else ""
        self.parts.append(f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr{tb}/><p:nvPr/>'
                          f"</p:nvSpPr>{sp_pr}{tx}</p:sp>")
        return sid

    def text(self, name: str, box: tuple, paras: str, **kw) -> int:
        return self.shape(name, box, text=paras, txbox=True, **kw)

    def line(self, name: str, x1: float, y1: float, x2: float, y2: float, colour: str, width: float = 0.75,
             arrow: bool = False) -> int:
        sid = self._id()
        box = (min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
        head = '<a:tailEnd type="triangle" w="med" len="med"/>' if arrow else ""
        self.parts.append(
            f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f'<p:spPr>{xfrm(box, flip_v=y2 < y1)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
            f'<a:ln w="{emu(width)}">{solid(colour)}{head}</a:ln></p:spPr></p:cxnSp>')
        return sid

    def picture(self, name: str, rid: str, box: tuple, descr: str, svg_rid: str = None) -> int:
        sid = self._id()
        ext = ""
        if svg_rid:
            ext = (f'<a:extLst><a:ext uri="{SVG_EXT}"><asvg:svgBlip '
                   f'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{svg_rid}"/>'
                   "</a:ext></a:extLst>")
        self.parts.append(
            f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="{esc(name)}" descr="{esc(descr)}"/><p:cNvPicPr>'
            '<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}">{ext}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
        return sid

    def placeholder(self, name: str, ph: str, paras: str = "", box: tuple = None, bpr: str = "<a:bodyPr/>",
                    lst: str = "") -> int:
        sid = self._id()
        sp_pr = f"<p:spPr>{xfrm(box)}</p:spPr>" if box else "<p:spPr/>"
        self.parts.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            f"<p:nvPr><p:ph {ph}/></p:nvPr></p:nvSpPr>{sp_pr}<p:txBody>{bpr}<a:lstStyle>{lst}</a:lstStyle>"
            f"{paras or PH('')}</p:txBody></p:sp>")
        return sid

    def slide_number(self, ph: str, number: str = "\u2039#\u203A", box: tuple = None, bpr: str = "<a:bodyPr/>",
                     lst: str = "") -> int:
        field = (f'<a:p><a:fld id="{SLIDENUM_FIELD}" type="slidenum"><a:rPr lang="en-GB"/><a:t>{number}</a:t></a:fld>'
                 '<a:endParaRPr lang="en-GB"/></a:p>')
        return self.placeholder("Slide number", ph, field, box, bpr, lst)

    def table(self, name: str, x: float, y: float, widths: list, rows: list) -> int:
        """rows: list of (height, cells); a cell is the <a:tc> XML from tc()."""
        sid = self._id()
        grid = "".join(f'<a:gridCol w="{emu(w)}"/>' for w in widths)
        body = "".join(f'<a:tr h="{emu(h)}">{"".join(cells)}</a:tr>' for h, cells in rows)
        box = (x, y, sum(widths), sum(h for h, _ in rows))
        self.parts.append(
            f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvGraphicFramePr>'
            '<a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'{xfrm(box, "p:xfrm")}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'<a:tbl><a:tblPr firstRow="1" bandRow="0"><a:tableStyleId>{NO_STYLE_TABLE}</a:tableStyleId></a:tblPr>'
            f"<a:tblGrid>{grid}</a:tblGrid>{body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")
        return sid

    def chart(self, name: str, rid: str, box: tuple, descr: str) -> int:
        sid = self._id()
        self.parts.append(
            f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{esc(name)}" descr="{esc(descr)}"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'{xfrm(box, "p:xfrm")}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
            f'<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="{rid}"/></a:graphicData>'
            "</a:graphic></p:graphicFrame>")
        return sid

    def strip(self):
        """The green, charcoal and lime strip along the bottom edge, as in the BwB documents."""
        third = G.width / 3
        for i, (label, colour) in enumerate((("green", GREEN), ("charcoal", CHARCOAL), ("lime", LIME))):
            self.shape(f"Colour strip {label}", (i * third, G.height - G.strip, third, G.strip), fill=colour)

    def xml(self) -> str:
        return "".join(self.parts)


def tc(paras: str, fill: str = None, lines: dict = None, margins: tuple = (6, 4, 6, 4), anchor: str = "ctr",
       span: int = 0, hmerge: bool = False, rowspan: int = 0, vmerge: bool = False) -> str:
    """One table cell. lines maps L, R, T and B to (colour, width in points)."""
    attrs = (f' gridSpan="{span}"' if span else "") + (f' rowSpan="{rowspan}"' if rowspan else "")
    attrs += (' hMerge="1"' if hmerge else "") + (' vMerge="1"' if vmerge else "")
    borders = ""
    for side in "LRTB":
        spec = (lines or {}).get(side)
        if spec:
            borders += (f'<a:ln{side} w="{emu(spec[1])}" cap="flat" cmpd="sng" algn="ctr">{solid(spec[0])}'
                        f'<a:prstDash val="solid"/></a:ln{side}>')
        else:
            borders += f'<a:ln{side} w="0"><a:noFill/></a:ln{side}>'
    l, t, r, b = margins
    return (f'<a:tc{attrs}><a:txBody><a:bodyPr/><a:lstStyle/>{paras}</a:txBody>'
            f'<a:tcPr marL="{emu(l)}" marR="{emu(r)}" marT="{emu(t)}" marB="{emu(b)}" anchor="{anchor}">'
            f'{borders}{solid(fill) if fill else "<a:noFill/>"}</a:tcPr></a:tc>')


def creation_id(key: str) -> int:
    """A stable p14:creationId for a slide, layout or master, derived from its name."""
    return (zlib.crc32(f"bwb-slides/{key}".encode("utf-8")) & 0x7FFFFFFF) or 1


def creation_ext(key: str) -> str:
    return (f'<p:extLst><p:ext uri="{CREATION_EXT}"><p14:creationId xmlns:p14="{P14}" val="{creation_id(key)}"/>'
            "</p:ext></p:extLst>")


def sp_tree(content: str) -> str:
    return ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>'
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
            f"</p:grpSpPr>{content}</p:spTree>")
