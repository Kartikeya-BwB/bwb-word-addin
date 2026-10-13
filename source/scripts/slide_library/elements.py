"""Elements the library slides share: the title placeholder, the source line, the BwB callout, coloured runs
and key figures.
"""
import re

from .drawing import Canvas, P, PH, esc, rpr
from .sources import BODY, CHARCOAL, G, GREEN, NAVY, OFFWHITE


def title(cv: Canvas, text: str):
    cv.placeholder("Title", 'type="title"', PH(text))


def source(cv: Canvas, text: str):
    cv.text("Source", G.source, P(text, 8, CHARCOAL), anchor="b")


def callout(cv: Canvas, box: tuple, heading: str, text: str, name: str = "Callout"):
    """An off-white box with a green bar, as the Callout element in the BwB tab."""
    x, y, w, h = box
    cv.shape(name, box, fill=OFFWHITE, text=P(heading, 14, NAVY, bold=True, after=4) + P(text, 12), anchor="t",
             inset=(20, 12, 16, 12))
    cv.shape(f"{name} bar", (x, y, 4, h), fill=GREEN)


def coloured(parts: list, size: float, align: str = "l") -> str:
    """One paragraph of bold Aptos runs, each in its own colour, such as a lime step number on navy."""
    out = "".join(f"<a:r>{rpr(size, colour, BODY, True)}<a:t>{esc(text)}</a:t></a:r>" for text, colour in parts)
    return (f'<a:p><a:pPr marL="0" indent="0" algn="{align}"><a:buNone/></a:pPr>{out}'
            f'{rpr(size, parts[-1][1], BODY, tag="a:endParaRPr")}</a:p>')


def figure(value: str, size: float, colour: str) -> str:
    """A key figure in Aptos bold. A leading currency code is set smaller, as in the BwB documents."""
    m = re.match(r"([A-Z]{3}) (.+)", value)
    out = f'<a:r>{rpr(round(size * 0.42), colour, BODY, True)}<a:t>{m.group(1)} </a:t></a:r>' if m else ""
    out += f"<a:r>{rpr(size, colour, BODY, True)}<a:t>{esc(m.group(2) if m else value)}</a:t></a:r>"
    return (f'<a:p><a:pPr marL="0" indent="0" algn="l"><a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:buNone/></a:pPr>'
            f'{out}{rpr(size, colour, BODY, tag="a:endParaRPr")}</a:p>')
