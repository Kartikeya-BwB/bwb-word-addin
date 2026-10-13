#!/usr/bin/env python3
"""Build the add-in's template assets from the BwB skill template (a dev-time step; the outputs are committed).

Usage:
  python3 scripts/make-base-template.py [--skill SKILL_DIR]

SKILL_DIR is the bwb-document-format skill folder (or set $BWB_SKILL_DIR).

Writes, relative to source/:
  public/templates/bwb-base.docx   the BwB template with its sample body replaced by one empty paragraph in the content
                                   section (logo header, Page X of Y footer). Sample figures, sample icons and the cover
                                   photograph are removed. Styles, numbering, settings (updateFields), theme, font table,
                                   web settings, every header and footer with its relationships, empty footnotes and
                                   endnotes, and the brand images the builder adds by name (navy, strip, both logos, header
                                   logo) are kept, so building a document needs no other download except the cover photo.
  public/templates/bwb-cover.jpg   the cover photograph, about 1400 px wide, JPEG quality about 82
  public/assets/brand/             logo-white.png, logo-black.png, header-logo.png
  src/docx/template/               styles.xml and numbering.xml for the insertOoxml path

The template's numbering.xml carries two restart instances (numId 20 and 21) left over from its sample lists. They are
removed here with the same expression the builder uses when it saves, so the parts start from list definitions 1 to 8.
The zip is written with fixed timestamps so that rerunning the script gives the same bytes.
"""
import argparse
import io
import os
import re
import sys
import zipfile

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
WML = "application/vnd.openxmlformats-officedocument.wordprocessingml."
FIXED_TIME = (2026, 10, 5, 9, 53, 41)
FIXED_ISO = "2026-10-05T09:53:41Z"

# parts copied unchanged from the template
COPY = [
    "word/styles.xml", "word/settings.xml", "word/theme/theme1.xml", "word/fontTable.xml", "word/webSettings.xml",
    "word/header1.xml", "word/header2.xml", "word/footer1.xml", "word/footer2.xml", "word/footer3.xml",
    "word/_rels/header1.xml.rels", "word/_rels/footer1.xml.rels", "word/_rels/footer2.xml.rels", "word/endnotes.xml",
]
# brand images the builder looks up by name, or that the headers and footers use
MEDIA = ["header_logo.png", "logo_black.png", "logo_white.png", "navy.png", "strip.png"]
BRAND_EXPORT = {"logo_white.png": "logo-white.png", "logo_black.png": "logo-black.png", "header_logo.png": "header-logo.png"}

# document relationships kept from the template (the cover photo, rId15, is dropped), plus the notes parts
DOC_RELS = [
    ("rId1", "styles", "styles.xml"), ("rId2", "numbering", "numbering.xml"), ("rId3", "settings", "settings.xml"),
    ("rId4", "webSettings", "webSettings.xml"), ("rId5", "theme", "theme/theme1.xml"), ("rId6", "fontTable", "fontTable.xml"),
    ("rId7", "header", "header1.xml"), ("rId8", "header", "header2.xml"),
    ("rId9", "footer", "footer1.xml"), ("rId10", "footer", "footer2.xml"), ("rId11", "footer", "footer3.xml"),
    ("rId12", "image", "media/logo_black.png"), ("rId13", "image", "media/logo_white.png"),
    ("rId14", "image", "media/navy.png"), ("rId16", "image", "media/strip.png"),
    ("rId17", "footnotes", "footnotes.xml"), ("rId18", "endnotes", "endnotes.xml"),
]

CT_OVERRIDES = [
    ("word/document.xml", WML + "document.main+xml"), ("word/styles.xml", WML + "styles+xml"),
    ("word/numbering.xml", WML + "numbering+xml"), ("word/settings.xml", WML + "settings+xml"),
    ("word/webSettings.xml", WML + "webSettings+xml"), ("word/footnotes.xml", WML + "footnotes+xml"),
    ("word/endnotes.xml", WML + "endnotes+xml"), ("word/fontTable.xml", WML + "fontTable+xml"),
    ("word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
    ("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    ("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
]

# the content section of the builder (Doc._sect("content")): logo header rId7, Page X of Y footer rId10
CONTENT_SECT = ('<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:footerReference w:type="default" r:id="rId10"/>'
                '<w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1224" w:right="1224" w:bottom="1224" '
                'w:left="1224" w:header="283" w:footer="778" w:gutter="0"/><w:cols w:space="720"/>'
                '<w:docGrid w:linePitch="360"/></w:sectPr>')


def root_ns(xml):
    """The namespace declarations of the template's document element, reused for the new parts."""
    m = re.search(r"<w:document (.*?)>", xml, re.S)
    if not m:
        raise SystemExit("word/document.xml in the template has no w:document element")
    return m.group(1)


def strip_restarts(numbering):
    # same expression as builder.Doc.save: numbered-list restart instances (numId 20 to 99) belong to a document body
    return re.sub(r'<w:num w:numId="(?:[2-9]\d)"[^>]*>.*?</w:num>', "", numbering, flags=re.S)


def rels_xml(items):
    out = [XML_DECL + f'<Relationships xmlns="{REL_NS}">']
    for rid, typ, target in items:
        out.append(f'<Relationship Id="{rid}" Type="{RT}{typ}" Target="{target}"/>')
    out.append("</Relationships>")
    return "".join(out)


def content_types(names):
    exts = {"rels": "application/vnd.openxmlformats-package.relationships+xml", "xml": "application/xml",
            "png": "image/png"}
    out = [XML_DECL + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">']
    for e in sorted(exts):
        out.append(f'<Default Extension="{e}" ContentType="{exts[e]}"/>')
    for part, ct in CT_OVERRIDES:
        if part in names:
            out.append(f'<Override PartName="/{part}" ContentType="{ct}"/>')
    for p in sorted(names):
        m = re.match(r"word/(header|footer)\d+\.xml$", p)
        if m:
            out.append(f'<Override PartName="/{p}" ContentType="{WML}{m.group(1)}+xml"/>')
    out.append("</Types>")
    return "".join(out)


def core_xml():
    return (XML_DECL +
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>BwB Document Template</dc:title><dc:subject></dc:subject>"
            "<dc:creator>Bankers without Boundaries</dc:creator><cp:keywords></cp:keywords>"
            "<dc:description>Base package for documents built by the BwB Word add-in</dc:description>"
            "<cp:lastModifiedBy>Bankers without Boundaries</cp:lastModifiedBy><cp:revision>1</cp:revision>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{FIXED_ISO}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{FIXED_ISO}</dcterms:modified></cp:coreProperties>')


def app_xml():
    return (XML_DECL + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Microsoft Office Word</Application><Company>Bankers without Boundaries</Company>"
            "<AppVersion>16.0000</AppVersion></Properties>")


def empty_footnotes(ns):
    # the same separators the builder writes when a document has no footnotes
    sep = ('<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
           '</w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
           '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
           'w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>')
    return XML_DECL + f"<w:footnotes {ns}>" + sep + "</w:footnotes>"


def write_zip(path, files):
    order = ["[Content_Types].xml", "_rels/.rels"]
    names = order + sorted(n for n in files if n not in order)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            info = zipfile.ZipInfo(n, date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            data = files[n]
            z.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)
    with open(path, "wb") as f:
        f.write(buf.getvalue())


def make_cover(src, dst, width=1400, quality=82, limit=450_000):
    """Resize the cover photograph and encode it as a baseline JPEG at the highest quality, up to `quality`, that keeps
    it under `limit`. The aerial photo is full of fine water texture, so at 1400 px quality 82 it is about 740 KB; the
    loop lowers the quality instead of the width because the photo is printed full-bleed (21 cm) on the cover."""
    im = Image.open(src).convert("RGB")
    h = round(im.size[1] * width / im.size[0])
    im = im.resize((width, h), Image.LANCZOS)
    q = quality
    while True:
        b = io.BytesIO()
        im.save(b, "JPEG", quality=q, optimize=True, dpi=(220, 220))
        if b.tell() <= limit or q <= 50:
            break
        q -= 1
    with open(dst, "wb") as f:
        f.write(b.getvalue())
    return (width, h), q, b.tell()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--skill", default=os.environ.get("BWB_SKILL_DIR"), required="BWB_SKILL_DIR" not in os.environ,
                    help="the bwb-document-format skill folder")
    a = ap.parse_args()
    tpl_path = os.path.join(a.skill, "assets", "BwB_Document_Template.docx")
    with zipfile.ZipFile(tpl_path) as z:
        tpl = {i.filename: z.read(i.filename) for i in z.infolist() if not i.is_dir()}

    ns = root_ns(tpl["word/document.xml"].decode("utf-8"))
    numbering = strip_restarts(tpl["word/numbering.xml"].decode("utf-8"))
    files = {n: tpl[n] for n in COPY}
    files["word/numbering.xml"] = numbering
    files["word/footnotes.xml"] = empty_footnotes(ns)
    files["word/document.xml"] = XML_DECL + f"<w:document {ns}><w:body><w:p/>{CONTENT_SECT}</w:body></w:document>"
    for m in MEDIA:
        files[f"word/media/{m}"] = tpl[f"word/media/{m}"]
    files["word/_rels/document.xml.rels"] = rels_xml(DOC_RELS)
    files["docProps/core.xml"] = core_xml()
    files["docProps/app.xml"] = app_xml()
    files["_rels/.rels"] = (XML_DECL + f'<Relationships xmlns="{REL_NS}">'
                            f'<Relationship Id="rId1" Type="{RT}officeDocument" Target="word/document.xml"/>'
                            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/'
                            'metadata/core-properties" Target="docProps/core.xml"/>'
                            f'<Relationship Id="rId3" Type="{RT}extended-properties" Target="docProps/app.xml"/>'
                            "</Relationships>")
    files["[Content_Types].xml"] = content_types(set(files))

    out_tpl = os.path.join(ROOT, "public", "templates")
    out_brand = os.path.join(ROOT, "public", "assets", "brand")
    out_src = os.path.join(ROOT, "src", "docx", "template")
    for d in (out_tpl, out_brand, out_src):
        os.makedirs(d, exist_ok=True)
    base = os.path.join(out_tpl, "bwb-base.docx")
    write_zip(base, files)
    size, q, nbytes = make_cover(os.path.join(a.skill, "assets", "brand", "cover_default.jpg"),
                                 os.path.join(out_tpl, "bwb-cover.jpg"))
    for src, dst in BRAND_EXPORT.items():
        with open(os.path.join(out_brand, dst), "wb") as f:
            f.write(tpl[f"word/media/{src}"])
    with open(os.path.join(out_src, "styles.xml"), "wb") as f:
        f.write(tpl["word/styles.xml"])
    with open(os.path.join(out_src, "numbering.xml"), "w", encoding="utf-8") as f:
        f.write(numbering)

    print(f"bwb-base.docx {os.path.getsize(base)} bytes, {len(files)} parts")
    print(f"bwb-cover.jpg {size[0]}x{size[1]} quality {q}, {nbytes} bytes")
    try:
        import docx  # python-docx, only to prove the package opens
        docx.Document(base)
        print("python-docx opens bwb-base.docx")
    except ImportError:
        print("python-docx is not installed; open check skipped", file=sys.stderr)


if __name__ == "__main__":
    main()
