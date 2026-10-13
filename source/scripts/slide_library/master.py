"""The theme, the default BwB table style, the slide master and the seven layouts, and the package without
slides that python-pptx opens to add them.
"""
import io
import zipfile

from .assets import trimmed_png
from .drawing import Canvas, NS, P, PH, body_pr, creation_ext, emu, esc, lvl_ppr, solid, sp_tree
from .sources import (
    BLACK, BODY, BRAND, CB, CH, CHARCOAL, CT, FIXED_ISO, G, GREEN, GREEN50, GREY, HEAD, LIME, NAVY, OFFWHITE, RIGHT,
    WHITE,
)


P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"
GUIDE_EXT = "{27BBF7A9-308A-43DC-89C8-2F10F3537804}"

BWB_TABLE_STYLE = "{6A0B2E4D-5C3F-4E8B-9C1A-2B7D4F6E8A10}"
LAYOUT_NAMES = ["BwB Title", "BwB Section divider", "BwB Title and content", "BwB Two content", "BwB Title only",
                "BwB Blank", "BwB Closing"]
NAVY_LAYOUTS = ("BwB Title", "BwB Section divider", "BwB Closing")
WHITE_LOGO_LAYOUTS = ("BwB Title", "BwB Closing")


def theme_xml() -> str:
    colours = [("dk1", BLACK), ("lt1", WHITE), ("dk2", NAVY), ("lt2", OFFWHITE), ("accent1", NAVY),
               ("accent2", GREEN), ("accent3", GREEN50), ("accent4", CHARCOAL), ("accent5", LIME),
               ("accent6", GREY), ("hlink", GREEN), ("folHlink", CHARCOAL)]
    scheme = "".join(f'<a:{k}><a:srgbClr val="{v}"/></a:{k}>' for k, v in colours)
    fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    lines = "".join(f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr">{fill}<a:prstDash val="solid"/>'
                    '<a:miter lim="800000"/></a:ln>' for w in (9525, 12700, 19050))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="BwB">'
            f'<a:themeElements><a:clrScheme name="BwB">{scheme}</a:clrScheme>'
            f'<a:fontScheme name="BwB"><a:majorFont><a:latin typeface="{HEAD}"/><a:ea typeface=""/><a:cs typeface=""/>'
            f'</a:majorFont><a:minorFont><a:latin typeface="{BODY}"/><a:ea typeface=""/><a:cs typeface=""/>'
            '</a:minorFont></a:fontScheme>'
            f'<a:fmtScheme name="BwB"><a:fillStyleLst>{fill * 3}</a:fillStyleLst><a:lnStyleLst>{lines}</a:lnStyleLst>'
            '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/>'
            '</a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
            f'<a:bgFillStyleLst>{fill * 3}</a:bgFillStyleLst></a:fmtScheme></a:themeElements>'
            # new shapes: navy with white Aptos text; new lines: navy; new text boxes: Aptos 14 black, no margins
            '<a:objectDefaults><a:spDef><a:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
            '<a:ln><a:noFill/></a:ln></a:spPr><a:bodyPr rtlCol="0" anchor="ctr" lIns="72000" tIns="72000" '
            'rIns="72000" bIns="72000"/><a:lstStyle><a:defPPr><a:defRPr lang="en-GB" sz="1400"><a:solidFill>'
            '<a:schemeClr val="lt1"/></a:solidFill></a:defRPr></a:defPPr></a:lstStyle></a:spDef>'
            '<a:lnDef><a:spPr><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln>'
            '</a:spPr><a:bodyPr/><a:lstStyle/></a:lnDef>'
            '<a:txDef><a:spPr><a:noFill/></a:spPr><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" '
            'rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle><a:defPPr><a:defRPr lang="en-GB" sz="1400">'
            '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:defRPr></a:defPPr></a:lstStyle></a:txDef>'
            '</a:objectDefaults><a:extraClrSchemeLst/></a:theme>')


def table_styles_xml() -> str:
    """A BwB table style, the default for new tables: navy header row, light grey rules, total row in off-white."""
    grey = f'<a:ln w="9525" cmpd="sng">{solid(GREY)}</a:ln>'
    none = "<a:ln><a:noFill/></a:ln>"

    def txt(colour: str, bold: bool) -> str:
        b = ' b="on"' if bold else ""
        return (f'<a:tcTxStyle{b}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef>'
                f'<a:srgbClr val="{colour}"/></a:tcTxStyle>')

    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{BWB_TABLE_STYLE}">'
            f'<a:tblStyle styleId="{BWB_TABLE_STYLE}" styleName="BwB table"><a:wholeTbl>{txt(BLACK, False)}'
            f'<a:tcStyle><a:tcBdr><a:left>{none}</a:left><a:right>{none}</a:right><a:top>{grey}</a:top>'
            f'<a:bottom>{grey}</a:bottom><a:insideH>{grey}</a:insideH><a:insideV>{none}</a:insideV></a:tcBdr>'
            '<a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl>'
            f'<a:lastRow>{txt(BLACK, True)}<a:tcStyle><a:tcBdr><a:top><a:ln w="12700" cmpd="sng">{solid(NAVY)}</a:ln>'
            f'</a:top></a:tcBdr><a:fill>{solid(OFFWHITE)}</a:fill></a:tcStyle></a:lastRow>'
            f'<a:firstRow>{txt(WHITE, True)}<a:tcStyle><a:tcBdr/><a:fill>{solid(NAVY)}</a:fill></a:tcStyle>'
            '</a:firstRow></a:tblStyle></a:tblStyleLst>')


def master_xml(logo_rid: str, logo_aspect: float) -> str:
    cv = Canvas()
    cv.placeholder("Title", 'type="title"', PH("Action title: say the so-what in one sentence"), G.title,
                   body_pr("b"))
    cv.placeholder("Text", 'type="body" idx="1"',
                   PH("Text in Aptos 14 pt: lead with the point") + PH("Second level: the supporting detail", 1)
                   + PH("Third level", 2), G.content, body_pr("t"))
    fx, fy, fw, fh = G.footer
    cv.slide_number('type="sldNum" sz="quarter" idx="4"', box=(fx + fw - 80, fy, 80, fh), bpr=body_pr("ctr"),
                    lst=lvl_ppr(1, 9, CHARCOAL, align="r"))
    logo_h = 16
    cv.picture("BwB logo", logo_rid, (fx, fy + (fh - logo_h) / 2, logo_h * logo_aspect, logo_h),
               "Bankers without Boundaries logo")
    cv.strip()
    layouts = "".join(f'<p:sldLayoutId id="{2147483649 + i}" r:id="rId{i + 1}"/>' for i in range(len(LAYOUT_NAMES)))
    title_style = lvl_ppr(1, 24, NAVY, "+mj-lt", bold=True, line=90)
    body_levels = [lvl_ppr(1, 14, BLACK, after=8, bullet="square"), lvl_ppr(2, 12, BLACK, after=6, bullet="circle")]
    body_levels += [lvl_ppr(n, 12, BLACK, after=4, bullet="dash") for n in range(3, 10)]
    other = "".join(lvl_ppr(n, 14, BLACK) for n in range(1, 10))
    guides = [("vert", G.margin_left), ("vert", G.x(6) - G.gutter), ("vert", G.x(6)), ("vert", RIGHT),
              ("horz", G.title[1]), ("horz", G.title[1] + G.title[3]), ("horz", CT), ("horz", CB),
              ("horz", G.source[1]), ("horz", G.footer[1])]
    guide_xml = "".join(f'<p15:guide id="{i + 1}" orient="{o}" pos="{int(round(v * 8))}" userDrawn="1"><p15:clr>'
                        '<a:srgbClr val="F26B43"/></p15:clr></p15:guide>' for i, (o, v) in enumerate(guides))
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldMaster {NS}>'
            f'<p:cSld><p:bg><p:bgPr>{solid(WHITE)}<a:effectLst/></p:bgPr></p:bg>{sp_tree(cv.xml())}'
            f'{creation_ext("master")}</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" '
            'accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            f'<p:sldLayoutIdLst>{layouts}</p:sldLayoutIdLst><p:hf hdr="0" ftr="0" dt="0"/>'
            f'<p:txStyles><p:titleStyle>{title_style}</p:titleStyle><p:bodyStyle>{"".join(body_levels)}</p:bodyStyle>'
            f'<p:otherStyle><a:defPPr><a:defRPr lang="en-GB"/></a:defPPr>{other}</p:otherStyle></p:txStyles>'
            f'<p:extLst><p:ext uri="{GUIDE_EXT}"><p15:sldGuideLst xmlns:p15="{P15}">{guide_xml}</p15:sldGuideLst>'
            "</p:ext></p:extLst></p:sldMaster>")


def layout_xml(name: str, white_logo_rid: str, white_logo_aspect: float) -> str:
    """One of the seven layouts. The navy layouts hide the master's footer and draw their own."""
    cv = Canvas()
    navy = name in NAVY_LAYOUTS
    slide_number = 'type="sldNum" sz="quarter" idx="12"'
    wide = G.w(10)

    def logo():
        cv.picture("BwB logo", white_logo_rid, (G.margin_left, G.margin_top, 48 * white_logo_aspect, 48),
                   "Bankers without Boundaries logo")

    if name == "BwB Title":
        logo()
        cv.placeholder("Title", 'type="ctrTitle" hasCustomPrompt="1"', PH("Presentation Title in Title Case"),
                       (G.margin_left, 168, wide, 168), body_pr("b"),
                       lvl_ppr(1, 40, WHITE, "+mj-lt", bold=True, line=90))
        cv.shape("Lime rule", (G.margin_left, 352, G.w(2), 4), fill=LIME)
        cv.placeholder("Subtitle", 'type="subTitle" idx="1" hasCustomPrompt="1"',
                       PH("Subtitle: the client, the project and the purpose of the meeting"),
                       (G.margin_left, 372, wide, 56), body_pr("t"), lvl_ppr(1, 18, WHITE))
        cv.placeholder("Date", 'type="body" sz="quarter" idx="13" hasCustomPrompt="1"', PH("Date: 5 October 2026"),
                       (G.margin_left, 440, G.w(6), 24), body_pr("t"), lvl_ppr(1, 14, WHITE))
        cv.strip()
    elif name == "BwB Section divider":
        cv.placeholder("Section number", 'type="body" sz="quarter" idx="13" hasCustomPrompt="1"', PH("01"),
                       (G.margin_left, 96, G.w(6), 160), body_pr("b"),
                       lvl_ppr(1, 120, LIME, "+mj-lt", bold=True, line=80))
        cv.placeholder("Title", 'type="title" hasCustomPrompt="1"', PH("Section title in sentence case"),
                       (G.margin_left, 280, wide, 96), body_pr("t"),
                       lvl_ppr(1, 36, WHITE, "+mj-lt", bold=True, line=90))
        cv.placeholder("Description", 'type="body" sz="quarter" idx="14" hasCustomPrompt="1"',
                       PH("One sentence on what this section shows"), (G.margin_left, 392, G.w(8), 48),
                       body_pr("t"), lvl_ppr(1, 16, WHITE))
        cv.strip()
    elif name == "BwB Closing":
        logo()
        cv.placeholder("Title", 'type="title" hasCustomPrompt="1"', PH("Thank you"),
                       (G.margin_left, 168, wide, 96), body_pr("b"),
                       lvl_ppr(1, 40, WHITE, "+mj-lt", bold=True, line=90))
        cv.shape("Lime rule", (G.margin_left, 280, G.w(2), 4), fill=LIME)
        for i in (1, 2):
            cv.placeholder(f"Contact {i}", f'type="body" idx="{i}" hasCustomPrompt="1"',
                           PH("**Name Surname**") + PH("Role") + PH("name.surname@bwb.earth"),
                           (G.x(4 * (i - 1)), 304, G.w(4), 120), body_pr("t"), lvl_ppr(1, 14, WHITE, after=2))
        cv.text("Website", (G.margin_left, 476, G.w(4), 20), P("bwb.earth", 14, WHITE, bold=True), anchor="b")
        cv.strip()
    else:
        if name != "BwB Blank":
            cv.placeholder("Title", 'type="title" hasCustomPrompt="1"',
                           PH("Action title: say the so-what in one sentence, on no more than two lines"))
        prompt = PH("Text in Aptos 14 pt: lead with the point, then the evidence") + PH("Second level", 1)
        if name == "BwB Title and content":
            cv.placeholder("Content", 'idx="1" hasCustomPrompt="1"', prompt)
        elif name == "BwB Two content":
            half = G.w(6)
            cv.placeholder("Left content", 'sz="half" idx="1" hasCustomPrompt="1"', prompt, (G.x(0), CT, half, CH))
            cv.placeholder("Right content", 'sz="half" idx="2" hasCustomPrompt="1"', prompt, (G.x(6), CT, half, CH))
        cv.slide_number(slide_number)
    kind = {"BwB Title": "title", "BwB Section divider": "secHead", "BwB Title and content": "obj",
            "BwB Two content": "twoObj", "BwB Title only": "titleOnly", "BwB Blank": "blank"}.get(name)
    attrs = (f' type="{kind}"' if kind else "") + (' showMasterSp="0"' if navy else "")
    bg = f"<p:bg><p:bgPr>{solid(NAVY)}<a:effectLst/></p:bgPr></p:bg>" if navy else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldLayout {NS}{attrs} preserve="1" '
            f'userDrawn="1"><p:cSld name="{name}">{bg}{sp_tree(cv.xml())}{creation_ext("layout/" + name)}</p:cSld>'
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>")


def core_xml(title: str) -> str:
    """Document properties. The starter deck has no title, author or dates, so a deck made from it starts clean."""
    props = ""
    if title:
        props = (f"<dc:title>{esc(title)}</dc:title><dc:creator>Bankers without Boundaries</dc:creator>"
                 "<cp:lastModifiedBy>Bankers without Boundaries</cp:lastModifiedBy>"
                 f'<dcterms:created xsi:type="dcterms:W3CDTF">{FIXED_ISO}</dcterms:created>'
                 f'<dcterms:modified xsi:type="dcterms:W3CDTF">{FIXED_ISO}</dcterms:modified>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"{props}<cp:revision>1</cp:revision><dc:language>en-GB</dc:language></cp:coreProperties>")


def base_package(slide_count: int, title: str) -> bytes:
    """A pptx with the theme, the master and the seven layouts and no slides, for python-pptx to open."""
    footer_logo, footer_aspect = trimmed_png(BRAND / "header-logo.png")
    white_logo, white_aspect = trimmed_png(BRAND / "logo-white.png")
    rel = ('<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{}" '
           'Target="{}"/>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{}</Relationships>')
    ct = "application/vnd.openxmlformats-officedocument.presentationml."
    overrides = [("/ppt/presentation.xml", ct + "presentation.main+xml"),
                 ("/ppt/slideMasters/slideMaster1.xml", ct + "slideMaster+xml"),
                 ("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
                 ("/ppt/presProps.xml", ct + "presProps+xml"), ("/ppt/viewProps.xml", ct + "viewProps+xml"),
                 ("/ppt/tableStyles.xml", ct + "tableStyles+xml"),
                 ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
                 ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")]
    overrides += [(f"/ppt/slideLayouts/slideLayout{i + 1}.xml", ct + "slideLayout+xml")
                  for i in range(len(LAYOUT_NAMES))]
    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
            + "".join(f'<Override PartName="{p}" ContentType="{c}"/>' for p, c in overrides) + "</Types>"),
        "_rels/.rels": rels.format(
            rel.format("rId1", "officeDocument", "ppt/presentation.xml")
            + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
              'core-properties" Target="docProps/core.xml"/>'
            + rel.format("rId3", "extended-properties", "docProps/app.xml")),
        "docProps/core.xml": core_xml(title),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties '
            'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Microsoft Office PowerPoint</Application><PresentationFormat>Widescreen</PresentationFormat>"
            f"<Slides>{slide_count}</Slides><Company>Bankers without Boundaries</Company>"
            "<AppVersion>16.0000</AppVersion></Properties>"),
        "ppt/_rels/presentation.xml.rels": rels.format(
            rel.format("rId1", "slideMaster", "slideMasters/slideMaster1.xml")
            + rel.format("rId2", "theme", "theme/theme1.xml")
            + rel.format("rId3", "presProps", "presProps.xml") + rel.format("rId4", "viewProps", "viewProps.xml")
            + rel.format("rId5", "tableStyles", "tableStyles.xml")),
        "ppt/presentation.xml": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation {NS}>'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst/>'
            f'<p:sldSz cx="{emu(G.width)}" cy="{emu(G.height)}"/><p:notesSz cx="6858000" cy="9144000"/>'
            f'<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-GB"/></a:defPPr>'
            + "".join(lvl_ppr(n, 14, BLACK) for n in range(1, 10)) + "</p:defaultTextStyle></p:presentation>"),
        "ppt/presProps.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr {NS}/>',
        "ppt/viewProps.xml": (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr {NS}><p:normalViewPr>'
            '<p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr>'
            '<p:cViewPr varScale="1"><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale>'
            '<p:origin x="0" y="0"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:gridSpacing '
            f'cx="{emu(4)}" cy="{emu(4)}"/></p:viewPr>'),
        "ppt/tableStyles.xml": table_styles_xml(),
        "ppt/theme/theme1.xml": theme_xml(),
        "ppt/media/image1.png": footer_logo,
        "ppt/media/image2.png": white_logo,
        "ppt/slideMasters/slideMaster1.xml": master_xml(f"rId{len(LAYOUT_NAMES) + 2}", footer_aspect),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": rels.format(
            "".join(rel.format(f"rId{i + 1}", "slideLayout", f"../slideLayouts/slideLayout{i + 1}.xml")
                    for i in range(len(LAYOUT_NAMES)))
            + rel.format(f"rId{len(LAYOUT_NAMES) + 1}", "theme", "../theme/theme1.xml")
            + rel.format(f"rId{len(LAYOUT_NAMES) + 2}", "image", "../media/image1.png")),
    }
    for i, name in enumerate(LAYOUT_NAMES):
        files[f"ppt/slideLayouts/slideLayout{i + 1}.xml"] = layout_xml(name, "rId2", white_aspect)
        logo = rel.format("rId2", "image", "../media/image2.png") if name in WHITE_LOGO_LAYOUTS else ""
        files[f"ppt/slideLayouts/_rels/slideLayout{i + 1}.xml.rels"] = rels.format(
            rel.format("rId1", "slideMaster", "../slideMasters/slideMaster1.xml") + logo)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    return buf.getvalue()
