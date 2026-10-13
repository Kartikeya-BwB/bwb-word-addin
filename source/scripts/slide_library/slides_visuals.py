"""Library slides for data and visuals: key figures, table, process, matrix, timeline, workplan, icon grid,
and image and text. Their text is written as guidance for the author.
"""
from .assets import Assets, photo
from .drawing import Canvas, P, tc
from .elements import callout, coloured, figure, source, title
from .sources import BLACK, CB, CH, CHARCOAL, CT, G, GREEN, GREY, GREY20, HEAD, LIME, NAVY, OFFWHITE, RIGHT, WHITE


FIGURES = [("USD 250m", "Headline figure, with its unit and year",
            "Lead with the figure that carries the message. Only this one is lime."),
           ("38%", "Second figure: say what it measures",
            "Round figures to what the reader needs: 38%, not 37.84%."),
           ("12", "Countries, projects or partners",
            "Use the % sign in figures and \u2018per cent\u2019 in sentences."),
           ("2030", "A date or a target year",
            "Write currency as a code and a figure: GBP 5m, USD 250m, EUR 1.2bn.")]


def s_key_figures(cv: Canvas, a: Assets):
    title(cv, "Action title: what the numbers show, in one sentence")
    cv.text("Introduction", (G.x(0), CT, G.w(8), 40),
            P("One or two lines that set up the figures: the scope, the period and where the data comes from.", 14))
    top, band = CT + 56, 176
    cv.shape("Navy band", (0, top, G.width, band), fill=NAVY)
    for i, (value, label, note) in enumerate(FIGURES):
        x, w = G.x(3 * i), G.w(3)
        if i:
            cv.line(f"Key figure divider {i}", x - G.gutter / 2, top + 28, x - G.gutter / 2, top + band - 28, WHITE,
                    0.5)
        cv.text(f"Key figure {i + 1}", (x, top + 24, w, 64), figure(value, 44, LIME if i == 0 else WHITE), anchor="b")
        cv.text(f"Key figure label {i + 1}", (x, top + 100, w, 48), P(label, 14, WHITE))
        cv.text(f"Key figure note {i + 1}", (x, top + band + 20, w, CB - top - band - 20), P(note, 12))
    source(cv, "Source: organisation, title of the publication (year). Every figure needs a source.")


TABLE = [("Green bonds", "1,250", "1,480", "18", "Figures right-aligned, with a thousands separator"),
         ("Sustainability-linked loans", "860", "910", "6", "Units in the column heading, not in each cell"),
         ("Blended finance", "420", "510", "21", "The % sign in tables, \u2018per cent\u2019 in sentences"),
         ("Guarantees", "180", "240", "33", "Light grey rules between rows, no vertical lines"),
         ("Results-based finance", "95", "120", "26", "Order the rows so that the order means something"),
         ("Total", "2,805", "3,260", "16", "Totals in bold, on the off-white row")]


def s_table(cv: Canvas, a: Assets):
    title(cv, "Action title: the comparison the table makes, in one sentence")
    num = 112
    widths = [G.w(3), num, num, num, G.width - G.margin_left - G.margin_right - G.w(3) - 3 * num]
    head = ["Instrument", "2024 (USD m)", "2025 (USD m)", "Change (%)", "How the table is set"]
    rows = [(44, [tc(P(t, 12, WHITE, bold=True, align="r" if 0 < j < 4 else "l"), NAVY, margins=(8, 4, 8, 4))
                  for j, t in enumerate(head)])]
    for i, cells in enumerate(TABLE):
        total = i == len(TABLE) - 1
        lines = {"T": (NAVY, 1), "B": (GREY, 0.75)} if total else {"B": (GREY, 0.75)}
        rows.append((44, [tc(P(t, 12, CHARCOAL if j == 4 else BLACK, bold=total and j < 4,
                               align="r" if 0 < j < 4 else "l"),
                             OFFWHITE if total else None, lines, margins=(8, 4, 8, 4)) for j, t in enumerate(cells)]))
    cv.table("Table", G.margin_left, CT, widths, rows)
    source(cv, "Source: organisation, dataset (year). Illustrative figures.")


STEPS = [("Diagnose", "Weeks 1 to 4", ["Frame the question and the decision it informs",
                                       "Collect the data and interview stakeholders",
                                       "Test early findings with the client team"], "Diagnostic note"),
         ("Design", "Weeks 5 to 8", ["Develop the options and test them with the client",
                                     "Agree the criteria for choosing between them",
                                     "Shortlist two or three options to detail"], "Options paper"),
         ("Structure", "Weeks 9 to 12", ["Detail the preferred option: instruments, terms and risks",
                                         "Model the costs and the benefits", "Draft the term sheet for investors"],
          "Structuring report"),
         ("Mobilise", "Weeks 13 to 16", ["Engage investors and partners", "Plan the launch and the first 100 days",
                                         "Set up the reporting on results"], "Implementation plan")]


def s_process(cv: Canvas, a: Assets):
    title(cv, "Action title: how the work moves from question to answer in four steps")
    depth, h, pad = 16, 52, 14
    for i, (step, when, items, output) in enumerate(STEPS):
        x, w = G.x(3 * i), G.w(3)
        reach = G.gutter - 4 if i < len(STEPS) - 1 else 0
        cv.shape(f"Step {i + 1}", (x, CT, w + reach, h), geom="homePlate" if i == 0 else "chevron",
                 adj={"adj": int(depth / h * 100000)}, fill=NAVY, anchor="ctr", inset=(depth, 0, depth, 0),
                 text=coloured([(f"{i + 1} ", LIME), (step, WHITE)], 16, "ctr"))
        top = CT + h + 12
        cv.shape(f"Step panel {i + 1}", (x, top, w, CB - top), fill=OFFWHITE)
        cv.text(f"Step timing {i + 1}", (x + pad, top + 14, w - 2 * pad, 18), P(when, 12, CHARCOAL, bold=True))
        cv.text(f"Step text {i + 1}", (x + pad, top + 42, w - 2 * pad, 170),
                "".join(P(t, 12, bullet="square", after=6) for t in items))
        cv.line(f"Step rule {i + 1}", x + pad, CB - 50, x + w - pad, CB - 50, GREY20)
        cv.text(f"Step output {i + 1}", (x + pad, CB - 40, w - 2 * pad, 30), P(f"**Output:** {output}", 12, NAVY))


QUADRANTS = [("Plan carefully", "High impact, but it needs preparation"),
             ("Prioritise", "High impact and ready to start"),
             ("Deprioritise", "Little gain for the effort"),
             ("Quick wins", "Easy to do, with modest gains")]
OPTIONS = [("A", "Option A: name it and say what it involves", (0.86, 0.33)),
           ("B", "Option B: one line each, in the same form", (0.66, 0.40)),
           ("C", "Option C: high impact, but it needs work first", (0.30, 0.36)),
           ("D", "Option D: easy, with a modest gain", (0.82, 0.82)),
           ("E", "Option E: say why it is not worth pursuing", (0.20, 0.84))]


def s_matrix(cv: Canvas, a: Assets):
    title(cv, "Action title: which options to pursue first, by impact and feasibility")
    x0, x1 = G.x(0) + 40, G.x(8) - G.gutter
    y0, y1 = CT, CB - 40
    gap = 8
    qw, qh = (x1 - x0 - gap) / 2, (y1 - y0 - gap) / 2
    for i, (head, text) in enumerate(QUADRANTS):
        qx, qy = x0 + (i % 2) * (qw + gap), y0 + (i // 2) * (qh + gap)
        hot = i == 1
        cv.shape(f"Quadrant {i + 1}", (qx, qy, qw, qh), fill=NAVY if hot else OFFWHITE, anchor="t",
                 inset=(16, 14, 16, 14), text=P(head, 16, WHITE if hot else NAVY, HEAD, bold=True, after=2)
                 + P(text, 12, WHITE if hot else BLACK))
    cv.line("Impact axis", G.x(0) + 28, y1, G.x(0) + 28, y0, CHARCOAL, 1, arrow=True)
    cv.text("Impact label", (G.x(0), y0, 20, y1 - y0), P("Impact", 12, CHARCOAL, bold=True, align="ctr"),
            anchor="ctr", vert="vert270")
    cv.line("Feasibility axis", x0, y1 + 12, x1, y1 + 12, CHARCOAL, 1, arrow=True)
    cv.text("Feasibility label", (x0, y1 + 18, x1 - x0, 20), P("Feasibility", 12, CHARCOAL, bold=True, align="ctr"))
    d = 24
    for letter, _, (fx, fy) in OPTIONS:
        cx, cy = x0 + fx * (x1 - x0), y0 + fy * (y1 - y0)
        cv.shape(f"Option {letter}", (cx - d / 2, cy - d / 2, d, d), geom="ellipse", fill=GREEN, line=WHITE,
                 line_w=1.5, anchor="ctr", text=P(letter, 12, WHITE, bold=True, align="ctr"))
    x, w = G.x(8), G.w(4)
    cv.text("Options heading", (x, CT, w, 24), P("Options assessed", 16, NAVY, HEAD, bold=True))
    for i, (letter, text, _) in enumerate(OPTIONS):
        y = CT + 48 + i * 52
        cv.shape(f"Option key {letter}", (x, y, d, d), geom="ellipse", fill=GREEN, anchor="ctr",
                 text=P(letter, 12, WHITE, bold=True, align="ctr"))
        cv.text(f"Option text {letter}", (x + 36, y + 3, w - 36, 44), P(text, 12))


PHASES = ["Phase 1: diagnostic", "Phase 2: design", "Phase 3: delivery"]
MILESTONES = [("January 2027", "Kick-off", "Inception meeting and data request", True),
              ("February 2027", "Inception report", "Approach, workplan and risks agreed", True),
              ("April 2027", "Diagnostic note", "Findings shared with the steering group", False),
              ("June 2027", "Options workshop", "Options tested with stakeholders", False),
              ("August 2027", "Draft report", "Recommendations shared for comment", False),
              ("October 2027", "Final report", "Final version and handover", False)]


def s_timeline(cv: Canvas, a: Assets):
    title(cv, "Action title: the milestones that matter between now and completion")
    for i, phase in enumerate(PHASES):
        cv.shape(f"Phase {i + 1}", (G.x(4 * i), CT, G.w(4), 28), fill=OFFWHITE, anchor="ctr", inset=(10, 0, 10, 0),
                 text=P(phase, 12, NAVY, bold=True))
    line_y = CT + 116
    cv.line("Timeline", G.margin_left, line_y, RIGHT, line_y, NAVY, 1.5)
    d = 14
    for i, (date, head, text, done) in enumerate(MILESTONES):
        x, w = G.x(2 * i), G.w(2)
        cv.text(f"Milestone date {i + 1}", (x, line_y - 40, w, 22), P(date, 12, CHARCOAL, bold=True), anchor="b")
        cv.shape(f"Milestone {i + 1}", (x, line_y - d / 2, d, d), geom="ellipse", fill=GREEN if done else WHITE,
                 line=GREEN, line_w=2)
        cv.text(f"Milestone title {i + 1}", (x, line_y + 22, w, 20), P(head, 14, NAVY, bold=True))
        cv.text(f"Milestone text {i + 1}", (x, line_y + 46, w, 60), P(text, 12))
    ly = line_y + 128
    cv.shape("Legend completed", (G.margin_left, ly, 10, 10), geom="ellipse", fill=GREEN, line=GREEN, line_w=1.5)
    cv.text("Legend completed text", (G.margin_left + 16, ly - 3, 80, 16), P("Completed", 10, CHARCOAL))
    cv.shape("Legend planned", (G.margin_left + 100, ly, 10, 10), geom="ellipse", fill=WHITE, line=GREEN, line_w=1.5)
    cv.text("Legend planned text", (G.margin_left + 116, ly - 3, 80, 16), P("Planned", 10, CHARCOAL))
    callout(cv, (G.x(0), CB - 64, G.w(12), 64), "Critical path",
            "Name the dependency that would delay the final milestone, who owns it and when it must be resolved.")


WORKPLAN = [("1 Inception", None, None),
            ("Kick-off meeting and data request", (1, 1), None),
            ("Inception report", (2, 3), 3),
            ("2 Diagnostic", None, None),
            ("Desk review and stakeholder interviews", (3, 7), None),
            ("Diagnostic note", (7, 8), 8),
            ("3 Options and recommendations", None, None),
            ("Options analysis and financial modelling", (8, 10), None),
            ("Validation workshop", (11, 11), 11),
            ("Final report", (11, 12), 12)]


def s_workplan(cv: Canvas, a: Assets):
    title(cv, "Action title: the workplan delivers the final report in week 12")
    weeks = 12
    first = G.x(4) - G.margin_left
    week_w = (G.width - G.margin_left - G.margin_right - first) / weeks
    white = (WHITE, 1)
    head1 = [tc(P("Workstream and activity", 11, WHITE, bold=True), NAVY, {"B": white}, rowspan=2, anchor="b",
                margins=(8, 4, 8, 6))]
    for m in range(3):
        head1.append(tc(P(f"Month {m + 1}", 11, WHITE, bold=True, align="ctr"), NAVY, {"L": white, "B": white},
                        span=4))
        head1 += [tc(P(""), NAVY, hmerge=True) for _ in range(3)]
    head2 = [tc(P(""), NAVY, vmerge=True)]
    head2 += [tc(P(str(n), 10, WHITE, align="ctr"), NAVY, {"L": white} if n in (1, 5, 9) else None,
                 margins=(0, 2, 0, 2)) for n in range(1, weeks + 1)]
    rows = [(24, head1), (22, head2)]
    for label, span, due in WORKPLAN:
        group = span is None
        cells = [tc(P(label, 11, NAVY if group else BLACK, bold=group), OFFWHITE if group else None,
                    {"B": (GREY, 0.75)}, margins=(8 if group else 22, 3, 8, 3))]
        for n in range(1, weeks + 1):
            fill = OFFWHITE if group else None
            if span and span[0] <= n <= span[1]:
                fill = NAVY if n == due else GREEN
            lines = {"B": (GREY, 0.75)}
            if n in (5, 9) and fill not in (GREEN, NAVY):
                lines["L"] = (GREY, 0.75)
            cells.append(tc(P(""), fill, lines, margins=(0, 3, 0, 3)))
        rows.append((28, cells))
    cv.table("Workplan", G.margin_left, CT, [first] + [week_w] * weeks, rows)
    ly = CT + 24 + 22 + 28 * len(WORKPLAN) + 14
    cv.shape("Legend activity", (G.margin_left, ly, 16, 10), fill=GREEN)
    cv.text("Legend activity text", (G.margin_left + 22, ly - 3, 80, 16), P("Activity", 10, CHARCOAL))
    cv.shape("Legend deliverable", (G.margin_left + 100, ly, 16, 10), fill=NAVY)
    cv.text("Legend deliverable text", (G.margin_left + 122, ly - 3, 120, 16), P("Deliverable due", 10, CHARCOAL))


ICONS = [("landmark", "Public finance", "One sentence on what the capability covers and who it serves."),
         ("leaf", "Nature and climate", "Choose each icon to name the topic: icons are never decoration."),
         ("handshake", "Partnerships", "Use Lucide icons in green on white, all at the same size."),
         ("trending-up", "Capital markets", "Keep titles to two or three words, in sentence case."),
         ("shield-check", "Risk and safeguards", "Keep the sentences parallel and about the same length."),
         ("graduation-cap", "Capacity building", "Six items at most; use three or four when each needs more text.")]


def s_icons(cv: Canvas, a: Assets):
    title(cv, "Action title: six capabilities that cover the client\u2019s needs end to end")
    gap = 16
    rh = (CH - gap) / 2
    for i, (icon, head, text) in enumerate(ICONS):
        x, w, y = G.x(4 * (i % 3)), G.w(4), CT + (i // 3) * (rh + gap)
        cv.line(f"Item rule {i + 1}", x, y, x + w, y, GREY)
        png, svg = a.icon(icon, GREEN)
        cv.picture(f"Icon {i + 1}", png, (x, y + 20, 40, 40), f"{head} icon", svg)
        cv.text(f"Item title {i + 1}", (x, y + 76, w, 22), P(head, 16, NAVY, HEAD, bold=True))
        cv.text(f"Item text {i + 1}", (x, y + 104, w, rh - 108), P(text, 12))


def s_image_text(cv: Canvas, a: Assets):
    title(cv, "Action title: one message, with a photo that shows its subject")
    w = G.w(6)
    cv.text("Subheading", (G.x(0), CT, w, 52),
            P("Use a photo that shows the place, the people or the asset in question", 18, NAVY, HEAD, bold=True,
              line=95))
    items = ["Choose from the BwB photo library: Photos in the BwB tab", "Crop the photo to its frame; never stretch it",
             "Keep text off the photo and credit it in the source line", "Add alt text that says what the photo shows"]
    cv.text("Text", (G.x(0), CT + 72, w, 180), "".join(P(t, 14, bullet="square", after=8) for t in items))
    callout(cv, (G.x(0), CB - 64, w, 64), "Key point",
            "What the reader should take from the photo and the text, in a line or two.")
    rid = a.image(photo("hero-solar-woman", w / CH, 0.68, 0.5, 1.0, 960))
    cv.picture("Photo", rid, (G.x(6), CT, w, CH), "A person standing beside a solar farm at dusk")
    source(cv, "Photo: solar farm at dusk, BwB image library. Credit every photo that is not BwB\u2019s own.")
