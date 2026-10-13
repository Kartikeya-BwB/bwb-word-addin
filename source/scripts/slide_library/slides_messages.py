"""Library slides for messages and people: executive summary, text and callout, two columns, case study, quote
and team. Their text is written as guidance for the author.
"""
from .assets import Assets, photo
from .drawing import Canvas, P, PH
from .elements import callout, figure, source, title
from .sources import BODY, CB, CH, CHARCOAL, CT, G, GREEN, GREY, HEAD, LIME, NAVY, RIGHT, WHITE


SUMMARY = [("State the answer to the client\u2019s question in one sentence",
            ["Put the conclusion first: a reader who stops here should still know what we recommend",
             "Use the client\u2019s own words for the question"]),
           ("Give the finding that matters most, with its evidence",
            ["Back it with one or two facts: a figure, a comparison or a trend",
             "Source each fact on the slide that shows it in full"]),
           ("Close with what the client should do next",
            ["Name the decision, the owner and the timing",
             "Keep each message to two lines, so the page reads in a minute"])]


def s_summary(cv: Canvas, a: Assets):
    title(cv, "Executive summary")
    gap = 16
    row = (CH - 2 * gap) / 3
    for i, (head, points) in enumerate(SUMMARY):
        y = CT + i * (row + gap)
        cv.shape(f"Message marker {i + 1}", (G.x(0), y, 32, 32), fill=GREEN, anchor="ctr",
                 text=P(str(i + 1), 16, WHITE, BODY, bold=True, align="ctr"))
        cv.text(f"Message {i + 1}", (G.x(1), y, G.w(5), row), P(head, 18, NAVY, HEAD, bold=True, line=95))
        cv.text(f"Message detail {i + 1}", (G.x(6), y + 2, G.w(6), row),
                "".join(P(t, 14, bullet="square", after=6) for t in points))
        if i < 2:
            cv.line(f"Message rule {i + 1}", G.margin_left, y + row + gap / 2, RIGHT, y + row + gap / 2, GREY)


def s_bullets(cv: Canvas, a: Assets):
    title(cv, "Action title: say the so-what in one sentence, on no more than two lines")
    items = [(0, "Lead with the point, then give the evidence for it"),
             (1, "Use the second level for supporting detail, at most three under a point"),
             (0, "Keep to five or six bullets of one or two lines each"),
             (0, "Write full sentences in plain English, with UK spelling and no filler"),
             (1, "Spell out an abbreviation the first time it appears"),
             (0, "Use figures where they make the point: GBP 5m, 38 per cent, 2025 to 2030"),
             (1, "Give the source of every figure in the source line at the foot of the slide"),
             (0, "Make the bullets parallel: start each with the same kind of word"),
             (0, "Bold only the **few words** the reader must not miss")]
    cv.placeholder("Content", 'idx="1"', "".join(PH(t, lvl) for lvl, t in items), (G.x(0), CT, G.w(8), CH))
    callout(cv, (G.x(8), CT, G.w(4), 92), "Key takeaway",
            "Use a callout for the one message the slide exists to make. Keep it to three lines.")
    source(cv, "Source: organisation, title of the publication (year). Notes go on the same line, in Aptos 8 pt "
               "charcoal.")


def s_two_columns(cv: Canvas, a: Assets):
    title(cv, "Action title: compare two options, periods or groups side by side")
    cols = [("Option A: name it in a few words",
             ["Say what the option involves, in one line", "Give its main benefit, with a figure if there is one",
              "Give its main risk or cost", "Say what it would take to deliver"]),
            ("Option B: name it in a few words",
             ["Match the left column point for point, so the reader can compare across",
              "Use the same units and periods in both columns", "Keep the points parallel and about the same length",
              "Leave the verdict to the box below"])]
    box = 64
    for i, (head, items) in enumerate(cols):
        x, w = G.x(6 * i), G.w(6)
        cv.text(f"Column heading {i + 1}", (x, CT, w, 28), P(head, 18, NAVY, HEAD, bold=True), anchor="b")
        cv.line(f"Column rule {i + 1}", x, CT + 38, x + w, CT + 38, GREEN, 1.5)
        cv.placeholder("Left content" if i == 0 else "Right content", f'sz="half" idx="{i + 1}"',
                       "".join(PH(t) for t in items), (x, CT + 56, w, CH - 56 - box - 16))
    callout(cv, (G.x(0), CB - box, G.w(12), box), "Recommendation",
            "Say which option we recommend and the one reason that decides it.")


CASE = [("The challenge", "Two or three sentences on the client\u2019s situation and the question they faced. Name the "
                          "constraint that made it hard."),
        ("What we did", "The approach in two or three sentences: the analysis, the structure we designed and the "
                        "partners we brought in.")]
RESULTS = [("USD 120m", "mobilised from private investors"), ("3x", "leverage on the public funds committed"),
           ("18", "months from mandate to financial close"), ("4", "countries covered by the programme")]


def s_case(cv: Canvas, a: Assets):
    title(cv, "Action title: what the case shows for the client\u2019s decision")
    px, pw, pad = G.x(0), G.w(4), 20
    cv.shape("Case panel", (px, CT, pw, CH), fill=NAVY)
    cv.text("Case label", (px + pad, CT + 24, pw - 2 * pad, 16), P("Case study", 12, LIME, bold=True))
    cv.text("Case title", (px + pad, CT + 48, pw - 2 * pad, 64),
            P("Project or client name, on one or two lines", 24, WHITE, HEAD, bold=True, line=95))
    cv.text("Case details", (px + pad, CT + 128, pw - 2 * pad, 80),
            P("Country, year", 14, WHITE, after=4) + P("Sector and instrument", 14, WHITE, after=4)
            + P("Client: sovereign, city or company", 14, WHITE))
    png, svg = a.icon("landmark", WHITE)
    cv.picture("Case icon", png, (px + pad, CB - pad - 40, 40, 40), "Landmark icon", svg)
    for i, (head, text) in enumerate(CASE):
        x, w = G.x(4 + 4 * i), G.w(4)
        cv.text(f"Case heading {i + 1}", (x, CT, w, 24), P(head, 16, NAVY, HEAD, bold=True))
        cv.text(f"Case text {i + 1}", (x, CT + 32, w, 120), P(text, 12))
    rule = CT + 164
    cv.line("Results rule", G.x(4), rule, RIGHT, rule, GREY)
    cv.text("Results heading", (G.x(4), rule + 14, G.w(8), 24), P("Results", 16, NAVY, HEAD, bold=True))
    for i, (value, label) in enumerate(RESULTS):
        x, w = G.x(4 + 2 * i), G.w(2)
        cv.text(f"Result {i + 1}", (x, rule + 48, w, 44), figure(value, 32, GREEN), anchor="b")
        cv.text(f"Result label {i + 1}", (x, rule + 100, w, 60), P(label, 12))
    source(cv, "Source: BwB project records (year). Use figures the client has agreed we can publish.")


def s_quote(cv: Canvas, a: Assets):
    title(cv, "Action title: what stakeholders told us, in their own words")
    cv.shape("Quote panel", (G.x(0), CT, G.w(12), CH), fill=NAVY)
    x, w = G.x(1), G.w(10)
    cv.text("Quote mark", (x, CT + 20, 80, 80), P("\u201C", 120, LIME, HEAD, bold=True, line=100))
    cv.text("Quote", (x, CT + 116, w, 130),
            P("Put here the one quotation that carries the point, in the speaker\u2019s own words, cut to two or "
              "three lines.", 28, WHITE, HEAD, italic=True, line=100))
    cv.text("Attribution", (x, CB - 92, w, 56),
            P("**Name Surname**, role, organisation", 14, WHITE, after=4)
            + P("Interview, October 2026. Quote with permission, or describe the speaker: \u2018a senior official at "
                "the ministry of finance\u2019.", 12, WHITE), anchor="b")


TEAM = [("team-discussion", (0.583, 0.242, 0.281), "Engagement lead",
         "Accountable for quality and for the client relationship. Give years of experience and two similar projects."),
        ("team-interview", (0.365, 0.461, 0.271), "Project manager",
         "Runs the workplan day to day. Name the sectors and the regions they know best."),
        ("team-meeting-table", (0.763, 0.466, 0.280), "Sector expert",
         "Brings the technical depth. Say what they have done that this project needs."),
        ("team-workshop", (0.659, 0.313, 0.313), "Financial analyst",
         "Builds the models and the evidence base. Keep every profile to the same length.")]


def s_team(cv: Canvas, a: Assets):
    title(cv, "Action title: the team brings the mix of skills this project needs")
    size = G.w(3)
    for i, (image, (cx, cy, scale), role, text) in enumerate(TEAM):
        x = G.x(3 * i)
        rid = a.image(photo(image, 1.0, cx, cy, scale, 480))
        cv.picture(f"Photo {i + 1}", rid, (x, CT, size, size), "Team member photo: replace with their portrait")
        cv.text(f"Name {i + 1}", (x, CT + size + 14, size, 22), P("Name Surname", 16, NAVY, HEAD, bold=True))
        cv.text(f"Role {i + 1}", (x, CT + size + 38, size, 18), P(role, 12, CHARCOAL))
        cv.text(f"Profile {i + 1}", (x, CT + size + 64, size, CB - CT - size - 64), P(text, 12))
