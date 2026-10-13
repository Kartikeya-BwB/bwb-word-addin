"""Library slides for structure and closing: title, agenda, section divider, closing and disclaimer. Their text
is written as guidance for the author.
"""
from .assets import Assets
from .drawing import Canvas, P, PH
from .elements import title
from .sources import BODY, CH, CHARCOAL, CT, G, GREEN, GREY, GREY20, HEAD, NAVY, OFFWHITE, RIGHT, read_disclaimer


def s_title(cv: Canvas, a: Assets):
    cv.placeholder("Title", 'type="ctrTitle"', PH("Presentation Title in Title Case, on One or Two Lines"))
    cv.placeholder("Subtitle", 'type="subTitle" idx="1"',
                   PH("Subtitle: the client, the project and the purpose of the meeting"))
    cv.placeholder("Date", 'type="body" sz="quarter" idx="13"', PH("Date in the form 5 October 2026"))


AGENDA = [("Context and objectives", "Why we are here and the decision this meeting supports"),
          ("Findings", "What the diagnostic shows, with the evidence behind each point"),
          ("Options", "The options we assessed and how they compare"),
          ("Recommendations", "What we recommend, why, and what it will take"),
          ("Next steps", "Who does what by when, and the decisions we need today")]


def s_agenda(cv: Canvas, a: Assets):
    title(cv, "Agenda")
    row = CH / len(AGENDA)
    cv.line("Agenda rule top", G.margin_left, CT, RIGHT, CT, NAVY, 1)
    for i, (head, text) in enumerate(AGENDA):
        y = CT + i * row
        cv.text(f"Agenda number {i + 1}", (G.x(0), y, G.w(1), row), P(f"{i + 1:02d}", 24, GREEN, BODY, bold=True),
                anchor="ctr")
        cv.text(f"Agenda item {i + 1}", (G.x(1), y, G.w(5), row), P(head, 18, NAVY, HEAD, bold=True), anchor="ctr")
        cv.text(f"Agenda description {i + 1}", (G.x(6), y, G.w(6), row), P(text, 14, CHARCOAL), anchor="ctr")
        cv.line(f"Agenda rule {i + 1}", G.margin_left, y + row, RIGHT, y + row, GREY)


def s_divider(cv: Canvas, a: Assets):
    cv.placeholder("Section number", 'type="body" sz="quarter" idx="13"', PH("01"))
    cv.placeholder("Title", 'type="title"', PH("Section title in sentence case, as it appears on the agenda"))
    cv.placeholder("Description", 'type="body" sz="quarter" idx="14"',
                   PH("One sentence on what this section shows and why it matters to the reader"))


def s_closing(cv: Canvas, a: Assets):
    cv.placeholder("Title", 'type="title"', PH("Thank you"))
    for i in (1, 2):
        cv.placeholder(f"Contact {i}", f'type="body" idx="{i}"',
                       PH("**Name Surname**") + PH("Role") + PH("name.surname@bwb.earth") + PH("Phone number"))


def s_disclaimer(cv: Canvas, a: Assets):
    title(cv, "Disclaimer")
    paras, note = read_disclaimer()
    half = (len(paras) + 1) // 2
    pad, height = 24, 264
    cv.shape("Disclaimer panel", (G.x(0), CT, G.w(12), height), fill=OFFWHITE, line=GREY20, line_w=0.5)
    col = (G.w(12) - 3 * pad) / 2
    for i, group in enumerate((paras[:half], paras[half:])):
        body = "".join(P(t, 11, after=8, line=105) for t in group)
        if i:
            body += P(note, 11, CHARCOAL, line=105)
        cv.text(f"Disclaimer text {i + 1}", (G.x(0) + pad + i * (col + pad), CT + pad, col, height - 2 * pad), body)
