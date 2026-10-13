"""The library catalogue: each slide's fixed sldId, its layout, the function that draws it and the text the
add-in's index shows.
"""
from .chart import s_chart
from .drawing import creation_id
from .slides_messages import s_bullets, s_case, s_quote, s_summary, s_team, s_two_columns
from .slides_structure import s_agenda, s_closing, s_disclaimer, s_divider, s_title
from .slides_visuals import s_icons, s_image_text, s_key_figures, s_matrix, s_process, s_table, s_timeline, s_workplan


CATEGORIES = ("Structure", "Messages", "Data", "Visuals", "People", "Closing")


class Spec:
    def __init__(self, sid: int, slide_id: str, layout: str, build, name: str, category: str, description: str,
                 tags: list):
        self.sid, self.id, self.layout, self.build = sid, slide_id, layout, build
        self.title, self.category, self.description, self.tags = name, category, description, tags

    @property
    def slide_ref(self) -> str:
        return f"{self.sid}#{creation_id('slide/' + self.id)}"


# The sldId of each slide is fixed here, not by its position, so reordering the library keeps every slide's id.
SLIDES = [
    Spec(256, "title", "BwB Title", s_title, "Title", "Structure",
         "Opening slide on navy with the logo, the presentation title, a subtitle and the date.",
         ["title", "cover", "opening", "navy"]),
    Spec(257, "agenda", "BwB Title only", s_agenda, "Agenda", "Structure",
         "Five numbered sections, each with a one-line description.", ["agenda", "contents", "sections", "overview"]),
    Spec(258, "section-divider", "BwB Section divider", s_divider, "Section divider", "Structure",
         "Navy divider with a large lime section number, the section title and a line of context.",
         ["divider", "section", "chapter", "navy"]),
    Spec(259, "executive-summary", "BwB Title only", s_summary, "Executive summary", "Messages",
         "Three key messages, each with a green marker, a headline and the evidence behind it.",
         ["summary", "key messages", "findings", "answer first"]),
    Spec(260, "bullets", "BwB Title and content", s_bullets, "Text and callout", "Messages",
         "An action title, two levels of bullets and a callout for the key takeaway.",
         ["text", "bullets", "callout", "takeaway"]),
    Spec(261, "two-columns", "BwB Two content", s_two_columns, "Two columns", "Messages",
         "Two columns with headings and bullets, for options or periods side by side, and a recommendation.",
         ["comparison", "options", "before and after", "columns"]),
    Spec(262, "key-figures", "BwB Title only", s_key_figures, "Key figures", "Data",
         "Four headline numbers on a navy band, the first in lime, with a note under each.",
         ["figures", "statistics", "numbers", "kpi", "navy"]),
    Spec(263, "chart-and-text", "BwB Title only", s_chart, "Chart and text", "Data",
         "A native column chart in the BwB colours, ready for Edit Data, with a panel on what it shows.",
         ["chart", "graph", "data", "commentary"]),
    Spec(264, "table", "BwB Title only", s_table, "Table", "Data",
         "A data table with a navy header row, right-aligned figures and a total row.",
         ["table", "figures", "comparison", "data"]),
    Spec(265, "process", "BwB Title only", s_process, "Process", "Visuals",
         "Four chevrons for the steps of a process, with the timing, activities and output of each.",
         ["process", "steps", "approach", "methodology", "chevrons"]),
    Spec(266, "matrix", "BwB Title only", s_matrix, "Two-by-two matrix", "Visuals",
         "Four quadrants on impact and feasibility, with the options plotted and listed.",
         ["matrix", "prioritisation", "quadrants", "options"]),
    Spec(267, "timeline", "BwB Title only", s_timeline, "Timeline", "Visuals",
         "Six milestones on a line, grouped into phases, with the critical path below.",
         ["timeline", "milestones", "roadmap", "phases"]),
    Spec(268, "workplan", "BwB Title only", s_workplan, "Workplan", "Data",
         "A Gantt-style table of activities over twelve weeks, with deliverables marked.",
         ["workplan", "gantt", "schedule", "plan"]),
    Spec(269, "icon-grid", "BwB Title only", s_icons, "Icon grid", "Visuals",
         "Six items in a grid, each with a green Lucide icon, a title and a sentence.",
         ["icons", "capabilities", "services", "grid"]),
    Spec(270, "image-and-text", "BwB Title only", s_image_text, "Image and text", "Visuals",
         "A photo from the BwB library beside a heading, bullets and a key point.", ["photo", "image", "picture", "text"]),
    Spec(271, "case-study", "BwB Title only", s_case, "Case study", "Messages",
         "A navy panel naming the case, the challenge and the approach, and four results.",
         ["case study", "example", "project", "results", "credentials"]),
    Spec(272, "quote", "BwB Title only", s_quote, "Quote", "People",
         "One quotation in Candara italic on a navy panel, with the speaker\u2019s name and role.",
         ["quote", "testimonial", "stakeholder", "interview"]),
    Spec(273, "team", "BwB Title only", s_team, "Team", "People",
         "Four photo frames with names, roles and a short profile.", ["team", "people", "staff", "profiles"]),
    Spec(274, "closing", "BwB Closing", s_closing, "Closing", "Closing",
         "Navy closing slide with two contacts and the BwB website.", ["closing", "contact", "thank you", "end"]),
    Spec(275, "disclaimer", "BwB Title only", s_disclaimer, "Disclaimer", "Closing",
         "The standard BwB disclaimer and the FCA note, in an off-white panel.",
         ["disclaimer", "legal", "compliance", "fca"]),
]
