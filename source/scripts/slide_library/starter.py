"""The starter deck that "New BwB deck" opens: the BwB master with a blank title slide and a blank title and
content slide.
"""
from .assets import Assets
from .catalogue import Spec
from .drawing import Canvas


def s_starter_title(cv: Canvas, a: Assets):
    for name, ph in (("Title", 'type="ctrTitle"'), ("Subtitle", 'type="subTitle" idx="1"'),
                     ("Date", 'type="body" sz="quarter" idx="13"')):
        cv.placeholder(name, ph)


def s_starter_content(cv: Canvas, a: Assets):
    cv.placeholder("Title", 'type="title"')
    cv.placeholder("Content", 'idx="1"')


# The starter deck's placeholders are empty, so PowerPoint shows the layouts' prompts, such as the action title prompt.
STARTER = [
    Spec(256, "starter-title", "BwB Title", s_starter_title, "Title", "Structure", "", []),
    Spec(257, "starter-content", "BwB Title and content", s_starter_content, "Title and content", "Messages", "", []),
]
