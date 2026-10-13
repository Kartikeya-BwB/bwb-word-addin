"""The chart and text slide: a native column chart in the BwB style, with the workbook behind it so that
authors can right-click it and choose Edit Data.
"""
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.oxml import parse_xml
from pptx.util import Pt

from .assets import Assets
from .drawing import Canvas, P
from .elements import source, title
from .sources import BODY, CH, CHARCOAL, CT, G, GREEN, GREEN50, GREY20, HEAD, NAVY, OFFWHITE


CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def style_chart(chart):
    """BwB chart style: green and green 50 series, Aptos labels in charcoal, no title, no gridlines, no frame."""
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.TOP
    chart.legend.include_in_layout = False
    chart.font.name = BODY
    chart.font.size = Pt(10)
    chart.font.color.rgb = RGBColor.from_string(CHARCOAL)
    plot = chart.plots[0]
    plot.gap_width = 60
    plot.overlap = -8
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.position = XL_LABEL_POSITION.OUTSIDE_END
    labels.number_format = "#,##0"
    labels.number_format_is_linked = False
    labels.font.size = Pt(10)
    labels.font.color.rgb = RGBColor.from_string(CHARCOAL)
    for series, colour in zip(plot.series, (GREEN, GREEN50)):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = RGBColor.from_string(colour)
        series.format.line.fill.background()
    value = chart.value_axis
    value.visible = False
    value.has_major_gridlines = False
    value.major_tick_mark = XL_TICK_MARK.NONE
    category = chart.category_axis
    category.major_tick_mark = XL_TICK_MARK.NONE
    category.format.line.color.rgb = RGBColor.from_string(GREY20)
    category.tick_labels.font.size = Pt(10)
    category.tick_labels.font.color.rgb = RGBColor.from_string(CHARCOAL)
    space = chart._chartSpace

    def tag(name: str) -> str:
        return f"{{{CHART_NS}}}{name}"

    for axis in space.iter(tag("axId"), tag("crossAx")):
        axis.set("val", axis.get("val", "").lstrip("-"))  # python-pptx writes negative ids; the schema wants unsigned
    for hidden in space.iter(tag("delete")):
        hidden.set("val", hidden.get("val", "1"))  # say "hidden" outright rather than rely on the schema default
    lang = space.find(tag("lang"))
    if lang is not None:
        lang.set("val", "en-GB")
    if space.find(tag("roundedCorners")) is None:
        anchor = lang if lang is not None else space.find(tag("date1904"))
        corners = parse_xml(f'<c:roundedCorners xmlns:c="{CHART_NS}" val="0"/>')
        if anchor is not None:
            anchor.addnext(corners)
        else:
            space.insert(0, corners)
    if space.find(tag("spPr")) is None:
        space.find(tag("chart")).addnext(parse_xml(
            f'<c:spPr xmlns:c="{CHART_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            "<a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>"))


def s_chart(cv: Canvas, a: Assets):
    title(cv, "Action title: the conclusion the chart supports, not a description of it")
    w8 = G.w(8)
    cv.text("Chart heading", (G.x(0), CT, w8, 20),
            P("Chart heading: what is measured, the unit and the period (USD m, 2021 to 2025)", 12, NAVY, bold=True))
    data = CategoryChartData(number_format="#,##0")
    data.categories = ["2021", "2022", "2023", "2024", "2025"]
    data.add_series("Committed", (120, 150, 185, 210, 260))
    data.add_series("Disbursed", (80, 105, 140, 170, 205))
    rid, chart = a.chart(XL_CHART_TYPE.COLUMN_CLUSTERED, data)
    style_chart(chart)
    cv.chart("Chart", rid, (G.x(0), CT + 28, w8, CH - 28),
             "Column chart of two series over five years. Right-click it and choose Edit Data to use your own figures")
    x, w = G.x(8), G.w(4)
    points = ["Name the pattern: the trend, the gap or the outlier", "Give the number that proves it, read from the chart",
              "Say what it means for the decision in hand",
              "Right-click the chart and choose Edit Data to put in your own figures"]
    cv.shape("Commentary", (x, CT, w, CH), fill=OFFWHITE, anchor="t", inset=(20, 20, 20, 20),
             text=P("What the chart shows", 16, NAVY, HEAD, bold=True, after=12)
             + "".join(P(t, 12, bullet="square", after=8) for t in points))
    source(cv, "Source: organisation, dataset (year). Charts use green and green 50 on white, with no title inside "
               "the chart.")
