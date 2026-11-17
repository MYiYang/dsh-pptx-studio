# -*- coding: utf-8 -*-
"""Render a deck.json model into a real .pptx with python-pptx.

Usage:
  python pptx_build.py <deck.json> <out.pptx> [--template <template.pptx>] [--no-effects]

The deck model is the contract shared with the DSH plugin host and its client
canvas; see REQUIREMENTS.md for the schema. Everything here is deliberately
small and dependency-free beyond python-pptx, because this script runs inside
the user's DSH process on Windows.

Beyond plain shapes this writer emits the features a .pptx is expected to have:
  - slide transitions          -> <p:transition> (fade / push / wipe / …)
  - entrance animations        -> <p:timing> (appear / fade / wipe / zoom)
  - speaker notes              -> notes slide text
  - per-slide background       -> solid fill
  - hidden slides              -> show="0"
  - text shadows               -> <a:outerShdw>
  - native charts              -> python-pptx chart parts
"""
import json
import math
import os
import sys

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

# The theme is read once per build and then used as the fallback for every
# missing font/colour, so a deck declared in the model really looks themed.
FONT = "微软雅黑"
THEME_INK = "0F172A"
THEME_MUTED = "64748B"
THEME_ACCENT = "2563EB"
THEME_SURFACE = "F1F5F9"
THEME_BACKGROUND = "FFFFFF"

TRANSITION_TYPES = {
    "fade": "<p:fade/>",
    "cut": "<p:cut/>",
    "dissolve": "<p:dissolve/>",
    "circle": "<p:circle/>",
    "random": "<p:random/>",
}

CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE_MARKERS,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
}

ANIM_PRESETS = {
    "appear": (1, 0),
    "fade": (10, 0),
    "wipe": (22, 0),
    "zoom": (23, 0),
}

WIPE_FILTERS = {
    "left": "wipe(left)",
    "right": "wipe(right)",
    "up": "wipe(up)",
    "down": "wipe(down)",
}


def hexcolor(value, fallback=None):
    """Turn a model colour into an RGBColor (never throws)."""
    fallback = fallback or THEME_INK
    if not value:
        value = fallback
    value = str(value).strip().lstrip("#")
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    if len(value) == 8:
        # Canvas colours may carry an alpha channel (#RRGGBBAA); OOXML wants RGB.
        value = value[:6]
    try:
        return RGBColor.from_string(value.upper())
    except Exception:
        return RGBColor.from_string(fallback.lstrip("#"))


def set_font(run, name, size=None, bold=None, color=None, italic=None, underline=None):
    f = run.font
    if size is not None:
        f.size = Pt(float(size))
    if bold is not None:
        f.bold = bool(bold)
    if italic is not None:
        f.italic = bool(italic)
    if underline is not None:
        f.underline = bool(underline)
    if color is not None:
        f.color.rgb = hexcolor(color)
    f.name = name or FONT
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        latin.addnext(ea)
    ea.set("typeface", name or FONT)
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = rPr.makeelement(qn("a:cs"), {})
        ea.addnext(cs)
    cs.set("typeface", name or FONT)


def remove_all_slides(prs):
    """Drop every slide but keep masters/layouts/theme (template mode)."""
    xml_slides = prs.slides._sldIdLst
    for sld_id in list(xml_slides):
        rId = sld_id.get(qn("r:id"))
        prs.part.drop_rel(rId)
        xml_slides.remove(sld_id)


def emu_box(el):
    x, y, w, h = el.get("box", [0, 0, 1, 1])
    return Inches(float(x)), Inches(float(y)), Inches(float(w)), Inches(float(h))


def apply_shadow(shape, shadow):
    """Write <a:outerShdw> into the shape properties (best effort)."""
    if not shadow:
        return
    if shadow is True:
        shadow = {}
    if not isinstance(shadow, dict):
        return
    spPr = getattr(shape._element, "spPr", None)
    if spPr is None:
        return
    for old in spPr.findall(qn("a:effectLst")):
        spPr.remove(old)
    blur = int(float(shadow.get("blur", 4)) * 12700)
    distance = int(float(shadow.get("distance", 3)) * 12700)
    direction = int(float(shadow.get("direction", 90)) * 60000)
    color = str(shadow.get("color") or "#000000").lstrip("#")
    if len(color) == 8:
        color = color[:6]
    alpha = int(max(0.0, min(1.0, float(shadow.get("alpha", 0.35)))) * 100000)
    xml = (
        '<a:effectLst %s><a:outerShdw blurRad="%d" dist="%d" dir="%d" rotWithShape="0">'
        '<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr></a:outerShdw></a:effectLst>'
        % (nsdecls("a"), blur, distance, direction, color.upper(), alpha)
    )
    spPr.append(parse_xml(xml))


def no_default_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_text(slide, el):
    left, top, width, height = emu_box(el)
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.01)
    tf.margin_bottom = Inches(0.01)
    style = el.get("style") or {}
    anchor = str(style.get("anchor") or "top").lower()
    tf.vertical_anchor = {
        "top": MSO_ANCHOR.TOP,
        "middle": MSO_ANCHOR.MIDDLE,
        "bottom": MSO_ANCHOR.BOTTOM,
    }.get(anchor, MSO_ANCHOR.TOP)
    text = str(el.get("text") or "")
    paragraphs = el.get("paragraphs")
    if not paragraphs:
        paragraphs = [{"text": line} for line in text.split("\n")]
    first = True
    for para in paragraphs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY,
        }.get(str(para.get("align") or style.get("align") or "left").lower(), PP_ALIGN.LEFT)
        if para.get("lineSpacing"):
            p.line_spacing = float(para["lineSpacing"])
        elif style.get("lineSpacing"):
            p.line_spacing = float(style["lineSpacing"])
        if para.get("spaceAfter") is not None:
            p.space_after = Pt(float(para["spaceAfter"]))
        elif style.get("spaceAfter") is not None:
            p.space_after = Pt(float(style["spaceAfter"]))
        runs = para.get("runs")
        if runs:
            for r in runs:
                run = p.add_run()
                run.text = str(r.get("text") or "")
                set_font(
                    run,
                    str(r.get("font") or para.get("font") or style.get("font") or FONT),
                    r.get("size", para.get("size", style.get("size", 18))),
                    r.get("bold", para.get("bold", style.get("bold", False))),
                    r.get("color", para.get("color", style.get("color") or THEME_INK)),
                    r.get("italic", para.get("italic", style.get("italic", False))),
                    r.get("underline", para.get("underline", style.get("underline", False))),
                )
        else:
            run = p.add_run()
            run.text = str(para.get("text") or "")
            set_font(
                run,
                str(para.get("font") or style.get("font") or FONT),
                para.get("size", style.get("size", 18)),
                para.get("bold", style.get("bold", style.get("bold", False))),
                para.get("color", style.get("color") or THEME_INK),
                para.get("italic", style.get("italic", False)),
                para.get("underline", style.get("underline", False)),
            )
    no_default_shadow(box)
    apply_shadow(box, el.get("shadow"))
    return box


def add_rect(slide, el):
    left, top, width, height = emu_box(el)
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if el.get("radius") else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if el.get("radius"):
        try:
            shape.adjustments[0] = float(el["radius"])
        except Exception:
            pass
    fill = el.get("fill")
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = hexcolor(fill, "FFFFFF")
    line = el.get("line")
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = hexcolor(line, "E2E8F0")
        shape.line.width = Pt(float(el.get("lineWidth") or 1))
    no_default_shadow(shape)
    apply_shadow(shape, el.get("shadow"))
    text = el.get("text")
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        style = el.get("style") or {}
        p = tf.paragraphs[0]
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(
            str(style.get("align") or "left").lower(), PP_ALIGN.LEFT)
        run = p.add_run()
        run.text = str(text)
        set_font(
            run,
            str(style.get("font") or FONT),
            style.get("size", 14),
            style.get("bold", False),
            style.get("color") or THEME_INK,
            style.get("italic", False),
            style.get("underline", False),
        )
    return shape


def add_image(slide, el, base_dir):
    left, top, width, height = emu_box(el)
    src = str(el.get("src") or "")
    path = src if os.path.isabs(src) else os.path.join(base_dir, src)
    if not path or not os.path.exists(path):
        return add_rect(slide, {**el, "fill": el.get("fill") or "EEF2F7", "line": "CBD5E1",
                                "text": el.get("text") or "[image]"})
    picture = slide.shapes.add_picture(path, left, top, width, height)
    apply_shadow(picture, el.get("shadow"))
    return picture


def add_line(slide, el):
    left, top, width, height = emu_box(el)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, max(height, Inches(0.01)))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hexcolor(el.get("line") or el.get("fill") or "E2E8F0")
    shape.line.fill.background()
    no_default_shadow(shape)
    apply_shadow(shape, el.get("shadow"))
    return shape


def add_table(slide, el):
    left, top, width, height = emu_box(el)
    data = el.get("table") or {}
    rows = data.get("rows") or [[""]]
    cols = max(len(r) for r in rows) if rows else 1
    shape = slide.shapes.add_table(len(rows), cols, left, top, width, height)
    table = shape.table
    style = el.get("style") or {}
    size = float(style.get("size") or 11)
    for r, row in enumerate(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = ""
            cell.margin_left = Inches(0.05)
            cell.margin_right = Inches(0.05)
            cell.margin_top = Inches(0.02)
            cell.margin_bottom = Inches(0.02)
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = str(row[c]) if c < len(row) else ""
            set_font(
                run,
                str(style.get("font") or FONT),
                size,
                r == 0,
                style.get("color") or THEME_INK,
                False,
                False,
            )
            if r == 0:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = hexcolor(THEME_SURFACE, "F1F5F9")
                except Exception:
                    pass
    return shape


# CT_ChartSpace child order (ECMA-376 part 1, 21.2.2.29): spPr must sit
# between c:chart and c:txPr or PowerPoint rejects the chart part.
_CHART_SPACE_ORDER = (
    "c:date1904", "c:lang", "c:roundedCorners", "c:style", "c:clrMapOvr",
    "c:pivotSource", "c:protection", "c:chart", "c:spPr", "c:txPr",
    "c:externalData", "c:printSettings", "c:userShapes", "c:extLst",
)


def axis_scale(numbers):
    """Nice round value-axis bounds covering the data.

    The canvas draws the very same scale (client.js#axisScale), so pinning the
    exported chart to it is what keeps the file and the preview identical —
    PowerPoint's own automatic axis would pick different bounds and the bars
    would land at different heights on the two sides.
    @param {list[float]} numbers - every data point of the chart.
    @returns {tuple} (bottom, top, step, ticks).
    """
    flat = [float(n) for n in (numbers or []) if n is not None]
    if not flat:
        flat = [0.0]
    lo = min(0.0, min(flat))
    hi = max(1.0, max(flat))
    span = hi - lo or 1.0
    raw = span / 4.0
    mag = 10.0 ** math.floor(math.log10(raw)) if raw > 0 else 1.0
    norm = raw / mag
    step = (1 if norm <= 1 else 2 if norm <= 2 else 2.5 if norm <= 2.5 else 5 if norm <= 5 else 10) * mag
    bottom = math.floor(lo / step) * step
    top = math.ceil(hi / step) * step
    ticks = []
    value = bottom
    while value <= top + step / 2.0:
        ticks.append(round(value, 3))
        value += step
    return bottom, top, step, ticks


def _luma(hex_text):
    """Rough perceptual brightness (0-255) of an #RRGGBB colour."""
    value = str(hex_text or "").strip().lstrip("#")
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    if len(value) < 6:
        return 255
    try:
        r, g, b = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
    except ValueError:
        return 255
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _slide_bg_hex(slide):
    """The slide's own background colour, falling back to the deck theme."""
    try:
        rgb = slide.background.fill.fore_color.rgb
        if rgb is not None:
            return str(rgb)
    except Exception:
        pass
    return THEME_BACKGROUND


def _transparent_spPr(parent, order=None):
    """Give a chart part a transparent fill; order keeps OOXML schema-valid."""
    existing = parent.find(qn("c:spPr"))
    if existing is not None:
        parent.remove(existing)
    spPr = parse_xml('<c:spPr %s><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>' % nsdecls("c", "a"))
    if order:
        tags = [qn(tag) for tag in order]
        pos = tags.index(qn("c:spPr"))
        for child in parent:
            if child.tag in tags and tags.index(child.tag) > pos:
                child.addprevious(spPr)
                break
        else:
            parent.append(spPr)
    else:
        parent.append(spPr)
    return spPr


def _rounded_bars(series, corner):
    """Round the top of a column/bar series (prstGeom roundRect on the series).

    prstGeom must precede the fill inside c:spPr, so it goes in first.
    """
    adj = int(max(0.0, min(0.5, float(corner))) * 100000)
    spPr = series._element.get_or_add_spPr()
    for tag in ("a:prstGeom", "a:custGeom"):
        old = spPr.find(qn(tag))
        if old is not None:
            spPr.remove(old)
    spPr.insert(0, parse_xml('<a:prstGeom %s prst="roundRect"><a:avLst><a:gd name="adj" fmla="val %d"/></a:avLst></a:prstGeom>' % (nsdecls("a"), adj)))


def _hide_axes(chart):
    """Drop both axes, the gridlines and the auto title: a design sheet that
    draws bare bars with value labels on top, no ruler around them."""
    try:
        for axis in (chart.category_axis, chart.value_axis):
            axis.visible = False
        chart.value_axis.has_major_gridlines = False
    except Exception:
        pass
    try:
        chart_el = chart._chartSpace.find(qn("c:chart"))
        auto = chart_el.find(qn("c:autoTitleDeleted"))
        if auto is None:
            auto = parse_xml('<c:autoTitleDeleted %s val="1"/>' % nsdecls("c"))
            plot_area = chart_el.find(qn("c:plotArea"))
            if plot_area is not None:
                plot_area.addprevious(auto)
            else:
                chart_el.append(auto)
        else:
            auto.set("val", "1")
    except Exception:
        pass


def theme_chart(chart, slide, kind, hide_axes=False):
    """Bend a native chart onto the deck theme.

    python-pptx charts arrive Office-default: black text, white chart space,
    theme-grid grey gridlines. On a dark slide that is a bright hole with
    invisible labels, so here the chart space and plot area go transparent,
    and text / axis / gridline colours follow the deck theme (light ink on a
    dark slide, dark ink on a light one).
    """
    dark = _luma(_slide_bg_hex(slide)) < 110
    text_hex = THEME_SURFACE if dark else THEME_INK
    axis_hex = THEME_MUTED
    try:
        chart.font.color.rgb = hexcolor(text_hex)
    except Exception:
        pass
    try:
        _transparent_spPr(chart._chartSpace, _CHART_SPACE_ORDER)
    except Exception:
        pass
    try:
        plot_area = chart.plots[0]._element.getparent()
        _transparent_spPr(plot_area)
    except Exception:
        pass
    if kind in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        try:
            chart.plots[0].data_labels.font.color.rgb = hexcolor(text_hex)
        except Exception:
            pass
    elif not hide_axes:
        for axis in (chart.category_axis, chart.value_axis):
            try:
                axis.tick_labels.font.color.rgb = hexcolor(text_hex)
                axis.format.line.color.rgb = hexcolor(axis_hex)
            except Exception:
                pass
        try:
            value_axis = chart.value_axis
            value_axis.has_major_gridlines = True
            value_axis.major_gridlines.format.line.color.rgb = hexcolor(axis_hex)
        except Exception:
            pass
    else:
        # Data labels on an axis-less chart carry the numbers the ruler would.
        try:
            chart.plots[0].data_labels.font.color.rgb = hexcolor(text_hex)
        except Exception:
            pass
    if chart.has_title:
        try:
            for para in chart.chart_title.text_frame.paragraphs:
                for run in para.runs:
                    run.font.color.rgb = hexcolor(text_hex)
        except Exception:
            pass
    if chart.has_legend:
        try:
            chart.legend.font.color.rgb = hexcolor(text_hex)
        except Exception:
            pass


def add_chart(slide, el):
    """Native, editable chart part (bar / line / pie / doughnut / area)."""
    left, top, width, height = emu_box(el)
    model = el.get("chart") or {}
    series_list = list(model.get("series") or [])
    if not series_list:
        return add_placeholder(slide, {**el, "text": "chart without series"})
    categories = [str(c) for c in (model.get("categories") or [])]
    longest = max((len(s.get("values") or []) for s in series_list), default=0)
    if longest == 0:
        return add_placeholder(slide, {**el, "text": "chart without values"})
    if len(categories) < longest:
        categories = categories + [str(i + 1) for i in range(len(categories), longest)]
    categories = categories[:longest]

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for series in series_list:
        values = [float(v) for v in (series.get("values") or [])]
        if len(values) < len(categories):
            values = values + [0.0] * (len(categories) - len(values))
        chart_data.add_series(str(series.get("name") or "Series"), values[:len(categories)])

    kind = CHART_TYPES.get(str(model.get("kind") or "column"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    frame = slide.shapes.add_chart(kind, left, top, width, height, chart_data)
    chart = frame.chart
    style = el.get("style") or {}
    try:
        chart.font.size = Pt(float(style.get("size") or 11))
        chart.font.name = FONT
    except Exception:
        pass
    title = str(model.get("title") or "")
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    chart.has_legend = len(series_list) > 1
    if chart.has_legend:
        try:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
        except Exception:
            pass

    plot = chart.plots[0]
    if kind in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        plot.has_data_labels = True
        try:
            plot.data_labels.font.size = Pt(9)
        except Exception:
            pass
    else:
        try:
            plot.gap_width = 60
        except Exception:
            pass
        if model.get("dataLabels") is True:
            try:
                plot.has_data_labels = True
                plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass

    # axis:"none" = bare bars with value labels, the design-sheet look.
    hide_axes = str(model.get("axis") or "auto") == "none"
    if hide_axes:
        _hide_axes(chart)

    palette = [THEME_ACCENT, "0EA5E9", "10B981", "F59E0B", "EF4444", "8B5CF6"]
    gradient = model.get("gradient")
    if not (isinstance(gradient, list) and len(gradient) >= 2):
        gradient = None
    corner = model.get("corner")
    if not isinstance(corner, (int, float)):
        corner = None
    for index, series_model in enumerate(series_list):
        if index >= len(chart.series):
            break
        series = chart.series[index]
        color = series_model.get("color")
        try:
            if kind in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
                for i, point in enumerate(series.points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = hexcolor(color or palette[i % len(palette)])
            elif kind == XL_CHART_TYPE.LINE_MARKERS:
                series.format.line.color.rgb = hexcolor(color or THEME_ACCENT)
            elif gradient:
                fill = series.format.fill
                fill.gradient()
                fill.gradient_angle = 90.0  # top -> bottom, like the CSS sheet
                stops = fill.gradient_stops
                stops[0].color.rgb = hexcolor(gradient[0])
                stops[0].position = 0.0
                stops[-1].color.rgb = hexcolor(gradient[1])
                stops[-1].position = 1.0
            else:
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = hexcolor(color or THEME_ACCENT)
        except Exception:
            pass
        if corner is not None and kind not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT, XL_CHART_TYPE.LINE_MARKERS):
            _rounded_bars(series, corner)
    # Pin the value axis to the canvas scale (client.js#axisScale) so the
    # exported bars stand at exactly the heights the preview shows.
    if kind not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
        numbers = [float(v) for one in series_list for v in (one.get("values") or [])]
        bottom, top, step, _ticks = axis_scale(numbers)
        try:
            chart.value_axis.minimum_scale = bottom
            chart.value_axis.maximum_scale = top
            chart.value_axis.major_unit = step
        except Exception:
            pass
    theme_chart(chart, slide, kind, hide_axes)
    return frame


def add_placeholder(slide, el):
    left, top, width, height = emu_box(el)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.06
    shape.fill.solid()
    shape.fill.fore_color.rgb = hexcolor("F1F5F9")
    shape.line.color.rgb = hexcolor("CBD5E1")
    shape.line.width = Pt(1)
    no_default_shadow(shape)
    apply_shadow(shape, el.get("shadow"))
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "[%s] %s" % (el.get("type") or "object", el.get("text") or "")
    set_font(run, FONT, 11, False, THEME_MUTED)
    return shape


BUILDERS = {
    "text": add_text,
    "rect": add_rect,
    "image": add_image,
    "line": add_line,
    "table": add_table,
    "chart": add_chart,
}


def direction_code(direction):
    return {"left": "l", "right": "r", "up": "u", "down": "d"}.get(str(direction), "l")


def transition_xml(transition):
    """Build one <p:transition> element for a slide."""
    kind = str(transition.get("type") or "fade")
    speed = {"slow": "slow", "medium": "med", "fast": "fast"}.get(str(transition.get("speed")), "med")
    attrs = ['spd="%s"' % speed]
    if transition.get("advanceOnClick") is False:
        attrs.append('advClick="0"')
    else:
        attrs.append('advClick="1"')
    advance = int(transition.get("advanceAfterMs") or 0)
    if advance > 0:
        attrs.append('advTm="%d"' % advance)
    direction = direction_code(transition.get("direction"))
    if kind == "push":
        body = '<p:push dir="%s"/>' % direction
    elif kind == "wipe":
        body = '<p:wipe dir="%s"/>' % direction
    elif kind == "cover":
        body = '<p:cover dir="%s"/>' % direction
    elif kind == "pull":
        body = '<p:pull dir="%s"/>' % direction
    elif kind == "zoom":
        body = '<p:zoom dir="%s"/>' % ("in" if str(transition.get("direction")) != "right" else "out")
    elif kind == "split":
        orient = "horz" if str(transition.get("direction")) in ("left", "right") else "vert"
        body = '<p:split orient="%s" dir="out"/>' % orient
    else:
        body = TRANSITION_TYPES.get(kind, "<p:fade/>")
    return '<p:transition %s %s>%s</p:transition>' % (nsdecls("p"), " ".join(attrs), body)


def animation_xml(slide_model, shape_ids):
    """Build the <p:timing> entrance-animation tree for one slide."""
    elements = slide_model.get("elements") or []
    entries = []
    for index, el in enumerate(elements):
        anim = el.get("anim")
        if not isinstance(anim, dict):
            continue
        effect = str(anim.get("effect") or "none")
        if effect == "none" or effect not in ANIM_PRESETS:
            continue
        if index >= len(shape_ids) or shape_ids[index] is None:
            continue
        entries.append((index, el, anim, shape_ids[index]))
    if not entries:
        return ""
    entries.sort(key=lambda item: (int(item[2].get("order") or 0), item[0]))

    next_id = 3
    groups = []
    for _index, _el, anim, shape_id in entries:
        effect = str(anim.get("effect"))
        preset_id, preset_subtype = ANIM_PRESETS[effect]
        duration = int(anim.get("duration") or 500)
        delay = int(anim.get("delay") or 0)
        group_id, inner_id, effect_id, behavior_id = next_id, next_id + 1, next_id + 2, next_id + 3
        next_id += 4
        if effect == "appear":
            behavior = (
                '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold"/>'
                '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl></p:cBhvr>'
                '<p:to><p:strVal val="visible"/></p:to></p:set>'
                % (behavior_id, shape_id)
            )
        elif effect == "wipe":
            behavior = (
                '<p:animEffect transition="in" filter="%s"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '</p:cBhvr></p:animEffect>'
                % (WIPE_FILTERS.get(str(anim.get("direction")), "wipe(left)"), behavior_id, duration, shape_id)
            )
        elif effect == "zoom":
            behavior = (
                '<p:animEffect transition="in" filter="zoom(in)"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '</p:cBhvr></p:animEffect>'
                % (behavior_id, duration, shape_id)
            )
        else:
            behavior = (
                '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
                '<p:cTn id="%d" dur="%d"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
                '</p:cBhvr></p:animEffect>'
                % (behavior_id, duration, shape_id)
            )
        groups.append(
            '<p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            '<p:childTnLst><p:par><p:cTn id="%d" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            '<p:childTnLst><p:par><p:cTn id="%d" presetID="%d" presetClass="entr" presetSubtype="%d" '
            'fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="%d"/></p:stCondLst>'
            '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
            '</p:childTnLst></p:cTn></p:par>'
            % (group_id, inner_id, effect_id, preset_id, preset_subtype, delay, behavior)
        )

    return (
        '<p:timing %s><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        % (nsdecls("p"), "".join(groups))
    )


def apply_background(slide, background):
    if not background or not background.get("color"):
        return
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = hexcolor(background["color"], "FFFFFF")
    except Exception as exc:
        sys.stderr.write("background failed: %s\n" % exc)


def build(deck_path, out_path, template=None, effects=True):
    global FONT, THEME_INK, THEME_MUTED, THEME_ACCENT, THEME_SURFACE, THEME_BACKGROUND
    # Anything that fails to render is a data problem the caller must see: the
    # old code only wrote it to stderr, where the tool receipt never looked.
    failures = []
    with open(deck_path, encoding="utf-8") as fh:
        deck = json.load(fh)

    theme = deck.get("theme") or {}
    colors = theme.get("colors") or {}
    fonts = theme.get("fonts") or {}
    FONT = str(fonts.get("latin") or FONT)
    THEME_INK = str(colors.get("ink") or THEME_INK).lstrip("#")
    THEME_MUTED = str(colors.get("muted") or THEME_MUTED).lstrip("#")
    THEME_ACCENT = str(colors.get("accent") or THEME_ACCENT).lstrip("#")
    THEME_SURFACE = str(colors.get("surface") or THEME_SURFACE).lstrip("#")
    THEME_BACKGROUND = str(colors.get("background") or THEME_BACKGROUND).lstrip("#")

    if template:
        prs = Presentation(template)
        remove_all_slides(prs)
    else:
        prs = Presentation()

    size = deck.get("size") or {}
    if size.get("w"):
        prs.slide_width = Inches(float(size["w"]))
    if size.get("h"):
        prs.slide_height = Inches(float(size["h"]))

    base_dir = os.path.dirname(os.path.abspath(deck_path))
    layouts = list(prs.slide_layouts)
    blank = layouts[6] if len(layouts) > 6 else layouts[0]

    for slide_model in deck.get("slides") or []:
        layout_index = slide_model.get("layoutIndex")
        layout = layouts[int(layout_index)] if layout_index is not None and 0 <= int(layout_index) < len(layouts) else blank
        slide = prs.slides.add_slide(layout)
        # Drop inherited placeholders so the canvas and the file agree on geometry.
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)

        apply_background(slide, slide_model.get("background"))
        if slide_model.get("hidden") is True:
            slide._element.set("show", "0")

        shape_ids = []
        for el in slide_model.get("elements") or []:
            kind = str(el.get("type") or "text").lower()
            builder = BUILDERS.get(kind)
            try:
                if builder is None:
                    shape = add_placeholder(slide, el)
                elif kind == "image":
                    shape = add_image(slide, el, base_dir)
                else:
                    shape = builder(slide, el)
            except Exception as exc:  # one bad element must not kill the deck
                line = "element %s failed: %s" % (el.get("id"), exc)
                sys.stderr.write(line + "\n")
                failures.append(line)
                shape = add_placeholder(slide, {**el, "text": "render error"})
            shape_ids.append(getattr(shape, "shape_id", None))

        notes = slide_model.get("notes")
        if isinstance(notes, str) and notes.strip() != "":
            try:
                slide.notes_slide.notes_text_frame.text = notes
            except Exception as exc:
                line = "notes failed: %s" % exc
                sys.stderr.write(line + "\n")
                failures.append(line)

        if effects:
            transition = slide_model.get("transition")
            if isinstance(transition, dict) and str(transition.get("type") or "none") != "none":
                try:
                    slide._element.append(parse_xml(transition_xml(transition)))
                except Exception as exc:
                    line = "transition failed: %s" % exc
                    sys.stderr.write(line + "\n")
                    failures.append(line)
            timing = animation_xml(slide_model, shape_ids)
            if timing:
                try:
                    slide._element.append(parse_xml(timing))
                except Exception as exc:
                    line = "timing failed: %s" % exc
                    sys.stderr.write(line + "\n")
                    failures.append(line)

    title = str(deck.get("title") or "deck")
    prs.core_properties.title = title
    prs.core_properties.comments = "dsh-pptx-studio deck %s" % deck.get("id", "")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    prs.save(out_path)

    # Reopen: cheap structural self-check so the tool can report real numbers.
    check = Presentation(out_path)
    transitions = 0
    timings = 0
    animations = 0
    notes_count = 0
    hidden = 0
    charts = 0
    for slide in check.slides:
        if slide._element.find(qn("p:transition")) is not None:
            transitions += 1
        if slide._element.find(qn("p:timing")) is not None:
            timings += 1
        for tag in ("p:animEffect", "p:set", "p:animMotion", "p:animScale"):
            animations += sum(1 for _ in slide._element.iter(qn(tag)))
        if slide.has_notes_slide:
            try:
                if slide.notes_slide.notes_text_frame.text.strip() != "":
                    notes_count += 1
            except Exception:
                pass
        if slide._element.get("show") == "0":
            hidden += 1
        charts += sum(1 for shape in slide.shapes if getattr(shape, "has_chart", False))
    result = {
        "ok": True,
        "path": os.path.abspath(out_path),
        "bytes": os.path.getsize(out_path),
        "slides": len(check.slides),
        "shapes": sum(len(s.shapes) for s in check.slides),
        "width": check.slide_width / 914400,
        "height": check.slide_height / 914400,
        "features": {
            "transitions": transitions,
            "timings": timings,
            "animations": animations,
            "notes": notes_count,
            "hidden": hidden,
            "charts": charts,
        },
    }
    if failures:
        result["warnings"] = failures
    print(json.dumps(result, ensure_ascii=False))
    return result


def main(argv):
    args = [a for a in argv[1:]]
    template = None
    effects = True
    if "--template" in args:
        i = args.index("--template")
        template = args[i + 1]
        del args[i:i + 2]
    if "--no-effects" in args:
        args.remove("--no-effects")
        effects = False
    if len(args) < 2:
        sys.stderr.write("usage: pptx_build.py <deck.json> <out.pptx> [--template t.pptx] [--no-effects]\n")
        return 2
    try:
        build(args[0], args[1], template, effects)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
