# -*- coding: utf-8 -*-
"""Render a deck.json model into a real .pptx with python-pptx.

Usage:
  python pptx_build.py <deck.json> <out.pptx> [--template <template.pptx>]

The deck model is the contract shared with the DSH plugin host and its client
canvas; see REQUIREMENTS.md for the schema. Everything here is deliberately
small and dependency-free beyond python-pptx, because this script runs inside
the user's DSH process on Windows.
"""
import json
import os
import sys

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

FONT = "微软雅黑"


def hexcolor(value, fallback="0F172A"):
    if not value:
        value = fallback
    value = str(value).strip().lstrip("#")
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    try:
        return RGBColor.from_string(value.upper())
    except Exception:
        return RGBColor.from_string(fallback)


def set_font(run, name, size=None, bold=None, color=None):
    f = run.font
    if size is not None:
        f.size = Pt(float(size))
    if bold is not None:
        f.bold = bool(bold)
    if color is not None:
        f.color.rgb = hexcolor(color)
    f.name = name
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        latin.addnext(ea)
    ea.set("typeface", name)
    cs = rPr.find(qn("a:cs"))
    if cs is None:
        cs = rPr.makeelement(qn("a:cs"), {})
        ea.addnext(cs)
    cs.set("typeface", name)


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
        if para.get("spaceAfter") is not None:
            p.space_after = Pt(float(para["spaceAfter"]))
        runs = para.get("runs")
        if runs:
            for r in runs:
                run = p.add_run()
                run.text = str(r.get("text") or "")
                set_font(
                    run,
                    str(r.get("font") or style.get("font") or FONT),
                    r.get("size", para.get("size", style.get("size", 18))),
                    r.get("bold", para.get("bold", style.get("bold", False))),
                    r.get("color", para.get("color", style.get("color"))),
                )
        else:
            run = p.add_run()
            run.text = str(para.get("text") or "")
            set_font(
                run,
                str(para.get("font") or style.get("font") or FONT),
                para.get("size", style.get("size", 18)),
                para.get("bold", style.get("bold", False)),
                para.get("color", style.get("color")),
            )
    return box


def add_rect(slide, el):
    left, top, width, height = emu_box(el)
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if el.get("radius") else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if el.get("radius"):
        shape.adjustments[0] = float(el["radius"])
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
    shape.shadow.inherit = False
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
        set_font(run, str(style.get("font") or FONT), style.get("size", 14), style.get("bold", False), style.get("color"))
    return shape


def add_image(slide, el, base_dir):
    left, top, width, height = emu_box(el)
    src = str(el.get("src") or "")
    path = src if os.path.isabs(src) else os.path.join(base_dir, src)
    if not path or not os.path.exists(path):
        return add_rect(slide, {**el, "fill": el.get("fill") or "EEF2F7", "line": "CBD5E1",
                                "text": el.get("text") or "[image]"})
    return slide.shapes.add_picture(path, left, top, width, height)


def add_line(slide, el):
    left, top, width, height = emu_box(el)
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, max(height, Inches(0.01)))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hexcolor(el.get("line") or el.get("fill") or "E2E8F0")
    shape.line.fill.background()
    shape.shadow.inherit = False
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
            p = cell.text_frame.paragraphs[0]
            run = p.add_run()
            run.text = str(row[c]) if c < len(row) else ""
            set_font(run, str(style.get("font") or FONT), size, r == 0, style.get("color"))
    return shape


def add_placeholder(slide, el):
    left, top, width, height = emu_box(el)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.06
    shape.fill.solid()
    shape.fill.fore_color.rgb = hexcolor("F1F5F9")
    shape.line.color.rgb = hexcolor("CBD5E1")
    shape.line.width = Pt(1)
    shape.shadow.inherit = False
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = "[%s] %s" % (el.get("type") or "object", el.get("text") or "")
    set_font(run, FONT, 11, False, "64748B")
    return shape


BUILDERS = {
    "text": add_text,
    "rect": add_rect,
    "image": add_image,
    "line": add_line,
    "table": add_table,
}


def build(deck_path, out_path, template=None):
    with open(deck_path, encoding="utf-8") as fh:
        deck = json.load(fh)

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

    for index, slide_model in enumerate(deck.get("slides") or []):
        layout_index = slide_model.get("layoutIndex")
        layout = layouts[int(layout_index)] if layout_index is not None and 0 <= int(layout_index) < len(layouts) else blank
        slide = prs.slides.add_slide(layout)
        # Drop inherited placeholders so the canvas and the file agree on geometry.
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)
        for el in slide_model.get("elements") or []:
            kind = str(el.get("type") or "text").lower()
            builder = BUILDERS.get(kind)
            try:
                if builder is None:
                    add_placeholder(slide, el)
                elif kind == "image":
                    add_image(slide, el, base_dir)
                else:
                    builder(slide, el)
            except Exception as exc:  # one bad element must not kill the deck
                sys.stderr.write("element %s failed: %s\n" % (el.get("id"), exc))
                add_placeholder(slide, {**el, "text": "render error"})

    title = str(deck.get("title") or "deck")
    prs.core_properties.title = title
    prs.core_properties.comments = "dsh-pptx-studio deck %s" % deck.get("id", "")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    prs.save(out_path)

    # Reopen: cheap structural self-check so the tool can report real numbers.
    check = Presentation(out_path)
    result = {
        "ok": True,
        "path": os.path.abspath(out_path),
        "bytes": os.path.getsize(out_path),
        "slides": len(check.slides),
        "shapes": sum(len(s.shapes) for s in check.slides),
        "width": check.slide_width / 914400,
        "height": check.slide_height / 914400,
    }
    print(json.dumps(result, ensure_ascii=False))
    return result


def main(argv):
    args = [a for a in argv[1:]]
    template = None
    if "--template" in args:
        i = args.index("--template")
        template = args[i + 1]
        del args[i:i + 2]
    if len(args) < 2:
        sys.stderr.write("usage: pptx_build.py <deck.json> <out.pptx> [--template t.pptx]\n")
        return 2
    try:
        build(args[0], args[1], template)
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
