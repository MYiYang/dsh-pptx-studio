# -*- coding: utf-8 -*-
"""Inspect a .pptx template (or any deck) and print a JSON summary.

Usage:
  python pptx_inspect.py <file.pptx> [--slides]

Output contract (consumed by the DSH plugin host):
  { ok, path, size:{w,h}, layouts:[{index,name,placeholders:[...]}],
    theme:{colors:{...}, fonts:{major:{latin,ea}, minor:{latin,ea}}},
    slides?:[{index,elements:[...]}] }
"""
import json
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
EMU = 914400.0


def q(tag):
    return "{%s}%s" % (A, tag)


def theme_summary(prs):
    out = {"colors": {}, "fonts": {"major": {}, "minor": {}}}
    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
    except Exception as exc:  # no theme / unexpected package
        out["error"] = str(exc)
        return out
    scheme = root.find(".//%s" % q("clrScheme"))
    if scheme is not None:
        for child in scheme:
            name = etree.QName(child).localname
            srgb = child.find("%s" % q("srgbClr"))
            sys_clr = child.find("%s" % q("sysClr"))
            if srgb is not None:
                out["colors"][name] = "#" + srgb.get("val", "")
            elif sys_clr is not None:
                out["colors"][name] = "#" + sys_clr.get("lastClr", "000000")
    for which in ("major", "minor"):
        node = root.find(".//%s" % q(which + "Font"))
        if node is None:
            continue
        for tag in ("latin", "ea"):
            el = node.find("%s" % q(tag))
            if el is not None:
                out["fonts"][which][tag] = el.get("typeface", "")
    return out


def placeholder_summary(shape):
    ph = shape.placeholder_format
    return {
        "idx": ph.idx,
        "type": str(ph.type).replace("PLACEHOLDER_TYPE.", "").lower() if ph.type is not None else "none",
        "name": shape.name,
        "box": [round(shape.left / EMU, 3), round(shape.top / EMU, 3),
                round(shape.width / EMU, 3), round(shape.height / EMU, 3)],
    }


def text_of(shape):
    try:
        if shape.has_text_frame:
            return shape.text_frame.text
    except Exception:
        pass
    return ""


def shapes_summary(shapes):
    items = []
    for shape in shapes:
        item = {
            "name": shape.name,
            "type": str(shape.shape_type).split(".")[-1].lower() if shape.shape_type is not None else "unknown",
            "box": [round(shape.left / EMU, 3), round(shape.top / EMU, 3),
                    round(shape.width / EMU, 3), round(shape.height / EMU, 3)],
            "text": text_of(shape),
        }
        if shape.shape_type is not None and "PLACEHOLDER" in str(shape.shape_type):
            try:
                item["placeholder"] = placeholder_summary(shape)
            except Exception:
                pass
        items.append(item)
    return items


def inspect(path, with_slides=False):
    prs = Presentation(path)
    result = {
        "ok": True,
        "path": path,
        "size": {"w": round(prs.slide_width / EMU, 3), "h": round(prs.slide_height / EMU, 3)},
        "layouts": [],
        "theme": theme_summary(prs),
        "slideCount": len(prs.slides),
    }
    for i, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for shape in layout.placeholders:
            try:
                placeholders.append(placeholder_summary(shape))
            except Exception:
                pass
        result["layouts"].append({"index": i, "name": layout.name, "placeholders": placeholders})
    if with_slides:
        result["slides"] = [{"index": i, "elements": shapes_summary(s.shapes)} for i, s in enumerate(prs.slides)]
    return result


def main(argv):
    args = argv[1:]
    with_slides = "--slides" in args
    if with_slides:
        args.remove("--slides")
    if not args:
        print(json.dumps({"ok": False, "error": "usage: pptx_inspect.py <file.pptx> [--slides]"}))
        return 2
    try:
        print(json.dumps(inspect(args[0], with_slides), ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
