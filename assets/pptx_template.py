# -*- coding: utf-8 -*-
"""Decompose a .pptx template into a reusable design spec (the "template model").

This is what turns "a deck the user likes" into something a new deck can be
built out of. It answers four questions per template slide:

  1. what is this page for?          -> role (cover / toc / section / content / ending)
  2. what does its background look   -> plate render (PowerPoint COM, text hidden)
     like without any copy on it?
  3. where does copy go, in what     -> slots: box, font, size, colour, align,
     style, and how much fits?          anchor, maxChars
  4. what is the visual language?    -> design.palette / fonts / text colours

Usage:
  python pptx_template.py <template.pptx> <outDir> [--max-slides N] [--no-render]
                          [--width 1600] [--height 900]

Output: one JSON object on stdout (contract consumed by the DSH plugin host).
Images land in <outDir>/{reference,plates}/s<N>.png (only when PowerPoint was
found; otherwise the spec still carries slots and palette).
"""
import base64
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU = 914400.0

NS = {"a": A, "p": P}

# Placeholder types that exist on every slide and should be filled by the
# template/PowerPoint itself, not by the author.
AUTO_PLACEHOLDERS = {"date", "footer", "slide_number", "slide number"}

# Role heuristics: this is a labelling job, not a classification problem. The
# point is to tell the model "this page is a section divider", never to be
# exact about the designer's intent.
TOC_RE = re.compile(r"(目\s*录|contents?|mulu|index|目录页)", re.I)
ENDING_RE = re.compile(r"(谢\s*谢|感\s*谢|thank|q\s*&\s*a|结束|汇报完毕)", re.I)
SECTION_NUM_RE = re.compile(r"^[壹贰叁肆伍陆柒捌玖拾0-9一二三四五六七八九十]{1,3}$")


def q(tag):
    return "{%s}%s" % (A, tag)


def pq(tag):
    return "{%s}%s" % (P, tag)


def emu_in(value):
    """EMU -> inches, rounded to 3 decimals (the deck model's unit)."""
    return round(float(value) / EMU, 3)


def hex_or_none(node):
    """First colour under a fill/line node as #RRGGBB, or None."""
    if node is None:
        return None
    srgb = node.find(".//" + q("srgbClr"))
    if srgb is not None and srgb.get("val"):
        return "#" + srgb.get("val")
    scheme = node.find(".//" + q("schemeClr"))
    if scheme is not None and scheme.get("val"):
        return "scheme:" + scheme.get("val")
    return None


def theme_colors(prs):
    """All named colours from the presentation theme (for scheme: resolution)."""
    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
    except Exception:
        return {}
    scheme = root.find(".//" + q("clrScheme"))
    out = {}
    if scheme is None:
        return out
    for child in scheme:
        name = etree.QName(child).localname
        srgb = child.find(q("srgbClr"))
        sys_clr = child.find(q("sysClr"))
        if srgb is not None:
            out[name] = "#" + srgb.get("val", "000000")
        elif sys_clr is not None:
            out[name] = "#" + sys_clr.get("lastClr", "000000")
    return out


def resolve_color(value, scheme):
    if not value:
        return None
    if value.startswith("scheme:"):
        return scheme.get(value.split(":", 1)[1])
    return value


def theme_fonts(prs):
    try:
        part = prs.slide_masters[0].part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
    except Exception:
        return {"major": {}, "minor": {}}
    out = {"major": {}, "minor": {}}
    for which in ("major", "minor"):
        node = root.find(".//" + q(which + "Font"))
        if node is None:
            continue
        for tag in ("latin", "ea", "cs"):
            el = node.find(q(tag))
            if el is not None and el.get("typeface"):
                out[which][tag] = el.get("typeface")
    return out


# ── transforms ────────────────────────────────────────────────────────────────
# Group shapes position their children in their own coordinate space, so a text
# box inside a group needs the group's off/ext/chOff/chExt composed in. Ignoring
# this is the classic way slot boxes come out scrambled.

def xfrm_of(node):
    """(off.x, off.y, ext.cx, ext.cy) of a shape node, or None."""
    xfrm = node.find("./p:spPr/a:xfrm", NS)
    if xfrm is None:
        xfrm = node.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    if off is None or ext is None:
        return None
    return (int(off.get("x", 0)), int(off.get("y", 0)), int(ext.get("cx", 0)), int(ext.get("cy", 0)))


def group_child_transform(node, parent):
    """Child-space -> absolute transform for a p:grpSp."""
    xfrm = node.find("./p:grpSpPr/a:xfrm", NS)
    if xfrm is None:
        return parent
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    ch_off = xfrm.find("./a:chOff", NS)
    ch_ext = xfrm.find("./a:chExt", NS)
    if off is None or ext is None or ch_off is None or ch_ext is None:
        return parent
    ox, oy = int(off.get("x", 0)), int(off.get("y", 0))
    cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
    chx, chy = int(ch_off.get("x", 0)), int(ch_off.get("y", 0))
    chcx, chcy = int(ch_ext.get("cx", 0)), int(ch_ext.get("cy", 0))
    sx = (cx / float(chcx)) if chcx else 1.0
    sy = (cy / float(chcy)) if chcy else 1.0
    psx, ptx, psy, pty = parent
    return (sx * psx, (ox - chx * sx) * psx + ptx, sy * psy, (oy - chy * sy) * psy + pty)


IDENTITY = (1.0, 0.0, 1.0, 0.0)


def apply_transform(box, transform):
    sx, tx, sy, ty = transform
    return [box[0] * sx + tx, box[1] * sy + ty, box[2] * sx, box[3] * sy]


# ── text extraction ───────────────────────────────────────────────────────────

def run_style(run_node, inherited):
    """Font size/bold/colour/font of one a:r, falling back to inherited values."""
    style = dict(inherited)
    rpr = run_node.find("./a:rPr", NS)
    if rpr is None:
        return style
    if rpr.get("sz"):
        style["size"] = round(int(rpr.get("sz")) / 100.0, 2)
    if rpr.get("b") is not None:
        style["bold"] = rpr.get("b") in ("1", "true")
    if rpr.get("i") is not None:
        style["italic"] = rpr.get("i") in ("1", "true")
    if rpr.get("u") is not None and rpr.get("u") != "none":
        style["underline"] = True
    fill = rpr.find("./a:solidFill", NS)
    color = hex_or_none(fill) if fill is not None else None
    if color:
        style["color"] = color
    latin = rpr.find("./a:latin", NS)
    if latin is not None and latin.get("typeface"):
        style["font"] = latin.get("typeface")
    ea = rpr.find("./a:ea", NS)
    if ea is not None and ea.get("typeface") and "font" not in style:
        style["font"] = ea.get("typeface")
    return style


def def_rpr_style(shape_node, inherited):
    """Paragraph defaults from a:pPr/a:defRPr (templates often style here)."""
    style = dict(inherited)
    node = shape_node.find(".//a:pPr/a:defRPr", NS)
    if node is None:
        return style
    if node.get("sz"):
        style["size"] = round(int(node.get("sz")) / 100.0, 2)
    if node.get("b") is not None:
        style["bold"] = node.get("b") in ("1", "true")
    fill = node.find("./a:solidFill", NS)
    color = hex_or_none(fill) if fill is not None else None
    if color:
        style["color"] = color
    latin = node.find("./a:latin", NS)
    if latin is not None and latin.get("typeface"):
        style["font"] = latin.get("typeface")
    return style


def text_slots_of(shape_node, base_style):
    """Text lines of one p:sp, each with its own style, plus aggregated style."""
    paragraphs = []
    for p_node in shape_node.findall("./p:txBody/a:p", NS):
        style = def_rpr_style(shape_node, base_style)
        ppr = p_node.find("./a:pPr", NS)
        if ppr is not None:
            if ppr.get("algn"):
                style["align"] = ppr.get("algn")
            if ppr.get("marL") is not None:
                style["marL"] = int(ppr.get("marL"))
            ln = ppr.find("./a:lnSpc/a:spcPct", NS)
            if ln is not None and ln.get("val"):
                style["lineSpacing"] = round(int(ln.get("val")) / 100000.0, 3)
        runs = []
        for child in p_node:
            tag = etree.QName(child).localname
            if tag == "r":
                t = child.find("./a:t", NS)
                text = t.text if t is not None and t.text else ""
                runs.append({"text": text, "style": run_style(child, style)})
            elif tag == "br":
                runs.append({"text": "\n", "style": dict(style)})
            elif tag == "fld":
                t = child.find("./a:t", NS)
                runs.append({"text": t.text if t is not None and t.text else "", "style": dict(style)})
        text = "".join(r["text"] for r in runs)
        if text.strip() or runs:
            paragraphs.append({"text": text, "style": style, "runs": runs})
    body_pr = shape_node.find("./p:txBody/a:bodyPr", NS)
    anchor = "top"
    if body_pr is not None:
        anchor = body_pr.get("anchor") or "top"
        anchor = {"ctr": "middle", "t": "top", "b": "bottom"}.get(anchor, anchor)
    return paragraphs, anchor


def slug(value, fallback):
    value = re.sub(r"[^a-z0-9]+", "-", (value or "").lower()).strip("-")
    return value or fallback


# ── rendering (PowerPoint COM) ────────────────────────────────────────────────

def find_powershell():
    for name in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        r"C:\Program Files\PowerShell\7\pwsh.exe",
    ):
        if os.path.exists(candidate):
            return candidate
    return None


def render_slides(pptx_path, out_dir, width, height, max_slides, script_path):
    """Render reference + plate PNGs through PowerPoint. Returns a dict."""
    powershell = find_powershell()
    if not powershell:
        return {"ok": False, "engine": "powerpoint", "error": "PowerShell not found"}
    if not os.path.exists(script_path):
        return {"ok": False, "engine": "powerpoint", "error": "pptx_render.ps1 missing"}
    args = [
        powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path,
        "-Src", os.path.abspath(pptx_path),
        "-OutDir", os.path.abspath(out_dir),
        "-Width", str(width), "-Height", str(height),
        "-MaxSlides", str(max_slides or 0),
    ]
    try:
        proc = subprocess.run(
            args,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=300,
            # PowerShell writes UTF-16-ish console output on some hosts; we only
            # parse the last JSON line, so decode defensively.
            errors="replace",
        )
    except subprocess.TimeoutExpired:
        return {"ok": False, "engine": "powerpoint", "error": "render timed out after 300s"}
    except Exception as exc:  # noqa: BLE001 - any spawn failure means "no COM"
        return {"ok": False, "engine": "powerpoint", "error": str(exc)}
    text = proc.stdout.decode("utf-8", "replace") if isinstance(proc.stdout, bytes) else str(proc.stdout)
    match = re.search(r"\{.*\}\s*$", text.strip(), re.S)
    if not match:
        detail = text.strip()[-400:] or (proc.stderr.decode("utf-8", "replace")[-400:] if isinstance(proc.stderr, bytes) else "")
        return {"ok": False, "engine": "powerpoint", "error": "no JSON from renderer", "detail": detail}
    try:
        parsed = json.loads(match.group(0))
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "engine": "powerpoint", "error": "bad renderer JSON: %s" % exc}
    if not parsed.get("ok"):
        return {"ok": False, "engine": "powerpoint", "error": parsed.get("error", "render failed")}
    return parsed


# ── fallback rendering (no PowerPoint): compose pictures + solid shapes ───────

def pillow_render(prs, out_dir, width=1280, height=720):
    """Best-effort plate without PowerPoint: paste pictures and solid shapes.

    Groups are flattened recursively; gradients/text are skipped. It is a rough
    plate, which is why the spec reports engine='pillow' so the caller can warn.
    """
    try:
        from PIL import Image, ImageDraw
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "engine": "pillow", "error": "Pillow unavailable: %s" % exc}
    os.makedirs(os.path.join(out_dir, "plates"), exist_ok=True)
    digits = 4  # EMU rounding precision for raster pixels

    def scale(box):
        sx = width / float(prs.slide_width)
        sy = height / float(prs.slide_height)
        return (box[0] * sx, box[1] * sy, box[2] * sx, box[3] * sy)

    # The generator form above is awkward with relationship lookups; do it with
    # an explicit walk instead.
    parts = {}

    def walk(node, transform, slide_part):
        tm = group_child_transform(node, transform) if node.tag == pq("grpSp") else transform
        for child in node:
            tag = etree.QName(child).localname
            if tag == "grpSp":
                walk(child, tm, slide_part)
            elif tag == "pic":
                blip = child.find(".//a:blip", NS)
                rid = None
                if blip is not None:
                    rid = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if not rid:
                    continue
                try:
                    blob = slide_part.related_part(rid).blob
                except Exception:
                    continue
                parts[id(child)] = blob
                yield ("pic", child, tm, blob)
            elif tag == "sp":
                fill = child.find("./p:spPr/a:solidFill", NS)
                geom = child.find("./p:spPr/a:prstGeom", NS)
                yield ("sp", child, tm, (hex_or_none(fill), geom.get("prst") if geom is not None else None))

    pages = []
    for index, slide in enumerate(prs.slides):
        canvas = Image.new("RGB", (width, height), (255, 255, 255))
        draw = ImageDraw.Draw(canvas, "RGBA")
        for kind, node, tm, payload in walk(slide.shapes._spTree, IDENTITY, slide.part):
            box = xfrm_of(node)
            if box is None:
                continue
            abs_box = apply_transform([box[0], box[1], box[2], box[3]], tm)
            rect = scale(abs_box)
            if kind == "pic":
                try:
                    import io
                    tile = Image.open(io.BytesIO(payload)).convert("RGBA")
                except Exception:
                    continue
                tile = tile.resize((max(1, int(rect[2])), max(1, int(rect[3]))), Image.LANCZOS)
                canvas.paste(tile, (int(rect[0]), int(rect[1])), tile)
            else:
                color, geom = payload
                rgb = resolve_color(color, {})
                if not rgb or not rgb.startswith("#"):
                    continue
                value = tuple(int(rgb[i:i + 2], 16) for i in (1, 3, 5))
                if geom in ("ellipse",):
                    draw.ellipse(rect, fill=value)
                elif geom in ("roundRect",):
                    draw.rounded_rectangle(rect, radius=min(rect[2], rect[3]) / 6.0, fill=value)
                else:
                    draw.rectangle(rect, fill=value)
        plate = os.path.join(out_dir, "plates", "s%d.png" % index)
        canvas.save(plate)
        pages.append({"index": index, "plate": plate})
    return {"ok": True, "engine": "pillow", "slides": len(pages), "width": width, "height": height, "pages": pages}


# ── asset optimisation ────────────────────────────────────────────────────────
# PowerPoint hands back ~1MB PNGs per slide. Two consumers want two sizes: the
# exported .pptx embeds the plate as a picture (1600px is plenty at 13.3in), and
# the sandboxed HTML sheet has to inline every image as a data URL (so it needs
# a small JPEG or the document becomes tens of MB). Both are produced here so
# the plugin host never has to re-encode images.

THUMB_WIDTH = 720


def optimize_assets(out_dir, plate_width=1600, quality=88):
    """PNG renders -> plates/sN.jpg (+ thumbs/sN.jpg). Returns what changed."""
    try:
        from PIL import Image
    except Exception:
        return {"ok": False, "error": "Pillow unavailable; keeping PNG renders"}
    thumbs_dir = os.path.join(out_dir, "thumbs")
    os.makedirs(thumbs_dir, exist_ok=True)
    converted = {"plates": 0, "thumbs": 0}
    for kind, target_quality, target_width in (("plates", quality, plate_width), ("reference", 84, plate_width)):
        source_dir = os.path.join(out_dir, kind)
        if not os.path.isdir(source_dir):
            continue
        for name in sorted(os.listdir(source_dir)):
            if not name.lower().endswith(".png"):
                continue
            stem = os.path.splitext(name)[0]
            source = os.path.join(source_dir, name)
            target = os.path.join(source_dir, stem + ".jpg")
            try:
                image = Image.open(source).convert("RGB")
                if target_width and image.width > target_width:
                    ratio = target_width / float(image.width)
                    image = image.resize((target_width, max(1, int(image.height * ratio))), Image.LANCZOS)
                image.save(target, "JPEG", quality=target_quality, optimize=True)
                os.remove(source)
                converted["plates" if kind == "plates" else "reference"] += 1
                if kind == "plates":
                    thumb = image.copy()
                    if thumb.width > THUMB_WIDTH:
                        ratio = THUMB_WIDTH / float(thumb.width)
                        thumb = thumb.resize((THUMB_WIDTH, max(1, int(thumb.height * ratio))), Image.LANCZOS)
                    thumb.save(os.path.join(thumbs_dir, stem + ".jpg"), "JPEG", quality=72, optimize=True)
                    converted["thumbs"] += 1
            except Exception:
                continue
    return {"ok": True, **converted}


def rewrite_asset_paths(slides, out_dir):
    """Point the spec at the .jpg files when the optimisation actually ran."""
    for record in slides:
        for key in ("plate", "reference"):
            value = record.get(key)
            if not value:
                continue
            stem, ext = os.path.splitext(value)
            if ext.lower() != ".png":
                continue
            if os.path.exists(os.path.join(out_dir, stem + ".jpg")):
                record[key] = stem + ".jpg"
            elif not os.path.exists(os.path.join(out_dir, value)):
                record[key] = None


# ── text contrast ─────────────────────────────────────────────────────────────
# A run that inherits its colour from the layout has no fill in the OOXML, so it
# comes back as "no colour". Dropping the deck's dominant ink on it is wrong the
# moment the field sits on a dark plaque (this template has several), so the
# choice is made against the actual plate pixels behind each box.

INK_CANDIDATES = ["#FFFFFF", "#FFF8EC"]


def _relative_luminance(rgb):
    channels = []
    for value in rgb[:3]:
        c = value / 255.0
        channels.append(c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4)
    r, g, b = channels
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast(a, b):
    la, lb = _relative_luminance(a), _relative_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def _hex_rgb(value):
    text = value.lstrip("#")
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


def background_luminance(plate_path, box, slide_w, slide_h):
    """Average colour + luminance of the plate under a field's box."""
    try:
        from PIL import Image
    except Exception:
        return None
    try:
        image = Image.open(plate_path).convert("RGB")
    except Exception:
        return None
    width, height = image.size
    sx = width / float(slide_w)
    sy = height / float(slide_h)
    # Inset a little: box edges often sit on a frame or a drop shadow.
    x0 = max(0, int((box[0] + box[2] * 0.08) * sx))
    y0 = max(0, int((box[1] + box[3] * 0.08) * sy))
    x1 = min(width, int((box[0] + box[2] * 0.92) * sx))
    y1 = min(height, int((box[1] + box[3] * 0.92) * sy))
    if x1 - x0 < 2 or y1 - y0 < 2:
        return None
    patch = image.crop((x0, y0, x1, y1)).resize((1, 1), Image.LANCZOS)
    rgb = patch.getpixel((0, 0))
    return {"rgb": tuple(rgb), "luminance": _relative_luminance(rgb)}


def apply_contrast_colors(slides, out_dir, slide_w, slide_h):
    """Give every colour-less field a colour that reads on its own background."""
    for record in slides:
        plate = record.get("plate")
        plate_path = os.path.join(out_dir, plate) if plate else None
        if not plate_path or not os.path.exists(plate_path):
            continue
        for field in record.get("fields", []):
            if field.get("style", {}).get("color"):
                continue
            sample = background_luminance(plate_path, field["box"], slide_w, slide_h)
            if sample is None:
                continue
            bg = sample["rgb"]
            # Prefer a warm near-white / deep ink over pure black-white pairs so
            # the copy still belongs to the template's palette family.
            best = None
            for candidate in ("#FFFFFF", "#FFF8EC", "#2B1206", "#4A2410"):
                score = _contrast(bg, _hex_rgb(candidate))
                if best is None or score > best[0]:
                    best = (score, candidate)
            field["style"]["color"] = best[1]
            field["style"]["contrastPicked"] = round(best[0], 2)


# ── design tokens ─────────────────────────────────────────────────────────────

def palette_from_image(path, limit=6):
    """Dominant colours + their share, from a rendered plate."""
    try:
        from PIL import Image
    except Exception:
        return {"palette": [], "background": None}
    try:
        image = Image.open(path).convert("RGB")
    except Exception:
        return {"palette": [], "background": None}
    small = image.resize((160, 90))
    quantized = small.quantize(colors=8, method=Image.MEDIANCUT)
    palette_entries = quantized.getpalette() or []
    counts = quantized.getcolors() or []
    total = float(sum(c for c, _ in counts)) or 1.0
    entries = []
    for count, idx in counts:
        rgb = tuple(palette_entries[idx * 3: idx * 3 + 3])
        entries.append({"hex": "#%02X%02X%02X" % rgb, "share": round(count / total, 3)})
    entries.sort(key=lambda item: item["share"], reverse=True)
    # Drop near-duplicates so the palette reads as a design choice, not noise.
    kept = []
    for entry in entries:
        if any(color_distance(entry["hex"], other["hex"]) < 26 for other in kept):
            continue
        kept.append(entry)
        if len(kept) >= limit:
            break
    center = image.getpixel((image.width // 2, int(image.height * 0.86)))
    return {
        "palette": kept,
        "background": "#%02X%02X%02X" % tuple(center[:3]),
    }


def color_distance(a, b):
    ar, ag, ab = int(a[1:3], 16), int(a[3:5], 16), int(a[5:7], 16)
    br, bg, bb = int(b[1:3], 16), int(b[3:5], 16), int(b[5:7], 16)
    return ((ar - br) ** 2 + (ag - bg) ** 2 + (ab - bb) ** 2) ** 0.5


def dhash(path):
    """64-bit difference hash, used to group slides that share a chrome."""
    try:
        from PIL import Image
    except Exception:
        return None
    try:
        image = Image.open(path).convert("L").resize((9, 8))
    except Exception:
        return None
    bits = 0
    for y in range(8):
        for x in range(8):
            bits = (bits << 1) | (1 if image.getpixel((x, y)) < image.getpixel((x + 1, y)) else 0)
    return bits


def hash_distance(a, b):
    if a is None or b is None:
        return 64
    return bin(a ^ b).count("1")


# ── roles ─────────────────────────────────────────────────────────────────────

def classify_role(index, total, slots, paragraph_count, char_count, fields=None):
    texts = " ".join(slot["text"] for slot in slots)
    if index == 0:
        return "cover"
    if TOC_RE.search(texts):
        return "toc"
    if ENDING_RE.search(texts):
        return "ending"
    if index == total - 1 and char_count < 30 and len(slots) <= 2:
        return "ending"
    # Per-character boxes inflate the slot count; judge a divider by its fields.
    visuals = len(fields) if fields is not None else len(slots)
    if visuals <= 1 and char_count <= 14:
        return "section"
    if visuals <= 2 and char_count <= 8:
        return "section"
    return "content"


# ── slots ─────────────────────────────────────────────────────────────────────

def slot_style(paragraphs):
    """Dominant style of a text shape: biggest run wins (that is the title)."""
    best = {"size": 18.0, "bold": False, "italic": False, "underline": False,
            "color": None, "font": None, "align": paragraphs[0]["style"].get("align", "l") if paragraphs else "l"}
    best_size = -1.0
    for para in paragraphs:
        for run in para["runs"]:
            style = run["style"]
            size = float(style.get("size") or 18.0)
            if size > best_size:
                best_size = size
                best = {
                    "size": round(size, 2),
                    "bold": bool(style.get("bold")),
                    "italic": bool(style.get("italic")),
                    "underline": bool(style.get("underline")),
                    "color": style.get("color"),
                    "font": style.get("font"),
                    "align": para["style"].get("align", "l"),
                }
    return best


def union_box(boxes):
    left = min(b[0] for b in boxes)
    top = min(b[1] for b in boxes)
    right = max(b[0] + b[2] for b in boxes)
    bottom = max(b[1] + b[3] for b in boxes)
    return [round(left, 3), round(top, 3), round(right - left, 3), round(bottom - top, 3)]


def same_line(a, b):
    """Two slots sit on one visual line when their vertical centres are close."""
    ca = a["box"][1] + a["box"][3] / 2.0
    cb = b["box"][1] + b["box"][3] / 2.0
    return abs(ca - cb) <= max(0.06, min(a["box"][3], b["box"][3]) * 0.45)


def similar_style(a, b):
    sa, sb = a["style"], b["style"]
    if abs(float(sa.get("size") or 0) - float(sb.get("size") or 0)) > 0.75:
        return False
    if (sa.get("font") or "") != (sb.get("font") or ""):
        return False
    if (sa.get("color") or "") != (sb.get("color") or ""):
        return False
    if bool(sa.get("bold")) != bool(sb.get("bold")):
        return False
    return True


def style_key(slot):
    style = slot["style"]
    return (round(float(style.get("size") or 0), 1), style.get("font") or "", style.get("color") or "",
            bool(style.get("bold")))


def char_width(style):
    return float(style.get("size") or 16) / 72.0


def is_single_char(slot):
    return len(slot["text"].strip()) == 1


def center_of(box):
    return (box[0] + box[2] / 2.0, box[1] + box[3] / 2.0)


def same_band(a, b):
    """Whether two same-styled boxes belong to one visual line/block.

    Staggered per-character titles (this template's cover alternates baselines)
    sit within a few character widths of each other, so the band test is
    generous while the chaining test stays strict about horizontal distance.
    """
    ax, ay = center_of(a["box"])
    bx, by = center_of(b["box"])
    cw = max(char_width(a["style"]), char_width(b["style"]))
    if is_single_char(a) and is_single_char(b):
        return abs(ay - by) <= cw * 4.0
    return abs(ay - by) <= max(0.06, min(a["box"][3], b["box"][3]) * 0.45)


def chained(a, b):
    """Whether b continues a's run (to its right) rather than being a new item."""
    ax, _ = center_of(a["box"])
    bx, _ = center_of(b["box"])
    cw = max(char_width(a["style"]), char_width(b["style"]))
    if is_single_char(a) and is_single_char(b):
        # Per-character boxes overlap their neighbours; spacing is the signal.
        return bx > ax and abs(bx - ax) <= cw * 3.2 and abs(center_of(a["box"])[1] - center_of(b["box"])[1]) <= cw * 4.0
    gap = b["box"][0] - (a["box"][0] + a["box"][2])
    return -cw * 0.6 <= gap <= cw * 1.2


def reading_order(field):
    return (round(field["box"][1], 1), field["box"][0])


def make_field(members, index):
    run = sorted(members, key=lambda slot: slot["box"][0])
    text = "".join(slot["text"] for slot in run)
    per_char = len(run) > 1 and all(is_single_char(slot) for slot in run)
    box = union_box([slot["box"] for slot in run])
    style = dict(run[0]["style"])
    if per_char:
        style["align"] = "ctr"
    return {
        "id": "field%d" % index,
        "slots": [slot["id"] for slot in run],
        "box": box,
        "sample": text,
        "perChar": per_char,
        "style": style,
        "fit": fit_estimate(box, style["size"], style["align"]),
        "fillAs": "line" if per_char else "slots",
    }


def name_fields(fields):
    """Give fields names a human (and a model) can aim copy at.

    Largest type on the page is the title, the next one down is the subtitle,
    everything else is body1..N in reading order. Slot-level ids stay as the
    template had them; field ids are the authoring surface.
    """
    if not fields:
        return fields
    ordered = sorted(fields, key=lambda f: (-float(f["style"].get("size") or 0), f["box"][1], f["box"][0]))
    used = set()
    for rank, field in enumerate(ordered):
        if rank == 0:
            base = "title"
        elif rank == 1 and float(field["style"].get("size") or 0) >= float(ordered[0]["style"].get("size") or 0) * 0.45:
            base = "subtitle"
        else:
            base = "body"
        # The cover's year badge and runner sit above the title; name them for
        # what they are instead of numbering them like body copy.
        sample = field["sample"].strip()
        if re.match(r"^[0-9０-９零一二三四五六七八九十贰叁肆]{2,6}$", sample.replace(" ", "")):
            base = "badge"
        elif re.match(r"^(汇报人|主讲人|报告人|author|presented)", sample, re.I):
            base = "credit"
        elif re.search(r"(目录|contents|mulu)", sample, re.I):
            base = "tocTitle"
        name = base
        n = 1
        while name in used:
            n += 1
            name = "%s%d" % (base, n)
        used.add(name)
        field["id"] = name
    return sorted(fields, key=reading_order)


def build_fields(slots):
    """Merge per-character text boxes into the fields a human would name.

    Designers break a title into one text box per character so they can kern it
    (this template's cover is 敦/煌/魅/力 in four boxes). Filling four boxes
    independently is how a copy change turns into scrambled type, so the spec
    exposes *fields*: a run of same-styled boxes on one line/block, ordered left
    to right, with the union box used when new copy is dropped in.
    """
    candidates = [slot for slot in slots if slot["kind"] not in ("footer", "date", "slideNumber")]
    buckets = {}
    for slot in candidates:
        buckets.setdefault(style_key(slot), []).append(slot)
    raw_fields = []
    for slots_in_bucket in buckets.values():
        ordered = sorted(slots_in_bucket, key=lambda slot: (slot["box"][0], slot["box"][1]))
        bands = []
        for slot in ordered:
            for band in bands:
                if same_band(band[-1], slot):
                    band.append(slot)
                    break
            else:
                bands.append([slot])
        for band in bands:
            band.sort(key=lambda slot: (slot["box"][0], slot["box"][1]))
            run = [band[0]]
            for slot in band[1:]:
                if chained(run[-1], slot):
                    run.append(slot)
                else:
                    raw_fields.append(make_field(run, len(raw_fields) + 1))
                    run = [slot]
            raw_fields.append(make_field(run, len(raw_fields) + 1))
    raw_fields.sort(key=reading_order)
    fields = name_fields(raw_fields)

    # Parallel copy (three identical columns) gets a shared group id so the
    # caller can fill one and mirror it instead of guessing.
    groups = []
    for field in fields:
        for group in groups:
            head = next((f for f in fields if f["id"] == group["fields"][0]), None)
            if head and head["sample"] == field["sample"]                     and abs(head["box"][2] - field["box"][2]) < 0.05                     and abs(head["box"][3] - field["box"][3]) < 0.05:
                group["fields"].append(field["id"])
                field["repeatGroup"] = group["id"]
                break
        else:
            groups.append({"id": "r%d" % (len(groups) + 1), "fields": [field["id"]]})
    for group in groups:
        if len(group["fields"]) < 2:
            for field in fields:
                if field["id"] in group["fields"]:
                    field.pop("repeatGroup", None)
    return fields, [group for group in groups if len(group["fields"]) > 1]


def fit_estimate(box, size_pt, align):
    """How much copy fits in a slot — the number the model needs to not overflow."""
    width_in, height_in = box[2], box[3]
    line_pt = max(1.0, size_pt) * 1.25
    lines = max(1, int(height_in * 72.0 / line_pt))
    cjk_per_line = max(1, int(width_in * 72.0 / (size_pt * 1.02)))
    latin_per_line = max(1, int(width_in * 72.0 / (size_pt * 0.55)))
    return {
        "lines": lines,
        "cjkPerLine": cjk_per_line,
        "latinPerLine": latin_per_line,
        "maxCjk": cjk_per_line * lines,
        "maxLatin": latin_per_line * lines,
    }


def slot_kind(shape_node, placeholder_type, has_text):
    ptype = (placeholder_type or "").lower()
    if "slide number" in ptype or "slide_number" in ptype:
        return "slideNumber"
    if "date" in ptype:
        return "date"
    if "footer" in ptype:
        return "footer"
    if "title" in ptype or "ctrTitle" in (placeholder_type or ""):
        return "title"
    return "text"


def collect_slots(slide, slide_index, layout, master, scheme, theme_font):
    """Every text-bearing shape on a slide, with absolute geometry and style."""
    slots = []
    used_ids = set()

    def inherited_style():
        return {"size": 18.0, "bold": False, "italic": False, "underline": False,
                "color": None, "font": theme_font, "align": "l"}

    def visit(node, transform, layout_ph_by_idx):
        tag = etree.QName(node).localname
        if tag == "grpSp":
            child_tf = group_child_transform(node, transform)
            for child in node:
                visit(child, child_tf, layout_ph_by_idx)
            return
        if tag not in ("sp",):
            return
        # Placeholders may omit geometry on the slide and inherit it.
        ph = node.find("./p:nvSpPr/p:nvPr/p:ph", NS)
        ph_idx = ph.get("idx") if ph is not None else None
        ph_type = (ph.get("type") if ph is not None else None) or "body"
        box = xfrm_of(node)
        if box is None and ph_idx is not None and layout_ph_by_idx.get(ph_idx):
            inherited_box = layout_ph_by_idx[ph_idx]
            box = inherited_box
        if box is None:
            return
        # Compose in EMU space (the transform's offsets are EMU), then convert.
        abs_box = apply_transform([float(box[0]), float(box[1]), float(box[2]), float(box[3])], transform)
        box = [emu_in(v) for v in abs_box]
        paragraphs, anchor = text_slots_of(node, inherited_style())
        text = "\n".join(para["text"] for para in paragraphs).strip()
        if not text:
            return
        style = slot_style(paragraphs)
        style["color"] = resolve_color(style.get("color"), scheme)
        style["anchor"] = anchor
        for run in [r for para in paragraphs for r in para["runs"]]:
            if run["style"].get("color"):
                run["style"]["color"] = resolve_color(run["style"]["color"], scheme)
        kind = slot_kind(node, ph_type, True)
        nv = node.find("./p:nvSpPr/p:cNvPr", NS)
        name = nv.get("name") if nv is not None else "text"
        base = kind if kind != "text" else ("title" if style["size"] >= 24 else "body")
        sid = slug(base, "text")
        n = 1
        while sid in used_ids:
            n += 1
            sid = "%s%d" % (slug(base, "text"), n)
        used_ids.add(sid)
        fit = fit_estimate(box, style["size"], style["align"])
        slots.append({
            "id": sid,
            "name": name,
            "kind": kind,
            "box": box,
            "text": text,
            "sample": text,
            "style": style,
            "paragraphs": [{"text": para["text"], "runs": [{"text": r["text"], "style": r["style"]} for r in para["runs"]]}
                           for para in paragraphs],
            "fit": fit,
            "optional": kind in ("footer", "date", "slideNumber"),
        })

    layout_ph_by_idx = {}
    try:
        for shape in layout.placeholders:
            fmt = shape.placeholder_format
            try:
                from pptx.util import Emu as _Emu
                layout_ph_by_idx[str(fmt.idx)] = (
                    int(shape.left or 0), int(shape.top or 0),
                    int(shape.width or 0), int(shape.height or 0),
                )
            except Exception:
                continue
    except Exception:
        pass
    # The tree root itself is p:spTree (not a shape), so start from its children.
    for child in slide.shapes._spTree:
        visit(child, IDENTITY, layout_ph_by_idx)
    return slots


# ── main ──────────────────────────────────────────────────────────────────────

def decompose(pptx_path, out_dir, max_slides=0, do_render=True, width=1600, height=900, script_dir=None):
    prs = Presentation(pptx_path)
    scheme = theme_colors(prs)
    fonts = theme_fonts(prs)
    theme_font = fonts.get("minor", {}).get("ea") or fonts.get("minor", {}).get("latin") or "微软雅黑"

    os.makedirs(out_dir, exist_ok=True)
    render = {"ok": False, "engine": "none", "error": "render disabled"}
    if do_render:
        script = os.path.join(script_dir or os.path.dirname(os.path.abspath(__file__)), "pptx_render.ps1")
        render = render_slides(pptx_path, out_dir, width, height, max_slides, script)
        if not render.get("ok"):
            pillow = pillow_render(prs, out_dir, 1280, 720)
            pillow["fallbackFrom"] = render.get("error")
            render = pillow

    total = len(prs.slides)
    limit = total if not max_slides else min(total, max_slides)

    slides = []
    font_counter = {}
    color_counter = {}
    for index, slide in enumerate(prs.slides):
        if index >= limit:
            break
        layout = slide.slide_layout
        slots = collect_slots(slide, index, layout, prs.slide_masters[0], scheme, theme_font)
        for slot in slots:
            font = slot["style"].get("font")
            if font:
                font_counter[font] = font_counter.get(font, 0) + 1
            color = slot["style"].get("color")
            if color:
                color_counter[color] = color_counter.get(color, 0) + 1
        fields, repeat_groups = build_fields(slots)
        chars = sum(len(slot["text"]) for slot in slots)
        role = classify_role(index, total, slots, len(slots), chars, fields)
        plate_rel = "plates/s%d.png" % index
        ref_rel = "reference/s%d.png" % index
        plate_abs = os.path.join(out_dir, plate_rel)
        ref_abs = os.path.join(out_dir, ref_rel)
        record = {
            "index": index,
            "role": role,
            "layoutIndex": list(prs.slide_layouts).index(layout) if layout in list(prs.slide_layouts) else 0,
            "layoutName": layout.name,
            "plate": plate_rel if os.path.exists(plate_abs) else None,
            "reference": ref_rel if os.path.exists(ref_abs) else None,
            "slots": slots,
            "fields": fields,
            "repeatGroups": [g for g in repeat_groups if len(g["fields"]) > 1],
        }
        fillable = [f for f in fields if not re.match(r"^(badge|tocTitle)$", f["id"])]
        record["capacity"] = {
            "fields": len(fillable),
            "maxCjk": sum(f["fit"]["maxCjk"] for f in fillable),
            "bodyMaxCjk": sum(f["fit"]["maxCjk"] for f in fillable if f["id"].startswith("body")),
        }
        record["guidance"] = "%s页：%s" % (
            role,
            "；".join("%s=%r" % (f["id"], f["sample"][:14] + ("…" if len(f["sample"]) > 14 else ""))
                     for f in fields[:6]) or "无可用文本槽",
        )
        if os.path.exists(plate_abs):
            record["plateHash"] = dhash(plate_abs)
            tokens = palette_from_image(plate_abs)
            record["palette"] = tokens["palette"]
            if index == 0:
                record["_backgroundHint"] = tokens["background"]
        slides.append(record)

    # Shrink the renders, then point the spec at what survived.
    optimised = optimize_assets(out_dir, width)
    rewrite_asset_paths(slides, out_dir)

    # Inherited colours have to be decided against the rendered plate.
    try:
        apply_contrast_colors(slides, out_dir, emu_in(prs.slide_width), emu_in(prs.slide_height))
    except Exception as exc:  # noqa: BLE001 - a colour fallback must never kill a decomposition
        sys.stderr.write("contrast pass skipped: %s\n" % exc)

    # Group slides whose plate looks the same: those are interchangeable layouts.
    families = []
    for record in slides:
        placed = False
        for family in families:
            anchor = next((s for s in slides if s["index"] == family["slides"][0]), None)
            if anchor and hash_distance(record.get("plateHash"), anchor.get("plateHash")) <= 8:
                family["slides"].append(record["index"])
                record["family"] = family["id"]
                placed = True
                break
        if not placed:
            family = {"id": "f%d" % (len(families) + 1), "role": record["role"], "slides": [record["index"]]}
            record["family"] = family["id"]
            families.append(family)

    for record in slides:
        record.pop("plateHash", None)
        record.pop("_backgroundHint", None)

    # Deck-level palette prefers the cover plate, then merges the rest.
    cover_plate = None
    for candidate in ("s0.jpg", "s0.png"):
        path = os.path.join(out_dir, "plates", candidate)
        if os.path.exists(path):
            cover_plate = path
            break
    design = {"palette": [], "background": None, "fonts": {"title": theme_font, "body": theme_font},
              "fontUsage": [], "textColors": []}
    if cover_plate:
        tokens = palette_from_image(cover_plate)
        design["palette"] = tokens["palette"]
        design["background"] = tokens["background"]
    if not design["palette"]:
        merged = {}
        for record in slides:
            for entry in record.get("palette", []):
                merged[entry["hex"]] = merged.get(entry["hex"], 0) + entry["share"]
        ranked = sorted(merged.items(), key=lambda kv: kv[1], reverse=True)[:8]
        total_share = sum(v for _, v in ranked) or 1.0
        design["palette"] = [{"hex": h, "share": round(v / total_share, 3)} for h, v in ranked]
    design["fontUsage"] = [{"name": k, "count": v} for k, v in
                           sorted(font_counter.items(), key=lambda kv: kv[1], reverse=True)[:8]]
    if design["fontUsage"]:
        design["fonts"]["body"] = design["fontUsage"][0]["name"]
        design["fonts"]["title"] = design["fontUsage"][0]["name"]
    design["textColors"] = [{"hex": k, "count": v} for k, v in
                            sorted(color_counter.items(), key=lambda kv: kv[1], reverse=True)[:8]]

    return {
        "ok": True,
        "version": 2,
        "path": pptx_path,
        "size": {"w": round(prs.slide_width / EMU, 3), "h": round(prs.slide_height / EMU, 3)},
        "slideCount": total,
        "decomposedSlides": len(slides),
        "engine": render.get("engine"),
        "assets": optimised,
        "render": {
            "ok": bool(render.get("ok")),
            "engine": render.get("engine"),
            "error": render.get("error"),
            "fallbackFrom": render.get("fallbackFrom"),
            "width": render.get("width"),
            "height": render.get("height"),
            "textHidden": render.get("engine") == "powerpoint",
        },
        "theme": {"colors": scheme, "fonts": fonts},
        "design": design,
        "families": families,
        "slides": slides,
    }


def main(argv):
    args = [a for a in argv[1:]]
    flags = {"--no-render": False, "--max-slides": 0, "--width": 1600, "--height": 900}
    positional = []
    i = 0
    while i < len(args):
        item = args[i]
        if item in ("--no-render",):
            flags["--no-render"] = True
        elif item == "--max-slides":
            i += 1
            flags["--max-slides"] = int(args[i])
        elif item == "--width":
            i += 1
            flags["--width"] = int(args[i])
        elif item == "--height":
            i += 1
            flags["--height"] = int(args[i])
        else:
            positional.append(item)
        i += 1
    if len(positional) < 2:
        print(json.dumps({"ok": False, "error": "usage: pptx_template.py <file.pptx> <outDir> [--no-render] [--max-slides N]"}))
        return 2
    try:
        result = decompose(
            positional[0], positional[1],
            max_slides=flags["--max-slides"],
            do_render=not flags["--no-render"],
            width=flags["--width"], height=flags["--height"],
        )
        print(json.dumps(result, ensure_ascii=False))
        return 0
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
