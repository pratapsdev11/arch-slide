#!/usr/bin/env python3
"""arch-slide: editable architecture slides from a JSON spec, in any org's style.

  text     <deck.pptx>                              dump slide text, tables, notes
  template <deck.pptx> --slide N -o base.pptx       base template: slide N's layout + repeated chrome, no slides
  style    <deck.pptx> --slide N -o style.json      draft style tokens from slide N
  icons    <svg_dir> <out_dir> [--color HEX]        SVG -> PNG (one-time)
  build    <spec.json>... --style S -o out.pptx     render; prints a JSON report
  preview  <deck.pptx>                              one PNG per slide next to the deck
"""
import argparse
import copy
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

HERE = Path(__file__).resolve().parent
STYLE_HOME = Path.home() / ".arch-slide" / "styles"
EMU = 914400

# Neutral defaults. A style.json only needs the keys it changes.
DEFAULT_STYLE = {
    "template": None,          # base.pptx next to style.json; None = blank 16:9 + chrome below
    "layout": "Architecture",
    "font": "Calibri",
    "sizes": {"title": 20, "header": 10, "name": 9, "desc": 8, "legend": 8, "minFontPt": 8},
    "limits": {"nameMaxChars": 32, "descMaxChars": 70},
    "colors": {"primary": "1F4E79", "headerText": "1F4E79", "text": "262626", "muted": "595959",
               "panelFill": "FFFFFF", "panelBorder": "BDD7EE", "cardFill": "F7F9FC", "cardBorder": "DDE3EA"},
    "panel": {"header": "text", "radius": 0.06, "border": 0.75, "shadow": True, "pad": 0.06},
    "gap": [0.32, 0.08, 0.06],  # sibling gap by panel depth; depth 0 leaves room for arrows
    "node": {"icon": 0.32, "iconLeftRatio": 1.25},
    "content": {"left": 0.3, "top": 1.0, "right": 0.3, "bottom": 0.45},
    "slide": {"w": 13.333, "h": 7.5},
    "chrome": {"accentBar": True, "footerBand": 0.25, "footerText": "", "logo": None},  # blank template only
    "flows": {"data": {"color": "000000", "label": "Data Flow"},
              "analytics": {"color": "1F4E79", "label": "Processing / Analytics"},
              "action": {"color": "ED7D31", "label": "Actions / Feedback"}},
    "legend": {"title": "Legend"},
}


def E(inches):
    return Emu(int(round(inches * EMU)))


def rgb(hex_):
    return RGBColor.from_string(hex_.lstrip("#").upper())


def merge(base, over):
    out = dict(base)
    for k, v in over.items():
        deep = isinstance(v, dict) and isinstance(base.get(k), dict) and k != "flows"
        out[k] = merge(base[k], v) if deep else v
    return out


def load_style(ref):
    if not ref:
        return dict(DEFAULT_STYLE, _dir=HERE)
    p = Path(ref).expanduser()
    if not p.exists():
        p = STYLE_HOME / ref
    if p.is_dir():
        p = p / "style.json"
    if not p.is_file():
        sys.exit(f"style not found: {ref} (looked in {STYLE_HOME})")
    s = merge(DEFAULT_STYLE, json.loads(p.read_text(encoding="utf-8")))
    s["_dir"] = p.parent
    return s


# ---------- shape helpers ----------

def plain(shape):
    """Drop the theme style python-pptx attaches, so only explicit formatting applies."""
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)
    return shape


def box_shape(shapes, kind, x, y, w, h, fill=None, line=None, width=0.75, dash=None, radius=None, name=None):
    sh = plain(shapes.add_shape(kind, E(x), E(y), E(w), E(h)))
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(width)
    else:
        sh.line.fill.background()
    if dash:
        sh.line.dash_style = dash
    if radius is not None:
        sh.adjustments[0] = min(0.5, radius / max(min(w, h), 0.01))
    if name:
        sh.name = name
    return sh


SHADOW = ('<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
          '<a:outerShdw blurRad="50800" dist="12700" dir="5400000" algn="t" rotWithShape="0">'
          '<a:srgbClr val="000000"><a:alpha val="18000"/></a:srgbClr></a:outerShdw></a:effectLst>')


def add_shadow(sh):
    sh._element.spPr.append(etree.fromstring(SHADOW))  # after a:ln, which box_shape always creates


def write(tf, paras, style, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, margin=0.03):
    """paras: [(text, pt, color, bold)]. Font never goes below style minFontPt."""
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = E(margin)
    for i, (text, size, color, bold) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text
        f = r.font
        f.size = Pt(max(size, style["sizes"]["minFontPt"]))
        f.bold = bold
        f.name = style["font"]
        f.color.rgb = rgb(color)


# ponytail: width-per-char heuristic (0.52 em), not real font metrics; the PNG preview is the real check
def text_lines(text, size, width_in):
    cpl = max(1, int(width_in * 72 / (size * 0.52)))
    return max(1, math.ceil(len(text) / cpl))


def text_height(paras, width_in):
    return sum(text_lines(t, s, width_in) * s * 1.2 for t, s, *_ in paras) / 72


def arrow(shapes, x1, y1, x2, y2, flow, start_head=False, end_head=True, elbow=False, name=None):
    kind = MSO_CONNECTOR.ELBOW if elbow else MSO_CONNECTOR.STRAIGHT
    cx = plain(shapes.add_connector(kind, E(x1), E(y1), E(x2), E(y2)))
    if elbow:  # one bend: leaves start horizontally, enters end vertically
        cx._element.spPr.find(qn("a:prstGeom")).set("prst", "bentConnector2")
    cx.line.color.rgb = rgb(flow["color"])
    cx.line.width = Pt(flow.get("width", 1.25))
    if flow.get("dash"):
        cx.line.dash_style = getattr(MSO_LINE_DASH_STYLE, flow["dash"].upper())
    ln = cx.line._get_or_add_ln()
    if start_head:
        etree.SubElement(ln, qn("a:headEnd"), type="triangle")
    if end_head:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle")
    if name:
        cx.name = name
    return cx


# ---------- layout ----------

def walk(box):
    yield box
    for k in box.get("children", []):
        yield from walk(k)


def place(ctx, box, x, y, w, h, depth):
    """Recursive split: each box splits its rect among children by weight along dir."""
    if box.get("legend"):
        ctx["legend_rect"] = (x, y, w, h)
        return
    if "name" in box:
        return draw_node(ctx, box, x, y, w, h)
    if "id" in box:
        ctx["rects"][box["id"]] = (x, y, w, h)
    if box.get("title") or box.get("panel"):
        x, y, w, h = draw_panel(ctx, box, x, y, w, h, depth)
        depth += 1
    kids = box.get("children", [])
    if not kids:
        return
    gaps = ctx["style"]["gap"]
    gap = box.get("gap", gaps[min(depth, len(gaps) - 1)])
    row = box.get("dir", "col") == "row"
    total = sum(k.get("weight", 1) for k in kids)
    pos, span = (x, w) if row else (y, h)
    span -= gap * (len(kids) - 1)
    for k in kids:
        size = span * k.get("weight", 1) / total
        if row:
            place(ctx, k, pos, y, size, h, depth)
        else:
            place(ctx, k, x, pos, w, size, depth)
        pos += size + gap


def draw_panel(ctx, box, x, y, w, h, depth):
    s, c, p = ctx["style"], ctx["style"]["colors"], ctx["style"]["panel"]
    shapes, top = ctx["slide"].shapes, depth == 0
    bid = box.get("id", "")
    hsize = s["sizes"]["header"] - (0 if top else 1)
    title = box.get("title", "")
    hh = text_lines(title, hsize, w - 0.1) * hsize * 1.2 / 72 + 0.1 if title else p["pad"]
    sh = box_shape(shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                   fill=c["panelFill"] if top else c["cardFill"],
                   line=c["panelBorder"] if top else c["cardBorder"],
                   width=p["border"], radius=p["radius"], name=f"{'zone' if top else 'card'}:{bid}")
    if p["shadow"] and top:
        add_shadow(sh)
    if title and top and p["header"] == "bar":
        bar = box_shape(shapes, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x, y, w, hh, fill=c["primary"],
                        radius=p["radius"], name=f"header:{bid}")
        write(bar.text_frame, [(title, hsize, "FFFFFF", True)], s, MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER)
    elif title:
        write(sh.text_frame, [(title, hsize, c["headerText"], True)], s, MSO_ANCHOR.TOP, PP_ALIGN.CENTER, 0.04)
    pad = p["pad"]
    return x + pad, y + hh, w - 2 * pad, h - hh - pad


def draw_node(ctx, n, x, y, w, h):
    s, c, z = ctx["style"], ctx["style"]["colors"], ctx["style"]["sizes"]
    nid = n.get("id") or n["name"]
    ctx["rects"][nid] = (x, y, w, h)
    if n.get("boxed"):  # node drawn as its own small card (agent cards etc.)
        p = s["panel"]
        box_shape(ctx["slide"].shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=c["cardFill"],
                  line=c["cardBorder"], width=p["border"], radius=p["radius"], name=f"card:{nid}")
        x, y, w, h = x + p["pad"], y + p["pad"], w - 2 * p["pad"], h - 2 * p["pad"]
    desc = n.get("desc") or []
    desc = [desc] if isinstance(desc, str) else desc
    dot = "• " if len(desc) > 1 else ""
    name_color = c["headerText"] if n.get("boxed") else c["text"]
    paras = [(n["name"], z["name"], name_color, True)] + [(dot + d, z["desc"], c["muted"], False) for d in desc]
    icon = n.get("icon")
    isz = min(s["node"]["icon"], h * 0.6, w * 0.45) if icon else 0
    left = n.get("iconPos", "left" if w >= s["node"]["iconLeftRatio"] * h else "top") == "left"
    tw = w - isz - 0.05 if left and icon else w
    th_need = text_height(paras, tw - 0.06) + 0.06
    content_h = max(isz, th_need) if left else isz + 0.03 + th_need
    y0 = y + max(0.0, (h - content_h) / 2)  # center the node vertically in its slot
    grp = ctx["slide"].shapes.add_group_shape()
    grp.name = f"node:{nid}"
    if icon:
        add_icon(ctx, grp.shapes, icon, x if left else x + (w - isz) / 2, y0, isz, nid)
    tx, ty = (x + isz + 0.05, y0) if left and icon else (x, y0 + (isz + 0.03 if icon else 0))
    tb = grp.shapes.add_textbox(E(tx), E(ty), E(tw), E(y + h - ty))
    tb.name = f"text:{nid}"
    write(tb.text_frame, paras, s, align=PP_ALIGN.LEFT if left else PP_ALIGN.CENTER)
    check_fit(ctx, nid, paras, tw - 0.06, y + h - ty - 0.06)


def find_icon(ctx, icon):
    for d in ctx["icon_dirs"]:
        p = d / f"{icon}.png"
        if p.is_file():
            return p
    return None


def add_icon(ctx, shapes, icon, x, y, size, nid):
    path = find_icon(ctx, icon)
    ctx["report"]["icons_used"][nid] = icon
    if not path:
        ctx["report"]["missing_icons"].append({"node": nid, "icon": icon})
        ph = box_shape(shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, size, size, fill="F2F2F2", line="A6A6A6",
                       dash=MSO_LINE_DASH_STYLE.DASH, radius=0.04, name=f"missing-icon:{nid}")
        write(ph.text_frame, [("?", 10, "7F7F7F", True)], ctx["style"], MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, 0)
        return
    from PIL import Image  # installed with python-pptx
    with Image.open(path) as im:
        iw, ih = im.size
    k = size / max(iw, ih)
    pic = shapes.add_picture(str(path), E(x + (size - iw * k) / 2), E(y + (size - ih * k) / 2), E(iw * k), E(ih * k))
    pic.name = f"icon:{icon}"


def check_fit(ctx, nid, paras, w, h):
    lim, rep = ctx["style"]["limits"], ctx["report"]
    need = text_height(paras, w)
    if need > h + 0.01:
        rep["overflow"].append({"id": nid, "need_in": round(need, 2), "have_in": round(h, 2)})
    for i, (t, *_) in enumerate(paras):
        cap = lim["nameMaxChars"] if i == 0 else lim["descMaxChars"]
        if len(t.lstrip("• ")) > cap:
            rep["too_long"].append({"id": nid, "text": t, "max": cap})


def draw_flow(ctx, f):
    rects, s = ctx["rects"], ctx["style"]
    if f.get("from") not in rects or f.get("to") not in rects or f.get("type") not in s["flows"]:
        return  # already listed in report["errors"] by check_spec
    (ax, ay, aw, ah), (bx, by, bw, bh) = rects[f["from"]], rects[f["to"]]
    sep_x = ax + aw <= bx or bx + bw <= ax
    sep_y = ay + ah <= by or by + bh <= ay
    elbow = sep_x and sep_y
    both = f.get("both", False)
    start_head, end_head = both, True
    if sep_x and not sep_y:      # side by side: horizontal at the middle of the vertical overlap
        ym = (max(ay, by) + min(ay + ah, by + bh)) / 2
        x1, x2 = (ax + aw, bx) if ax < bx else (ax, bx + bw)
        y1 = y2 = ym
    elif sep_y and not sep_x:    # stacked: vertical at the middle of the horizontal overlap
        xm = (max(ax, bx) + min(ax + aw, bx + bw)) / 2
        y1, y2 = (ay + ah, by) if ay < by else (ay, by + bh)
        x1 = x2 = xm
    elif elbow:                  # diagonal: L from the wider box's side to the narrower box's top/bottom
        a_side = aw >= bw
        (sx, sy, sw, sh_), (tx, ty, tw, th) = (rects[f["from"]], rects[f["to"]]) if a_side else \
                                              (rects[f["to"]], rects[f["from"]])
        tcx = tx + tw / 2
        x1, y1 = (sx + sw if tcx > sx + sw else sx), sy + sh_ / 2
        x2, y2 = tcx, (ty if ty > y1 else ty + th)
        start_head, end_head = (both, True) if a_side else (True, both)
    else:
        ctx["report"]["errors"].append(f"flow {f['from']}->{f['to']}: boxes overlap (parent/child?), skipped")
        return
    flow = s["flows"][f["type"]]
    ctx["used"].add(f["type"])
    shapes = ctx["slide"].shapes
    arrow(shapes, x1, y1, x2, y2, flow, start_head, end_head, elbow, f"flow:{f['from']}->{f['to']}")
    if f.get("step") is not None:
        d = 0.22
        mx, my = (x1 + x2) / 2, (y1 if elbow else (y1 + y2) / 2)  # elbow: middle of the horizontal leg
        mx, my = (mx + d * 0.7, my) if x1 == x2 else (mx, my - d * 0.7)  # beside the line, not on it
        b = box_shape(shapes, MSO_SHAPE.OVAL, mx - d / 2, my - d / 2, d, d, fill=flow["color"], name=f"step:{f['step']}")
        write(b.text_frame, [(str(f["step"]), s["sizes"]["minFontPt"], "FFFFFF", True)], s,
              MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, 0)


def draw_legend(ctx):
    s, c = ctx["style"], ctx["style"]["colors"]
    used = [t for t in s["flows"] if t in ctx["used"]]
    if "legend_rect" not in ctx or not used:
        return
    x, y, w, h = ctx["legend_rect"]
    sz = s["sizes"]["legend"]
    grp = ctx["slide"].shapes.add_group_shape()
    grp.name = "legend"
    bg = box_shape(grp.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=c["panelFill"],
                   line=c["panelBorder"], radius=s["panel"]["radius"], name="legend:box")
    write(bg.text_frame, [(s["legend"]["title"], sz + 1, c["headerText"], True)], s, margin=0.06)
    head = (sz + 1) * 1.2 / 72 + 0.12
    row = min(0.24, (h - head - 0.04) / len(used))
    if row * 72 < sz * 1.2:
        ctx["report"]["overflow"].append({"id": "legend", "need_in": round(head + len(used) * sz * 1.2 / 72, 2),
                                          "have_in": round(h, 2)})
    for i, t in enumerate(used):
        ly = y + head + i * row + row / 2
        arrow(grp.shapes, x + 0.1, ly, x + 0.45, ly, s["flows"][t], name=f"legend:{t}")
        tb = grp.shapes.add_textbox(E(x + 0.52), E(ly - row / 2), E(w - 0.58), E(row))
        write(tb.text_frame, [(s["flows"][t]["label"], sz, c["text"], False)], s, MSO_ANCHOR.MIDDLE, margin=0.01)


def check_spec(spec, s, rep):
    ids = [b.get("id") or b["name"] for b in walk(spec["root"]) if "id" in b or "name" in b]
    dup = sorted({i for i in ids if ids.count(i) > 1})
    if dup:
        rep["errors"].append(f"duplicate ids: {dup}")
    for f in spec.get("flows", []):
        for end in ("from", "to"):
            if f.get(end) not in ids:
                rep["errors"].append(f"flow {f.get('from')}->{f.get('to')}: unknown {end} id {f.get(end)!r}")
        if f.get("type") not in s["flows"]:
            rep["errors"].append(f"flow {f.get('from')}->{f.get('to')}: type {f.get('type')!r} "
                                 f"not in style flows {list(s['flows'])}")


# ---------- template / slides ----------

def drop_slides(prs):
    lst = prs.slides._sldIdLst
    for sid in list(lst):
        lst.remove(sid)
        prs.part.drop_rel(sid.rId)
    for sec in prs.part._element.iter("{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst"):
        ext = sec.getparent()
        ext.getparent().remove(ext)  # sections list slide ids that no longer exist


def find_layout(prs, name):
    for m in prs.slide_masters:
        for lay in m.slide_layouts:
            if lay.name == name:
                return lay
    sys.exit(f"layout {name!r} not in template; run `build.py template` on the org deck first")


def blank_chrome(ctx, W, H):
    s, c = ctx["style"], ctx["style"]["colors"]
    ch, shapes = s["chrome"], ctx["slide"].shapes
    if ch.get("accentBar"):
        box_shape(shapes, MSO_SHAPE.RECTANGLE, 0, 0.22, 0.12, 0.62, fill=c["primary"], name="chrome:accent")
    band = ch.get("footerBand") or 0
    if band:
        fb = box_shape(shapes, MSO_SHAPE.RECTANGLE, 0, H - band, W, band, fill=c["primary"], name="chrome:footer")
        if ch.get("footerText"):
            write(fb.text_frame, [(ch["footerText"], s["sizes"]["minFontPt"] + 1, "FFFFFF", True)], s,
                  MSO_ANCHOR.MIDDLE, PP_ALIGN.RIGHT, 0.2)
    if ch.get("logo"):
        logo = shapes.add_picture(str(s["_dir"] / ch["logo"]), E(0), E(0.18), height=E(0.65))
        logo.left = E(W - 0.25) - logo.width
        logo.name = "chrome:logo"


def set_title(ctx, text, W):
    slide, s = ctx["slide"], ctx["style"]
    if slide.shapes.title is not None:
        slide.shapes.title.text = text
        return
    tb = slide.shapes.add_textbox(E(0.3), E(0.22), E(W - 2.8), E(0.62))
    tb.name = "title"
    write(tb.text_frame, [(text, s["sizes"]["title"], s["colors"]["text"], True)], s, MSO_ANCHOR.MIDDLE)


def cmd_build(a):
    s = load_style(a.style)
    tpl = s.get("template")
    if tpl:
        prs = Presentation(str(s["_dir"] / tpl))
        drop_slides(prs)
        layout = find_layout(prs, s["layout"])
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = E(s["slide"]["w"]), E(s["slide"]["h"])
        layout = prs.slide_layouts[6]  # Blank
    W, H = prs.slide_width / EMU, prs.slide_height / EMU
    rep = {"out": str(Path(a.out).resolve()), "errors": [], "missing_icons": [], "overflow": [],
           "too_long": [], "icons_used": {}}
    icon_dirs = [s["_dir"] / "icons", HERE / "assets" / "icons"]
    for spec_path in a.specs:
        spec = json.loads(Path(spec_path).read_text(encoding="utf-8"))
        check_spec(spec, s, rep)
        slide = prs.slides.add_slide(layout)
        for ph in list(slide.placeholders):  # templates not prepped by `template` may carry content placeholders
            if ph.placeholder_format.type not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                ph._element.getparent().remove(ph._element)
        ctx = {"slide": slide, "style": s, "report": rep, "rects": {}, "used": set(), "icon_dirs": icon_dirs}
        if not tpl:
            blank_chrome(ctx, W, H)
        set_title(ctx, spec.get("title", ""), W)
        m = s["content"]
        place(ctx, spec["root"], m["left"], m["top"], W - m["left"] - m["right"], H - m["top"] - m["bottom"], 0)
        for f in spec.get("flows", []):
            draw_flow(ctx, f)
        draw_legend(ctx)
    prs.save(a.out)
    print(json.dumps(rep, indent=1, ensure_ascii=False))


def cmd_template(a):
    prs = Presentation(a.deck)
    slides = list(prs.slides)
    src = slides[a.slide - 1]
    layout = src.slide_layout
    peers = [x for x in slides if x is not src and x.slide_layout.part is layout.part] or \
            [x for x in slides if x is not src]

    def sig(sh):
        return sh.shape_type, sh.left, sh.top, sh.width, sh.height

    share = Counter()
    for p in peers:
        share.update({sig(sh) for sh in p.shapes if not sh.is_placeholder})
    W, H, edge = prs.slide_width, prs.slide_height, E(0.3)

    def at_edge(sh):
        return sh.left <= edge or sh.top <= edge or sh.left + sh.width >= W - edge or sh.top + sh.height >= H - edge

    # ponytail: chrome = repeated on another slide AND touching a slide edge; --keep overrides when that guesses wrong
    def sp_text(sp):
        return "".join(t.text or "" for t in sp.iter(qn("a:t"))).strip()

    peer_texts = {sp_text(sp) for p in peers for sp in p._element.iter(qn("p:sp"))}

    def per_slide(e):
        """Placeholders and text that changes per slide (the old title) are not chrome; a repeated tagline is."""
        return e.tag == qn("p:sp") and (e.find(f".//{qn('p:ph')}") is not None or
                                        (sp_text(e) and sp_text(e) not in peer_texts))

    keep_names = [k.strip() for k in a.keep.split(",")] if a.keep else None
    cands = [sh for sh in src.shapes if not sh.is_placeholder]
    chrome = [sh for sh in cands if (sh.name in keep_names if keep_names else
                                     share[sig(sh)] and at_edge(sh) and not per_slide(sh._element))]

    tree = layout.shapes._spTree
    next_id = max(int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))) + 1
    for sh in chrome:
        el = copy.deepcopy(sh._element)
        for sp in [e for e in el.iter(qn("p:sp")) if e is not el and per_slide(e)]:
            sp.getparent().remove(sp)
        for e in el.iter():
            rid = e.get(qn("r:embed"))
            if rid:
                e.set(qn("r:embed"), layout.part.relate_to(src.part.related_part(rid), RT.IMAGE))
            if e.tag == qn("a:hlinkClick"):
                e.getparent().remove(e)
        for cnv in el.iter(qn("p:cNvPr")):
            cnv.set("id", str(next_id))
            next_id += 1
        tree.insert_element_before(el, "p:extLst")

    layout._element.find(qn("p:cSld")).set("name", a.name)
    keep = (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE, PP_PLACEHOLDER.DATE,
            PP_PLACEHOLDER.FOOTER, PP_PLACEHOLDER.SLIDE_NUMBER)
    for ph in list(layout.placeholders):
        if ph.placeholder_format.type not in keep:
            ph._element.getparent().remove(ph._element)

    drop_slides(prs)
    master = layout.slide_master
    mlist = prs.part._element.find(qn("p:sldMasterIdLst"))
    for mid in list(mlist):
        rid = mid.get(qn("r:id"))
        if prs.part.related_part(rid) is not master.part:
            mlist.remove(mid)
            prs.part.drop_rel(rid)
    llist = master._element.find(qn("p:sldLayoutIdLst"))
    for lid in list(llist):
        rid = lid.get(qn("r:id"))
        if master.part.related_part(rid) is not layout.part:
            llist.remove(lid)
            master.part.drop_rel(rid)

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    prs.save(a.out)
    i = lambda v: round(v / EMU, 2)  # noqa: E731
    print(json.dumps({"out": str(Path(a.out).resolve()), "layout": a.name,
                      "slide_in": [i(prs.slide_width), i(prs.slide_height)],
                      "shapes_on_slide": [{"name": sh.name, "xywh_in": [i(sh.left), i(sh.top), i(sh.width), i(sh.height)],
                                           "on_other_slides": share[sig(sh)], "at_edge": at_edge(sh),
                                           "copied_to_layout": sh in chrome} for sh in cands]}, indent=1))


def cmd_style(a):
    prs = Presentation(a.deck)
    sl = prs.slides[a.slide - 1]
    fills, lines, texts, fonts, sizes = Counter(), Counter(), Counter(), Counter(), Counter()
    for e in sl._element.iter(qn("a:srgbClr")):
        anc = {x.tag for x in e.iterancestors()}
        v = e.get("val").upper()
        if qn("a:rPr") in anc or qn("a:defRPr") in anc or qn("a:endParaRPr") in anc:
            texts[v] += 1
        elif qn("a:ln") in anc:
            lines[v] += 1
        else:
            fills[v] += 1
    theme_xml = etree.fromstring(sl.slide_layout.slide_master.part.part_related_by(RT.THEME).blob)
    theme_font = {t: theme_xml.find(f".//{{*}}{t}/{{*}}latin").get("typeface") for t in ("majorFont", "minorFont")}
    for e in sl._element.iter(qn("a:latin")):
        tf = e.get("typeface")
        fonts[theme_font["majorFont"] if tf == "+mj-lt" else theme_font["minorFont"] if tf == "+mn-lt" else tf] += 1
    for e in sl._element.iter(qn("a:rPr")):
        if e.get("sz"):
            sizes[int(e.get("sz")) / 100] += 1

    def sat(h):
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return max(r, g, b) - min(r, g, b)

    colored = [c for c, _ in (fills + lines).most_common() if sat(c) > 60]
    primary = colored[0] if colored else DEFAULT_STYLE["colors"]["primary"]
    text = next((c for c, _ in texts.most_common() if sat(c) <= 30), DEFAULT_STYLE["colors"]["text"])
    body = sizes.most_common(1)[0][0] if sizes else DEFAULT_STYLE["sizes"]["name"]
    draft = {
        "name": a.name,
        "template": "base.pptx",
        "layout": "Architecture",
        "font": fonts.most_common(1)[0][0] if fonts else theme_font["minorFont"],
        "sizes": {"name": body, "desc": max(body - 1, 8), "header": body + 1.5},
        "colors": {"primary": primary, "headerText": primary, "text": text},
        "flows": {"data": {"color": "000000", "label": "Data Flow"},
                  "analytics": {"color": primary, "label": "Processing / Analytics"},
                  "action": {"color": colored[1] if len(colored) > 1 else "2E5BC0", "label": "Actions / Feedback"}},
        "_observed": {"fills": fills.most_common(10), "lines": lines.most_common(8), "text": texts.most_common(8),
                      "fonts": fonts.most_common(5), "sizes_pt": sizes.most_common(8), "theme_fonts": theme_font},
    }
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(draft, indent=1), encoding="utf-8")
    print(json.dumps(draft, indent=1))


def cmd_icons(a):
    import pymupdf
    out = Path(a.out_dir)
    out.mkdir(parents=True, exist_ok=True)
    done, skipped = 0, []
    for svg in sorted(Path(a.svg_dir).rglob("*.svg")):
        name = re.sub(r"^\d+-icon-service-", "", svg.stem)  # Azure pack: 10021-icon-service-Data-Factories
        name = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
        dst = out / f"{name}.png"
        if dst.exists():
            skipped.append(svg.name)
            continue
        data = svg.read_bytes()
        if a.color:
            data = data.replace(b"currentColor", f"#{a.color.lstrip('#')}".encode())
        page = pymupdf.open(stream=data, filetype="svg")[0]
        z = a.size / max(page.rect.width, page.rect.height)
        page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=True).save(str(dst))
        done += 1
    print(json.dumps({"converted": done, "skipped_existing": len(skipped), "out": str(out.resolve())}))


def cmd_preview(a):
    deck = Path(a.deck).resolve()
    prs = Presentation(str(deck))
    W = 1600
    Hpx = int(W * prs.slide_height / prs.slide_width)
    stem = deck.with_suffix("")
    outs = []
    if sys.platform == "win32":  # PowerPoint renders exactly what users will see
        q = lambda p: str(p).replace("'", "''")  # noqa: E731
        ps = (f"$ErrorActionPreference='Stop'; $app=New-Object -ComObject PowerPoint.Application; "
              f"$p=$app.Presentations.Open('{q(deck)}', $true, $false, $false); $i=1; "
              f"foreach($s in $p.Slides){{ $s.Export('{q(stem)}-' + $i + '.png','PNG',{W},{Hpx}); $i++ }}; "
              f"$p.Close(); if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}")
        r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
        if r.returncode == 0:
            outs = [f"{stem}-{i}.png" for i in range(1, len(prs.slides) + 1)]
    if not outs:
        soffice = shutil.which("soffice") or shutil.which("libreoffice") or next(
            (p for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                         "/Applications/LibreOffice.app/Contents/MacOS/soffice") if Path(p).exists()), None)
        if not soffice:
            sys.exit("no renderer: install LibreOffice (or PowerPoint on Windows)")
        import pymupdf
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(deck)],
                           check=True, capture_output=True)
            for i, page in enumerate(pymupdf.open(Path(tmp) / f"{deck.stem}.pdf"), 1):
                page.get_pixmap(dpi=int(W / (prs.slide_width / EMU))).save(f"{stem}-{i}.png")
                outs.append(f"{stem}-{i}.png")
    print(json.dumps({"png": outs}, indent=1))


def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def cmd_text(a):
    for i, sl in enumerate(Presentation(a.deck).slides, 1):
        print(f"\n## Slide {i} ({sl.slide_layout.name})")
        for sh in iter_shapes(sl.shapes):
            if sh.has_text_frame and sh.text_frame.text.strip():
                print(sh.text_frame.text.strip())
            if getattr(sh, "has_table", False) and sh.has_table:
                for r in sh.table.rows:
                    print(" | ".join(c.text.strip() for c in r.cells))
        if sl.has_notes_slide and sl.notes_slide.notes_text_frame.text.strip():
            print("Notes:", sl.notes_slide.notes_text_frame.text.strip())


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("text"); p.add_argument("deck")  # noqa: E702
    p = sub.add_parser("template"); p.add_argument("deck"); p.add_argument("--slide", type=int, required=True)  # noqa: E702
    p.add_argument("-o", "--out", required=True); p.add_argument("--name", default="Architecture")  # noqa: E702
    p.add_argument("--keep", help="comma-separated shape names to copy as chrome (overrides auto-detect)")
    p = sub.add_parser("style"); p.add_argument("deck"); p.add_argument("--slide", type=int, required=True)  # noqa: E702
    p.add_argument("-o", "--out", required=True); p.add_argument("--name", default="org")  # noqa: E702
    p = sub.add_parser("icons"); p.add_argument("svg_dir"); p.add_argument("out_dir")  # noqa: E702
    p.add_argument("--color"); p.add_argument("--size", type=int, default=256)  # noqa: E702
    p = sub.add_parser("build"); p.add_argument("specs", nargs="+")  # noqa: E702
    p.add_argument("--style", help="style dir/file or saved org name; omit for neutral"); p.add_argument("-o", "--out", required=True)  # noqa: E702
    p = sub.add_parser("preview"); p.add_argument("deck")  # noqa: E702
    a = ap.parse_args()
    {"text": cmd_text, "template": cmd_template, "style": cmd_style, "icons": cmd_icons,
     "build": cmd_build, "preview": cmd_preview}[a.cmd](a)


if __name__ == "__main__":
    main()
