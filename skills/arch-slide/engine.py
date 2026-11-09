"""arch-slide engine: spec -> content-aware layout -> routed arrows -> native PPTX shapes -> lint.

Layers:
  1. style   : org tokens + org rules (style.json), merged over neutral defaults
  2. model   : spec tree (+ Mermaid sequence -> flows, see seq.py)
  3. layout  : boxes sized from measured content (layout.py), extra space shared by weight; then a second pass
               slides nodes inside their card so connected nodes line up and arrows run straight (align_offsets)
  4. routing : orthogonal arrows around boxes (layout.route), labels placed clear of boxes
  5. render  : native, grouped, named PowerPoint shapes
  6. lint    : rule engine (lint.py) -> lint.json next to the spec
"""
import itertools
import json
import re
import statistics
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import layout as L
import lint
import seq as S

HERE = Path(__file__).resolve().parent
STYLE_HOME = Path.home() / ".arch-slide" / "styles"
ICON_HOME = Path.home() / ".arch-slide" / "icons"
EMU = 914400
TM = 0.03  # text box inner margin, inches

# Neutral defaults. A style.json only needs the keys it changes; "rules" holds the org's guideline values.
DEFAULT_STYLE = {
    "template": None,          # base.pptx next to style.json; None = blank 16:9 + chrome below
    "layout": "Architecture",
    "font": "Calibri",
    "sizes": {"title": 20, "subtitle": 12, "header": 10, "name": 9, "desc": 8, "legend": 8, "minFontPt": 8},
    "colors": {"primary": "1F4E79", "headerText": "1F4E79", "text": "262626", "muted": "595959", "subtitle": "1F4E79",
               "panelFill": "FFFFFF", "panelBorder": "BDD7EE", "cardFill": "F7F9FC", "cardBorder": "DDE3EA",
               "boundary": "8C8C8C"},
    "panel": {"header": "text", "align": "center", "radius": 0.06, "border": 0.75, "shadow": True, "pad": 0.06},
    "gap": [0.32, 0.08, 0.06],  # sibling gap by panel depth; depth 0 leaves room for arrows
    # icon height = ratio x name font size; one size per orientation keeps icons consistent
    "node": {"iconLeft": 2.8, "iconTop": 3.6, "iconLeftMinWidth": 1.6},
    "content": {"left": 0.3, "top": 1.0, "right": 0.3, "bottom": 0.45},
    "slide": {"w": 13.333, "h": 7.5},
    "chrome": {"accentBar": True, "footerBand": 0.25, "footerText": "", "logo": None},  # blank template only
    "flows": {"data": {"color": "000000", "label": "Data Flow"},
              "analytics": {"color": "1F4E79", "label": "Processing / Analytics"},
              "action": {"color": "ED7D31", "label": "Actions / Feedback"}},
    "legend": {"title": "Legend"},
    "rules": {},
}


def E(inches):
    return Emu(int(round(inches * EMU)))


def rgb(hex_):
    return RGBColor.from_string(hex_.lstrip("#").upper())


def merge(base, over):
    out = dict(base)
    for k, v in over.items():
        deep = isinstance(v, dict) and isinstance(base.get(k), dict) and k not in ("flows", "rules")
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


def fs(style, size):
    return max(size, style["sizes"]["minFontPt"])


def write(tf, paras, style, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, margin=TM):
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
        f.size = Pt(fs(style, size))
        f.bold = bold
        f.name = style["font"]
        f.color.rgb = rgb(color)


def _line_style(sh, flow, start_head, end_head):
    sh.line.color.rgb = rgb(flow["color"])
    sh.line.width = Pt(flow.get("width", 1.25))
    if flow.get("dash"):
        sh.line.dash_style = getattr(MSO_LINE_DASH_STYLE, flow["dash"].upper())
    ln = sh.line._get_or_add_ln()
    if start_head:
        etree.SubElement(ln, qn("a:headEnd"), type="triangle")
    if end_head:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle")


def arrow(shapes, pts, flow, start_head=False, end_head=True, name=None):
    """Straight connector for 2 points, editable freeform polyline for routed paths."""
    if len(pts) == 2:
        (x1, y1), (x2, y2) = pts
        sh = plain(shapes.add_connector(MSO_CONNECTOR.STRAIGHT, E(x1), E(y1), E(x2), E(y2)))
    else:
        fb = shapes.build_freeform(E(pts[0][0]), E(pts[0][1]), scale=1.0)
        fb.add_line_segments([(E(x), E(y)) for x, y in pts[1:]], close=False)
        sh = plain(fb.convert_to_shape())
        sh.fill.background()
    _line_style(sh, flow, start_head, end_head)
    if name:
        sh.name = name
    return sh


# ---------- measuring ----------

def header_dims(ctx, box, w, depth):
    """(header height, gap under it, header paragraphs, is a filled bar, text margin)."""
    s, c, p = ctx["style"], ctx["style"]["colors"], ctx["style"]["panel"]
    top = depth == 0
    hsize = fs(s, s["sizes"]["header"] - (0 if top else 1))
    title, sub = box.get("title", ""), box.get("subtitle", "")
    bar = bool(title) and top and p["header"] == "bar"
    paras = [(title, hsize, "FFFFFF" if bar else c["headerText"], True)] if title else []
    if sub:
        paras.append((sub, fs(s, hsize - 2), "FFFFFF" if bar else c["muted"], False))
    margin = 0.08 if (p.get("align") == "left" and not bar) else 0.04
    hh = (L.measure(paras, s["font"], w - 2 * margin)[0] + 2 * margin + 0.02) if paras else p["pad"]
    return hh, (p["pad"] if bar else 0), paras, bar, margin


def boundary_dims(ctx, box):
    s = ctx["style"]
    size = fs(s, s["sizes"]["header"] - 1)
    return L.line_height(s["font"], size, True) + 0.1, size


def node_dims(ctx, n, w):
    s, c, z = ctx["style"], ctx["style"]["colors"], ctx["style"]["sizes"]
    pad = s["panel"]["pad"] if n.get("boxed") else 0
    iw = w - 2 * pad
    icon = n.get("icon")
    orient = n.get("iconPos") or ("left" if iw >= s["node"]["iconLeftMinWidth"] else "top")
    name_pt = fs(s, z["name"])
    isz = s["node"]["iconLeft" if orient == "left" else "iconTop"] * name_pt / 72 if icon else 0
    desc = n.get("desc") or []
    desc = [desc] if isinstance(desc, str) else desc
    dot = "• " if len(desc) > 1 else ""
    name_color = c["headerText"] if n.get("boxed") else c["text"]
    paras = [(n["name"], name_pt, name_color, True)] + [(dot + d, fs(s, z["desc"]), c["muted"], False) for d in desc]
    tw = iw - (isz + 0.06 if icon and orient == "left" else 0)
    th, counts, too_wide = L.measure(paras, s["font"], tw - 2 * TM)
    th += 2 * TM
    content = max(isz, th) if orient == "left" else isz + (0.04 if icon else 0) + th
    return {"orient": orient, "isz": isz, "paras": paras, "tw": tw, "th": th, "counts": counts,
            "too_wide": too_wide, "pad": pad, "content": content, "min_h": content + 2 * pad + 0.04,
            "kinds": ["name"] + ["desc"] * len(desc), "strings": [n["name"]] + list(desc), "name_pt": name_pt}


def legend_min_h(ctx):
    s = ctx["style"]
    n = max(1, len(s["flows"]))
    return L.line_height(s["font"], s["sizes"]["legend"] + 1, True) + 0.16 + n * 0.2


def box_kind(box, depth):
    if box.get("kind") == "boundary":
        return "boundary"
    if box.get("title") or box.get("panel"):
        return "zone" if depth == 0 else "card"
    return None


def gap_for(ctx, box, depth):
    gaps = ctx["style"]["gap"]
    return box.get("gap", gaps[min(depth, len(gaps) - 1)])


def split_row(kids, w, gap):
    total = sum(k.get("weight", 1) for k in kids)
    span = w - gap * (len(kids) - 1)
    return [span * k.get("weight", 1) / total for k in kids]


def min_h(ctx, box, w, depth):
    """Smallest height that shows this box's content at width w."""
    if box.get("legend"):
        return legend_min_h(ctx)
    if "name" in box:
        return node_dims(ctx, box, w)["min_h"]
    kind, pad = box_kind(box, depth), ctx["style"]["panel"]["pad"]
    extra, iw, d2 = 0.0, w, depth
    if kind == "boundary":
        extra, iw = boundary_dims(ctx, box)[0] + pad, w - 2 * pad
    elif kind:
        hh, under, *_ = header_dims(ctx, box, w, depth)
        extra, iw, d2 = hh + under + pad, w - 2 * pad, depth + 1
    kids = box.get("children", [])
    if not kids:
        return extra + 0.2
    gap = gap_for(ctx, box, d2)
    if box.get("dir", "col") == "row":
        inner = max(min_h(ctx, k, sz, d2) for k, sz in zip(kids, split_row(kids, iw, gap)))
    else:
        inner = sum(min_h(ctx, k, iw, d2) for k in kids) + gap * (len(kids) - 1)
    return extra + inner


def split_col(ctx, kids, w, h, gap, depth):
    """Each child gets its measured minimum; leftover height is shared by weight."""
    mins = [min_h(ctx, k, w, depth) for k in kids]
    avail = h - gap * (len(kids) - 1)
    extra = avail - sum(mins)
    if extra >= 0:
        total = sum(k.get("weight", 1) for k in kids)
        return [m + extra * k.get("weight", 1) / total for m, k in zip(mins, kids)]
    ctx["squeezed"] = True  # content doesn't fit: shrink proportionally; lint's text-fit will name the boxes
    return [m * avail / sum(mins) for m in mins]


# ---------- alignment: line connected nodes up so arrows run straight ----------

def _spacer(k):
    return not (k.get("children") or "name" in k or k.get("title") or k.get("kind") or k.get("panel")
                or k.get("legend"))


def _slot_kind(k):
    """What a column child is, for alignment: a node, a plain row of nodes, a spacer, or None (anything else)."""
    if "name" in k:
        return "node"
    if _spacer(k):
        return "spacer"
    if k.get("dir") == "row" and not (k.get("title") or k.get("kind") or k.get("panel")) \
            and all("name" in c or _spacer(c) for c in k["children"]):
        return "row"
    return None


def _node_ids(k):
    return [c.get("id") or c["name"] for c in ([k] if "name" in k else k.get("children", [])) if "name" in c]


def _isotonic(t, w):
    """Weighted least-squares fit to t that never decreases (pool adjacent violators)."""
    blocks = []
    for v, wt in zip(t, w):
        blocks.append([v, wt, 1])
        while len(blocks) > 1 and blocks[-2][0] > blocks[-1][0]:
            v2, w2, n2 = blocks.pop()
            v1, w1, n1 = blocks.pop()
            blocks.append([(v1 * w1 + v2 * w2) / (w1 + w2), w1 + w2, n1 + n2])
    return [v for v, _, n in blocks for _ in range(n)]


FLOW_GAP = 0.3  # room for a visible arrow (and its step number) between two linked nodes stacked in one card


def _spread(col, anchored):
    """Children with no partner beside them: spread between the anchored ones (each gap = its minimum + an equal
    share of the free space), packed from the top; those after the last anchor keep the column's rhythm."""
    y0, y1, gp, mins, top, n = col["y"], col["y"] + col["h"], col["gp"], col["mins"], col["top"], len(col["mins"])
    idx = sorted(anchored)

    def lay(first, stop, start, extra):
        pos = start
        for i in range(first, stop):
            top[i] = pos
            pos += mins[i] + (gp[i] if i < n - 1 else 0) + extra

    if not idx:  # only vertical links in this column: pack from the top with a modest, even spacing
        lay(0, n, y0, min(0.25, (col["h"] - sum(mins) - sum(gp)) / max(1, n - 1)))
        return
    a = idx[0]
    if a:
        lay(0, a, y0, (top[a] - y0 - sum(mins[:a]) - sum(gp[:a])) / a)
    for a, b in zip(idx, idx[1:]):
        if b - a > 1:
            extra = (top[b] - top[a] - sum(mins[a:b]) - sum(gp[a:b])) / (b - a)
            lay(a + 1, b, top[a] + mins[a] + gp[a] + extra, extra)
    last = idx[-1]
    if last < n - 1:
        extras = [top[j + 1] - top[j] - mins[j] - gp[j] for j in range(last)]
        avail = y1 - top[last] - sum(mins[last:]) - sum(gp[last:])
        extra = max(0.0, min(statistics.median(extras) if extras else 0.0, avail / (n - 1 - last)))
        lay(last + 1, n, top[last] + mins[last] + gp[last] + extra, extra)


def align_offsets(ctx, flows, sweeps=4):
    """For every column of nodes (a card's node list), place each node as close as possible to the height of
    the boxes it connects to beside it, keeping order, min heights and gaps; two linked nodes stacked in the
    column get FLOW_GAP between them. Frames never move: only the free space inside each column is redistributed.
    Returns {container id: (child heights, gap before each child, unused height at the bottom)}."""
    rects, cols = ctx["rects"], ctx["cols"]
    centre = {k: r[1] + r[3] / 2 for k, r in rects.items()}
    home = {n: (ci, ki) for ci, col in enumerate(cols) for ki, ids in enumerate(col["ids"]) for n in ids}

    def beside(a, b):  # side by side, so a horizontal arrow can join them
        (ax, _, aw, _), (bx, _, bw, _) = rects[a], rects[b]
        return ax + aw <= bx or bx + bw <= ax

    edges = [(f["from"], f["to"]) for f in flows if f.get("from") in rects and f.get("to") in rects]
    linked = {(home[a][0], *sorted((home[a][1], home[b][1]))) for a, b in edges
              if a in home and b in home and home[a][0] == home[b][0]}
    for ci, col in enumerate(cols):
        g, n = col["gap"], len(col["mins"])
        links = [j for j in range(n - 1) if (ci, j, j + 1) in linked]
        room = col["h"] - sum(col["mins"]) - g * (n - 1)  # what a tight card can spare for its arrows
        extra = min(max(0.0, FLOW_GAP - g), max(0.0, room) / max(1, len(links)))
        col["gp"] = [g + extra if j in links else g for j in range(n - 1)]
    done = set()
    for _ in range(sweeps):  # Gauss-Seidel: each column sees the columns already moved in this sweep
        for ci, col in enumerate(cols):
            mins, gp, n = col["mins"], col["gp"], len(col["mins"])
            spare = col["h"] - sum(mins) - sum(gp)
            want = [[] for _ in range(n)]
            for a, b in edges:
                for u, v in ((a, b), (b, a)):
                    if home.get(u, (-1,))[0] == ci and home.get(v, (-1,))[0] != ci and beside(u, v):
                        want[home[u][1]].append(centre[v])
            if spare < -1e-9 or not (any(want) or gp != [col["gap"]] * (n - 1)):
                continue
            t, w, acc = [], [], 0.0
            for i in range(n):  # children with no partner beside them hold their place, weakly
                top = statistics.median(want[i]) - mins[i] / 2 if want[i] else col["top"][i]
                t.append(top - acc)
                w.append(len(want[i]) or 0.01)
                acc += mins[i] + (gp[i] if i < n - 1 else 0)
            acc = 0.0
            for i, z in enumerate(_isotonic(t, w)):
                col["top"][i] = min(max(z, col["y"]), col["y"] + spare) + acc
                acc += mins[i] + (gp[i] if i < n - 1 else 0)
            _spread(col, {i for i in range(n) if want[i]})
            for i in range(n):
                for nid in col["ids"][i]:
                    centre[nid] = col["top"][i] + mins[i] / 2
            done.add(ci)
    out = {}
    for ci in done:
        col = cols[ci]
        ends = [col["y"]] + [t + m + col["gap"] for t, m in zip(col["top"], col["mins"])]
        out[col["id"]] = (col["mins"], [max(0.0, t - e) for t, e in zip(col["top"], ends)],
                          max(0.0, col["y"] + col["h"] - ends[-1] + col["gap"]))
    return out


# ---------- layout + draw ----------

def record(ctx, bid, kind, rect, depth, anc, parent, **extra):
    b = {"id": bid, "kind": kind, "rect": rect, "depth": depth, "ancestors": list(anc), "parent": parent, **extra}
    ctx["boxes"].append(b)
    ctx["info"][bid] = b
    if kind in ("zone", "card", "node", "boundary", "group"):
        ctx["rects"][bid] = rect


def arrange(ctx, box, x, y, w, h, depth, anc, parent, fit=False):
    if box.get("legend"):
        ctx["legend_rect"] = (x, y, w, h)
        record(ctx, "legend", "legend", (x, y, w, h), depth, anc, parent)
        return
    if "name" in box:
        return draw_node(ctx, box, x, y, w, h, depth, anc, parent)
    bid = box.get("id") or f"_g{next(ctx['counter'])}"
    if box.get("participants"):
        ctx["participants"][bid] = box["participants"]
    kind = box_kind(box, depth)
    if fit and kind in ("card", "boundary") and bid in ctx["align"]:
        h -= ctx["align"][bid][2]  # no empty card bottom: the space stays free below it
    if kind == "boundary":
        x, y, w, h = draw_boundary(ctx, box, bid, x, y, w, h, depth, anc, parent)
        anc = anc + [bid]
    elif kind:
        x, y, w, h = draw_panel(ctx, box, bid, kind, x, y, w, h, depth, anc, parent)
        anc, depth = anc + [bid], depth + 1
    elif "id" in box:
        record(ctx, bid, "group", (x, y, w, h), depth, anc, parent)
        anc = anc + [bid]
    kids = box.get("children", [])
    if not kids:
        return
    gap = gap_for(ctx, box, depth)
    row = box.get("dir", "col") == "row"
    if row and depth == 0:
        ctx["columns"] = max(ctx["columns"], sum(1 for k in kids if k.get("title") or k.get("kind")))
    pre = [0.0] * len(kids)
    if row:
        sizes = split_row(kids, w, gap)
    elif bid in ctx["align"]:  # second pass: nodes at their min height, moved into line with their partners
        sizes, pre, _ = ctx["align"][bid]
    else:
        sizes = split_col(ctx, kids, w, h, gap, depth)
        if box.get("align", True) and all(_slot_kind(k) for k in kids):
            mins = [min_h(ctx, k, w, depth) for k in kids]
            tops = itertools.accumulate([y] + [sz + gap for sz in sizes[:-1]])
            ctx["cols"].append({"id": bid, "y": y, "h": h, "gap": gap, "mins": mins,
                                "ids": [_node_ids(k) for k in kids],
                                "top": [t + (sz - m) / 2 for t, sz, m in zip(tops, sizes, mins)]})
    pos = x if row else y
    for k, size, p in zip(kids, sizes, pre):
        pos += p
        if row:
            arrange(ctx, k, pos, y, size, h, depth, anc, bid)
        else:
            arrange(ctx, k, x, pos, w, size, depth, anc, bid, fit=True)
        pos += size + gap


def draw_panel(ctx, box, bid, kind, x, y, w, h, depth, anc, parent):
    s, c, p = ctx["style"], ctx["style"]["colors"], ctx["style"]["panel"]
    shapes, top = ctx["slide"].shapes, kind == "zone"
    hh, under, paras, bar, margin = header_dims(ctx, box, w, depth)
    fill, line = (c["panelFill"], c["panelBorder"]) if top else (c["cardFill"], c["cardBorder"])
    sh = box_shape(shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=fill, line=line,
                   width=p["border"], radius=p["radius"], name=f"{kind}:{bid}")
    if p["shadow"] and top:
        add_shadow(sh)
    record(ctx, bid, kind, (x, y, w, h), depth, anc, parent, title=box.get("title"))
    if bar:
        hb = box_shape(shapes, MSO_SHAPE.ROUND_2_SAME_RECTANGLE, x, y, w, hh, fill=c["primary"],
                       radius=p["radius"], name=f"header:{bid}")
        write(hb.text_frame, paras, s, MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, margin)
        record(ctx, f"header:{bid}", "header", (x, y, w, hh), depth, anc + [bid], bid, owner=bid)
        ctx["color_pairs"].append(("FFFFFF", c["primary"], f"header:{bid}"))
    elif paras:
        align = PP_ALIGN.LEFT if p.get("align") == "left" else PP_ALIGN.CENTER
        write(sh.text_frame, paras, s, MSO_ANCHOR.TOP, align, margin)
        ctx["color_pairs"].append((c["headerText"], fill or "FFFFFF", f"{kind}:{bid}"))
        # the title text band: arrows and labels keep off it like a header bar
        record(ctx, f"title:{bid}", "header", (x, y, w, hh), depth, anc + [bid], bid, owner=bid)
    pad = p["pad"]
    return x + pad, y + hh + under, w - 2 * pad, h - hh - pad - under


def draw_boundary(ctx, box, bid, x, y, w, h, depth, anc, parent):
    """Trust / network boundary: dashed outline with a label, no fill. Doesn't change nesting depth."""
    s, c, pad = ctx["style"], ctx["style"]["colors"], ctx["style"]["panel"]["pad"]
    lh, size = boundary_dims(ctx, box)
    sh = box_shape(ctx["slide"].shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, line=c["boundary"], width=1,
                   dash=MSO_LINE_DASH_STYLE.DASH, radius=ctx["style"]["panel"]["radius"], name=f"boundary:{bid}")
    record(ctx, bid, "boundary", (x, y, w, h), depth, anc, parent, title=box.get("title"))
    if box.get("title"):
        write(sh.text_frame, [(box["title"], size, c["boundary"], True)], s, MSO_ANCHOR.TOP, PP_ALIGN.LEFT, 0.06)
        tw = L.text_width(box["title"], s["font"], size, True) + 0.16
        record(ctx, f"title:{bid}", "header", (x, y, min(w, tw), lh), depth, anc + [bid], bid, owner=bid)
    return x + pad, y + lh, w - 2 * pad, h - lh - pad


def draw_node(ctx, n, x, y, w, h, depth, anc, parent):
    s, c = ctx["style"], ctx["style"]["colors"]
    d = node_dims(ctx, n, w)
    nid = n.get("id") or n["name"]
    if n.get("participants"):
        ctx["participants"][nid] = n["participants"]
    record(ctx, nid, "node", (x, y, w, h), depth, anc, parent)
    if n.get("boxed"):
        p = s["panel"]
        box_shape(ctx["slide"].shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=c["cardFill"],
                  line=c["cardBorder"], width=p["border"], radius=p["radius"], name=f"card:{nid}")
    pad, isz, left = d["pad"], d["isz"], d["orient"] == "left"
    ix0, iy0, iw, ih = x + pad, y + pad, w - 2 * pad, h - 2 * pad
    y0 = iy0 + max(0.0, (ih - d["content"]) / 2)  # centre the node in its slot
    grp = ctx["slide"].shapes.add_group_shape()
    grp.name = f"node:{nid}"
    icon = n.get("icon")
    if icon:
        ix, iy = (ix0, y0 + (d["content"] - isz) / 2) if left else (ix0 + (iw - isz) / 2, y0)
        add_icon(ctx, grp.shapes, n, ix, iy, isz, nid, depth, d)
    if left:
        tx, ty = ix0 + (isz + 0.06 if icon else 0), y0 + (d["content"] - d["th"]) / 2
    else:
        tx, ty = ix0, y0 + (isz + 0.04 if icon else 0)
    tb = grp.shapes.add_textbox(E(tx), E(ty), E(d["tw"]), E(d["th"]))
    tb.name = f"text:{nid}"
    write(tb.text_frame, d["paras"], s, align=PP_ALIGN.LEFT if left else PP_ALIGN.CENTER)
    ctx["texts"].append({"id": nid, "need": d["content"] + 2 * pad, "have": h, "lines": d["counts"],
                         "kinds": d["kinds"], "strings": d["strings"], "too_wide": d["too_wide"]})
    ctx["color_pairs"].append((c["muted"], c["cardFill"] or "FFFFFF", f"node:{nid}"))


def find_icon(ctx, icon):
    for d in ctx["icon_dirs"]:
        p = d / f"{icon}.png"
        if p.is_file():
            return p
    return None


def add_icon(ctx, shapes, n, x, y, size, nid, depth, d):
    icon = n["icon"]
    path = find_icon(ctx, icon)
    ctx["icons"].append({"id": nid, "icon": icon, "size": size, "name_pt": d["name_pt"], "orientation": d["orient"],
                         "depth": depth, "missing": path is None, "stand_in": bool(n.get("iconStandIn"))})
    if not path:
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


# ---------- flows ----------

def lift(ctx, bid, level):
    """Move a node id up to the card or zone that contains it."""
    b = ctx["info"].get(bid)
    if not b or level == "node":
        return bid
    def fits(x):  # untitled inner panels aren't meaningful arrow ends; skip up to a titled card or the zone
        i = ctx["info"].get(x, {})
        return i.get("kind") == "zone" or (level == "card" and i.get("kind") == "card" and i.get("title"))

    if fits(bid):
        return bid
    for a in reversed(b["ancestors"]):
        if fits(a):
            return a
    return bid


def steps_of(f):
    if f.get("steps"):
        return list(f["steps"])
    if f.get("step") is not None:
        return [f["step"]]
    m = re.match(r"^\s*(\d+(?:\s*[·,]\s*\d+)*)", str(f.get("label") or ""))
    return [int(v) for v in re.findall(r"\d+", m.group(1))] if m else []


def resolve_flows(ctx, spec, spec_dir):
    seqd, skipped, flows = None, [], spec.get("flows", [])
    if spec.get("sequence"):
        seqd = S.parse((spec_dir / spec["sequence"]).read_text(encoding="utf-8"))
    if flows == "from-sequence":
        if not seqd:
            ctx["spec_errors"].append('"flows": "from-sequence" needs "sequence": "<file.mmd>"')
            return [], None, []
        owner = {p: bid for bid, parts in ctx["participants"].items() for p in parts}
        opts = spec.get("flowOptions", {})
        flows, skipped = S.derive_flows(seqd, owner, lambda b: lift(ctx, b, opts.get("level", "node")), opts)
    return flows, seqd, skipped


def route_flows(ctx, flows, W, H):
    s, m = ctx["style"], ctx["style"]["content"]
    # arrows stay in the content area: above the zones they cut through the title, below it into the footer
    bounds = (max(0, m["left"] - 0.25), ctx["top"],
              min(W, W - m["right"] + 0.25) - max(0, m["left"] - 0.25), H - m["bottom"] - ctx["top"])
    # a column (zone) body may be crossed between its cards, at a small cost; its header bar never
    blockers = [b for b in ctx["boxes"] if b["kind"] in ("node", "card", "legend", "header", "boundary")]
    zones = [b for b in ctx["boxes"] if b["kind"] == "zone"]
    used, routes = {}, []
    valid = [f for f in flows if f.get("from") in ctx["rects"] and f.get("to") in ctx["rects"]
             and f.get("type") in s["flows"]]

    def dist(f):
        a, b = ctx["rects"][f["from"]], ctx["rects"][f["to"]]
        return abs(a[0] + a[2] / 2 - b[0] - b[2] / 2) + abs(a[1] + a[3] / 2 - b[1] - b[3] / 2)

    for f in sorted(valid, key=dist):  # short, simple routes first; long ones go around them
        a, b = f["from"], f["to"]
        # an arrow may touch its endpoints, their containers and their contents; everything else is an obstacle
        exempt = {a, b, *ctx["info"][a]["ancestors"], *ctx["info"][b]["ancestors"]}
        exempt |= {o["id"] for o in blockers if (a in o["ancestors"] or b in o["ancestors"]) and o["kind"] != "header"}
        # header bars are never crossed, not even the endpoint's own: arrows over header text are unreadable
        obs = [(r[0] - 0.02, r[1] - 0.02, r[2] + 0.04, r[3] + 0.04) for r in
               (o["rect"] for o in blockers if o["kind"] == "header" or o["id"] not in exempt)]
        # the endpoints' own cards may be crossed, but an arrow riding along their border looks like part of it
        soft = [ctx["rects"][i] for i in {*ctx["info"][a]["ancestors"], *ctx["info"][b]["ancestors"]}
                if ctx["info"].get(i, {}).get("kind") in ("zone", "card", "boundary")]
        cross = [z["rect"] for z in zones if z["id"] not in exempt]
        pts = L.route(ctx["rects"][a], ctx["rects"][b], obs, used, bounds, soft=soft, cross=cross)
        routed = pts is not None
        if not routed:
            ra, rb = ctx["rects"][a], ctx["rects"][b]
            pts = [(ra[0] + ra[2] / 2, ra[1] + ra[3] / 2), (rb[0] + rb[2] / 2, rb[1] + rb[3] / 2)]
        routes.append({"id": f"{a}->{b}", "flow": f, "pts": pts, "routed": routed, "exempt": exempt,
                       "steps": steps_of(f)})
    return routes


def draw_routes(ctx, routes):
    s, c = ctx["style"], ctx["style"]["colors"]
    shapes = ctx["slide"].shapes
    # labels may sit in card padding, never on node text, header bars, the legend or other labels
    avoid = [b["rect"] for b in ctx["boxes"] if b["kind"] in ("node", "legend", "header")]
    t = 0.03  # nor on a dashed boundary line
    for x, y, w, h in (b["rect"] for b in ctx["boxes"] if b["kind"] == "boundary"):
        avoid += [(x, y - t, w, 2 * t), (x, y + h - t, w, 2 * t), (x - t, y, 2 * t, h), (x + w - t, y, 2 * t, h)]
    # nor across another arrow's line
    lines = {r["id"]: [(min(p[0], q[0]) - 0.015, min(p[1], q[1]) - 0.015, abs(p[0] - q[0]) + 0.03,
                        abs(p[1] - q[1]) + 0.03) for p, q in L.segments(r["pts"])] for r in routes}
    for r in routes:
        others = [seg for k, v in lines.items() if k != r["id"] for seg in v]
        f = r["flow"]
        flow = s["flows"][f["type"]]
        ctx["used"].add(f["type"])
        arrow(shapes, r["pts"], flow, f.get("both", False), True, f"flow:{r['id']}")
        label = f.get("label")
        if label and f.get("step") is not None and not str(label).lstrip().startswith(str(f["step"])):
            label = f"{f['step']}. {label}"
        note = f.get("note")
        if label or note:
            name_pt, desc_pt = fs(s, s["sizes"]["name"]), fs(s, s["sizes"]["desc"])
            paras = ([(label, name_pt, flow["color"], True)] if label else []) + \
                    ([(note, desc_pt, c["muted"], False)] if note else [])
            w = max([L.text_width(label, s["font"], name_pt, True)] if label else [0]) + 0.08
            if note:
                w = max(w, min(1.5, L.text_width(note, s["font"], desc_pt) + 0.08))
            # a label beside a short horizontal run (a gutter) wraps to the run's length instead of spilling over
            runs = [abs(p[0] - q[0]) for p, q in L.segments(r["pts"]) if abs(p[1] - q[1]) < 1e-9]
            if runs and len(runs) == len(L.segments(r["pts"])) and max(runs) < w:
                w = max(max(runs) - 0.04, 0.5)
            h = L.measure(paras, s["font"], w - 0.04)[0] + 0.06
            rect, hit = L.place_label(r["pts"], w, h, avoid, others)
            tb = shapes.add_textbox(E(rect[0]), E(rect[1]), E(rect[2]), E(rect[3]))
            tb.name = f"label:{r['id']}"
            write(tb.text_frame, paras, s, MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, 0.02)
        elif r["steps"]:
            rect, hit = L.place_label(r["pts"], 0.24, 0.24, avoid, others)
            b = box_shape(shapes, MSO_SHAPE.OVAL, *rect, fill=flow["color"], name=f"step:{r['id']}")
            write(b.text_frame, [(str(r["steps"][0]), s["sizes"]["minFontPt"], "FFFFFF", True)], s,
                  MSO_ANCHOR.MIDDLE, PP_ALIGN.CENTER, 0)
        else:
            continue
        r["label_rect"], r["label_collided"] = rect, hit
        avoid.append(rect)
        ctx["boxes"].append({"id": f"label:{r['id']}", "kind": "label", "rect": rect, "depth": 0,
                             "ancestors": [], "parent": "_labels"})


def draw_legend(ctx):
    s, c = ctx["style"], ctx["style"]["colors"]
    used = [t for t in s["flows"] if t in ctx["used"]]
    if "legend_rect" not in ctx or not used:
        return
    x, y, w, h = ctx["legend_rect"]
    sz = fs(s, s["sizes"]["legend"])
    grp = ctx["slide"].shapes.add_group_shape()
    grp.name = "legend"
    bg = box_shape(grp.shapes, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill=c["panelFill"] or "FFFFFF",
                   line=c["panelBorder"] or c["cardBorder"], radius=s["panel"]["radius"], name="legend:box")
    write(bg.text_frame, [(s["legend"]["title"], sz + 1, c["headerText"], True)], s, margin=0.06)
    head = L.line_height(s["font"], sz + 1, True) + 0.14
    row = min(0.24, (h - head - 0.04) / len(used))
    for i, t in enumerate(used):
        ly = y + head + i * row + row / 2
        arrow(grp.shapes, [(x + 0.1, ly), (x + 0.45, ly)], s["flows"][t], name=f"legend:{t}")
        tb = grp.shapes.add_textbox(E(x + 0.52), E(ly - row / 2), E(w - 0.58), E(row))
        write(tb.text_frame, [(s["flows"][t]["label"], sz, c["text"], False)], s, MSO_ANCHOR.MIDDLE, margin=0.01)
    ctx["texts"].append({"id": "legend", "need": head + len(used) * L.line_height(s["font"], sz) + 0.06, "have": h,
                         "lines": [], "kinds": [], "strings": [], "too_wide": False})


# ---------- slide chrome ----------

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
    """Returns (left, bottom, width) of the title, in inches."""
    slide, s = ctx["slide"], ctx["style"]
    t = slide.shapes.title
    if t is not None:  # the template's title placeholder: keep its place, use the style's type
        write(t.text_frame, [(text, s["sizes"]["title"], s["colors"]["text"], True)], s, MSO_ANCHOR.MIDDLE)
        return t.left / EMU, (t.top + t.height) / EMU, t.width / EMU
    tb = slide.shapes.add_textbox(E(0.3), E(0.22), E(W - 2.8), E(0.62))
    tb.name = "title"
    write(tb.text_frame, [(text, s["sizes"]["title"], s["colors"]["text"], True)], s, MSO_ANCHOR.MIDDLE)
    return 0.3, 0.84, W - 2.8


def draw_subtitle(ctx, text, left, top, width):
    s = ctx["style"]
    size = s["sizes"]["subtitle"]
    h = L.line_height(s["font"], size) + 0.06
    tb = ctx["slide"].shapes.add_textbox(E(left + 0.07), E(top), E(width), E(h))
    tb.name = "subtitle"
    write(tb.text_frame, [(text, size, s["colors"]["subtitle"], False)], s, margin=0.03)
    ctx["color_pairs"].append((s["colors"]["subtitle"], "FFFFFF", "subtitle"))
    return top + h


# ---------- build ----------

def check_spec(spec, style, errors):
    ids = [b.get("id") or b["name"] for b in _walk(spec["root"]) if "id" in b or "name" in b]
    dup = sorted({i for i in ids if ids.count(i) > 1})
    if dup:
        errors.append(f"duplicate ids: {dup}")
    if isinstance(spec.get("flows"), list):
        for f in spec["flows"]:
            for end in ("from", "to"):
                if f.get(end) not in ids:
                    errors.append(f"flow {f.get('from')}->{f.get('to')}: unknown {end} id {f.get(end)!r}")
            if f.get("type") not in style["flows"]:
                errors.append(f"flow {f.get('from')}->{f.get('to')}: type {f.get('type')!r} not in style flows "
                              f"{list(style['flows'])}")


def _walk(box):
    yield box
    for k in box.get("children", []):
        yield from _walk(k)


def new_ctx(slide, style, icon_dirs):
    return {"slide": slide, "style": style, "rects": {}, "info": {}, "boxes": [], "texts": [], "icons": [],
            "color_pairs": [], "participants": {}, "used": set(), "icon_dirs": icon_dirs, "columns": 0,
            "counter": itertools.count(1), "spec_errors": [], "squeezed": False, "cols": [], "align": {}}


def build(spec_paths, style_ref, out):
    """Builds the deck and writes lint.json next to each spec. Returns the report."""
    s = load_style(style_ref)
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
    icon_dirs = [s["_dir"] / "icons", ICON_HOME, HERE / "assets" / "icons"]
    report = {"out": str(Path(out).resolve()), "metrics": "exact" if L.exact_metrics(s["font"]) else
              f"estimated ({s['font']} not installed)", "slides": []}
    lint_files = {}
    scratch = Presentation()  # pass 1 draws here, only to measure
    for spec_path in spec_paths:
        spec_path = Path(spec_path)
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        slide = prs.slides.add_slide(layout)
        for ph in list(slide.placeholders):  # templates not prepped by `template` may carry content placeholders
            if ph.placeholder_format.type not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                ph._element.getparent().remove(ph._element)
        ctx = new_ctx(slide, s, icon_dirs)
        check_spec(spec, s, ctx["spec_errors"])
        if not tpl:
            blank_chrome(ctx, W, H)
        tl, tb, tw = set_title(ctx, spec.get("title", ""), W)
        m = s["content"]
        top = m["top"]
        if spec.get("subtitle"):
            top = max(top, draw_subtitle(ctx, spec["subtitle"], tl, tb - 0.06, tw) + 0.08)
        ctx["top"] = top
        content = (m["left"], top, W - m["left"] - m["right"], H - top - m["bottom"])
        # pass 1 (scratch slide): measure the layout and find the flows; then line connected nodes up
        probe = new_ctx(scratch.slides.add_slide(scratch.slide_layouts[6]), s, icon_dirs)
        arrange(probe, spec["root"], *content, 0, [], "_root")
        ctx["align"] = align_offsets(probe, resolve_flows(probe, spec, spec_path.parent)[0])
        arrange(ctx, spec["root"], *content, 0, [], "_root")
        flows, seqd, skipped = resolve_flows(ctx, spec, spec_path.parent)
        routes = route_flows(ctx, flows, W, H)
        draw_routes(ctx, routes)
        draw_legend(ctx)
        model = {"style": s, "boxes": ctx["boxes"], "texts": ctx["texts"], "icons": ctx["icons"], "routes": routes,
                 "content": content, "columns": ctx["columns"], "color_pairs": ctx["color_pairs"], "sequence": seqd,
                 "skipped": skipped, "has_legend": "legend_rect" in ctx, "flow_types": ctx["used"]}
        result = lint.run(model, s, spec)
        for e in ctx["spec_errors"]:
            result["issues"].insert(0, {"rule": "spec", "severity": "error", "id": "spec", "message": e,
                                        "fix": "fix the spec"})
            result["errors"] += 1
            result["ok"] = False
        slide_rep = {"spec": str(spec_path), "flows": len(routes), "lint": result,
                     "icons_used": {i["id"]: i["icon"] for i in ctx["icons"]}}
        report["slides"].append(slide_rep)
        lint_files.setdefault(spec_path.parent, []).append(slide_rep)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    try:
        prs.save(out)
    except PermissionError:  # Windows locks a deck that is open in PowerPoint
        print(json.dumps({"error": f"cannot write {Path(out).resolve()}: it is open in PowerPoint (or read-only). "
                                   "Close it and build again, or build to another path."}))
        sys.exit(2)
    for folder, slides in lint_files.items():  # lint.json next to the spec: the Stop-hook quality gate reads it
        (folder / "lint.json").write_text(json.dumps(
            {"out": report["out"], "errors": sum(x["lint"]["errors"] for x in slides),
             "slides": slides}, indent=1, ensure_ascii=False), encoding="utf-8")
    report["ok"] = all(x["lint"]["ok"] for x in report["slides"])
    report["lint_files"] = [str((f / "lint.json").resolve()) for f in lint_files]
    return report
