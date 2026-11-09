"""Measurement and routing for arch-slide.

- Text is measured with the real font file (PIL) when it is installed, wrapped the way PowerPoint wraps it.
  Without the font, widths fall back to an estimate and the report says so.
- Arrows are routed orthogonally around boxes with A* on a grid (bends and shared lanes cost extra).
All units are inches.
"""
import heapq
import math
import os
import shutil
import subprocess
from functools import lru_cache
from pathlib import Path

PX_PER_PT = 20  # fonts are loaded at 20 px per point; inches = px / PX_PER_PT / 72


# ---------- fonts and text ----------

@lru_cache(None)
def _font_index():
    dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts",
            Path(os.environ.get("LOCALAPPDATA", "~")).expanduser() / "Microsoft" / "Windows" / "Fonts",
            Path.home() / "Library" / "Fonts", Path("/Library/Fonts"), Path("/System/Library/Fonts"),
            Path("/usr/share/fonts"), Path("/usr/local/share/fonts"), Path.home() / ".fonts",
            Path.home() / ".local" / "share" / "fonts"]
    idx = {}
    for d in dirs:
        if d.is_dir():
            for p in d.rglob("*"):
                if p.suffix.lower() in (".ttf", ".otf", ".ttc"):
                    idx.setdefault(p.name.lower(), p)
    return idx


@lru_cache(None)
def font_file(family, bold=False):
    """Path of the font file for a family, or None when it isn't installed."""
    key = family.lower().replace(" ", "")
    names = ([f"{key}b", f"{key}bd", f"{key}-bold", f"{key}bold", f"{key}-semibold"] if bold else []) + \
            [key, f"{key}-regular", f"{key}regular"]
    idx = _font_index()
    for n in names:
        for ext in (".ttf", ".otf", ".ttc"):
            if n + ext in idx:
                return str(idx[n + ext])
    if shutil.which("fc-match"):
        r = subprocess.run(["fc-match", "-f", "%{file}", f"{family}:{'bold' if bold else 'regular'}"],
                           capture_output=True, text=True)
        if r.returncode == 0 and key[:5] in Path(r.stdout).name.lower().replace(" ", ""):
            return r.stdout
    return None


@lru_cache(None)
def _pil(family, bold, size):
    f = font_file(family, bold) or (font_file(family, False) if bold else None)
    if not f:
        return None
    from PIL import ImageFont
    return ImageFont.truetype(f, size=max(1, round(size * PX_PER_PT)))


def exact_metrics(family):
    return font_file(family) is not None


def text_width(text, family, size, bold=False):
    f = _pil(family, bold, size)
    if f is None:  # ponytail: estimate when the font isn't installed; the report flags it
        return len(text) * size * (0.56 if bold else 0.52) / 72
    return f.getlength(text) / PX_PER_PT / 72


def line_height(family, size, bold=False):
    """PowerPoint single spacing: the font's own line height, at least 1.2 x size."""
    f = _pil(family, bold, size)
    ratio = 1.2
    if f is not None:
        a, d = f.getmetrics()
        ratio = max(ratio, (a + d) / (size * PX_PER_PT))
    return size * ratio / 72


def wrap(text, family, size, width, bold=False):
    """Greedy word wrap, like PowerPoint."""
    lines, cur = [], ""
    for w in text.split():
        trial = f"{cur} {w}" if cur else w
        if not cur or text_width(trial, family, size, bold) <= width:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    lines.append(cur)
    return lines


def measure(paras, family, width):
    """paras [(text, pt, color, bold)] in a box `width` wide (inches, inside margins).
    Returns (height, line counts per paragraph, a word wider than the box?)."""
    h, counts, too_wide = 0.0, [], False
    for t, size, _, bold in paras:
        ls = wrap(t, family, size, width, bold)
        counts.append(len(ls))
        h += len(ls) * line_height(family, size, bold)
        too_wide = too_wide or any(text_width(w, family, size, bold) > width + 1e-6 for w in t.split())
    return h, counts, too_wide


# ---------- geometry ----------

def overlap(a, b, tol=0.005):
    return a[0] + tol < b[0] + b[2] and b[0] + tol < a[0] + a[2] and \
        a[1] + tol < b[1] + b[3] and b[1] + tol < a[1] + a[3]


def inside(inner, outer, tol=0.005):
    return inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol and \
        inner[0] + inner[2] <= outer[0] + outer[2] + tol and inner[1] + inner[3] <= outer[1] + outer[3] + tol


def seg_hits_rect(p, q, r, tol=0.01):
    """Axis-aligned segment p-q passes through the interior of rect r."""
    x0, x1 = sorted((p[0], q[0]))
    y0, y1 = sorted((p[1], q[1]))
    return x0 < r[0] + r[2] - tol and x1 > r[0] + tol and y0 < r[1] + r[3] - tol and y1 > r[1] + tol


def segs_cross(a, b):
    """Proper crossing of two axis-aligned segments (one horizontal, one vertical)."""
    (p, q), (r, s) = a, b
    ah, bh = abs(p[1] - q[1]) < 1e-9, abs(r[1] - s[1]) < 1e-9
    if ah == bh:
        return False
    (h0, h1), (v0, v1) = ((p, q), (r, s)) if ah else ((r, s), (p, q))
    hx0, hx1 = sorted((h0[0], h1[0]))
    vy0, vy1 = sorted((v0[1], v1[1]))
    return hx0 < v0[0] < hx1 and vy0 < h0[1] < vy1


def segments(pts):
    return list(zip(pts, pts[1:]))


# ---------- routing ----------

DIRS = [(0, -1), (1, 0), (0, 1), (-1, 0)]  # up, right, down, left


def _ports(r, other):
    """One port per side. Where the two boxes share a span (x for top/bottom, y for left/right), both boxes
    use the middle of that span, so the ports line up and a straight arrow is possible."""
    x, y, w, h = r
    ox, oy, ow, oh = other
    m = min(0.12, w / 4, h / 4)

    def aim(lo, hi, olo, ohi):
        a, b = max(lo, olo), min(hi, ohi)
        target = (a + b) / 2 if b - a > 2 * m else (olo + ohi) / 2
        return min(max(target, lo + m), hi - m)

    px, py = aim(x, x + w, ox, ox + ow), aim(y, y + h, oy, oy + oh)
    return [((px, y), 0), ((x + w, py), 1), ((px, y + h), 2), ((x, py), 3)]


def route(src, dst, obstacles, used, bounds, cell=0.04, turn=10, crowd=6, hug=2, soft=()):
    """Orthogonal path from an edge of `src` to an edge of `dst` avoiding `obstacles` (rects).
    `used` counts grid cells taken by earlier routes; sharing them costs extra, so routes spread out.
    `soft` rects may be crossed, but running along their border costs extra.
    Returns a list of points, or None when no path exists."""
    bx, by, bw, bh = bounds
    nx, ny = int(bw / cell), int(bh / cell)

    def ci(x):
        return int((x - bx) / cell)

    def cj(y):
        return int((y - by) / cell)

    def cx(i):
        return bx + (i + 0.5) * cell

    def cy(j):
        return by + (j + 0.5) * cell

    # ports already in line and nothing between them: a straight arrow, no search needed
    for (p, d), (q, e) in ((a, b) for a in _ports(src, dst) for b in _ports(dst, src)):
        k = 0 if d in (0, 2) else 1  # the coordinate both ports must share
        ahead = (q[0] - p[0]) * DIRS[d][0] + (q[1] - p[1]) * DIRS[d][1]
        if e != (d + 2) % 4 or abs(p[k] - q[k]) > 1e-6 or ahead <= 0 or                 any(seg_hits_rect(p, q, o) for o in obstacles):
            continue
        if k:  # horizontal
            cells = [(i, cj(p[1])) for i in range(ci(min(p[0], q[0])), ci(max(p[0], q[0])) + 1)]
        else:
            cells = [(ci(p[0]), j) for j in range(cj(min(p[1], q[1])), cj(max(p[1], q[1])) + 1)]
        if any(c in used for c in cells):
            continue  # an earlier arrow runs here: let the search find a lane of its own
        for c in cells:
            used[c] = 1
        return [p, q]

    blocked = set()
    for (x, y, w, h) in obstacles:  # a cell is blocked when its centre lies inside the obstacle
        i0, i1 = max(0, math.ceil((x - bx) / cell - 0.5)), min(nx - 1, math.floor((x + w - bx) / cell - 0.5))
        j0, j1 = max(0, math.ceil((y - by) / cell - 0.5)), min(ny - 1, math.floor((y + h - by) / cell - 0.5))
        for i in range(i0, i1 + 1):
            for j in range(j0, j1 + 1):
                blocked.add((i, j))
    # cells hugging an obstacle cost extra, so arrows prefer open gutters to slivers between boxes
    near = {(i + dx, j + dy) for (i, j) in blocked for dx, dy in DIRS} - blocked
    for (x, y, w, h) in soft:  # a 3-cell band along each border
        i0, i1, j0, j1 = ci(x), ci(x + w), cj(y), cj(y + h)
        for k in (-1, 0, 1):
            near |= {(i, j + k) for i in range(i0, i1 + 1) for j in (j0, j1)}
            near |= {(i + k, j) for j in range(j0, j1 + 1) for i in (i0, i1)}

    def port_cell(pt, d):
        # a full cell clear of the edge, so an obstacle's safety margin never swallows the port
        i, j = ci(pt[0] + DIRS[d][0] * cell * 1.1), cj(pt[1] + DIRS[d][1] * cell * 1.1)
        ok = 0 <= i < nx and 0 <= j < ny and (i, j) not in blocked
        edge = (cx(i), pt[1]) if d in (0, 2) else (pt[0], cy(j))  # snap the edge point onto the grid line
        return ok, (i, j), edge

    goals = {}
    for pt, d in _ports(dst, src):
        ok, c, edge = port_cell(pt, d)
        if ok:
            goals[c] = edge
    if not goals:
        return None

    def hfun(i, j):
        return min(abs(i - gi) + abs(j - gj) for gi, gj in goals)

    heap, best, prev, start_edge = [], {}, {}, {}
    for pt, d in _ports(src, dst):
        ok, (i, j), edge = port_cell(pt, d)
        if ok:
            s = (i, j, d)
            best[s], prev[s], start_edge[s] = 0, None, edge
            heapq.heappush(heap, (hfun(i, j), 0, s))
    while heap:
        _, g, s = heapq.heappop(heap)
        if g > best.get(s, 1e18):
            continue
        i, j, d = s
        if (i, j) in goals:
            cells, st = [], s
            while st is not None:
                cells.append(st)
                first = st
                st = prev[st]
            cells.reverse()
            for c in cells:
                used[(c[0], c[1])] = used.get((c[0], c[1]), 0) + 1
            pts = [start_edge[first]] + [(cx(c[0]), cy(c[1])) for c in cells] + [goals[(i, j)]]
            return simplify(pts)
        for nd, (dx, dy) in enumerate(DIRS):
            if nd == (d + 2) % 4:
                continue  # no U-turns
            ni, nj = i + dx, j + dy
            if not (0 <= ni < nx and 0 <= nj < ny) or (ni, nj) in blocked:
                continue
            ng = g + 1 + (turn if nd != d else 0) + crowd * used.get((ni, nj), 0) + (hug if (ni, nj) in near else 0)
            ns = (ni, nj, nd)
            if ng < best.get(ns, 1e18):
                best[ns], prev[ns] = ng, s
                heapq.heappush(heap, (ng + hfun(ni, nj), ng, ns))
    return None


def simplify(pts):
    """Drop repeated and collinear points."""
    out = []
    for p in (tuple(round(v, 6) for v in q) for q in pts):
        if out and p == out[-1]:
            continue
        if len(out) >= 2 and ((out[-2][0] == out[-1][0] == p[0]) or (out[-2][1] == out[-1][1] == p[1])):
            out[-1] = p
        else:
            out.append(p)
    return out


def _area(a, b):
    return max(0.0, min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])) *         max(0.0, min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]))


def place_label(pts, w, h, avoid, lines=()):
    """Put a w x h label beside a segment, longest first, at its middle, then nearer either end, clear of the
    `avoid` rects and of other arrows' `lines`. Returns (rect, collided); when every spot collides, the one with
    the least overlap, where a line through the label counts as much as covering a third of a small box."""
    best = None
    for p, q in sorted(segments(pts), key=lambda s: -(abs(s[0][0] - s[1][0]) + abs(s[0][1] - s[1][1]))):
        for f in (0.5, 0.3, 0.7):
            mx, my = p[0] + (q[0] - p[0]) * f, p[1] + (q[1] - p[1]) * f
            if abs(p[1] - q[1]) < 1e-9:  # horizontal: above, then below
                cands = [(mx - w / 2, my - h - 0.03, w, h), (mx - w / 2, my + 0.03, w, h)]
            else:                        # vertical: right, then left
                cands = [(mx + 0.05, my - h / 2, w, h), (mx - w - 0.05, my - h / 2, w, h)]
            for r in cands:
                if not any(overlap(r, o) for o in (*avoid, *lines)):
                    return r, False
                hit = sum(_area(r, o) for o in avoid) + 0.1 * sum(overlap(r, o) for o in lines)
                if best is None or hit < best[0]:
                    best = (hit, r)
    return best[1], True
