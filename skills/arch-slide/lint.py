"""Rule engine for arch-slide.

Rule TYPES live here; rule VALUES come from the org (style.json "rules", learned from its guidelines,
screenshots or text) and the project (spec.json "rules"). Resolution: defaults <- style <- spec.
Each rule: {"severity": "error" | "warn" | "off", ...params}.
"""
import re
from itertools import combinations

from layout import inside, overlap, seg_hits_rect, segments, segs_cross

DEFAULT_RULES = {
    "overlap":           {"severity": "error"},
    "text-fit":          {"severity": "error"},
    "min-font":          {"severity": "error", "minPt": None},             # None -> style sizes.minFontPt
    "name-lines":        {"severity": "warn", "max": 2},
    "desc-lines":        {"severity": "warn", "max": 2},
    "name-chars":        {"severity": "warn", "max": 32},
    "desc-chars":        {"severity": "warn", "max": 70},
    "icon-ratio":        {"severity": "warn", "min": 2.4, "max": 5.0},     # icon height / name font size
    "icon-consistency":  {"severity": "warn"},                             # same level + orientation = same size
    "missing-icon":      {"severity": "error"},
    "generic-icon":      {"severity": "warn"},                             # a generic icon stands in for a product
    "unrouted-flow":     {"severity": "error"},
    "arrow-through-box": {"severity": "error"},
    "arrow-crossings":   {"severity": "warn", "max": 2},
    "arrow-bends":       {"severity": "warn", "max": 3},
    "short-arrow":       {"severity": "warn", "min": 0.25},                # inches; shorter reads as a lone arrowhead
    "label-collision":   {"severity": "warn"},
    "safe-area":         {"severity": "error"},
    "density":           {"severity": "warn", "maxNodes": 40, "maxColumns": 6},
    "nodes-per-card":    {"severity": "warn", "max": 6},
    "contrast":          {"severity": "warn", "min": 4.5},
    "traceability":      {"severity": "error"},
    "unused-messages":   {"severity": "warn"},
    "legend":            {"severity": "off"},                              # orgs that require a legend turn it on
    "title-case":        {"severity": "off", "case": "upper"},             # e.g. "column headers in UPPERCASE"
    "boundary-label":    {"severity": "warn"},
}


def resolve(style, spec):
    rules = {}
    for k, v in DEFAULT_RULES.items():
        rules[k] = {**v, **style.get("rules", {}).get(k, {}), **spec.get("rules", {}).get(k, {})}
    unknown = sorted((set(style.get("rules", {})) | set(spec.get("rules", {}))) - set(DEFAULT_RULES))
    return rules, unknown


def _lum(hex_):
    c = [int(hex_.lstrip("#")[i:i + 2], 16) / 255 for i in (0, 2, 4)]
    c = [v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4 for v in c]
    return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]


def contrast(fg, bg):
    a, b = sorted((_lum(fg), _lum(bg)), reverse=True)
    return (a + 0.05) / (b + 0.05)


# Each check yields (target id, message, how to fix).

def r_overlap(m, c):
    kids = {}
    for b in m["boxes"]:
        if b["kind"] in ("zone", "card", "node", "legend", "boundary"):
            kids.setdefault(b["parent"], []).append(b)
    for sib in kids.values():
        for a, b in combinations(sib, 2):
            if overlap(a["rect"], b["rect"]):
                yield a["id"], f"overlaps {b['id']}", "change weights or move one box"


def r_text_fit(m, c):
    for t in m["texts"]:
        if t["need"] > t["have"] + 0.01:
            yield t["id"], f"text needs {t['need']:.2f} in, box has {t['have']:.2f} in", \
                "shorten the text, give the box more weight, or split the slide"
        if t["too_wide"]:
            yield t["id"], "a word is wider than its box", "shorten or hyphenate the long word"


def r_min_font(m, c):
    floor = c["minPt"] or m["style"]["sizes"]["minFontPt"]
    for k, v in m["style"]["sizes"].items():
        if k != "minFontPt" and v < floor:
            yield f"sizes.{k}", f"{v} pt is below the {floor} pt minimum", f"raise sizes.{k} to {floor}"


def _lines(m, c, kind):
    for t in m["texts"]:
        for k, n in zip(t["kinds"], t["lines"]):
            if k == kind and n > c["max"]:
                yield t["id"], f"{kind} wraps to {n} lines (max {c['max']})", "shorten it or widen the box"


def r_name_lines(m, c):
    yield from _lines(m, c, "name")


def r_desc_lines(m, c):
    yield from _lines(m, c, "desc")


def _chars(m, c, kind):
    for t in m["texts"]:
        for k, s in zip(t["kinds"], t["strings"]):
            if k == kind and len(s) > c["max"]:
                yield t["id"], f"{kind} is {len(s)} characters (max {c['max']})", "shorten it"


def r_name_chars(m, c):
    yield from _chars(m, c, "name")


def r_desc_chars(m, c):
    yield from _chars(m, c, "desc")


def r_icon_ratio(m, c):
    for i in m["icons"]:
        ratio = i["size"] * 72 / i["name_pt"]
        if not c["min"] <= ratio <= c["max"]:
            yield i["id"], f"icon is {ratio:.1f}x the name font (allowed {c['min']}-{c['max']})", \
                "adjust node.iconLeft / node.iconTop in the style"


def r_icon_consistency(m, c):
    groups = {}
    for i in m["icons"]:
        groups.setdefault((i["depth"], i["orientation"]), set()).add(round(i["size"], 3))
    for (depth, orient), sizes in groups.items():
        if len(sizes) > 1:
            yield f"depth {depth}", f"{orient} icons have {len(sizes)} different sizes", "give the nodes more room"


def r_missing_icon(m, c):
    for i in m["icons"]:
        if i["missing"]:
            yield i["id"], f"no icon file for {i['icon']!r} (placeholder drawn)", \
                "run `build.py fetch ...` or Glob ~/.arch-slide/icons for the right id"


def r_generic_icon(m, c):
    for i in m["icons"]:
        if i.get("stand_in"):
            yield i["id"], f"generic icon {i['icon']!r} stands in for a product", "tell the user; drop the official PNG in"


def r_unrouted(m, c):
    for r in m["routes"]:
        if not r["routed"]:
            yield r["id"], "no clear path; drawn straight", "move the boxes closer or add gaps between them"


def r_arrow_through(m, c):
    for r in m["routes"]:
        for p, q in segments(r["pts"]):
            for b in m["boxes"]:
                if b["kind"] in ("node", "card", "legend", "header", "boundary") and \
                        (b["kind"] == "header" or b["id"] not in r["exempt"]) and seg_hits_rect(p, q, b["rect"]):
                    yield r["id"], f"arrow passes through {b['id']}", "re-arrange so the path is clear"
                    break


def r_crossings(m, c):
    n = sum(1 for a, b in combinations(m["routes"], 2)
            for s in segments(a["pts"]) for t in segments(b["pts"]) if segs_cross(s, t))
    if n > c["max"]:
        yield "flows", f"{n} arrow crossings (max {c['max']})", "reorder columns or cards so connected boxes sit nearer"


def r_bends(m, c):
    for r in m["routes"]:
        if len(r["pts"]) - 2 > c["max"]:
            yield r["id"], f"{len(r['pts']) - 2} bends (max {c['max']})", "place the two boxes nearer each other"


def r_short_arrow(m, c):
    for r in m["routes"]:
        n = sum(abs(p[0] - q[0]) + abs(p[1] - q[1]) for p, q in segments(r["pts"]))
        if n < c["min"]:
            yield r["id"], f"arrow is only {n:.2f} in long (min {c['min']})",                 "give the card more height (weight), fewer nodes, or put the two nodes side by side"


def r_label_collision(m, c):
    for r in m["routes"]:
        if r.get("label_collided"):
            yield r["id"], "label overlaps a box or another label", "shorten the label or widen the gap"


def r_safe_area(m, c):
    for b in m["boxes"]:
        if b["kind"] in ("zone", "card", "node", "legend", "boundary") and not inside(b["rect"], m["content"]):
            yield b["id"], "outside the content area", "reduce content or margins"


def r_density(m, c):
    nodes = sum(1 for b in m["boxes"] if b["kind"] == "node")
    if nodes > c["maxNodes"]:
        yield "slide", f"{nodes} nodes (max {c['maxNodes']})", "split into an overview and a detail slide"
    if m["columns"] > c["maxColumns"]:
        yield "slide", f"{m['columns']} columns (max {c['maxColumns']})", "merge columns or split the slide"


def r_nodes_per_card(m, c):
    counts = {}
    for b in m["boxes"]:
        if b["kind"] == "node":
            counts[b["parent"]] = counts.get(b["parent"], 0) + 1
    for k, n in counts.items():
        if n > c["max"]:
            yield k, f"{n} nodes in one container (max {c['max']})", "split it into two cards"


def r_contrast(m, c):
    for fg, bg, where in m["color_pairs"]:
        if fg and bg and contrast(fg, bg) < c["min"]:
            yield where, f"contrast {contrast(fg, bg):.1f}:1 (min {c['min']}:1)", "darken the text or lighten the fill"


def r_traceability(m, c):
    if not m["sequence"]:
        return
    known = {msg["n"] for msg in m["sequence"]["messages"]}
    for r in m["routes"]:
        bad = [n for n in r["steps"] if n not in known]
        if bad:
            yield r["id"], f"steps {bad} are not in the sequence diagram", "fix the step numbers or the diagram"


def r_unused_messages(m, c):
    for s in m["skipped"]:
        yield f"message {s['n']}", f"not drawn: {s['why']}", "map the participant to a node, or accept it"


def r_legend(m, c):
    if len(m["flow_types"]) > 1 and not m["has_legend"]:
        yield "slide", f"{len(m['flow_types'])} flow types but no legend", 'add {"legend": true} to the spec'


SMALL_WORDS = {"a", "an", "and", "as", "at", "by", "for", "in", "of", "on", "or", "the", "to", "via", "with"}


def _title_case(t):  # every word capitalised, except small words after the first; acronyms (API, SAP) are fine
    words = [w for w in re.split(r"[\s/&-]+", t) if w and w[0].isalpha()]
    return all(w[0].isupper() or (i and w in SMALL_WORDS) for i, w in enumerate(words))


def r_title_case(m, c):
    for b in m["boxes"]:
        t = b.get("title")
        if b["kind"] == "zone" and t and ((c["case"] == "upper" and t != t.upper()) or
                                         (c["case"] == "title" and not _title_case(t))):
            yield b["id"], f"title {t!r} is not {c['case']} case", f"write it in {c['case']} case"


def r_boundary_label(m, c):
    for b in m["boxes"]:
        if b["kind"] == "boundary" and not b.get("title"):
            yield b["id"], "boundary has no label", "name it (e.g. 'Client network', 'Azure subscription')"


CHECKS = {
    "overlap": r_overlap, "text-fit": r_text_fit, "min-font": r_min_font, "name-lines": r_name_lines,
    "desc-lines": r_desc_lines, "name-chars": r_name_chars, "desc-chars": r_desc_chars,
    "icon-ratio": r_icon_ratio, "icon-consistency": r_icon_consistency, "missing-icon": r_missing_icon,
    "generic-icon": r_generic_icon, "unrouted-flow": r_unrouted, "arrow-through-box": r_arrow_through,
    "arrow-crossings": r_crossings, "arrow-bends": r_bends, "short-arrow": r_short_arrow, "label-collision": r_label_collision,
    "safe-area": r_safe_area, "density": r_density, "nodes-per-card": r_nodes_per_card, "contrast": r_contrast,
    "traceability": r_traceability, "unused-messages": r_unused_messages, "legend": r_legend,
    "title-case": r_title_case, "boundary-label": r_boundary_label,
}
assert set(CHECKS) == set(DEFAULT_RULES)


def run(model, style, spec):
    rules, unknown = resolve(style, spec)
    issues = [{"rule": "config", "severity": "warn", "id": k, "message": "unknown rule name", "fix": "check spelling"}
              for k in unknown]
    for name, conf in rules.items():
        if conf["severity"] == "off":
            continue
        for target, msg, fix in CHECKS[name](model, conf):
            issues.append({"rule": name, "severity": conf["severity"], "id": target, "message": msg, "fix": fix})
    errors = sum(i["severity"] == "error" for i in issues)
    warns = sum(i["severity"] == "warn" for i in issues)
    return {"ok": errors == 0, "score": max(0, 100 - 10 * errors - 3 * warns), "errors": errors,
            "warnings": warns, "rules": {k: v for k, v in rules.items() if v["severity"] != "off"}, "issues": issues}
