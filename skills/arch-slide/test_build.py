"""Smoke checks for arch-slide. Run: python test_build.py   (needs python-pptx, pymupdf)"""
import json
import subprocess
import sys
import tempfile
from itertools import combinations
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import layout as L  # noqa: E402


def run(*args):
    args = [*args, "--json"] if args[0] == "build" else list(args)
    r = subprocess.run([sys.executable, str(HERE / "build.py"), *args], capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def issues(rep, rule=None):
    return [i for s in rep["slides"] for i in s["lint"]["issues"] if rule is None or i["rule"] == rule]


def test_build(tmp):
    spec = json.loads((HERE / "examples" / "sample.json").read_text(encoding="utf-8"))
    sources = spec["root"]["children"][0]["children"][0]["children"]
    sources[0]["desc"] = "x" * 120          # one 120-char word: text-fit + desc-chars, font not shrunk
    sources[1]["icon"] = "nope/missing"     # missing icon -> placeholder + lint error
    (tmp / "spec.json").write_text(json.dumps(spec), encoding="utf-8")
    out = tmp / "out" / "out.pptx"
    rep = run("build", str(tmp / "spec.json"), "-o", str(out))
    assert any(i["id"] == "files" for i in issues(rep, "missing-icon")), issues(rep)
    assert any(i["id"] == "erp" for i in issues(rep, "desc-chars"))
    assert not issues(rep, "unrouted-flow") and not issues(rep, "arrow-through-box"), issues(rep)
    assert (tmp / "lint.json").exists() and not rep["ok"]

    prs = Presentation(str(out))
    slide = prs.slides[0]
    W, H = prs.slide_width, prs.slide_height
    for sh in slide.shapes:
        assert -Inches(0.01) <= sh.left and sh.left + sh.width <= W + Inches(0.01), sh.name
        assert -Inches(0.01) <= sh.top and sh.top + sh.height <= H + Inches(0.01), sh.name
    zones = [sh for sh in slide.shapes if sh.name.startswith("zone:")]
    assert len(zones) == 6, [z.name for z in zones]  # 4 zones + 2 bands in sample.json
    for a, b in combinations(zones, 2):
        assert not L.overlap((a.left, a.top, a.width, a.height), (b.left, b.top, b.width, b.height), 0), (a.name, b.name)
    names = {e.get("name") for e in slide._element.iter(qn("p:cNvPr"))}
    assert {"missing-icon:files", "node:files", "legend"} <= names
    assert min(int(e.get("sz")) for e in slide._element.iter(qn("a:rPr")) if e.get("sz")) >= 800
    pics = [p.find(f".//{qn('p:cNvPr')}").get("name") for p in slide._element.iter(qn("p:pic"))]
    assert all(n.startswith("icon:") for n in pics), pics  # no images of text


def test_router():
    a, b, wall = (0.5, 2.0, 1.0, 1.0), (4.0, 2.0, 1.0, 1.0), (2.2, 1.0, 0.6, 3.0)
    p = L.route(a, b, [wall], {}, (0, 0, 6, 6))
    assert p and not any(L.seg_hits_rect(s, t, wall) for s, t in L.segments(p)), p
    assert L.route(a, b, [], {}, (0, 0, 6, 6)) == [(1.5, 2.5), (4.0, 2.5)] or len(L.route(a, b, [], {}, (0, 0, 6, 6))) == 2


def test_sequence_flows(tmp):
    (tmp / "flow.mmd").write_text("""sequenceDiagram
    autonumber
    participant U as User
    participant API
    participant DB
    U->>API: submit
    API->>DB: save
    DB-->>API: ok
    API->>API: audit
    API-->>U: done
""", encoding="utf-8")
    spec = {"title": "Seq", "sequence": "flow.mmd", "flows": "from-sequence",
            "root": {"dir": "row", "children": [
                {"id": "users", "title": "Users", "children": [{"id": "u", "name": "User", "participants": ["U"]}]},
                {"id": "app", "title": "App", "children": [{"id": "api", "name": "API", "participants": ["API"]}]},
                {"id": "data", "title": "Data", "children": [{"id": "db", "name": "Database", "participants": ["DB"]}]}]}}
    (tmp / "seq.json").write_text(json.dumps(spec), encoding="utf-8")
    rep = run("build", str(tmp / "seq.json"), "-o", str(tmp / "seq.pptx"))
    assert rep["slides"][0]["flows"] == 2                              # U<->API, API<->DB
    assert not issues(rep, "traceability") and not issues(rep, "unrouted-flow"), issues(rep)
    assert [i["id"] for i in issues(rep, "unused-messages")] == ["message 4"]  # the self-message
    labels = [sh.text_frame.text for sh in Presentation(str(tmp / "seq.pptx")).slides[0].shapes
              if sh.name.startswith("label:")]
    assert any(t.startswith("1 · 5") for t in labels) and any(t.startswith("2 · 3") for t in labels), labels


def test_rules_from_org_and_project(tmp):
    org = tmp / "org"
    org.mkdir()
    (org / "style.json").write_text(json.dumps({"rules": {
        "min-font": {"minPt": 10},                           # org guideline: nothing under 10 pt
        "title-case": {"severity": "error", "case": "upper"}}}), encoding="utf-8")
    spec = {"title": "Rules", "rules": {"density": {"maxNodes": 1}},   # project-level override
            "root": {"dir": "row", "children": [
                {"id": "z", "title": "Mixed Case", "children": [{"id": "a", "name": "A"}, {"id": "b", "name": "B"}]}]}}
    (tmp / "rules.json").write_text(json.dumps(spec), encoding="utf-8")
    rep = run("build", str(tmp / "rules.json"), "--style", str(org), "-o", str(tmp / "rules.pptx"))
    sev = {(i["rule"], i["severity"]) for i in issues(rep)}
    assert ("min-font", "error") in sev and ("title-case", "error") in sev and ("density", "warn") in sev, sev


def test_align(tmp):
    # one node on the right linked to the 3rd of 4 nodes on the left: it must slide level with it, arrow straight
    spec = {"title": "Align", "root": {"dir": "row", "children": [
        {"id": "L", "title": "Left", "children": [{"id": "ca", "title": "Card A", "children": [
            {"id": f"a{i}", "name": f"Node {i}"} for i in range(1, 5)]}]},
        {"id": "R", "title": "Right", "children": [{"id": "cb", "title": "Card B", "children": [
            {"id": "b", "name": "Bee"}]}]}]},
        "flows": [{"from": "b", "to": "a3", "type": "data"}, {"from": "a3", "to": "a4", "type": "data"}]}
    (tmp / "align.json").write_text(json.dumps(spec), encoding="utf-8")
    rep = run("build", str(tmp / "align.json"), "-o", str(tmp / "align.pptx"))
    assert rep["ok"], issues(rep)
    sh = {s.name: s for s in Presentation(str(tmp / "align.pptx")).slides[0].shapes}
    mid = lambda s: s.top + s.height / 2  # noqa: E731
    assert abs(mid(sh["node:b"]) - mid(sh["node:a3"])) < Inches(0.05), (mid(sh["node:b"]), mid(sh["node:a3"]))
    assert type(sh["flow:b->a3"]).__name__ == "Connector"  # 2 points: a straight arrow
    gap = sh["node:a4"].top - (sh["node:a3"].top + sh["node:a3"].height)
    assert gap >= Inches(0.28), gap / 914400  # linked nodes stacked in a card leave room for a visible arrow
    card = sh["card:cb"]
    assert card.top <= sh["node:b"].top and sh["node:b"].top + sh["node:b"].height <= card.top + card.height


def test_template(tmp):
    # 3 slides: same footer band at the bottom edge, a different title and a content box on each
    prs = Presentation()
    for i in range(3):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, prs.slide_height - Inches(0.3), prs.slide_width, Inches(0.3))
        band.name = "Footer Band"
        band.fill.solid()
        band.fill.fore_color.rgb = RGBColor(0xC0, 0, 0)
        s.shapes.add_textbox(0, 0, Inches(6), Inches(0.6)).text_frame.text = f"Title {i}"
        s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(2), Inches(4), Inches(3)).name = "Content"
    prs.save(tmp / "deck.pptx")
    rep = run("template", str(tmp / "deck.pptx"), "--slide", "1", "-o", str(tmp / "base.pptx"))
    copied = {s["name"] for s in rep["shapes_on_slide"] if s["copied_to_layout"]}
    assert "Footer Band" in copied and "Content" not in copied, rep
    base = Presentation(str(tmp / "base.pptx"))
    assert len(base.slides) == 0
    layout = [lay for m in base.slide_masters for lay in m.slide_layouts]
    assert [lay.name for lay in layout] == ["Architecture"]
    texts = "".join(t.text or "" for t in layout[0]._element.iter(qn("a:t")))
    assert "Footer Band" in {sh.name for sh in layout[0].shapes} and "Title" not in texts


def test_icons(tmp):
    # gradient-filled SVG (how Azure icons are drawn) must come out in colour, not black
    src = tmp / "svg"
    src.mkdir()
    (src / "10001-icon-service-Grad-Box.svg").write_text(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><defs><linearGradient id="g">'
        '<stop offset="0" stop-color="#ff0000"/><stop offset="1" stop-color="#ff0000"/></linearGradient></defs>'
        '<rect width="10" height="10" fill="url(#g)"/></svg>', encoding="utf-8")
    rep = run("icons", str(src), str(tmp / "png"))
    assert rep["converted"] == 1 and rep["failed"] == 0, rep
    from PIL import Image
    px = Image.open(tmp / "png" / "grad-box.png").convert("RGB").getpixel((128, 128))
    assert px[0] > 200 and px[1] < 60, px


if __name__ == "__main__":
    test_router()
    for t in (test_build, test_sequence_flows, test_rules_from_org_and_project, test_align, test_template, test_icons):
        with tempfile.TemporaryDirectory() as d:
            t(Path(d))
    print("ok")
