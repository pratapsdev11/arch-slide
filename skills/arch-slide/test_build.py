"""Smoke check for build.py. Run: python test_build.py   (needs python-pptx)"""
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


def run(*args):
    r = subprocess.run([sys.executable, str(HERE / "build.py"), *args],
                       capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def overlap(a, b):
    return a.left < b.left + b.width and b.left < a.left + a.width and \
        a.top < b.top + b.height and b.top < a.top + a.height


def test_build(tmp):
    spec = json.loads((HERE / "examples" / "sample.json").read_text(encoding="utf-8"))
    sources = spec["root"]["children"][0]["children"][0]["children"]
    sources[0]["desc"] = "x" * 120          # long text -> reported, font not shrunk
    sources[1]["icon"] = "nope/missing"     # missing icon -> placeholder + report
    (tmp / "spec.json").write_text(json.dumps(spec), encoding="utf-8")
    out = tmp / "out.pptx"
    rep = run("build", str(tmp / "spec.json"), "-o", str(out))

    assert not rep["errors"], rep["errors"]
    assert {"node": "files", "icon": "nope/missing"} in rep["missing_icons"]
    assert any(t["id"] == "erp" for t in rep["too_long"])

    prs = Presentation(str(out))
    slide = prs.slides[0]
    W, H = prs.slide_width, prs.slide_height
    for sh in slide.shapes:
        assert -Inches(0.01) <= sh.left and sh.left + sh.width <= W + Inches(0.01), sh.name
        assert -Inches(0.01) <= sh.top and sh.top + sh.height <= H + Inches(0.01), sh.name
    zones = [sh for sh in slide.shapes if sh.name.startswith("zone:")]
    assert len(zones) == 6, [z.name for z in zones]  # 4 zones + 2 bands in sample.json
    for a, b in combinations(zones, 2):
        assert not overlap(a, b), (a.name, b.name)
    names = {e.get("name") for e in slide._element.iter(qn("p:cNvPr"))}
    assert "missing-icon:files" in names and "node:files" in names and "legend" in names
    min_sz = min(int(e.get("sz")) for e in slide._element.iter(qn("a:rPr")) if e.get("sz"))
    assert min_sz >= 800, min_sz            # neutral minFontPt = 8
    pics = [e for e in slide._element.iter(qn("p:pic"))]
    assert all(p.find(f".//{qn('p:cNvPr')}").get("name").startswith("icon:") for p in pics)  # no images of text


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


if __name__ == "__main__":
    with tempfile.TemporaryDirectory() as d:
        test_build(Path(d))
        test_template(Path(d))
    print("ok")
