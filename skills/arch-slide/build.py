#!/usr/bin/env python3
"""arch-slide: editable architecture slides from a JSON spec, in any org's style.

  text     <deck.pptx>                              dump slide text, tables, notes
  template <deck.pptx> --slide N -o base.pptx       base template: slide N's layout + repeated chrome, no slides
  style    <deck.pptx> --slide N -o style.json      draft style tokens from slide N
  fetch    azure|power-platform|generic|brands      download official icon packs once -> ~/.arch-slide/icons
  icons    <svg_dir> <out_dir> [--color HEX]        convert any SVG folder -> PNG (manual / offline route)
  seq      <flow.mmd>                               parse a Mermaid sequence diagram (participants, numbered messages)
  build    <spec.json>... --style S -o out.pptx     layout + route + render + lint; writes lint.json next to the spec
  preview  <deck.pptx> [--out-dir DIR]              one PNG per slide (next to the deck by default)
"""
import argparse
import copy
import json
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parent))  # engine, layout, lint, seq live next to this file
from engine import DEFAULT_STYLE, EMU, ICON_HOME, E, build, drop_slides  # noqa: E402
import seq as S  # noqa: E402


def cmd_build(a):
    build(a.specs, a.style, a.out, a.strict)


def cmd_seq(a):
    print(json.dumps(S.parse(Path(a.mmd).read_text(encoding="utf-8")), indent=1, ensure_ascii=False))


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


def flatten_gradients(svg):
    """MuPDF paints SVG gradients black; swap each gradient for its middle stop colour.
    ponytail: flat colours, not gradients. Use a full SVG renderer (resvg) if gradients must survive."""
    colors, refs = {}, {}
    for tag in re.finditer(rb"<(linearGradient|radialGradient)\b([^>]*?)(/?)>", svg):
        gid = re.search(rb'\bid="([^"]+)"', tag.group(2))
        if not gid:
            continue
        href = re.search(rb'href="#([^"]+)"', tag.group(2))
        body = b"" if tag.group(3) else svg[tag.end():svg.find(b"</" + tag.group(1) + b">", tag.end())]
        stops = re.findall(rb'stop-color\s*[=:]\s*"?\s*(#[0-9a-fA-F]{3,6})', body)
        if stops:
            colors[gid.group(1)] = stops[len(stops) // 2]
        elif href:
            refs[gid.group(1)] = href.group(1)
    for gid, target in refs.items():
        if target in colors:
            colors[gid] = colors[target]
    for gid, col in colors.items():
        svg = svg.replace(b"url(#" + gid + b")", col)
    return svg


def convert(items, out, color=None, size=256, brand=None, force=False):
    """items: (stem, svg bytes). Writes <out>/<normalised-name>.png. brand: {stem: hex} fill per icon."""
    import pymupdf
    pymupdf.TOOLS.mupdf_display_errors(False)
    out.mkdir(parents=True, exist_ok=True)
    done = skipped = failed = 0
    for stem, data in sorted(items):
        name = re.sub(r"^\d+-icon-service-", "", stem)  # Azure pack: 10021-icon-service-Data-Factories
        name = re.sub(r"[-_]scalable$", "", name, flags=re.I)  # Power Platform pack: PowerApps_scalable
        name = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")
        dst = out / f"{name}.png"
        if dst.exists() and not force:
            skipped += 1
            continue
        if color:
            data = data.replace(b"currentColor", f"#{color.lstrip('#')}".encode())
        if brand and stem in brand:
            data = data.replace(b"<svg ", f'<svg fill="#{brand[stem]}" '.encode(), 1)
        data = flatten_gradients(data)
        try:
            page = pymupdf.open(stream=data, filetype="svg")[0]
            z = size / max(page.rect.width, page.rect.height)
            page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=True).save(str(dst))
            done += 1
        except Exception:  # a broken SVG in a pack shouldn't stop the rest
            failed += 1
    return {"converted": done, "skipped_existing": skipped, "failed": failed, "out": str(out.resolve())}


def cmd_icons(a):
    items = ((p.stem, p.read_bytes()) for p in Path(a.svg_dir).rglob("*.svg"))
    print(json.dumps(convert(items, Path(a.out_dir), a.color, a.size, force=a.force)))


# Official sources only. "page" packs: the current zip link is read off the vendor's download page,
# so a new pack version doesn't break the fetch.
PACKS = {
    "azure": ("azure", "page", "https://learn.microsoft.com/en-us/azure/architecture/icons/",
              r"https://[^\"']+Azure_Public_Service_Icons[^\"']*\.zip"),
    "power-platform": ("microsoft", "page", "https://learn.microsoft.com/en-us/power-platform/guidance/icons",
                       r"https://[^\"']+Power-Platform-icons[^\"']*\.zip"),
    "generic": ("generic", "npm", "lucide-static", None),   # Lucide line icons, ISC
    "brands": ("brand", "npm", "simple-icons", None),       # Simple Icons brand logos, CC0
}


def cmd_fetch(a):
    import io
    import tarfile
    import urllib.request
    import zipfile

    def get(url):
        req = urllib.request.Request(url, headers={"User-Agent": "arch-slide"})
        with urllib.request.urlopen(req, timeout=300) as r:
            return r.read()

    report = {}
    for pack in a.packs:
        folder, kind, src, pat = PACKS[pack]
        try:
            if kind == "npm":
                meta = json.loads(get(f"https://registry.npmjs.org/{src}/latest"))
                version = f"{src}@{meta['version']}"
                tar = tarfile.open(fileobj=io.BytesIO(get(meta["dist"]["tarball"])), mode="r:gz")
                files = {m.name: tar.extractfile(m).read() for m in tar.getmembers() if m.isfile() and (
                    (m.name.startswith("package/icons/") and m.name.endswith(".svg")) or
                    m.name == "package/data/simple-icons.json")}
            else:
                link = re.search(pat, get(src).decode("utf-8", "ignore"))
                if not link:
                    raise RuntimeError(f"no download link found on {src}")
                version = link.group(0).rsplit("/", 1)[-1]
                z = zipfile.ZipFile(io.BytesIO(get(link.group(0))))
                files = {n: z.read(n) for n in z.namelist() if n.lower().endswith(".svg")}
        except Exception as e:  # offline / blocked network: say how to do it by hand
            manual = src if kind == "page" else f"https://www.npmjs.com/package/{src}"
            report[pack] = {"error": str(e), "manual": f"download from {manual}, unzip, then: "
                                                       f"build.py icons <svg_dir> {ICON_HOME / folder}"}
            continue
        data = files.pop("package/data/simple-icons.json", None)
        brand = {i["slug"]: i["hex"] for i in json.loads(data) if "slug" in i} if data else None
        items = ((Path(n).stem, b) for n, b in files.items())
        report[pack] = dict(convert(items, ICON_HOME / folder, a.color if pack == "generic" else None,
                                    a.size, brand, a.force), source=version, id_prefix=f"{folder}/")
    print(json.dumps(report, indent=1))


def cmd_preview(a):
    deck = Path(a.deck).resolve()
    prs = Presentation(str(deck))
    W = 1600
    Hpx = int(W * prs.slide_height / prs.slide_width)
    out_dir = Path(a.out_dir).resolve() if a.out_dir else deck.parent  # keep PNGs out of the deliverables folder
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = out_dir / deck.stem
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
    p.add_argument("--force", action="store_true", help="overwrite existing PNGs")
    p = sub.add_parser("fetch"); p.add_argument("packs", nargs="+", choices=list(PACKS))  # noqa: E702
    p.add_argument("--color", default="404040", help="line colour for the generic pack")
    p.add_argument("--size", type=int, default=256); p.add_argument("--force", action="store_true")  # noqa: E702
    p = sub.add_parser("seq"); p.add_argument("mmd")  # noqa: E702
    p = sub.add_parser("build"); p.add_argument("specs", nargs="+")  # noqa: E702
    p.add_argument("--strict", action="store_true", help="exit 1 when lint has errors")
    p.add_argument("--style", help="style dir/file or saved org name; omit for neutral"); p.add_argument("-o", "--out", required=True)  # noqa: E702
    p = sub.add_parser("preview"); p.add_argument("deck"); p.add_argument("--out-dir")  # noqa: E702
    a = ap.parse_args()
    {"text": cmd_text, "template": cmd_template, "style": cmd_style, "icons": cmd_icons,
     "build": cmd_build, "preview": cmd_preview, "fetch": cmd_fetch, "seq": cmd_seq}[a.cmd](a)


if __name__ == "__main__":
    main()
