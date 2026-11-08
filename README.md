# arch-slide

A Claude Code plugin that turns a **SOW or idea deck (PPTX)** into a **separate, architecture-only PowerPoint
file**:
- native, editable shapes with official icons;
- in **your organisation's own style and rules**;
- every arrow traced to a numbered sequence diagram.

```
SOW ─► brief ─► sequence diagram (.mmd) ─► spec ─► engine ─► lint ─► preview ─► independent review ─► .pptx
        you approve   you approve       style + rules     (Stop hook won't let Claude finish until lint and review pass)
```

## Install

```
/plugin marketplace add pratapsdev11/arch-slide
/plugin install arch-slide@arch-slide
pip install python-pptx pymupdf
```
Previews need **PowerPoint** (Windows) or **LibreOffice** (`soffice`; any OS).

## Use

```
/arch-slide:arch-slide  path/to/SOW.pptx
/arch-slide:arch-style  (set up or change your org's style and rules)
```
Or just ask: *"make an architecture slide for the third use case in this deck"*.

## How it works

| Layer | What it does |
|---|---|
| **Sequence** (`seq.py`) | Parses the Mermaid `sequenceDiagram`. Messages become arrows between the boxes that own each participant. Steps between the same two boxes merge into one arrow (`6 · 10 · 16`). Nothing is retyped, so traceability holds by construction. |
| **Measure** (`layout.py`) | Text is measured with the real font file and wrapped the way PowerPoint wraps it. The report says when a font is missing and it had to estimate. |
| **Layout** (`engine.py`) | Every box first gets the height its content needs. Leftover space is shared by `weight`. Icons are a fixed multiple of the name font, so they stay consistent. Supports zones, cards, boxed nodes and dashed **trust boundaries**. |
| **Routing** (`layout.py`) | Orthogonal A* arrows that go around boxes and header bars. Bends, shared lanes and narrow slivers cost extra, so arrows are straight wherever the layout allows. Labels go beside the arrow, clear of text, and wrap to fit narrow gutters. |
| **Render** | Native PowerPoint shapes: grouped, named (`zone:…`, `node:…`, `flow:…`), and editable as ordinary shapes. |
| **Lint** (`lint.py`) | 25 rule types (overlap, text fit, font minimum, icon ratio, arrows through boxes, crossings, bends, label collisions, contrast, density, traceability, legend, header case, …). **Rule values come from your org's style** and can be overridden per project. Writes `lint.json` with a score and a fix for each issue. |
| **Review** (`agents/arch-reviewer`) | A separate subagent checks the preview against the org reference, the brief and the flow. The builder doesn't grade its own work. |
| **Gate** (`hooks/`) | A Stop hook blocks "done" while the latest build has lint errors or no fresh `VERDICT: PASS`. |

## Your org's style and rules

`/arch-slide:arch-style` learns both from whatever you give it:

| You provide | Result |
|---|---|
| Your template or any deck in the org style | exact colours, fonts, sizes, logo, footer (`template` + `style` commands) |
| Screenshots of your architecture slides | tokens read off the images (marked as estimates), plus visible conventions as rules |
| Written guidelines ("nothing under 9 pt, always a legend") | rules, e.g. `"min-font": {"minPt": 9}`, `"legend": {"severity": "error"}` |

Saved in `~/.arch-slide/styles/<org>/`, outside the plugin, so updates never overwrite them. Each project can
tighten or loosen a rule in its spec (`"rules": {...}`). The full catalogue is in
[skills/arch-slide/references/rules.md](skills/arch-slide/references/rules.md).

## Icons

Nothing is bundled. Fetch once, from official sources, into `~/.arch-slide/icons/` (about a minute, ~6,400 icons):
```
python skills/arch-slide/build.py fetch azure power-platform generic brands
```

| Pack | Source | Licence |
|---|---|---|
| `azure/` | Microsoft Azure Architecture Icons | Microsoft's icon terms |
| `microsoft/` | Microsoft Power Platform icons | Microsoft's icon terms |
| `generic/` | Lucide line icons | ISC |
| `brand/` | Simple Icons, in brand colours | CC0 (logos remain their owners' trademarks) |

Microsoft publishes no open Teams, Excel, SharePoint or Outlook icons. Drop official PNGs into
`~/.arch-slide/icons/brand/`; otherwise a generic stand-in is used, flagged and reported. Behind a firewall,
`fetch` prints the manual download page, and `build.py icons <svg_dir> <out_dir>` converts the SVGs.

## Output

- `./architecture/<UseCase>-Architecture.pptx`: the deliverable, alone in its folder.
- `./arch-work/<use-case>/`: brief, `.mmd` flow, spec, `lint.json`, `review.md`, previews.

## Commands (also usable without Claude)

```
build.py text <deck.pptx>                         dump slide text, tables, notes
build.py seq <flow.mmd>                           participants + numbered messages
build.py template <deck.pptx> --slide N -o base.pptx
build.py style <deck.pptx> --slide N -o style.json
build.py fetch azure power-platform generic brands
build.py build <spec.json>... --style <org> -o out.pptx [--strict]
build.py preview out.pptx --out-dir previews
python skills/arch-slide/test_build.py            self-checks
```

## Privacy

SOW content stays on your machine. The only network access is `fetch`, which downloads icon packs from fixed
official URLs and sends nothing about your work.
