# arch-slide

A Claude Code plugin that turns a **SOW or idea deck (PPTX)** into a **separate, architecture-only
PowerPoint file**: native, editable shapes in **your organisation's own style** (template, colours,
fonts, logo, footer), with your icon set.

```
SOW deck ──► 1. understand + pick stack ──► 2. numbered flow ──► 3. org style ──► 4. slide (.pptx)
               (you approve)                  (you approve)        (you approve)    (previewed + fixed)
```

## Install

```
/plugin marketplace add pratapsdev11/arch-slide
/plugin install arch-slide@arch-slide
```

Then install the Python dependencies once:

```
pip install python-pptx pymupdf
```

For previews, Claude renders the slide to PNG and checks it. That needs **PowerPoint** (Windows)
or **LibreOffice** (`soffice` on PATH; macOS / Linux / Windows).

## Use

```
/arch-slide:arch-slide  path/to/SOW.pptx
```

Or just ask: *"make an architecture slide for the PO reconciliation use case in this deck"*.
Claude asks for your org style the first time and saves it for reuse.

### Your org style

Give Claude any one of these:

| You provide | What happens |
|---|---|
| Your org's template / any deck in its style | `build.py template` builds a clean base (logo, footer, title area); `build.py style` extracts exact colours, fonts and sizes |
| 1–3 screenshots of your architecture slides | Claude writes the style from the images (colours marked approximate) |
| Words ("dark blue headers, rounded boxes, no legend") | Applied over neutral defaults |

Styles are saved in `~/.arch-slide/styles/<org>/` (`style.json`, `base.pptx`, optional `icons/`),
outside the plugin, so plugin updates never overwrite them.

### Icons

No icons are bundled. Add your own once:

- **Azure**: download the official *Azure Architecture Icons* pack from Microsoft, unzip it, then
  `python skills/arch-slide/build.py icons <unzipped>/Icons skills/arch-slide/assets/icons/azure`
- **Generic line icons**: any SVG set using `currentColor` (e.g. Lucide, Tabler), recoloured to your brand:
  `... build.py icons <svg_dir> skills/arch-slide/assets/icons/generic --color C00000`
- **Brand / product logos**: drop PNGs (transparent, ≥256 px) into `assets/icons/brand/`.
- Org-only icons can live in `~/.arch-slide/styles/<org>/icons/<folder>/`.

A spec refers to an icon as `folder/file` (e.g. `azure/data-factories`). A missing icon renders as a
visible `?` placeholder and is listed in the build report; it is never silently replaced.

**Licensing:** check each vendor's icon terms (e.g. Microsoft's Azure Architecture Icons page)
before redistributing icon files in a fork of this plugin.

## What's inside

```
.claude-plugin/plugin.json, marketplace.json
skills/arch-slide/
  SKILL.md          the 4-checkpoint workflow Claude follows
  build.py          text | template | style | icons | build | preview
  test_build.py     smoke check: python skills/arch-slide/test_build.py
  examples/sample.json
docs/usage.puml     sequence diagram of a full run
```

`build.py` can also be used without Claude:

```
python build.py build my-spec.json --style my-org -o out.pptx   # prints a JSON report
python build.py preview out.pptx                                  # out-1.png, ...
```

## Privacy

SOW content stays in your local working folder (`./arch-work/`). The plugin makes no network calls.
