---
name: arch-style
description: Use when the user wants to set up or change their organisation's architecture-slide style. They might give a template or deck in the org style, screenshots of their architecture diagrams, or written style guidelines ("no legend, red headers, nothing under 9 pt"). Learns the visual tokens AND the org's rules and saves them for reuse by arch-slide.
---

# arch-style

Learns an org's architecture-slide style into `~/.arch-slide/styles/<org>/`:
- `style.json`: visual tokens (colours, fonts, sizes, header treatment, icon sizes) + `"rules"` (the org's guidelines
  as lint rule values) + `"_notes"` (spec conventions) + `"_estimated"` (values read off screenshots);
- `base.pptx`: the org template (logo, footer, title area), when a deck is given;
- `icons/` (optional): org-only icons, which take priority over the shared packs.

`B` = `python "<this skill's base directory>/../arch-slide/build.py"`.
Rule catalogue: `<this skill's base directory>/../arch-slide/references/rules.md`.

## 1. Collect
Ask for any of:
- the org template or any deck in the org style (best: exact values);
- 1–3 screenshots of real architecture slides;
- written guidelines.

Ask for the org name (used as the folder name). If screenshots disagree with each other, ask which is the standard.

## 2. Tokens
- **Deck given:**
  - `B template deck.pptx --slide N -o ~/.arch-slide/styles/<org>/base.pptx`, where N is a clean content slide with
    the logo and footer. Check `shapes_on_slide`: only chrome should be `copied_to_layout`. Fix with `--keep "A,B"`.
  - `B style deck.pptx --slide M -o ~/.arch-slide/styles/<org>/style.json`, where M is a diagram slide if there is one.
    Map `_observed` onto roles: `primary`, `headerText`, `text`, `muted`, `panelFill/Border`, `cardFill/Border`.
- **Screenshots / words only:** write `style.json` with `"template": null`. The renderer draws the title, accent bar and
  footer from `chrome`. List every colour read off a screenshot in `"_estimated"` and tell the user.
- Typical keys:
  - `font`, `sizes{title,subtitle,header,name,desc,legend,minFontPt}`;
  - `panel{header: "text"|"bar", align: "center"|"left", radius, border, shadow, pad}`;
  - `colors{...}`: `panelFill` / `panelBorder` set to `null` give borderless columns under header bars;
  - `gap[...]`, `node{iconLeft, iconTop, iconLeftMinWidth}`, `content{...}`;
  - `flows{<type>: {color, label, dash?}}`, `legend{title}`.

## 3. Rules (the org's guidelines, made checkable)
Read the guidelines and the screenshots for rules, then write them into `style.json` → `"rules"`, using the catalogue:
- **Stated rules:** "minimum 9 pt", "always a legend", "max 5 services per box", "headers in capitals".
- **Visible conventions:** no legend in any screenshot → `legend: off`; titles all uppercase → `title-case`;
  one-line descriptions → `desc-lines: {max: 1}`; arrows never cross → `arrow-crossings: {max: 0}`.
- **Icon-to-text proportion:** measure it on the screenshot (icon height ÷ label cap height ≈ ratio). Set `node.iconLeft` /
  `node.iconTop` and an `icon-ratio` band around them.

Put spec conventions the engine can't check in `"_notes"` (e.g. "always wrap nodes in a titled card").

## 4. Confirm
- Render and preview the sample in one command: `B build "<arch-slide dir>/examples/sample.json" --style <org>
  -o ./arch-work/style-sample.pptx --preview ./arch-work/style-preview`.
- Compare it side by side with the reference. Adjust tokens until it matches.
- Show the user the image and a summary:
  - the tokens;
  - every rule and where it came from (stated, or seen in the screenshot);
  - which values are estimated.
- Save on confirmation.
