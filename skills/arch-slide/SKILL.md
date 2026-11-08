---
name: arch-slide
description: Use when the user wants an architecture diagram or architecture slide from a SOW, idea deck, or solution description, or wants to set up their organisation's slide style for architecture diagrams. Produces a new, editable PPTX with native shapes in the org's own style.
---

# arch-slide

Turn a SOW / idea deck into a **separate PPTX containing only architecture slide(s)**, as native,
editable shapes in the user's org style. Claude does the thinking (understanding, stack, flow, spec);
`build.py` (in this skill's base directory) does the drawing, deterministically.

`B` below = `python "<this skill's base directory>/build.py"`. Requires `pip install python-pptx pymupdf`
(check once with `python -c "import pptx, pymupdf"`; if missing, tell the user the install command).

Work in `./arch-work/<use-case-slug>/` in the user's current directory. Save every step's output there so
work can resume. **Stop at each checkpoint** and get the user's confirmation before moving on.
If the user only wants to set up a style, go straight to Checkpoint 3.

## Inputs to ask for up front
1. The SOW / idea deck (PPTX) or pasted text. Which slides / use case, if the deck has several.
2. The org style source, any one of: a saved style name (`ls ~/.arch-slide/styles`), the org's template PPTX
   (best: exact colours, logo, footer), 1-3 screenshots of their architecture slides, or a description in words.

## Checkpoint 1: understand + choose stack
- `B text deck.pptx` → read all slide text, tables and notes. Never edit the original deck.
- Write `01-brief.md`: problem, business goal, data sources, systems (e.g. SAP S/4HANA), users, expected
  outputs, constraints, **assumptions / unknowns**, then 2-3 stack options (Azure / AWS / hybrid) with a
  recommendation, trade-offs and a concrete service list each. Mark every assumed component `(assumed)`.
  Recommend only what the SOW supports; don't invent client commitments.
- Ask with a multiple-choice question: which stack, and is the summary right. Iterate until approved.

## Checkpoint 2: numbered flow
- Write `02-flow.md`: a numbered message list (1, 2, 3 ...) plus a ```` ```mermaid sequenceDiagram ```` block of the
  end-to-end flow on the chosen stack. Show the numbered list in chat (the terminal can't render Mermaid;
  never paste client content into online renderers).
- Ask where more depth is needed: auth, error paths, human approval, write-back, monitoring. Iterate until "okay".
- These numbers become `step` badges on arrows. Every numbered arrow must map to a message here.

## Checkpoint 3: org style
Styles live in `~/.arch-slide/styles/<org>/` (`style.json` + optional `base.pptx` + optional `icons/`),
outside the plugin so updates never overwrite them.
- **Saved style exists** → show its summary (colours, header look, flow colours) and offer reuse.
- **From the org's PPTX**:
  1. Pick a clean content slide that shows the logo/footer. Run `B template deck.pptx --slide N -o ~/.arch-slide/styles/<org>/base.pptx`.
     Check `shapes_on_slide` in the output: `copied_to_layout` should be true only for chrome (bars, bands, logos).
     If it guessed wrong, re-run with `--keep "Shape A,Shape B"`.
  2. Run `B style deck.pptx --slide M -o ~/.arch-slide/styles/<org>/style.json`, where M is an existing diagram slide if there is one, else N.
     Read `_observed` and fix the roles (primary, headerText, text, muted, panelFill/Border, cardFill/Border).
- **From screenshots / words only**: write `style.json` by hand with `"template": null` (blank 16:9). The renderer
  draws the title, accent bar and footer from `chrome`. Colours read off a screenshot are approximate: list
  them in `"_estimated"` and say so. If screenshots disagree with each other, ask which one is the standard.
- Render `examples/sample.json` (in this skill's directory) with the style, preview it, and show the user a short
  style summary + the image. Adjust until they confirm.

`style.json` keys (all optional; anything missing falls back to neutral defaults in `build.py`):
`template, layout, font, sizes{title,header,name,desc,legend,minFontPt}, limits{nameMaxChars,descMaxChars},
colors{primary,headerText,text,muted,panelFill,panelBorder,cardFill,cardBorder},
panel{header:"text"|"bar", radius, border, shadow, pad}, gap[depth0,depth1,depth2], node{icon,iconLeftRatio},
content{left,top,right,bottom}, chrome{accentBar,footerBand,footerText,logo}, flows{<type>:{color,label,dash?}}, legend{title}`.
Flow types are whatever the org's legend uses (e.g. data / analytics / action / feedback).

## Checkpoint 4: build the slide
1. Write `arch-spec.json` (format below) from the approved flow + stack.
2. `B build arch-spec.json --style <org> -o "<UseCase>-Architecture.pptx"` → read the JSON report:
   - `errors`: fix the spec (bad ids, unknown flow types, overlapping boxes).
   - `missing_icons`: Glob the icon folders for a better name (`assets/icons/**`, `~/.arch-slide/styles/<org>/icons/**`).
     If none exists, leave the placeholder; never substitute an unrelated icon.
   - `overflow` / `too_long`: **shorten the text** or rebalance `weight`s. Never lower fonts below `minFontPt`
     (the renderer clamps anyway). Too dense? Split into an overview + detail slide (pass several spec files to one `build`).
3. `B preview "<UseCase>-Architecture.pptx"` → Read each PNG. Check for overlapping shapes, clipped text, arrows
   crossing content, placeholders, legend or footer collisions. Fix the spec, rebuild, and repeat until clean.
4. Report: output path, the preview, missing icons (placeholders), the `icons_used` list for a quick check,
   and the assumptions the user should double-check.

## Spec format
A tree of boxes. Each box splits its area among `children` by `weight` along `dir`.
```json
{
  "title": "Use Case — Architecture",
  "root": {"dir": "col", "children": [
    {"id": "pipeline", "dir": "row", "weight": 3, "children": [
      {"id": "sources", "title": "1. Sources", "children": [
        {"id": "erp", "name": "SAP S/4HANA", "desc": ["PO, invoices", "Master data"], "icon": "brand/sap"}]},
      {"id": "ai", "title": "2. Analytics & AI", "weight": 2, "children": [
        {"dir": "row", "children": [
          {"id": "a1", "name": "Scoring Agent", "boxed": true, "iconPos": "top", "desc": "Rank invoices", "icon": "generic/robot"}]}]}]},
    {"dir": "row", "children": [
      {"id": "gov", "title": "3. Governance", "dir": "row", "weight": 4, "children": []},
      {"legend": true}]}]},
  "flows": [{"from": "sources", "to": "ai", "type": "data", "step": 1},
            {"from": "pipeline", "to": "gov", "type": "action", "both": true}]
}
```
- Box: `id`, `title` (draws a panel; no title = invisible grouping), `dir` `row|col` (default `col`), `weight` (default 1), `gap`, `children`.
- Node (leaf, has `name`): `id`, `name`, `desc` (string or list = bullets), `icon` (`folder/file` without `.png`),
  `iconPos` `left|top` (auto by shape), `boxed: true` (draws it as a small card, e.g. an agent).
- `{"legend": true}`: legend box; lists only the flow types actually used.
- Flow: `from`/`to` = any box or node id; `type` = key in style `flows`; `both` = arrowheads both ends; `step` = number badge.
  Side by side → horizontal arrow; stacked → vertical; diagonal → L-shaped connector.
- Top-level zones are best kept ≤ 5 across and bands ≤ 3. Shape names (`zone:…`, `card:…`, `node:…`) show in PowerPoint's Selection Pane.

## Icons
Looked up as `<icon>.png` in `~/.arch-slide/styles/<org>/icons/`, then in `assets/icons/` of this skill.
Convert an SVG pack once: `B icons <svg_dir> <out_dir> [--color HEX]` (`--color` recolours `currentColor` line icons).
Never download icons at runtime.

## Guardrails
- Don't invent scope; mark assumed components; recommend only what the SOW supports.
- Every numbered arrow traces to a message in `02-flow.md`.
- SOW content is client-confidential: keep it in the local working directory; don't send it to external services.
- Missing icon → visible placeholder + reported. Never silently substitute.
- Never modify the user's original deck; output is always a new file.
- Ask when ambiguous (stack, which style, which screenshot is canonical) rather than guessing.
