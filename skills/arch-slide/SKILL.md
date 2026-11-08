---
name: arch-slide
description: Use when the user wants an architecture diagram or architecture slide from a SOW, idea deck, solution description or sequence diagram. Produces a separate, editable PPTX (native shapes, official icons) in the user's org style, traced to a numbered sequence diagram and checked by a rule engine and an independent reviewer.
---

# arch-slide

Turns a SOW / idea deck into a **separate PPTX with only the architecture slide(s)**. You do the thinking; the
engine does the drawing deterministically:

```
SOW -> brief -> sequence diagram (.mmd, the source of truth) -> spec (participants -> boxes)
    -> engine: measure text (real fonts) -> layout sized to the content -> route arrows around boxes
    -> native PPTX -> lint (org + project rules) -> preview -> arch-reviewer -> deliver
```

`B` = `python "<this skill's base directory>/build.py"`. Needs `pip install python-pptx pymupdf`
(check with `python -c "import pptx, pymupdf"`). Working files go in `./arch-work/<slug>/`. The deliverable goes in
`./architecture/<UseCase>-Architecture.pptx`, alone in its folder.
**Stop at every checkpoint** for the user's confirmation. Never edit the user's original deck.

Read when needed: [references/spec.md](references/spec.md) (spec format) and [references/rules.md](references/rules.md) (lint rules).

## Setup (once per machine)
- If `~/.arch-slide/icons/` is empty, run `B fetch azure power-platform generic brands` (about 1 min, official sources only).
  Never build a slide full of placeholders just because the packs haven't been fetched.

## Checkpoint 1: understand + choose stack
- `B text deck.pptx` → read every slide, table and note.
- Write `01-brief.md`:
  - problem, goal, data sources, systems, users, outputs, constraints;
  - **assumptions / unknowns**;
  - 2–3 stack options with a recommendation and trade-offs.
- Mark everything not in the SOW as `(assumed)`. Recommend only what the SOW supports.
- Ask with a multiple-choice question: which stack, and is the summary right. Ask about the 1–3 unknowns that change the
  drawing (e.g. what a component really does, where approvals happen).

## Checkpoint 2: sequence diagram (source of truth)
- Write `02-flow.mmd`: a Mermaid `sequenceDiagram` with `autonumber`, one participant per real component.
  Also write `02-flow.md`: the same messages as a numbered list, for reading.
- `B seq 02-flow.mmd` → check the participants and message numbers parse as intended.
- Show the numbered list in chat; never paste client content into online renderers.
- Ask where more depth is needed (auth, errors, approvals, write-back, monitoring). Iterate until "okay".

## Checkpoint 3: org style + rules
- Saved style (`ls ~/.arch-slide/styles`) → show its summary and its `rules`, then offer reuse.
- New org, or the user gives screenshots / a guideline → use the **arch-style** skill. It learns the tokens AND the org's
  rules (min font, icon ratio, legend, header case, nodes per card, ...) into `~/.arch-slide/styles/<org>/style.json`.

## Checkpoint 4: model → build → lint → review
1. Write `arch-spec.json` (see references/spec.md):
   - zones / cards / nodes / boundaries, sized by `weight`;
   - each node lists the sequence `participants` it represents;
   - `"sequence": "02-flow.mmd"` and `"flows": "from-sequence"`, so arrows and step numbers come from the diagram.
     Use `flowOptions` for the level (node / card / zone), types and labels;
   - icon ids from real files: Glob `~/.arch-slide/icons/<pack>/*keyword*`. If no product icon exists, use a generic
     one and set `"iconStandIn": true`;
   - project-specific rule overrides in `"rules"`, only when this project needs them (e.g. a big landscape:
     `"density": {"maxNodes": 60}`). Say why.
2. `B build arch-spec.json --style <org> -o "./architecture/<UseCase>-Architecture.pptx"`
   → read `lint` per slide (score, errors, warnings; each issue has a `fix`). Fix the **spec**, never the rules,
   to make errors go away. Repeat until errors = 0 and every remaining warning is deliberate.
   Too dense to pass? Split into an overview + detail slide (pass both specs to one `build`).
3. `B preview "./architecture/<UseCase>-Architecture.pptx" --out-dir ./arch-work/<slug>/preview` → look at the PNG yourself.
4. Dispatch the **arch-reviewer** agent with:
   - the preview PNG path(s);
   - `arch-spec.json`, `lint.json` and `02-flow.mmd`;
   - the style folder, and the org reference screenshot if any;
   - `01-brief.md`.

   Save its answer verbatim to `./arch-work/<slug>/review.md`.
   On FAIL: fix, rebuild, re-preview, re-review.
5. A **Stop hook** blocks finishing while `lint.json` has errors or `review.md` isn't a fresh `VERDICT: PASS`.
   If something genuinely can't pass, tell the user exactly what and why. Don't hide it.
6. Report:
   - the output path and the preview;
   - the lint score and any accepted warnings;
   - icon stand-ins, and steps not drawn (`unused-messages`);
   - the assumptions to check with the client.

## Guardrails
- Don't invent scope. Every component traces to the brief; every arrow traces to a numbered message.
- SOW content stays local. The only network use is `fetch` (fixed official icon URLs, nothing about the work is sent).
- Never pass off a different product's icon. A stand-in is generic, flagged `iconStandIn`, and reported.
- Fonts never go below the org's minimum: shorten text or re-layout instead.
- Ask when ambiguous (stack, style, which screenshot is canonical) rather than guessing.
