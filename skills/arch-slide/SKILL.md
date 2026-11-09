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

Read when needed: [references/rules.md](references/rules.md) (lint rules). The spec format ([references/spec.md](references/spec.md)) is the builder agent's job.

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

## Checkpoint 4: build → review (in subagents, to keep this conversation small)
1. Dispatch the **arch-builder** agent with:
   - the build.py path (`<this skill's base directory>/build.py`) and `<this skill's base directory>/references/spec.md`;
   - the work folder `./arch-work/<slug>/` (holds `01-brief.md` and `02-flow.mmd`) and the style name;
   - the deliverable path `./architecture/<UseCase>-Architecture.pptx`;
   - any project rule overrides the user asked for (e.g. a big landscape: `"density": {"maxNodes": 60}`).

   It writes the spec, builds, fixes the spec until lint has no errors, renders the preview, and replies in
   ≤ 15 lines. Don't redo its work here, and don't read the engine sources.
2. Dispatch the **arch-reviewer** agent once with:
   - the preview PNG, `arch-spec.json`, `lint.json`, `02-flow.mmd` and `01-brief.md`;
   - the style folder, and the org reference screenshot if any.

   Save its answer verbatim to `./arch-work/<slug>/review.md`.
3. **FAIL** → send the findings to the same arch-builder (SendMessage), then review once more. **PASS with minor
   findings** → have the builder apply them; no new review is needed. At most 2 review rounds: if the slide still
   fails, tell the user what is left and why.
4. A **Stop hook** blocks finishing while `lint.json` has errors, or `review.md` is missing or says FAIL.
5. Report:
   - the output path and the preview (look at it once yourself);
   - the lint score and any accepted warnings;
   - icon stand-ins, and steps not drawn (`unused-messages`);
   - the assumptions to check with the client.

## Guardrails
- Don't invent scope. Every component traces to the brief; every arrow traces to a numbered message.
- SOW content stays local. The only network use is `fetch` (fixed official icon URLs, nothing about the work is sent).
- Never pass off a different product's icon. A stand-in is generic, flagged `iconStandIn`, and reported.
- Fonts never go below the org's minimum: shorten text or re-layout instead.
- Ask when ambiguous (stack, style, which screenshot is canonical) rather than guessing.
