# Lint rules

The engine knows the rule **types**. The **values** belong to the org (`style.json` → `"rules"`, learned from its
guidelines, screenshots or text) and can be tightened or loosened per project (`arch-spec.json` → `"rules"`).
Resolution: built-in defaults ← org style ← project spec.
Each rule: `{"severity": "error" | "warn" | "off", ...params}`. Errors block delivery (Stop-hook gate).

| Rule | Default | Params | Checks |
|---|---|---|---|
| `overlap` | error | – | sibling boxes overlap |
| `text-fit` | error | – | measured text doesn't fit its box, or a word is wider than the box |
| `min-font` | error | `minPt` (null = `sizes.minFontPt`) | any style size below the org minimum |
| `name-lines` / `desc-lines` | warn | `max` 2 | names / descriptions wrap to too many lines |
| `name-chars` / `desc-chars` | warn | `max` 32 / 70 | text too long |
| `icon-ratio` | warn | `min` 2.4, `max` 5.0 | icon height ÷ name font size |
| `icon-consistency` | warn | – | icons at the same level and orientation differ in size |
| `missing-icon` | error | – | placeholder drawn |
| `generic-icon` | warn | – | a generic icon stands in for a product (`iconStandIn`) |
| `unrouted-flow` | error | – | no clear path; the arrow was drawn straight |
| `arrow-through-box` | error | – | an arrow passes through a box that isn't its endpoint or a container of one |
| `arrow-crossings` | warn | `max` 2 | arrow crossings on the slide |
| `arrow-bends` | warn | `max` 3 | bends in one arrow |
| `short-arrow` | warn | `min` 0.25 (in) | an arrow too short to read (just an arrowhead) |
| `label-collision` | warn | – | an arrow label overlaps a box or another label |
| `safe-area` | error | – | a box is outside the content area |
| `density` | warn | `maxNodes` 40, `maxColumns` 6 | slide too dense |
| `nodes-per-card` | warn | `max` 6 | too many nodes in one container |
| `contrast` | warn | `min` 4.5 | WCAG contrast of text on its fill |
| `traceability` | error | – | an arrow's step numbers aren't in the sequence diagram |
| `unused-messages` | warn | – | sequence messages not drawn (self-messages, unmapped participants) |
| `legend` | off | – | flows of 2+ types but no legend |
| `title-case` | off | `case` upper / title | zone titles not in the org's case |
| `boundary-label` | warn | – | a boundary without a label |

## Turning an org's guidelines into rules (examples)
| Guideline (screenshot or text) | style.json `rules` |
|---|---|
| "Nothing smaller than 9 pt" | `"min-font": {"minPt": 9}` |
| "Every diagram has a legend" | `"legend": {"severity": "error"}` |
| "Column headers in capitals" | `"title-case": {"severity": "warn", "case": "upper"}` |
| "Max 5 services per box" | `"nodes-per-card": {"max": 5}` |
| "Icons about 3× the label height" | `"icon-ratio": {"min": 2.6, "max": 3.4}` and `node.iconLeft: 3` |
| "Arrows must never cross" | `"arrow-crossings": {"severity": "error", "max": 0}` |
| "Descriptions are one line" | `"desc-lines": {"max": 1}` |
| Screenshot shows no legend | `"legend": {"severity": "off"}` |

Project overrides (in the spec) are for real project needs, e.g. a large landscape slide:
`"density": {"maxNodes": 60}`. Never loosen a rule just to make an error go away. Fix the slide.
