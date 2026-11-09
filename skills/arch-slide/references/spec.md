# arch-spec.json

A tree of boxes. Each box splits its area among `children` along `dir`. Every child first gets the height its
content needs (text is measured with the real font). Leftover space is shared by `weight`. Widths in a row are
shared by `weight`.

**Alignment is automatic.** In a column of nodes (a card's node list, or plain rows of nodes), each node slides to
the height of the boxes it connects to beside it, so those arrows run straight. Nodes with no partner beside them
are spaced evenly, and the card ends where its content ends. Never add spacer boxes or tune `weight`s to straighten
arrows. To fix a bent arrow, put the two boxes in neighbouring columns, or reorder siblings. `"align": false` on a
box keeps the plain weight layout for its children.

```json
{
  "title": "Use Case — Architecture",
  "subtitle": "Technical Architecture (Azure + ERP)",
  "sequence": "02-flow.mmd",
  "flows": "from-sequence",
  "flowOptions": {"level": "node", "default": "data",
                  "overrides": [{"from": "api", "to": "approver", "type": "action", "label": "Approval"}]},
  "rules": {"density": {"maxNodes": 50}},
  "root": {"dir": "col", "children": [
    {"id": "cloud", "kind": "boundary", "title": "Azure subscription", "dir": "row", "weight": 3, "children": [
      {"id": "ingest", "title": "Ingest", "subtitle": "Data in", "children": [
        {"id": "adf", "name": "Azure Data Factory", "desc": ["Pipelines", "Schedules"],
         "icon": "azure/data-factories", "participants": ["ADF"]}]},
      {"id": "app", "title": "Process", "weight": 2, "children": [
        {"dir": "row", "children": [
          {"id": "api", "name": "Scoring API", "boxed": true, "iconPos": "top", "desc": "Rules + model",
           "icon": "azure/function-apps", "participants": ["API"]}]}]}]},
    {"id": "erp", "title": "ERP (system of record)", "participants": ["ERP"], "children": [
      {"panel": true, "dir": "row", "children": [
        {"id": "orders", "name": "Orders", "icon": "brand/sap"}]}]},
    {"legend": true}
  ]}
}
```

## Boxes
| Key | Meaning |
|---|---|
| `id` | unique; flows and lint refer to it |
| `title` | draws a panel: a **zone** at the top level, a **card** below that. No title = invisible grouping |
| `subtitle` | second header line (e.g. "Capture in the field") |
| `kind: "boundary"` | dashed trust/network boundary with a label (on-prem, VNet, subscription). Doesn't change nesting depth |
| `panel: true` | card without a header |
| `dir` | `row` or `col` (default `col`) |
| `weight` | share of the leftover space (default 1) |
| `gap` | override the sibling gap |
| `align` | `false` = don't auto-align this box's children (default on) |
| `participants` | sequence participants this box represents (zones and cards can represent a whole system) |

## Nodes (leaf boxes: they have `name`)
| Key | Meaning |
|---|---|
| `name`, `desc` | `desc` is a string, or a list (shown as bullets) |
| `icon` | `<folder>/<file>` without `.png` (e.g. `azure/key-vaults`, `brand/sap`, `generic/user`) |
| `iconStandIn: true` | a generic icon stands in for a product with no official icon; lint reports it |
| `iconPos` | `left` or `top`; by default `left` when the node is at least `node.iconLeftMinWidth` wide |
| `boxed: true` | draw the node as its own small card (agents, microservices) |
| `participants` | sequence participants this node represents |

Icon size = style `node.iconLeft` / `node.iconTop` × the name font size, so icons at the same level are always
the same size. Missing room is a lint error (`text-fit`), never a silently shrunk icon or font.

## Flows
Either derive them from the sequence diagram (preferred):
- `"flows": "from-sequence"` with `"sequence": "<file.mmd>"`. Each message becomes an arrow between the boxes
  owning its participants.
- Messages between the same two boxes merge into one arrow labelled with all their step numbers (`6 · 10 · 16`).
  Two-way traffic gets heads at both ends.
- Self-messages and unmapped participants are reported by the `unused-messages` rule.
- `flowOptions.level`: `node` (default), `card` or `zone`. This lifts arrows up to containers for a calmer overview slide.
- `flowOptions.default`: the flow type to use (a key in the style's `flows`). `overrides`: per pair `type`, `label`, `note`.

Or list them explicitly: `[{"from", "to", "type", "both"?, "step"? | "steps"?, "label"?, "note"?}]`. Step numbers
are checked against the sequence diagram when one is referenced (`traceability` rule).

Arrows are routed orthogonally around boxes and boundaries they don't belong to, and never along a card's border.
They are straight when the ports line up; otherwise they bend around obstacles. A zone with no fill and no border
(header bar only) can be crossed between its cards. Labels go beside a segment, clear of boxes and other arrows.

## Legend
`{"legend": true}` is a box like any other. It lists only the flow types actually used. Whether a legend is
required is an org rule (`legend`).

## Rules
`"rules"` in the spec overrides the org's rules for this project only. See rules.md.
