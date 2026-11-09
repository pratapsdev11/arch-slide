"""Mermaid sequenceDiagram -> participants + numbered messages -> architecture flows.

The sequence diagram is the single source of truth: message numbers match Mermaid's `autonumber`,
and flows derived from it carry those numbers, so every arrow traces back to a message.
"""
import re

ARROW = r"-->>|->>|-->|->|--x|-x|--\)|-\)"
MSG = re.compile(rf"^(?P<src>[^\s:]+?)\s*(?P<arrow>{ARROW})\s*[+-]?\s*(?P<dst>[^\s:]+?)\s*:\s*(?P<text>.*)$")
DECL = re.compile(r"^(participant|actor)\s+(\S+)(?:\s+as\s+(.+))?$")
BLOCK = re.compile(r"^(loop|alt|opt|par|critical|break|rect)\b\s*(.*)$")


def parse(text):
    """Returns {"participants": [{id, label, kind}], "messages": [{n, from, to, text, reply, block}]}."""
    parts, msgs, blocks = {}, [], []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("%%") or line in ("sequenceDiagram", "autonumber") or line.startswith("```"):
            continue
        if line.startswith(("Note ", "note ", "activate ", "deactivate ", "title ", "box ", "autonumber ")):
            continue
        m = DECL.match(line)
        if m:
            parts[m[2]] = {"id": m[2], "label": (m[3] or m[2]).strip(), "kind": m[1]}
            continue
        m = BLOCK.match(line)
        if m:
            blocks.append(f"{m[1]} {m[2]}".strip())
            continue
        if re.match(r"^(else|and|option)\b", line):
            if blocks:
                blocks[-1] = blocks[-1].split()[0] + " " + line
            continue
        if line == "end":
            if blocks:
                blocks.pop()
            continue
        m = MSG.match(line)
        if m:
            for p in (m["src"], m["dst"]):
                parts.setdefault(p, {"id": p, "label": p, "kind": "participant"})
            msgs.append({"n": len(msgs) + 1, "from": m["src"], "to": m["dst"], "text": m["text"].strip(),
                         "reply": m["arrow"].startswith("--"), "block": blocks[-1] if blocks else None})
    return {"participants": list(parts.values()), "messages": msgs}


def _short(text, n=26):
    t = text.strip().rstrip(".")
    if len(t) > n:
        t = t[:n].rsplit(" ", 1)[0]
    return t[:1].upper() + t[1:]


def derive_flows(seq, owner, lift, opts, tops=()):
    """Turn messages into flows between boxes.

    owner: participant id -> node id (from each node's "participants" list)
    lift:  (node id, node id) -> the two box ids to draw the arrow between (node, card, zone or nested level)
    tops:  ids of the top-level columns; with labels "none", an arrow between two columns with no override label
           gets the first message's text, shortened, so the main hand-offs still say what they are
    opts:  {"level": node|card|zone, "labels": steps|none, "default": type, "overrides": [{from, to, type?, label?, note?}]}
           labels "none": plain arrows; step numbers are kept on the flow for traceability but not drawn.
    Returns (flows, skipped messages)."""
    block = opts.get("level", "node") != "node"  # messages inside one card/zone are collapsed on purpose
    pairs, skipped = {}, []
    for msg in seq["messages"]:
        a, b = owner.get(msg["from"]), owner.get(msg["to"])
        if a is None or b is None:
            skipped.append({"n": msg["n"], "why": f"participant {msg['from'] if a is None else msg['to']} not on the slide"})
            continue
        a, b = lift(a, b)
        if a == b:
            if not block:
                skipped.append({"n": msg["n"], "why": "inside one box"})
            continue
        key = tuple(sorted((a, b)))
        f = pairs.setdefault(key, {"from": a, "to": b, "steps": [], "texts": [], "dirs": set()})
        f["steps"].append(msg["n"])
        f["texts"].append(msg["text"])
        f["dirs"].add((a, b))
    over = {(o["from"], o["to"]): o for o in opts.get("overrides", [])}
    flows = []
    for (a, b), f in pairs.items():
        o = over.get((f["from"], f["to"])) or over.get((f["to"], f["from"])) or {}
        steps = " · ".join(str(n) for n in f["steps"])
        plain = opts.get("labels") == "none"
        if plain:
            label = o.get("label") or (_short(f["texts"][0]) if f["from"] in tops and f["to"] in tops else None)
        else:
            label = f"{steps}. {o['label']}" if o.get("label") else steps
        flows.append({
            "from": f["from"], "to": f["to"], "type": o.get("type", opts.get("default", "data")),
            "both": len(f["dirs"]) > 1, "steps": f["steps"], "plain": plain, "label": label,
            "note": o.get("note"),  # message text stays in the sequence diagram; arrows carry step numbers
            "messages": f["texts"],
        })
    return flows, skipped


if __name__ == "__main__":  # self-check
    s = parse("""sequenceDiagram
    autonumber
    participant A as App
    actor U as User
    U->>A: login
    alt bad password
      A-->>U: retry
    end
    A->>A: audit
    A->>+DB: save
    """)
    assert [p["id"] for p in s["participants"]] == ["A", "U", "DB"], s
    assert [m["n"] for m in s["messages"]] == [1, 2, 3, 4]
    assert s["messages"][1]["reply"] and s["messages"][1]["block"] == "alt bad password"
    flows, skipped = derive_flows(s, {"A": "app", "U": "user", "DB": "db"}, lambda x, y: (x, y), {})
    assert {(f["from"], f["to"], f["both"]) for f in flows} == {("user", "app", True), ("app", "db", False)}
    assert skipped == [{"n": 3, "why": "inside one box"}]
    print("ok")
