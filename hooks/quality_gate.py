"""Stop hook: block "done" while the latest arch-slide build has lint errors, no review, or a FAIL review.

A PASS stays valid when minor findings are fixed afterwards, so small fixes don't force another review round.

Looks only at ./arch-work/*/lint.json written in the last 6 hours, so other sessions are unaffected.
Blocks once per stop attempt (stop_hook_active): Claude must fix the slide or tell the user why it can't pass.
"""
import json
import sys
import time
from pathlib import Path


def main():
    try:
        data = json.load(sys.stdin)
    except ValueError:
        return
    if data.get("stop_hook_active"):
        return
    work = Path(data.get("cwd") or ".") / "arch-work"
    lints = sorted(work.glob("*/lint.json"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not lints or time.time() - lints[0].stat().st_mtime > 6 * 3600:
        return
    latest = lints[0]
    lint = json.loads(latest.read_text(encoding="utf-8"))
    problems = []
    errors = [i for s in lint.get("slides", []) for i in s["lint"]["issues"] if i["severity"] == "error"]
    if errors:
        top = "; ".join(f"{e['rule']} ({e['id']}): {e['message']}" for e in errors[:4])
        problems.append(f"{len(errors)} lint error(s) in {latest.parent.name}: {top}")
    review = latest.with_name("review.md")
    if not review.exists():
        problems.append("the build has no review: run the arch-reviewer agent on the preview PNG and save "
                        f"its verdict to {review}")
    elif not review.read_text(encoding="utf-8").lstrip().upper().startswith("VERDICT: PASS"):
        problems.append(f"{review.name} does not say VERDICT: PASS: fix the findings with arch-builder and "
                        "review again (2 rounds at most, then tell the user what is left)")
    if problems:
        print(json.dumps({"decision": "block", "reason": "arch-slide quality gate: " + " | ".join(problems) +
                          ". Fix the spec, rebuild, preview and review, or tell the user exactly why it can't pass."}))


if __name__ == "__main__":
    main()
