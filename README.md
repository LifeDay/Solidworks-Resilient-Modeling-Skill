# Solidworks-Skills

A Claude Code skill for SolidWorks modeling: it teaches Claude to build parts and assemblies using the **Resilient Modeling Strategy (RMS)** — a feature-tree discipline that keeps generated models editable by both the build script and the SolidWorks GUI.

## Contents

| File | Purpose |
|---|---|
| [SKILL.md](SKILL.md) | The skill definition Claude Code loads: the six-folder feature-tree structure, parameterization rules, sketch and selection policy, naming conventions, and escape-hatch process. |
| [api-notes.md](api-notes.md) | SolidWorks API (pywin32) reference for the calls the skill and checker depend on — folders, `EquationMgr`, feature descriptions, traversal, suppression testing, sketch constraint status. Signatures shift between SolidWorks releases; verify against the local API Help. |
| [rms_check.py](rms_check.py) | A standalone checker that connects to a running SolidWorks session and audits the active document against the RMS rules. |

## What RMS gives you

Scripted feature trees tend to fail the first time someone edits a dimension, because features get wired to whatever geometry happened to be under the cursor when they were created. RMS fixes this by:

- Grouping every feature into one of six ordered folders (`1-Ref`, `2-Construction`, `3-Core`, `4-Detail`, `5-Modify`, `6-Quarantine`) so dependencies only ever point earlier in the tree.
- Driving every meaningful dimension from named global variables and equations instead of literal values, so the model is editable in the GUI, not just by re-running the script.
- Enforcing a deterministic selection policy (select by name, then by parameter-derived coordinate, then by a feature's own created-entity list) instead of picking raw topological indices.

See [SKILL.md](SKILL.md) for the full ruleset.

## Using the skill

**Option 1 — Claude Code plugin marketplace (recommended):**

```
/plugin marketplace add LifeDay/Solidworks-Skills
/plugin install solidworks-rms@solidworks-skills
```

**Option 2 — manual clone:**

Clone this repo's contents into your Claude Code skills directory (`~/.claude/skills/solidworks-rms/`, or a project's `.claude/skills/solidworks-rms/`) so `solidworks-rms` is available.

Either way, the skill activates automatically for SolidWorks modeling work — writing or editing macros that create geometry, driving SolidWorks via MCP, reviewing an existing feature tree, or building a part family — even if you don't say "RMS" explicitly.

## Verifying a model

With the target part or assembly open and active in SolidWorks:

```
python rms_check.py
```

This walks the feature tree and reports PASS/FAIL per rule, including a per-feature suppression test for every `4-Detail` feature (the fastest way to catch a Detail-to-Detail reference that slipped through).

Useful flags:

- `--dump-types` — print every feature's name, group, and API type string, then exit. Run this first against an unfamiliar part, since feature type strings vary across SolidWorks versions and the checker's classification sets may need calibrating.
- `--no-suppress` — skip the suppression test (read-only; makes no changes to the model).
- `--exceptions <file>` — path to a JSON file of approved rule waivers (`{rule_id: reason}`), default `rms_exceptions.json`. Waived rules are reported as accepted deviations instead of failures.

### Requirements

- SolidWorks running with the target document active.
- Python with [`pywin32`](https://pypi.org/project/pywin32/) installed (`pip install pywin32`).
- Windows, since the checker talks to SolidWorks over COM.
