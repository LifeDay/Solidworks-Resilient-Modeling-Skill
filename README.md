# Solidworks-Skills

A Claude Code skill for SolidWorks modeling: it teaches Claude to build parts and assemblies using the **Resilient Modeling Strategy (RMS)** — a feature-tree discipline that keeps generated models editable by both the build script and the SolidWorks GUI.

## Contents

The plugin manifests live in `.claude-plugin/`; the skill itself lives in
`skills/solidworks-rms/`, which is where Claude Code looks for plugin skills.

| File | Purpose |
|---|---|
| [SKILL.md](skills/solidworks-rms/SKILL.md) | The skill definition Claude Code loads: the six-folder feature-tree structure, parameterization rules, sketch and selection policy, naming conventions, and escape-hatch process. |
| [api-notes.md](skills/solidworks-rms/api-notes.md) | Index into the API references below, plus the five findings that cost the most time. |
| [references/troubleshooting.md](skills/solidworks-rms/references/troubleshooting.md) | **Symptom → cause → fix**, keyed by the error string or behaviour you actually have in hand. Start here when something breaks. |
| [references/api-recipes.md](skills/solidworks-rms/references/api-recipes.md) | Sequences confirmed end-to-end against a live session, with real positional signatures. |
| [references/dispatch-quirks.md](skills/solidworks-rms/references/dispatch-quirks.md) | How pywin32 late binding behaves against this API — auto-invoking getters, typed nulls, reading signatures from the typelib. |
| [scripts/sw_helpers.py](skills/solidworks-rms/scripts/sw_helpers.py) | The COM helpers every recipe assumes: `connect` (forced late binding), `call0`, `none_dispatch`, `select_by_id2`, `select_origin`, `add_equation`, `wrap_in_folder`, `no_input_dim_dialog`, sketch-coordinate and geometry-measuring helpers, plus the live-session safety guards (`assert_scratch_doc`, `safe_close`, document tags, `save_as`). |
| [scripts/sw_preflight.py](skills/solidworks-rms/scripts/sw_preflight.py) | Environment check to run before any build script. |
| [capabilities.yaml](skills/solidworks-rms/capabilities.yaml) | What has actually been verified, and on which SolidWorks version. |
| [rms_check.py](skills/solidworks-rms/rms_check.py) | A standalone checker that connects to a running SolidWorks session and audits the active document against the RMS rules. |

All paths in the table are under `skills/solidworks-rms/`.

Signatures and enum values shift between SolidWorks releases. Everything here
was confirmed against **SW2026 SP1.1** on one machine; verify against the local
API Help before relying on it elsewhere.

## What RMS gives you

Scripted feature trees tend to fail the first time someone edits a dimension, because features get wired to whatever geometry happened to be under the cursor when they were created. RMS fixes this by:

- Grouping every feature into one of six ordered folders (`1-Ref`, `2-Construction`, `3-Core`, `4-Detail`, `5-Modify`, `6-Quarantine`) so dependencies only ever point earlier in the tree.
- Driving the true design parameters from named global variables and equations instead of literal values, so the model is editable in the GUI, not just by re-running the script.
- Enforcing a deterministic selection policy (select by name, then by parameter-derived coordinate, then by a feature's own created-entity list) instead of picking raw topological indices.

See [SKILL.md](skills/solidworks-rms/SKILL.md) for the full ruleset.

## Using the skill

**Option 1 — Claude Code plugin marketplace (recommended):**

```
/plugin marketplace add LifeDay/Solidworks-Skills
/plugin install solidworks-rms@solidworks-skills
```

To pick up a new release of an existing install, run
`/plugin marketplace update solidworks-skills` and then update the plugin.

**Option 2 — manual install:**

Copy the `skills/solidworks-rms/` folder (not the whole repo) to `~/.claude/skills/solidworks-rms/`, or to a project's `.claude/skills/solidworks-rms/`, so that `SKILL.md` sits directly inside it.

Either way, the skill activates automatically for SolidWorks modeling work — writing or editing macros that create geometry, driving SolidWorks via MCP, reviewing an existing feature tree, or building a part family — even if you don't say "RMS" explicitly.

## Before running anything

```
python skills/solidworks-rms/scripts/sw_preflight.py --fix
```

(Paths are shown from a clone of this repo. For an installed skill, use the
script's full path inside the install directory.)

Checks Python bitness, `pywin32`, and an attachable SolidWorks; warns when a
cached gen_py module would make plain `Dispatch` early-bound; lists the
documents that were already open; and clears the **"Input dimension value"**
system option, whose modal dialog makes scripted dimensioning appear to hang and
can leave the session in a state where the next save crashes SolidWorks.

It deliberately will not start SolidWorks (`--launch` opts in) — and it detects
the orphaned invisible instance a previous script's `Dispatch` call may have
left holding a license seat.

## Verifying a model

With the target part or assembly open and active in SolidWorks:

```
python skills/solidworks-rms/rms_check.py
```

This audits the **feature tree**, not the geometry — a part with a mis-centered
hole passes every rule. Zero rebuild errors is necessary, not sufficient.

This walks the feature tree and reports PASS/FAIL per rule, including a per-feature suppression test for every `4-Detail` feature (the fastest way to catch a Detail-to-Detail reference that slipped through).

Useful flags:

- `--dump-types` — print every feature's name, group, and API type string, then exit. Run this first against an unfamiliar part, since feature type strings vary across SolidWorks versions and the checker's classification sets may need calibrating.
- `--no-suppress` — skip the suppression test (read-only; makes no changes to the model).
- `--exceptions <file>` — path to a JSON file of approved rule waivers (`{rule_id: reason}`), default `rms_exceptions.json`. Waived rules are reported as accepted deviations instead of failures.

### Requirements

- SolidWorks running with the target document active.
- Python with [`pywin32`](https://pypi.org/project/pywin32/) installed (`pip install pywin32`).
- Windows, since the checker talks to SolidWorks over COM.
