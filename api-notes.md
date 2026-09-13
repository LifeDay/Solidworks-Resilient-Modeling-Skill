# SolidWorks API notes

The API knowledge this skill depends on, split by how you need to reach it. This
page is an index; the content lives in [references/](references/).

Everything here was confirmed against a live **SW2026 SP1.1 (rev 34.1.1)**
session unless an entry says otherwise. **Signatures and enum values shift
between releases — verify against the local SolidWorks API Help, or against the
installed typelib, before relying on any of it.** Where a call takes more than a
couple of arguments, check the argument order rather than assuming.

| Start here when… | Go to |
|---|---|
| Something is broken and you have an error string, a `None`, or a hang | **[references/troubleshooting.md](references/troubleshooting.md)** — symptom → cause → fix |
| You are about to write a call and want the working sequence | **[references/api-recipes.md](references/api-recipes.md)** — confirmed end-to-end recipes |
| You are wondering how pywin32 late binding behaves here | **[references/dispatch-quirks.md](references/dispatch-quirks.md)** — auto-invoking getters, typed nulls, typelib reading |
| You need to know whether something is actually verified | **[capabilities.yaml](capabilities.yaml)** — per-capability status by version |

## The five things that cost the most time

If you read nothing else:

1. **Zero-arg getters are properties here.** `doc.GetTitle`, not
   `doc.GetTitle()`. A `callable()`-based wrapper cannot fix this and will make
   it worse — use `sw_helpers.call0`.
   ([dispatch-quirks](references/dispatch-quirks.md#zero-argument-getters-auto-invoke--call-them-without-parentheses))
2. **Turn off "Input dimension value" before any scripted dimensioning.** It
   opens a modal on every dimension call; killing the client while it is open
   makes the next `Save()` crash SolidWorks.
   ([troubleshooting](references/troubleshooting.md#a-dimension-call-hangs-forever-but-solidworks-is-responding))
3. **`EquationMgr.Add3` silently fails — use `Add2` and check `GetCount`.**
   ([troubleshooting](references/troubleshooting.md#add3-adds-no-equation-and-returns--1))
4. **A clean rebuild does not mean correct geometry.** A wrong-point coincident
   relation produced a fully-defined sketch, zero rebuild errors, and wrong
   parts. Check face counts.
   ([troubleshooting](references/troubleshooting.md#a-hole-is-in-the-wrong-place-but-the-sketch-is-fully-defined-and-rebuilds-clean))
5. **You are driving the user's real session, not a sandbox.** Re-verify
   document identity in the same breath as any mutating call.
   ([troubleshooting](references/troubleshooting.md#a-script-created-a-sketch-in-the-users-real-document))

## Helpers

Do not re-derive `call0`, `none_dispatch`, `select_by_id2`, `last_feature`,
`add_equation` or `wrap_in_folder` per script — re-deriving them is how two
SolidWorks crashes happened. They ship in
[`scripts/sw_helpers.py`](scripts/sw_helpers.py), and every recipe assumes them.

Run [`scripts/sw_preflight.py`](scripts/sw_preflight.py) before any build script.
