---
name: solidworks-rms
description: Build SolidWorks parts and assemblies via script using the Resilient Modeling Strategy — six ordered feature-tree folders, global-variable-driven dimensions, and a deterministic selection policy that survives edits. Use this skill whenever the work touches SolidWorks modeling in any form: writing or editing a pywin32/VBA/C# macro that creates geometry, driving SolidWorks through an MCP server, reviewing or repairing an existing feature tree, building a part family, or answering "how should I structure this model." Use it even when the user does not say "RMS" or "resilient" — if a .sldprt or .sldasm is going to be created or modified, this skill applies.
---

# SolidWorks Resilient Modeling Strategy

> **Paths.** Every file path in this skill (`scripts/...`, `references/...`,
> `rms_check.py`, `capabilities.yaml`) is relative to this skill's base
> directory — written `<skill dir>` below — not to the current working
> directory, which is the user's project. Run scripts by their full path, e.g.
> `python "<skill dir>/scripts/sw_preflight.py" --fix`.

## Why this exists

A generated feature tree fails in a specific way. The geometry looks right, the user changes one dimension, and the tree collapses — because features reference whatever entity happened to be under the cursor when they were created, and nothing records what the model was supposed to mean.

The Resilient Modeling Strategy (Richard Gebhard, 2013) fixes the topology half of that problem by ordering features so parent-child chains run one direction only. This skill adds the half RMS omits: the dimensional intent lives in named global variables, and the script that produced the model is kept as the source of truth.

Two artifacts, always:

1. **The build script** — a parameterized macro whose header declares every driving value. This is what gets version-controlled and edited.
2. **The part or assembly** — carrying those same values as SolidWorks global variables with dimensions driven by equations, so it is independently editable in the GUI.

Both edit surfaces must produce the same rebuild. A model that can only be changed by re-running the script has failed half the requirement.

## Working safely against a live session

A script talking to `SldWorks.Application` over COM is talking to the user's actual,
already-running SolidWorks session — not a sandbox. In one observed session,
`sw.NewPart()` returned the user's real `ActiveDoc` instead of creating a new part.
Trusting that handle led to an exploratory `InsertSketch` call landing a stray
sketch in the user's real document, and a later `CloseDoc` call — using a title
captured earlier in the script — closing that same real, unsaved document.

**Run `python "<skill dir>/scripts/sw_preflight.py" --fix` before any build script.** It
confirms an attachable SolidWorks, lists the documents that were already open
(pass these as `known_user_titles`), and clears the "Input dimension value"
option whose modal dialog is the single most expensive failure in this repo's
history. It refuses to start SolidWorks itself, and detects the orphaned
headless instance that a previous script's `Dispatch` call may have left behind.

Rules, not suggestions:

- **Never call a destructive or mutating method** (`CloseDoc`, `QuitDoc`,
  `EditDelete`, `SetSuppression2`, and the like) **without re-fetching and
  re-checking the target document's identity in the same breath as the call.**
  A title or handle captured earlier in the script is not trustworthy evidence of
  what that handle still points to — re-read `GetTitle` (no parens — see
  api-notes.md) or a custom property immediately before the call that acts on it.
- **Never run exploratory API probes** (`InsertSketch`, or anything else poking
  at geometry to see what happens) **against `sw.ActiveDoc` or an unverified
  "scratch" document.** The user's own document may be open in the same process.
- When a scratch document is needed, **give it an identifying custom property**
  at creation, and **assert `doc.GetTitle not in known_user_titles`** before any
  call that mutates or closes it. Do not proceed on an assumption that a "new"
  document call actually created something new.

`scripts/sw_helpers.py` ships `assert_scratch_doc`, `safe_close`, `tag_doc`,
`verify_tag` and `find_tagged_doc`, which enforce exactly these rules. Use them
rather than hand-rolling the check.

**Build in stages.** Split a build into short scripts run as separate processes.
Long single scripts have crashed SolidWorks twice; a seven-stage build of one
part never did. Stage 1 records the user's open titles in a small state file
and tags the new document. Every later stage finds the document by that tag
(`find_tagged_doc`), never by a remembered title, and calls `verify_tag`
immediately before each mutating call.

## The six groups

Every feature belongs to exactly one top-level folder. Build in this order — SolidWorks folders must hold contiguous features and features cannot be reordered past a dependency, so there is no reorganize-afterward path.

| Folder | Contains | May reference | Never |
|---|---|---|---|
| `1-Ref` | Reference planes, axes, coordinate systems, layout sketches, imported reference geometry | Origin and default planes only | Solid bodies |
| `2-Construction` | Surface bodies, 3D / composite / projected curves, split lines | `1-Ref` | Solid bodies |
| `3-Core` | Base solid features: extrude, revolve, sweep, loft, thicken. Structural fillets. Shell last. | `1-Ref`, `2-Construction`, earlier `3-Core` | Material removal that isn't shaping the core |
| `4-Detail` | Bosses, cuts, pockets, holes, threads. Holes last within the group. | `1-Ref`, `2-Construction`, `3-Core` **only** | Other `4-Detail` features (see exception below) |
| `5-Modify` | Draft, mirror, pattern — transforms before replications — then any final features | Anything earlier | — |
| `6-Quarantine` | Cosmetic chamfers and fillets. Chamfers first, then fillets largest-radius first. | Anything earlier | Being a parent to anything |

Two rules carry most of the weight.

**`4-Detail` features do not reference each other.** Each one attaches to the Core or to reference geometry. This is what makes a detail feature independently suppressible, deletable, and movable. The permitted exception is a genuinely coupled pair — a boss and the hole through it — which must be placed adjacent and wrapped in a named subfolder so the coupling is visible.

**Nothing references `6-Quarantine`.** Fillets and chamfers consume edges and renumber topology, which is exactly why they are last and why nothing may depend on them. A fillet must also never consume its own defining surfaces; if a radius is large enough to swallow the face that positions it, the radius is wrong or the fillet belongs in Core.

### Fillets: Core or Quarantine

Structural fillets go in `3-Core`. These are fillets that change what the part *is* rather than how it looks — interior fillets that must exist before Shell for correct wall behaviour, fillets that the draft angle depends on, fillets carrying a stress-concentration requirement. Everything else is a cosmetic round and goes in `6-Quarantine`.

When the call is ambiguous, ask. The test to apply: if removing this fillet changes a dimension, a wall thickness, or a manufacturing decision, it is structural.

### Ordering inside Core

Order Core so a later feature absorbs an earlier feature's overrun. That keeps necessary cuts inside Core with no exception. In a cast connecting rod, pockets that had to stop at the ring were cut *before* the ring existed:

1. Arm envelope boss, deliberately oversize.
2. Both pocket cuts, running past the arm end.
3. Ring boss, which refills the pocket runout.
4. Drafted plan-profile trim cut.

`rms_check.py` accepts cuts in `3-Core`. Whether a cut "shapes the core" is a judgement you make, not an automated check.

## Parameterization

Only true design parameters are driven by global variables — values a user would actually change. Every other dimension is set directly and given a meaningful name. A large equation set slows rebuilds and hides where values come from, so do not route a dimension through a global just because it could conceivably change.

The script declares every value in one header block, split into design parameters and the rest. It writes the design parameters into the model as global variables through `EquationMgr`, and drives the dimensions that consume them with equations referencing those variables. A design parameter typed into a sketch as `50` is a dead end; a dimension driven by `"plate_width"` is an editable model.

Name variables for what they mean in the design, not for the feature that consumes them — `bore_dia`, `wall_t`, `bolt_circle_dia`, `mount_hole_qty`. Values derived from a design parameter get their own equations (`clearance_dia = bolt_dia + 0.4`) so the relationship is recorded rather than baked into a number.

Values that exist only for construction are not globals: sketch overruns that make a profile reach past the part, such as `pocket_runout`, are named dimensions set directly. Label them in the header as non-design values, so nobody edits one expecting the part to change.

When a drawing dimension falls on a drafted wall, decide which height it applies at before building. Record that in the header, and check the widths it implies against the drawing. A 7 mm rim measured at the pocket floor and one measured at the rim top produce different parts. The rim-top reading needed derived globals with trig, and trig in equations takes degrees.

For part families, drive configurations from a design table generated by the script. Do not emit N separate part files.

## Sketches

- One sketch per feature. Do not share a sketch across features; the coupling it creates is invisible in the tree and defeats the group rules.
- Every sketch fully defined before the feature consumes it. An under-defined sketch is a model that moves when you aren't looking.
- Dimension and constrain to the origin, to reference planes from `1-Ref`, and to other sketch entities. Dimension to model edges only when there is no alternative, and only to Core geometry.
- Sketch on named reference planes from `1-Ref`, which are stable across every edit. Faces of Core features that can't move on their own — the top face of the base extrude — are also acceptable. Any other face is not.
- Use feature fillets, not sketch fillets. The exception is a radius that is part of a 2D profile's definition, such as a slot end or a cam profile.

## Multibody parts

Avoid multibody parts when possible. When one is necessary, use it as a modeling technique — a tool body for a combine, a local operation — never as a substitute for an assembly of separately made parts. If the bodies would be made, bought or inspected separately, they are components.

## Selection policy

RMS tells a human "if you can see it in the background, it is OK to link to it." A script has no eyes, so the rule becomes mechanical: **a feature may only reference entities created in a strictly earlier group.**

In descending order of preference:

1. **Select by name.** `SelectByID2` with the name of a plane, sketch, axis, or feature created earlier in the script. Names the script created are names the script controls.
2. **Select by parameter-derived coordinate.** When a face or edge must be picked, compute the pick point from the parameter set — `x = plate_width / 2` — so the selection tracks the design. Never hard-code a coordinate observed in a GUI session.
3. **Select from a feature's created-entity list.** For fillets, walk the faces or edges a known feature produced rather than picking from the whole body.

Never select by raw topological index harvested from an interactive session. Those indices are the primary cause of trees that rebuild correctly once and never again.

## Naming and intent

Leave default feature names in place (`Boss-Extrude3`, `Cut-Extrude1`, `Fillet7`). The default carries the feature type, which is the stable and useful part, and avoids names that go stale when a variable changes.

Put the design intent in **Feature Properties → Description** via the `Description` property on the feature. Every feature gets one. Write what the feature is for, not what it is — `Bearing seat, press fit to 6802` rather than `Cylindrical cut`.

Tell the user to enable Tree Display → Show Feature Descriptions, since the intent is invisible otherwise.

Folders are named exactly `1-Ref`, `2-Construction`, `3-Core`, `4-Detail`, `5-Modify`, `6-Quarantine`. The checker matches on these names.

## Assemblies

The same logic applies one level up.

- Mate to reference geometry — planes, axes, and coordinate systems belonging to the components — rather than to faces and edges, for the same topological-stability reason.
- Fix or fully constrain the first component; float the rest.
- Keep the mate list shallow. Long mate chains propagate failure the way parent-child chains do in a part.
- Drive component positions from assembly-level global variables where they are design parameters rather than incidental.
- Name mates for intent through the same Description mechanism.
- Avoid top-down (in-context) design when possible. When it is necessary, carry the shared interfaces in a skeleton or layout part and reference that; never create in-context references between sibling parts. Lock the references before release.
- Take standard hardware from one shared Toolbox in the vault, set to create parts rather than configurations, with parts named by standard and size.

## Drawings

- Show the model's dimensions on the drawing (model items) when the model is dimensioned the way the part is inspected. Otherwise add reference dimensions in the drawing.
- Do not re-dimension a part in the drawing just to change the dimensioning scheme. Fix the model.

## Verification

Run the checker before declaring a model finished:

```
python "<skill dir>/rms_check.py"
```

The checker audits the **feature tree**, not the geometry. A part with a wrong
dimension or a mis-centered hole passes every rule here. Zero rebuild errors and
a fully-defined sketch are necessary, not sufficient — this has bitten before.

It connects to the running SolidWorks session, walks the active document, and reports per-rule PASS/FAIL. It also tests that each `4-Detail` feature suppresses individually without a rebuild error, which is the fastest way to catch a Detail-to-Detail reference that slipped through.

First time against an unfamiliar part, run `python "<skill dir>/rms_check.py" --dump-types` to print every feature's name and API type string. Feature type names vary across SolidWorks versions and the checker's classification sets at the top of the file may need calibrating against real parts.

Report the checker output to the user. Do not describe a model as complete while a rule is failing.

## Escape hatches

These rules are defaults, not laws, but the bar for breaking one is high.

When a rule genuinely cannot be met, **stop and ask before proceeding**. Do not silently deviate and do not bury the deviation in a comment. State:

1. Which rule, precisely.
2. Why the geometry or the requirement forces it.
3. What the alternative would cost.
4. What breaks later as a result.

Then wait for confirmation. If the user approves, record the exception in the build script header and in the affected feature's Description, and add it to `rms_exceptions.json` so the checker reports it as an accepted deviation rather than a failure.

Needing more than one or two exceptions on a part usually means the Core is wrong, not that the rules are.

## API notes

`api-notes.md` indexes the API knowledge this skill depends on:

- **`references/troubleshooting.md`** — symptom → cause → fix. Go here first
  when something returns `None`, raises, hangs, or produces wrong geometry.
- **`references/api-recipes.md`** — sequences confirmed end-to-end, with real
  positional signatures.
- **`references/dispatch-quirks.md`** — how pywin32 late binding behaves here.

**Do not re-derive the COM helpers.** `scripts/sw_helpers.py` ships `connect`,
`call0`, `none_dispatch`, `select_by_id2`, `select_origin`, `last_feature`,
`add_equation`, `equations`, `wrap_in_folder`, `no_input_dim_dialog`, the
sketch-transform helpers (`sketch_frame`, `model_to_sketch`, `sketch_to_model`,
`runs_horizontal`), `save_as`, the document-tag guards, and the geometry reads
`volume` and `body_extents`. Every recipe assumes them. Re-deriving them per
script is how two SolidWorks crashes happened.

**Attach with `connect()`, never plain `win32com.client.Dispatch`.** It forces
late binding. On a machine with a gen_py cache for the SldWorks typelib, plain
`Dispatch` is silently early-bound, and every bare-getter call in the references
then returns a bound method instead of a value.

Build scripts live in the user's project, so put the skill's `scripts` folder on
`sys.path` by absolute path (or copy `sw_helpers.py` next to the build script):

```python
import sys; sys.path.insert(0, r"<skill dir>/scripts")
from sw_helpers import connect, call0, select_by_id2, mm, no_input_dim_dialog
```

Verify signatures against the local SolidWorks API Help before relying on them;
they shift between releases.

## Saying what is actually verified

`capabilities.yaml` records what has been driven against a real session, and on
which SolidWorks version. **Check it before promising a capability.** Findings
so far come from a single install (SW2026 SP1.1), so `verified` there means
"verified on 2026 SP1.1" and nothing stronger.

If the user's version is not listed, or the capability is `unverified`, say so
plainly and proceed as an experiment — do not present untested behaviour as
known-good. When something new is confirmed or refuted, add it to the ledger and
to `references/troubleshooting.md` rather than leaving it in conversation.

Two standing limits worth stating to the user up front:

- **A clean rebuild does not prove correct geometry.** `rms_check.py` validates
  tree structure only. Confirm shape another way: volume, face counts and
  normals, exact extents (`body_extents`), and standard-view screenshots after
  each major stage. A knife-edge rim that passed every structural check was
  caught only by looking at a side view.
- **Assemblies, drawings, design tables, sheet metal and surfacing are
  unverified here.** The guidance for them is reasoned from the part rules, not
  tested.
