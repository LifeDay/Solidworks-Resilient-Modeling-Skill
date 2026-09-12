---
name: solidworks-rms
description: Build SolidWorks parts and assemblies via script using the Resilient Modeling Strategy — six ordered feature-tree folders, global-variable-driven dimensions, and a deterministic selection policy that survives edits. Use this skill whenever the work touches SolidWorks modeling in any form: writing or editing a pywin32/VBA/C# macro that creates geometry, driving SolidWorks through an MCP server, reviewing or repairing an existing feature tree, building a part family, or answering "how should I structure this model." Use it even when the user does not say "RMS" or "resilient" — if a .sldprt or .sldasm is going to be created or modified, this skill applies.
---

# SolidWorks Resilient Modeling Strategy

## Why this exists

A generated feature tree fails in a specific way. The geometry looks right, the user changes one dimension, and the tree collapses — because features reference whatever entity happened to be under the cursor when they were created, and nothing records what the model was supposed to mean.

The Resilient Modeling Strategy (Richard Gebhard, 2013) fixes the topology half of that problem by ordering features so parent-child chains run one direction only. This skill adds the half RMS omits: the dimensional intent lives in named global variables, and the script that produced the model is kept as the source of truth.

Two artifacts, always:

1. **The build script** — a parameterized macro whose header declares every driving value. This is what gets version-controlled and edited.
2. **The part or assembly** — carrying those same values as SolidWorks global variables with dimensions driven by equations, so it is independently editable in the GUI.

Both edit surfaces must produce the same rebuild. A model that can only be changed by re-running the script has failed half the requirement.

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

## Parameterization

Every dimension that could plausibly change is driven by a global variable. Literals belong only to values that are geometrically fixed by something else.

The script declares parameters in one header block. It then writes them into the model as global variables through `EquationMgr`, and drives dimensions with equations referencing those variables rather than setting dimension values directly. A dimension set to `50` is a dead end; a dimension set to `"plate_width"` is an editable model.

Name variables for what they mean in the design, not for the feature that consumes them — `bore_dia`, `wall_t`, `bolt_circle_dia`, `mount_hole_qty`. Derived values get their own equations (`clearance_dia = bolt_dia + 0.4`) so the relationship is recorded rather than baked into a number.

For part families, drive configurations from a design table generated by the script. Do not emit N separate part files.

## Sketches

- One sketch per feature. Do not share a sketch across features; the coupling it creates is invisible in the tree and defeats the group rules.
- Every sketch fully defined before the feature consumes it. An under-defined sketch is a model that moves when you aren't looking.
- Dimension and constrain to the origin, to reference planes from `1-Ref`, and to other sketch entities. Dimension to model edges only when there is no alternative, and only to Core geometry.
- Sketch on named reference planes rather than model faces wherever the geometry permits. A plane from `1-Ref` is stable across every edit; a face is not.

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

## Verification

Run the checker before declaring a model finished:

```
python scripts/rms_check.py
```

It connects to the running SolidWorks session, walks the active document, and reports per-rule PASS/FAIL. It also tests that each `4-Detail` feature suppresses individually without a rebuild error, which is the fastest way to catch a Detail-to-Detail reference that slipped through.

First time against an unfamiliar part, run `python scripts/rms_check.py --dump-types` to print every feature's name and API type string. Feature type names vary across SolidWorks versions and the checker's classification sets at the top of the file may need calibrating against real parts.

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

See `references/api-notes.md` for the SolidWorks API calls this skill depends on — folders, equations, descriptions, suppression, parent/child traversal, and sketch constraint status. Verify signatures against the local SolidWorks API Help before relying on them; they shift between releases.
