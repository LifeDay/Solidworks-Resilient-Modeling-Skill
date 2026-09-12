# solidworks-rms skill — findings from a live build session (2026-09-12)

Context: tried to build a NEMA34 mounting flange via pywin32 against a live SolidWorks
2026 (rev 34.1.1) session, following the skill's api-notes.md. Got the base plate
(sketch + boss-extrude) working end to end; hit enough API friction along the way,
plus one real safety incident, to be worth folding back into the skill before the
next attempt.

## Safety incident (highest priority to fix)

`sw.NewPart` is not a reliable way to get an isolated scratch document — in one
call it returned the user's real, already-open `ActiveDoc` instead of creating a
new part. Two consequences followed from trusting the returned handle:

1. An API-probing call (`InsertSketch`) landed on the user's real document and left
   a stray empty `Sketch1` in it.
2. A later cleanup step (`CloseDoc` using a title captured earlier in the same
   script) closed that same real, unsaved document, discarding it. It happened to
   be a blank untitled part with nothing of value yet, so no real work was lost —
   but the mechanism would just as easily have closed a real in-progress part.

**Root cause:** nothing in the workflow re-verifies document identity immediately
before a destructive call. A title/handle captured at the top of a script is
treated as still valid several calls later, in a session where the user's own
document lives in the same `SldWorks.Application` process with no isolation.

**Recommended skill change:** add a hard rule to SKILL.md (or api-notes.md) —
*never* call `CloseDoc`/`QuitDoc`/`EditDelete` etc. without re-fetching and
re-checking the target document's title/identity in the same breath as the call,
and never run exploratory API probes against `sw.ActiveDoc` or an unverified
"scratch" doc when a user document might be open in the same session. Prefer
creating scratch docs with an identifying custom property or a title check
(`assert doc.GetTitle not in {known_user_titles}`) before any mutating call.

## Dynamic-dispatch (pywin32 late binding) quirks

These cost the most trial-and-error and should go straight into api-notes.md:

- **Zero-argument getter-style methods auto-invoke on attribute access** and must
  be called *without* parentheses — `doc.FirstFeature`, `feat.GetNextFeature`,
  `doc.GetEquationMgr`, `sketchSeg.GetType`, `sketchSeg.GetStartPoint2`,
  `doc.GetTitle`, `ext.GetWhatsWrongCount`. Calling them with `()` raises
  `TypeError: 'X' object is not callable` because the value is already resolved.
- **Multi-arg methods with an Object/IDispatch parameter reject a bare Python
  `None`** with `Type mismatch` (COM `DISP_E_TYPEMISMATCH`). Fix: wrap it —
  `win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)`. This alone fixed
  `SelectByID2`'s persistent "Type mismatch" error on its `Callout` parameter;
  none of the other arguments needed explicit VARIANT typing.
- **Early binding doesn't work against this SW build.**
  `win32com.client.gencache.EnsureDispatch("SldWorks.Application")` (and
  `CastTo`, which calls it internally) fails with `Element not found` from
  `GetTypeInfo()` on essentially every live SW object tried. Early binding had
  to be abandoned for actual calls; stuck with plain `win32com.client.Dispatch`
  throughout.
  - **Useful workaround:** `gencache.EnsureModule(clsid, 0, major, minor)`
    against the SolidWorks type library's CLSID directly (found via
    `win32com.client.selecttlb.EnumTlbs()`) *does* succeed and generates a
    readable `.py` file in `gen_py` full of real method signatures (argument
    names, order, and count) — even though the generated module can't be used
    for dispatch. Reading that file is the fastest way to get a method's true
    signature instead of guessing. SolidWorks 2026's main typelib CLSID:
    `{83A33D31-27C5-11CE-BFD4-00400513BB57}`, major/minor `34/0`.
  - Enum constants (`swEndConditions_e`, `swSelectType_e`,
    `swFeatureSuppressionAction_e`, `swConstrainedStatus_e`,
    `swInConfigurationOpts_e`, `swFeatureTreeFolderType_e`, ...) are **not** in
    that typelib — they're in a separate one, "SOLIDWORKS 2026 Constant type
    library", CLSID `{4687F359-55D0-4CD3-B6CF-2EB42C11F989}`. Same
    `EnsureModule` trick works there.
- **`EquationMgr.Add3` silently fails** on this build — returns `-1` and does not
  add the equation, no exception raised. `Add2(index, equation_text,
  use_automatic_solve_order)` works correctly. api-notes.md already hedges
  ("older releases expose Add2 ... with fewer arguments") but frames it as a
  fallback for *older* releases; on this 2026 build it's actually `Add2` that
  works and `Add3` that's broken, which is backwards from what the note implies.

## Method-location / naming surprises (found by trial, not in the notes)

- `InsertSketch2` doesn't exist via dynamic dispatch on this build's
  `SketchManager` — use `InsertSketch(bool)` instead (toggles sketch edit mode
  the same way).
- `InsertAxis2(AutoSize)` — the call that turns two selected planes into a
  reference axis — lives on `IModelDoc2` (the document object), **not** on
  `FeatureManager`, despite every other feature-creation call living there.
- `AddDimension2(x, y, z)` also lives on `IModelDoc2`, not on
  `ModelDocExtension` — easy to guess wrong since most modern selection/geometry
  calls (`SelectByID2`, `GetWhatsWrongCount`) are on `Extension`.
- To select a *feature* (as opposed to a face/edge/plane/sketch) by name via
  `SelectByID2` for use in a folder-wrap or pattern-seed operation, the `Type`
  string must be `"BODYFEATURE"` — not the feature's own `GetTypeName2()` string
  (e.g. `"Extrusion"`), not `"FEAT"`, not `""`.
- `SketchManager.CreateCenterRectangle` returns **6** sketch segments, not 4: the
  4 real profile edges plus 2 construction diagonal lines (corner-to-corner,
  used internally for the symmetric constraint). Don't assume return-order
  correspondence to "bottom/left/top/right" — inspect
  `.ConstructionGeometry` (bool) to filter to the 4 real edges, then
  `.GetStartPoint2`/`.GetEndPoint2` (both zero-arg, no-parens) to tell
  horizontal from vertical.
- `FeatureManager.InsertFeatureTreeFolder2(2)` (the "Containing" folder type)
  against a **prior selection of the features to enclose** is the working
  pattern for RMS's per-group folders — confirmed end-to-end: select the
  group's features (`Type="SKETCH"` / `"BODYFEATURE"` as appropriate), call
  `InsertFeatureTreeFolder2(2)`, then rename the returned folder feature's
  `.Name` to the RMS group name (e.g. `"3-Core"`). This matches api-notes.md's
  guidance ("create the folder when the group starts") and is worth promoting
  from a hint to a concrete confirmed recipe with the exact call.
  - Caveat found while confirming this: flat tree traversal
    (`FirstFeature`/`GetNextFeature`) shows a synthetic
    `"<FolderName>___EndTag___"` marker feature immediately after a folder's
    contents. `rms_check.py`'s `walk()` doesn't special-case this, but it
    happens to be harmless (the end-tag is itself `FtrFolder`-typed, and every
    checker rule that matters skips `FOLDER_TYPE` features) — worth a one-line
    comment in `rms_check.py` so a future maintainer doesn't "fix" it into a
    bug.

## Confirmed-working recipe (worth adding to api-notes.md verbatim)

Base plate (sketch on a default plane → boss extrude), the one full slice that
was built and rebuilt with zero errors:

```python
ext = doc.Extension
sm = doc.SketchManager
fm = doc.FeatureManager

# sketch a centered square on Front Plane
select_by_id2(ext, "Front Plane", "PLANE")   # SelectByID2 w/ VARIANT(VT_DISPATCH, None) for Callout
sm.InsertSketch(True)
sm.CreateCenterRectangle(0, 0, 0, half, half, 0)
sm.InsertSketch(True)                         # exit sketch
sk1 = last_feature(doc)                       # walk FirstFeature/GetNextFeature to the tail

# extrude it
select_by_id2(ext, sk1.Name, "SKETCH")
boss = fm.FeatureExtrusion3(
    True, False, False,      # Sd, Flip, Dir
    0, 0,                     # T1, T2 (0 = swEndCondBlind)
    0.010, 0.010,             # D1, D2 depths (metres)
    False, False, False, False,
    0, 0,
    False, False, False, False,
    True, True, True,         # Merge, UseFeatScope, UseAutoSelect
    0, 0, False,
)
```

23 positional args, confirmed matching `IFeatureManager::FeatureExtrusion3`'s
real signature (verified against the `gencache.EnsureModule`-generated file,
not just trial and error).

## Suggested next steps for the skill

1. Add a "Working safely against a live session" section to SKILL.md covering
   the incident above, ahead of any API detail.
2. Rewrite api-notes.md's "Connecting" section to lead with the dynamic-dispatch
   quirks (no-parens zero-arg calls, `VARIANT(VT_DISPATCH, None)`) since they
   affect nearly every other call in the reference.
3. Add the `EnsureModule`-by-CLSID trick as the documented way to look up real
   signatures, replacing the current "verify against SolidWorks API Help"
   instruction (which is fine for a human but useless for an agent without a
   browser to a licensed help site).
4. Add the method-location table (`InsertAxis2` and `AddDimension2` on
   `IModelDoc2`, feature-selection needs `Type="BODYFEATURE"`) since these are
   exactly the kind of thing that "looks right" and fails silently or
   confusingly otherwise.
5. Land the confirmed folder-wrap recipe and the confirmed `FeatureExtrusion3`
   call as copy-pasteable examples rather than descriptions.
