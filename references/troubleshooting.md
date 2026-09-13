# Troubleshooting

Keyed by the **symptom you actually have in hand** — an error string, a return
value, an observed behaviour. Every entry is something that was hit for real and
cost time; the fixes are the confirmed ones, not speculation.

Each entry carries a **Verified on** tag. Where a workaround has since been
superseded by a root cause, it is marked so rather than deleted.

## Index

**Crashes and hangs**
- [A dimension call hangs forever but SolidWorks is Responding](#a-dimension-call-hangs-forever-but-solidworks-is-responding)
- [SolidWorks crashes on Save() with RPC -2147023170](#solidworks-crashes-on-save-with-rpc--2147023170)
- [Dimensioning a rectangle edge hangs or crashes SolidWorks](#dimensioning-a-rectangle-edge-hangs-or-crashes-solidworks)
- [A long script crashes SolidWorks partway through](#a-long-script-crashes-solidworks-partway-through)
- [The script deadlocks after adding a settle/wait helper](#the-script-deadlocks-after-adding-a-settlewait-helper)

**Calls that return None or silently do nothing**
- [Add3 adds no equation and returns -1](#add3-adds-no-equation-and-returns--1)
- [FeatureCut / FeatureCut4 returns None](#featurecut--featurecut4-returns-none)
- [FeatureFillet3 returns None](#featurefillet3-returns-none)
- [MoveToFolder does nothing, or raises a type mismatch](#movetofolder-does-nothing-or-raises-a-type-mismatch)
- [AddHorizontalDimension2 returns None for every entity in one sketch](#addhorizontaldimension2-returns-none-for-every-entity-in-one-sketch)

**COM type and member errors**
- ['X' object is not callable, or 'Member not found' (-2147352573)](#x-object-is-not-callable-or-member-not-found--2147352573)
- [DISP_E_TYPEMISMATCH / 'Type mismatch' on a call with an optional Object](#disp_e_typemismatch--type-mismatch-on-a-call-with-an-optional-object)
- ['Element not found' from EnsureDispatch or CastTo](#element-not-found-from-ensuredispatch-or-castto)
- ['Unable to read write-only property' from GetCoords()](#unable-to-read-write-only-property-from-getcoords)
- [SelectByID2 failed for a folder](#selectbyid2-failed-for-a-folder)

**Session and connection**
- [GetActiveObject raises 'Operation unavailable' but SLDWORKS.exe is running](#getactiveobject-raises-operation-unavailable-but-sldworksexe-is-running)
- [A script created a sketch in the user's real document](#a-script-created-a-sketch-in-the-users-real-document)
- [GetDocuments raises instead of returning an empty list](#getdocuments-raises-instead-of-returning-an-empty-list)
- [ActivateDoc2 raises 'Type mismatch'](#activatedoc2-raises-type-mismatch)

**Wrong geometry, no error**
- [A hole is in the wrong place but the sketch is fully defined and rebuilds clean](#a-hole-is-in-the-wrong-place-but-the-sketch-is-fully-defined-and-rebuilds-clean)
- [AddDimension2 gave one diagonal dimension instead of H and V](#adddimension2-gave-one-diagonal-dimension-instead-of-h-and-v)
- [ORIGIN selection picked up a dimension instead of the origin](#origin-selection-picked-up-a-dimension-instead-of-the-origin)
- [An enum constant has an unexpected value](#an-enum-constant-has-an-unexpected-value)
- [rms_check.py reports SKIP for detail.holes_last](#rms_checkpy-reports-skip-for-detailholes_last)

---

## A dimension call hangs forever but SolidWorks is Responding

**Symptom.** `AddDimension2`, `AddHorizontalDimension2` or `AddVerticalDimension2`
never returns. Checked from outside (`Get-Process`, `MainWindowTitle`),
SolidWorks reports `Responding = True` the whole time.

**Cause.** The System Option **"Input dimension value"**
(`swInputDimValOnCreate`, value `10` in `swUserPreferenceToggle_e`) pops a modal
*Modify* dialog after every API call that creates a dimension. The script is not
hung — it is a normal modal wait, with the dialog possibly behind another window.

**Fix.** Disable it at the source rather than working around each hang:

```python
with no_input_dim_dialog(sw):       # scripts/sw_helpers.py
    doc.AddHorizontalDimension2(x, y, 0)
```

This is an App-level global preference, not document-scoped, so restore it
afterwards for the user's normal interactive work — the context manager does.
`python scripts/sw_preflight.py --fix` checks and clears it before a build.

**If you find one already hung: do not kill the client.** Dismiss the dialog
properly first (`EnumWindows` + `SetForegroundWindow` + `SendKeys("{ENTER}")`,
or `PostMessage(hwnd, 0x0010, 0, 0)` where Cancel is acceptable). Killing it
leads directly to the next entry.

*Verified on: SW2026 SP1.1. This root cause supersedes the earlier
"COM defect" reading of the rectangle-edge hang below.*

## SolidWorks crashes on Save() with RPC -2147023170

**Symptom.** `Save()` or `SaveAs()` crashes the process outright — RPC failure
`-2147023170`, full app exit with a SOLIDWORKS Error Report dialog — with no
other operation in between.

**Cause.** The Python client was killed while a modal *Modify* dialog was open
(see above). The session is left in a state where the next save crashes.
Reproduced 3 times, isolated by testing `Save()` immediately after the kill with
no intervening feature creation.

**Fix.** Prevention only: keep `swInputDimValOnCreate` off so the modal never
appears, and dismiss rather than kill if one does. Once a session is in this
state, restart SolidWorks.

*Verified on: SW2026 SP1.1.*

## Dimensioning a rectangle edge hangs or crashes SolidWorks

**Symptom.** Selecting one of the 4 real edges from `CreateCenterRectangle` via
`edge.Select4(False, none_dispatch())` and calling `doc.AddDimension2(x, y, z)`
hangs indefinitely, or crashes the process (RPC `-2147023170`). Happened on the
first such call in a freshly created document, so it is not accumulated-state
clutter.

**Cause.** Partly the modal dialog above. The hang is fully explained by it;
whether a genuine defect in dimensioning rectangle-tool line entities also
exists was never isolated, because the corner-point recipe below removed the
need.

**Fix.** Dimension the rectangle's **corner point against the origin** instead of
an edge's length. One dimension fully constrains a symmetric center rectangle —
see
[api-recipes.md](api-recipes.md#fully-defining-a-center-rectangle-with-one-dimension).

The corner-point-plus-origin approach also worked reliably for a freestanding
`CreatePoint` and for a circle's center.

*Verified on: SW2026 SP1.1. Partially superseded — try the recipe with
`no_input_dim_dialog` before assuming a defect.*

## A long script crashes SolidWorks partway through

**Symptom.** A single script that creates and dimensions several
sketches/features in one continuous COM session crashes SolidWorks partway
through. Short, single-purpose scripts run as separate processes against the
same live session never crash — even exercising the exact same calls.

**Cause.** Unresolved. The pattern points at sustained rapid-fire API pressure
within one continuous automation run rather than any single call.

**Fix.** Remove any message pumping (see next entry). If it persists, split the
build into several smaller scripts run as separate processes with a real pause
between them, or prompt the user to restart SolidWorks between stages.

*Verified on: SW2026 SP1.1. Root cause not found.*

## The script deadlocks after adding a settle/wait helper

**Symptom.** A helper that calls `pythoncom.PumpWaitingMessages()` a few times
plus a short sleep after every UI-mutating call makes a *different* call hang
indefinitely. SolidWorks stays `Responding = True`; the Python **client** thread
is the one stuck.

**Cause.** COM re-entrancy triggered by pumping messages while SolidWorks is
still mid-operation.

**Fix.** **Do not call `PumpWaitingMessages()` as a generic "let the UI catch up"
measure.** A plain `time.sleep()` is slower but cannot deadlock the caller.
`sw_helpers.settle()` is deliberately just a sleep.

*Verified on: SW2026 SP1.1.*

---

## Add3 adds no equation and returns -1

**Symptom.** `EquationMgr.Add3(...)` returns `-1`, raises nothing, and adds no
equation. `GetCount` is unchanged.

**Cause.** `Add3` is the documented primary call, with `Add2`/`Add` framed as
fallbacks for older releases — but on this build `Add3` silently no-ops.
Confirmed reproducible twice in separate scratch documents, independent of
`ConfigurationOption` (`swAllConfiguration=1` and `swThisConfiguration=2` both
no-op).

**Fix.** Use `Add2(index, equation_text, use_automatic_solve_order)`, and verify
the count moved. `sw_helpers.add_equation` does both.

If you must call `Add3`, its trailing `ConfigurationName` `Object` parameter
rejects `None`, `""` and `[]` alike; only
`VARIANT(pythoncom.VT_EMPTY, None)` gets the call to execute — at which point it
still returns `-1`. `GetCount`, `Equation(i)` and `Status(i)` verify what
actually landed regardless of which `Add*` you used.

*Verified on: SW2026 SP1.1 (rev 34.1.1.11). Untested on other versions — treat
"Add3 is broken on 2026" as confirmed and "Add3 is broken in general" as
unverified.*

## FeatureCut / FeatureCut4 returns None

**Symptom.** A plain sketch-based cut returns `None` unconditionally, across
every combination of end condition (`Blind` with oversized depth, `ThroughAll`,
`ThroughAllBoth`), `Flip`, and `NormalCut`. The same cut works in the GUI.

**Cause.** The `Dir` argument. Documented only as a direction-flip flag, and
easy to assume defaults fine at `False` by analogy with `FeatureExtrusion3`
(where `Dir=False` works).

**Fix.** Set **`Dir=True`**. Confirmed by a parameter sweep
(`itertools.product` over `Flip`, `Dir`, `NormalCut`, `UseFeatScope`/
`UseAutoSelect`): every `Dir=False` combination returned `None`; every
`Dir=True` combination with `UseFeatScope=True, UseAutoSelect=True` succeeded.

*Verified on: SW2026 SP1.1.*

## FeatureFillet3 returns None

**Symptom.** A basic constant-radius edge fillet returns `None` with
`Options=0, Ftyp=swFeatureFilletType_Simple(0)` — a reasonable-looking "no
special flags" default — and also with the recorded-macro-style `Options=195`.

**Cause.** `Options` needs the uniform-radius bit set.

**Fix.** `Options=2` (`swFeatureFilletUniformRadius`) with `Ftyp=0`. Found by
sweeping candidate `Options` values against the same fully-valid edge selection;
edge selection was verified separately as correct
(`GetSelectedObjectCount2(-1) == 4`), so it was not the problem here.

For selecting the edges reliably in the first place, see
[api-recipes.md](api-recipes.md#selecting-the-edges-reliably) — coordinate-guess
picks are flaky.

*Verified on: SW2026 SP1.1.*

## MoveToFolder does nothing, or raises a type mismatch

**Symptom.** `MoveToFolder(MoveToFeat, MoveFromFeat, IsFolder)` silently does
nothing when passed string names — no exception, feature stays outside the
folder — and raises `DISP_E_TYPEMISMATCH` when passed `IFeature` objects.

**Fix.** Do not spend time on it. To add a feature to an existing folder,
dissolve and re-wrap: `EditDelete` on an `"FTRFOLDER"`-type selection, then
re-select every member in tree order and call `InsertFeatureTreeFolder2(2)`
again. `sw_helpers.wrap_in_folder` handles the wrap half.

Better still, create each folder as its group completes — SolidWorks folders
must hold contiguous features and cannot reorder past a dependency.

*Verified on: SW2026 SP1.1.*

## AddHorizontalDimension2 returns None for every entity in one sketch

**Symptom.** After several failed dimension attempts against a sketch, the call
returns `None` for *every* entity in it — including a brand-new freestanding
`CreatePoint`. The identical call on a fresh sketch in the same document and
same session succeeds immediately. `GetConstrainedStatus` reads normally
(`2`, under-defined) and relation queries show nothing pathological.

**Cause.** Per-sketch state corruption. Not session-wide, not point-type
specific. No root cause found — this is a step beyond what the API surfaces.

**Fix.** Cheap and reliable: if a sketch has had more than one or two failed
`AddDimension2`/`AddHorizontalDimension2`/`AddVerticalDimension2` calls, stop
debugging it in place. **Delete the sketch** (and any partial equations
referencing its dimensions) and rebuild it fresh in a single continuous pass.

Every sketch built start-to-finish in one pass survived; the one repeatedly
re-entered across separate failed attempts never recovered.

*Verified on: SW2026 SP1.1.*

---

## 'X' object is not callable, or 'Member not found' (-2147352573)

**Symptom.** `doc.GetTitle()` raises `TypeError: 'str' object is not callable`,
or a COM call raises `-2147352573 'Member not found'`.

**Cause.** Zero-argument getters auto-invoke on bare attribute access under late
binding. The value is already resolved by the time you add `()`.

**Fix.** Drop the parentheses: `doc.GetTitle`, `feat.GetNextFeature`,
`call0(obj, name)`. **Do not write a `callable()`-based wrapper** — a resolved
`CDispatch` is itself callable, so the check cannot work. Full explanation in
[dispatch-quirks.md](dispatch-quirks.md#zero-argument-getters-auto-invoke--call-them-without-parentheses).

Known exceptions that *do* need parens: `body.GetFaces()`, `body.GetEdges()`.

*Verified on: SW2026 SP1.1.*

## DISP_E_TYPEMISMATCH / 'Type mismatch' on a call with an optional Object

**Symptom.** `SelectByID2` (or similar) raises `Type mismatch` when the
`Callout` argument is a bare Python `None`.

**Fix.** Wrap it: `VARIANT(pythoncom.VT_DISPATCH, None)`, i.e.
`sw_helpers.none_dispatch()`. For parameters that reject even that — notably
`Add3`'s `ConfigurationName` — use `VARIANT(pythoncom.VT_EMPTY, None)`.

*Verified on: SW2026 SP1.1.*

## 'Element not found' from EnsureDispatch or CastTo

**Symptom.** `gencache.EnsureDispatch("SldWorks.Application")` fails with
`Element not found` from `GetTypeInfo()`, on essentially every live object.

**Cause.** Early binding is broken against this build.

**Fix.** Use plain `win32com.client.Dispatch` for all actual calls. Early
binding is not a first-resort fix for a call returning `None` — suspect wrong
argument count/type or a misplaced method first.

To read a real signature, use `gencache.EnsureModule(clsid, 0, major, minor)`,
which succeeds where `EnsureDispatch` fails and writes a readable `.py` with
real argument order and count. See
[dispatch-quirks.md](dispatch-quirks.md#reading-a-real-signature-without-api-help).

*Verified on: SW2026 SP1.1 (rev 34.1.1).*

## 'Unable to read write-only property' from GetCoords()

**Symptom.** `ISketchPoint.GetCoords()` raises this via late-bound dispatch,
with or without parens.

**Fix.** Don't read sketch point coordinates back. Drive geometry by the
parameters you created it with, and dimension by selecting the entities
themselves (`Select4`), not by reading positions out.

*Verified on: SW2026 SP1.1.*

## SelectByID2 failed for a folder

**Symptom.** Selecting a feature-tree folder by name with `Type="BODYFEATURE"`
raises `SelectByID2 failed`.

**Fix.** Folders need `Type="FTRFOLDER"`. `"BODYFEATURE"` is correct for regular
features but not for folders. Full type-string table in
[dispatch-quirks.md](dispatch-quirks.md#selectbyid2-type-strings).

*Verified on: SW2026 SP1.1.*

---

## GetActiveObject raises 'Operation unavailable' but SLDWORKS.exe is running

**Symptom.** `win32com.client.GetActiveObject("SldWorks.Application")` raises
`-2147221021 'Operation unavailable'`, yet `Get-Process SLDWORKS` shows a
process that is `Responding`, with an empty `MainWindowTitle`.

**Cause.** A `Dispatch`-launched SolidWorks starts **invisible** and never
registers in the COM running object table. It holds a license seat and has no
window to close.

**Fix.** `Dispatch` *will* attach to it — which means the next script silently
does its work in an invisible session. Clean it up: attach with `Dispatch`,
confirm `GetDocuments` is empty and `Visible` is `False`, then call `ExitApp()`.

`python scripts/sw_preflight.py` detects exactly this state by cross-checking
the process list against COM, and refuses to proceed.

To avoid creating one: use `GetActiveObject` to attach-or-fail, and only
`Dispatch` when you deliberately intend to launch.

*Verified on: SW2026 SP1.1.*

## A script created a sketch in the user's real document

**Symptom.** An exploratory `InsertSketch` call landed a stray sketch in the
user's actual open document. In the same session, a later `CloseDoc` — using a
title captured earlier in the script — closed that same real, unsaved document.

**Cause.** `sw.NewPart()` returned the user's existing `ActiveDoc` instead of
creating a new part. A script talking to `SldWorks.Application` is talking to
the user's live session; there is no sandbox.

**Fix.** Three rules, in
[SKILL.md](../SKILL.md#working-safely-against-a-live-session), enforced by
`sw_helpers.assert_scratch_doc` and `safe_close`:

1. Never call a mutating or destructive method without re-fetching and
   re-checking the target document's identity **in the same breath as the call**.
   An earlier-captured title or handle is not evidence of what it still points at.
2. Never run exploratory probes against `sw.ActiveDoc` or an unverified
   "scratch" document.
3. Capture `open_titles(sw)` before creating anything, give scratch documents an
   identifying custom property, and assert the target is not a user document
   before every mutating call.

*Verified on: SW2026 SP1.1.*

## GetDocuments raises instead of returning an empty list

**Symptom.** Iterating `sw.GetDocuments` fails when no documents are open.

**Fix.** It returns `None`, not an empty tuple or collection. Guard with
`(call0(sw, "GetDocuments") or ())` — `sw_helpers.open_titles` does.

*Verified on: SW2026 SP1.1.*

## ActivateDoc2 raises 'Type mismatch'

**Symptom.** `sw.ActivateDoc2(name, useUserPreferences, errors)` raises `Type
mismatch` when `Errors` (a ByRef Long) is passed a plain `0`.

**Fix.** Avoid the call rather than fight the byref marshaling — it is
unnecessary when only one document is open.

*Verified on: SW2026 SP1.1.*

---

## A hole is in the wrong place but the sketch is fully defined and rebuilds clean

**Symptom.** Geometry is visibly wrong, yet the sketch reports fully defined
(status `3`), the feature rebuilds with zero errors, and nothing in the API
surfaces a problem.

**Cause.** For a full circle from `CreateCircleByRadius`, `GetStartPoint2` /
`GetEndPoint2` return a point on the circle's **boundary**, not its center. The
center is `GetCenterPoint2`. Adding a coincident relation to the origin using
the start point — by false analogy with the rectangle-corner recipe, where it is
correct — silently builds a circle tangent through the origin from one side.

The relation itself was correct (`GetRelationType` = `9`,
`swConstraintType_COINCIDENT`) — right relation kind, wrong point.

**Fix.** Use `GetCenterPoint2` for circle centers. When unsure whether a
"center to origin" fix actually took, **check face count on the resulting body**
(`body.GetFaces()` — this one needs the parens) against the expected count
(flat faces plus one cylindrical face per hole). Rebuild success alone does not
prove correct geometry.

This is the general lesson: a clean rebuild is necessary, not sufficient. There
is no automated geometry check in this repo yet.

*Verified on: SW2026 SP1.1.*

## AddDimension2 gave one diagonal dimension instead of H and V

**Symptom.** Selecting two points and calling `AddDimension2(x, y, z)` produces
a single aligned point-to-point distance, not independent horizontal and
vertical dimensions.

**Fix.** Call `AddHorizontalDimension2` and `AddVerticalDimension2` explicitly
(both take `(X, Y, Z)` placement args the same way). **Re-select both entities
before each call** — selection does not persist across them.

*Verified on: SW2026 SP1.1.*

## ORIGIN selection picked up a dimension instead of the origin

**Symptom.** `ext.SelectByID2("", "ORIGIN", 0, 0, 0, ...)` selects a *dimension*
(`swSelDIMENSIONS`, type `14`) rather than the origin point
(`swSelSKETCHPOINTS`, type `11`), and a selection-dependent call then returns
`None`.

**Cause.** `"ORIGIN"` selection is type/coordinate based, not name based. In a
sketch cluttered with overlapping prior test dimensions near the origin, it
picks the wrong thing.

**Fix.** Check `doc.SelectionManager.GetSelectedObjectType3(index, -1)` to
confirm what you got (`sw_helpers.selected_type`). In practice the fix was
working in a fresh, uncluttered sketch rather than reusing one across many
manual test iterations.

*Verified on: SW2026 SP1.1.*

## An enum constant has an unexpected value

**Symptom.** A remembered or guessed enum integer produces wrong behaviour or no
behaviour.

**Cause.** Enum values are not stable across versions and are easy to
misremember. Two confirmed here: `swDefaultTemplatePart` is **8** in
`swUserPreferenceStringValue_e`, not `4`; `swConstrainedStatus_e` is
**1**=unknown, **2**=under, **3**=fully, **4**=over, not the `1/2/3` assumed.

**Fix.** Resolve against the *installed* constants typelib via
`sw_helpers.constants_module()`, or the `gencache.EnsureModule` recipe in
[dispatch-quirks.md](dispatch-quirks.md#reading-a-real-signature-without-api-help).
Never hard-code an enum you have not verified on the target version.

*Verified on: SW2026 SP1.1.*

## rms_check.py reports SKIP for detail.holes_last

**Symptom.** An all-cut `4-Detail` folder produces `SKIP  detail.holes_last`
instead of PASS or FAIL.

**Cause.** Cuts created via `FeatureCut4` report `GetTypeName2 == "ICE"` on this
build, not `"Cut"`. `"ICE"` is now in the checker's `CUT_TYPES` and `HOLE_TYPES`,
but a differently-created cut may still land outside the classification sets.

**Fix.** Run `python rms_check.py --dump-types` against the part and add the
observed type strings to the sets at the top of the checker. Cosmetic only — it
does not cause a false FAIL.

*Verified on: SW2026 SP1.1.*
