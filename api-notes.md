# SolidWorks API notes

The calls this skill depends on, with the caveat that signatures and enum values change between releases. **Verify against the local SolidWorks API Help before relying on any of these.** Where a call takes more than a couple of arguments, check the argument order rather than assuming.

## Connecting (pywin32) — and its dynamic-dispatch quirks

```python
import win32com.client

sw = win32com.client.Dispatch("SldWorks.Application")
sw.Visible = True
model = sw.ActiveDoc
```

`Dispatch` gives late binding. Late binding means every call is resolved by name
at call time rather than against a known signature, and that shows up as real
behavioral quirks, not just a performance difference:

- **Zero-argument getter-style methods auto-invoke on attribute access and must
  be called without parentheses.** `doc.FirstFeature`, `feat.GetNextFeature`,
  `doc.GetEquationMgr`, `sketchSeg.GetType`, `sketchSeg.GetStartPoint2`,
  `doc.GetTitle`, `ext.GetWhatsWrongCount` — all of these are properties from
  pywin32's point of view, even though API Help documents them as methods.
  Calling any of them with `()` raises `TypeError: 'X' object is not callable`
  because the value is already resolved by the time you try to call it.
- **Multi-arg methods with an Object/IDispatch parameter reject a bare Python
  `None`** with `DISP_E_TYPEMISMATCH` ("Type mismatch"). Wrap it instead:
  `win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)`. Confirmed fix for
  `SelectByID2`'s `Callout` parameter; no other argument there needed explicit
  VARIANT typing.
- **Early binding is a known-bad path on SW2026 (rev 34.1.1), not a first-resort
  fix.** `win32com.client.gencache.EnsureDispatch("SldWorks.Application")` (and
  `CastTo`, which calls it internally) failed with `Element not found` from
  `GetTypeInfo()` on essentially every live object tried in that session. Stick
  with plain `win32com.client.Dispatch` for actual calls. If a call returns
  `None` unexpectedly, suspect a wrong argument count/type or a misplaced method
  (see below) before reaching for early binding.
- **To find a method's real signature without API Help access**, generate the
  typelib module directly by CLSID rather than through `EnsureDispatch`:
  ```python
  win32com.client.gencache.EnsureModule(clsid, 0, major, minor)
  ```
  This succeeds even when `EnsureDispatch` fails, and writes a readable `.py`
  into `gen_py` with real method names, argument order, and argument count — it
  just can't itself be used for dispatch. Find the CLSID with
  `win32com.client.selecttlb.EnumTlbs()`.
  - SW2026 main typelib ("SldWorks 2026 Type Library"): CLSID
    `{83A33D31-27C5-11CE-BFD4-00400513BB57}`, major/minor `34/0`.
  - Enum constants (`swEndConditions_e`, `swSelectType_e`,
    `swFeatureSuppressionAction_e`, `swConstrainedStatus_e`,
    `swInConfigurationOpts_e`, `swFeatureTreeFolderType_e`, ...) live in a
    **separate** typelib, "SOLIDWORKS 2026 Constant type library", CLSID
    `{4687F359-55D0-4CD3-B6CF-2EB42C11F989}`. Same `EnsureModule` trick works
    there.

All lengths in the API are **metres**, regardless of document units.

## Folders

- `FeatureManager.InsertFeatureTreeFolder2(type)` — creates a folder. The type enum (`swFeatureTreeFolderType_e`) distinguishes an empty folder from one containing the current selection.
- `FeatureManager.MoveToFolder(folderName, featureName, moveOnlySelected)` — moves a feature into a folder.
- Folders appear in tree traversal with `GetTypeName2()` returning `"FtrFolder"`.

Because folders require contiguous features and SolidWorks will not reorder past a dependency, create the folder when the group starts and add features as they are built. Attempting to sort a finished tree into folders will fail on any part of realistic complexity.

### Confirmed recipe: wrapping a group's features in a folder

`InsertFeatureTreeFolder2(2)` — type `2` is the "Containing" folder, which wraps
the *current selection* rather than creating an empty folder. Select the group's
features first (`Type="SKETCH"` or `Type="BODYFEATURE"` as appropriate — see the
feature-selection note below for why `"BODYFEATURE"` and not `GetTypeName2()`'s
value), then wrap and rename:

```python
# select_by_id2 selects with Callout wrapped as VARIANT(VT_DISPATCH, None)
select_by_id2(ext, "Boss-Extrude1", "BODYFEATURE", append=False)
select_by_id2(ext, "Fillet1", "BODYFEATURE", append=True)

folder = fm.InsertFeatureTreeFolder2(2)   # 2 = swFeatureTreeFolderType_Containing
folder.Name = "3-Core"
```

Confirmed end-to-end against SW2026. Caveat: flat tree traversal
(`FirstFeature`/`GetNextFeature`) surfaces a synthetic
`"<FolderName>___EndTag___"` marker feature immediately after a folder's
contents — see the comment in `rms_check.py`'s `walk()` for why this is
harmless and should stay that way.

## Global variables and equations

`EquationMgr` hangs off `ModelDoc2`:

```python
eq = model.GetEquationMgr  # zero-arg getter — no parens, see dynamic-dispatch note above
eq.Add2(-1, '"plate_width" = 120', True)          # global variable
eq.Add2(-1, '"D1@Sketch1" = "plate_width"', True)  # driven dimension
```

Index `-1` appends. **Use `Add2`, not `Add3`.** `Add3` is the documented primary
call and `Add2`/`Add` are framed as fallbacks for older releases, but on SW2026
(rev 34.1.1.11, SP1.1) `Add3` silently fails — returns `-1`, raises nothing,
adds no equation — while `Add2(index, equation_text, use_automatic_solve_order)`
works correctly. Confirmed reproducible twice, in separate scratch documents,
independent of `ConfigurationOption` (`swAllConfiguration=1` and
`swThisConfiguration=2` both silently no-op). This still hasn't been tested
against any other SolidWorks version — no second install was available — so
treat "`Add3` is broken on 2026" as confirmed and "`Add3` is broken in general"
as unverified; try `Add2` first regardless and check `eq.GetCount` to confirm.

If you do need `Add3`'s signature
(`Add3(Index, Equation, ConfigurationOption, ConfigurationName)`), its trailing
`ConfigurationName` `Object` parameter is pickier than the `VT_DISPATCH` fix
documented above for `SelectByID2`: a bare `None`, `""`, or `[]` all raise
`DISP_E_TYPEMISMATCH` here. Use `win32com.client.VARIANT(pythoncom.VT_EMPTY, None)`
to get the call to actually execute — at which point, on this build, it still
just returns `-1`. `eq.GetCount`, `eq.Equation(i)`, and `eq.Status(i)` are
useful for verifying what actually landed, regardless of which `Add*` call you
used.

A dimension driven by an equation reports as driven; one set through `model.Parameter("D1@Sketch1").SystemValue = 0.12` does not. The checker uses this distinction to measure how much of the model is genuinely parameterized.

## Feature descriptions

`Feature.Description` is a read/write property. Empty string means no description. This is where design intent lives under this skill's naming policy.

The user must enable Tree Display → Show Feature Descriptions for it to appear in the FeatureManager.

## Traversal

```python
feat = model.FirstFeature()
while feat:
    name = feat.Name
    type_name = feat.GetTypeName2()
    feat = feat.GetNextFeature()
```

Folder contents are reached through `GetFirstSubFeature()` / `GetNextSubFeature()` on the folder feature, or by continuing the flat traversal — behaviour differs by version, so handle both.

## Parent and child relationships

- `Feature.GetChildren()` returns the dependent features.
- `Feature.GetParents()` returns what it depends on. Not present in every release; degrade to building the graph from `GetChildren()` in reverse if it returns `None`.

These are the basis for the cross-group reference check. A feature in group N whose parent sits in group > N is a violation.

## Suppression testing

```python
feat.SetSuppression2(swSuppressFeature, swThisConfiguration, None)
model.ForceRebuild3(False)
# inspect errors
feat.SetSuppression2(swUnSuppressFeature, swThisConfiguration, None)
```

`swSuppressFeature` and `swUnSuppressFeature` come from `swFeatureSuppressionAction_e`; configuration scope from `swInConfigurationOpts_e`. Always restore suppression state, including on exception, or the check leaves the model damaged.

Rebuild errors are read through `ModelDocExtension.GetWhatsWrongCount()` and `GetWhatsWrong()`, or per-feature through `Feature.GetErrorCode2()`.

## Sketch constraint status

```python
sketch = feat.GetSpecificFeature2()
status = sketch.GetConstrainedStatus()
```

Compare against the named values in `swConstrainedStatus_e` rather than hard-coded integers — under-defined, fully defined, over-defined, and the various invalid states are all distinguished there.

## Method-location and naming surprises

None of these are guessable from where sibling calls live — each "looks right"
on pattern-matching from nearby API calls and is wrong:

- `InsertSketch2` does not exist via dynamic dispatch on `SketchManager` on this
  build — use `InsertSketch(bool)` instead (same effect: toggles sketch edit
  mode).
- `InsertAxis2(AutoSize)` lives on `IModelDoc2` (the document), **not** on
  `FeatureManager`, despite every other feature-creation call living on
  `FeatureManager`.
- `AddDimension2(x, y, z)` lives on `IModelDoc2`, **not** on
  `ModelDocExtension` — easy to guess wrong since `SelectByID2` and
  `GetWhatsWrongCount` both live on `Extension`.
- To select a *feature* (not a face/edge/plane/sketch) by name via
  `SelectByID2` — needed for a folder-wrap or pattern-seed selection — the
  `Type` string must be `"BODYFEATURE"`. Not `GetTypeName2()`'s value (e.g.
  `"Extrusion"`), not `"FEAT"`, not `""`.
- `SketchManager.CreateCenterRectangle` returns **6** sketch segments, not 4:
  the 4 real profile edges plus 2 construction diagonals (corner-to-corner,
  used internally for the symmetric constraint). Don't assume return order
  maps to bottom/left/top/right. Filter on `.ConstructionGeometry` (bool) first
  to drop the diagonals, then classify horizontal vs. vertical with
  `.GetStartPoint2` / `.GetEndPoint2` (zero-arg, no parens).

## Confirmed recipe: base plate (sketch on a default plane → boss extrude)

The one full slice confirmed end-to-end with zero rebuild errors against SW2026:

```python
ext = doc.Extension
sm = doc.SketchManager
fm = doc.FeatureManager

# sketch a centered square on Front Plane
select_by_id2(ext, "Front Plane", "PLANE")   # Callout = VARIANT(VT_DISPATCH, None)
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

23 positional args, confirmed against `IFeatureManager::FeatureExtrusion3`'s real
signature via the `gencache.EnsureModule`-generated file (see Connecting, above)
— not just trial and error.

## Session log: NEMA34 mounting plate build (SW2026 SP1.1, Academic)

Findings from actually driving a build end-to-end, kept here because they
contradict or extend the notes above and cost two SolidWorks crashes to learn.

### The auto-invoke quirk is much broader than the four examples above

Every zero-arg `Get...`-style member tried on this build auto-invoked on
bare attribute access, not just the ones already listed: `GetTypeName2`,
`GetCount` (on `EquationMgr`), `GetConstrainedStatus`, `GetName` (on
`ISketchSegment`), `GetNameForSelection` (on `IDimension`), `GetDimension`
(on `IDisplayDimension`), `GetWhatsWrongCount`, `ActiveSketch`,
`GetSketchSegments`, `GetLength`, `EditRebuild3`. Treat this as the default,
not the exception: **try bare attribute access first for anything zero-arg,
and only add `()` if that raises `AttributeError`.**

The dangerous failure mode is a naive helper that checks `callable(value)`
to decide whether to also call it: a *resolved* `CDispatch` object is
itself callable (it implements `__call__` for the "invoke this COM method"
path), so `callable()` can't distinguish "still needs calling" from
"already resolved, calling it again is wrong". Calling an already-resolved
value raises either `TypeError: 'X' object is not callable` (plain Python
type) or the COM error `-2147352573 'Member not found'` (resolved
`CDispatch`) - both are the same underlying mistake. A correct zero-arg
helper is just `getattr(obj, name)` with **no** fallback call:

```python
def call0(obj, name):
    return getattr(obj, name)
```

### `ISketchPoint.GetCoords()` does not behave as documented

Calling it (with or without parens) raised `Unable to read write-only
property` via late-bound dispatch on this build. Don't use it to read a
sketch point's coordinates. There was no need found for reading them back
in practice - drive geometry by the parameters you created it with, and
dimension by selecting the entities themselves (`Select4`), not by
reading positions back out.

### Dimensioning a `CreateCenterRectangle` LINE segment directly is unsafe

Selecting one of the 4 real edges returned by `CreateCenterRectangle`
(via `edge.Select4(False, none_dispatch)`) and then calling
`doc.AddDimension2(x, y, z)` on it reproducibly **hung indefinitely** in
one run and **crashed the SolidWorks process outright** (RPC failure,
`-2147023170`, full app exit with a SOLIDWORKS Error Report dialog) in
another. Both happened on the very first such call in an otherwise-clean,
freshly created document, so this is not a clutter/accumulated-state
artifact - it looks like a genuine defect in this SW build's handling of
API-driven dimensioning of rectangle-tool line entities specifically.

**Workaround, confirmed working repeatedly:** dimension the rectangle's
*corner point* against the *origin* instead of dimensioning an edge's
length:

```python
corner_pt = edges[0].GetStartPoint2          # bare attribute - auto-invokes
doc.ClearSelection2(True)
corner_pt.Select4(False, none_dispatch)
ext.SelectByID2("", "ORIGIN", 0, 0, 0, True, 0, none_dispatch, 0)  # append
half_dim = doc.AddHorizontalDimension2(placement_x, placement_y, 0)
```

Empirically this **fully constrains** a symmetric (about-origin)
center-rectangle with a *single* dimension - confirmed by directly editing
the resulting dimension's `SystemValue` and observing all 4 edge lengths
scale together, so it is a true parametric square/rectangle control, not a
DOF miscount. A second `AddVerticalDimension2` call on the same corner
point + origin then correctly returns `None` (SolidWorks recognizing it
would be redundant) rather than hanging - use `sketch.GetConstrainedStatus
== 3` (see enum note below) to confirm full definition after the one
dimension rather than assuming a second is needed.

This same corner-point-plus-origin recipe, and the equivalent for a
freestanding `CreatePoint` and for a `CreateCircleByRadius` circle's
center, all worked reliably and repeatedly in isolation - the problem is
specific to dimensioning a rectangle tool's *line* entities directly.

### `pythoncom.PumpWaitingMessages()` inside a throttle helper is dangerous

An early fix for the hang above added a `settle()` helper that called
`pythoncom.PumpWaitingMessages()` a few times plus a short sleep after
every UI-mutating call. This made a *different* call
(`AddHorizontalDimension2` on the corner-point recipe above) hang
indefinitely on a later run, even though SolidWorks itself stayed
`Responding = True` the whole time (checked via `Get-Process` /
`MainWindowTitle` from outside) - the Python *client* thread was the one
stuck, apparently in COM re-entrancy triggered by pumping messages while
SolidWorks was still mid-operation. Removing the message-pump calls and
keeping only `time.sleep(seconds)` avoided this. **Do not call
`PumpWaitingMessages()` as a generic "let the UI catch up" measure against
a live SolidWorks session; a plain sleep is slower but doesn't risk
deadlocking the calling script.**

### `AddDimension2` on two points creates a point-to-point distance, not H+V

Selecting two points (e.g. a point and the origin) and calling
`doc.AddDimension2(x, y, z)` creates a single **aligned/point-to-point**
distance dimension (the straight-line distance, unsigned) - it does not
give you independent horizontal and vertical dimensions. For that, call
`AddHorizontalDimension2` and `AddVerticalDimension2` explicitly (both take
`(X, Y, Z)` placement args the same way `AddDimension2` does). Re-select
the same two entities before each call; do not assume selection persists
across the two dimension-adding calls.

### Selecting the sketch origin

```python
ext.SelectByID2("", "ORIGIN", 0, 0, 0, append, 0, none_dispatch, 0)
```

Empty name, `Type="ORIGIN"`. This is coordinate/type based, not name based
- in a sketch cluttered with many overlapping prior test dimensions near
the origin, this picked up a *dimension* (`swSelDIMENSIONS`, selection
type `14`) instead of the origin point (`swSelSKETCHPOINTS`, type `11`).
Check `doc.SelectionManager.GetSelectedObjectType3(index, -1)` if a
selection-dependent call mysteriously returns `None` in a sketch that has
accumulated a lot of test geometry; the fix in practice was just working
in a fresh, uncluttered sketch rather than reusing one across many manual
test iterations.

### Other confirmed odds and ends

- `sw.GetDocuments` returns `None`, not an empty tuple/collection, when no
  documents are open. Guard with `(sw.GetDocuments or ())`.
- `swDefaultTemplatePart = 8` in `swUserPreferenceStringValue_e` (not 4 -
  double check every enum value against the *installed* constants typelib
  rather than assuming a remembered value; see the `EnsureModule` trick
  above for `swUserPreferenceStringValue_e`'s home and every other enum).
- `sw.ActivateDoc2(name, useUserPreferences, errors)` raised `Type
  mismatch` when called with a plain `0` for the `Errors` (ByRef Long)
  argument on this build. Unneeded when there is only one open document;
  avoid it rather than fight the byref marshaling.
- Confirmed exact positional signatures against the live install (for
  future scripts, verify again the same way rather than trusting these
  across versions): `FeatureCut4` takes 27 positional args, in the same
  order as `FeatureExtrusion3`'s 23 plus `NormalCut`,
  `AssemblyFeatureScope`, `AutoSelectComponents`,
  `PropagateFeatureToParts`, `OptimizeGeometry`. `FeatureFillet3` takes 14:
  `Options, R1, R2, Rho, Ftyp, OverflowType, ConicRhoType, Radii,
  Dist2Arr, RhoArr, SetBackDistances, PointRadiusArray, PointDist2Array,
  PointRhoArray` - a plain constant-radius fillet passes `None` for the
  last 7 (array) params.
- **`rms_check.py`'s `swConstrainedStatus_e` values are wrong for this SW
  version** - see the dedicated note under Verification below.

### Session-stability observation (unresolved)

Both SolidWorks crashes happened partway through a *long* single script
run that had already created and dimensioned several sketches/features in
one continuous COM session; short, single-purpose test scripts (one script
= one or two API calls, run as separate Python processes against the same
live SW session) never crashed it, even when they exercised the exact same
calls that crashed inside the long script. This suggests the instability
is tied to sustained rapid-fire API pressure within one continuous
automation run rather than to any single call in isolation. No fix was
confirmed beyond removing the message-pumping (above); if a build script
keeps crashing SolidWorks partway through, consider splitting it into
several smaller scripts run as separate processes with a real pause (not
just `time.sleep` inside the same process) between them, or prompting the
user to restart SolidWorks between stages.

## Feature type strings

`GetTypeName2()` returns strings like `Extrusion`, `Cut`, `Revolution`, `Sweep`, `Loft`, `Fillet`, `Chamfer`, `Shell`, `Draft`, `LPattern`, `CirPattern`, `MirrorPattern`, `HoleWzd`, `ProfileFeature` (a sketch), `RefPlane`, `RefAxis`, `CoordSys`, `FtrFolder`.

These vary by version and by how the feature was created. Run `rms_check.py --dump-types` against known-good parts and adjust the classification sets at the top of the checker rather than trusting this list.

## Session log addendum: NEMA34 mounting plate, take 2 (SW2026 SP1.1, Academic)

A second full build of this same part, in a separate session, surfaced the
*root cause* behind several mysteries in the log above, plus new sharp edges.
Read this before repeating any of the workarounds above - some are now
obsolete.

### The dimension-add "hang" above was a modal dialog, not a COM defect

The System Options toggle **"Input dimension value" (`swInputDimValOnCreate`,
value `10` in `swUserPreferenceToggle_e`)** pops a modal "Modify" dialog box
after *every* API call that creates a new dimension
(`AddDimension2`, `AddHorizontalDimension2`, `AddVerticalDimension2`, ...) when
enabled - which explains the "hangs indefinitely, but SolidWorks itself stays
`Responding = True`" symptom from the earlier session as a normal modal wait,
not a defect. Fix at the source instead of working around each hang:

```python
sw.SetUserPreferenceToggle(10, False)   # disable before any scripted dimensioning
...
sw.SetUserPreferenceToggle(10, True)    # restore afterward - this is a global
                                          # App-level System Option, not scoped
                                          # to the document, so put it back for
                                          # the user's normal interactive use
```

Confirm the enum value against the live constants typelib rather than trusting
`10` blindly (`gencache.EnsureModule` on the constants CLSID, then
`mod.constants.swInputDimValOnCreate` - see the "Global variables and
equations" section above for the CLSID/EnsureModule recipe applied to
`swUserPreferenceToggle_e` members generally).

**Critically: killing the Python client while this modal sits open (instead of
dismissing it) leaves the live SolidWorks session in a state where the very
next `Save()` or `SaveAs()` call - even with no other operation in between -
crashes the process outright** (RPC failure, `-2147023170`, same signature as
the crash logged above). Reproduced 3 times this session, isolated by testing
`Save()` immediately after the kill with no intervening feature creation. If a
dimension call is ever found hung with this preference still on, do not
just kill-and-continue: find the "Modify" window and dismiss it properly first
(`EnumWindows` + `SetForegroundWindow` + `SendKeys("{ENTER}")` to accept, or
`PostMessage(hwnd, 0x0010, 0, 0)` / WM_CLOSE for a dialog where Cancel is
acceptable), then continue. With the preference disabled from the start,
this entire failure mode does not arise.

### `ISketchArc.GetCenterPoint2` vs `GetStartPoint2` - easy to get backwards

For a full circle (created via `CreateCircleByRadius`), `GetStartPoint2` /
`GetEndPoint2` return a point on the circle's boundary (at the tool's
parametric start angle), not the center. The center is a distinct accessor,
`GetCenterPoint2` (confirmed via `gencache.EnsureModule` listing
`ISketchArc`'s methods: `GetCenterPoint`, `GetCenterPoint2`, `GetEndPoint`,
`GetEndPoint2`, `GetStartPoint`, `GetStartPoint2`).

Using `GetStartPoint2` and adding a coincident relation to the origin (by
analogy with the rectangle-corner recipe above, which is correct for a
rectangle corner) silently builds a circle that is tangent through the
origin from one side, not centered on it - the sketch still reports fully
defined (status 3) and the feature rebuilds with zero errors, so nothing
surfaces the mistake short of actually checking the geometry. Confirmed by
reproducing it (center bore) and fixing it: relation type on the point,
read via `pt.GetRelations()[i].GetRelationType`, was `9` (`swConstraintType_
COINCIDENT` - confirmed correct relation kind, wrong point). Verify circle
center relations land on `GetCenterPoint2`, and when unsure whether a "center
to origin" fix actually worked, check face count on the resulting body
(`body.GetFaces()` - note the `()`, this one does not auto-invoke on bare
attribute access, unlike most zero-arg getters) against the expected count
(flat faces + one cylindrical face per hole) rather than trusting rebuild
success alone.

### A sketch that has absorbed several failed dimension-add attempts can become permanently unable to accept new dimensions

After `GetStartPoint2` (wrong point, see above) caused `AddHorizontalDimension2`
to return `None` repeatedly on a 4-circle sketch, switching to the correct
`GetCenterPoint2` point still returned `None` - for every circle in that
sketch, and even for a brand-new freestanding `CreatePoint` point added to
that same sketch. The identical call on a fresh scratch sketch (different
sketch, same document, same live SolidWorks session) succeeded immediately.
This narrows the fault to per-sketch state, not a session-wide or point-type
issue. `sketch.GetConstrainedStatus` still read normally (`2`, under-defined,
as expected), and relation queries on the existing entities showed nothing
pathological.

No root cause was found (this is a step beyond what the API surfaces), but
the practical fix is cheap: if a sketch has had more than one or two failed
`AddDimension2`/`AddHorizontalDimension2`/`AddVerticalDimension2` calls
against it, do not keep debugging it in place - delete the sketch (and any
partial equations referencing its dimensions) and rebuild it fresh in a single
continuous pass. Every sketch built start-to-finish in one pass this session
(no re-entry after `InsertSketch(True)` toggled it closed) worked on the first
or second try; the one sketch that was repeatedly re-entered across several
separate failed attempts never recovered.

### `FeatureCut` / `FeatureCut4`: the `Dir` argument is not optional-false

A plain sketch-based cut (`FeatureManager.FeatureCut` or `FeatureCut4`) on a
circle sketched exactly on the same plane as the extrude's start face
returned `None` unconditionally - across every combination of end condition
(`Blind` with an oversized depth, `ThroughAll`, `ThroughAllBoth`), `Flip`, and
`NormalCut` - until `Dir` was set `True`, at which point the identical call
(same end condition, same `Flip=False`) succeeded immediately. Confirmed by a
parameter sweep (`itertools.product` over `Flip`, `Dir`, `NormalCut`,
`UseFeatScope`/`UseAutoSelect` pairs) against a live cut attempt: every
`Dir=False` combination returned `None`; every `Dir=True` combination with
`UseFeatScope=True, UseAutoSelect=True` succeeded. `Dir` is documented only as
a direction-flip flag and is easy to assume defaults fine at `False` by
analogy with `FeatureExtrusion3` (where `Dir=False` works); for a cut on
this build, treat `Dir=True` as the default to try first, not `False`.

Resulting cut features report `GetTypeName2() == "ICE"` on this build/route
(not `"Cut"`) - `rms_check.py`'s `SOLID_TYPES` set already listed `"ICE"`
(apparently anticipated from an earlier calibration pass) but neither
`CUT_TYPES` nor `HOLE_TYPES` include it, so `detail.holes_last` reports SKIP
rather than PASS/FAIL for an all-cut 4-Detail folder on this build. Cosmetic
only - doesn't cause a false FAIL - but add `"ICE"` to `CUT_TYPES` (and to
`HOLE_TYPES` if the cuts are functionally holes) when calibrating the checker
against a part built this way.

`MoveToFolder(MoveToFeat, MoveFromFeat, IsFolder)` (folder name, feature name,
bool, per this build's real signature via `gencache`) rejected both a
string-name call (`SelectByID2`-friendly names for both args - silently did
nothing, feature stayed outside the folder despite no exception) and an
`IFeature`-object call (`DISP_E_TYPEMISMATCH`). Do not spend time on it -
the folder-recreation recipe already documented above (dissolve the folder
with `EditDelete` on a `"FTRFOLDER"`-type selection, then re-select every
member feature in tree order and call `InsertFeatureTreeFolder2(2)` again)
is the confirmed-reliable way to add a later feature to an existing named
folder, at the cost of recreating the wrapper each time.

Selecting a folder by name for this dissolve step needs
`Type="FTRFOLDER"` in `SelectByID2` - not `"BODYFEATURE"` (works for regular
features per the note above, but raises `SelectByID2 failed` for a folder).

### `FeatureFillet` / `FeatureFillet3`: `Options` needs the uniform-radius bit set, not `0`

A basic constant-radius edge fillet (4 vertical edges, single radius) returned
`None` from `FeatureFillet3` with `Options=0, Ftyp=swFeatureFilletType_Simple
(0)` - a seemingly reasonable "no special flags" default - and also with the
recorded-macro-style `Options=195` combined with `Ftyp=0`. It succeeded with
`Options=2` (`swFeatureFilletUniformRadius`, confirmed via the constants
typelib) and `Ftyp=0`. Found by sweeping a short list of candidate `Options`
values (`2, 195, 1, 128, 64`) against the same fully-valid edge selection -
edge selection was not the problem in this failure (verified separately: 4
edges selected correctly every time, `GetSelectedObjectCount2(-1) == 4`).
Treat `Options=2` as the default first try for a plain constant-radius round,
not `0`.

Selecting the target edges by a coordinate guess (`SelectByID2("", "EDGE", x,
y, z, ...)`) at each corner's nominal `(x, y)` was flaky - 3 of 4 picks
"succeeded" (`True`) but the actual post-hoc selected count was only 2, and
the 4th pick returned `False` outright. Switched to the reliable method:
iterate `body.GetEdges()`, filter `edge.GetCurve.IsLine` (bare attribute, no
parens) plus start/end-vertex coordinates matching the expected corner and
edge length, and `Select4(True, none_dispatch)` each match directly - 4-for-4
reliable. Vertex coordinates from `vertex.GetPoint` come back as a plain
Python tuple already in `(x, y, z)` model order (no unwrapping needed), but
double-check which axis is the sketch-plane pair and which is the
extrude/depth direction for the plane actually used - sketching on Front
Plane in this document put the extrude/depth direction in tuple index
0, with the two in-sketch-plane coordinates in indices 1 and 2 (i.e.
`(depth, sketch_y, sketch_z)`, not the `(x, y, z)` naively expected from the
sketch's own 2D coordinate system) - confirmed by dumping raw vertex tuples
for known edges before writing the filter, which is the right way to settle
this rather than assuming an axis order.
