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
eq = model.GetEquationMgr()
eq.Add2(-1, '"plate_width" = 120', True)          # global variable
eq.Add2(-1, '"D1@Sketch1" = "plate_width"', True)  # driven dimension
```

Index `-1` appends. **Try `Add2` first.** `Add3` is the documented primary call
and `Add2`/`Add` are framed as fallbacks for older releases, but on SW2026
(rev 34.1.1) `Add3` silently failed — returned `-1`, raised nothing, added no
equation — while `Add2(index, equation_text, use_automatic_solve_order)` worked
correctly. That session only tested one build, so treat this as "verify which
one actually lands an equation on the target version" rather than a blanket
"`Add3` is broken" — but don't assume newer-is-safer here. `eq.GetCount()`,
`eq.Equation(i)`, and `eq.Status(i)` are useful for verifying what actually
landed, regardless of which `Add*` call you used.

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

## Feature type strings

`GetTypeName2()` returns strings like `Extrusion`, `Cut`, `Revolution`, `Sweep`, `Loft`, `Fillet`, `Chamfer`, `Shell`, `Draft`, `LPattern`, `CirPattern`, `MirrorPattern`, `HoleWzd`, `ProfileFeature` (a sketch), `RefPlane`, `RefAxis`, `CoordSys`, `FtrFolder`.

These vary by version and by how the feature was created. Run `rms_check.py --dump-types` against known-good parts and adjust the classification sets at the top of the checker rather than trusting this list.
