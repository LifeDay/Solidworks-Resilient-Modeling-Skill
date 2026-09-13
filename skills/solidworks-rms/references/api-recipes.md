# Confirmed API recipes

Sequences that have been driven end-to-end against a live session with zero
rebuild errors. Each one encodes several non-obvious decisions; prefer copying a
recipe whole over reassembling it from API Help.

Verified on **SW2026 SP1.1 (rev 34.1.1)**. Signatures shift between releases —
re-confirm with the `gencache.EnsureModule` trick in
[dispatch-quirks.md](dispatch-quirks.md#reading-a-real-signature-without-api-help)
rather than trusting these across versions.

Every sample assumes the helpers. `<skill dir>` is this skill's base directory;
build scripts run from the user's project, so the path must be absolute:

```python
import sys; sys.path.insert(0, r"<skill dir>/scripts")
from sw_helpers import (connect, call0, select_by_id2, last_feature,
                        add_equation, wrap_in_folder, no_input_dim_dialog, mm,
                        none_dispatch, circular_edges)
```

## Before anything: preflight

```
python "<skill dir>/scripts/sw_preflight.py" --fix
```

Confirms bitness, an attachable SolidWorks, and — critically — that the
"Input dimension value" system option is off. With it on, every scripted
dimension call opens a modal dialog and the script appears to hang. See
[troubleshooting.md](troubleshooting.md#a-dimension-call-hangs-forever-but-solidworks-is-responding).

Its document list is what you pass as `known_user_titles`; nothing mutating
should run without `assert_scratch_doc`.

## Where sketch coordinates land

**`SketchManager.Create*(x, y, z, ...)` arguments are not global model
coordinates.** Measured by sketching a test circle at distinct coordinate
values, extruding it, and reading the resulting body's bounding box:

| Sketch plane | arg 1 → | arg 2 → | arg 3 | Plane sits at |
|---|---|---|---|---|
| Front Plane | global **Y** | global **Z** | discarded | X = 0 |
| Right Plane | global **−X** (negated) | global **Z** | discarded | Y = 0 |
| Top Plane | *not measured* | | | |

This is not the textbook "Front Plane = XY" layout, so it is very likely a
property of the part template in use on the test machine, not of the API. **Treat
the table as that install only.** Before placing geometry on a plane you have not
measured against your own template, repeat the probe in a scratch document: one
circle at distinct values such as `(0.011, 0.023, 0.037)`, a short extrude, then
compare the body's bounding-box centre with what you passed. Mixing up two axes
or missing the Right Plane sign flip produces a model that rebuilds clean with
features in the wrong place.

Model-space reads — `vertex.GetPoint`, `ICurve.CircleParams`, bounding boxes —
come back in **global** coordinates. Convert with this mapping before comparing
them against the values you sketched with.

## Base plate: sketch on a default plane, then extrude

The first slice confirmed end-to-end with zero rebuild errors.

```python
ext = doc.Extension
sm = doc.SketchManager
fm = doc.FeatureManager

# sketch a centered square on Front Plane
select_by_id2(ext, "Front Plane", "PLANE")
sm.InsertSketch(True)
sm.CreateCenterRectangle(0, 0, 0, half, half, 0)
sm.InsertSketch(True)                         # exit sketch
sk1 = last_feature(doc)                       # the creating call returns no handle

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

23 positional args, confirmed against `IFeatureManager::FeatureExtrusion3`'s
real signature via the generated typelib — not trial and error.

**`CreateCenterRectangle` returns 6 sketch segments, not 4**: the 4 profile
edges plus 2 construction diagonals used for the symmetric constraint. Return
order does not map to bottom/left/top/right. Filter on `.ConstructionGeometry`
first to drop the diagonals, then classify horizontal vs. vertical with
`.GetStartPoint2` / `.GetEndPoint2` (zero-arg, no parens).

## Fully defining a center rectangle with one dimension

Do **not** dimension a `CreateCenterRectangle` line segment directly — it hangs
or crashes the process (see
[troubleshooting.md](troubleshooting.md#dimensioning-a-rectangle-edge-hangs-or-crashes-solidworks)).
Dimension the rectangle's *corner point* against the *origin* instead:

```python
with no_input_dim_dialog(sw):
    corner_pt = edges[0].GetStartPoint2          # bare attribute - auto-invokes
    doc.ClearSelection2(True)
    corner_pt.Select4(False, none_dispatch())
    select_by_id2(ext, "", "ORIGIN", append=True)
    half_dim = doc.AddHorizontalDimension2(placement_x, placement_y, 0)
```

This **fully constrains** a symmetric about-origin center rectangle with a
*single* dimension — confirmed by editing the resulting dimension's
`SystemValue` and watching all 4 edge lengths scale together, so it is true
parametric control, not a DOF miscount. A second `AddVerticalDimension2` on the
same pair correctly returns `None` (SolidWorks recognising redundancy) rather
than hanging. Check `call0(sketch, "GetConstrainedStatus") == 3` to confirm
full definition rather than assuming a second dimension is needed.

The same corner-point-plus-origin approach works for a freestanding
`CreatePoint` and for a circle's center.

### Horizontal and vertical are separate calls

`AddDimension2` on two selected points creates a single **aligned**
point-to-point distance (unsigned straight-line), not independent H and V
dimensions. For those, call `AddHorizontalDimension2` and
`AddVerticalDimension2` explicitly — both take `(X, Y, Z)` placement args the
same way. **Re-select both entities before each call**; selection does not
persist across them.

### Circle centers

For a full circle from `CreateCircleByRadius`, `GetStartPoint2` / `GetEndPoint2`
return a point on the *boundary*, at the tool's parametric start angle. The
center is `GetCenterPoint2`. Getting this backwards silently builds a circle
tangent through the origin that still reports fully defined and rebuilds without
error — see
[troubleshooting.md](troubleshooting.md#a-hole-is-in-the-wrong-place-but-the-sketch-is-fully-defined-and-rebuilds-clean).

## Cuts

```python
select_by_id2(ext, sketch_name, "SKETCH")
cut = fm.FeatureCut4(
    True, False, True,        # Sd, Flip, Dir  <- Dir MUST be True
    ...                       # 27 positional args total
)
```

**`Dir=True` is the default to try first, not `False`.** A plain sketch-based
cut returned `None` unconditionally across every combination of end condition,
`Flip`, and `NormalCut` until `Dir` was set `True`. Confirmed by a full
parameter sweep: every `Dir=False` combination returned `None`; every `Dir=True`
combination with `UseFeatScope=True, UseAutoSelect=True` succeeded. This is easy
to get wrong by analogy with `FeatureExtrusion3`, where `Dir=False` works.

`FeatureCut4` takes 27 positional args: the same order as `FeatureExtrusion3`'s
23, plus `NormalCut`, `AssemblyFeatureScope`, `AutoSelectComponents`,
`PropagateFeatureToParts`, `OptimizeGeometry`.

Cut features created this way report `GetTypeName2 == "ICE"` on this build, not
`"Cut"`.

## Constant-radius fillet

```python
fillet = fm.FeatureFillet3(
    2,                        # Options: swFeatureFilletUniformRadius - NOT 0
    radius, 0, 0,             # R1, R2, Rho
    0,                        # Ftyp: swFeatureFilletType_Simple
    0, 0,                     # OverflowType, ConicRhoType
    None, None, None, None, None, None, None,   # 7 array params
)
```

**`Options=2`, not `0`.** A basic constant-radius edge fillet returned `None`
with the seemingly reasonable `Options=0, Ftyp=0`, and also with the
recorded-macro-style `Options=195`. It succeeded with `Options=2`
(`swFeatureFilletUniformRadius`, confirmed against the constants typelib).

`FeatureFillet3` takes 14 positional args, confirmed against the live install:

```
Options, R1, R2, Rho, Ftyp, OverflowType, ConicRhoType,
Radii, Dist2Arr, RhoArr, SetBackDistances,
PointRadiusArray, PointDist2Array, PointRhoArray
```

A plain constant-radius fillet passes `None` for the last 7 (the array params).

### Selecting the edges reliably

Picking edges by coordinate guess (`SelectByID2("", "EDGE", x, y, z, ...)`) was
flaky: 3 of 4 picks returned `True` while the actual selected count was 2, and
the 4th returned `False`. Iterate the body's edges and filter instead — 4-for-4
reliable:

```python
for edge in body.GetEdges():          # note the parens; this one does not auto-invoke
    if not edge.GetCurve.IsLine:      # bare attribute
        continue
    start = edge.GetStartVertex().GetPoint
    ...                               # match expected corner + edge length
    edge.Select4(True, none_dispatch())
```

`vertex.GetPoint` returns a plain Python tuple already in model `(x, y, z)`
order. **Check which axis is the depth direction for the plane you actually
used** — sketching on Front Plane put the extrude direction at tuple index 0,
with the two in-sketch-plane coordinates at indices 1 and 2, i.e.
`(depth, sketch_y, sketch_z)`, not the `(x, y, z)` naively expected from the
sketch's own 2D system. That matches the measured plane mapping in
[where sketch coordinates land](#where-sketch-coordinates-land). Dump raw
vertex tuples for known edges before writing the filter rather than assuming
an axis order.

### Circular edges

The same lesson held for circular fillet edges: coordinate-guess `SelectByID2`
picks were unreliable, but matching each circular edge's centre and radius
worked:

```python
doc.ClearSelection2(True)
for edge in circular_edges(body, centre_xyz, radius):   # global coords, metres
    edge.Select4(True, none_dispatch())
```

`circular_edges` reads `ICurve.CircleParams`. That is a bare property returning
`(cx, cy, cz, nx, ny, nz, r)`; there is no `GetCircleParams()` method. Compute
`centre_xyz` from the parameter set *through the plane mapping*, not from the raw
sketch arguments. Confirm with `selected_count(doc)` before calling
`FeatureFillet3`.

## Sweep paths: chain arcs with CreateTangentArc

**Build a multi-segment path with `CreateTangentArc`, not `Create3PointArc`.**
A `Create3PointArc` next to a separately created `CreateLine` is not
topologically connected, even at numerically identical endpoints, and
`InsertProtrusionSwept4` returns `None` for the path. See
[troubleshooting.md](troubleshooting.md#insertprotrusionswept4-returns-none-for-a-line-plus-arc-path).
`CreateTangentArc` continues from the previous entity's *actual* endpoint, and
the identical sweep succeeded immediately.

```python
select_by_id2(ext, "Right Plane", "PLANE")
sm.InsertSketch(True)
sm.CreateLine(x0, y0, 0, x1, y1, 0)
sm.CreateTangentArc(x1, y1, 0, x2, y2, 0, arc_type)   # (start, end, swTangentArcTypes_e)
sm.InsertSketch(True)
path = last_feature(doc)
```

Signatures from the SW2026 typelib:

- `ISketchManager.CreateTangentArc(X1, Y1, Z1, X2, Y2, Z2, ArcType)`.
  `swTangentArcTypes_e`: `swForward=1`, `swLeft=2`, `swBack=3`, `swRight=4`.
  The session did not record which value it used. Resolve through
  `constants_module()` rather than hard-coding.
- `IFeatureManager.InsertProtrusionSwept4` takes **20** positional args:

  ```
  Propagate, Alignment, TwistCtrlOption, KeepTangency, BAdvancedSmoothing,
  StartMatchingType, EndMatchingType, IsThinBody, Thickness1, Thickness2,
  ThinType, PathAlign, Merge, UseFeatScope, UseAutoSelect, TwistAngle,
  BMergeSmoothFaces, CircularProfile, CircularProfileDiameter, Direction
  ```

  With `CircularProfile=True` and a diameter, no profile sketch is needed. Per
  API Help, only the path gets selected; the selection marks the session used
  were not recorded. An explicit profile sketch plus
  path sketch also works, but did not fix the unconnected-path failure.

What was confirmed: arc-only, line-only, line+line (sharp corner) and
line → `CreateTangentArc` paths all swept. Not tested: a `CreateLine` placed
*after* an arc. If one returns `None`, suspect the same connectivity problem.
Also watch for
[exactly horizontal `CreateLine` calls](troubleshooting.md#createline-returns-none-for-a-horizontal-line),
which return `None` outright.

## Saving, opening and screenshots

```python
import os
doc.Save()                                              # already-named document, in place
doc = sw.OpenDoc(os.path.abspath(path), doc_type)       # swDocumentTypes_e; returns doc or None
doc.SaveBMP(os.path.abspath("view.bmp"), 1280, 720)     # relative path -> False, silently
```

- **`OpenDoc`, not `OpenDoc6`.** `OpenDoc6`'s `ByRef Long` `Errors`/`Warnings`
  raise `Type mismatch` with `None` or `[0]`. The same applies to the 6-arg
  `SaveAs`; no working SaveAs-to-new-name form is recorded yet. See
  [troubleshooting.md](troubleshooting.md#saveas-or-opendoc6-raises-type-mismatch).
- **`SaveBMP` needs an absolute path.** A relative one returns `False` with no
  file written.
- A document opened this way is subject to the same identity rules as any
  other. If the file is already open in the user's session, assume `OpenDoc`
  hands back *their* document; this is unconfirmed. Run `assert_scratch_doc`
  before mutating it.

## Global variables and driven dimensions

```python
eq = call0(doc, "GetEquationMgr")
add_equation(eq, '"plate_width" = 120')            # global variable
add_equation(eq, '"D1@Sketch1" = "plate_width"')   # driven dimension
```

`add_equation` uses `Add2(index, text, use_automatic_solve_order)` with index
`-1` to append, then verifies `GetCount` actually incremented — because `Add3`
fails *silently* on this build. See
[troubleshooting.md](troubleshooting.md#add3-adds-no-equation-and-returns--1).

A dimension driven by an equation reports as driven; one set through
`model.Parameter("D1@Sketch1").SystemValue = 0.12` does not. `rms_check.py` uses
that distinction to measure how much of the model is genuinely parameterized.

## Wrapping a group's features in a folder

```python
wrap_in_folder(doc, ["Boss-Extrude1", "Fillet1"], "3-Core")
wrap_in_folder(doc, ["Sketch1"], "1-Ref", types={"Sketch1": "SKETCH"})
```

`InsertFeatureTreeFolder2(2)` — type `2` is "Containing", which wraps the
*current selection* rather than creating an empty folder. Select the group's
features first (`"BODYFEATURE"` for features, `"SKETCH"` for sketches), then
wrap and rename.

Create each folder **as its group completes**. SolidWorks folders must hold
contiguous features and will not reorder past a dependency, so there is no
sort-the-tree-afterwards path on a part of any realistic complexity.

Flat traversal surfaces a synthetic `"<FolderName>___EndTag___"` marker feature
immediately after a folder's contents. It is itself `FtrFolder`-typed, so
type-based filters skip it naturally — do not special-case it away.

**To add a feature to an existing folder**, dissolve and re-wrap: `EditDelete`
on an `"FTRFOLDER"`-type selection, then re-select every member in tree order
and call `InsertFeatureTreeFolder2(2)` again. `MoveToFolder` does not work here
— see [troubleshooting.md](troubleshooting.md#movetofolder-does-nothing-or-raises-a-type-mismatch).

## Traversal

```python
feat = call0(doc, "FirstFeature")
while feat:
    name = feat.Name
    type_name = call0(feat, "GetTypeName2")
    feat = call0(feat, "GetNextFeature")
```

`sw_helpers.iter_features(doc)`, `last_feature(doc)` and
`feature_by_name(doc, name)` wrap this.

Folder contents are reached through `GetFirstSubFeature` / `GetNextSubFeature`
on the folder feature, or by continuing the flat traversal — behaviour differs
by version, so handle both.

### Parents and children

- `Feature.GetChildren` returns dependent features.
- `Feature.GetParents` returns what it depends on. Not present in every release;
  degrade to building the graph from `GetChildren` in reverse if it returns
  `None`.

These are the basis for the cross-group reference check: a feature in group N
whose parent sits in group > N is a violation.

## Feature descriptions

`Feature.Description` is a read/write property; empty string means none. This is
where design intent lives under this skill's naming policy.

Tell the user to enable **Tree Display → Show Feature Descriptions**, since the
intent is invisible otherwise.

## Suppression testing

```python
feat.SetSuppression2(SUPPRESS, THIS_CONFIG, None)
doc.ForceRebuild3(False)
errors = call0(doc.Extension, "GetWhatsWrongCount")
feat.SetSuppression2(UNSUPPRESS, THIS_CONFIG, None)
```

`swSuppressFeature` / `swUnSuppressFeature` come from
`swFeatureSuppressionAction_e`; configuration scope from
`swInConfigurationOpts_e`. **Always restore suppression state, including on
exception**, or the check leaves the model damaged.

Rebuild errors read through `ModelDocExtension.GetWhatsWrongCount` and
`GetWhatsWrong`, or per-feature through `Feature.GetErrorCode2()`.

## Sketch constraint status

```python
sketch = call0(feat, "GetSpecificFeature2")
status = call0(sketch, "GetConstrainedStatus")
```

`swConstrainedStatus_e` on this build: `1` = unknown, `2` = under-defined,
`3` = fully defined, `4` = over-defined (`5`–`7` are solver-error states).
Resolve against the constants typelib rather than hard-coding — these were
*not* the `1/2/3` originally assumed.

## Feature type strings

`GetTypeName2` returns strings like `Extrusion`, `Cut`, `ICE`, `Revolution`,
`Sweep`, `Loft`, `Fillet`, `Chamfer`, `Shell`, `Draft`, `LPattern`,
`CirPattern`, `MirrorPattern`, `HoleWzd`, `ProfileFeature` (a sketch),
`RefPlane`, `RefAxis`, `CoordSys`, `FtrFolder`.

These vary by version and by how the feature was created. Run
`python "<skill dir>/rms_check.py" --dump-types` against known-good parts and adjust the
classification sets at the top of the checker rather than trusting this list.
