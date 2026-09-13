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
from sw_helpers import (connect, open_titles, assert_scratch_doc, call0,
                        select_by_id2, select_origin, last_feature, iter_features,
                        add_equation, equations, wrap_in_folder,
                        no_input_dim_dialog, mm, deg, none_dispatch, circular_edges,
                        sketch_frame, model_to_sketch, sketch_to_model,
                        runs_horizontal, tag_doc, verify_tag, find_tagged_doc,
                        save_as, volume, body_extents)
```

Get `sw` from `connect()`, never from `win32com.client.Dispatch`. `connect()`
forces late binding. Attached any other way on a machine with a gen_py cache,
every bare getter below returns a bound method instead of a value; see
[dispatch-quirks.md](dispatch-quirks.md#a-gen_py-cache-silently-switches-to-early-binding).

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

## Creating the build document

```python
sw = connect()
known = open_titles(sw)                          # before creating anything
tmpl = sw.GetUserPreferenceStringValue(8)        # swDefaultTemplatePart
doc = sw.NewDocument(tmpl, 0, 0, 0)
assert_scratch_doc(doc, known)
tag_doc(doc, "rms_build", "bracket_v1")          # CustomPropertyManager("").Add3(key, 30, value, 2)
save_as(doc, r"C:\work\bracket.SLDPRT", known)   # SaveAs3(abspath, 0, 1); see Saving
```

A build split into stage scripts finds the document again by its tag, not by a
title remembered from an earlier run. It re-checks the tag immediately before
every mutating call:

```python
doc = find_tagged_doc(sw, "rms_build", "bracket_v1", known)
...
verify_tag(doc, known, "rms_build", "bracket_v1")   # in the same breath as the call
fm.FeatureExtrusion3(...)
```

Persist `known` between processes; a small JSON file next to the build scripts
worked. The user's titles captured before stage 1 are the ones that stay
untouchable.

## Where sketch coordinates land

**`SketchManager.Create*(x, y, z, ...)` arguments are sketch coordinates, not
global model coordinates.** Read the mapping from the open sketch rather than
assuming it:

```python
select_by_id2(ext, "Top Plane", "PLANE")
sm.InsertSketch(True)
frame = sketch_frame(doc)                               # ActiveSketch.ModelToSketchTransform.ArrayData
u, v, _ = model_to_sketch(frame, (mm(-100), 0, 0))      # raises if the point is off the plane
sm.CreateCircleByRadius(u, v, 0, mm(7.5))
placement = sketch_to_model(frame, u + mm(10), v)       # dimension placement points are model coords
```

`ArrayData` is 16 doubles: rotation `R` (3×3), translation `t`, scale `s`, then
3 unused. A model point `p` maps as `p' = s·(p·R) + t` with `p` as a row vector,
and the inverse is `p = ((p' − t) / s)·Rᵀ`. The helpers apply exactly that. It
placed geometry correctly on Right Plane, Top Plane and two offset planes that
do not pass through the origin, checked against bounding boxes and face
positions. The API's own route, `MathUtility.CreatePoint(...).MultiplyTransform`,
does not work under late binding; see
[troubleshooting.md](troubleshooting.md#mathutilitycreatepoint-raises-member-not-found).

For reference, the mapping on the test machine's part template:

| Sketch plane | arg 1 → | arg 2 → | arg 3 | Plane sits at |
|---|---|---|---|---|
| Front Plane | global **Y** | global **Z** | discarded | X = 0 |
| Right Plane | global **−X** (negated) | global **Z** | discarded | Y = 0 |
| Top Plane | global **Y** | global **−X** (negated) | discarded | Z = 0 |

Front and Top were measured by extruding a probe circle and reading its bounding
box. Right was measured the same way, and its transform gave
`R = [[-1,0,0],[0,0,1],[0,1,0]]`, which agrees. This is not the textbook
"Front Plane = XY" layout, so it is very likely a property of the template, not
of the API. The transform is the portable answer.

**Horizontal and vertical belong to the sketch, not the model.** "Both points
on global X" is `sgHORIZONTALPOINTS2D` only if global X runs horizontally in
that sketch, and on Top Plane it runs vertically. Getting it backwards dragged
the geometry, and the next dimension call returned `None`. Derive the choice:

```python
along_x = runs_horizontal(frame, (1, 0, 0))
relation = "sgHORIZONTALPOINTS2D" if along_x else "sgVERTICALPOINTS2D"
add_dim = doc.AddHorizontalDimension2 if along_x else doc.AddVerticalDimension2
```

Model-space reads — `vertex.GetPoint`, `ICurve.CircleParams`, `face.Normal`,
bounding boxes — come back in **global** coordinates. Convert before comparing
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

**Mid-plane:** `T1=6` (`swEndCondMidPlane`), with `D1` as the *total* depth:
`FeatureExtrusion3(True, False, False, 6, 0, depth, depth, ...)` with the rest as
above. A second boss built this way reported `GetTypeName2 == "ICE"`, not
`"Extrusion"`.

**`CreateCenterRectangle` returns 6 sketch segments, not 4**: the 4 profile
edges plus 2 construction diagonals used for the symmetric constraint. Return
order does not map to bottom/left/top/right. Filter on `.ConstructionGeometry`
first to drop the diagonals, then classify horizontal vs. vertical with
`.GetStartPoint2` / `.GetEndPoint2` (zero-arg, no parens).

## Offset reference planes

```python
doc.ClearSelection2(True)
select_by_id2(ext, "Top Plane", "PLANE")
upper = fm.InsertRefPlane(8, mm(7), 0, 0, 0, 0)          # swRefPlaneReferenceConstraint_Distance
add_equation(eq, '"D1@%s" = "web_half"' % upper.Name)

doc.ClearSelection2(True)
select_by_id2(ext, "Top Plane", "PLANE")
lower = fm.InsertRefPlane(8 | 256, mm(7), 0, 0, 0, 0)    # | _OptionFlip: the other side
```

The offset is `D1@PlaneN`, and an equation can drive it. Sketches on these planes
worked with `sketch_frame` and `select_origin` even though the planes miss the
origin. A cut sketched on the flipped plane needed the opposite `Dir`; see
[Cuts](#cuts). Wrap planes into `1-Ref` with `types={name: "PLANE"}`.

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
    select_origin(ext, append=True)              # ("", "ORIGIN") worked here, but has failed elsewhere
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

## Closed profiles with explicit relations

For a profile of lines and arcs, turn inference off, create the entities, then
add every relation yourself. With `AddToDB = True` nothing is inferred, not even
the endpoint merge between two chained `CreateLine` calls, so the DOF count is
exact. Every sketch built this way reached `GetConstrainedStatus == 3` on the
first full run once its relation set was right.

```python
frame = sketch_frame(doc)
p, q, r, c = (model_to_sketch(frame, pt) for pt in (p_model, q_model, r_model, c_model))

sm.AddToDB = True
line = sm.CreateLine(p[0], p[1], 0, q[0], q[1], 0)
arc = sm.CreateArc(c[0], c[1], 0, q[0], q[1], 0, r[0], r[1], 0, direction)   # centre, start, end
sm.AddToDB = False

def relate(entities, code):                   # sketch entities/points, or "ORIGIN"
    doc.ClearSelection2(True)
    for i, e in enumerate(entities):
        if isinstance(e, str):
            select_origin(ext, append=i > 0)
        else:
            e.Select4(i > 0, none_dispatch())
    doc.SketchAddConstraints(code)
    doc.ClearSelection2(True)

relate([line.GetEndPoint2, arc.GetStartPoint2], "sgMERGEPOINTS")
relate([line, arc], "sgTANGENT")
relate([arc.GetCenterPoint2, "ORIGIN"], "sgHORIZONTALPOINTS2D")   # choose with runs_horizontal
```

Relation codes used successfully: `sgMERGEPOINTS`, `sgCOINCIDENT` (point–curve,
and point–origin), `sgTANGENT` (line–arc, line–circle), `sgVERTICAL2D`,
`sgHORIZONTALPOINTS2D`, `sgVERTICALPOINTS2D`.

- **`CreateArc` plus `sgMERGEPOINTS`** with the neighbouring line ends produced
  closed profiles that cut and extruded correctly. It is an alternative to
  `CreateTangentArc` when tangency is added as a relation. `direction` is `+1`
  for counter-clockwise from start to end, `-1` for clockwise, worked out in
  *sketch* coordinates after mapping.
- **Pick horizontal/vertical relations from the transform**, as in
  [where sketch coordinates land](#where-sketch-coordinates-land).
- **Not tested:** whether an exactly horizontal `CreateLine`
  ([returns `None`](troubleshooting.md#createline-returns-none-for-a-horizontal-line))
  behaves differently with `AddToDB = True`. Every line in these sketches was
  laid out off-horizontal.

### Angle dimensions

`AddDimension2` on two selected lines creates an angle dimension. The placement
point decides whether it measures the angle or its supplement: the session got
105° where 75° was equally valid. Read the value back and write the equation
that matches:

```python
doc.ClearSelection2(True)
line_a.Select4(False, none_dispatch())
line_b.Select4(True, none_dispatch())
with no_input_dim_dialog(sw):
    dd = doc.AddDimension2(*sketch_to_model(frame, u, v))   # placement, in model coordinates
dim = dd.GetDimension2(0)
expr = '90 + "arm_half_angle"' if dim.SystemValue > math.pi / 2 else '90 - "arm_half_angle"'
# after InsertSketch(True): add_equation(eq, '"%s@%s" = %s' % (dim.Name, sketch.Name, expr))
```

## Cuts

`FeatureCut4` takes 27 positional args, in this order in the SW2026 typelib:

```python
select_by_id2(ext, sketch_name, "SKETCH")
cut = fm.FeatureCut4(
    True, False, True,            # Sd, Flip, Dir - Dir is a direction, see below
    1, 0,                         # T1, T2: swEndCondThroughAll = 1, Blind = 0
    mm(100), 0,                   # D1, D2 (metres)
    False, False,                 # Dchk1, Dchk2: draft on, per direction
    False, False,                 # Ddir1, Ddir2: draft outward
    0, 0,                         # Dang1, Dang2 (radians)
    False, False, False, False,   # OffsetReverse1/2, TranslateSurface1/2
    False,                        # NormalCut
    True, True,                   # UseFeatScope, UseAutoSelect
    False, False, False,          # AssemblyFeatureScope, AutoSelectComponents, PropagateFeatureToParts
    0, 0, False,                  # T0, StartOffset, FlipStartOffset
    False,                        # OptimizeGeometry
)
```

Compared with `FeatureExtrusion3`'s 23, `Merge` becomes `NormalCut`, the three
assembly-scope flags follow `UseAutoSelect`, and `OptimizeGeometry` comes last.

**`Dir` picks which side of the sketch plane the cut goes to.** Try `True`
first. In the first session a plain cut returned `None` for every `Dir=False`
combination in a full parameter sweep, and succeeded for every `Dir=True`
combination with `UseFeatScope=True, UseAutoSelect=True`. That reflected one
sketch's position relative to its material, not a rule. A later pocket on a
*flipped* offset plane needed `Dir=False`. With `True` it produced a valid
feature that cut toward the web. Don't hard-code `Dir`: create the cut, check it,
and undo and retry if it's wrong.

```python
cut = None
for dir_flag in (True, False):
    verify_tag(doc, known, TAG, VALUE)
    doc.ClearSelection2(True)
    select_by_id2(ext, sketch_name, "SKETCH")
    cut = fm.FeatureCut4(True, False, dir_flag, ...)
    if cut is None:
        continue
    if cut_is_right():                  # the checks below
        break
    verify_tag(doc, known, TAG, VALUE)
    doc.EditUndo2(1)                    # removes the failed feature cleanly
    cut = None
```

Undo left nothing behind that a later tree and equation audit could find, but
feature numbering still advances (`Cut-Extrude4` was the first surviving lower
cut). Find features by Description, not default name. Checks that worked, over
`body.GetFaces()` for each body in `doc.GetBodies2(0, True)`:

- **Floor direction:** the planar face at the floor height (centre of
  `face.GetBox`) has a `face.Normal` Z component pointing away from the web.
- **Draft direction:** for 7° walls, faces with `0.08 < |nz| < 0.16` must have
  `nz * z_centre > 0`.
- **Volume:** `volume(doc)` falls inside the expected range.

`face.Normal` is meaningful for planar faces; the checks skipped any face whose
normal did not have unit length.

Cut features created this way report `GetTypeName2 == "ICE"` on this build, not
`"Cut"`. So did a mid-plane boss.

### Draft while cutting

- **Pocket:** `Dchk1=True, Ddir1=True, Dang1=deg(7)` widened the pocket away
  from the sketch plane, which is the correct casting draft.
- **Outside trim (keep inside, cut outside):** `Sd=False, Flip=True, T1=1, T2=1`
  with draft in both directions (`Dchk1=Dchk2=True`, `Dang1=Dang2`).
  `Ddir1=Ddir2=True` gave side faces that narrow away from the sketch plane,
  which was right here; `False` was wrong. Only the normal and volume checks
  showed which, so run them every time.
- **Driving the draft angle:** it appears as a feature display dimension
  (`D3@...`). Find it by value and drive it from the variable:

```python
dd = cut.GetFirstDisplayDimension                 # bare
while dd:
    dim = dd.GetDimension2(0)
    if abs(dim.SystemValue - deg(7)) < 1e-6:
        add_equation(eq, '"%s@%s" = "draft_ang"' % (dim.Name, cut.Name))
    dd = cut.GetNextDisplayDimension(dd)          # takes an argument, so parens
```

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
`centre_xyz` from the parameter set in model coordinates, not from the raw
sketch arguments. Confirm with `selected_count(doc)` before calling
`FeatureFillet3`.

### Edges between two features

Choosing edges by the features that created their two faces fits the selection
policy's "created-entity list" rule well, and survives dimension changes.
Fillets on 10 and 16 edges picked this way built in one `FeatureFillet3` call
each.

```python
by_desc = {f.Name: (f.Description or "") for f in iter_features(doc)}
trim = {n for n, d in by_desc.items() if d.startswith("Plan trim")}
bore = {n for n, d in by_desc.items() if d.startswith("Main bore")}

def rule(a, b):                                   # "trim face next to anything but the bore"
    return any(x in trim and y not in trim and y not in bore for x, y in ((a, b), (b, a)))

picked = []
for body in doc.GetBodies2(0, True) or ():        # 0 = swSolidBody
    for face in body.GetFaces() or ():            # parens
        for edge in face.GetEdges or ():          # bare - unlike body.GetEdges()
            pair = edge.GetTwoAdjacentFaces2 or ()
            if len(pair) != 2 or pair[0] is None or pair[1] is None:
                continue
            if not rule(pair[0].GetFeature.Name, pair[1].GetFeature.Name):
                continue
            if not any(sw.IsSame(edge, p) == 1 for p in picked):   # reached from both faces
                picked.append(edge)
```

Identify features by Description rather than default name. Undo-and-retry
shifts the numbering (see [Cuts](#cuts)).

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

A later session closed extrude and cut profiles with `CreateArc` plus explicit
`sgMERGEPOINTS` relations under `AddToDB = True`
([closed profiles](#closed-profiles-with-explicit-relations)). That may be
another way to connect a sweep path; as a sweep path it is untested.

## Saving, opening and screenshots

```python
import os
doc.Save()                                              # already-named document, in place
title = save_as(doc, path, known)                       # new name: SaveAs3(abspath, 0, 1)
doc = sw.OpenDoc(os.path.abspath(path), doc_type)       # swDocumentTypes_e; returns doc or None
doc.SaveBMP(os.path.abspath("view.bmp"), 1280, 720)     # relative path -> False, silently
```

- **`SaveAs3(Name, Version, Options)` for a new name.** Options `1` is Silent.
  `2` is Copy, which opens a modal Save As dialog and blocks the script after
  the file is already written. With `1` it returned `0` and renamed the document
  in place (`GetTitle` became the new file name), and `doc.Save()` worked
  afterwards. `save_as` also refuses an existing target and runs the identity
  check.
- **`OpenDoc`, not `OpenDoc6`.** `OpenDoc6`'s `ByRef Long` `Errors`/`Warnings`
  raise `Type mismatch` with `None` or `[0]`, and so does the 6-arg `SaveAs`. See
  [troubleshooting.md](troubleshooting.md#saveas-or-opendoc6-raises-type-mismatch).
- **`SaveBMP` needs an absolute path.** A relative one returns `False` with no
  file written.
- A document opened this way is subject to the same identity rules as any
  other. If the file is already open in the user's session, assume `OpenDoc`
  hands back *their* document; this is unconfirmed. Run `assert_scratch_doc`
  before mutating it.

### Standard-view screenshots

```python
VIEWS = {"front": 1, "back": 2, "left": 3, "right": 4, "top": 5, "bottom": 6}   # swStandardViews_e
for name, view_id in VIEWS.items():
    doc.ShowNamedView2("", view_id)
    doc.SaveBMP(os.path.abspath("stage3_%s.bmp" % name), 1280, 720)
```

Saving the six standard views after each major stage was cheap. It caught a
[design defect every other check missed](troubleshooting.md#a-clean-fully-defined-rms-passing-model-still-has-a-design-defect).
The files are BMP; convert them to PNG to look at them (for example with
Pillow). View ids are from the SW2026 constants typelib.

## Global variables and driven dimensions

```python
eq = call0(doc, "GetEquationMgr")
add_equation(eq, '"plate_width" = 120')              # global variable
add_equation(eq, '"D1@Sketch1" = "plate_width"')     # driven dimension
add_equation(eq, '"D2@Sketch5" = "inset"', index=7)  # insert at a position
for i, text, value, status in equations(eq):         # status -1 = broken
    ...
```

`add_equation` uses `Add2(index, text, use_automatic_solve_order)`: index `-1`
appends and any other value inserts there. It then verifies `GetCount` actually
incremented, because `Add3` fails *silently* on this build. See
[troubleshooting.md](troubleshooting.md#add3-adds-no-equation-and-returns--1).

- **Trig in equation text uses degrees**, not radians: `tan(45) + atn(1)`
  evaluates to `46`. `sin`, `cos`, `tan`, `atn`, `arcsin` and `sqr` (square root)
  all worked, matching Python to 4 decimal places.
- **`Status` has no index.** It reports on the equation most recently
  evaluated, so `equations()` reads `Value(i)` first.
- **Don't delete a global that dimensions reference just to change it.** The
  dimension equations go stale and block re-adding it; the recovery order is in
  [troubleshooting.md](troubleshooting.md#add2-returns--1-when-re-adding-a-global-that-dimensions-reference).
- **Construction-only values**, such as sketch overruns that exist only to reach
  past the part, are globals too. Label them as non-design values in the header.

A dimension driven by an equation reports as driven; one set through
`model.Parameter("D1@Sketch1").SystemValue = 0.12` does not. `rms_check.py` uses
that distinction to measure how much of the model is genuinely parameterized.

## Measuring the result

A clean rebuild says nothing about shape. Checks that have caught real errors:

```python
v = volume(doc)                                   # m³: doc.Extension.CreateMassProperty.Volume
for body in doc.GetBodies2(0, True) or ():        # 0 = swSolidBody
    lo, hi = body_extents(body)                   # exact-ish extents, from tessellation
    faces = body.GetFaces() or ()
    for face in faces:
        n, box = face.Normal, face.GetBox         # (nx, ny, nz); (x0, y0, z0, x1, y1, z1)
```

- **Volume** after every feature, against a hand estimate. It is cheap, and it
  catches a cut that went the wrong way.
- **Extents:** `body_extents` takes min/max over `face.GetTessTriangles(True)`
  vertices. `doc.GetPartBox(True)` is padded (X max 42.196 against a true
  42.000), and `body.GetBodyBox()` returned `None` on a cut body. Tessellated
  extents can fall slightly short where the extreme lies on a curved face.
- **Face normals and boxes** on planar and drafted faces, as in the direction
  checks under [Cuts](#cuts).
- **Face counts:** `len(faces)` against the expected count.
- **Look at it:** take [standard-view screenshots](#standard-view-screenshots)
  after each stage.

## Wrapping a group's features in a folder

```python
wrap_in_folder(doc, ["Boss-Extrude1", "Fillet1"], "3-Core")
wrap_in_folder(doc, ["Sketch1"], "1-Ref", types={"Sketch1": "SKETCH"})
wrap_in_folder(doc, ["Plane1", "Plane2"], "1-Ref", types={"Plane1": "PLANE", "Plane2": "PLANE"})
```

`InsertFeatureTreeFolder2(2)` — type `2` is "Containing", which wraps the
*current selection* rather than creating an empty folder. Select the group's
features first (`"BODYFEATURE"` for features, `"SKETCH"` for sketches), then
wrap and rename.

Create each folder **as its group completes**. SolidWorks folders must hold
contiguous features and will not reorder past a dependency, so there is no
sort-the-tree-afterwards path on a part of any realistic complexity.

Flat traversal surfaces a synthetic `"...___EndTag___"` marker feature
immediately after a folder's contents. It keeps the folder's *default* name:
after renaming to `3-Core` it is still `Folder2___EndTag___`, so match on the
`___EndTag___` suffix. It is itself `FtrFolder`-typed, so type-based filters
skip it naturally — do not special-case it away.

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

These vary by version and by how the feature was created. On SW2026, `ICE`
covers `FeatureCut4` cuts *and* some `FeatureExtrusion3` bosses: of two mid-plane
bosses in one part, the first reported `Extrusion` and the second `ICE`. Run
`python "<skill dir>/rms_check.py" --dump-types` against known-good parts and adjust the
classification sets at the top of the checker rather than trusting this list.
