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
- [A SaveAs call opened a Save As dialog and the script hangs](#a-saveas-call-opened-a-save-as-dialog-and-the-script-hangs)
- [Dimensioning a rectangle edge hangs or crashes SolidWorks](#dimensioning-a-rectangle-edge-hangs-or-crashes-solidworks)
- [A long script crashes SolidWorks partway through](#a-long-script-crashes-solidworks-partway-through)
- [The script deadlocks after adding a settle/wait helper](#the-script-deadlocks-after-adding-a-settlewait-helper)

**Calls that return None or silently do nothing**
- [Add3 adds no equation and returns -1](#add3-adds-no-equation-and-returns--1)
- [FeatureCut / FeatureCut4 returns None](#featurecut--featurecut4-returns-none)
- [FeatureFillet3 returns None](#featurefillet3-returns-none)
- [MoveToFolder does nothing, or raises a type mismatch](#movetofolder-does-nothing-or-raises-a-type-mismatch)
- [AddHorizontalDimension2 returns None for every entity in one sketch](#addhorizontaldimension2-returns-none-for-every-entity-in-one-sketch)
- [CreateLine returns None for a horizontal line](#createline-returns-none-for-a-horizontal-line)
- [InsertProtrusionSwept4 returns None for a line-plus-arc path](#insertprotrusionswept4-returns-none-for-a-line-plus-arc-path)
- [SaveBMP returns False and writes nothing](#savebmp-returns-false-and-writes-nothing)
- [SelectByID2 fails for the origin inside a sketch](#selectbyid2-fails-for-the-origin-inside-a-sketch)
- [A dimension call returns None right after a horizontal or vertical relation](#a-dimension-call-returns-none-right-after-a-horizontal-or-vertical-relation)
- [Add2 returns -1 when re-adding a global that dimensions reference](#add2-returns--1-when-re-adding-a-global-that-dimensions-reference)
- [Setting an equation's text in place does nothing](#setting-an-equations-text-in-place-does-nothing)
- [GetBodyBox returns None](#getbodybox-returns-none)

**COM type and member errors**
- ['X' object is not callable, or 'Member not found' (-2147352573)](#x-object-is-not-callable-or-member-not-found--2147352573)
- [Bare getters return bound methods, or 'method' object is not iterable](#bare-getters-return-bound-methods-or-method-object-is-not-iterable)
- [DISP_E_TYPEMISMATCH / 'Type mismatch' on a call with an optional Object](#disp_e_typemismatch--type-mismatch-on-a-call-with-an-optional-object)
- ['Element not found' from EnsureDispatch or CastTo](#element-not-found-from-ensuredispatch-or-castto)
- ['Unable to read write-only property' from GetCoords()](#unable-to-read-write-only-property-from-getcoords)
- [SelectByID2 failed for a folder](#selectbyid2-failed-for-a-folder)
- [AttributeError from GetCircleParams()](#attributeerror-from-getcircleparams)
- [MathUtility.CreatePoint raises 'Member not found'](#mathutilitycreatepoint-raises-member-not-found)
- [SaveAs or OpenDoc6 raises 'Type mismatch'](#saveas-or-opendoc6-raises-type-mismatch)

**Session and connection**
- [GetActiveObject raises 'Operation unavailable' but SLDWORKS.exe is running](#getactiveobject-raises-operation-unavailable-but-sldworksexe-is-running)
- [A script created a sketch in the user's real document](#a-script-created-a-sketch-in-the-users-real-document)
- [GetDocuments raises instead of returning an empty list](#getdocuments-raises-instead-of-returning-an-empty-list)
- [ActivateDoc2 raises 'Type mismatch'](#activatedoc2-raises-type-mismatch)
- [A failed build stage left a sketch open](#a-failed-build-stage-left-a-sketch-open)

**Wrong geometry, no error**
- [A hole is in the wrong place but the sketch is fully defined and rebuilds clean](#a-hole-is-in-the-wrong-place-but-the-sketch-is-fully-defined-and-rebuilds-clean)
- [Sketch geometry lands on the wrong global axis, or mirrored](#sketch-geometry-lands-on-the-wrong-global-axis-or-mirrored)
- [AddDimension2 gave one diagonal dimension instead of H and V](#adddimension2-gave-one-diagonal-dimension-instead-of-h-and-v)
- [A cut succeeds but removes material on the wrong side](#a-cut-succeeds-but-removes-material-on-the-wrong-side)
- [ORIGIN selection picked up a dimension instead of the origin](#origin-selection-picked-up-a-dimension-instead-of-the-origin)
- [GetPartBox reports extents larger than the part](#getpartbox-reports-extents-larger-than-the-part)
- [A clean, fully defined, RMS-passing model still has a design defect](#a-clean-fully-defined-rms-passing-model-still-has-a-design-defect)
- [An enum constant has an unexpected value](#an-enum-constant-has-an-unexpected-value)

**rms_check.py**
- [rms_check.py reports SKIP or WARN for detail.holes_last](#rms_checkpy-reports-skip-or-warn-for-detailholes_last)
- [grouping.all_features_in_a_group fails on Comments, Selection Sets, Markups](#groupingall_features_in_a_group-fails-on-comments-selection-sets-markups)
- [detail.no_internal_references flags a cut's own sketch](#detailno_internal_references-flags-a-cuts-own-sketch)

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
`python "<skill dir>/scripts/sw_preflight.py" --fix` checks and clears it before a build.

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

## A SaveAs call opened a Save As dialog and the script hangs

**Symptom.** `doc.SaveAs3(path, 0, 2)` never returns, and SolidWorks shows a
Save As dialog. The target file has already been written.

**Cause.** In `swSaveAsOptions_e`, `2` is **Copy**, not Silent. **Silent is
`1`.** Both values are confirmed against the SW2026 constants typelib.

**Fix.** `doc.SaveAs3(os.path.abspath(path), 0, 1)` returns `0` and renames the
document in place. `sw_helpers.save_as` wraps it with the identity check. Check
that the target does not exist before the call: a dialog-raising attempt may
already have written it.

If the dialog is already open, cancel it in SolidWorks. In the one recorded
case the Python client had already been interrupted, the user cancelled the
dialog by hand, and later `SaveAs3` and `Save()` calls did not crash, with no
restart. That contrasts with the `Save()` crash above, but it is one observation
with a different dialog. Don't read it as evidence the crash risk is gone.

*Verified on: SW2026 SP1.1 (option values). Single observation (no crash after
the interrupt).*

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

Staging has since held up. Seven short stage scripts sharing one module built a
complete part without crashing SolidWorks. Each stage located the build document
with `find_tagged_doc` rather than a remembered title, and called `verify_tag`
immediately before every mutating call. The user's original open titles came
from a small JSON state file written by stage 1. Recipe in
[api-recipes.md](api-recipes.md#creating-the-build-document).

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
still returns `-1`. `GetCount`, `Equation(i)`, and `Value(i)` followed by the
index-less `Status` property (`sw_helpers.equations`) verify what actually
landed regardless of which `Add*` you used.

*Verified on: SW2026 SP1.1 (rev 34.1.1.11). Untested on other versions — treat
"Add3 is broken on 2026" as confirmed and "Add3 is broken in general" as
unverified.*

## FeatureCut / FeatureCut4 returns None

**Symptom.** A plain sketch-based cut returns `None` unconditionally, across
every combination of end condition (`Blind` with oversized depth, `ThroughAll`,
`ThroughAllBoth`), `Flip`, and `NormalCut`. The same cut works in the GUI.

**Cause.** The `Dir` argument, which picks the side of the sketch plane the
cut goes to. It is easy to assume it defaults fine at `False` by analogy with
`FeatureExtrusion3`, where `Dir=False` works. Most likely the chosen side held
no material to cut (inferred).

**Fix.** Try **`Dir=True`** first, then `False`. In the first session's parameter
sweep (`itertools.product` over `Flip`, `Dir`, `NormalCut`, `UseFeatScope`/
`UseAutoSelect`), every `Dir=False` combination returned `None`, and every
`Dir=True` combination with `UseFeatScope=True, UseAutoSelect=True` succeeded.
That reflected one sketch's position relative to its material, not a rule. A
later cut on a *flipped* offset plane needed `Dir=False` and returned a valid
feature. Check the result either way: the wrong `Dir` can also succeed and
[cut the wrong side](#a-cut-succeeds-but-removes-material-on-the-wrong-side).

*Verified on: SW2026 SP1.1. Partially superseded: this entry originally read
`Dir=True` as mandatory.*

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

## CreateLine returns None for a horizontal line

**Symptom.** `SketchManager.CreateLine(x1, y, z1, x2, y, z2)` returns `None`
when both endpoints have exactly the same Y argument. Vertical and generic
diagonal lines in the same sketch are created normally.

**Cause.** Unknown. Isolated against a fresh document, so it is not the
accumulated per-sketch state of the entry above. `CreateCenterRectangle` is not
affected — its horizontal edges are created fine.

**Fix.** Offset one endpoint's Y by 0.01 mm (`mm(0.01)`); the line is then
created. That bakes a 10 µm slope into the sketch, so where the line is meant to
be horizontal, add a Horizontal relation or a dimension afterwards rather than
trusting the coordinate — that follow-up step has **not** been tested here.

Remember Y here is the sketch argument, not global Y — see
[where sketch coordinates land](api-recipes.md#where-sketch-coordinates-land).

*Verified on: SW2026 SP1.1.*

## InsertProtrusionSwept4 returns None for a line-plus-arc path

**Symptom.** A sweep whose path is a `CreateLine` joined to a `Create3PointArc`
returns `None`, even with the arc endpoint numerically identical to the line
endpoint. Ruled out, each tried separately: arc radius (checked correct),
`KeepTangency` on and off, an explicit profile sketch plus path sketch instead
of `CircularProfile`, and clean round-number coordinates instead of computed
ones. Arc-only, line-only, and line+line (sharp corner) paths all swept fine.

**Cause.** Identical coordinates are not a topological connection.
`Create3PointArc` and a separately created `CreateLine` are two independent
entities whose endpoints merely overlap, so the path is not a connected chain.
Why line+line *does* merge is not explained.

**Fix.** Create each arc in a multi-segment path with **`CreateTangentArc`**,
which starts from the previous entity's actual endpoint. The identical sweep
succeeded immediately. Recipe in
[api-recipes.md](api-recipes.md#sweep-paths-chain-arcs-with-createtangentarc).

The same reasoning should apply to loft paths and guide curves; that is
untested.

*Verified on: SW2026 SP1.1.*

## SaveBMP returns False and writes nothing

**Symptom.** `doc.SaveBMP(path, width, height)` returns `False`, raises
nothing, and no file appears.

**Cause.** `path` was relative.

**Fix.** Pass an absolute path: `doc.SaveBMP(os.path.abspath(path), w, h)`.
Passing absolute paths to every file-taking call (`OpenDoc`, `SaveAs`) is the
safe default; only `SaveBMP` has been confirmed to fail on a relative one.

*Verified on: SW2026 SP1.1.*

## SelectByID2 fails for the origin inside a sketch

**Symptom.** `select_by_id2(ext, "", "ORIGIN")` raises `SelectByID2 failed for
'' as 'ORIGIN'` in a fresh, uncluttered sketch on Right Plane.

**Cause.** Unknown. The same call worked in the center-rectangle recipe on Front
Plane, and in a cluttered sketch it
[picked a dimension instead](#origin-selection-picked-up-a-dimension-instead-of-the-origin).
The empty-name form is not dependable.

**Fix.** Select the origin by name as an external sketch point:
`sw_helpers.select_origin(ext, append)`, which is
`SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, append, 0, none_dispatch(), 0)`.
It selected type `25` (`swSelEXTSKETCHPOINTS`). It worked for coincident,
horizontal-points and vertical-points relations and for horizontal and vertical
dimensions, on default planes and on offset planes that do not pass through the
origin. `("", "EXTSKETCHPOINT")` also selected type 25.

*Verified on: SW2026 SP1.1, in 7 sketches.*

## A dimension call returns None right after a horizontal or vertical relation

**Symptom.** `AddHorizontalDimension2` or `AddVerticalDimension2` returns `None`
straight after `SketchAddConstraints("sgHORIZONTALPOINTS2D")` or
`("sgVERTICALPOINTS2D")`, and the sketch geometry has jumped.

**Cause.** The relation was picked from the *model* axis. "Both points on global
X" is horizontal only if global X runs horizontally in that sketch, and in a Top
Plane sketch on the test template it runs vertically. The wrong relation dragged
the geometry, and the next dimension was refused.

**Fix.** Derive both the relation and the dimension call from the sketch
transform: `runs_horizontal(sketch_frame(doc), (1, 0, 0))`. See
[api-recipes.md](api-recipes.md#where-sketch-coordinates-land). If the sketch
has since absorbed several failed calls,
[rebuild it rather than debugging in place](#addhorizontaldimension2-returns-none-for-every-entity-in-one-sketch).

*Verified on: SW2026 SP1.1.*

## Add2 returns -1 when re-adding a global that dimensions reference

**Symptom.** A global was replaced by `eq.Delete(i)` then `add_equation(...)`.
`Add2` returns `-1` and the count does not move for
`"pocket_inset_rad" = "end_rad" - "pocket_inset_hole"`. After that, a second
global feeding the same sketch is rejected too.

- **Accepted:** a plain constant under the same name (`"pocket_inset_rad" = 4.397`),
  and the identical expression under an unused name
  (`"zz_b" = "end_rad" - "pocket_inset_hole"`).
- **Error state:** `GetWhatsWrongCount` is 1, attributed to the Equations
  feature. The dependent dimension equations (`"D2@Sketch5" = ...` and so on)
  evaluate to `0`, though the sketch dimensions keep their old values and the
  sketches still report fully defined.
- **Ruled out:** list order. A forward reference inserted at index 0 was
  accepted, with automatic solve order on.

**Cause (inferred).** While the global was missing, the dimension equations that
use it went into an error state, and that state doesn't clear when the global
comes back. SolidWorks then rejects any non-constant equation for a global that
feeds an erroring dimension equation.

**Fix.**
1. Delete the dependent dimension equations. The what's-wrong count drops to 0.
2. Delete and re-add the globals with their new expressions. They are accepted now.
3. Re-add the dimension equations at their original positions with
   `add_equation(eq, text, index=original_index)`.
4. Rebuild: 0 errors, and the values matched Python exactly.

Better: don't delete a referenced global in order to change it. No way to edit
an equation in place has worked yet ([next entry](#setting-an-equations-text-in-place-does-nothing)),
so for now use the order above. Check each equation's health with
`sw_helpers.equations(eq)`, which reads `Value(i)` before the index-less `Status`
property; `-1` marks a broken one.

*Verified on: SW2026 SP1.1 (symptom, fix). Inferred (cause).*

## Setting an equation's text in place does nothing

**Symptom.** Writing `Equation(i)` as a property, with
`eq._oleobj_.Invoke(dispid, 0, pythoncom.DISPATCH_PROPERTYPUT, 0, i, text)`,
raises nothing and leaves the text unchanged.

**Fix.** None confirmed. `SetEquationAndConfigurationOption(Index, Equation,
WhichConfigurations, ConfigNames)` exists in the SW2026 typelib and has not been
tried. For now, change an equation by deleting and re-adding it, in the order
[above](#add2-returns--1-when-re-adding-a-global-that-dimensions-reference) when
dimensions reference it.

*Verified on: SW2026 SP1.1 (the silent no-op).*

## GetBodyBox returns None

**Symptom.** `body.GetBodyBox()` returns `None`. It had worked on a plain
extrusion earlier in the same session, and failed on the finished, cut body.

**Cause.** Unknown.

**Fix.** Use `sw_helpers.body_extents(body)`, which takes min/max over
tessellation vertices. `doc.GetPartBox(True)` is not a substitute where tenths
of a millimetre matter:
[it is padded](#getpartbox-reports-extents-larger-than-the-part).

*Single observation: SW2026 SP1.1.*

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
`face.GetEdges` is **not** one of them: `face.GetEdges()` raises `TypeError:
'tuple' object is not callable`. The typelib declares both `GetEdges` the same
way, so it can't settle which is which.

If the opposite happens and a bare getter gives a bound method instead of a
value, you are early-bound; see the next entry.

*Verified on: SW2026 SP1.1.*

## Bare getters return bound methods, or 'method' object is not iterable

**Symptom.** `sw_preflight.py` or your own script prints the version as
`<bound method ISldWorks.RevisionNumber of <win32com.gen_py.SldWorks 2026 Type Library...>>`.
Iterating `sw.GetDocuments` raises `'method' object is not iterable`. `doc.GetTitle`,
`feat.GetTypeName2` and every other bare getter give bound methods instead of
values. `type(sw)` is a `win32com.gen_py...` class.

**Cause.** Early binding you didn't ask for. pywin32's gen_py cache holds a module
for the SldWorks typelib, so `GetActiveObject` and `Dispatch` return the
generated class. `gencache.EnsureModule` on the main typelib, the
signature-reading trick, writes that module.

**Fix.** Don't add parentheses everywhere. Force late binding with
`dynamic.Dispatch(app._oleobj_)`, which is what `sw_helpers.connect()` and
`late_bound()` do. Objects reached from a late-bound parent stay late-bound.
`sw_preflight.py` reports the condition as a `com.late_binding` WARN, and
`rms_check.py` re-wraps on its own. Details in
[dispatch-quirks.md](dispatch-quirks.md#a-gen_py-cache-silently-switches-to-early-binding).

*Verified on: SW2026 SP1.1 (rev 34.1.1).*

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
That module stays cached, and from then on plain `Dispatch` on that machine is
early-bound; see
[bare getters return bound methods](#bare-getters-return-bound-methods-or-method-object-is-not-iterable).
Code that attaches through `connect()` is unaffected.

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

## AttributeError from GetCircleParams()

**Symptom.** `edge.GetCurve.GetCircleParams()` raises `AttributeError`.

**Cause.** No such method exists. Pattern-matching from the many `Get...`
getters gives the wrong name here.

**Fix.** Read the bare property **`CircleParams`** (no `Get`, no parens):

```python
cx, cy, cz, nx, ny, nz, r = edge.GetCurve.CircleParams   # centre, axis, radius; global, metres
```

`sw_helpers.circular_edges` wraps this for picking fillet edges.

*Verified on: SW2026 SP1.1.*

## MathUtility.CreatePoint raises 'Member not found'

**Symptom.** `sw.GetMathUtility.CreatePoint(arr)` raises `-2147352573 Member not
found`, whether `arr` is a list, a tuple or a `VT_ARRAY | VT_R8` VARIANT. That
puts `MultiplyTransform` out of reach for mapping model points into a sketch.

**Fix.** Apply the transform in Python. `sketch_frame(doc)` reads
`ActiveSketch.ModelToSketchTransform.ArrayData`, and `model_to_sketch` /
`sketch_to_model` apply it. See
[api-recipes.md](api-recipes.md#where-sketch-coordinates-land).

*Verified on: SW2026 SP1.1.*

## SaveAs or OpenDoc6 raises 'Type mismatch'

**Symptom.** `SaveAs(...)` raises `Type mismatch` when `ExportData`, `Errors`
and `Warnings` are passed `None`. `sw.OpenDoc6(...)` raises `Type mismatch` with
either `None` or `[0]` for `Errors` / `Warnings`.

**Cause.** `Errors` and `Warnings` are `ByRef Long` out-parameters, and late
binding does not marshal plain Python values into them — the same problem as
[ActivateDoc2](#activatedoc2-raises-type-mismatch). See
[dispatch-quirks.md](dispatch-quirks.md#byref-long-out-parameters).

The session reported this as `IModelDoc2.SaveAs`. In the typelib, though, the
6-arg form is `IModelDocExtension.SaveAs`, and `IModelDoc2.SaveAs` takes only
`NewName`.

**Fix.** Use the calls without out-parameters:

- Save an already-named document in place: **`doc.Save()`**.
- Save to a new name: **`doc.SaveAs3(os.path.abspath(path), 0, 1)`**, or
  `sw_helpers.save_as`. Options `1` is Silent; `2` is Copy and
  [opens a modal dialog](#a-saveas-call-opened-a-save-as-dialog-and-the-script-hangs).
  It returns `0` and renames the document in place, and `Save()` works afterwards.
- Open a document: **`sw.OpenDoc(os.path.abspath(path), doc_type)`**, where
  `doc_type` comes from `swDocumentTypes_e` (resolve with
  `constants_module()`). It returns the document, or `None` with no reason.

`VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)` for `Errors` and `Warnings`
remains untested. With `SaveAs3` working, there is little reason to try it.

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

`python "<skill dir>/scripts/sw_preflight.py"` detects exactly this state by cross-checking
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

## A failed build stage left a sketch open

**Symptom.** A stage raised while building a sketch. The document is still in
sketch edit mode with the partial entities, and the next run creates another
sketch while that one is still open.

**Fix, with no dialog:** exit the sketch, then delete the leftover with
`DeleteSelection2`, which avoids the confirmation dialog `EditDelete` can raise.

```python
sm = doc.SketchManager
if call0(sm, "ActiveSketch") is not None:
    sm.InsertSketch(True)                       # toggles - only call it while a sketch is open
leftover = last_feature(doc)
if call0(leftover, "GetTypeName2") == "ProfileFeature" and not call0(leftover, "GetChildren"):
    verify_tag(doc, known, TAG, VALUE)
    doc.ClearSelection2(True)
    select_by_id2(doc.Extension, leftover.Name, "SKETCH")
    doc.Extension.DeleteSelection2(0)
```

Check that the sketch has no children before deleting it. Put each stage's
sketch building in `try`/`except` so this cleanup runs automatically.

*Verified on: SW2026 SP1.1, used twice (exit, select, `DeleteSelection2(0)`).
The `ActiveSketch` guard is a precaution against the toggle, not a recorded
failure.*

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
is no automated geometry check in this repo yet. The manual checks that have
worked — volume, face normals, tessellated extents, standard-view screenshots —
are in [api-recipes.md](api-recipes.md#measuring-the-result).

*Verified on: SW2026 SP1.1.*

## Sketch geometry lands on the wrong global axis, or mirrored

**Symptom.** A feature has the right shape but sits in the wrong place: offsets
end up along a different global axis than intended, or come out mirrored on one
plane. A fillet edge match or face check computed in global coordinates finds
nothing. The model rebuilds clean.

**Cause.** `SketchManager.Create*(x, y, z, ...)` arguments are sketch
coordinates, not global `(X, Y, Z)`, and the mapping differs per plane. On this
install's template: Front Plane arg1 → global Y, arg2 → global Z. Right Plane
arg1 → global **−X**, arg2 → global Z. Top Plane arg1 → global Y, arg2 → global
**−X**. arg3 is discarded on all three.

**Fix.** Don't depend on the table. Read the open sketch's own transform with
`sketch_frame(doc)`, then map with `model_to_sketch` / `sketch_to_model`. That
also covers offset planes and templates nobody has probed. Recipe in
[api-recipes.md](api-recipes.md#where-sketch-coordinates-land). Values read back
from the model (`GetPoint`, `CircleParams`, `Normal`) are global, so convert
before comparing.

*Verified on: SW2026 SP1.1 (table: Front, Right and Top Plane; transform: Right,
Top and two offset planes).*

## AddDimension2 gave one diagonal dimension instead of H and V

**Symptom.** Selecting two points and calling `AddDimension2(x, y, z)` produces
a single aligned point-to-point distance, not independent horizontal and
vertical dimensions.

**Fix.** Call `AddHorizontalDimension2` and `AddVerticalDimension2` explicitly
(both take `(X, Y, Z)` placement args the same way). **Re-select both entities
before each call** — selection does not persist across them.

*Verified on: SW2026 SP1.1.*

## A cut succeeds but removes material on the wrong side

**Symptom.** `FeatureCut4` returns a valid feature and the rebuild is clean, but
the cut went the wrong way. In the recorded case, an upper pocket sketched on an
offset plane at Z = +7 cut correctly with `Dir=True`. The identical lower pocket,
on a *flipped* offset plane at Z = −7, cut toward the web instead of away from it.

**Cause.** `Dir` picks the side of the sketch plane, and flipping the plane
reverses which side `True` means.

**Fix.** Don't hard-code `Dir`. Create the cut, check it, and if it's wrong, call
`doc.EditUndo2(1)` and try the other value. The lower pocket needed
`Dir=False`. Undo removed the failed feature cleanly, though feature numbering
still advances, so find features by Description rather than default name. The
checks that caught it are in [api-recipes.md](api-recipes.md#cuts).

*Verified on: SW2026 SP1.1.*

## ORIGIN selection picked up a dimension instead of the origin

**Symptom.** `ext.SelectByID2("", "ORIGIN", 0, 0, 0, ...)` selects a *dimension*
(`swSelDIMENSIONS`, type `14`) rather than the origin point
(`swSelSKETCHPOINTS`, type `11`), and a selection-dependent call then returns
`None`.

**Cause.** `"ORIGIN"` selection is type/coordinate based, not name based. In a
sketch cluttered with overlapping prior test dimensions near the origin, it
picks the wrong thing.

**Fix.** Select the origin by name instead, with `sw_helpers.select_origin(ext, append)`,
i.e. `("Point1@Origin", "EXTSKETCHPOINT")`; see
[SelectByID2 fails for the origin inside a sketch](#selectbyid2-fails-for-the-origin-inside-a-sketch).
Check `doc.SelectionManager.GetSelectedObjectType3(index, -1)` to confirm what
you got (`sw_helpers.selected_type`); `select_origin` gives type `25`. Before
that form was found, the workaround was a fresh, uncluttered sketch.

*Verified on: SW2026 SP1.1.*

## GetPartBox reports extents larger than the part

**Symptom.** After filleting, `doc.GetPartBox(True)` gave X max **42.196** where
the true extreme is **42.000**, and Y ±34.705 against a true ±34.53.

**Cause.** The box is padded, not tight.

**Fix.** Where tenths of a millimetre matter, use `sw_helpers.body_extents(body)`,
which takes min/max over `face.GetTessTriangles(True)` vertices, or check planar
faces analytically with `face.Normal` and `face.GetBox`.

*Verified on: SW2026 SP1.1, compared against tessellation.*

## A clean, fully defined, RMS-passing model still has a design defect

**Symptom.** 0 rebuild errors, every sketch fully defined, bounding boxes exact,
and the part is still wrong. In the recorded case, a drafted pocket left a
knife-edge rim near the ring. It was 7 mm wide at the pocket floor, but draft on
both sides thinned it to nothing at Z ≈ 32, removing part of the top face.

**Cause.** A design-intent error, not an API one. The drawing's 7 mm rim is
measured at the rim top; the model applied it at the floor. With draft on both
the outer faces and the pocket walls, the two readings give very different parts.

**Fix.** Look at the part. Save
[standard-view screenshots](api-recipes.md#standard-view-screenshots) after each
major stage. This defect showed in an orthographic side view as a flat step where
the drawing has a straight tangent face. When a dimension falls on a drafted
wall, decide which height it applies at before building, and compare the widths
it implies against the drawing. Here the floor width near the ring came out
≈19 mm under the top reading, ≈33 mm under the floor reading, and ≈17 mm scaled
from the drawing.

*Verified on: SW2026 SP1.1. A second, different example of the
[wrong-point lesson](#a-hole-is-in-the-wrong-place-but-the-sketch-is-fully-defined-and-rebuilds-clean).*

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

---

## rms_check.py reports SKIP or WARN for detail.holes_last

**Symptom.** An all-cut `4-Detail` folder produces `SKIP  detail.holes_last`, or
`WARN  detail.holes_last  holes interleaved with other detail features`.

**Cause.** Two calibration gaps. The SKIP: cuts created via `FeatureCut4` report
`GetTypeName2 == "ICE"` on this build, not `"Cut"`. The WARN: the sketch in
front of each cut counted as a non-hole feature between the holes.

**Fix.** Both are fixed in the checker. `"ICE"` is in `CUT_TYPES` and
`HOLE_TYPES`, and the rule skips sketches. `"ICE"` also covers some bosses, so a
boss in `4-Detail` still counts as a hole here. If a part trips this, run
`python "<skill dir>/rms_check.py" --dump-types` and adjust the sets at the top
of the checker. Cosmetic only — it does not cause a false FAIL.

*Verified on: SW2026 SP1.1 (both symptoms). The sketch-skipping fix has not yet
been re-run against a live part.*

## grouping.all_features_in_a_group fails on Comments, Selection Sets, Markups

**Symptom.** Every SW2026 part fails `grouping.all_features_in_a_group`, listing
Comments, Selection Sets and Markups as ungrouped.

**Cause.** SW2026 adds system folders of type `CommentsFolder`,
`SelectionSetFolder` and `InkMarkupFolder`, and `TOLERATED_LOOSE` did not list
them.

**Fix.** Fixed in the checker. On another version, run `--dump-types` and add any
new system types to `TOLERATED_LOOSE`.

*Verified on: SW2026 SP1.1 (rev 34.1.1).*

## detail.no_internal_references flags a cut's own sketch

**Symptom.** `FAIL  detail.no_internal_references  Cut-Extrude7 -> Sketch9` on a
part where each detail feature has its own sketch.

**Cause.** The rule read the sketch-to-feature parent link as a coupling
between two detail features. Under the one-sketch-per-feature rule, every
sketch-based detail feature has that link, so the rule could never pass.

**Fix.** Fixed in the checker: a sketch whose only child is the feature it
defines is skipped. A sketch with several consumers is still reported, here and
by `sketches.one_sketch_per_feature`. With an older copy of the checker, waive
the rule in `rms_exceptions.json` and give this as the reason.

*Verified on: SW2026 SP1.1 (the false positive). The fix has not yet been re-run
against a live part.*
