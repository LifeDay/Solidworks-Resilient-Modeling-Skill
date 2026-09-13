# pywin32 dynamic dispatch against SolidWorks

How late binding behaves against this API, and why the obvious workarounds are
wrong. Read this before writing any COM call; most of the entries in
[troubleshooting.md](troubleshooting.md) are downstream of something here.

Verified against **SW2026 SP1.1 (rev 34.1.1)** unless noted. Signatures and enum
values shift between releases — see [capabilities.yaml](../capabilities.yaml) for
what has actually been confirmed on which version.

## Connecting

```python
from sw_helpers import connect
sw = connect()          # attach-or-fail, with late binding forced
model = sw.ActiveDoc
```

which is:

```python
import win32com.client
from win32com.client import dynamic

app = win32com.client.GetActiveObject("SldWorks.Application")
sw = dynamic.Dispatch(app._oleobj_)
```

Everything on this page assumes **late binding**: every call is resolved by
name at call time rather than against a known signature. That produces real
behavioural quirks, not merely a performance difference.

Plain `win32com.client.Dispatch` and `GetActiveObject` give late binding only
while no generated module for the SldWorks typelib is cached; see
[a gen_py cache silently switches to early binding](#a-gen_py-cache-silently-switches-to-early-binding).
`dynamic.Dispatch` gives it unconditionally. Objects returned by calls on a
late-bound parent are late-bound too, because pywin32's `dynamic.CDispatch`
wraps return values with `dynamic.Dispatch`. Wrapping the Application object is
enough (`sw_helpers.late_bound` does it, and is harmless to repeat).

**`Dispatch` starts SolidWorks if it is not already running** — invisibly
(`Visible` is `False`), holding a license seat, with no window to find. Worse,
such an instance never registers in the COM running object table, so a later
`GetActiveObject("SldWorks.Application")` raises `-2147221021 'Operation
unavailable'` while the process sits there `Responding`. `Dispatch` *will*
attach to it, so the next script silently does its work in an invisible
session. To attach-or-fail honestly, use `GetActiveObject`; to clean up an
orphan, attach with `Dispatch`, confirm `GetDocuments` is empty and `Visible`
is `False`, then call `ExitApp()`. `scripts/sw_preflight.py` detects this state
by cross-checking the process list against COM.

## Zero-argument getters auto-invoke — call them without parentheses

**This is the default, not the exception.** Every zero-arg `Get...`-style member
tried on this build resolves on bare attribute access, even though API Help
documents them as methods:

`FirstFeature`, `GetNextFeature`, `GetEquationMgr`, `GetTitle`, `GetTypeName2`,
`GetCount`, `GetConstrainedStatus`, `GetWhatsWrongCount`, `GetStartPoint2`,
`GetEndPoint2`, `GetCenterPoint2`, `GetName`, `GetNameForSelection`,
`GetDimension`, `ActiveSketch`, `GetSketchSegments`, `GetLength`,
`EditRebuild3`, `GetFirstSubFeature`, `GetNextSubFeature`, `GetChildren`,
`GetDefinition`, `GetSpecificFeature2`, `RevisionNumber`, `GetDocuments`,
`GetFirstDisplayDimension`, `CreateMassProperty`, and on topology
`face.GetEdges`, `face.GetFeature`, `face.GetBox`, `face.GetArea`,
`edge.GetTwoAdjacentFaces2`, `edge.GetCurve`.

Adding `()` raises `TypeError: 'X' object is not callable`, because the value is
already resolved by the time you try to call it.

**Try bare attribute access first for anything zero-arg; add `()` only if that
raises `AttributeError`.** Known exceptions: `body.GetFaces()` and
`body.GetEdges()` do *not* auto-invoke and need the parentheses. `face.GetEdges`
does auto-invoke; `face.GetEdges()` raises `'tuple' object is not callable`.
The generated SW2026 typelib declares `IBody2.GetEdges` and `IFace2.GetEdges`
identically, as zero-argument methods, so the typelib cannot tell you which
form works. Only a live call can.

Members that take arguments always need parentheses: `CustomPropertyManager("")`,
`GetDimension2(0)`, `GetNextDisplayDimension(dd)`, `GetTessTriangles(True)`,
`Equation(i)`, `Value(i)`, `IsSame(a, b)`. True properties never take them:
`face.Normal`, `ModelToSketchTransform`, `ArrayData`, `IMassProperty.Volume`,
and `EquationMgr.Status`, which has **no** index argument.

If a bare getter hands you `<bound method ...>` instead of a value, the object
is early-bound, not late-bound. Adding parentheses is the wrong fix; see
[below](#a-gen_py-cache-silently-switches-to-early-binding).

### Why a `callable()` check cannot fix this

The tempting helper — get the attribute, and call it if it's callable — is
wrong, and it is what other SolidWorks skills ship.

A *resolved* `CDispatch` object is itself callable: it implements `__call__` for
the "invoke this COM method" path. So `callable()` cannot distinguish "still
needs calling" from "already resolved, calling it again is wrong". Calling an
already-resolved value raises either `TypeError: 'X' object is not callable`
(plain Python type) or COM `-2147352573 'Member not found'` (resolved
`CDispatch`) — the same underlying mistake wearing two different errors.

The correct zero-arg helper has no fallback call at all, and is shipped as
`call0` in [`scripts/sw_helpers.py`](../scripts/sw_helpers.py):

```python
def call0(obj, name):
    return getattr(obj, name)
```

## Typed nulls for optional Object parameters

Multi-argument methods with an `Object`/`IDispatch` parameter reject a bare
Python `None` with `DISP_E_TYPEMISMATCH` ("Type mismatch"). Wrap it:

```python
win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)   # sw_helpers.none_dispatch()
```

Confirmed fix for `SelectByID2`'s `Callout` parameter; no other argument there
needed explicit VARIANT typing.

Some parameters are pickier still. `EquationMgr.Add3`'s trailing
`ConfigurationName` rejects `None`, `""` *and* `[]`; it needs

```python
win32com.client.VARIANT(pythoncom.VT_EMPTY, None)      # sw_helpers.none_empty()
```

— at which point, on this build, `Add3` still just returns `-1`. See
[troubleshooting.md](troubleshooting.md#add3-adds-no-equation-and-returns--1).

### ByRef Long out-parameters

`Errors` / `Warnings` parameters declared `ByRef Long` (type `16387`,
`VT_BYREF | VT_I4`, in the generated typelib) reject every plain Python value
tried: `None`, `0` and `[0]` all raise `Type mismatch`. Confirmed on
`sw.ActivateDoc2`, `sw.OpenDoc6`, and the 6-arg `SaveAs(Name, Version, Options,
ExportData, Errors, Warnings)`.

**Use the overload without out-parameters** rather than fight the marshaling:
`sw.OpenDoc(Name, Type)` instead of `OpenDoc6`, `doc.Save()` for an
already-named document. Passing `VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)`
is the obvious next thing to try and has **not** been tested here. See
[troubleshooting.md](troubleshooting.md#saveas-or-opendoc6-raises-type-mismatch).

## Early binding is a known-bad path

`win32com.client.gencache.EnsureDispatch("SldWorks.Application")` — and
`CastTo`, which calls it internally — fails with `Element not found` from
`GetTypeInfo()` on essentially every live object tried on SW2026. **Stick with
late binding for actual calls.** If a call returns `None` unexpectedly,
suspect a wrong argument count/type or a misplaced method before reaching for
early binding.

### A gen_py cache silently switches to early binding

Asking for early binding fails loudly. Getting it by accident does not.

**Symptom.** `sw.RevisionNumber` prints as
`<bound method ISldWorks.RevisionNumber of <win32com.gen_py.SldWorks 2026 Type Library...>>`;
iterating `sw.GetDocuments` raises `'method' object is not iterable`;
`type(sw)` is `win32com.gen_py.83A33D31-...x0x34x0.SldWorks`. From then on
every bare getter in this skill returns a bound method.

**Cause.** A generated module for the SldWorks typelib exists in pywin32's
gen_py cache (`win32com.__gen_path__`). `GetActiveObject` passes the ProgID's
CLSID to `gencache.GetClassForCLSID`, and `Dispatch` looks the object's type up
the same way. When a generated class exists, you get it. Running
`gencache.EnsureModule` on the main SldWorks typelib, the signature-reading
trick below, writes exactly that module. On the machine where this was found,
the cache predated the session, and what created it is unknown.

**Fix.** Force late binding on whatever you attach to:
`dynamic.Dispatch(app._oleobj_)`, i.e. `sw_helpers.connect()` or
`late_bound(obj)`. `sw_preflight.py` reports the cached state as a
`com.late_binding` WARN, and `rms_check.py` re-wraps on its own.

*Verified on: SW2026 SP1.1 (rev 34.1.1). Mechanism read from pywin32's
`win32com/client/__init__.py`.*

## Reading a real signature without API Help

Generate the typelib module directly by CLSID rather than through
`EnsureDispatch`. This succeeds even where `EnsureDispatch` fails, and writes a
readable `.py` into `gen_py` with real method names, argument order, and
argument count — it just cannot itself be used for dispatch.

```python
import inspect
mod = win32com.client.gencache.EnsureModule(clsid, 0, major, minor)
inspect.signature(mod.IFeatureManager.FeatureCut4)     # real names and order
"Status" in mod.IEquationMgr._prop_map_get_            # True: a property, not a method
```

**Generating the main SldWorks typelib this way leaves that module cached, and
from then on plain `Dispatch` on that machine is early-bound**. See
[above](#a-gen_py-cache-silently-switches-to-early-binding). Code that goes
through `connect()` is unaffected. The constants typelib holds only enums, so
caching it (as `constants_module()` does) switches nothing.

The typelib's split into properties and methods does not predict which
zero-arg members auto-invoke. Every `Get...` member listed above is declared
as a method.

Find CLSIDs with `win32com.client.selecttlb.EnumTlbs()`.

| Typelib | CLSID | major/minor |
|---|---|---|
| SldWorks 2026 Type Library | `{83A33D31-27C5-11CE-BFD4-00400513BB57}` | 34 / 0 |
| SOLIDWORKS 2026 Constant type library | `{4687F359-55D0-4CD3-B6CF-2EB42C11F989}` | 34 / 0 |

Enum constants live in the **separate** constants typelib: `swEndConditions_e`,
`swSelectType_e`, `swFeatureSuppressionAction_e`, `swConstrainedStatus_e`,
`swInConfigurationOpts_e`, `swFeatureTreeFolderType_e`,
`swUserPreferenceToggle_e`, `swUserPreferenceStringValue_e`, and the rest.

**Resolve enum values against the installed constants typelib rather than a
remembered integer.** `swDefaultTemplatePart` is `8` here, not the `4` that was
assumed; `swConstrainedStatus_e` is `1`=unknown, `2`=under, `3`=fully,
`4`=over, not the `1/2/3` originally guessed.
`sw_helpers.constants_module()` wraps this.

## Units

All lengths in the API are **metres** and all angles **radians**, regardless of
document units. `sw_helpers.mm()` and `sw_helpers.deg()` convert.

**Equation text is the exception.** Values there are in document units, and
the trig functions take and return **degrees**: `"x" = tan(45) + atn(1)`
evaluates to `46`. `sin`, `cos`, `tan`, `atn`, `arcsin` and `sqr` (square root)
all worked in global-variable equations and matched Python to 4 decimal places.

Sketch `Create*` arguments are **not** global `(X, Y, Z)`: their mapping to
model space depends on the plane, and on two planes includes a sign flip. Read
it from the sketch's own transform. See
[api-recipes.md](api-recipes.md#where-sketch-coordinates-land).

## Method-location surprises

None of these are guessable by pattern-matching from nearby calls — each "looks
right" and is wrong:

- **`ICurve.CircleParams` is a property**, with no `Get` prefix. It returns
  `(cx, cy, cz, nx, ny, nz, r)`. `curve.GetCircleParams()` raises
  `AttributeError`. The typelib declares it as a property, unlike the
  neighbouring `IsCircle` / `IsLine`, which are zero-arg methods that auto-invoke.
- **The 6-arg `SaveAs(Name, Version, Options, ExportData, Errors, Warnings)` is
  declared on `IModelDocExtension`** in the typelib. `IModelDoc2.SaveAs` takes
  only `NewName`. Check which object you are actually calling before debugging
  its arguments.

- **`InsertSketch2` does not exist** via dynamic dispatch on `SketchManager` on
  this build. Use `InsertSketch(bool)` — same effect, toggles sketch edit mode.
- **`InsertAxis2(AutoSize)` lives on `IModelDoc2`**, not on `FeatureManager`,
  despite every other feature-creation call living on `FeatureManager`.
- **`AddDimension2(x, y, z)` lives on `IModelDoc2`**, not on
  `ModelDocExtension` — easy to guess wrong, since `SelectByID2` and
  `GetWhatsWrongCount` both live on `Extension`.
- **`IEquationMgr.Status` is a property with no index**, unlike its neighbours
  `Equation(i)` and `Value(i)`. `eq.Status(i)` raises `'int' object is not
  callable`. It reports on the equation most recently evaluated, so call
  `eq.Value(i)` and then read `eq.Status`; `-1` marks a broken equation.
- **`SaveAs3(Name, Version, Options)` on `IModelDoc2`** is the save-to-new-name
  call that works: no ByRef out-parameters. In `swSaveAsOptions_e`, `1` is
  Silent and `2` is Copy, which opens a modal dialog.
- **`MathUtility.CreatePoint(...)` raises `Member not found`** under late binding
  with a list, a tuple or a `VT_ARRAY | VT_R8` VARIANT. Apply
  `ModelToSketchTransform.ArrayData` in Python instead (`sw_helpers.sketch_frame`).
- **`GetBodies2` and `GetPartBox` are declared on `IPartDoc`**, not `IModelDoc2`.
  They work on a part's `doc` anyway, because late binding resolves names on the
  real object. Look under `IPartDoc` when reading their signatures.

## `SelectByID2` type strings

The `Type` string is its own small vocabulary, unrelated to `GetTypeName2()`:

| To select | Type string |
|---|---|
| A feature | `"BODYFEATURE"` — not `GetTypeName2()`'s value (e.g. `"Extrusion"`), not `"FEAT"`, not `""` |
| A feature-tree folder | `"FTRFOLDER"` — `"BODYFEATURE"` raises `SelectByID2 failed` |
| A sketch | `"SKETCH"` |
| A reference plane | `"PLANE"` |
| The origin, from inside a sketch | `"EXTSKETCHPOINT"` with the name `"Point1@Origin"`: `sw_helpers.select_origin` |
| The sketch origin (older form, unreliable) | `"ORIGIN"`, with an **empty** name — type/coordinate based, not name based |

**Select the origin by name:**
`ext.SelectByID2("Point1@Origin", "EXTSKETCHPOINT", 0, 0, 0, append, 0, none_dispatch(), 0)`.
It selected type `25` (`swSelEXTSKETCHPOINTS`) and worked for coincident and
point-alignment relations and for horizontal and vertical dimensions, in 7
sketches on default planes and on offset planes that do not pass through the
origin. `("", "EXTSKETCHPOINT")` also selected type 25. Like `"Front Plane"`,
the name is a feature name and may differ on a localised install.

The empty-name `"ORIGIN"` form has two recorded failures. In a sketch cluttered
with overlapping prior test dimensions it picked up a *dimension*
(`swSelDIMENSIONS`, type `14`) instead of the origin point (`swSelSKETCHPOINTS`,
type `11`). In a fresh, uncluttered Right Plane sketch it raised
`SelectByID2 failed` outright. It did work for the center-rectangle recipe on
Front Plane. Check `doc.SelectionManager.GetSelectedObjectType3(index, -1)` when
a selection-dependent call mysteriously returns `None`.
