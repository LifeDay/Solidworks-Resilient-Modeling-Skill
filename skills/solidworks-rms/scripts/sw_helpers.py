#!/usr/bin/env python
"""
sw_helpers.py - the SolidWorks COM helpers this skill's recipes assume.

Every code sample in references/api-recipes.md calls into this module rather
than redefining `call0`, `none_dispatch`, `select_by_id2` and friends inline.
Re-deriving them per script is how two SolidWorks crashes happened; import
them instead.

    import sys; sys.path.insert(0, r"<skill dir>/scripts")   # absolute, not cwd-relative
    from sw_helpers import connect, call0, select_by_id2, mm, no_input_dim_dialog

Everything here is calibrated against SW2026 SP1.1 (rev 34.1.1) under *late*
binding, which connect() forces. On a machine with a gen_py cache for the
SldWorks typelib, plain win32com.client.Dispatch / GetActiveObject silently
return early-bound objects instead, and every bare-getter call below breaks.
See capabilities.yaml for what is verified on which version, and
references/troubleshooting.md for the symptom each guard exists to prevent.
"""

import math
import os
import time

try:
    import pythoncom
    import win32com.client
    from win32com.client import VARIANT, dynamic
except ImportError:  # pragma: no cover - environment problem, not a code path
    raise SystemExit("pywin32 is required:  pip install pywin32")


# SW2026 typelibs. Find others with win32com.client.selecttlb.EnumTlbs().
TLB_MAIN = "{83A33D31-27C5-11CE-BFD4-00400513BB57}"
TLB_CONSTANTS = "{4687F359-55D0-4CD3-B6CF-2EB42C11F989}"

# swUserPreferenceToggle_e. Verify against the live constants typelib with
# constants_module().swInputDimValOnCreate rather than trusting this number.
SW_INPUT_DIM_VAL_ON_CREATE = 10

# swFeatureTreeFolderType_e: wrap the current selection rather than create empty.
FOLDER_CONTAINING = 2

# swFeatureFilletType_e / options. A plain constant-radius round needs the
# uniform-radius bit set; Options=0 returns None.
FILLET_UNIFORM_RADIUS = 2
FILLET_TYPE_SIMPLE = 0

# swSaveAsOptions_e / swSaveAsVersion_e. Silent is 1; 2 is Copy, which opens a
# modal Save As dialog and leaves the script blocked.
SAVE_AS_SILENT = 1
SAVE_AS_CURRENT_VERSION = 0

# swCustomInfoType_e text, swCustomPropertyAddOption_e replace-existing.
CUSTOM_INFO_TEXT = 30
CUSTOM_PROPERTY_REPLACE = 2


# --------------------------------------------------------------------------
# Dynamic dispatch
# --------------------------------------------------------------------------

def call0(obj, name):
    """Read a zero-argument COM 'method' that late binding resolves eagerly.

    On this build essentially every zero-arg Get...-style member auto-invokes
    on bare attribute access: GetTypeName2, GetNextFeature, GetEquationMgr,
    GetCount, GetTitle, GetConstrainedStatus, ActiveSketch, GetWhatsWrongCount.

    There is deliberately no fallback call. A *resolved* CDispatch is itself
    callable, so `callable(value)` cannot distinguish "still needs calling"
    from "already resolved" - calling again raises TypeError or COM
    -2147352573 'Member not found'. Any wrapper that tries to be clever here
    is wrong; this one-liner is the whole correct implementation.

    It only holds for a late-bound `obj`. If this returns a bound method, the
    object is early-bound - get it through connect() or late_bound().
    """
    return getattr(obj, name)


def late_bound(obj):
    """Re-wrap a COM object for pure late binding.

    win32com.client.Dispatch and GetActiveObject return a makepy-generated,
    early-bound class whenever gen_py holds a module for the SldWorks typelib -
    and gencache.EnsureModule on that typelib is one way such a module appears.
    Early-bound, a bare `doc.GetTitle` is a bound method rather than the title,
    so every recipe in this skill breaks.

    pywin32's dynamic.CDispatch wraps the objects its calls return with
    dynamic.Dispatch, so re-wrapping the Application object is enough; objects
    reached from it are late-bound too. Wrapping again is harmless.
    """
    if obj is None or type(obj) is dynamic.CDispatch:
        return obj
    return dynamic.Dispatch(obj._oleobj_)


def none_dispatch():
    """A typed null for an Object/IDispatch parameter (e.g. SelectByID2's Callout).

    A bare Python None raises DISP_E_TYPEMISMATCH on these.
    """
    return VARIANT(pythoncom.VT_DISPATCH, None)


def none_empty():
    """A typed empty for parameters that reject even VT_DISPATCH null.

    Needed by EquationMgr.Add3's trailing ConfigurationName, where None, ""
    and [] all raise DISP_E_TYPEMISMATCH.
    """
    return VARIANT(pythoncom.VT_EMPTY, None)


def constants_module():
    """Return the generated constants typelib module, or None if unavailable.

    Use `constants_module().constants.swSomeEnum_e_Member` to resolve an enum
    against the *installed* SolidWorks rather than a remembered integer - the
    values drift, and swDefaultTemplatePart was 8 here where 4 was expected.

    Note this uses EnsureModule, not EnsureDispatch: EnsureDispatch (and CastTo,
    which calls it) fails with 'Element not found' on SW2026 and must not be
    used for actual dispatch. The constants typelib holds only enums, so caching
    it does not switch any object to early binding; caching TLB_MAIN does.
    """
    try:
        return win32com.client.gencache.EnsureModule(TLB_CONSTANTS, 0, 34, 0)
    except Exception:
        return None


# --------------------------------------------------------------------------
# Units - the API is metres and radians regardless of document units
# --------------------------------------------------------------------------

def mm(value):
    """Millimetres to metres."""
    return value / 1000.0


def deg(value):
    """Degrees to radians. (Equation *text* is the exception: its trig is in degrees.)"""
    return value * math.pi / 180.0


# --------------------------------------------------------------------------
# Connecting, and not clobbering the user's session
# --------------------------------------------------------------------------

def connect(visible=True, launch=False):
    """Attach to the running SolidWorks and return a late-bound Application.

    This attaches to the user's *real, running* session - there is no sandbox.
    Anything mutating must go through assert_scratch_doc / safe_close below.

    GetActiveObject attaches or fails honestly. Dispatch starts an invisible
    SolidWorks that never registers in the running object table, so it is used
    only with launch=True.
    """
    try:
        app = win32com.client.GetActiveObject("SldWorks.Application")
    except pythoncom.com_error:
        if not launch:
            raise RuntimeError(
                "SolidWorks is not running, or not COM-attachable. Run "
                "scripts/sw_preflight.py, or pass launch=True to start one.")
        app = win32com.client.Dispatch("SldWorks.Application")
    sw = late_bound(app)
    sw.Visible = visible
    return sw


def open_titles(sw):
    """Titles of every currently open document.

    GetDocuments returns None, not an empty collection, when nothing is open.
    Capture this *before* creating anything so the result can be passed to
    assert_scratch_doc as known_user_titles.
    """
    docs = call0(sw, "GetDocuments") or ()
    return [call0(d, "GetTitle") for d in docs]


def assert_scratch_doc(doc, known_user_titles):
    """Refuse to proceed unless `doc` is demonstrably not a user document.

    NewPart() has been observed returning the user's existing ActiveDoc instead
    of a new part. Re-read the title now rather than trusting a handle or a
    title captured earlier in the script.
    """
    if doc is None:
        raise RuntimeError("No document - refusing to operate on ActiveDoc by default.")
    title = call0(doc, "GetTitle")
    if title in set(known_user_titles):
        raise RuntimeError(
            "Refusing to touch {!r}: it was already open before this script started. "
            "This is a user document, not a scratch document.".format(title))
    return title


def safe_close(sw, doc, known_user_titles):
    """Close `doc` only after re-verifying, in the same breath, what it is.

    A title captured earlier is not evidence of what a handle still points to;
    an earlier-captured title is exactly what closed a user's unsaved document
    in one observed session.
    """
    title = assert_scratch_doc(doc, known_user_titles)
    sw.CloseDoc(title)
    return title


def tag_doc(doc, key, value):
    """Stamp `doc` with a document-level custom property identifying the build.

    A title is not a durable identity: SaveAs renames the document, and a
    multi-process build cannot trust a title remembered from an earlier stage.
    A tag can be re-read immediately before every mutating call (verify_tag)
    and used to find the document again (find_tagged_doc).
    """
    doc.Extension.CustomPropertyManager("").Add3(
        key, CUSTOM_INFO_TEXT, value, CUSTOM_PROPERTY_REPLACE)
    if doc_tag(doc, key) != value:
        raise RuntimeError("custom property {!r} did not take".format(key))


def doc_tag(doc, key):
    """The document-level custom property `key`, or None if it is absent."""
    try:
        return doc.Extension.CustomPropertyManager("").Get(key) or None
    except Exception:
        return None


def verify_tag(doc, known_user_titles, key, value):
    """assert_scratch_doc plus a tag check. Call it in the same breath as the mutation."""
    title = assert_scratch_doc(doc, known_user_titles)
    if doc_tag(doc, key) != value:
        raise RuntimeError(
            "{!r} does not carry {}={!r} - refusing to mutate it.".format(title, key, value))
    return title


def find_tagged_doc(sw, key, value, known_user_titles):
    """The open document carrying key=value, never one of the user's own.

    Lets each stage of a build run as a separate process and locate the build
    document by tag rather than by a remembered title.
    """
    for doc in call0(sw, "GetDocuments") or ():
        if call0(doc, "GetTitle") in set(known_user_titles):
            continue
        if doc_tag(doc, key) == value:
            verify_tag(doc, known_user_titles, key, value)
            return doc
    raise RuntimeError("No open document carries {}={!r}.".format(key, value))


def save_as(doc, path, known_user_titles):
    """Save `doc` under a new name, silently, and return its new title.

    SaveAs3(Name, Version, Options) has no ByRef out-parameters, so it works
    under late binding where the 6-arg SaveAs and OpenDoc6 raise Type mismatch.
    Options must be swSaveAsOptions_Silent (1): 2 is Copy, which opens a modal
    Save As dialog - after the file has already been written.

    SaveAs renames the document in place, so it is a mutating call and goes
    through the identity check. It refuses to overwrite an existing file.
    """
    path = os.path.abspath(path)
    if os.path.exists(path):
        raise RuntimeError("Refusing to overwrite existing {!r}.".format(path))
    assert_scratch_doc(doc, known_user_titles)
    err = doc.SaveAs3(path, SAVE_AS_CURRENT_VERSION, SAVE_AS_SILENT)
    if err != 0:
        raise RuntimeError("SaveAs3 returned {} for {!r}".format(err, path))
    return call0(doc, "GetTitle")


# --------------------------------------------------------------------------
# Selection
# --------------------------------------------------------------------------

def select_by_id2(ext, name, typ, append=False, x=0.0, y=0.0, z=0.0, mark=0):
    """SelectByID2 with Callout wrapped as a typed null.

    Type strings that are not guessable from sibling calls:
      "BODYFEATURE"     a feature (NOT GetTypeName2()'s value, not "FEAT", not "")
      "FTRFOLDER"       a feature-tree folder (raises SelectByID2 failed on BODYFEATURE)
      "SKETCH"          a sketch
      "PLANE"           a reference plane
      "EXTSKETCHPOINT"  the origin from inside a sketch, as "Point1@Origin" -
                        use select_origin(). The older ("", "ORIGIN") form is
                        coordinate based: it has picked up a dimension in a
                        cluttered sketch and raised outright in a clean one.
    """
    ok = ext.SelectByID2(name, typ, x, y, z, append, mark, none_dispatch(), 0)
    if not ok:
        raise RuntimeError("SelectByID2 failed for {!r} as {!r}".format(name, typ))
    return ok


def select_origin(ext, append=False):
    """Select the part origin from inside a sketch, for a relation or dimension.

    ("", "ORIGIN") raised SelectByID2 failed in a fresh Right Plane sketch on
    SW2026. This name-based form worked in 7 sketches on default and offset
    planes, selecting type 25 (swSelEXTSKETCHPOINTS). Like "Front Plane", the
    name is a feature name and may differ on a localised install.
    """
    return select_by_id2(ext, "Point1@Origin", "EXTSKETCHPOINT", append=append)


def selected_count(doc):
    """How many entities are actually selected, for verifying a multi-pick."""
    return doc.SelectionManager.GetSelectedObjectCount2(-1)


def selected_type(doc, index=1):
    """swSelectType_e of a selection, for diagnosing a wrong pick.

    11 = swSelSKETCHPOINTS, 14 = swSelDIMENSIONS, 25 = swSelEXTSKETCHPOINTS
    (the origin via select_origin). If a selection-dependent call mysteriously
    returns None, check this before suspecting the call itself.
    """
    return doc.SelectionManager.GetSelectedObjectType3(index, -1)


def circular_edges(body, center, radius, tol=1e-6):
    """Edges of `body` that are circles of `radius` centred on `center`.

    `center` is a model-space (x, y, z) tuple in metres, computed from the
    parameter set. Sketch Create* arguments are NOT global coordinates, so map
    them first - see sketch_to_model and references/api-recipes.md.

    Coordinate-guess SelectByID2("", "EDGE", x, y, z) picks were unreliable for
    fillet edges; matching each circular edge's CircleParams was reliable.
    CircleParams is a bare property returning (cx, cy, cz, nx, ny, nz, r) -
    there is no GetCircleParams() method, and calling one raises AttributeError.

        for edge in circular_edges(body, (0, y, z), mm(4)):
            edge.Select4(True, none_dispatch())
    """
    found = []
    for edge in body.GetEdges() or ():   # body.GetEdges does not auto-invoke; face.GetEdges does
        curve = edge.GetCurve
        if not curve.IsCircle:
            continue
        cx, cy, cz, _nx, _ny, _nz, r = curve.CircleParams
        if (abs(r - radius) <= tol
                and abs(cx - center[0]) <= tol
                and abs(cy - center[1]) <= tol
                and abs(cz - center[2]) <= tol):
            found.append(edge)
    return found


# --------------------------------------------------------------------------
# Sketch coordinates
# --------------------------------------------------------------------------

def sketch_frame(doc):
    """The active sketch's model-to-sketch transform, as (R, t, s).

    ModelToSketchTransform.ArrayData is 16 doubles: a 3x3 rotation, a
    translation, a scale, then 3 unused. Applying it in Python is the working
    route - MathUtility.CreatePoint(...).MultiplyTransform raised 'Member not
    found' under late binding with a list, a tuple and a VT_ARRAY|VT_R8 VARIANT.

    Read this rather than a per-plane table: it covers every plane, including
    offset planes and part templates nobody has probed. Only valid while the
    sketch is open for editing.
    """
    sketch = call0(doc.SketchManager, "ActiveSketch")
    if sketch is None:
        raise RuntimeError("No active sketch - call sketch_frame while editing one.")
    a = sketch.ModelToSketchTransform.ArrayData
    return [a[0:3], a[3:6], a[6:9]], a[9:12], a[12]


def model_to_sketch(frame, point, tol=1e-9):
    """Model (x, y, z) in metres -> sketch (u, v, w) in metres, for Create* calls.

    p' = s * (p . R) + t, with p as a row vector. Raises if the point is not on
    the sketch plane (|w| > tol), which catches a point computed for the wrong
    plane before it becomes wrong geometry.
    """
    R, t, s = frame
    out = tuple(s * sum(point[i] * R[i][j] for i in range(3)) + t[j] for j in range(3))
    if abs(out[2]) > tol:
        raise ValueError("model point {} is off the sketch plane (w={:g})".format(point, out[2]))
    return out


def sketch_to_model(frame, u, v):
    """Sketch (u, v) in metres -> model (x, y, z) in metres. Inverse of model_to_sketch.

    p = ((p' - t) / s) . R^T. Dimension placement points go through this.
    """
    R, t, s = frame
    q = ((u - t[0]) / s, (v - t[1]) / s, (0.0 - t[2]) / s)
    return tuple(sum(q[j] * R[i][j] for j in range(3)) for i in range(3))


def runs_horizontal(frame, direction):
    """True if a model-space direction runs horizontally in the open sketch.

    Choose sgHORIZONTALPOINTS2D vs sgVERTICALPOINTS2D, and
    AddHorizontalDimension2 vs AddVerticalDimension2, from this - not from the
    plane's name. Global X runs vertically in a Top Plane sketch on the test
    template, and the wrong relation drags the geometry.
    """
    R, _t, s = frame
    u, v = (s * sum(direction[i] * R[i][j] for i in range(3)) for j in range(2))
    return abs(u) > abs(v)


# --------------------------------------------------------------------------
# Tree traversal
# --------------------------------------------------------------------------

def iter_features(doc):
    """Yield every feature in flat tree order.

    A folder contributes a synthetic "...___EndTag___" marker feature after its
    contents. It keeps the folder's default name ("Folder1___EndTag___") even
    after the folder is renamed. It is FtrFolder-typed, so type-based filters
    skip it naturally - do not special-case it away.
    """
    feat = call0(doc, "FirstFeature")
    while feat:
        yield feat
        feat = call0(feat, "GetNextFeature")


def last_feature(doc):
    """The feature most recently added to the tree.

    The reliable way to get a handle on what you just created when the
    creating call did not return one (e.g. after InsertSketch).
    """
    result = None
    for feat in iter_features(doc):
        result = feat
    return result


def feature_by_name(doc, name):
    for feat in iter_features(doc):
        if feat.Name == name:
            return feat
    return None


# --------------------------------------------------------------------------
# Equations and global variables
# --------------------------------------------------------------------------

def add_equation(eq, text, index=-1):
    """Add a global variable or driven dimension, and verify it landed.

    Uses Add2. Add3 is the documented primary call but silently fails on
    SW2026 SP1.1 - returns -1, raises nothing, adds no equation - independent
    of ConfigurationOption. The count check is the point of this wrapper: a
    silent no-op is the failure mode being guarded against.

    `index` -1 appends; any other value inserts at that position, which is how
    to put a deleted dimension equation back where it was. Returns the index.

        add_equation(eq, '"plate_width" = 120')
        add_equation(eq, '"D1@Sketch1" = "plate_width"')
    """
    before = call0(eq, "GetCount")
    eq.Add2(index, text, True)
    after = call0(eq, "GetCount")
    if after <= before:
        raise RuntimeError(
            "Equation {!r} was not added (count stayed at {}). If it redefines a "
            "global that was deleted while dimensions referenced it, or this build "
            "needs Add3, see references/troubleshooting.md.".format(text, before))
    return after - 1 if index < 0 else index


def equations(eq):
    """Every equation as (index, text, value, status), for verifying what landed.

    Status is a property with no index argument: it reports on the equation
    most recently evaluated, so Value(i) is read first. -1 marks a broken
    equation. Value and status are both None when evaluating it raised, since
    Status would then still describe the previous equation.
    """
    rows = []
    for i in range(call0(eq, "GetCount")):
        try:
            value = eq.Value(i)
            status = call0(eq, "Status")
        except Exception:
            value = status = None
        rows.append((i, eq.Equation(i), value, status))
    return rows


# --------------------------------------------------------------------------
# Folders
# --------------------------------------------------------------------------

def wrap_in_folder(doc, feature_names, folder_name, types=None):
    """Wrap already-created features in a named folder.

    Selects each feature, then InsertFeatureTreeFolder2(2) - the "Containing"
    type, which wraps the current selection rather than creating an empty
    folder - and renames the result.

    `types` maps a feature name to its SelectByID2 type string; anything absent
    defaults to "BODYFEATURE". Sketches need "SKETCH", reference planes "PLANE".

    Call this as each RMS group completes. SolidWorks folders must hold
    contiguous features and will not reorder past a dependency, so there is no
    sort-the-tree-afterwards path.

    Note MoveToFolder is not a usable alternative on this build: it silently
    does nothing when passed names and raises DISP_E_TYPEMISMATCH when passed
    IFeature objects. To add a feature to an existing folder, dissolve the
    folder (EditDelete on an "FTRFOLDER" selection) and re-wrap.
    """
    types = types or {}
    ext = doc.Extension
    doc.ClearSelection2(True)
    for i, name in enumerate(feature_names):
        select_by_id2(ext, name, types.get(name, "BODYFEATURE"), append=(i > 0))
    folder = doc.FeatureManager.InsertFeatureTreeFolder2(FOLDER_CONTAINING)
    if folder is None:
        raise RuntimeError("InsertFeatureTreeFolder2 returned None for {!r}".format(folder_name))
    folder.Name = folder_name
    return folder


# --------------------------------------------------------------------------
# The modal-dialog guard
# --------------------------------------------------------------------------

class no_input_dim_dialog(object):
    """Disable System Options -> "Input dimension value" for the duration.

    With it enabled, every API call that creates a dimension (AddDimension2,
    AddHorizontalDimension2, AddVerticalDimension2) pops a modal Modify dialog
    and the script appears to hang while SolidWorks itself stays Responding.

    The expensive part is what follows: killing the Python client while that
    modal sits open leaves the session in a state where the next Save() or
    SaveAs() crashes SolidWorks outright (RPC -2147023170). Reproduced three
    times. Wrap any scripted dimensioning in this and the failure cannot arise.

    This is an App-level global preference, not document-scoped, so it is
    restored on exit for the user's normal interactive work.

        with no_input_dim_dialog(sw):
            doc.AddHorizontalDimension2(x, y, 0)
    """

    def __init__(self, sw, toggle=None):
        self.sw = sw
        self.toggle = toggle if toggle is not None else _input_dim_toggle()
        self.previous = None

    def __enter__(self):
        self.previous = self.sw.GetUserPreferenceToggle(self.toggle)
        self.sw.SetUserPreferenceToggle(self.toggle, False)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.previous is not None:
            self.sw.SetUserPreferenceToggle(self.toggle, self.previous)
        return False


def _input_dim_toggle():
    """Resolve swInputDimValOnCreate from the live typelib, falling back to 10."""
    mod = constants_module()
    if mod is not None:
        try:
            return mod.constants.swInputDimValOnCreate
        except Exception:
            pass
    return SW_INPUT_DIM_VAL_ON_CREATE


def settle(seconds=0.3):
    """Let SolidWorks catch up after a UI-mutating call.

    A plain sleep, deliberately. Do NOT add pythoncom.PumpWaitingMessages() -
    pumping messages while SolidWorks is mid-operation deadlocked the *client*
    thread in one session while SolidWorks itself stayed responsive.
    """
    time.sleep(seconds)


# --------------------------------------------------------------------------
# Rebuild state and geometry reads
# --------------------------------------------------------------------------

def rebuild_errors(doc):
    """Number of outstanding what's-wrong entries after a rebuild."""
    return call0(doc.Extension, "GetWhatsWrongCount")


def force_rebuild(doc):
    """Rebuild and return the error count. Zero errors is necessary, not sufficient -
    a geometrically wrong model rebuilds cleanly (see troubleshooting.md on
    GetCenterPoint2)."""
    doc.ForceRebuild3(False)
    return rebuild_errors(doc)


def volume(doc):
    """Solid volume of the part in cubic metres. A cheap check after every feature."""
    return doc.Extension.CreateMassProperty.Volume


def body_extents(body):
    """((xmin, ymin, zmin), (xmax, ymax, zmax)) of a body in metres, from tessellation.

    doc.GetPartBox(True) pads its result (42.196 against a true 42.000 after
    filleting), and body.GetBodyBox() returned None on a cut body. Tessellation
    vertices lie on the real surfaces, so these extents never exceed the part;
    where the extreme is on a curved face they can fall short of it by the
    tessellation chord tolerance.
    """
    lo, hi = [math.inf] * 3, [-math.inf] * 3
    for face in body.GetFaces() or ():           # GetFaces needs parens
        coords = face.GetTessTriangles(True) or ()
        for k in range(0, len(coords) - 2, 3):
            for axis in range(3):
                c = coords[k + axis]
                lo[axis] = min(lo[axis], c)
                hi[axis] = max(hi[axis], c)
    return tuple(lo), tuple(hi)
