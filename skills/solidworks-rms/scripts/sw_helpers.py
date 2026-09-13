#!/usr/bin/env python
"""
sw_helpers.py - the SolidWorks COM helpers this skill's recipes assume.

Every code sample in references/api-recipes.md calls into this module rather
than redefining `call0`, `none_dispatch`, `select_by_id2` and friends inline.
Re-deriving them per script is how two SolidWorks crashes happened; import
them instead.

    import sys; sys.path.insert(0, r"<skill dir>/scripts")   # absolute, not cwd-relative
    from sw_helpers import connect, call0, select_by_id2, mm, no_input_dim_dialog

Everything here is calibrated against SW2026 SP1.1 (rev 34.1.1) via plain
late-bound `win32com.client.Dispatch`. See capabilities.yaml for what is
verified on which version, and references/troubleshooting.md for the symptom
each guard exists to prevent.
"""

import time

try:
    import pythoncom
    import win32com.client
    from win32com.client import VARIANT
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
    """
    return getattr(obj, name)


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
    used for actual dispatch.
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
    """Degrees to radians."""
    import math
    return value * math.pi / 180.0


# --------------------------------------------------------------------------
# Connecting, and not clobbering the user's session
# --------------------------------------------------------------------------

def connect(visible=True):
    """Attach to SolidWorks. Returns the Application object.

    This attaches to the user's *real, running* session - there is no sandbox.
    Anything mutating must go through assert_scratch_doc / safe_close below.
    """
    sw = win32com.client.Dispatch("SldWorks.Application")
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


# --------------------------------------------------------------------------
# Selection
# --------------------------------------------------------------------------

def select_by_id2(ext, name, typ, append=False, x=0.0, y=0.0, z=0.0, mark=0):
    """SelectByID2 with Callout wrapped as a typed null.

    Type strings that are not guessable from sibling calls:
      "BODYFEATURE"  a feature (NOT GetTypeName2()'s value, not "FEAT", not "")
      "FTRFOLDER"    a feature-tree folder (raises SelectByID2 failed on BODYFEATURE)
      "SKETCH"       a sketch
      "PLANE"        a reference plane
      "ORIGIN"       the sketch origin - pass name="" ; this is type/coordinate
                     based, so in a sketch cluttered with prior test dimensions
                     it can pick up a dimension instead. Work in a fresh sketch.
    """
    ok = ext.SelectByID2(name, typ, x, y, z, append, mark, none_dispatch(), 0)
    if not ok:
        raise RuntimeError("SelectByID2 failed for {!r} as {!r}".format(name, typ))
    return ok


def selected_count(doc):
    """How many entities are actually selected, for verifying a multi-pick."""
    return doc.SelectionManager.GetSelectedObjectCount2(-1)


def selected_type(doc, index=1):
    """swSelectType_e of a selection, for diagnosing a wrong pick.

    11 = swSelSKETCHPOINTS, 14 = swSelDIMENSIONS. If a selection-dependent call
    mysteriously returns None, check this before suspecting the call itself.
    """
    return doc.SelectionManager.GetSelectedObjectType3(index, -1)


def circular_edges(body, center, radius, tol=1e-6):
    """Edges of `body` that are circles of `radius` centred on `center`.

    `center` is a model-space (x, y, z) tuple in metres, computed from the
    parameter set. Sketch Create* arguments are NOT global coordinates, so map
    them first - see references/api-recipes.md, "Where sketch coordinates land".

    Coordinate-guess SelectByID2("", "EDGE", x, y, z) picks were unreliable for
    fillet edges; matching each circular edge's CircleParams was reliable.
    CircleParams is a bare property returning (cx, cy, cz, nx, ny, nz, r) -
    there is no GetCircleParams() method, and calling one raises AttributeError.

        for edge in circular_edges(body, (0, y, z), mm(4)):
            edge.Select4(True, none_dispatch())
    """
    found = []
    for edge in body.GetEdges() or ():   # GetEdges does not auto-invoke
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
# Tree traversal
# --------------------------------------------------------------------------

def iter_features(doc):
    """Yield every feature in flat tree order.

    A folder contributes a synthetic "<FolderName>___EndTag___" marker feature
    after its contents. It is FtrFolder-typed, so type-based filters skip it
    naturally - do not special-case it away.
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

def add_equation(eq, text):
    """Add a global variable or driven dimension, and verify it landed.

    Uses Add2. Add3 is the documented primary call but silently fails on
    SW2026 SP1.1 - returns -1, raises nothing, adds no equation - independent
    of ConfigurationOption. The count check is the point of this wrapper: a
    silent no-op is the failure mode being guarded against.

        add_equation(eq, '"plate_width" = 120')
        add_equation(eq, '"D1@Sketch1" = "plate_width"')
    """
    before = call0(eq, "GetCount")
    eq.Add2(-1, text, True)
    after = call0(eq, "GetCount")
    if after <= before:
        raise RuntimeError(
            "Equation {!r} was not added (count stayed at {}). "
            "If this build needs Add3, see references/troubleshooting.md.".format(text, before))
    return after - 1


def equations(eq):
    """Every equation as (index, text, status), for verifying what actually landed."""
    return [(i, eq.Equation(i), eq.Status(i)) for i in range(call0(eq, "GetCount"))]


# --------------------------------------------------------------------------
# Folders
# --------------------------------------------------------------------------

def wrap_in_folder(doc, feature_names, folder_name, types=None):
    """Wrap already-created features in a named folder.

    Selects each feature, then InsertFeatureTreeFolder2(2) - the "Containing"
    type, which wraps the current selection rather than creating an empty
    folder - and renames the result.

    `types` maps a feature name to its SelectByID2 type string; anything absent
    defaults to "BODYFEATURE". Sketches need "SKETCH".

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
# Rebuild state
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
