#!/usr/bin/env python
"""
rms_check.py - Resilient Modeling Strategy compliance checker for SolidWorks.

Connects to a running SolidWorks session, walks the active part document, and
reports per-rule PASS / FAIL / WARN.

    python rms_check.py                    # check the active document
    python rms_check.py --dump-types       # print every feature + API type name
    python rms_check.py --no-suppress      # skip the suppression test (faster, non-mutating)
    python rms_check.py --exceptions rms_exceptions.json

STATUS: zero-arg getter calls (GetTypeName2, GetNextFeature, GetChildren,
GetConstrainedStatus, GetEquationMgr, GetCount, GetWhatsWrongCount, GetTitle,
GetDefinition, GetSpecificFeature2, FirstFeature, GetFirstSubFeature,
GetNextSubFeature) were switched from `.Method()` to bare `.Method` after
confirming against a live SW2026 SP1.1 session that dynamic dispatch
auto-invokes these on attribute access and that calling the resolved value
again raises 'Member not found' or TypeError - see api-notes.md's "Session
log" section. The swConstrainedStatus_e values below were also corrected
against the live constants typelib (1=unknown, 2=under, 3=fully, 4=over -
NOT 1/2/3 as originally guessed). The feature-type classification sets
below are still the most likely remaining thing to need adjustment for a
new SW version - run --dump-types on known-good parts and calibrate.
"""

import argparse
import json
import os
import sys
from collections import OrderedDict

try:
    import win32com.client
except ImportError:
    sys.exit("pywin32 is required:  pip install pywin32")


# --------------------------------------------------------------------------
# Configuration - calibrate these against --dump-types output on real parts
# --------------------------------------------------------------------------

GROUPS = ["1-Ref", "2-Construction", "3-Core", "4-Detail", "5-Modify", "6-Quarantine"]
GROUP_INDEX = {name: i for i, name in enumerate(GROUPS)}

FOLDER_TYPE = "FtrFolder"
SKETCH_TYPES = {"ProfileFeature", "3DProfileFeature"}

# Features that create solid material. Banned from 1-Ref and 2-Construction.
SOLID_TYPES = {
    "Extrusion", "Revolution", "Sweep", "Loft", "Boss", "Thicken",
    "BaseFlange", "Boundary", "ImportedFeature", "ICE",
}

# Features that remove material.
CUT_TYPES = {"Cut", "CutRevolve", "SweepCut", "LoftCut", "Hole", "HoleWzd", "SimpleHole"}

HOLE_TYPES = {"HoleWzd", "SimpleHole", "Hole"}
FILLET_TYPES = {"Fillet"}
CHAMFER_TYPES = {"Chamfer"}
SHELL_TYPES = {"Shell"}
DRAFT_TYPES = {"Draft"}
PATTERN_TYPES = {"LPattern", "CirPattern", "MirrorPattern", "MirrorSolid",
                 "CurvePattern", "TablePattern", "FillPattern", "DerivedLPattern"}

# Reference geometry - always legal in 1-Ref.
REF_TYPES = {"RefPlane", "RefAxis", "CoordSys", "RefPoint", "OriginProfileFeature"}

# Surfaces and curves - legal in 2-Construction.
CONSTRUCTION_TYPES = {
    "SurfaceExtrude", "SurfaceRevolve", "SurfaceSweep", "SurfaceLoft",
    "SurfaceKnit", "SurfaceTrim", "SurfaceExtend", "SurfacePlanar",
    "SurfaceOffset", "SurfaceRuled", "SurfaceRadiate", "SurfaceFill",
    "CompositeCurve", "Reference Curve", "ProjectionCurve", "HelixCurve",
    "SplitLine", "SurfaceCut",
}

# Features SolidWorks creates itself that live outside the group scheme.
TOLERATED_LOOSE = {
    "OriginProfileFeature", "MateGroup", "HistoryFolder", "SensorFolder",
    "DetailCabinet", "CommentFolder", "FavoriteFolder", "SolidBodyFolder",
    "SurfaceBodyFolder", "DocsFolder", "MaterialFolder", "EnvFolder",
    "LightFolder", "EqnFolder", "RefPlane", "OriginPoint", "CoordSys",
}


# --------------------------------------------------------------------------
# Result collection
# --------------------------------------------------------------------------

class Report(object):
    def __init__(self, exceptions):
        self.rows = []
        self.exceptions = exceptions or {}

    def add(self, rule, status, detail=""):
        waived = self.exceptions.get(rule)
        if status == "FAIL" and waived:
            status = "WAIVED"
            detail = "{}  [approved: {}]".format(detail, waived)
        self.rows.append((rule, status, detail))

    def emit(self):
        width = max(len(r[0]) for r in self.rows) + 2
        counts = {"PASS": 0, "FAIL": 0, "WARN": 0, "WAIVED": 0, "SKIP": 0}
        print("")
        for rule, status, detail in self.rows:
            counts[status] = counts.get(status, 0) + 1
            line = "  {:<6} {:<{w}} {}".format(status, rule, detail, w=width)
            print(line.rstrip())
        print("")
        print("  {} passed, {} failed, {} warnings, {} waived".format(
            counts["PASS"], counts["FAIL"], counts["WARN"], counts["WAIVED"]))
        print("")
        return 1 if counts["FAIL"] else 0


# --------------------------------------------------------------------------
# Tree walking
# --------------------------------------------------------------------------

def walk(model):
    """Return an ordered list of (feature, group_name_or_None, depth).

    Flat traversal (FirstFeature/GetNextFeature) surfaces a synthetic
    "<FolderName>___EndTag___" feature immediately after a folder's contents.
    It is itself FtrFolder-typed, so every rule below that matters skips it
    the same way it skips real folders - do not special-case it away.
    """
    out = []
    feat = model.FirstFeature
    current_group = None
    while feat:
        tname = feat.GetTypeName2
        name = feat.Name
        if tname == FOLDER_TYPE and name in GROUP_INDEX:
            current_group = name
            out.append((feat, name, 0))
            for sub in subfeatures(feat):
                out.append((sub, name, 1))
        else:
            out.append((feat, current_group, 0))
        feat = feat.GetNextFeature
    return out


def subfeatures(folder):
    """Yield features inside a folder, tolerating version differences."""
    try:
        sub = folder.GetFirstSubFeature
    except Exception:
        return
    while sub:
        yield sub
        try:
            sub = sub.GetNextSubFeature
        except Exception:
            return


def group_of(feat, entries):
    for f, g, _ in entries:
        if f.Name == feat.Name:
            return g
    return None


def is_content(tname, name):
    """True if this feature is real modelling content we expect to be grouped."""
    if tname == FOLDER_TYPE:
        return False
    if tname in TOLERATED_LOOSE:
        return False
    if name in ("Front Plane", "Top Plane", "Right Plane", "Origin"):
        return False
    return True


# --------------------------------------------------------------------------
# Individual checks
# --------------------------------------------------------------------------

def check_folders(entries, rep):
    found = [g for f, g, d in entries if f.GetTypeName2 == FOLDER_TYPE and d == 0 and g]
    seen = [g for g in GROUPS if g in found]
    missing = [g for g in GROUPS if g not in found]
    if missing:
        rep.add("folders.present", "WARN", "absent: " + ", ".join(missing))
    else:
        rep.add("folders.present", "PASS")

    order = [GROUP_INDEX[g] for g in found if g in GROUP_INDEX]
    rep.add("folders.ordered", "PASS" if order == sorted(order) else "FAIL",
            "" if order == sorted(order) else "folder order: " + " ".join(found))


def check_grouping(entries, rep):
    loose = [f.Name for f, g, d in entries
             if g is None and is_content(f.GetTypeName2, f.Name)]
    rep.add("grouping.all_features_in_a_group",
            "PASS" if not loose else "FAIL",
            "" if not loose else "ungrouped: " + ", ".join(loose[:8]))


def check_no_solids_in_ref(entries, rep):
    bad = [f.Name for f, g, d in entries
           if g in ("1-Ref", "2-Construction")
           and f.GetTypeName2 in (SOLID_TYPES | CUT_TYPES)]
    rep.add("groups.no_solids_in_ref_or_construction",
            "PASS" if not bad else "FAIL",
            "" if not bad else ", ".join(bad))


def check_shell_last_in_core(entries, rep):
    core = [f for f, g, d in entries if g == "3-Core" and f.GetTypeName2 != FOLDER_TYPE]
    shells = [i for i, f in enumerate(core) if f.GetTypeName2 in SHELL_TYPES]
    if not shells:
        rep.add("core.shell_last", "SKIP", "no shell feature")
    elif max(shells) == len(core) - 1:
        rep.add("core.shell_last", "PASS")
    else:
        rep.add("core.shell_last", "FAIL",
                "{} features follow the shell".format(len(core) - 1 - max(shells)))


def check_holes_last_in_detail(entries, rep):
    detail = [f for f, g, d in entries if g == "4-Detail" and f.GetTypeName2 != FOLDER_TYPE]
    holes = [i for i, f in enumerate(detail) if f.GetTypeName2 in HOLE_TYPES]
    if not holes:
        rep.add("detail.holes_last", "SKIP", "no hole features")
        return
    tail = list(range(len(detail) - len(holes), len(detail)))
    rep.add("detail.holes_last", "PASS" if sorted(holes) == tail else "WARN",
            "" if sorted(holes) == tail else "holes interleaved with other detail features")


def check_quarantine_order(entries, rep):
    q = [f for f, g, d in entries if g == "6-Quarantine" and f.GetTypeName2 != FOLDER_TYPE]
    if not q:
        rep.add("quarantine.order", "SKIP", "empty")
        return

    kinds = [f.GetTypeName2 for f in q]
    last_chamfer = max([i for i, k in enumerate(kinds) if k in CHAMFER_TYPES] or [-1])
    first_fillet = min([i for i, k in enumerate(kinds) if k in FILLET_TYPES] or [len(q)])
    rep.add("quarantine.chamfers_before_fillets",
            "PASS" if last_chamfer < first_fillet else "FAIL")

    radii = []
    for f in q:
        if f.GetTypeName2 not in FILLET_TYPES:
            continue
        try:
            data = f.GetDefinition
            radii.append(data.DefaultRadius)
        except Exception:
            radii.append(None)
    known = [r for r in radii if r is not None]
    if len(known) < 2:
        rep.add("quarantine.largest_fillet_first", "SKIP", "fewer than two readable radii")
    else:
        ok = all(known[i] >= known[i + 1] for i in range(len(known) - 1))
        rep.add("quarantine.largest_fillet_first", "PASS" if ok else "FAIL",
                "" if ok else "radii: " + ", ".join("{:.4g}".format(r * 1000) for r in known))

    strays = [f.Name for f in q if f.GetTypeName2 not in (FILLET_TYPES | CHAMFER_TYPES)]
    rep.add("quarantine.only_fillets_and_chamfers",
            "PASS" if not strays else "FAIL",
            "" if not strays else ", ".join(strays))


def check_modify_order(entries, rep):
    m = [f for f, g, d in entries if g == "5-Modify" and f.GetTypeName2 != FOLDER_TYPE]
    if not m:
        rep.add("modify.transform_before_replicate", "SKIP", "empty")
        return
    kinds = [f.GetTypeName2 for f in m]
    last_draft = max([i for i, k in enumerate(kinds) if k in DRAFT_TYPES] or [-1])
    first_pat = min([i for i, k in enumerate(kinds) if k in PATTERN_TYPES] or [len(m)])
    rep.add("modify.transform_before_replicate",
            "PASS" if last_draft < first_pat else "WARN",
            "" if last_draft < first_pat else "a draft follows a pattern")


def build_group_map(entries):
    return {f.Name: g for f, g, d in entries}


def check_reference_direction(model, entries, rep):
    """No feature may depend on a feature in a later group."""
    gmap = build_group_map(entries)
    violations = []
    quarantine_parents = []

    for feat, group, depth in entries:
        if group is None or feat.GetTypeName2 == FOLDER_TYPE:
            continue
        try:
            children = feat.GetChildren
        except Exception:
            children = None
        if not children:
            continue
        for child in children:
            cname = getattr(child, "Name", None)
            if cname is None:
                continue
            cgroup = gmap.get(cname)
            if cgroup is None or cgroup not in GROUP_INDEX or group not in GROUP_INDEX:
                continue
            # child depends on feat; child must sit in the same or a later group
            if GROUP_INDEX[cgroup] < GROUP_INDEX[group]:
                violations.append("{} ({}) -> {} ({})".format(cname, cgroup, feat.Name, group))
            if group == "6-Quarantine":
                quarantine_parents.append("{} depends on {}".format(cname, feat.Name))

    rep.add("refs.direction", "PASS" if not violations else "FAIL",
            "" if not violations else "; ".join(violations[:6]))
    rep.add("refs.quarantine_has_no_children",
            "PASS" if not quarantine_parents else "FAIL",
            "" if not quarantine_parents else "; ".join(quarantine_parents[:6]))


def check_detail_internal_refs(entries, rep):
    """4-Detail features must not reference each other outside a declared subfolder."""
    detail_names = set()
    subfolder_members = set()
    for feat, group, depth in entries:
        if group != "4-Detail":
            continue
        if feat.GetTypeName2 == FOLDER_TYPE and feat.Name != "4-Detail":
            for sub in subfeatures(feat):
                subfolder_members.add(sub.Name)
            continue
        detail_names.add(feat.Name)

    bad = []
    for feat, group, depth in entries:
        if group != "4-Detail" or feat.GetTypeName2 == FOLDER_TYPE:
            continue
        try:
            children = feat.GetChildren or []
        except Exception:
            children = []
        for child in children:
            cname = getattr(child, "Name", None)
            if cname in detail_names and cname != feat.Name:
                if cname in subfolder_members and feat.Name in subfolder_members:
                    continue  # declared coupled pair
                bad.append("{} -> {}".format(cname, feat.Name))

    rep.add("detail.no_internal_references", "PASS" if not bad else "FAIL",
            "" if not bad else "; ".join(bad[:6]))


def check_descriptions(entries, rep):
    missing = []
    for feat, group, depth in entries:
        if feat.GetTypeName2 == FOLDER_TYPE:
            continue
        if not is_content(feat.GetTypeName2, feat.Name):
            continue
        try:
            desc = feat.Description
        except Exception:
            desc = None
        if not desc or not str(desc).strip():
            missing.append(feat.Name)
    rep.add("intent.every_feature_described", "PASS" if not missing else "FAIL",
            "" if not missing else "{} without a description: {}".format(
                len(missing), ", ".join(missing[:8])))


def check_sketches(entries, rep):
    under = []
    over = []
    consumers = {}
    for feat, group, depth in entries:
        if feat.GetTypeName2 not in SKETCH_TYPES:
            continue
        try:
            sk = feat.GetSpecificFeature2
            status = sk.GetConstrainedStatus
        except Exception:
            continue
        # swConstrainedStatus_e (confirmed against the live constants typelib on
        # SW2026 SP1.1 - do not trust a remembered value, re-verify per version):
        # 1 = swUnknownConstraint, 2 = swUnderConstrained, 3 = swFullyConstrained,
        # 4 = swOverConstrained (5-7 are solver-error states, treated as over/bad here).
        if status in (1, 2):
            under.append(feat.Name)
        elif status >= 4:
            over.append(feat.Name)
        try:
            kids = feat.GetChildren or []
        except Exception:
            kids = []
        consumers[feat.Name] = len([k for k in kids if getattr(k, "Name", None)])

    rep.add("sketches.fully_defined", "PASS" if not under else "FAIL",
            "" if not under else "under-defined: " + ", ".join(under[:8]))
    if over:
        rep.add("sketches.not_over_defined", "FAIL", ", ".join(over[:8]))
    else:
        rep.add("sketches.not_over_defined", "PASS")

    shared = [n for n, c in consumers.items() if c > 1]
    rep.add("sketches.one_sketch_per_feature", "PASS" if not shared else "FAIL",
            "" if not shared else "shared: " + ", ".join(shared[:8]))


def check_parameterization(model, rep):
    try:
        eq = model.GetEquationMgr
        count = eq.GetCount
    except Exception:
        rep.add("params.global_variables_present", "SKIP", "EquationMgr unavailable")
        return

    globals_, driven = 0, 0
    for i in range(count):
        text = eq.Equation(i) or ""
        lhs = text.split("=")[0].strip().strip('"')
        if "@" in lhs:
            driven += 1
        else:
            globals_ += 1

    rep.add("params.global_variables_present", "PASS" if globals_ else "FAIL",
            "{} global variables".format(globals_))
    rep.add("params.dimensions_driven_by_equations", "PASS" if driven else "WARN",
            "{} driven dimensions".format(driven))


def check_detail_suppression(model, entries, rep):
    """Each 4-Detail feature must suppress individually without a rebuild error."""
    ext = model.Extension
    targets = [f for f, g, d in entries
               if g == "4-Detail" and f.GetTypeName2 != FOLDER_TYPE]
    if not targets:
        rep.add("detail.individually_suppressible", "SKIP", "no detail features")
        return

    SUPPRESS, UNSUPPRESS = 0, 1       # swFeatureSuppressionAction_e
    THIS_CONFIG = 1                   # swInConfigurationOpts_e

    failures = []
    for feat in targets:
        try:
            feat.SetSuppression2(SUPPRESS, THIS_CONFIG, None)
            model.ForceRebuild3(False)
            if ext.GetWhatsWrongCount > 0:
                failures.append(feat.Name)
        except Exception as exc:
            failures.append("{} ({})".format(feat.Name, exc))
        finally:
            try:
                feat.SetSuppression2(UNSUPPRESS, THIS_CONFIG, None)
            except Exception:
                pass

    try:
        model.ForceRebuild3(False)
    except Exception:
        pass

    rep.add("detail.individually_suppressible", "PASS" if not failures else "FAIL",
            "" if not failures else "errors on: " + ", ".join(failures[:6]))


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def dump_types(entries):
    print("")
    print("  {:<10} {:<34} {}".format("GROUP", "FEATURE", "API TYPE"))
    for feat, group, depth in entries:
        print("  {:<10} {:<34} {}".format(
            group or "-", ("  " * depth) + feat.Name, feat.GetTypeName2))
    print("")


def main():
    ap = argparse.ArgumentParser(description="RMS compliance checker for SolidWorks")
    ap.add_argument("--dump-types", action="store_true",
                    help="print every feature with its API type name and exit")
    ap.add_argument("--no-suppress", action="store_true",
                    help="skip the detail suppression test (does not modify the model)")
    ap.add_argument("--exceptions", default="rms_exceptions.json",
                    help="JSON file of approved rule waivers: {rule_id: reason}")
    args = ap.parse_args()

    sw = win32com.client.Dispatch("SldWorks.Application")
    model = sw.ActiveDoc
    if model is None:
        sys.exit("No active SolidWorks document.")

    print("\n  Document: {}".format(model.GetTitle))

    entries = walk(model)

    if args.dump_types:
        dump_types(entries)
        return 0

    waivers = {}
    if os.path.exists(args.exceptions):
        with open(args.exceptions) as fh:
            waivers = json.load(fh)

    rep = Report(waivers)

    check_folders(entries, rep)
    check_grouping(entries, rep)
    check_no_solids_in_ref(entries, rep)
    check_shell_last_in_core(entries, rep)
    check_holes_last_in_detail(entries, rep)
    check_modify_order(entries, rep)
    check_quarantine_order(entries, rep)
    check_reference_direction(model, entries, rep)
    check_detail_internal_refs(entries, rep)
    check_descriptions(entries, rep)
    check_sketches(entries, rep)
    check_parameterization(model, rep)

    if args.no_suppress:
        rep.add("detail.individually_suppressible", "SKIP", "--no-suppress")
    else:
        check_detail_suppression(model, entries, rep)

    return rep.emit()


if __name__ == "__main__":
    sys.exit(main())
