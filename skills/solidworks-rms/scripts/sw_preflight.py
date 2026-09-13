#!/usr/bin/env python
"""
sw_preflight.py - check the environment before driving SolidWorks from a script.

Run this first, every time. Each check exists because its absence cost a real
session; the remedies printed on failure are the confirmed ones.

    python scripts/sw_preflight.py             # report only
    python scripts/sw_preflight.py --fix       # also turn off the input-dimension dialog
    python scripts/sw_preflight.py --json      # machine-readable, for a build script

Exits non-zero if anything that would break or endanger a build script is wrong.
The "Open documents" list it prints is what a build script should pass as
known_user_titles to sw_helpers.assert_scratch_doc.
"""

import argparse
import json
import os
import struct
import subprocess
import sys

CHECKS = []


def check(name, critical=True):
    def register(fn):
        CHECKS.append((name, critical, fn))
        return fn
    return register


class Result(object):
    def __init__(self, ok, detail="", remedy=""):
        self.ok = ok
        self.detail = detail
        self.remedy = remedy


def _sldworks_pids():
    """PIDs of running SLDWORKS.exe processes, independent of COM."""
    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq SLDWORKS.exe", "/NH", "/FO", "CSV"],
            stderr=subprocess.DEVNULL, text=True)
    except Exception:
        return []
    pids = []
    for line in out.splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) > 1 and parts[0].upper().startswith("SLDWORKS"):
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------

@check("python.bitness")
def _bitness(state):
    bits = struct.calcsize("P") * 8
    if bits != 64:
        return Result(False, "{}-bit Python".format(bits),
                      "SolidWorks is 64-bit; COM calls from 32-bit Python fail or "
                      "silently misbehave. Install 64-bit Python.")
    return Result(True, "64-bit")


@check("pywin32.installed")
def _pywin32(state):
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except ImportError as exc:
        return Result(False, str(exc), "pip install pywin32")
    return Result(True)


@check("solidworks.running")
def _reachable(state):
    """Attach to a running SolidWorks. Deliberately does NOT start one.

    Dispatch() launches an *invisible* SolidWorks instance when none is running,
    which holds a license seat, has no window to find, and is not what anyone
    asking "is my environment ready" wants to happen. GetActiveObject attaches
    or fails honestly. --launch opts in to starting one.
    """
    import win32com.client
    try:
        _attach(state, win32com.client.GetActiveObject("SldWorks.Application"))
        return Result(True, "attached to running instance")
    except Exception:
        pass

    # GetActiveObject failing does NOT mean no SolidWorks is running. A
    # Dispatch-launched headless instance never registers in the COM running
    # object table, so GetActiveObject raises -2147221021 'Operation
    # unavailable' while the process sits there Responding, holding a license
    # seat, with no window to close. Confirmed on SW2026 SP1.1.
    orphans = _sldworks_pids()
    if orphans and not state.get("launch"):
        return Result(
            False, "process {} exists but is not COM-attachable".format(
                ", ".join(str(p) for p in orphans)),
            "That is almost certainly an orphaned headless instance from an earlier "
            "script. Dispatch() will silently attach to it and your work will happen "
            "in an invisible session. Close it first: attach with Dispatch, confirm "
            "GetDocuments is empty and Visible is False, then call ExitApp().")

    if not state.get("launch"):
        return Result(False, "not running",
                      "Start SolidWorks and open the target document, then re-run. "
                      "Use --launch to start a headless instance instead.")

    try:
        _attach(state, win32com.client.Dispatch("SldWorks.Application"))
        state["sw"].Visible = True
        state["launched"] = True
    except Exception as exc:
        return Result(False, str(exc),
                      "COM registration may be missing - try a version-qualified "
                      "ProgID such as SldWorks.Application.32.")
    return Result(True, "launched a new instance (--launch)")


def _attach(state, app):
    """Store a late-bound handle, remembering whether pywin32 handed back an early-bound one."""
    from win32com.client import dynamic
    state["early_bound"] = type(app).__module__.startswith("win32com.gen_py")
    state["sw"] = dynamic.Dispatch(app._oleobj_)


@check("com.late_binding", critical=False)
def _binding(state):
    """Whether plain Dispatch / GetActiveObject are early-bound on this machine.

    With a gen_py module cached for the SldWorks typelib, both return makepy
    classes: zero-arg getters come back as bound methods, and the version and
    document checks below would print '<bound method ...>' and "'method' object
    is not iterable". This script, sw_helpers.connect() and rms_check.py re-wrap
    with dynamic.Dispatch, so they are unaffected. Hand-written Dispatch calls
    are not. Confirmed on SW2026 SP1.1.
    """
    if state.get("sw") is None:
        return Result(None, "skipped - no connection")
    if not state.get("early_bound"):
        return Result(True, "late-bound")
    import win32com
    return Result(
        False, "gen_py cache present - plain Dispatch is early-bound here (re-wrapped)",
        "Attach through sw_helpers.connect(), or wrap any object from "
        "win32com.client.Dispatch/GetActiveObject with "
        "win32com.client.dynamic.Dispatch(obj._oleobj_). Cache folder: "
        + win32com.__gen_path__)


@check("solidworks.version", critical=False)
def _version(state):
    sw = state.get("sw")
    if sw is None:
        return Result(None, "skipped - no connection")
    try:
        rev = getattr(sw, "RevisionNumber")
    except Exception as exc:
        return Result(False, str(exc), "")
    state["revision"] = str(rev)
    return Result(True, str(rev),
                  "" if str(rev).startswith("34.") else
                  "Findings in this repo are verified against SW2026 (rev 34.x). "
                  "Treat capabilities.yaml entries as unverified on this version.")


@check("solidworks.documents", critical=False)
def _documents(state):
    sw = state.get("sw")
    if sw is None:
        return Result(None, "skipped - no connection")
    try:
        docs = getattr(sw, "GetDocuments") or ()
        titles = [getattr(d, "GetTitle") for d in docs]
    except Exception as exc:
        return Result(False, str(exc), "")
    state["open_titles"] = titles
    if not titles:
        return Result(True, "none open")
    return Result(True, "{}: {}".format(len(titles), ", ".join(titles)))


@check("options.input_dimension_dialog")
def _input_dim(state):
    """The single highest-value check here.

    With this preference on, every scripted dimension call opens a modal Modify
    dialog. The script looks hung while SolidWorks reports Responding, and
    killing the client while the modal is open makes the next Save() crash
    SolidWorks outright (RPC -2147023170, reproduced three times).
    """
    sw = state.get("sw")
    if sw is None:
        return Result(None, "skipped - no connection")

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from sw_helpers import _input_dim_toggle

    toggle = _input_dim_toggle()
    state["input_dim_toggle"] = toggle
    try:
        enabled = sw.GetUserPreferenceToggle(toggle)
    except Exception as exc:
        return Result(False, str(exc), "")

    if not enabled:
        return Result(True, "off")

    if state.get("fix"):
        sw.SetUserPreferenceToggle(toggle, False)
        return Result(True, "was on - now disabled by --fix")

    return Result(
        False, "ON",
        'Turn off Tools > Options > System Options > General > "Input dimension '
        'value", or re-run with --fix, or wrap dimensioning in '
        "sw_helpers.no_input_dim_dialog(sw).")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------

def run(fix=False, launch=False):
    state = {"fix": fix, "launch": launch}
    rows = []
    for name, critical, fn in CHECKS:
        try:
            result = fn(state)
        except Exception as exc:  # a check must never take the run down with it
            result = Result(False, "check raised: {}".format(exc))
        rows.append((name, critical, result))
    return rows, state


def main():
    ap = argparse.ArgumentParser(description="Environment preflight for SolidWorks scripting")
    ap.add_argument("--fix", action="store_true",
                    help="disable the input-dimension dialog if it is on")
    ap.add_argument("--launch", action="store_true",
                    help="start SolidWorks if it is not already running")
    ap.add_argument("--json", action="store_true", help="emit JSON instead of a table")
    args = ap.parse_args()

    rows, state = run(fix=args.fix, launch=args.launch)
    failed = [n for n, critical, r in rows if critical and r.ok is False]

    if args.json:
        print(json.dumps({
            "ok": not failed,
            "revision": state.get("revision"),
            "early_bound_by_default": state.get("early_bound"),
            "open_titles": state.get("open_titles", []),
            "checks": {n: {"ok": r.ok, "detail": r.detail, "remedy": r.remedy}
                       for n, _, r in rows},
        }, indent=2))
        return 1 if failed else 0

    width = max(len(n) for n, _, _ in rows) + 2
    print("")
    for name, critical, result in rows:
        if result.ok is None:
            status = "SKIP"
        elif result.ok:
            status = "OK"
        else:
            status = "FAIL" if critical else "WARN"
        print("  {:<6} {:<{w}} {}".format(status, name, result.detail, w=width).rstrip())
        if result.remedy:
            for line in _wrap(result.remedy):
                print("         {}".format(line))
    print("")

    if failed:
        print("  Not safe to run a build script: {}\n".format(", ".join(failed)))
        return 1

    titles = state.get("open_titles", [])
    if titles:
        print("  Pass these to assert_scratch_doc as known_user_titles:")
        print("    {}\n".format(json.dumps(titles)))
    print("  Ready.\n")
    return 0


def _wrap(text, width=68):
    words, line, out = text.split(), "", []
    for word in words:
        if line and len(line) + 1 + len(word) > width:
            out.append(line)
            line = word
        else:
            line = "{} {}".format(line, word).strip()
    if line:
        out.append(line)
    return out


if __name__ == "__main__":
    sys.exit(main())
