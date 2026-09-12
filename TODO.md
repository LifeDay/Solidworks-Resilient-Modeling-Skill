# TODO — skill improvements from 2026-09-12 live session

Source: [solidworks-rms-session-findings.md](solidworks-rms-session-findings.md), a live pywin32 build
session against SolidWorks 2026 (rev 34.1.1). Items ordered by priority.

## 1. Safety: document isolation (highest priority)

`sw.NewPart()` returned the user's real `ActiveDoc` instead of a new part in one
observed call. Trusting the returned handle led to a stray sketch landing in the
user's real document, and a later `CloseDoc` call (using a title captured earlier)
closing that same real, unsaved document.

- [x] Add a "Working safely against a live session" section to [SKILL.md](SKILL.md),
      ahead of any API detail. Cover the incident and the rule below.
- [x] Hard rule: never call `CloseDoc` / `QuitDoc` / `EditDelete` or any other
      destructive/mutating call without re-fetching and re-checking the target
      document's identity (title, or a custom property) in the same breath as the
      call — never rely on a title/handle captured earlier in the script.
- [x] Never run exploratory API probes (e.g. `InsertSketch`) against `sw.ActiveDoc`
      or an unverified "scratch" doc when a user document might be open in the same
      `SldWorks.Application` process.
- [x] Recommend creating scratch docs with an identifying custom property, and
      asserting `doc.GetTitle() not in known_user_titles` before any mutating call.

## 2. api-notes.md: lead with dynamic-dispatch quirks

The "Connecting" section currently undersells how much late-binding behavior
differs from normal API-Help-documented calls. Rewrite it to lead with:

- [x] Zero-argument getter-style methods auto-invoke on attribute access and must
      be called **without parentheses** (`doc.FirstFeature`, `feat.GetNextFeature`,
      `doc.GetEquationMgr`, `sketchSeg.GetType`, `sketchSeg.GetStartPoint2`,
      `doc.GetTitle`, `ext.GetWhatsWrongCount`). Calling with `()` raises
      `TypeError: 'X' object is not callable`.
- [x] Multi-arg methods with an Object/IDispatch parameter reject a bare Python
      `None` (`DISP_E_TYPEMISMATCH`) — wrap it as
      `win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)`. Confirmed fix for
      `SelectByID2`'s `Callout` parameter.
- [x] Correct the current early-binding advice: `gencache.EnsureDispatch` (and
      `CastTo`, which calls it internally) failed with `Element not found` on
      essentially every live SW 2026 object tried. Early binding could not be used
      for actual calls this session — note this as a known-bad path, not a
      first-resort fix for `None`-return issues.
- [x] Document the `EnsureModule`-by-CLSID trick as the replacement for "verify
      against SolidWorks API Help" (that instruction is fine for a human, useless
      for an agent without browser access to a licensed help site):
      `win32com.client.gencache.EnsureModule(clsid, 0, major, minor)` against the
      SolidWorks type library CLSID (found via `win32com.client.selecttlb.EnumTlbs()`)
      generates a readable `.py` in `gen_py` with real method signatures, even
      though the generated module can't itself be used for dispatch.
      - Main typelib CLSID for SW2026: `{83A33D31-27C5-11CE-BFD4-00400513BB57}`,
        major/minor `34/0`.
      - Enum constants (`swEndConditions_e`, `swSelectType_e`,
        `swFeatureSuppressionAction_e`, `swConstrainedStatus_e`,
        `swInConfigurationOpts_e`, `swFeatureTreeFolderType_e`, ...) live in a
        separate typelib, "SOLIDWORKS 2026 Constant type library",
        CLSID `{4687F359-55D0-4CD3-B6CF-2EB42C11F989}`.

## 3. api-notes.md: fix EquationMgr guidance

- [x] `Add3` is documented as the primary call with `Add2`/`Add` framed as a
      fallback for *older* releases. On the SW2026 build tested, `Add3` silently
      failed (returned `-1`, no exception, no equation added) while
      `Add2(index, equation_text, use_automatic_solve_order)` worked correctly.
      Reframe: try `Add2` first (or note the `Add3` failure mode explicitly) rather
      than implying newer = `Add3`-safe.

## 4. api-notes.md: method-location / naming corrections

None of these are in the current notes and each "looks right" while being wrong:

- [x] `InsertSketch2` does not exist via dynamic dispatch on this build's
      `SketchManager` — use `InsertSketch(bool)` instead.
- [x] `InsertAxis2(AutoSize)` lives on `IModelDoc2`, **not** `FeatureManager`,
      despite every other feature-creation call living on `FeatureManager`.
- [x] `AddDimension2(x, y, z)` lives on `IModelDoc2`, **not**
      `ModelDocExtension` — easy to guess wrong since `SelectByID2` and
      `GetWhatsWrongCount` are on `Extension`.
- [x] To select a *feature* (not face/edge/plane/sketch) by name via `SelectByID2`
      for a folder-wrap or pattern seed, the `Type` string must be `"BODYFEATURE"`
      — not `GetTypeName2()`'s value (e.g. `"Extrusion"`), not `"FEAT"`, not `""`.
- [x] `SketchManager.CreateCenterRectangle` returns **6** sketch segments (4 real
      profile edges + 2 construction diagonals), not 4. Don't assume return-order
      maps to bottom/left/top/right — filter on `.ConstructionGeometry` (bool)
      first, then use `.GetStartPoint2` / `.GetEndPoint2` (zero-arg, no parens) to
      classify horizontal vs. vertical.

## 5. Confirmed recipes — promote from description to copy-pasteable code

- [x] Folder-wrap recipe: select the group's features (`Type="SKETCH"` /
      `"BODYFEATURE"` as appropriate), call `FeatureManager.InsertFeatureTreeFolder2(2)`
      (the "Containing" folder type against a prior selection), then rename the
      returned folder feature's `.Name` to the RMS group name (e.g. `"3-Core"`).
      Add this as a concrete example in [SKILL.md](SKILL.md)'s Folders section or
      api-notes.md, replacing the current hint-level description.
- [x] Base-plate recipe (sketch on default plane → boss extrude) — add the full
      confirmed code block from the findings doc verbatim to api-notes.md,
      including the 23-positional-arg `FeatureExtrusion3` call (verified against
      the `gencache.EnsureModule`-generated signature).

## 6. rms_check.py: folder end-tag marker

- [x] Flat tree traversal (`FirstFeature`/`GetNextFeature`) shows a synthetic
      `"<FolderName>___EndTag___"` marker feature immediately after a folder's
      contents. `walk()` in [rms_check.py](rms_check.py) doesn't special-case this
      — currently harmless (the end-tag is itself `FtrFolder`-typed and every rule
      that matters skips `FOLDER_TYPE` features) but add a one-line comment so a
      future maintainer doesn't "fix" it into a bug.

## Not yet actioned / needs a decision

- Confirm whether `Add3`'s failure is 2026-specific or applies more broadly before
  rewriting the guidance as a blanket "use Add2" — session only tested one build.
