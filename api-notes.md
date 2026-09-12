# SolidWorks API notes

The calls this skill depends on, with the caveat that signatures and enum values change between releases. **Verify against the local SolidWorks API Help before relying on any of these.** Where a call takes more than a couple of arguments, check the argument order rather than assuming.

## Connecting (pywin32)

```python
import win32com.client

sw = win32com.client.Dispatch("SldWorks.Application")
sw.Visible = True
model = sw.ActiveDoc
```

`Dispatch` gives late binding, which is fine for most calls. Methods with `out` parameters — anything returning status through an argument — need early binding via `win32com.client.gencache.EnsureDispatch("SldWorks.Application")`. If a call returns `None` unexpectedly, early binding is the first thing to try.

All lengths in the API are **metres**, regardless of document units.

## Folders

- `FeatureManager.InsertFeatureTreeFolder2(type)` — creates a folder. The type enum (`swFeatureTreeFolderType_e`) distinguishes an empty folder from one containing the current selection.
- `FeatureManager.MoveToFolder(folderName, featureName, moveOnlySelected)` — moves a feature into a folder.
- Folders appear in tree traversal with `GetTypeName2()` returning `"FtrFolder"`.

Because folders require contiguous features and SolidWorks will not reorder past a dependency, create the folder when the group starts and add features as they are built. Attempting to sort a finished tree into folders will fail on any part of realistic complexity.

## Global variables and equations

`EquationMgr` hangs off `ModelDoc2`:

```python
eq = model.GetEquationMgr()
eq.Add3(-1, '"plate_width" = 120', True, 0)          # global variable
eq.Add3(-1, '"D1@Sketch1" = "plate_width"', True, 0)  # driven dimension
```

Index `-1` appends. Older releases expose `Add2` or `Add` with fewer arguments. `eq.GetCount()`, `eq.Equation(i)`, and `eq.Status(i)` are useful for verifying what actually landed.

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

## Feature type strings

`GetTypeName2()` returns strings like `Extrusion`, `Cut`, `Revolution`, `Sweep`, `Loft`, `Fillet`, `Chamfer`, `Shell`, `Draft`, `LPattern`, `CirPattern`, `MirrorPattern`, `HoleWzd`, `ProfileFeature` (a sketch), `RefPlane`, `RefAxis`, `CoordSys`, `FtrFolder`.

These vary by version and by how the feature was created. Run `rms_check.py --dump-types` against known-good parts and adjust the classification sets at the top of the checker rather than trusting this list.
