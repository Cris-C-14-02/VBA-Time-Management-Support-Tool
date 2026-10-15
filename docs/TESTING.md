# Verification

## Native Excel regression tests

Verified on **October 15, 2026**, using Windows desktop Excel **16.0**. The tests run in a disposable workbook copy. They load real UserForms, exercise the service modules and assert actual worksheet values and chart counts.

| Check | Result verified |
| --- | --- |
| Calendar and duration validation | Valid leap day accepted; impossible dates and nonpositive estimates rejected |
| Forms | Task and sleep forms load; six categories and seven weekdays; school-field visibility changes |
| School fields | A deadline before the start date is rejected without appending a master row |
| Derived views | Master tasks appear in the correct category, sorted newest first |
| Date boundaries | Today is separate from tomorrow through Date + 7; Date + 8 excluded |
| Deletion | Removing a category updates master, category view, daily list and chart |
| Repeatability | Repeated refresh leaves one category chart and clears old list rows |
| Sleep | 0 and 24 accepted; >24 rejected; recorded values 0 and 8 average to 4 across two days |
| Workbook isolation | A second active workbook remains empty |

## Run locally

The optional native runner requires Windows, installed desktop Excel and programmatic VBA project access already permitted by the local Office policy. It does not change Office security settings. If that access is unavailable, automated native tests cannot run in that environment.

From the repository root:

```powershell
./scripts/Test-Workbook.ps1
```

To create a separate workbook with the synthetic screenshot fixture after the tests pass:

```powershell
./scripts/Test-Workbook.ps1 -DemoOutput ./demo.xlsm
```

Choose a new output path. The runner refuses to overwrite an existing demo and never writes test records to the input workbook. The test module is removed before saving the demonstration copy.

## Source synchronization

```powershell
python -m pip install -r scripts/requirements.txt
python scripts/extract_vba.py "Group 21 Time Managemnt Tool.xlsm" --check
```

This compares every extracted VBA component and the workbook hash against the committed snapshot. It does not establish runtime correctness; the native tests above exercise runtime behavior.

## Visual and preservation checks

The original and revised primary views were rendered with Excel for layout inspection. README screenshots were exported by Excel from separate synthetic fixtures. Source input values, worksheet names and order, and existing form designer resources were compared against the original workbook. Only the expected code and presentation changes were made.

## Remaining manual checks

The automated suite loads the forms but does not click every original shape, exercise every invalid submission through a visible dialog, or test every Excel version. Before distributing a release, click the add, submit, navigation and clear controls in a fresh working copy. Confirm that invalid submissions stay open with an explanatory message, successful submissions refresh or invalidate the appropriate output, and long entered text remains readable. External resource links, if present, were not checked.
