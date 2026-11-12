# inputScenarioPicker: setup

`05_Price_Setting_Inputs` gets its own scenario picker. It always shows the
loaded scenario, and choosing another one loads it straight away. The picker on
`04_SCN_Manager` works as before: it picks what Load, Edit and Delete act on,
and nothing loads until you click Load.

## Step 1: Import the module

Import or update `ScenarioUI.bas` in the VBE.

## Step 2: Add the picker cell

1. Choose a cell on `05_Price_Setting_Inputs` for the picker. You can merge it
   across a few columns if you want.
2. Name it `inputScenarioPicker` in the **Name Box** or **Formulas > Name Manager**.
   If the cell is merged, the name can point at the whole merged area or just its
   top-left cell.
3. Keep it **out of** all of the following:
   - the summary tables written from W10 (`tblMetricsOut*`). A summary rebuild
     would write over the picker.
   - every table in `ScenarioDirtyOutputTables`, plus `tblActiveATL` and
     `tblActiveDiscretions`. Changing the picker would mark the scenario as
     having unsaved changes.
   - `activeEffectiveDate`, for the same reason.
4. You don't need to set up Data Validation. The next refresh points the cell
   at `scenarioPickerList`, the same list the manager's picker uses.

## Step 3: Hook up the sheet

Double-click `05_Price_Setting_Inputs` in the Project Explorer to open its sheet
module.

If it already has a `Worksheet_Change`, add this line as the **first** line
inside it, before any `MarkActiveDirtyIfRelevant` call:

```vba
    If HandleInputScenarioPickerChange(Target) Then Exit Sub
```

If it has no `Worksheet_Change`, add this:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If HandleInputScenarioPickerChange(Target) Then Exit Sub
    MarkActiveDirtyIfRelevant Target
End Sub
```

## Step 4: Fill it in

Press **Alt+F8**, choose `RefreshScenarioManager`, and click **Run**. This sets
up the dropdown and writes the loaded scenario's name into the cell.

Add `RefreshScenarioManager` to `Workbook_Open` in `ThisWorkbook` as well, if
it isn't there already, so the picker is correct as soon as the file opens.

## Step 5 (optional): Status cell

In a cell next to the picker, enter `=scenarioStatus`. It shows "Unsaved
changes" or "Saved …" for the loaded scenario. It only updates when
`RefreshScenarioManager` runs, the same as the cell it copies on the manager
sheet.

## If 05 is sheet-protected

- The picker cell must be unlocked: **Format Cells > Protection**, untick
  **Locked**.
- Unprotect the sheet once before Step 4, because Excel won't let the macro add
  the dropdown to a protected sheet. Protect it again afterwards. Later refreshes
  find the dropdown already in place and leave it alone.

## How it behaves

- **Choosing a scenario on 05** loads it, recalculates, and sets the manager's
  picker to the same scenario.
- **Unsaved changes** bring up a prompt before the switch:
  - **Yes** saves first. If the inputs fail validation, you see why and nothing
    loads.
  - **No** throws the changes away and loads the new scenario.
  - **Cancel** keeps you where you are.
- **Loading on 04** with the Load button, or New, Save As, Revert, Run, or
  renaming in the editor, updates the 05 picker.
- **Changing 04's picker without clicking Load** leaves 05 alone.
- **Clearing the cell, pasting over it, or entering a name that isn't a saved
  scenario** puts the loaded scenario's name back.
- **Deleting the loaded scenario** leaves the 05 picker blank.
