# frmScenarioEditor: setup

The **Edit Scenario** and **New Scenario** buttons on `04_SCN_Manager` open
`frmScenarioEditor`. In this form you can change a scenario's name, description,
effective date, *Include in run* and *Include in compare*, or create a new
scenario. The **Current State** row is left out of the form, so it can't be
edited there.

`tblScenarios` is read-only. If you select a cell in the table, a note tells you
to use Edit Scenario. Double-clicking a scenario's row opens the editor on that
scenario. Excel blocks typing, pasting and row deletes in the table.

The form creates all of its controls in code, so you don't have to place
anything in the designer.

## Step 1: Import the modules

Import or update these modules in the VBE:

- `ScenarioTableLock.bas` (new)
- `ScenarioStore.bas`
- `ScenarioUI.bas`
- `ScenarioConfig.bas`

## Step 2: Add the form

1. In the VBE, right-click the VBAProject and choose **Insert > UserForm**.
2. In the Properties window (**F4**), set `(Name)` to `frmScenarioEditor`. You
   don't need to change the size or caption, because the code sets both.
3. Right-click the form and choose **View Code**. Delete anything already there.
4. Open `frmScenarioEditor.frm` in a text editor. Copy everything from
   `Option Explicit` to the end of the file, and paste it into the code window.
5. Run **Debug > Compile VBAProject** and check that it compiles cleanly.

When you next export the modules from the VBE, the real `.frm` and `.frx` replace
the code-only `.frm` in this folder.

## Step 3: Hook up the manager sheet

Double-click `04_SCN_Manager` in the Project Explorer to open its sheet module,
then add the following. If a `Worksheet_BeforeDoubleClick` already exists, add
the line to it instead.

```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If HandleScenarioTableDoubleClick(Target) Then Cancel = True
End Sub
```

## Step 4: Re-apply the lock on open

In `ThisWorkbook`, add `LockScenarioTable` to `Workbook_Open`:

```vba
Private Sub Workbook_Open()
    WriteModelInfo
    LockScenarioTable
End Sub
```

The `UserInterfaceOnly` part of the lock doesn't survive closing the file, so the
lock has to be applied again each time the workbook opens.

## Step 5: Lock the table (once)

Press **Alt+F8**, choose `PrepareScenarioTableLock`, and click **Run**. The macro
unlocks every cell on the sheet, then locks only `tblScenarios` and protects the
sheet. Run it again if you move the table.

## Step 6: Buttons

- **Edit Scenario**: draw a shape on `04_SCN_Manager`, right-click it, choose
  **Assign Macro**, and pick `EditScenarioButton`.
- **New Scenario**: no change needed. It still calls `NewScenarioButton`, which
  now opens the form.
- **Rename**: you can delete this button. If you keep it, `RenameScenarioButton`
  now opens the editor too.

You could also put a short label above the table: *Read-only. Edit via Edit
Scenario.*

## How it behaves

- **Edit mode** opens on the scenario in `scenarioPicker`. If the picker is empty,
  it opens on the loaded scenario instead.
  - Save is enabled once something has been changed.
  - Revert throws your edits away.
  - If you switch scenarios or close the form with unsaved edits, it asks whether
    to save first.
- **Renaming** a scenario also updates `scenarioPicker` and the output picker, if
  either one still shows the old name.
- **Loaded scenario**: if you change its effective date, `activeEffectiveDate` is
  updated as well. Without this, the next Save would copy the old date back.
  Unsaved input edits are not touched.
- **Changing the date of a scenario that has been run**: you get a warning that
  its stored results are out of date and need a re-run.
- **New mode** opens with the fields already filled in: a suggested unique name,
  today's date, and both Include boxes ticked.
  - **Create** builds the scenario from the blank default and loads it.
  - Closing the form creates nothing.
- **Validation**: the name can't be blank, can't be `Current State`, and can't
  match another scenario's name. The effective date has to be a real date; typing
  it as `dd-mmm-yyyy` avoids any day/month mix-up.

## Notes on the lock

- Only the cells in `tblScenarios` are locked. The picker and the other tables on
  the sheet can still be edited as before. On a protected sheet, though, Excel
  won't let you insert or delete whole sheet rows or columns, and typing directly
  below another table on this sheet won't grow that table.
- Macros lift the protection while they write. `BeginBatch` and `EndBatch` handle
  this, so every scenario button keeps working.
  - If a macro you add later fails with error 1004 on this sheet, it is writing
    outside a batch.
  - Wrap it in `BeginBatch`/`EndBatch`, or in
    `UnlockScenarioSheetIfLocked`/`RelockScenarioSheet`.
- Locking replaces any Data Validation that was on `tblScenarios` (for example
  TRUE/FALSE dropdowns) with the read-only input message.
- The password is `SHEET_PASSWORD` in `ScenarioTableLock.bas`. It is blank by
  default, the same as `ModelInfo`. If you need to unprotect the sheet by hand, use
  **Review > Unprotect Sheet**, and run `LockScenarioTable` afterwards.
