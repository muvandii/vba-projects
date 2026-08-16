# Automate Excel VBA Course — Module-by-Module Project Workbook

**Companion to:** `vba-course-projects-report.md`  
**Audience:** VBA learners building a client-ready portfolio  
**Format:** 10 chapter projects, broken into build modules, styling passes, test gates, and freelance delivery tasks  
**Recommended pace:** One chapter per week; 3–5 focused sessions per chapter

---

## How to Use This Workbook

For every chapter:

1. Create the project folder and macro-enabled workbook before writing code.
2. Complete modules in order. Do not move past a **checkpoint** until the listed test passes.
3. Keep `Option Explicit` at the top of every standard, worksheet, and workbook module.
4. Export finished VBA modules (`.bas`, `.cls`, or `.frm`) so code can be reviewed in Git.
5. Complete the styling pass only after the core workflow works.
6. Run the acceptance tests with a clean copy of the sample data.
7. Finish the freelance simulation: scope, estimate, client update, handoff, and portfolio evidence.

### Standard Portfolio Folder

```text
chapter-##_project-name/
├── workbook/
│   └── ProjectName.xlsm
├── sample-data/
│   └── sample-input.csv
├── src/
│   ├── modMain.bas
│   └── other-exported-modules.bas
├── screenshots/
│   ├── 01-input.png
│   ├── 02-running.png
│   └── 03-output.png
├── tests/
│   └── acceptance-tests.md
└── README.md
```

### Definition of Done for Every Project

- [ ] Workbook opens without broken links or repair warnings.
- [ ] Code compiles with **Debug → Compile VBAProject**.
- [ ] A non-technical user can find and run the main action.
- [ ] Raw input is preserved or recoverable.
- [ ] Valid input produces the documented output.
- [ ] Empty, malformed, boundary, and cancellation cases are tested.
- [ ] Any changed Excel application settings are restored after success or failure.
- [ ] Workbook contains an Instructions sheet and version number.
- [ ] README includes setup, macro-security note, usage, assumptions, and limitations.
- [ ] Source modules and three portfolio screenshots are included.

---

# Chapter 1 — VBA Basics

## Project: Personal Expense Tracker Macro

**Client outcome:** A one-click workflow posts an expense from an entry form to a formatted transaction log.  
**Primary concepts:** Subs, worksheets, ranges, cells, `Offset`, formatting, form-control buttons.  
**Target build time:** 1.5–2.5 hours.

## Workbook Blueprint

| Sheet | Purpose | Required elements |
|---|---|---|
| `Instructions` | User onboarding | Purpose, three-step workflow, macro notice, version |
| `Entry` | Simple input form | Date, Category, Description, Amount, Log and Clear buttons |
| `Log` | Append-only transaction list | Four headers, frozen row 1, currency/date formats |

Use these named cells on `Entry`: `inpDate`, `inpCategory`, `inpDescription`, and `inpAmount`. Named ranges make the final workbook easier to maintain than hard-coded addresses.

## Module 1.1 — Create the Shell

**Build tasks**

1. Save the workbook as `ExpenseTracker.xlsm`.
2. Rename sheets and add the headers shown above.
3. Put labels in `Entry!A3:A6` and input cells in `B3:B6`.
4. Put `Date`, `Category`, `Description`, and `Amount` in `Log!A1:D1`.
5. Add `Option Explicit` to `modExpenseTracker`.

**Checkpoint**

- [ ] All sheet names exactly match the specification.
- [ ] Named input cells resolve through the Name Box.
- [ ] The workbook closes and reopens as `.xlsm`.

## Module 1.2 — Post an Expense

Create `Public Sub LogExpense()`:

1. Set worksheet variables for `Entry` and `Log`.
2. Find the next row under the existing log.
3. Transfer values directly; do not use `.Select`, `.Activate`, copy, or paste.
4. Format date as `mmm d, yyyy` and amount as a currency.
5. Clear the four input cells and return focus to the date field.
6. Display a short success message.

```vba
Public Sub LogExpense()
    Dim wsEntry As Worksheet, wsLog As Worksheet
    Dim nextRow As Long

    Set wsEntry = ThisWorkbook.Worksheets("Entry")
    Set wsLog = ThisWorkbook.Worksheets("Log")
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    wsLog.Cells(nextRow, "A").Value = wsEntry.Range("inpDate").Value
    wsLog.Cells(nextRow, "B").Value = wsEntry.Range("inpCategory").Value
    wsLog.Cells(nextRow, "C").Value = wsEntry.Range("inpDescription").Value
    wsLog.Cells(nextRow, "D").Value = wsEntry.Range("inpAmount").Value
    wsLog.Cells(nextRow, "A").NumberFormat = "mmm d, yyyy"
    wsLog.Cells(nextRow, "D").NumberFormat = "$#,##0.00"

    wsEntry.Range("inpDate,inpCategory,inpDescription,inpAmount").ClearContents
    MsgBox "Expense logged.", vbInformation, "Expense Tracker"
End Sub
```

**Checkpoint:** Two consecutive runs create two consecutive rows and never overwrite row 1.

## Module 1.3 — Clear and Navigate

- Build `Public Sub ClearEntry()`.
- Add **Log Expense** and **Clear Form** buttons.
- Assign each macro and test after reopening the workbook.
- Add a hyperlink from `Log` back to `Entry`.

## Styling Pass

- Theme: navy `#17365D`, teal `#2F75B5`, pale yellow inputs `#FFF2CC`.
- Use one font family (Aptos or Calibri), 11 pt body and 18 pt title.
- Hide gridlines only on `Instructions` and `Entry`.
- Give input cells a thin border and unlocked appearance.
- Turn `Log!A1:D1` into a visually strong header; use an Excel Table only as an optional upgrade.
- Add a small status note: “Yellow cells are editable.”

## Test Gate

| Test | Action | Expected result |
|---|---|---|
| C1-T1 | Post a normal expense | Values appear in the next log row |
| C1-T2 | Post 0.5 as amount | Displays `$0.50` |
| C1-T3 | Enter a long description | Value remains intact; column wraps |
| C1-T4 | Click Clear | Inputs clear; no log row is added |
| C1-T5 | Run three times | Exactly three new rows are appended |

## Freelance Simulation

**Brief:** “Create a simple expense logger for a solo consultant who records 10–20 purchases daily.”  
**Scope statement:** One form, one append-only log, two buttons, and basic formatting; no dashboard or receipt storage.  
**Estimate:** 2 hours build + 30 minutes testing/documentation.  
**Client update to write:** A 100-word message explaining what works, what remains, and one assumption.  
**Portfolio evidence:** Empty form, completed form, populated log, and a 10-second click-to-result GIF.

**Upsell exercise:** Quote separately for category validation, monthly totals, and PDF export.

---

# Chapter 2 — Variables

## Project: Sales Commission Calculator

**Client outcome:** Tiered commissions are calculated accurately for a sales export.  
**Primary concepts:** Explicit data types, object variables, conversion, Boolean flags, scope.  
**Target build time:** 2–3 hours.

## Workbook Blueprint

| Sheet | Purpose |
|---|---|
| `Instructions` | Setup and tier explanation |
| `SalesData` | Salesperson, Region, SalesAmount, SaleDate, Commission, Status |
| `Summary` | Total sales, total commission, high-earner count |

## Module 2.1 — Typed Data Pipeline

Create `modCommission` and declare every variable intentionally:

```vba
Option Explicit

Public Sub CalculateCommissions()
    Dim wsData As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim salesAmount As Double, rate As Double, commission As Double
    Dim totalCommission As Double
    Dim salesperson As String
    Dim saleDate As Date
    Dim isHighEarner As Boolean

    Set wsData = ThisWorkbook.Worksheets("SalesData")
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    ' Processing loop is added in Module 2.2.
End Sub
```

**Build tasks**

- Explain in comments why row counters use `Long`, money uses `Double`, and worksheet references use object variables.
- Add 15 sample records spanning all commission tiers.
- Use a real Excel date value, not a date-looking string.
- Compile and deliberately misspell one variable to observe `Option Explicit` catching it; then fix it.

## Module 2.2 — Convert, Calculate, Write

For each populated row:

1. Read the salesperson with `CStr`.
2. Guard the amount with `IsNumeric`, then use `CDbl`.
3. Guard the date with `IsDate`, then use `CDate`.
4. Determine rate: below 5,000 = 5%; 5,000 through below 10,000 = 7%; 10,000+ = 10%.
5. Calculate commission and set `isHighEarner = (commission > 1000)`.
6. Write commission and status in one pass.

Extract the tier decision into a typed function:

```vba
Private Function CommissionRate(ByVal amount As Double) As Double
    Select Case amount
        Case Is < 5000: CommissionRate = 0.05
        Case Is < 10000: CommissionRate = 0.07
        Case Else: CommissionRate = 0.1
    End Select
End Function
```

**Checkpoint:** Values at 4,999.99, 5,000, 9,999.99, and 10,000 receive the documented rates.

## Module 2.3 — Summary and Auditability

- Accumulate total sales and commission in `Double` variables.
- Count high earners with a `Long`.
- Write results to `Summary`, including a “Last refreshed” timestamp.
- Add `Public Const APP_VERSION As String = "1.0.0"` and show it on `Instructions`.
- Add a `ClearResults` macro that clears only generated columns and summary values.

## Styling Pass

- Theme: dark green `#375623`, light green `#E2F0D9`, gold accent `#FFD966`.
- Convert rates to percentages and monetary output to the selected client currency.
- Apply a green fill to “High Earner”; do not rely on the star symbol alone.
- Use right alignment for amounts and centered dates/status.
- Keep raw inputs visually distinct from macro outputs.

## Test Gate

| Test | Input | Expected |
|---|---:|---|
| C2-T1 | 4,999.99 | 5% |
| C2-T2 | 5,000.00 | 7% |
| C2-T3 | 10,000.00 | 10% |
| C2-T4 | Blank amount | Skipped or flagged; no type mismatch |
| C2-T5 | Commission > 1,000 | High Earner status |
| C2-T6 | Rerun macro | Results refresh; totals are not doubled |

## Freelance Simulation

**Brief:** A manager sends a weekly CSV and fixed three-tier policy. Build a reusable calculator.  
**Discovery questions:** Is commission progressive or one rate on the whole amount? Which currency? How are returns handled? Are thresholds inclusive?  
**Change request:** The client asks to edit tiers without VBA. Respond with impact, price, and delivery change.  
**Portfolio claim to prove:** “Eliminates boundary and type errors.” Include a screenshot of the four boundary tests.

---

# Chapter 3 — Conditional Logic

## Project: Invoice Aging and Dunning System

**Client outcome:** Every invoice receives an aging bucket, action, urgency flag, and optional offer.  
**Primary concepts:** `If`, `ElseIf`, `Select Case`, `And`, `Or`, `Not`, guards.  
**Target build time:** 2–3 hours.

## Workbook Blueprint

- `Instructions` — as-of date, bucket definitions, warning that the tool does not send email.
- `Invoices` — InvoiceID, Customer, InvoiceDate, Amount, DueDate, DaysOverdue, Bucket, Action, Urgent, Offer.
- `AgingSummary` — count and value by bucket.
- Named cell `asOfDate` — makes test results repeatable; do not hard-code `Date` during development.

## Module 3.1 — Validate Before Branching

Create `modAging` with a controller `BuildAgingReport()`.

For each row:

- Invoice ID and customer must be nonblank.
- Amount must be numeric and nonnegative.
- Due date must pass `IsDate`.
- Invalid rows receive `Review input` and skip all date arithmetic.
- Valid rows calculate `daysOverdue = CLng(asOfDate - dueDate)`.

**Checkpoint:** A blank due date never becomes an unexpected 1899 date and never triggers an error.

## Module 3.2 — Assign Buckets

Use one auditable function:

```vba
Private Function AgingBucket(ByVal daysOverdue As Long) As String
    Select Case daysOverdue
        Case Is <= 0: AgingBucket = "Current"
        Case 1 To 30: AgingBucket = "1-30"
        Case 31 To 60: AgingBucket = "31-60"
        Case 61 To 90: AgingBucket = "61-90"
        Case Else: AgingBucket = "90+"
    End Select
End Function
```

Use a second `Select Case` to map bucket to action: no action, reminder email, follow-up call, collections review, or legal review.

## Module 3.3 — Compound Rules

- Urgent = overdue more than 60 days **and** amount over 10,000.
- Early settlement offer = 1–30 days overdue **and** amount over 5,000.
- Manual review = customer blank **or** amount/due date invalid.
- Never apply an offer to an urgent account.
- Summarize both count and value for each bucket.

Write each business rule in plain English above its corresponding condition.

## Styling Pass

- Current: pale green; 1–30: pale blue; 31–60: pale yellow; 61–90: orange; 90+: pale red.
- Include a legend and as-of date above the summary.
- Use icons only as secondary cues; text labels remain mandatory.
- Freeze identifiers and headers; use filters.
- Avoid fully saturated red rows, which reduce readability.

## Test Gate

- [ ] Days overdue of -1, 0, 1, 30, 31, 60, 61, 90, and 91 map correctly.
- [ ] Amount exactly 10,000 follows the stated urgent rule (`>`, not `>=`).
- [ ] Missing date is flagged and excluded from totals.
- [ ] A 90+ invoice gets legal review and no discount offer.
- [ ] Summary count equals the number of valid detail rows.
- [ ] Summary value reconciles to the valid invoice total.

## Freelance Simulation

**Brief:** Process an accounts-receivable export but do not contact customers automatically.  
**Risk note:** Draft a client-facing disclaimer that dunning actions are recommendations and legal escalation requires approval.  
**Demo script:** Explain one invoice at each boundary in under two minutes.  
**Change-control task:** Client changes 90+ to 120+. Identify the function, tests, documentation, and estimate affected.

---

# Chapter 4 — Loops

## Project: Multi-Sheet and Multi-File Consolidation Engine

**Client outcome:** Repeated store reports become one traceable master table.  
**Primary concepts:** `For`, `For Each`, `Do While`, nested loops, `Exit For`, counters.  
**Target build time:** 2.5–4 hours.

## Workbook Blueprint

| Sheet | Purpose |
|---|---|
| `Instructions` | Required source layout and refresh steps |
| `Control` | Folder path, run button, last run, files/rows processed |
| `Summary` | Consolidated output plus source metadata |
| `Exceptions` | Skipped file/sheet and reason |

Required source columns: Date, Product, Quantity, Revenue. Output adds SourceWorkbook, SourceSheet, and ImportedAt.

## Module 4.1 — Consolidate Internal Sheets

1. Loop `For Each wsSource In ThisWorkbook.Worksheets`.
2. Skip control/output/instruction sheets by name.
3. Determine the last source row.
4. Loop from row 2 through the last row.
5. Copy values to the next output row and append source metadata.
6. Maintain sheet and row counters.

**Checkpoint:** Four 10-row sheets produce exactly 40 detail rows, one header, and no blank separators.

## Module 4.2 — Consolidate a Folder

Build a separate `ImportFolder()` routine:

- Validate that the folder exists.
- Normalize the trailing path separator.
- Use `Dir(folderPath & "*.xlsx")`, then `Do While Len(fileName) > 0`.
- Never import the destination workbook itself.
- Open each source read-only and close it without saving.
- Record files with missing sheets or headers in `Exceptions`.
- Advance `Dir()` exactly once per loop iteration.

```vba
fileName = Dir(folderPath & "*.xlsx")
Do While Len(fileName) > 0
    ' Open, validate, import, and close source.
    fileName = Dir()
Loop
```

## Module 4.3 — Control and Recovery

- Add a `ClearSummary` routine with confirmation.
- Add an `Exit For` when a `STOP` sentinel is encountered in source column A.
- Guarantee an opened source workbook is closed if an error occurs.
- Show final files processed, rows imported, exceptions, and elapsed time.
- Decide and document refresh behavior: replace all output or append only. Default to replace all to prevent duplicates.

## Styling Pass

- Theme: charcoal `#404040`, blue `#5B9BD5`, light gray `#F2F2F2`.
- Make `Control` a dashboard with prominent **Select Folder**, **Run**, and **Clear** actions.
- Format `Summary` as a filterable table with metadata columns in muted gray.
- Use red only on exception count and failed rows.
- Add a visible run stamp and source folder.

## Test Gate

| Scenario | Expected |
|---|---|
| Four valid files | All rows imported once |
| Empty workbook/sheet | Logged and skipped |
| Wrong header | Logged with missing header name |
| Destination file inside source folder | Not self-imported |
| Source workbook already open | Handled without data loss |
| Error during file 2 | Open files close; clear error message appears |
| Second full refresh | No duplicate rows |

## Freelance Simulation

**Brief:** Twelve branch files follow a template, but occasional malformed files must not stop the batch.  
**Pricing exercise:** Quote internal-sheets, folder-import, and column-mapping tiers.  
**Status update:** Write a message reporting processed files, exceptions, and one decision needed.  
**Proof:** Record row reconciliation: sum of valid source rows equals destination rows.

---

# Chapter 5 — Advanced Cell Referencing

## Project: Dynamic Data Cleaner and Normalizer

**Client outcome:** Messy exports become clean, consistently typed, analysis-ready tables.  
**Primary concepts:** last row/column, `CurrentRegion`, `Offset`, `Resize`, `SpecialCells`, bulk formulas, copy destination.  
**Target build time:** 3–4 hours.

## Workbook Blueprint

- `Instructions` — accepted input and non-destructive promise.
- `RawData` — pasted client export; never modify it in the default workflow.
- `CleanData` — generated output.
- `CleaningLog` — operation, affected rows, warning, timestamp.
- `Config` — header row, required columns, date and amount columns.

## Module 5.1 — Detect the Data Boundary

- Reject an empty `RawData` sheet.
- Determine `lastRow` and `lastCol` dynamically.
- Compare `UsedRange` and a last-cell search; document why formatting can inflate `UsedRange`.
- Set a range with `Cells(firstRow, firstCol)` and `Cells(lastRow, lastCol)`.
- Copy values to `CleanData` before cleaning.

**Checkpoint:** Add data beyond the original size; rerun and verify the new boundary is included without code changes.

## Module 5.2 — Clean Safely

Implement operations in this order:

1. Copy raw values to output.
2. Unmerge output cells only.
3. Remove rows that are entirely blank—not rows containing one blank field.
4. Trim text values while preserving formulas if formulas are permitted.
5. Validate and normalize dates.
6. Validate and format numeric columns.
7. Standardize headers and detect duplicates.
8. Recalculate final boundaries.

Do not use `SpecialCells(xlCellTypeBlanks).EntireRow.Delete` without checking whether the **whole row** is blank; that pattern can delete legitimate partial records.

## Module 5.3 — Dynamic Formulas and Final Table

- Use `Offset(1).Resize(dataRowCount)` to target rows below headers.
- Add a quality flag with `.FormulaR1C1`.
- Convert formulas to values only if the client requests a static output.
- Format the resulting `CurrentRegion` and autofit with sensible maximum widths.
- Log counts: source rows, clean rows, blank rows removed, invalid dates, invalid amounts.

## Styling Pass

- Raw sheet tab: gray; clean sheet: green; warning log: orange.
- Add “RAW — DO NOT EDIT” and “GENERATED OUTPUT” banners.
- Use pale red cells for invalid fields and a QualityStatus text column.
- Keep alternating rows subtle and headers high contrast.
- Add before/after metric cards on `Instructions` or `CleaningLog`.

## Test Gate

- [ ] Merged title does not cause column loss.
- [ ] Fully blank rows are removed.
- [ ] Partially blank valid records remain.
- [ ] Leading/trailing spaces are trimmed.
- [ ] Text dates become real dates when valid.
- [ ] Invalid dates are flagged, not silently coerced.
- [ ] RawData is byte-for-byte unchanged at the cell-value level.
- [ ] Repeated run replaces CleanData cleanly.

## Freelance Simulation

**Brief:** The client provides one “representative” file; production exports may differ.  
**Discovery task:** Ask for required columns, locale/date format, formula policy, duplicate rule, and output format.  
**Assumption log:** Write five assumptions and assign each a validation owner.  
**Portfolio evidence:** Side-by-side messy/clean views plus a cleaning-log reconciliation.

---

# Chapter 6 — Message and Input Boxes

## Project: Interactive Shipment Entry and Validation Dashboard

**Client outcome:** Staff enter shipment records through a guided, cancellable workflow.  
**Primary concepts:** `MsgBox`, `Application.InputBox`, return values, validation loops.  
**Target build time:** 2–3 hours.

## Workbook Blueprint

| Sheet | Purpose |
|---|---|
| `Instructions` | Workflow and field rules |
| `Dashboard` | Add Shipment, Edit Last, Update Status buttons |
| `Shipments` | ID, Date, Client, Origin, Destination, WeightKg, Status, CreatedAt |
| `Lists` | Allowed statuses and optional locations |

## Module 6.1 — Capture One Valid Field at a Time

Create small functions rather than one giant procedure:

- `PromptRequiredText(prompt, title, wasCancelled) As String`
- `PromptPositiveNumber(prompt, title, wasCancelled) As Double`
- `PromptValidDate(prompt, title, wasCancelled) As Date`
- `NextShipmentId() As String`

Use `Application.InputBox` when a type argument is needed. Remember that Cancel returns Boolean `False`, which must be distinguished from numeric zero or text.

```vba
Dim response As Variant
response = Application.InputBox("Enter weight in kg:", "Shipment Weight", Type:=1)
If VarType(response) = vbBoolean And response = False Then Exit Sub
If CDbl(response) <= 0 Then
    MsgBox "Weight must be greater than zero.", vbExclamation
End If
```

## Module 6.2 — Build the Wizard

1. Prompt for client, origin, destination, date, and weight.
2. Allow cancellation at every step with no partial row written.
3. Build a readable confirmation summary.
4. Use `vbOKCancel + vbQuestion` before writing.
5. Generate the ID only when saving.
6. Write all fields to the next row in one commit-like block.
7. Confirm the new shipment ID.

**Checkpoint:** Cancel at each prompt; the shipment row count must remain unchanged.

## Module 6.3 — Edit and Update

- Prefill prompts with the last record’s current values.
- Confirm before overwriting.
- Provide status choices with clearly documented button meanings—or use a validated text/number prompt if three buttons are ambiguous.
- Handle an empty data table gracefully.
- Add a maximum retry count or a clear cancellation path to every validation loop.

## Styling Pass

- Theme: deep purple `#7030A0`, lavender `#E4DFEC`, blue action accent.
- Keep the Dashboard sparse: three large buttons, last shipment card, concise help.
- Place the data sheet behind the workflow, but do not hide it until testing is complete.
- Use consistent dialog titles containing the workbook/project name.
- Write prompts in plain language and include units/examples.

## Test Gate

| Test | Expected |
|---|---|
| Valid entry | One complete row, unique ID |
| Cancel at first/middle/final step | No row and no partial values |
| Blank client | Warning and retry |
| Text entered for weight | Excel type validation; no crash |
| Zero/negative weight | Rejected |
| Invalid date string | Rejected with example |
| Edit last when no records | Friendly message |

## Freelance Simulation

**Usability test:** Ask someone unfamiliar with the workbook to add a shipment without coaching; note every hesitation.  
**Client trade-off:** Explain when an InputBox wizard should be upgraded to a UserForm.  
**Handoff task:** Write a one-page operator guide with screenshots and a “How to cancel safely” note.  
**Upsell quote:** UserForm, dropdown lookup, auto-generated PDF label.

---

# Chapter 7 — Events

## Project: Worksheet Change Tracker and Audit Log

**Client outcome:** Monitored changes are logged automatically with user, timestamp, location, old value, and new value.  
**Primary concepts:** worksheet/workbook events, `Target`, `Intersect`, `EnableEvents`, event-safe cleanup.  
**Target build time:** 3–5 hours.

## Workbook Blueprint

- `Instructions` — monitored range, limitations, privacy note.
- `Budget` — Period, Category, Budgeted, Actual, Variance.
- `AuditLog` — Timestamp, User, Sheet, Address, OldValue, NewValue, Note.
- `Config` — monitored address and audit settings.

## Module 7.1 — Initialize on Open

In `ThisWorkbook`:

- Create or validate `AuditLog` during `Workbook_Open`.
- Add headers only if missing.
- Set it to `xlSheetVeryHidden` after setup.
- Store workbook version and open timestamp.
- Do not silently replace an existing audit sheet.

## Module 7.2 — Capture and Log Changes

In the `Budget` sheet module:

1. Use `Worksheet_SelectionChange` to cache the selected cell’s value.
2. In `Worksheet_Change`, exit if the changed range does not intersect the monitored range.
3. Decide how to handle multi-cell paste. The basic version logs one summary record; the premium version captures each cell accurately.
4. Disable events before writing to the log.
5. Use a single cleanup label that always re-enables events.
6. Record `Now`, `Environ$("Username")`, sheet, address, old value, and new value.

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanFail
    If Intersect(Target, Me.Range("B2:G50")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    LogBudgetChange Target, mPreviousValue

CleanExit:
    Application.EnableEvents = True
    Exit Sub
CleanFail:
    MsgBox "The change was made, but audit logging failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
```

**Important limitation:** Selection-based old-value caching is reliable for ordinary single-cell edits but not every paste, fill, external-link update, or macro-driven multi-cell change. Document this instead of claiming a forensic-grade audit trail.

## Module 7.3 — Save Notes and Protection

- Prompt for an optional session note in `Workbook_BeforeSave`.
- Prevent recursion and avoid canceling save unexpectedly.
- Add a controlled admin macro to reveal/hide the log.
- Protect the sheet structure only after testing recovery steps.
- Add an emergency `EnableEventsNow()` utility in a standard module for development.

## Styling Pass

- Budget input cells: pale yellow; formulas: pale gray and locked.
- Variance: green favorable, red unfavorable, with text/number support.
- AuditLog: compact table, filters, monospace for addresses if available.
- Instructions must state what is monitored and who can view the log.
- Avoid presenting `Environ$("Username")` as authenticated identity.

## Test Gate

- [ ] Single-cell monitored edit logs one record.
- [ ] Unmonitored edit creates no record.
- [ ] Old and new values are correct for a normal edit.
- [ ] Error in logger does not leave events disabled.
- [ ] Multi-cell paste follows documented behavior.
- [ ] Save note does not cause a save loop.
- [ ] Reopening preserves and re-hides the log.
- [ ] Admin recovery procedure works.

## Freelance Simulation

**Risk assessment:** List limitations of VBA audit logs versus SharePoint/version history/database controls.  
**Privacy message:** Draft notice covering username and timestamp collection.  
**Scope defense:** Decline a claim that the workbook is tamper-proof; propose appropriate alternatives.  
**Demo:** Edit monitored/unmonitored cells, trigger a controlled logging failure, and show event recovery.

---

# Chapter 8 — Application Settings

## Project: High-Performance Invoice Batch Processor

**Client outcome:** Thousands of invoices are calculated quickly without leaving Excel in a broken state.  
**Primary concepts:** screen updating, calculation, events, alerts, status bar, timing, cleanup.  
**Target build time:** 2–4 hours.

## Workbook Blueprint

| Sheet | Purpose |
|---|---|
| `Instructions` | Data schema and run steps |
| `Control` | Run/clear buttons and benchmark cards |
| `Invoices` | InvoiceID, Client, Date, Amount, TaxRate, Tax, Discount, Total |
| `RunLog` | Start, finish, rows, seconds, result/error |

## Module 8.1 — Save Application State

Do not restore settings to assumed defaults. Capture the user’s current values first:

```vba
Private Type TAppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type
```

Create `CaptureAppState`, `OptimizeApplication`, and `RestoreAppState`. Keep these responsibilities separate and testable.

## Module 8.2 — Batch Process

- Validate required headers before changing settings.
- Record start time with `Timer`.
- Loop through invoice rows and validate amount/tax rate.
- Calculate tax, discount, and total.
- Update the status bar every 100 rows, not every row.
- Write run statistics to `RunLog`.
- Use one success/error cleanup path.

For a batch that can cross midnight, account for `Timer` resetting at midnight or use `Now` for elapsed duration.

## Module 8.3 — Benchmark and Fail Safely

- Run a baseline version with normal settings on a copy of the data.
- Run the optimized version on equivalent data.
- Compare output—not only elapsed time.
- Inject a controlled error at row 1,000.
- Verify calculation mode, events, alerts, screen updating, and status bar are restored.
- Allow Esc interruption with a documented cancellation policy if enabled.

## Styling Pass

- Theme: black/charcoal with cyan performance accent and green success state.
- Control cards: rows processed, elapsed seconds, rows/second, last status.
- Keep progress text factual; do not promise a fixed completion time.
- RunLog uses green Success and red Failed statuses.
- Make generated columns visually distinct from imported columns.

## Test Gate

| Test | Pass condition |
|---|---|
| 5,000 valid rows | Correct output and run log |
| Invalid tax rate | Row flagged or controlled stop per policy |
| Injected error | All application states equal captured values |
| Manual calculation before run | Remains manual after run |
| Status bar had custom value | Restored afterward |
| Optimized vs baseline | Outputs match exactly |
| Rerun | No stale totals or duplicate run output |

## Freelance Simulation

**Performance report:** Record machine context, row count, baseline, optimized time, and speedup; avoid universal “10x” claims.  
**Client update:** Explain why faster code must still be tested for identical output.  
**Support plan:** Write recovery instructions for calculation mode/events and include an admin reset macro.  
**Upsell:** Array processing from Chapter 10, progress UserForm, scheduled folder batches.

---

# Chapter 9 — Advanced Procedures

## Project: Modular Financial Reporting System

**Client outcome:** A maintainable reporting pipeline imports a trial balance, calculates metrics, and produces a period report.  
**Primary concepts:** Functions/Subs, scope, `ByVal`/`ByRef`, optional parameters, `ParamArray`, orchestration.  
**Target build time:** 4–6 hours.

## Workbook Blueprint

- `Instructions` — process, version, accounting assumptions.
- `ReportConfig` — period, tax rate, source path, output currency.
- `TrialBalance` — Account, AccountType, Amount.
- `FinancialReport` — Revenue, COGS, gross profit/margin, operating expenses/profit, tax, net profit.
- `Validation` — balance checks and unmapped accounts.

## Module 9.1 — Design the Architecture

Export these standard modules:

| Module | Responsibility |
|---|---|
| `modMain` | Public controller only |
| `modConfig` | Constants and configuration loading |
| `modImport` | File selection and source import |
| `modValidation` | Schema, account mapping, balance checks |
| `modCalculations` | Pure typed financial functions |
| `modReporting` | Worksheet output and formatting |
| `modUtilities` | Shared generic helpers |

**Rule:** Calculation functions do not select sheets, open files, display dialogs, or format cells.

## Module 9.2 — Build Typed Functions

```vba
Public Function GrossProfit(ByVal revenue As Double, ByVal cogs As Double) As Double
    GrossProfit = revenue - cogs
End Function

Public Function MarginPercent(ByVal profit As Double, ByVal revenue As Double) As Variant
    If revenue = 0 Then
        MarginPercent = CVErr(xlErrDiv0)
    Else
        MarginPercent = profit / revenue
    End If
End Function
```

- Use `ByVal` for scalar inputs that should not be changed.
- Use `ByRef` only when mutation is part of the documented contract.
- Add an optional tax-rate argument with a clear default.
- Use `ParamArray` in a utility exercise, not when a typed array would provide a safer production API.
- Keep public surface area small; default helpers to `Private`.

## Module 9.3 — Orchestrate the Report

`RunFullReport()` should:

1. Load and validate configuration.
2. Import or refresh source data.
3. Validate headers and account mappings.
4. Calculate totals and ratios.
5. Generate the report.
6. Write validation/reconciliation results.
7. Show one final status message.

Each stage returns success/failure or raises a documented error; the controller decides whether to continue.

## Module 9.4 — Self-Tests

Create `RunCalculationTests()` with known inputs:

- Gross profit: 500,000 − 300,000 = 200,000.
- Gross margin: 200,000 / 500,000 = 40%.
- Zero revenue returns the documented error/blank behavior.
- Default tax rate is used when omitted.
- Explicit 15% tax overrides the default.
- Input arguments remain unchanged after `ByVal` calls.

Write results to the Immediate window or a dedicated test sheet.

## Styling Pass

- Theme: dark navy `#1F4E78`, white, slate gray, gold highlight.
- Use a report title, period, currency, generated timestamp, and version.
- Distinguish inputs/assumptions from calculated output.
- Use standard accounting presentation: negatives in parentheses, zero as dash if requested.
- Add a prominent validation status: Balanced, Warning, or Failed.
- Never use styling to hide an unreconciled result.

## Test Gate

- [ ] Every public function has a typed return and documented parameters.
- [ ] Zero-denominator behavior is intentional.
- [ ] Default and explicit optional arguments pass.
- [ ] Import failure stops calculation cleanly.
- [ ] Unmapped accounts appear in Validation.
- [ ] Report totals reconcile to source classifications.
- [ ] `modMain` reads as a short workflow, not an implementation dump.
- [ ] Exported `.bas` files are included in source control.

## Freelance Simulation

**Brief:** A finance manager needs quarterly reporting; accounting classifications are supplied by the client.  
**Professional boundary:** State that the tool automates provided rules and does not replace accounting review.  
**Maintainability review:** Ask another developer to locate and change the tax default; target under five minutes.  
**Change request:** Add EBITDA without modifying import logic. Prepare impact analysis and quote.

---

# Chapter 10 — Arrays

## Project: In-Memory Data Transformation Engine

**Client outcome:** Large exports are transformed in memory and written back efficiently.  
**Primary concepts:** 2D Variant arrays, bounds, `ReDim Preserve`, `Split`, `Join`, `Filter`, bulk I/O.  
**Target build time:** 4–6 hours.

## Workbook Blueprint

| Sheet | Purpose |
|---|---|
| `Instructions` | Schema, size guidance, run steps |
| `Control` | Transform and benchmark controls |
| `RawData` | ID, Name, Email, Department, Salary, Tags |
| `Transformed` | Source columns plus DepartmentCode, NormalizedTags, VIPFlag |
| `Benchmark` | Method, rows, seconds, checksum, status |
| `Exceptions` | Row ID and validation problem |

## Module 10.1 — Bulk Load and Inspect

- Validate `RawData` before timing.
- Set the source range from explicit boundaries or `CurrentRegion` after confirming it has no internal blank header.
- Load values in one statement: `data = sourceRange.Value2`.
- Read row and column counts with `LBound`/`UBound`; do not assume a zero lower bound.
- Create a separate output array sized for source plus three output columns.

**Checkpoint:** 50,000 source rows are loaded without any per-cell read inside the processing loop.

## Module 10.2 — Transform in Memory

For each array row:

1. Copy source values to the output array.
2. Map department to FIN, TECH, SALES, HR, or UNKNOWN.
3. Parse tags with `Split`.
4. Trim and uppercase each tag.
5. Reassemble with `Join`.
6. Assign VIP only when a complete normalized tag equals `VIP`, not merely when text contains those letters.
7. Validate salary and email; record exceptions separately.

Use a dictionary for department mapping if available, with late binding if you want to avoid a required reference.

## Module 10.3 — Dynamic Arrays Correctly

Practice `ReDim Preserve` in a small exercise:

- VBA preserves data only when resizing the **last dimension** of a multi-dimensional array.
- For growing record counts, either make records the last dimension, collect into a 1D structure, preallocate capacity, or create a correctly sized output array after a counting pass.
- In the production transform, prefer pre-sizing the output array; repeated `ReDim Preserve` is costly.

## Module 10.4 — Bulk Write and Benchmark

```vba
wsOut.Cells.ClearContents
wsOut.Range("A1").Resize(UBound(outputData, 1), _
    UBound(outputData, 2)).Value2 = outputData
```

- Write output once.
- Apply formats after the bulk write.
- Benchmark against a correct cell-by-cell implementation on identical data.
- Compute a checksum or compare sampled rows so speed cannot mask incorrect output.
- Clear large arrays with `Erase` when no longer needed.

## Styling Pass

- Theme: midnight blue `#203864`, bright teal `#00B0F0`, soft gray.
- Control cards show rows, duration, rows/second, and output exceptions.
- Avoid formatting all 1,048,576 rows; format only actual output.
- Use filters and freeze panes, but be cautious converting very large outputs to tables if performance suffers.
- Benchmark chart compares methods with honest axes and labeled row counts.

## Test Gate

| Test | Expected |
|---|---|
| 1-row dataset | Correct bounds and output |
| 50,000 rows | Completes within the locally agreed target |
| Unknown department | `UNKNOWN` and/or exception |
| Tags `vip,new` | `VIP,NEW`; VIP flag true |
| Tag `viper` | VIP flag false |
| Blank tags | Blank normalized value; no error |
| Invalid salary | Exception; no type mismatch |
| Array vs cell method | Same checksum and sampled outputs |
| RawData after run | Unchanged |

## Freelance Simulation

**Brief:** Transform recurring HR exports of 50,000–100,000 rows under a defined schema.  
**Benchmark report:** State hardware/environment, workbook size, row count, method, elapsed time, and correctness check.  
**Scope question:** What happens when the export exceeds an Excel sheet’s row limit? Propose Power Query, database, or CSV alternatives.  
**Portfolio demo:** Show raw data, click Transform, show benchmark and a filtered VIP result.

---

# Capstone Delivery Sprint

After completing all ten projects, choose Chapters 4, 7, and 10—or three projects matching your target market—and run this delivery sprint.

## Day 1 — Repository and Source Review

- [ ] Normalize folder structure.
- [ ] Export all VBA modules.
- [ ] Remove personal paths, usernames, and client data.
- [ ] Add a license or usage statement.
- [ ] Confirm `.xlsm` files are intentional and within repository size limits.

## Day 2 — Documentation

Every README should contain:

1. Business problem and result.
2. Features and chapter skills demonstrated.
3. Requirements and supported Excel platform.
4. Installation and macro-security instructions.
5. Step-by-step usage.
6. Input schema and sample data.
7. Expected output.
8. Test evidence.
9. Known limitations.
10. Screenshots/GIF and version history.

## Day 3 — Client Acceptance Pack

Prepare:

- A clean delivery ZIP or repository release.
- Original/backup workbook and sample data.
- One-page quick-start guide.
- Acceptance checklist signed off against the agreed scope.
- Change log and version number.
- A short support window and clear out-of-scope policy.

## Day 4 — Portfolio and Proposal Assets

For each selected project, write:

- **Title:** Outcome first, technology second.
- **50-word summary:** Client problem, automated workflow, measurable result.
- **Technical summary:** Objects, procedures, event/performance patterns, tests.
- **Three screenshots:** Input, action/progress, output.
- **15-second demo:** One workflow without editing cuts.
- **Proposal paragraph:** Mirror the buyer’s problem and cite the relevant project.

## Day 5 — Final Quality Review

Score each item from 1 to 5:

| Area | 1 | 3 | 5 |
|---|---|---|---|
| Functionality | Core flow fails | Standard cases pass | Edge cases and recovery pass |
| Code quality | Monolithic/implicit | Clear modules and types | Small APIs, cleanup, self-tests |
| UX | Requires VBA editor | Buttons and instructions | Accessible, polished, low-friction |
| Data safety | Overwrites raw input | Basic separation | Backup/recovery and audit trail |
| Documentation | Missing | Setup and usage | Tests, assumptions, limits, changes |
| Freelance readiness | Demo only | Deliverable package | Scoped, priced, supported, evidenced |

**Release threshold:** No category below 3, average at least 4, and every project-specific test gate complete.

---

# Reusable Freelance Templates

## Discovery Questions

1. What manual process should the workbook replace?
2. What are the exact input files, sheets, headers, formats, and typical row counts?
3. What output proves the job is complete?
4. Which Excel versions, operating systems, locales, and security restrictions apply?
5. What should happen to invalid, missing, duplicate, or partial records?
6. Must source data remain unchanged?
7. Who runs the tool, and how technical are they?
8. Is the workbook shared, and is simultaneous use expected?
9. Which actions require confirmation, approval, or an audit trail?
10. What is explicitly out of scope for version 1?

## Scope Template

```text
Objective:
Inputs supplied by client:
Workflow to automate:
Outputs/deliverables:
Business rules:
Acceptance tests:
Excel/platform requirements:
Assumptions:
Out of scope:
Timeline and milestones:
Price and payment terms:
Support period:
Change-request process:
```

## Progress Update Template

```text
Completed:
- ...

Demonstrated/tested:
- ...

In progress:
- ...

Decision or file needed from client:
- ...

Risks/assumptions:
- ...

Next milestone and date:
- ...
```

## Handoff Message Template

```text
The completed workbook and sample data are attached. Start on the Instructions
sheet, enable macros only after confirming the file source, and test first with
the included sample. The agreed acceptance tests have passed and are listed in
the README. Version, assumptions, and known limitations are documented there.
Please keep an untouched backup before using production data. The included
support period covers defects within the agreed scope; enhancements are quoted
separately.
```

## Change Request Log

| ID | Requested change | Reason | Scope impact | Time | Price | Approved | Version |
|---|---|---|---|---:|---:|---|---|
| CR-001 | Example | Example | Modules/tests/docs affected | 0.0 h | $0 | Pending | 1.1 |

---

# Final Portfolio Checklist

- [ ] Ten projects map clearly to the ten course chapters.
- [ ] Each project has a workbook blueprint and ordered module plan.
- [ ] Each includes a distinct styling direction.
- [ ] Each includes chapter-specific acceptance tests.
- [ ] Each includes realistic freelance discovery, scope, or handoff work.
- [ ] Source modules are exportable and reviewable outside Excel.
- [ ] Claims about security, auditability, and speed are appropriately qualified.
- [ ] Sample data is synthetic and contains no client information.
- [ ] Every workbook has a version, instructions, assumptions, and limitations.
- [ ] Portfolio assets demonstrate outcomes rather than showing code alone.

This workbook is an original companion plan. It applies the course topics to independent, client-oriented project builds and does not reproduce course exercises verbatim.
