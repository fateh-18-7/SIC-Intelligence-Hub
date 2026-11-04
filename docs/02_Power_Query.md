# Phase 5: Power Query automation

Status: **UNVERIFIED.** The M code was written to the Power Query spec, but no Power Query engine is available in my session. The Excel-side effects of a new file were simulated on the verified workbook copy; see "Demo" below.

## 1. Folder layout

```
SIC_Intelligence_Hub\data\                 Sales_2026-08.csv, Sales_2026-09.csv, Sales_2026-10.csv
SIC_Intelligence_Hub\demo_hold_back\       Sales_2026-11.csv  (move into data\ during the demo)
```

## 2. Create the parameter

Data > Get Data > Launch Power Query Editor > Home > Manage Parameters > New.
Name `pDataFolder`, Type Text, Current value = the full path of your data folder **ending with a backslash**, for example `C:\Users\you\Documents\SIC_Intelligence_Hub\data\`.

## 3. Queries

For each query: Home > New Source > Other Sources > Blank Query. Then Home > Advanced Editor, paste the code, and rename the query to the name shown.

### qSalesFiles
```
let
    Source = Folder.Files(pDataFolder),
    OnlyTopFolder = Table.SelectRows(Source, each [Folder Path] = pDataFolder),
    SalesCsv = Table.SelectRows(OnlyTopFolder, each Text.StartsWith([Name], "Sales_") and Text.Lower([Extension]) = ".csv")
in
    SalesCsv
```
Why: Folder.Files also reads subfolders. The filters keep only `Sales_*.csv` files in the top folder.

### qSchemaCheck
```
let
    Expected = Excel.CurrentWorkbook(){[Name="tExpectedSchema"]}[Content][ColumnName],
    Files = qSalesFiles,
    WithColumns = Table.AddColumn(Files, "Columns", each Table.ColumnNames(Table.PromoteHeaders(Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true]))),
    Missing = Table.AddColumn(WithColumns, "MissingColumns", each Text.Combine(List.Difference(Expected, [Columns]), ", "), type text),
    Extra = Table.AddColumn(Missing, "UnexpectedColumns", each Text.Combine(List.Difference([Columns], Expected), ", "), type text),
    Ok = Table.AddColumn(Extra, "SchemaOK", each [MissingColumns] = "" and [UnexpectedColumns] = "", type logical),
    Result = Table.RenameColumns(Table.SelectColumns(Ok, {"Name", "MissingColumns", "UnexpectedColumns", "SchemaOK"}), {{"Name", "SourceFile"}})
in
    Result
```
Why: the expected column list comes from the workbook table `tExpectedSchema` (sheet Schema), so it isn't hardcoded. A renamed column shows up as one missing column plus one unexpected column.

### qRawCapture (counts BEFORE cleaning)
```
let
    Files = qSalesFiles,
    Parsed = Table.AddColumn(Files, "Data", each Table.PromoteHeaders(Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true])),
    Rows = Table.AddColumn(Parsed, "RowsIn", each Table.RowCount([Data]), Int64.Type),
    Amount = Table.AddColumn(Rows, "AmountLocalIn", each List.Sum(List.Transform(Table.Column([Data], "AmountLocal"), each Number.From(_, "en-US"))), type number),
    Result = Table.RenameColumns(Table.SelectColumns(Amount, {"Name", "RowsIn", "AmountLocalIn"}), {{"Name", "SourceFile"}})
in
    Result
```
Why: it reads each file directly and independently of qSales, so check C1 compares two separate paths.

### qSales (typed load into tSales)
```
let
    Expected = Excel.CurrentWorkbook(){[Name="tExpectedSchema"]}[Content][ColumnName],
    Files = qSalesFiles,
    Parsed = Table.AddColumn(Files, "Data", each Table.PromoteHeaders(Csv.Document([Content], [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]), [PromoteAllScalars=true])),
    Keep = Table.SelectColumns(Parsed, {"Name", "Data"}),
    Expanded = Table.ExpandTableColumn(Keep, "Data", Expected),
    Renamed = Table.RenameColumns(Expanded, {{"Name", "SourceFile"}}),
    BlankToNull = Table.ReplaceValue(Renamed, "", null, Replacer.ReplaceValue, {"EmployeeID"}),
    Typed = Table.TransformColumnTypes(BlankToNull, {{"TransactionID", type text}, {"Date", type date}, {"EmployeeID", type text}, {"Region", type text}, {"ProductFamily", type text}, {"Currency", type text}, {"AmountLocal", type number}, {"SourceFile", type text}}, "en-US"),
    Ordered = Table.SelectColumns(Typed, {"TransactionID", "Date", "EmployeeID", "Region", "ProductFamily", "Currency", "AmountLocal", "SourceFile"})
in
    Ordered
```
Why: the culture `"en-US"` makes ISO dates and dot decimals load correctly on a Hungarian-locale PC. SourceFile comes from the file name.

Validation stays in Excel calculated columns. That keeps every rule visible and explainable in one place. Power Query does the schema check, raw capture and typing.

## 4. One-time switch from the static tables to the query tables (about 10 minutes)

Do this on a copy of the workbook first. First unprotect Control_Waterfall (Review > Unprotect Sheet; there's no password).

1. In the Queries & Connections pane, right-click qSalesFiles > Load To > **Only Create Connection**. Do the same for qRawCapture and qSales.
2. Right-click qSchemaCheck > Load To > Table > New worksheet. Rename that sheet `Schema_Check`.
3. Right-click qSales > Load To > Table > New worksheet. In Table Design, rename the new table to **tSalesNew**.
4. Add the calculated columns to tSalesNew. Type each header, then paste its formula from `Formula_Reference.md` (section tSales) into the first data row, in this order: Period, TxnMonth, IsDuplicate, MissingID, KnownEmp, ActiveOnDate, RegionOK, FXRate, AmountEUR, IsCreditNote, IsLate, Status, Reason, EligibleEUR, Severity, WorkflowStatus. In IsDuplicate, write `tSalesNew[TransactionID]` instead of `tSales[TransactionID]`.
5. Rename the old table on Sales_Raw from `tSales` to **tSales_old**. Excel repoints every formula to tSales_old.
6. Rename **tSalesNew** to **tSales**.
7. Home > Find & Select > Replace. Find `tSales_old[`, Replace with `tSales[`, Within: Workbook, Look in: Formulas, then Replace All.
8. Check that Controls shows the same results as before: only C8 failing, and RUN STATUS = HOLD.
9. Repeat steps 3 and 5–8 for the raw capture: load qRawCapture to a table named tRawCaptureNew, rename the old table to tRawCapture_old, rename the new one to tRawCapture, then replace `tRawCapture_old[` with `tRawCapture[`.
10. Delete the old Sales_Raw sheet and the old tRawCapture range. Add a Change_Log row.

Conflict flagged: if you deleted the old tables before steps 5–7, formulas would turn into #REF! permanently. That's why you rename first.

## 5. Run Log (no macros)

Sheet Run_Log, cells B4:B9, holds a live snapshot: files loaded, rows in, rows held, waterfall variance, checks failing and status. After each Refresh All:

1. Copy B4:B9, then Paste Special > Values (Transpose) into the next empty tRunLog row.
2. Type the timestamp: press Ctrl+; then a space, then Ctrl+Shift+;.

Limitation: logging is manual. `=NOW()` changes every time the workbook recalculates, so it can't record when a run happened. Only a macro or Power Automate can write a log automatically.

## 6. Demo: drop Sales_2026-11.csv and press Refresh All

I simulated this on the verified workbook copy (**VERIFIED** for the Excel side; the Power Query refresh itself is **UNVERIFIED**):

| Step | What you'll see |
|---|---|
| Move `demo_hold_back\Sales_2026-11.csv` into `data\`, then Data > Refresh All | Schema_Check shows 4 files, all SchemaOK = TRUE. Raw rows go from 59 to 73. Waterfall variance stays 0 |
| Status column | All 14 November rows are **Held - Missing FX rate**, because tFX has no 2026-11 rate, not even for EUR. C8 rises to 18, and the status is HOLD |
| Add a tFX row: EUR, 2026-11, 1 | The 14 rows become Valid. **C5 RosterGap = 14** and **C2a cross-cut = 515,000**, because there are no November payout rows or quotas yet. The status stays HOLD |
| Message for the interview | "The new month loads in seconds, nothing is lost (waterfall = 0), and the controls stop me paying until reference data, roster and quotas are in place." |

Fallback if the refresh fails live: open `demo_hold_back` and show the screenshots you took during rehearsal. Then walk through the waterfall using the static table.

## 7. Optional stretch: Power Automate (LICENCE-DEPENDENT, verify availability)

Flow: **When a file is created in a folder** (OneDrive for Business, folder `/SIC/data`) → **Condition**: the file name starts with `Sales_` → **Post message in a chat or channel** (Teams) or **Send an email**, with **counts only**: file name, size, time → **Start and wait for an approval** (approver: SIC Lead) → **Post the outcome**.

What stays manual: opening Excel and pressing Refresh All, reviewing exceptions, and the Excel approval cells. A cloud flow can't refresh a desktop workbook's queries without extra tooling. The notification must never include employee-level amounts.

## 8. Timing protocol (sheet Timing, table tTiming)

1. **Manual run:** paste three CSVs into tSales by hand, then check statuses. Time it with a stopwatch.
2. **Refresh run:** drop the files into data\ and press Refresh All.
3. **For each method:** do 3 runs. Record machine seconds (until the calculation finishes) and review minutes (reading Controls and exceptions) separately.
4. **Read the medians** at row 9. They show "not measured" until filled.

Put only measured numbers on your CV.
