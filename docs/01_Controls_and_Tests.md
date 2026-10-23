# Phases 3–4: Controls, exception workflow, gate, tests

All data is SYNTHETIC. Exact formulas: `Formula_Reference.md`.

## Tables and sheets added

| Sheet | Table / cells | Purpose |
|---|---|---|
| Controls | tControls (ControlID, Check, EvidenceSheet, Value, Result), D18 checks failing, D20–D22 approval inputs, D23 RUN STATUS | Check cards and status gate |
| Control_Waterfall | tRawCapture (SourceFile, RowsIn, AmountLocalIn), E5:G11 bucket block | C1 raw-to-clean waterfall |
| Control_CrossCut | tCrossCut (Region, AIS, PDx, PCS, RowTotal), C4, G4 | C2 cross-cut |
| Control_Recalc | tRecalc (step-by-step tiers, Variance, InPrior, InNext) | C3 independent recalculation |
| Control_RollForward | tRollForward (OpeningExpense … Variance, MemoLateItemsSalesEUR) | C4 period roll-forward |
| Exceptions | tExceptions (ExceptionID, Period, RecordID, Type, Severity, Owner, Status, CreatedDate, ResolvedDate, Resolution, AgeDays, AgeBucket) | Manual exception workflow |
| Exception_Log | spill range from A6 (paste formula) | Formula-driven log, separate from the workflow table |
| Exception_Types | tExceptionTypes (Type, Severity, DefaultOwner) | Severity master |
| Tests | tTests (TestID, Scenario, Input, ExpectedPayout, ActualPayout, Working, PassFail), C3 failed tests | Test suite |
| Controls_Matrix | tControlsMatrix (ControlID, Risk, Control, Evidence, TestID) | Risk → control → evidence → test |
| Change_Log | tChangeLog | Change history |
| Parameters | p_AsOfDate, p_AgeBucket1, p_AgeBucket2, p_ReportPeriod | No hardcoded thresholds |

New columns: tSales gets Severity and WorkflowStatus. tPayout gets Key, ExceptionType, Severity, WorkflowStatus and PayoutRegion.

## Why each control is independent

| Control | Compares | Why it isn't circular |
|---|---|---|
| C1 waterfall | Rows and amounts counted from the files (tRawCapture) vs the Valid + Held + Duplicate + Excluded statuses in tSales | The raw counts come from the files before any cleaning. In Phase 5 they come from their own Power Query query (qRawCapture) |
| C2 cross-cut | Region × family totals vs employee totals in tPayout | Two different aggregation paths. **Limit:** both read EligibleEUR, so C2 catches mapping and roster errors, not a wrong EligibleEUR. C1 and the tests cover that |
| C3 recalc | A step-by-step INDEX/MATCH + MIN/MAX tier calculation vs tPayout's LET/SUMPRODUCT | Different formulas, a separately typed roster, sales re-summed from AmountEUR rather than EligibleEUR |
| C4 roll-forward | Prior expense + movements vs current expense | Opening and closing come from tPayout. The movements come from tRecalc and tAdjustments |

## Status gate

```
Controls!D23  =IF(D18>0,"HOLD",IF(AND(LEN(TRIM(p_ApproverName))>1,ISNUMBER(p_ApprovalDate)),IF(AND(ISNUMBER(p_ApprovedTotal),ABS(p_ApprovedTotal-SUM(tPayout[FinalPayoutEUR]))<0.005),"APPROVED","READY FOR REVIEW - approval stale or total missing"),"READY FOR REVIEW"))
```
How the gate works:

- **Any check not passing means HOLD.** Each Result is FAIL when its value is an error, so a broken formula can't show green.
- **APPROVED needs three things a person types:** a name, a date, and the approved total. If any number changes after approval, the totals no longer match and the status drops back to READY FOR REVIEW.
- **Gate severity comes from the data, not the workflow table.** Severity is looked up from tExceptionTypes, so editing the workflow can't downgrade a High exception. Closing an exception needs a Resolution (check C9).

### Gate scenarios tested (VERIFIED, via the INDEX/MATCH twin)

| Scenario | Result |
|---|---|
| As delivered | HOLD (C8 = 4 open High: T1035, T1042, T1044, E106/2026-10) |
| Approver entered while HOLD | HOLD + "Approval entered while HOLD - ignored" |
| All exceptions closed with resolutions | READY FOR REVIEW |
| Closed without resolution | HOLD (C9 = 7) |
| Closed + name + date + correct total 13,930 | APPROVED |
| Approved, then a quota changed | READY FOR REVIEW - approval stale |

## Exception log (spill), pasted by you in Exception_Log!A6

```
=VSTACK(FILTER(HSTACK(tSales[TransactionID],tSales[Period],tSales[EmployeeID],tSales[Reason],tSales[Severity],tSales[WorkflowStatus]),tSales[Severity]<>""),FILTER(HSTACK(tPayout[Key],tPayout[Period],tPayout[EmployeeID],tPayout[ExceptionType],tPayout[Severity],tPayout[WorkflowStatus]),tPayout[Severity]<>""))
```
I left this for you to paste because a script can't write dynamic-array metadata. Written by script, Excel could treat the formula as legacy, prefix it with "@", and show only one row. Expected result: 11 rows (9 transaction exceptions + 2 payout-level). Requires Excel 365. **UNVERIFIED.**

## Test results: 18 tests, 18 PASS (VERIFIED)

The **expected** values were worked out by hand from the data, and the reasoning is in the Working column. The **actual** values are pulled from the engine by formula.

| Test | Scenario | Expected | Actual |
|---|---|---|---|
| TS01–TS05 | Boundaries 79.99 / 80.00 / 80.01 / 100.00 / 100.01 % | 0 / 0 / 0.20 / 400 / 400.40 | match |
| TS06 | Missing quota | HOLD - missing quota | match |
| TS07 | Inactive recipient | Excluded | match |
| TS08 | Duplicate counted once | 13,000 | match |
| TS09–TS10 | Unknown employee / blank ID | Held | match |
| TS11 | Credit note (E102 Sep) | 360 | match |
| TS12 | HUF conversion | 113,000 | match |
| TS13 | Missing FX row excluded (E105 Oct) | 720 | match |
| TS14 | Plan change (E101 Oct, v2) | 2,800 | match |
| TS15 | Unapproved adjustment excluded | 2,800 | match |
| TS16 | Territory change (E104 Sep) | 400 | match |
| TS17 | Late item credited in Oct | 148,000 | match |
| TS18 | Approved negative adjustment | 1,700 | match |

**Known limitation.** TS01–TS05 read tSanity, which uses the same formula text as tPayout, because no live payout row sits at those attainments. They show the formula's logic is right, not that the live cells are. TS11–TS18 do read live tPayout and tSales cells.

## Protection and conflicts flagged

| Feature | Conflict | Resolution |
|---|---|---|
| Sheet protection | Power Query can't refresh into a protected sheet, and tables on protected sheets can't grow | Sales_Raw and the raw-capture load sheet are **not** protected. Protected sheets have no password; it's a deterrent against accidental edits, not security |
| Calculated columns in a Power Query output table | Excel keeps them and fills them for new rows | See 02_Power_Query.md for the migration steps. **UNVERIFIED** until you run it |
| FILTER spill | Can't live inside a table | It sits on a separate sheet in plain cells |
| tRecalc PlanVersion uses approximate MATCH | Needs tPlan sorted by EffectiveFrom | Stated in the sheet note. A wrong order would show up as C3 variances |

## Self-audit

- **The ageing date is a fixed parameter.** p_AsOfDate is typed, not `=TODAY()`, so results are reproducible. Update it each run.
- **Prior-period exceptions stay in the gate.** The gate counts open exceptions from every period, which is a conservative choice. If you'd rather filter by period, that's a policy decision.
- **The waterfall can't catch a row lost before capture.** If a file is truncated before Power Query reads it, both sides miss the row. You'd need a control total from the source system.
