# Phase 6: Power BI (beginner steps)

Status: **UNVERIFIED.** Power BI isn't available in my session. Use the expected values in section 9 to check your build. Copilot and Smart Narrative are **LICENCE-DEPENDENT; verify availability.**

## 1. Before you start

- Open the workbook in Excel and press **Save** once. Power BI reads the values stored in the file, not the formulas, so always save in Excel after any change, then click Refresh in Power BI.
- Use Power BI Desktop (free) on Windows.

## 2. Load the tables

Home > Get data > Excel workbook > select `SIC_Intelligence_Hub.xlsx`. In the Navigator, tick these **tables** (blue table icon; ignore the sheets): `tSales`, `tPayout`, `tRecipients`, `tExceptions`, `tControls`, `tStatus`. Click **Transform Data**, not Load.

### Clean blanks (prevents the most common beginner error)

Excel shows "" for not-applicable values. Power BI then sees text in a number column. Fix it per table: select the query, then Home > Advanced Editor, and add the steps below before `in`. Change the last line to the new step name.

tPayout (the last step before your edit is usually `#"Changed Type"`; check the name in the editor):
```
    Blanks = Table.ReplaceValue(#"Changed Type", "", null, Replacer.ReplaceValue, {"QuotaEUR", "Attainment", "Tier1EUR", "Tier2EUR", "Tier3EUR", "CalcPayoutEUR", "FinalPayoutEUR", "ExceptionType", "Severity", "WorkflowStatus"}),
    Typed = Table.TransformColumnTypes(Blanks, {{"PeriodStart", type date}, {"EligibleSalesEUR", type number}, {"QuotaEUR", type number}, {"Attainment", type number}, {"Tier1EUR", type number}, {"Tier2EUR", type number}, {"Tier3EUR", type number}, {"CalcPayoutEUR", type number}, {"ApprovedAdjEUR", type number}, {"PendingAdjEUR", type number}, {"FinalPayoutEUR", type number}, {"HeldTxnCount", Int64.Type}})
```
tSales:
```
    Blanks = Table.ReplaceValue(#"Changed Type", "", null, Replacer.ReplaceValue, {"EmployeeID", "FXRate", "AmountEUR", "Reason", "Severity", "WorkflowStatus"}),
    Typed = Table.TransformColumnTypes(Blanks, {{"Date", type date}, {"AmountLocal", type number}, {"FXRate", type number}, {"AmountEUR", type number}, {"EligibleEUR", type number}})
```
Close & Apply.

## 3. Model (star schema)

Modeling > New table, one at a time:
```
Calendar = ADDCOLUMNS(CALENDAR(DATE(YEAR(MIN(tSales[Date])), MONTH(MIN(tSales[Date])), 1), EOMONTH(MAX(tSales[Date]), 0)), "Period", FORMAT([Date], "yyyy-mm"), "MonthStart", DATE(YEAR([Date]), MONTH([Date]), 1), "Month", FORMAT([Date], "mmm yyyy"))
```
```
DimRegion = DISTINCT(UNION(SELECTCOLUMNS(tSales, "Region", tSales[Region]), SELECTCOLUMNS(tPayout, "Region", tPayout[PayoutRegion])))
```
Select Calendar, then Table tools > **Mark as date table** > Date column = Date.

New column on tSales. This credits sales to the processing period, so a late item lands in October, the same as the payout:
```
PeriodStart = DATE(VALUE(LEFT(tSales[Period], 4)), VALUE(RIGHT(tSales[Period], 2)), 1)
```

Relationships (Model view; drag the field from the "1" side onto the "many" side; all single-direction):

| From (1) | To (many) |
|---|---|
| Calendar[Date] | tPayout[PeriodStart] |
| Calendar[Date] | tSales[PeriodStart] |
| tRecipients[EmployeeID] | tPayout[EmployeeID] |
| tRecipients[EmployeeID] | tSales[EmployeeID] |
| DimRegion[Region] | tPayout[PayoutRegion] |
| DimRegion[Region] | tSales[Region] |

tExceptions, tControls and tStatus stay unrelated, as standalone tables. Note that tPayout[PayoutRegion] is the region at period end, so E104's whole September payout counts under West.

## 4. Measures

Create a table to hold them: Home > Enter data > name it `_Measures` > Load. Then select it and use New measure for each of these.
```
Eligible Sales = SUM(tPayout[EligibleSalesEUR])
Total Quota = SUM(tPayout[QuotaEUR])
Eligible Sales With Quota = CALCULATE([Eligible Sales], NOT ISBLANK(tPayout[QuotaEUR]))
Overall Attainment = DIVIDE([Eligible Sales With Quota], [Total Quota])
Total Incentive = SUM(tPayout[FinalPayoutEUR])
Cost of Sales % = DIVIDE([Total Incentive], [Eligible Sales])
Incentive Prior Month = CALCULATE([Total Incentive], DATEADD('Calendar'[Date], -1, MONTH))
Incentive MoM % = DIVIDE([Total Incentive] - [Incentive Prior Month], [Incentive Prior Month])
Exception Count = CALCULATE(COUNTROWS(tSales), NOT ISBLANK(tSales[Severity]))
Exception Rate = DIVIDE([Exception Count], COUNTROWS(tSales))
Open High Exceptions = MAX(tStatus[OpenHighExceptions])
Checks Failing = MAX(tStatus[ChecksFailing])
Readiness Status = SELECTEDVALUE(tStatus[RunStatus], "Unknown")
Banner Text = "Run " & SELECTEDVALUE(tStatus[ReportPeriod]) & ": " & [Readiness Status]
Status Colour = VAR s = [Readiness Status] RETURN SWITCH(TRUE(), s = "APPROVED", "#2E7D32", LEFT(s, 5) = "READY", "#F9A825", "#C62828")
Exceptions Open = CALCULATE(COUNTROWS(tExceptions), tExceptions[Status] <> "Closed")
Exceptions Resolved = CALCULATE(COUNTROWS(tExceptions), tExceptions[Status] = "Closed")
```
Why Overall Attainment is weighted: total sales ÷ total quota treats €1 the same everywhere. An average of each person's percentage would give a small quota the same weight as a large one. E106's October row has no quota, so it's left out of both sides.

Formats: select each measure, then the Measure tools ribbon. Currency for the EUR measures; Percentage with 1 decimal for Attainment, MoM and Exception Rate; 2 decimals for Cost of Sales %.

## 5. Pages and visuals

### Page 1: Executive
```
+--------------------------------------------------------------+
| [Card: Banner Text, background = Status Colour]              |
+-----------+-----------+-----------+-----------+--------------+
| Total     | Eligible  | Overall   | Cost of   | Incentive    |
| Incentive | Sales     | Attainment| Sales %   | MoM %        |
+-----------+-----------+-----------+-----------+--------------+
| Line chart: Total Incentive by Calendar[Month]  | Bar chart:  |
|                                                 | Total Inc.  |
|                                                 | by DimRegion|
+-------------------------------------------------+-------------+
| Narrative (Smart Narrative / Copilot - verify availability)   |
+--------------------------------------------------------------+
Slicer: Calendar[Period]
```
To colour the banner: select the Card > Format > Visual > Callout value > Color > fx > Format style = **Field value** > Status Colour. (In older versions it's under Format > Background.)

### Page 2: Performance
- Clustered column chart: X = tRecipients[EmployeeID], Y = Eligible Sales and Total Quota.
- Table: EmployeeID, Period, Eligible Sales, Total Quota, Overall Attainment, Total Incentive, tPayout[ReviewStatus].
- Attainment distribution: Modeling > New column on tPayout: `AttainmentBand = IF(ISBLANK(tPayout[Attainment]), "No quota", IF(tPayout[Attainment] < 0.8, "<80%", IF(tPayout[Attainment] < 1, "80-100%", ">=100%")))`, then a column chart of the row count by AttainmentBand.
- Drill-through: create a page called "Recipient", drag tRecipients[EmployeeID] into **Drill through**, and add cards and a table. From Performance, right-click a bar > Drill through > Recipient.

### Page 3: Controls
- Table: tControls ControlID, Check, Value, Result. Conditional formatting on Result: Background color > Rules: equals "PASS" green, otherwise red.
- Cards: Checks Failing, Open High Exceptions, Exceptions Open, Exceptions Resolved.
- Stacked bar: count of tExceptions by Type, legend Severity.
- Column chart: count of tExceptions by AgeBucket, legend Status.

## 6. Row-level security

Modeling > Manage roles > New:
- **Regional Manager**: table DimRegion, rule `[Region] = "East"`. This is a static demo; a real setup maps users to regions and uses `USERPRINCIPALNAME()`.
- **SIC Analyst**: no rule, so it sees everything.

To test: Modeling > View as > tick Regional Manager. Totals should drop to East only. tExceptions, tControls and tStatus aren't linked to DimRegion, so a manager still sees them. In a real deployment you'd hide those pages from that role or filter them too.

## 7. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Measure is blank | No relationship, or a relationship in the wrong direction | Model view: check the 1 → * arrows from the table above |
| Numbers are all identical across months | A visual uses tSales[Date] while the relationship is on PeriodStart | Use Calendar fields on the axis |
| "We couldn't convert to Number" | "" text in a number column | Apply the Blanks step from section 2 |
| Old numbers after an Excel change | The file wasn't saved after recalculation | Save in Excel, then Refresh in Power BI |
| MoM % is blank | Calendar not marked as a date table, or no prior month in the filter | Mark as date table; Aug has no prior month, so blank there is correct |
| Duplicate-value error on a relationship | DimRegion or tRecipients has repeats or blanks | DimRegion uses DISTINCT. A blank EmployeeID in tSales is fine on the many side |

## 8. Smart Narrative / Copilot

Insert > Narrative (Smart narrative or Copilot, depending on your version and licence). If neither is available, show the AI draft from Excel's AI_Prompt and Claim_Verification sheets instead. Any narrative is commentary only. It never changes numbers or status.

## 9. Expected values (from the verified workbook)

| Measure | All periods | Aug | Sep | Oct |
|---|---|---|---|---|
| Total Incentive | 13,930 | 5,120 | 3,580 | 5,230 |
| Eligible Sales | 2,497,000 | | | 847,000 |
| Overall Attainment | | | | 102.6% |
| Incentive MoM % | | blank | −30.1% | +46.1% |
| Readiness Status | HOLD | | | |
| Open High Exceptions | 4 | | | |
