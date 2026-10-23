# Formula reference

Generated from the delivered workbook, so every formula below is exactly what the file contains. Type them in Excel as shown. `[@Col]` means this row of the same table.

## tControls (sheet Controls)

```
Column5               =IF(ISNUMBER([@Value]),IF([@Value]=0,"PASS","FAIL"),"FAIL")
```

## Single cells on Controls

```
D18   Checks not passing                        =COUNTIF(tControls[Result],"<>PASS")
E22   Approved total payout EUR (typed by appr  ="Current: "&FIXED(SUM(tPayout[FinalPayoutEUR]),2)
D23   RUN STATUS                                =IF(D18>0,"HOLD",IF(AND(LEN(TRIM(p_ApproverName))>1,ISNUMBER(p_ApprovalDate)),IF(AND(ISNUMBER(p_ApprovedTotal),ABS(p_ApprovedTotal-SUM(tPayout[FinalPayoutEUR]))<0.005),"APPROVED","READY FOR REVIEW - approval stale or total missing"),"READY FOR REVIEW"))
D24   Approval warning                          =IF(AND(D18>0,OR(LEN(TRIM(p_ApproverName))>0,ISNUMBER(p_ApprovalDate))),"Approval entered while HOLD - ignored","")
```

## tPayout (sheet Calc_Payout)

```
Column3               =DATE(LEFT([@Period],4),RIGHT([@Period],2),1)
Column4               =LET(fromD,MAXIFS(tPlan[EffectiveFrom],tPlan[EffectiveFrom],"<="&[@PeriodStart]),toD,XLOOKUP(fromD,tPlan[EffectiveFrom],tPlan[EffectiveTo],0),IF(toD<[@PeriodStart],"None",XLOOKUP(fromD,tPlan[EffectiveFrom],tPlan[PlanVersion])))
Column5               =SUMIFS(tSales[EligibleEUR],tSales[EmployeeID],[@EmployeeID],tSales[Period],[@Period])
Column6               =XLOOKUP([@EmployeeID]&"|"&[@Period],tQuotas[Key],tQuotas[QuotaEUR],"")
Column7               =IF([@QuotaEUR]="","",[@EligibleSalesEUR]/[@QuotaEUR])
Column8               =IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],key,[@PlanVersion]&"|"&1,lo,XLOOKUP(key,tPlan[Key],tPlan[LoPct]),hi,XLOOKUP(key,tPlan[Key],tPlan[HiPct]),rate,XLOOKUP(key,tPlan[Key],tPlan[Rate]),ROUND(rate*MAX(0,MIN(sales,hi*quota)-lo*quota),2)))
Column9               =IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],key,[@PlanVersion]&"|"&2,lo,XLOOKUP(key,tPlan[Key],tPlan[LoPct]),hi,XLOOKUP(key,tPlan[Key],tPlan[HiPct]),rate,XLOOKUP(key,tPlan[Key],tPlan[Rate]),ROUND(rate*MAX(0,MIN(sales,hi*quota)-lo*quota),2)))
Column10              =IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],key,[@PlanVersion]&"|"&3,lo,XLOOKUP(key,tPlan[Key],tPlan[LoPct]),hi,XLOOKUP(key,tPlan[Key],tPlan[HiPct]),rate,XLOOKUP(key,tPlan[Key],tPlan[Rate]),ROUND(rate*MAX(0,MIN(sales,hi*quota)-lo*quota),2)))
Column11              =IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],ROUND(SUMPRODUCT((tPlan[PlanVersion]=[@PlanVersion])*tPlan[Rate]*(sales>tPlan[LoPct]*quota)*(sales-(sales>tPlan[HiPct]*quota)*(sales-tPlan[HiPct]*quota)-tPlan[LoPct]*quota)),2)))
Column12              =SUMIFS(tAdjustments[AmountEUR],tAdjustments[EmployeeID],[@EmployeeID],tAdjustments[Period],[@Period],tAdjustments[ApprovalStatus],p_Approved)
Column13              =SUMIFS(tAdjustments[AmountEUR],tAdjustments[EmployeeID],[@EmployeeID],tAdjustments[Period],[@Period],tAdjustments[ApprovalStatus],"<>"&p_Approved)
Column14              =COUNTIFS(tSales[EmployeeID],[@EmployeeID],tSales[Period],[@Period],tSales[Status],"Held")
Column15              =IF([@CalcPayoutEUR]="","",[@CalcPayoutEUR]+[@ApprovedAdjEUR])
Column16              =IFS([@QuotaEUR]="","HOLD - missing quota",[@PlanVersion]="None","HOLD - no plan",[@HeldTxnCount]>0,"REVIEW - held transactions",[@PendingAdjEUR]<>0,"REVIEW - pending adjustment",TRUE,"Calculated")
Column17              =[@EmployeeID]&"|"&[@Period]
Column18              =IFS([@QuotaEUR]="","Missing quota",[@PlanVersion]="None","No plan in force",[@PendingAdjEUR]<>0,"Pending adjustment",TRUE,"")
Column19              =IF([@ExceptionType]="","",XLOOKUP([@ExceptionType],tExceptionTypes[Type],tExceptionTypes[Severity],"High"))
Column20              =IF([@Severity]="","",XLOOKUP([@Key],tExceptions[RecordID],tExceptions[Status],"Open"))
Column21              =IFERROR(LOOKUP(2,1/((tAssignments[EmployeeID]=[@EmployeeID])*(tAssignments[StartDate]<=EOMONTH([@PeriodStart],0))),tAssignments[Region]),"Unassigned")
```

## Single cells on Calc_Payout

```
E4    RosterGap (valid sales with no tPayout r  =SUMPRODUCT((tSales[Status]="Valid")*(COUNTIFS(tPayout[EmployeeID],tSales[EmployeeID],tPayout[Period],tSales[Period])=0))
```

## tSales (sheet Sales_Raw)

```
Period                =MID([@SourceFile],7,7)
TxnMonth              =YEAR([@Date])&"-"&TEXT(MONTH([@Date]),"00")
IsDuplicate           =COUNTIF(INDEX(tSales[TransactionID],1):[@TransactionID],[@TransactionID])>1
MissingID             =[@EmployeeID]=""
KnownEmp              =COUNTIF(tRecipients[EmployeeID],[@EmployeeID])>0
ActiveOnDate          =LET(fromD,XLOOKUP([@EmployeeID],tRecipients[EmployeeID],tRecipients[ActiveFrom],""),toD,XLOOKUP([@EmployeeID],tRecipients[EmployeeID],tRecipients[ActiveTo],""),IF(fromD="",FALSE,AND([@Date]>=fromD,[@Date]<=toD)))
RegionOK              =COUNTIFS(tAssignments[EmployeeID],[@EmployeeID],tAssignments[Region],[@Region],tAssignments[StartDate],"<="&[@Date],tAssignments[EndDate],">="&[@Date])>0
FXRate                =XLOOKUP([@Currency]&"|"&[@TxnMonth],tFX[Key],tFX[RateToEUR],"")
AmountEUR             =IF([@FXRate]="","",ROUND([@AmountLocal]*[@FXRate],2))
IsCreditNote          =[@AmountLocal]<0
IsLate                =[@TxnMonth]<>[@Period]
Status                =IFS([@IsDuplicate],"Duplicate",[@MissingID],"Held",NOT([@KnownEmp]),"Held",NOT([@ActiveOnDate]),"Excluded",NOT([@RegionOK]),"Held",[@FXRate]="","Held",TRUE,"Valid")
Reason                =IFS([@IsDuplicate],"Duplicate TransactionID",[@MissingID],"Missing EmployeeID",NOT([@KnownEmp]),"Unknown employee",NOT([@ActiveOnDate]),"Inactive on transaction date",NOT([@RegionOK]),"Region does not match assignment",[@FXRate]="","Missing FX rate",[@IsLate],"Late prior-period item",[@IsCreditNote],"Credit note",TRUE,"")
EligibleEUR           =IF([@Status]="Valid",[@AmountEUR],0)
Severity              =IF([@Reason]="","",XLOOKUP([@Reason],tExceptionTypes[Type],tExceptionTypes[Severity],"High"))
WorkflowStatus        =IF([@Severity]="","",XLOOKUP([@TransactionID],tExceptions[RecordID],tExceptions[Status],"Open"))
```

## tExceptions (sheet Exceptions)

```
Column11              =IF([@Status]="Closed",[@ResolvedDate]-[@CreatedDate],p_AsOfDate-[@CreatedDate])
Column12              =IF([@AgeDays]<=p_AgeBucket1,"0-"&p_AgeBucket1,IF([@AgeDays]<=p_AgeBucket2,(p_AgeBucket1+1)&"-"&p_AgeBucket2,">"&p_AgeBucket2))
```

## tTests (sheet Tests)

```
Column7               =IFERROR(IF(ISNUMBER([@ExpectedPayout]),IF(ABS([@ActualPayout]-[@ExpectedPayout])<0.005,"PASS","FAIL"),IF([@ActualPayout]=[@ExpectedPayout],"PASS","FAIL")),"FAIL")
```

## Single cells on Tests

```
C3    Failed tests:                             =COUNTIFS(tTests[TestID],"<>",tTests[PassFail],"<>PASS")
```

## tSanity (sheet Sanity_Checks)

```
Column7               =IF([@QuotaEUR]="","",LET(sales,[@SalesEUR],quota,[@QuotaEUR],ROUND(SUMPRODUCT((tPlan[PlanVersion]=[@PlanVersion])*tPlan[Rate]*(sales>tPlan[LoPct]*quota)*(sales-(sales>tPlan[HiPct]*quota)*(sales-tPlan[HiPct]*quota)-tPlan[LoPct]*quota)),2)))
Column8               =[@EnginePayout]-[@ExpectedPayout]
Column9               =IF(ABS([@Diff])<0.005,"PASS","FAIL")
```

## Single cells on Sanity_Checks

```
C3    Failed sanity checks:                     =ROWS(tSanity[Result])-COUNTIF(tSanity[Result],"PASS")
```

## Single cells on Control_Waterfall

```
F5                                              =SUM(tRawCapture[RowsIn])
G5                                              =SUM(tRawCapture[AmountLocalIn])
F6    SourceFile                                =COUNTIF(tSales[Status],"Valid")
G6    SourceFile                                =SUMIFS(tSales[AmountLocal],tSales[Status],"Valid")
F7    Sales_2026-08.csv                         =COUNTIF(tSales[Status],"Held")
G7    Sales_2026-08.csv                         =SUMIFS(tSales[AmountLocal],tSales[Status],"Held")
F8    Sales_2026-09.csv                         =COUNTIF(tSales[Status],"Duplicate")
G8    Sales_2026-09.csv                         =SUMIFS(tSales[AmountLocal],tSales[Status],"Duplicate")
F9    Sales_2026-10.csv                         =COUNTIF(tSales[Status],"Excluded")
G9    Sales_2026-10.csv                         =SUMIFS(tSales[AmountLocal],tSales[Status],"Excluded")
F10                                             =SUM(F6:F9)
G10                                             =SUM(G6:G9)
F11                                             =F5-F10
G11                                             =G5-G10
```

## tCrossCut (sheet Control_CrossCut)

```
Column2               =SUMIFS(tSales[EligibleEUR],tSales[Region],[@Region],tSales[ProductFamily],tCrossCut[[#Headers],[AIS]])
Column3               =SUMIFS(tSales[EligibleEUR],tSales[Region],[@Region],tSales[ProductFamily],tCrossCut[[#Headers],[PDx]])
Column4               =SUMIFS(tSales[EligibleEUR],tSales[Region],[@Region],tSales[ProductFamily],tCrossCut[[#Headers],[PCS]])
Column5               =SUM([@AIS],[@PDx],[@PCS])
```

## Single cells on Control_CrossCut

```
C4    Region x family total                     =SUM(tCrossCut[RowTotal])
G4    Region x family total                     =SUM(tPayout[EligibleSalesEUR])
```

## tRecalc (sheet Control_Recalc)

```
Column3               =DATE(LEFT([@Period],4),RIGHT([@Period],2),1)
Column4               =YEAR(EDATE([@PeriodStart],-1))&"-"&TEXT(MONTH(EDATE([@PeriodStart],-1)),"00")
Column5               =YEAR(EDATE([@PeriodStart],1))&"-"&TEXT(MONTH(EDATE([@PeriodStart],1)),"00")
Column6               =SUMIFS(tSales[AmountEUR],tSales[EmployeeID],[@EmployeeID],tSales[Period],[@Period],tSales[Status],"Valid")
Column7               =IFERROR(INDEX(tQuotas[QuotaEUR],MATCH([@EmployeeID]&"|"&[@Period],tQuotas[Key],0)),"")
Column8               =INDEX(tPlan[PlanVersion],MATCH([@PeriodStart],tPlan[EffectiveFrom],1))
Column9               =INDEX(tPlan[LoPct],MATCH([@PlanVersion]&"|"&1,tPlan[Key],0))
Column10              =INDEX(tPlan[HiPct],MATCH([@PlanVersion]&"|"&1,tPlan[Key],0))
Column11              =INDEX(tPlan[Rate],MATCH([@PlanVersion]&"|"&1,tPlan[Key],0))
Column12              =IF([@QuotaEUR]="","",MAX(0,MIN([@SalesEUR],[@T1Hi]*[@QuotaEUR])-[@T1Lo]*[@QuotaEUR]))
Column13              =IF([@T1Band]="","",ROUND([@T1Rate]*[@T1Band],2))
Column14              =INDEX(tPlan[LoPct],MATCH([@PlanVersion]&"|"&2,tPlan[Key],0))
Column15              =INDEX(tPlan[HiPct],MATCH([@PlanVersion]&"|"&2,tPlan[Key],0))
Column16              =INDEX(tPlan[Rate],MATCH([@PlanVersion]&"|"&2,tPlan[Key],0))
Column17              =IF([@QuotaEUR]="","",MAX(0,MIN([@SalesEUR],[@T2Hi]*[@QuotaEUR])-[@T2Lo]*[@QuotaEUR]))
Column18              =IF([@T2Band]="","",ROUND([@T2Rate]*[@T2Band],2))
Column19              =INDEX(tPlan[LoPct],MATCH([@PlanVersion]&"|"&3,tPlan[Key],0))
Column20              =INDEX(tPlan[HiPct],MATCH([@PlanVersion]&"|"&3,tPlan[Key],0))
Column21              =INDEX(tPlan[Rate],MATCH([@PlanVersion]&"|"&3,tPlan[Key],0))
Column22              =IF([@QuotaEUR]="","",MAX(0,MIN([@SalesEUR],[@T3Hi]*[@QuotaEUR])-[@T3Lo]*[@QuotaEUR]))
Column23              =IF([@T3Band]="","",ROUND([@T3Rate]*[@T3Band],2))
Column24              =IF([@QuotaEUR]="","",[@T1Pay]+[@T2Pay]+[@T3Pay])
Column25              =INDEX(tPayout[CalcPayoutEUR],MATCH([@EmployeeID]&"|"&[@Period],tPayout[Key],0))
Column26              =IF(AND([@RecalcPayout]="",[@PrimaryPayout]=""),0,IF(OR([@RecalcPayout]="",[@PrimaryPayout]=""),"MISMATCH",ROUND([@RecalcPayout]-[@PrimaryPayout],2)))
Column27              =COUNTIFS(tRecalc[EmployeeID],[@EmployeeID],tRecalc[Period],[@PriorPeriod])>0
Column28              =COUNTIFS(tRecalc[EmployeeID],[@EmployeeID],tRecalc[Period],[@NextPeriod])>0
```

## tRollForward (sheet Control_RollForward)

```
Column2               =YEAR(EDATE(DATE(LEFT([@Period],4),RIGHT([@Period],2),1),-1))&"-"&TEXT(MONTH(EDATE(DATE(LEFT([@Period],4),RIGHT([@Period],2),1),-1)),"00")
Column3               =SUMIFS(tPayout[FinalPayoutEUR],tPayout[Period],[@PriorPeriod])
Column4               =SUMIFS(tRecalc[RecalcPayout],tRecalc[Period],[@Period],tRecalc[InPrior],TRUE)-SUMIFS(tRecalc[RecalcPayout],tRecalc[Period],[@PriorPeriod],tRecalc[InNext],TRUE)
Column5               =SUMIFS(tRecalc[RecalcPayout],tRecalc[Period],[@Period],tRecalc[InPrior],FALSE)
Column6               =-SUMIFS(tRecalc[RecalcPayout],tRecalc[Period],[@PriorPeriod],tRecalc[InNext],FALSE)
Column7               =SUMIFS(tAdjustments[AmountEUR],tAdjustments[Period],[@Period],tAdjustments[ApprovalStatus],p_Approved)-SUMIFS(tAdjustments[AmountEUR],tAdjustments[Period],[@PriorPeriod],tAdjustments[ApprovalStatus],p_Approved)
Column8               =[@OpeningExpense]+[@ContinuingChange]+[@NewRecipients]+[@Leavers]+[@AdjustmentChange]
Column9               =SUMIFS(tPayout[FinalPayoutEUR],tPayout[Period],[@Period])
Column10              =ROUND([@ExpectedClosing]-[@ActualClosing],2)
Column11              =SUMIFS(tSales[EligibleEUR],tSales[Period],[@Period],tSales[IsLate],TRUE)
```

## tStatus (sheet Status_Export)

```
Column1               =p_ReportPeriod
Column2               =p_RunStatus
Column3               =p_ChecksFailing
Column4               =p_OpenHigh
Column5               =SUM(tPayout[FinalPayoutEUR])
```

## Single cells on AI_Prompt

```
B4    Report period                             =p_ReportPeriod
B5    Prior period                              =YEAR(EDATE(DATE(LEFT(B4,4),RIGHT(B4,2),1),-1))&"-"&TEXT(MONTH(EDATE(DATE(LEFT(B4,4),RIGHT(B4,2),1),-1)),"00")
B6    Total incentive EUR                       =SUMIFS(tPayout[FinalPayoutEUR],tPayout[Period],B4)
B7    Prior total incentive EUR                 =SUMIFS(tPayout[FinalPayoutEUR],tPayout[Period],B5)
B8    MoM change                                =IF(B7=0,"n/a",B6/B7-1)
B9    Eligible sales EUR                        =SUMIFS(tPayout[EligibleSalesEUR],tPayout[Period],B4)
B10   Quota EUR (rows with quota)               =SUMIFS(tPayout[QuotaEUR],tPayout[Period],B4)
B11   Weighted attainment                       =SUMIFS(tPayout[EligibleSalesEUR],tPayout[Period],B4,tPayout[QuotaEUR],">0")/B10
B12   Incentive % of eligible sales             =B6/B9
B13   Payout rows on HOLD                       =COUNTIFS(tPayout[Period],B4,tPayout[ReviewStatus],"HOLD*")
B14   Held or excluded transactions             =COUNTIFS(tSales[Period],B4,tSales[Status],"Held")+COUNTIFS(tSales[Period],B4,tSales[Status],"Excluded")
B15   Open High exceptions (all periods)        =p_OpenHigh
B16   Run status                                =p_RunStatus
A19                                             =TEXTJOIN(CHAR(10),TRUE,"SYNTHETIC DATA. Facts below are verified by Excel controls. Use only these facts.","Report period: "&B4,"Prior period: "&B5,"Total incentive EUR: "&FIXED(B6,2),"Prior total incentive EUR: "&FIXED(B7,2),"Month-on-month change: "&IF(ISNUMBER(B8),FIXED(B8*100,1)&"%","n/a"),"Eligible sales EUR: "&FIXED(B9,0),"Quota EUR (rows with quota): "&FIXED(B10,0),"Weighted attainment: "&FIXED(B11*100,1)&"%","Incentive as % of eligible sales: "&FIXED(B12*100,2)&"%","Payout rows on HOLD: "&B13,"Held or excluded transactions: "&B14,"Open High exceptions: "&B15,"Run status: "&B16)
```

## tClaims (sheet Claim_Verification)

```
Column6               =IF([@ValueInNarrative]="","Pending",IF(ISNUMBER([@SourceValue]),IF(ABS([@ValueInNarrative]-[@SourceValue])<=ABS([@SourceValue])*0.0005+0.005,"Y","N"),IF(EXACT([@ValueInNarrative],[@SourceValue]),"Y","N")))
```

## Single cells on Claim_Verification

```
C4    PUBLICATION STATUS                        =IF(COUNTIF(tClaims[Match],"Y")=ROWS(tClaims[Match]),"CLEARED for human review","BLOCKED - "&(ROWS(tClaims[Match])-COUNTIF(tClaims[Match],"Y"))&" claim(s) not verified")
```

## tAIEval (sheet AI_Eval)

```
Column8               =COUNTIF(tAIEval[@[Run1]:[Run3]],"Pass")
Column9               =COUNTIF(tAIEval[@[Run1]:[Run3]],"Pass")+COUNTIF(tAIEval[@[Run1]:[Run3]],"Fail")
Column10              =IF([@RunsDone]=0,"",[@Passes]/[@RunsDone])
```

## Single cells on AI_Eval

```
C3    Overall pass rate:                        =IF(SUM(tAIEval[RunsDone])=0,"not yet measured",SUM(tAIEval[Passes])/SUM(tAIEval[RunsDone]))
```

## tNotice (sheet Stakeholder_Pack)

```
Column2               =COUNTIFS(tSales[Period],p_ReportPeriod,tSales[Reason],[@Reason],tSales[WorkflowStatus],"<>Closed")
Column3               =SUMIFS(tSales[AmountLocal],tSales[Period],p_ReportPeriod,tSales[Reason],[@Reason],tSales[WorkflowStatus],"<>Closed")
```

## Single cells on Stakeholder_Pack

```
A4                                              ="Subject: SIC "&p_ReportPeriod&" - "&SUM(tNotice[OpenCount])&" transaction(s) need Commercial confirmation by <date>"
A5                                              ="Hello team, the "&p_ReportPeriod&" incentive run is on "&p_RunStatus&". The items below cannot be credited until you confirm. Reply with the corrected value or confirmation; we do not change source data on your behalf."
C15   Report period                             =p_ReportPeriod
C16   Incentive expense EUR                     =SUMIFS(tPayout[FinalPayoutEUR],tPayout[Period],p_ReportPeriod)
C17   Opening (prior period) EUR                =INDEX(tRollForward[OpeningExpense],MATCH(p_ReportPeriod,tRollForward[Period],0))
C18   Continuing recipients change              =INDEX(tRollForward[ContinuingChange],MATCH(p_ReportPeriod,tRollForward[Period],0))
C19   New recipients                            =INDEX(tRollForward[NewRecipients],MATCH(p_ReportPeriod,tRollForward[Period],0))
C20   Leavers                                   =INDEX(tRollForward[Leavers],MATCH(p_ReportPeriod,tRollForward[Period],0))
C21   Approved adjustments change               =INDEX(tRollForward[AdjustmentChange],MATCH(p_ReportPeriod,tRollForward[Period],0))
C22   Roll-forward variance                     =INDEX(tRollForward[Variance],MATCH(p_ReportPeriod,tRollForward[Period],0))
C23   Pending (unapproved) adjustments EUR      =SUMIFS(tPayout[PendingAdjEUR],tPayout[Period],p_ReportPeriod)
C24   Payout rows on HOLD (not accrued)         =COUNTIFS(tPayout[Period],p_ReportPeriod,tPayout[ReviewStatus],"HOLD*")
C25   Checks not passing                        =p_ChecksFailing
C26   Run status                                =p_RunStatus
C29   New hires                                 =COUNTIF(tRecipientChanges[ChangeType],"New hire")
C30   Leavers                                   =COUNTIF(tRecipientChanges[ChangeType],"Leaver")
C31   Territory changes                         =COUNTIF(tRecipientChanges[ChangeType],"Territory change")
C32   Sales not paid for changed recipients EU  =SUM(tRecipientChanges[SalesNotPaidEUR])
```

## Single cells on Run_Log

```
B4    Files loaded                              =ROWS(tRawCapture[SourceFile])
B5    Rows in                                   =SUM(tRawCapture[RowsIn])
B6    Rows held                                 =COUNTIF(tSales[Status],"Held")
B7    Waterfall row variance                    =Control_Waterfall!F11
B8    Checks not passing                        =p_ChecksFailing
B9    Run status                                =p_RunStatus
```

## Single cells on Timing

```
B9    Median                                    =IF(COUNT(tTiming[ManualMachineSec])=0,"not measured",MEDIAN(tTiming[ManualMachineSec]))
C9    Median                                    =IF(COUNT(tTiming[ManualReviewMin])=0,"not measured",MEDIAN(tTiming[ManualReviewMin]))
D9    Median                                    =IF(COUNT(tTiming[RefreshMachineSec])=0,"not measured",MEDIAN(tTiming[RefreshMachineSec]))
E9    Median                                    =IF(COUNT(tTiming[RefreshReviewMin])=0,"not measured",MEDIAN(tTiming[RefreshReviewMin]))
```

## tFX (sheet FX)

```
Key                   =[@Currency]&"|"&[@Period]
```

## tQuotas (sheet Quotas)

```
Key                   =[@EmployeeID]&"|"&[@Period]
```

## tPlan (sheet Plan)

```
Key                   =[@PlanVersion]&"|"&[@Tier]
```

## tRecipientChanges (sheet Recipient_Changes)

```
ChangePeriod          =YEAR([@EffectiveDate])&"-"&TEXT(MONTH([@EffectiveDate]),"00")
PayoutInChangePeriod  =IFERROR(INDEX(tPayout[FinalPayoutEUR],MATCH([@EmployeeID]&"|"&[@ChangePeriod],tPayout[Key],0)),"No payout row")
SalesNotPaidEUR       =SUMIFS(tSales[AmountEUR],tSales[EmployeeID],[@EmployeeID],tSales[Status],"Held")+SUMIFS(tSales[AmountEUR],tSales[EmployeeID],[@EmployeeID],tSales[Status],"Excluded")
```

## tTests ActualPayout (one formula per test)

```
TS01  =INDEX(tSanity[EnginePayout],MATCH("BD1",tSanity[CaseID],0))
TS02  =INDEX(tSanity[EnginePayout],MATCH("BD2",tSanity[CaseID],0))
TS03  =INDEX(tSanity[EnginePayout],MATCH("BD3",tSanity[CaseID],0))
TS04  =INDEX(tSanity[EnginePayout],MATCH("BD4",tSanity[CaseID],0))
TS05  =INDEX(tSanity[EnginePayout],MATCH("BD5",tSanity[CaseID],0))
TS06  =INDEX(tPayout[ReviewStatus],MATCH("E106|2026-10",tPayout[Key],0))
TS07  =INDEX(tSales[Status],MATCH("T1037",tSales[TransactionID],0))
TS08  =SUMIFS(tSales[EligibleEUR],tSales[TransactionID],"T1043")
TS09  =INDEX(tSales[Status],MATCH("T1015",tSales[TransactionID],0))
TS10  =INDEX(tSales[Status],MATCH("T1009",tSales[TransactionID],0))
TS11  =INDEX(tPayout[CalcPayoutEUR],MATCH("E102|2026-09",tPayout[Key],0))
TS12  =INDEX(tSales[AmountEUR],MATCH("T1032",tSales[TransactionID],0))
TS13  =INDEX(tPayout[CalcPayoutEUR],MATCH("E105|2026-10",tPayout[Key],0))
TS14  =INDEX(tPayout[CalcPayoutEUR],MATCH("E101|2026-10",tPayout[Key],0))
TS15  =INDEX(tPayout[FinalPayoutEUR],MATCH("E101|2026-10",tPayout[Key],0))
TS16  =INDEX(tPayout[CalcPayoutEUR],MATCH("E104|2026-09",tPayout[Key],0))
TS17  =INDEX(tPayout[EligibleSalesEUR],MATCH("E101|2026-10",tPayout[Key],0))
TS18  =INDEX(tPayout[FinalPayoutEUR],MATCH("E105|2026-09",tPayout[Key],0))
```

## tControls Value (one formula per control)

```
C1a   =Control_Waterfall!F11
C1b   =Control_Waterfall!G11
C2a   =ROUND(Control_CrossCut!C4-Control_CrossCut!G4,2)
C2b   =ROUND(SUM(tSales[EligibleEUR])-Control_CrossCut!C4,2)
C3a   =ROWS(tRecalc[Variance])-COUNTIF(tRecalc[Variance],0)
C3b   =ROWS(tRecalc[EmployeeID])-ROWS(tPayout[EmployeeID])
C4    =SUMPRODUCT(ABS(tRollForward[Variance]))
C5    =Calc_Payout!E4
C6    =Sanity_Checks!C3
C7    =COUNTIFS(tTests[TestID],"<>",tTests[PassFail],"<>PASS")
C8    =COUNTIFS(tSales[Severity],"High",tSales[WorkflowStatus],"<>Closed")+COUNTIFS(tPayout[Severity],"High",tPayout[WorkflowStatus],"<>Closed")
C9    =COUNTIFS(tExceptions[Status],"Closed",tExceptions[Resolution],"")
```

## Named ranges

| Name | Refers to |
|---|---|
| p_Approved | Parameters!$B$4 |
| p_AsOfDate | Parameters!$B$5 |
| p_AgeBucket1 | Parameters!$B$6 |
| p_AgeBucket2 | Parameters!$B$7 |
| p_ReportPeriod | Parameters!$B$8 |
| p_OpenHigh | Controls!$D$15 |
| p_ApprovedTotal | Controls!$D$22 |
| p_ApproverName | Controls!$D$20 |
| p_ApprovalDate | Controls!$D$21 |
| p_RunStatus | Controls!$D$23 |
| p_ChecksFailing | Controls!$D$18 |
