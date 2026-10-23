# SOP: monthly SIC cycle (synthetic, illustrative)

| # | When | Owner | Activity | Control | Evidence |
|---|---|---|---|---|---|
| 1 | Day 1 | SIC Analyst | Drop the month's files into /data, then Refresh All | Schema check + raw capture | Schema_Check, tRawCapture, Run_Log |
| 2 | Day 1 | SIC Analyst | Confirm FX rates, quotas, plan version and roster for the period | RosterGap = 0; plan version found | Calc_Payout E4 |
| 3 | Day 1 | HR / Sales Ops | Send recipient and territory changes | Recipient Change Log has an approver | tRecipientChanges |
| 4 | Day 2 | SIC Analyst | Review validation flags; assign an owner to each exception | Exception workflow | tExceptions |
| 5 | Day 2–4 | Commercial / Finance | Resolve exceptions (IDs, regions, FX) | Closed requires a Resolution | tExceptions, Controls C9 |
| 6 | Day 4 | SIC Analyst | Re-run; check C1–C4 and the tests | Independent controls | Controls sheet |
| 7 | Day 4 | SIC Analyst | Draft commentary with AI from the aggregate prompt; verify every claim | Claim verification | Claim_Verification |
| 8 | Day 5 | SIC Lead | Review and approve (name, date, approved total) | Gate: APPROVED only with human sign-off | Controls D20:D23 |
| 9 | Day 5 | SIC Analyst | Send the Finance summary and Commercial notice; publish the dashboard | Stakeholder pack | Stakeholder_Pack, Power BI |
| 10 | Day 5 | SIC Analyst | Log changes and the run; archive the files | Change Log + Run Log | Change_Log, Run_Log |

The same table is in the workbook on the SOP sheet (tSOP).
