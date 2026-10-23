# Phase 8: Interview kit

Framing to use every time: *"A synthetic prototype inspired by the SIC role. It doesn't reflect GE HealthCare's plans, systems or data."*

## 1. Job description → project (also in JD_Mapping.pptx)

| JD theme | Where it is in the project | Evidence to show |
|---|---|---|
| End-to-end SIC calculation | Effective-dated plan, FX, eligibility, marginal tiers, approved adjustments (Calc_Payout) | E105 Aug: 250k vs 200k = 2,800 |
| Data quality | 11 planted errors, validation columns, exception workflow with severity and ageing | Sales_Raw Status/Reason; Exceptions |
| Controls / accuracy | 4 independent reconciliations, 18 tests, sanity boundaries, gate tied to approval | Controls sheet: HOLD with 1 failing check |
| Power BI dashboards | 3 pages, star schema, weighted attainment, RLS | 03_Power_BI.md |
| AI exploration | Aggregate prompt, guarded system prompts, claim verification, eval harness | AI_Prompt, Claim_Verification, AI_Eval |
| Process improvement / automation | Folder import, schema check, run log, timing protocol | Drop-a-file demo |
| Large datasets | Logic is table-driven and row-count independent; the path to SQL/Python is in section 2 | Waterfall stays 0 when rows grow from 59 to 73 |
| Stakeholders | Commercial notice, Finance roll-forward summary, HR change report, SOP | Stakeholder_Pack, SOP |

## 2. Limitations and next steps

- **Plan complexity.** Real plans have product weightings, caps, accelerators, multiple measures, draws and SPIFFs. tPlan would need more columns and a measure dimension.
- **Proration.** E108 left mid-September but is measured against a full monthly quota. Next step: prorate quota by active days.
- **Territory split.** A mid-month change is credited by transaction date (correct), but the payout row's region is the region at period end. Splitting it would need proration.
- **Clawbacks.** Credit notes net in the booking month. A real policy might link back to the original sale, recover across periods, and set a floor on negative payouts.
- **Multi-currency policy.** Monthly average rates by transaction month. A real policy might use plan-rate FX or budget rates to stop currency moves affecting pay.
- **Timing of crediting.** Orders vs shipment vs revenue recognition is a configurable choice. This prototype credits by booking date. *Ask how they do it.*
- **Scale.** Excel is fine for thousands of rows. Next, move validation and aggregation to SQL or Python with the same rules and tests, and keep Excel and Power BI for review.
- **Automation.** Logging is manual (no macros). Power Automate notifications and approvals are LICENCE-DEPENDENT.

## 3. Five-minute demo script

| Time | What you show | What you say |
|---|---|---|
| 0:00 | README → Controls | "The problem: pay people correctly, on time, and be able to prove it. Synthetic data, inspired by the role." |
| 0:40 | Calc_Payout E105 Aug, then the formula bar | "Marginal tiers: each band pays only on sales inside it. 250k vs 200k: 800 + 2,000 = 2,800. Sanity_Checks proves the 80% and 100% boundaries." |
| 1:30 | Sales_Raw T1035 → Exceptions EX004 | "E104 moved East to West on 16 September. This East sale on the 25th is held. If Sales Ops confirms it, E104 passes 100% and the payout changes, so it needs a human decision." |
| 2:15 | Move Sales_2026-11.csv into data\ → Refresh All | "The waterfall stays 0, so nothing was lost. All 14 rows are held: there's no November FX rate. I add EUR = 1 and RosterGap shows 14: there are no payout rows or quotas yet. The controls stop the payment." |
| 3:15 | Power BI: Executive → Controls, View as Regional Manager | "Weighted attainment, a status banner from the same gate, and row-level security so East sees only East." |
| 4:00 | AI_Prompt A19 → AI draft → Claim_Verification | "Only aggregates go in. Every number is checked against its source cell, and one mismatch blocks publication." |
| 4:40 | Controls: approver name while HOLD → warning | "Excel can't approve itself. Approval needs a name, a date and the approved total. Change any number afterwards and the approval goes stale." |

**Fallbacks**

- **Refresh fails:** show rehearsal screenshots, then explain the waterfall from tRawCapture.
- **AI tool unavailable:** use a saved draft and run the claim check live.
- **Power BI fails:** show the Controls sheet and the expected-values table.

## 4. Fifteen likely questions with model answers

1. **How do you know the calculation is correct?** Three layers. Hand-derived tests (18, all passing) and boundary sanity checks. A second, independent calculation using different formulas that must match on every row. Reconciliations that tie the source files to the payout and the month to the prior month. Plus a gate that won't approve if any of those fail.
2. **What happens when territory changes mid-month?** Assignments are effective-dated. Each transaction is checked against the region valid on its own date. E104's sales on the 8th (East) and the 22nd (West) are valid; an East-tagged sale on the 25th is held for Sales Ops to confirm.
3. **Why marginal tiers?** No cliff. One more euro never changes the payout by more than one euro times the current rate. That's fairer, harder to game, and easier to explain. The 80.00% and 80.01% tests show the smooth start.
4. **How would you investigate a sudden jump in incentive expense?** Start with the roll-forward: split the change into continuing recipients, new hires, leavers and adjustments. Then check plan-version changes, late items, FX and big individual transactions. October's +46% breaks down as +1,100 from continuing recipients plus +550 from adjustments. The main driver is E101 (+2,640, at 148% under the new 5% top rate), partly offset by E105 (−1,280) and E106's payout on HOLD (−960).
5. **How would you use AI safely with compensation data?** Approved tools only. Aggregates in, never personal data. The prompt forbids calculating or approving. Every claim is checked against a source cell. A human approves. Behaviour is measured with an evaluation set, not assumed.
6. **What would you automate first on a real team?** The data intake and validation step: folder import, schema check, standard flags and the exception log. It's repetitive, error-prone and comes before every calculation. Measure the time before and after.
7. **Finance and Commercial disagree on a sales figure. What now?** Go back to the source record and the policy definition (orders vs revenue, booking date). Show the reconciliation trail. Don't change data on either side's behalf; hold the item, record the decision and the approver, and fix the rule if it was ambiguous.
8. **A duplicate appears across two months. What happens?** IsDuplicate catches the repeat, so it pays once. If the first was already paid, it's a recovery case: log it and agree an adjustment, never silently delete.
9. **Why HOLD rather than paying the valid part?** Policy choice. The gate is conservative for High exceptions. In practice you might pay unaffected recipients and hold only the affected rows. Statuses are already per row, so that would be a change to the gate rule, not a rebuild.
10. **How do you handle a missing quota?** That payout row is HOLD and can't be calculated. It's a High exception owned by Sales Ops. The sale stays valid, so nothing is lost when the quota arrives.
11. **What's the difference between your two payout calculations?** The primary uses LET + SUMPRODUCT across all tiers. The check uses INDEX/MATCH and MIN/MAX per tier in separate columns, its own roster, and re-sums sales from AmountEUR. Agreement means two methods reached the same answer, not one formula checking itself.
12. **How would this scale to 50 countries?** Same tables, more rows. Currency and plan versions are already data. I'd move transformation to SQL or Python, keep the test suite as regression tests, and use Power BI with RLS for distribution.
13. **What could still be wrong?** The waterfall can't see rows lost before capture. The cross-cut shares EligibleEUR with the employee view. Proration and clawbacks aren't modelled. Some tests use the sanity table rather than live rows. I'd say so up front.
14. **How do you document a monthly cycle?** The SOP sheet: ten steps with owner, control and evidence. Plus the Change Log, the Run Log and the Controls Matrix mapping each risk to a control and a test.
15. **Tell me about a mistake you caught.** My first sanity counter used COUNTIF("FAIL"). Errors aren't counted as FAIL, so a broken formula showed 0 failures, a false green. I changed every counter to "anything not PASS" and logged it in the Change Log.

## 5. Questions to ask them

- How did this year's sales reorganisation change territories or plans?
- Do you credit on orders, shipment or revenue recognition?
- Which tools does the team calculate in today, and where is the manual effort?
- Which AI tools are approved for finance data, and what review applies?
