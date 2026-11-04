import csv
import os
import random
from datetime import date
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

random.seed(2026)
OUT = "/tmp/sic/SIC_Intelligence_Hub"
DATA = os.path.join(OUT, "data")
REF = os.path.join(OUT, "reference")
DEMO = os.path.join(OUT, "demo_hold_back")
for d in (OUT, DATA, REF, DEMO):
    os.makedirs(d, exist_ok=True)

LABEL = "SYNTHETIC DATA - fictional prototype inspired by the SIC role. Not GE HealthCare plans, systems, people or data."

recipients = [
    ("E101", "Account Manager", "MGR-01", date(2024, 1, 1), date(2099, 12, 31)),
    ("E102", "Account Manager", "MGR-01", date(2024, 1, 1), date(2099, 12, 31)),
    ("E103", "Senior Account Manager", "MGR-01", date(2023, 6, 1), date(2099, 12, 31)),
    ("E104", "Account Manager", "MGR-02", date(2025, 3, 1), date(2099, 12, 31)),
    ("E105", "Senior Account Manager", "MGR-02", date(2022, 9, 1), date(2099, 12, 31)),
    ("E106", "Account Manager", "MGR-01", date(2025, 1, 1), date(2099, 12, 31)),
    ("E107", "Account Manager", "MGR-02", date(2026, 9, 1), date(2099, 12, 31)),
    ("E108", "Account Manager", "MGR-02", date(2024, 4, 1), date(2026, 9, 15)),
]

assignments = [
    ("E101", "North", date(2024, 1, 1), date(2099, 12, 31)),
    ("E102", "South", date(2024, 1, 1), date(2099, 12, 31)),
    ("E103", "North", date(2023, 6, 1), date(2099, 12, 31)),
    ("E104", "East", date(2025, 3, 1), date(2026, 9, 15)),
    ("E104", "West", date(2026, 9, 16), date(2099, 12, 31)),
    ("E105", "West", date(2022, 9, 1), date(2099, 12, 31)),
    ("E106", "South", date(2025, 1, 1), date(2099, 12, 31)),
    ("E107", "East", date(2026, 9, 1), date(2099, 12, 31)),
    ("E108", "East", date(2024, 4, 1), date(2026, 9, 15)),
]

periods = ["2026-08", "2026-09", "2026-10"]
quota_base = {"E101": 100000, "E102": 120000, "E103": 150000, "E104": 100000,
              "E105": 200000, "E106": 80000, "E107": 60000, "E108": 90000}
target_att = {
    "2026-08": {"E101": 0.70, "E102": 0.95, "E103": 1.10, "E104": 0.85, "E105": 1.25, "E106": 0.90, "E108": 0.75},
    "2026-09": {"E101": 0.88, "E102": 1.05, "E103": 0.78, "E104": 1.00, "E105": 1.15, "E106": 1.20, "E107": 0.60, "E108": 0.50},
    "2026-10": {"E101": 1.30, "E102": 0.82, "E103": 1.05, "E104": 0.92, "E105": 0.98, "E106": 1.10, "E107": 0.95},
}
quotas = []
for p in periods:
    for e, q in quota_base.items():
        if e == "E107" and p == "2026-08":
            continue
        if e == "E108" and p == "2026-10":
            continue
        if e == "E106" and p == "2026-10":
            continue
        quotas.append((e, p, q))

region_of = {"E101": "North", "E102": "South", "E103": "North", "E105": "West",
             "E106": "South", "E107": "East", "E108": "East"}
families = ["AIS", "PDx", "PCS"]
fam_w = [0.6, 0.25, 0.15]


def last_day(p):
    y, m = map(int, p.split("-"))
    return 31 if m in (8, 10) else 30


rows = []


def add(d, e, reg, fam, cur, amt, src):
    rows.append({"Date": d, "EmployeeID": e, "Region": reg, "ProductFamily": fam,
                 "Currency": cur, "AmountLocal": amt, "SourceFile": src})


for p in periods:
    y, m = map(int, p.split("-"))
    src = f"Sales_{p}.csv"
    for e, att in target_att[p].items():
        total = round(quota_base[e] * att / 1000) * 1000
        n = 3 if total >= 150000 else 2
        cuts = sorted(random.sample(range(1, int(total / 1000)), n - 1))
        parts = [b - a for a, b in zip([0] + cuts, cuts + [int(total / 1000)])]
        for k, part in enumerate(parts):
            amt = part * 1000
            if e == "E108" and p == "2026-09":
                d = date(y, m, [4, 11, 14][k])
            elif e == "E104" and p == "2026-09":
                d = date(y, m, [8, 22, 27][k])
            else:
                d = date(y, m, random.randint(1, last_day(p)))
            if e == "E104":
                reg = "East" if d <= date(2026, 9, 15) else "West"
            else:
                reg = region_of[e]
            add(d, e, reg, random.choices(families, fam_w)[0], "EUR", amt, src)

planted = []

hu_idx = next(i for i, r in enumerate(rows) if r["EmployeeID"] == "E102" and r["SourceFile"] == "Sales_2026-09.csv")
hu_eur = rows[hu_idx]["AmountLocal"]
rows[hu_idx]["Currency"] = "HUF"
rows[hu_idx]["AmountLocal"] = int(hu_eur / 0.0025)
rows[hu_idx]["_tag"] = "HUF"

add(date(2026, 9, 18), "E102", "South", "AIS", "EUR", -12000, "Sales_2026-09.csv"); rows[-1]["_tag"] = "CREDIT"
add(date(2026, 8, 20), "", "North", "PDx", "EUR", 9000, "Sales_2026-08.csv"); rows[-1]["_tag"] = "BLANKID"
add(date(2026, 8, 26), "E199", "West", "AIS", "EUR", 15000, "Sales_2026-08.csv"); rows[-1]["_tag"] = "UNKNOWN"
add(date(2026, 9, 25), "E108", "East", "AIS", "EUR", 22000, "Sales_2026-09.csv"); rows[-1]["_tag"] = "INACTIVE"
add(date(2026, 9, 25), "E104", "East", "PDx", "EUR", 14000, "Sales_2026-09.csv"); rows[-1]["_tag"] = "TERRITORY"
add(date(2026, 10, 9), "E103", "South", "PCS", "EUR", 11000, "Sales_2026-10.csv"); rows[-1]["_tag"] = "MISMATCH"
add(date(2026, 10, 14), "E105", "West", "AIS", "HUF", 8000000, "Sales_2026-10.csv"); rows[-1]["_tag"] = "MISSINGFX"
add(date(2026, 9, 29), "E101", "North", "AIS", "EUR", 18000, "Sales_2026-10.csv"); rows[-1]["_tag"] = "LATE"
add(date(2026, 10, 21), "E106", "South", "PDx", "EUR", 10000, "Sales_2026-10.csv"); rows[-1]["_tag"] = "NOQUOTA"

order = {"Sales_2026-08.csv": 0, "Sales_2026-09.csv": 1, "Sales_2026-10.csv": 2}
rows.sort(key=lambda r: (order[r["SourceFile"]], r["Date"], r["EmployeeID"]))
for i, r in enumerate(rows):
    r["TransactionID"] = f"T{1001 + i}"

dup_src = next(r for r in rows if r["SourceFile"] == "Sales_2026-10.csv" and r["EmployeeID"] == "E107")
dup = dict(dup_src); dup["_tag"] = "DUPLICATE"
pos = rows.index(dup_src) + 1
rows.insert(pos, dup)

fx = [("EUR", p, 1.0) for p in periods] + [("HUF", "2026-08", 0.00256), ("HUF", "2026-09", 0.0025)]

plan = [
    ("v1", date(2026, 1, 1), date(2026, 9, 30), 1, 0.0, 0.8, 0.00),
    ("v1", date(2026, 1, 1), date(2026, 9, 30), 2, 0.8, 1.0, 0.02),
    ("v1", date(2026, 1, 1), date(2026, 9, 30), 3, 1.0, 999.0, 0.04),
    ("v2", date(2026, 10, 1), date(2099, 12, 31), 1, 0.0, 0.8, 0.00),
    ("v2", date(2026, 10, 1), date(2099, 12, 31), 2, 0.8, 1.0, 0.02),
    ("v2", date(2026, 10, 1), date(2099, 12, 31), 3, 1.0, 999.0, 0.05),
]

adjustments = [
    ("ADJ01", "E103", "2026-08", 500, "Q2 SPIFF true-up", "Approved", "SIC Lead (synthetic)"),
    ("ADJ02", "E105", "2026-09", -300, "Recovery of prior overpayment", "Approved", "SIC Lead (synthetic)"),
    ("ADJ03", "E101", "2026-10", 750, "Manager discretionary award", "Pending", ""),
    ("ADJ04", "E104", "2026-10", 250, "Territory transition support", "Approved", "SIC Lead (synthetic)"),
]

recipient_changes = [
    ("RC01", "E107", date(2026, 9, 1), "New hire", "", "East / MGR-02", "HR Ops (synthetic)"),
    ("RC02", "E104", date(2026, 9, 16), "Territory change", "East", "West", "Sales Ops (synthetic)"),
    ("RC03", "E108", date(2026, 9, 15), "Leaver", "Active", "Inactive", "HR Ops (synthetic)"),
]

exceptions_hdr = ["ExceptionID", "Period", "RecordID", "Type", "Severity", "Owner", "Status",
                  "CreatedDate", "ResolvedDate", "Resolution"]
tests_hdr = ["TestID", "Scenario", "Input", "ExpectedPayout", "ActualPayout", "PassFail"]
sales_hdr = ["TransactionID", "Date", "EmployeeID", "Region", "ProductFamily", "Currency", "AmountLocal", "SourceFile"]

for p in periods:
    with open(os.path.join(DATA, f"Sales_{p}.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(sales_hdr[:-1])
        for r in rows:
            if r["SourceFile"] == f"Sales_{p}.csv":
                w.writerow([r["TransactionID"], r["Date"].isoformat(), r["EmployeeID"], r["Region"],
                            r["ProductFamily"], r["Currency"], r["AmountLocal"]])

nov = []
nid = 1001 + len(rows) + 50
for e, reg in [("E101", "North"), ("E102", "South"), ("E103", "North"), ("E104", "West"),
               ("E105", "West"), ("E106", "South"), ("E107", "East")]:
    for k in range(2):
        nov.append([f"T{nid}", date(2026, 11, random.randint(2, 27)).isoformat(), e, reg,
                    random.choices(families, fam_w)[0], "EUR", random.randint(10, 60) * 1000])
        nid += 1
nov.sort(key=lambda x: x[1])
with open(os.path.join(DEMO, "Sales_2026-11.csv"), "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(sales_hdr[:-1])
    w.writerows(nov)


def write_csv(name, hdr, data):
    with open(os.path.join(REF, name), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(hdr)
        for row in data:
            w.writerow([v.isoformat() if isinstance(v, date) else v for v in row])


write_csv("FX.csv", ["Currency", "Period", "RateToEUR"], fx)
write_csv("Recipients.csv", ["EmployeeID", "Role", "ManagerID", "ActiveFrom", "ActiveTo"], recipients)
write_csv("Assignments.csv", ["EmployeeID", "Region", "StartDate", "EndDate"], assignments)
write_csv("Quotas.csv", ["EmployeeID", "Period", "QuotaEUR"], quotas)
write_csv("Plan.csv", ["PlanVersion", "EffectiveFrom", "EffectiveTo", "Tier", "LoPct", "HiPct", "Rate"], plan)
write_csv("Adjustments.csv", ["AdjID", "EmployeeID", "Period", "AmountEUR", "Reason", "ApprovalStatus", "Approver"], adjustments)
write_csv("RecipientChanges.csv", ["ChangeID", "EmployeeID", "EffectiveDate", "ChangeType", "OldValue", "NewValue", "Approver"], recipient_changes)

wb = Workbook()
wb.remove(wb.active)
HFONT = Font(name="Arial", bold=True, color="FFFFFF")
BFONT = Font(name="Arial", size=10)
LFONT = Font(name="Arial", bold=True, color="C00000")


def sheet_table(title, tname, hdr, data, fmts=None, note=None):
    ws = wb.create_sheet(title)
    ws["A1"] = LABEL
    ws["A1"].font = LFONT
    if note:
        ws["A2"] = note
        ws["A2"].font = Font(name="Arial", italic=True, size=9)
    start = 4
    for c, h in enumerate(hdr, 1):
        ws.cell(start, c, h)
    body = data if data else [[None] * len(hdr)]
    for i, row in enumerate(body, start + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(i, c, v)
            cell.font = BFONT
            if fmts and hdr[c - 1] in fmts:
                cell.number_format = fmts[hdr[c - 1]]
    end_row = start + len(body)
    ref = f"A{start}:{get_column_letter(len(hdr))}{end_row}"
    t = Table(displayName=tname, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    for c, h in enumerate(hdr, 1):
        ws.column_dimensions[get_column_letter(c)].width = max(12, len(h) + 4)
    ws.freeze_panes = ws.cell(start + 1, 1)
    return ws


DATE = "yyyy-mm-dd"
sales_rows = [[r["TransactionID"], r["Date"], r["EmployeeID"] or None, r["Region"], r["ProductFamily"],
               r["Currency"], r["AmountLocal"], r["SourceFile"]] for r in rows]
ws = sheet_table("Sales_Raw", "tSales", sales_hdr, sales_rows,
                 {"Date": DATE, "AmountLocal": "#,##0.00"},
                 "Raw landing copy. From Phase 5 this table is replaced by the Power Query folder load.")
ws.column_dimensions["H"].width = 22
sheet_table("FX", "tFX", ["Currency", "Period", "RateToEUR"], fx, {"RateToEUR": "0.000000"},
            "Policy: monthly average rate for the transaction's calendar month, EUR per 1 unit of local currency. Synthetic rates.")
sheet_table("Recipients", "tRecipients", ["EmployeeID", "Role", "ManagerID", "ActiveFrom", "ActiveTo"], recipients,
            {"ActiveFrom": DATE, "ActiveTo": DATE}, "2099-12-31 = open-ended.")
sheet_table("Assignments", "tAssignments", ["EmployeeID", "Region", "StartDate", "EndDate"], assignments,
            {"StartDate": DATE, "EndDate": DATE}, "Effective-dated territory. E104 moves East to West on 2026-09-16.")
sheet_table("Quotas", "tQuotas", ["EmployeeID", "Period", "QuotaEUR"], quotas, {"QuotaEUR": "#,##0"},
            "Monthly quota in EUR. Period is text yyyy-mm.")
ws = sheet_table("Plan", "tPlan", ["PlanVersion", "EffectiveFrom", "EffectiveTo", "Tier", "LoPct", "HiPct", "Rate"], plan,
                 {"EffectiveFrom": DATE, "EffectiveTo": DATE, "LoPct": "0.00%", "HiPct": "0.00%", "Rate": "0.00%"},
                 "Marginal tiers. HiPct 99900% on the top tier means uncapped. v2 changes tier 3 from 4% to 5%.")
sheet_table("Adjustments", "tAdjustments", ["AdjID", "EmployeeID", "Period", "AmountEUR", "Reason", "ApprovalStatus", "Approver"],
            adjustments, {"AmountEUR": "#,##0.00"}, "Only ApprovalStatus = Approved may enter the payout.")
sheet_table("Recipient_Changes", "tRecipientChanges", ["ChangeID", "EmployeeID", "EffectiveDate", "ChangeType", "OldValue", "NewValue", "Approver"],
            recipient_changes, {"EffectiveDate": DATE}, "Audit trail for master-data changes.")
sheet_table("Exceptions", "tExceptions", exceptions_hdr, [], {"CreatedDate": DATE, "ResolvedDate": DATE},
            "Empty until Phase 3.")
sheet_table("Tests", "tTests", tests_hdr, [], None, "Empty until Phase 4.")

tag_info = {
    "DUPLICATE": ("Duplicate TransactionID", "Second occurrence = Duplicate; excluded from eligible sales. First occurrence stays Valid.", "Medium"),
    "BLANKID": ("Blank EmployeeID", "Held; cannot be credited until Commercial supplies the ID.", "High"),
    "UNKNOWN": ("Unknown employee (E199)", "Held; not in tRecipients.", "High"),
    "INACTIVE": ("Inactive recipient (E108 left 2026-09-15)", "Excluded; transaction date after ActiveTo. Route to Commercial for reassignment.", "High"),
    "TERRITORY": ("Territory change: E104 tagged East after move to West", "Held; region does not match assignment on transaction date.", "High"),
    "MISMATCH": ("Region mismatch: E103 is North, tagged South", "Held; confirm with Commercial.", "High"),
    "MISSINGFX": ("Missing FX rate (HUF 2026-10 not in tFX)", "Held; AmountEUR blank, flagged until rate is published.", "High"),
    "CREDIT": ("Negative credit note", "Valid; nets against eligible sales of the booking period. Flag for information.", "Low"),
    "LATE": ("Late prior-period transaction (dated 2026-09-29, arrived in Oct file)", "Valid; credited in processing period 2026-10, not restating closed Sep. Flag Medium; feeds roll-forward late items.", "Medium"),
    "NOQUOTA": ("Missing quota (E106, 2026-10)", "Sale is Valid; payout row for E106/2026-10 is HOLD because Q is missing.", "High"),
    "HUF": ("HUF transaction", "Valid; AmountEUR = AmountLocal x 0.0025 (HUF 2026-09).", "Info"),
}
planted_rows = []
for r in rows:
    tg = r.get("_tag")
    if tg:
        nm, tr, sev = tag_info[tg]
        planted_rows.append([r["TransactionID"], nm, r["EmployeeID"] or "(blank)", r["Date"], r["Currency"], r["AmountLocal"], r["SourceFile"], tr, sev])
planted_rows.append(["E106 / 2026-10", tag_info["NOQUOTA"][0] + " - quota row", "E106", None, None, None, "tQuotas", "No tQuotas row exists for E106 2026-10.", "High"])
planted_rows.append(["ADJ03", "Unapproved adjustment", "E101", None, "EUR", 750, "tAdjustments", "Excluded from payout while ApprovalStatus = Pending.", "Medium"])
sheet_table("Planted_Errors", "tPlantedErrors",
            ["RecordID", "Error", "EmployeeID", "Date", "Currency", "AmountLocal", "Source", "ExpectedTreatment", "Severity"],
            planted_rows, {"Date": DATE, "AmountLocal": "#,##0"}, "Answer key for the deliberate errors. Not used by any calculation.")

dd = [
    ("tSales", "TransactionID", "Text", "Unique key T####. Duplicates are a planted error."),
    ("tSales", "Date", "Date", "Transaction (booking) date."),
    ("tSales", "EmployeeID", "Text", "Credited recipient. Must exist in tRecipients."),
    ("tSales", "Region", "Text", "North / South / East / West. Must match tAssignments on Date."),
    ("tSales", "ProductFamily", "Text", "AIS (Advanced Imaging Solutions), PDx (Pharmaceutical Diagnostics), PCS (Patient Care Solutions)."),
    ("tSales", "Currency", "Text", "EUR or HUF."),
    ("tSales", "AmountLocal", "Number", "Amount in Currency. Negative = credit note."),
    ("tSales", "SourceFile", "Text", "File the row arrived in. Not in the CSVs; added by Power Query from the file name."),
    ("tFX", "Currency / Period / RateToEUR", "Text / Text yyyy-mm / Number", "EUR per 1 unit local, monthly average. Key = Currency & Period."),
    ("tRecipients", "EmployeeID / Role / ManagerID / ActiveFrom / ActiveTo", "Text / Text / Text / Date / Date", "Eligibility window. ManagerIDs are synthetic and not recipients."),
    ("tAssignments", "EmployeeID / Region / StartDate / EndDate", "Text / Text / Date / Date", "Effective-dated territory. Non-overlapping per employee."),
    ("tQuotas", "EmployeeID / Period / QuotaEUR", "Text / Text yyyy-mm / Number", "Monthly quota."),
    ("tPlan", "PlanVersion / EffectiveFrom / EffectiveTo / Tier / LoPct / HiPct / Rate", "Text / Date / Date / Int / % / % / %", "Band = LoPct..HiPct of quota; Rate pays only on sales inside the band."),
    ("tAdjustments", "AdjID / EmployeeID / Period / AmountEUR / Reason / ApprovalStatus / Approver", "Mixed", "Approved only enter payout."),
    ("tRecipientChanges", "ChangeID / EmployeeID / EffectiveDate / ChangeType / OldValue / NewValue / Approver", "Mixed", "Master-data change log (addition)."),
    ("tExceptions", "ExceptionID ... Resolution", "Mixed", "Manual workflow table, populated in Phase 3."),
    ("tTests", "TestID ... PassFail", "Mixed", "Test suite, populated in Phase 4."),
]
sheet_table("Data_Dictionary", "tDataDictionary", ["Table", "Column(s)", "Type", "Definition / rule"], dd)
wb["Data_Dictionary"].column_dimensions["B"].width = 60
wb["Data_Dictionary"].column_dimensions["D"].width = 90
wb["Planted_Errors"].column_dimensions["B"].width = 55
wb["Planted_Errors"].column_dimensions["H"].width = 95

rd = wb.create_sheet("README", 0)
lines = [LABEL, "", "SIC Intelligence Hub - Phase 1 synthetic data",
         "Periods: 2026-08, 2026-09, 2026-10. 8 recipients. Product families: AIS, PDx, PCS.",
         "Folders: data/ = Sales_*.csv for the folder import; reference/ = master tables; demo_hold_back/Sales_2026-11.csv stays OUT of data/ until the demo.",
         "Sheets: Sales_Raw (tSales), FX, Recipients, Assignments, Quotas, Plan, Adjustments, Recipient_Changes, Exceptions, Tests, Planted_Errors, Data_Dictionary."]
for i, l in enumerate(lines, 1):
    rd.cell(i, 1, l).font = LFONT if i == 1 else BFONT
rd.column_dimensions["A"].width = 140
wb.save(os.path.join(OUT, "SIC_Intelligence_Hub.xlsx"))

ref_S = defaultdict(float)
for r in rows:
    tg = r.get("_tag")
    if tg in ("DUPLICATE", "BLANKID", "UNKNOWN", "INACTIVE", "TERRITORY", "MISMATCH", "MISSINGFX"):
        continue
    p = r["SourceFile"][6:13]
    rate = 1.0 if r["Currency"] == "EUR" else dict(((c, pp), v) for c, pp, v in fx)[(r["Currency"], r["Date"].strftime("%Y-%m"))]
    ref_S[(r["EmployeeID"], p)] += r["AmountLocal"] * rate
print("ROWS", len(rows), {k: sum(1 for r in rows if r["SourceFile"] == k) for k in order})
print("NOV", len(nov))
for k in sorted(ref_S):
    print(k, round(ref_S[k], 2))
