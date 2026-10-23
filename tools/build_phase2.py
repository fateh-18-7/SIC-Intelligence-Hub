import re
import sys
import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn, TableFormula
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter, range_boundaries

SRC = "/tmp/sic/phase1_base.xlsx"
MODE = sys.argv[1]
OUT = sys.argv[2]

LET_VARS = ["sales", "quota", "key", "lo", "hi", "rate", "fromD", "toD"]
NEWFN = ["XLOOKUP", "LET", "IFS", "MAXIFS"]


def split_args(s):
    out, depth, cur, q, br = [], 0, "", False, 0
    for ch in s:
        if ch == '"':
            q = not q
        if not q:
            if ch == "[":
                br += 1
            elif ch == "]":
                br -= 1
            elif br == 0 and ch == "(":
                depth += 1
            elif br == 0 and ch == ")":
                depth -= 1
            elif br == 0 and depth == 0 and ch == ",":
                out.append(cur)
                cur = ""
                continue
        cur += ch
    out.append(cur)
    return out


def find_call(s, name):
    m = re.search(r"(?<![A-Za-z_.])" + name + r"\(", s)
    if not m:
        return None
    start = m.start()
    i = m.end()
    depth, q, br = 1, False, 0
    while depth:
        ch = s[i]
        if ch == '"':
            q = not q
        elif not q:
            if ch == "[":
                br += 1
            elif ch == "]":
                br -= 1
            elif br == 0 and ch == "(":
                depth += 1
            elif br == 0 and ch == ")":
                depth -= 1
        i += 1
    return start, m.end(), i


def sub_ident(expr, name, repl):
    out, i, q, br = "", 0, False, 0
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            q = not q
        if not q and ch == "[":
            br += 1
        if not q and ch == "]":
            br -= 1
        if not q and br == 0 and expr.startswith(name, i):
            before = expr[i - 1] if i else " "
            after = expr[i + len(name)] if i + len(name) < len(expr) else " "
            if not (before.isalnum() or before in "_.") and not (after.isalnum() or after in "_.("):
                out += repl
                i += len(name)
                continue
        out += ch
        i += 1
    return out


def expand_let(s):
    while True:
        c = find_call(s, "LET")
        if not c:
            return s
        a, b, e = c
        args = split_args(s[b:e - 1])
        body = args[-1]
        pairs = list(zip(args[0:-1:2], args[1:-1:2]))
        for idx in range(len(pairs)):
            n, v = pairs[idx]
            n = n.strip()
            v = "(" + v + ")"
            body = sub_ident(body, n, v)
            pairs = pairs[:idx + 1] + [(pn, sub_ident(pv, n, v)) for pn, pv in pairs[idx + 1:]]
        s = s[:a] + "(" + body + ")" + s[e:]


def expand_xlookup(s):
    while True:
        c = find_call(s, "XLOOKUP")
        if not c:
            return s
        a, b, e = c
        args = [expand_xlookup(x) for x in split_args(s[b:e - 1])]
        core = f"INDEX({args[2]},MATCH({args[0]},{args[1]},0))"
        rep = f"IFERROR({core},{args[3]})" if len(args) > 3 else core
        s = s[:a] + rep + s[e:]


def finalize(f, table):
    f = re.sub(r"\[@(\w+)\]", lambda m: f"{table}[[#This Row],[{m.group(1)}]]", f)
    if MODE == "lo":
        f = expand_let(f)
        f = expand_xlookup(f)
    else:
        def letfix(s):
            c = find_call(s, "LET")
            if not c:
                return s
            a, b, e = c
            args = split_args(s[b:e - 1])
            names = [x.strip() for x in args[0:-1:2]]
            inner = ",".join(args)
            for n in names:
                inner = sub_ident(inner, n, "_xlpm." + n)
            return s[:a] + "LET(" + letfix(inner) + ")" + letfix(s[e:])
        f = letfix(f)
    for fn in NEWFN:
        f = re.sub(r"(?<![A-Za-z_.])" + fn + r"\(", "_xlfn." + fn + "(", f)
    return f


wb = load_workbook(SRC)
FONT = Font(name="Arial", size=10)
HFONT = Font(name="Arial", bold=True)
CALC_FILL = PatternFill("solid", fgColor="EAF1FB")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def locate(tname):
    for ws in wb:
        if tname in ws.tables:
            return ws, ws.tables[tname]


def add_columns(tname, cols, fmts=None):
    ws, t = locate(tname)
    c1, r1, c2, r2 = range_boundaries(t.ref)
    nid = max(tc.id for tc in t.tableColumns) + 1
    for k, (hdr, formula) in enumerate(cols):
        c = c2 + 1 + k
        ws.cell(r1, c, hdr).font = HFONT
        f = finalize(formula, tname)
        for r in range(r1 + 1, r2 + 1):
            cell = ws.cell(r, c, f)
            cell.font = FONT
            cell.fill = CALC_FILL
            if fmts and hdr in fmts:
                cell.number_format = fmts[hdr]
        tc = TableColumn(id=nid + k, name=hdr)
        tc.calculatedColumnFormula = TableFormula(attr_text=f[1:])
        t.tableColumns.append(tc)
        ws.column_dimensions[get_column_letter(c)].width = max(12, len(hdr) + 3)
    t.ref = f"{get_column_letter(c1)}{r1}:{get_column_letter(c2 + len(cols))}{r2}"
    if t.autoFilter is not None:
        t.autoFilter.ref = t.ref


def new_table(sheet, tname, top, inputs, calc_cols, rows, fmts=None, note=None, label=True):
    ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
    if label:
        ws["A1"] = "SYNTHETIC DATA - fictional prototype inspired by the SIC role. Not GE HealthCare plans, systems, people or data."
        ws["A1"].font = Font(name="Arial", bold=True, color="C00000")
    if note:
        ws["A2"] = note
        ws["A2"].font = Font(name="Arial", italic=True, size=9)
    hdrs = inputs + [h for h, _ in calc_cols]
    for c, h in enumerate(hdrs, 1):
        ws.cell(top, c, h).font = HFONT
    finals = [finalize(f, tname) for _, f in calc_cols]
    for i, row in enumerate(rows, top + 1):
        for c, v in enumerate(row, 1):
            cell = ws.cell(i, c, v)
            cell.font = FONT
            cell.fill = INPUT_FILL
            if fmts and hdrs[c - 1] in fmts:
                cell.number_format = fmts[hdrs[c - 1]]
        for k, f in enumerate(finals):
            c = len(inputs) + 1 + k
            cell = ws.cell(i, c, f)
            cell.font = FONT
            cell.fill = CALC_FILL
            if fmts and hdrs[c - 1] in fmts:
                cell.number_format = fmts[hdrs[c - 1]]
    ref = f"A{top}:{get_column_letter(len(hdrs))}{top + len(rows)}"
    t = Table(displayName=tname, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    t._initialise_columns()
    for k, f in enumerate(finals):
        t.tableColumns[len(inputs) + k].calculatedColumnFormula = TableFormula(attr_text=f[1:])
    for c, h in enumerate(hdrs, 1):
        ws.column_dimensions[get_column_letter(c)].width = max(12, len(h) + 3)
    ws.freeze_panes = ws.cell(top + 1, 1)
    return ws


ps = wb.create_sheet("Parameters", 1)
ps["A1"] = "SYNTHETIC DATA - fictional prototype inspired by the SIC role. Not GE HealthCare plans, systems, people or data."
ps["A1"].font = Font(name="Arial", bold=True, color="C00000")
ps["A3"], ps["B3"], ps["C3"] = "Name", "Value", "Used by"
for c in "ABC":
    ps[f"{c}3"].font = HFONT
ps["A4"], ps["B4"], ps["C4"] = "p_Approved", "Approved", "tPayout[ApprovedAdjEUR], tPayout[PendingAdjEUR]"
ps["B4"].fill = INPUT_FILL
for c in "ABC":
    ps[f"{c}4"].font = FONT
ps.column_dimensions["A"].width = 16
ps.column_dimensions["B"].width = 14
ps.column_dimensions["C"].width = 60
wb.defined_names["p_Approved"] = DefinedName("p_Approved", attr_text="Parameters!$B$4")

add_columns("tFX", [("Key", '=[@Currency]&"|"&[@Period]')])
add_columns("tQuotas", [("Key", '=[@EmployeeID]&"|"&[@Period]')])
add_columns("tPlan", [("Key", '=[@PlanVersion]&"|"&[@Tier]')])

DATE = "yyyy-mm-dd"
add_columns("tSales", [
    ("Period", "=MID([@SourceFile],7,7)"),
    ("TxnMonth", '=YEAR([@Date])&"-"&TEXT(MONTH([@Date]),"00")'),
    ("IsDuplicate", "=COUNTIF(INDEX(tSales[TransactionID],1):[@TransactionID],[@TransactionID])>1"),
    ("MissingID", '=[@EmployeeID]=""'),
    ("KnownEmp", "=COUNTIF(tRecipients[EmployeeID],[@EmployeeID])>0"),
    ("ActiveOnDate", '=LET(fromD,XLOOKUP([@EmployeeID],tRecipients[EmployeeID],tRecipients[ActiveFrom],""),toD,XLOOKUP([@EmployeeID],tRecipients[EmployeeID],tRecipients[ActiveTo],""),IF(fromD="",FALSE,AND([@Date]>=fromD,[@Date]<=toD)))'),
    ("RegionOK", '=COUNTIFS(tAssignments[EmployeeID],[@EmployeeID],tAssignments[Region],[@Region],tAssignments[StartDate],"<="&[@Date],tAssignments[EndDate],">="&[@Date])>0'),
    ("FXRate", '=XLOOKUP([@Currency]&"|"&[@TxnMonth],tFX[Key],tFX[RateToEUR],"")'),
    ("AmountEUR", '=IF([@FXRate]="","",ROUND([@AmountLocal]*[@FXRate],2))'),
    ("IsCreditNote", "=[@AmountLocal]<0"),
    ("IsLate", "=[@TxnMonth]<>[@Period]"),
    ("Status", '=IFS([@IsDuplicate],"Duplicate",[@MissingID],"Held",NOT([@KnownEmp]),"Held",NOT([@ActiveOnDate]),"Excluded",NOT([@RegionOK]),"Held",[@FXRate]="","Held",TRUE,"Valid")'),
    ("Reason", '=IFS([@IsDuplicate],"Duplicate TransactionID",[@MissingID],"Missing EmployeeID",NOT([@KnownEmp]),"Unknown employee",NOT([@ActiveOnDate]),"Inactive on transaction date",NOT([@RegionOK]),"Region does not match assignment",[@FXRate]="","Missing FX rate",[@IsLate],"Late prior-period item",[@IsCreditNote],"Credit note",TRUE,"")'),
    ("EligibleEUR", '=IF([@Status]="Valid",[@AmountEUR],0)'),
], {"FXRate": "0.000000", "AmountEUR": "#,##0.00", "EligibleEUR": "#,##0.00"})


def tier(n):
    return ('=IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],key,[@PlanVersion]&"|"&' + str(n) +
            ',lo,XLOOKUP(key,tPlan[Key],tPlan[LoPct]),hi,XLOOKUP(key,tPlan[Key],tPlan[HiPct]),rate,XLOOKUP(key,tPlan[Key],tPlan[Rate]),ROUND(rate*MAX(0,MIN(sales,hi*quota)-lo*quota),2)))')


CALC = ('=IF([@QuotaEUR]="","",LET(sales,[@EligibleSalesEUR],quota,[@QuotaEUR],ROUND(SUMPRODUCT((tPlan[PlanVersion]=[@PlanVersion])*tPlan[Rate]*'
        '(sales>tPlan[LoPct]*quota)*(sales-(sales>tPlan[HiPct]*quota)*(sales-tPlan[HiPct]*quota)-tPlan[LoPct]*quota)),2)))')

payout_cols = [
    ("PeriodStart", "=DATE(LEFT([@Period],4),RIGHT([@Period],2),1)"),
    ("PlanVersion", '=LET(fromD,MAXIFS(tPlan[EffectiveFrom],tPlan[EffectiveFrom],"<="&[@PeriodStart]),toD,XLOOKUP(fromD,tPlan[EffectiveFrom],tPlan[EffectiveTo],0),IF(toD<[@PeriodStart],"None",XLOOKUP(fromD,tPlan[EffectiveFrom],tPlan[PlanVersion])))'),
    ("EligibleSalesEUR", "=SUMIFS(tSales[EligibleEUR],tSales[EmployeeID],[@EmployeeID],tSales[Period],[@Period])"),
    ("QuotaEUR", '=XLOOKUP([@EmployeeID]&"|"&[@Period],tQuotas[Key],tQuotas[QuotaEUR],"")'),
    ("Attainment", '=IF([@QuotaEUR]="","",[@EligibleSalesEUR]/[@QuotaEUR])'),
    ("Tier1EUR", tier(1)),
    ("Tier2EUR", tier(2)),
    ("Tier3EUR", tier(3)),
    ("CalcPayoutEUR", CALC),
    ("ApprovedAdjEUR", "=SUMIFS(tAdjustments[AmountEUR],tAdjustments[EmployeeID],[@EmployeeID],tAdjustments[Period],[@Period],tAdjustments[ApprovalStatus],p_Approved)"),
    ("PendingAdjEUR", '=SUMIFS(tAdjustments[AmountEUR],tAdjustments[EmployeeID],[@EmployeeID],tAdjustments[Period],[@Period],tAdjustments[ApprovalStatus],"<>"&p_Approved)'),
    ("HeldTxnCount", '=COUNTIFS(tSales[EmployeeID],[@EmployeeID],tSales[Period],[@Period],tSales[Status],"Held")'),
    ("FinalPayoutEUR", '=IF([@CalcPayoutEUR]="","",[@CalcPayoutEUR]+[@ApprovedAdjEUR])'),
    ("ReviewStatus", '=IFS([@QuotaEUR]="","HOLD - missing quota",[@PlanVersion]="None","HOLD - no plan",[@HeldTxnCount]>0,"REVIEW - held transactions",[@PendingAdjEUR]<>0,"REVIEW - pending adjustment",TRUE,"Calculated")'),
]
qws, qt = locate("tQuotas")
qrows = [[c.value for c in r[:2]] for r in qws[qt.ref]][1:]
roster = sorted(qrows + [["E106", "2026-10"]], key=lambda x: (x[1], x[0]))
money = "#,##0.00"
pws = new_table("Calc_Payout", "tPayout", 6, ["EmployeeID", "Period"], payout_cols, roster,
                {"PeriodStart": DATE, "EligibleSalesEUR": money, "QuotaEUR": "#,##0", "Attainment": "0.00%",
                 "Tier1EUR": money, "Tier2EUR": money, "Tier3EUR": money, "CalcPayoutEUR": money,
                 "ApprovedAdjEUR": money, "PendingAdjEUR": money, "FinalPayoutEUR": money},
                "Roster = one row per recipient x period (yellow = input). Blue = formulas. RosterGap must be 0.")
pws["A4"] = "RosterGap (valid sales with no tPayout row):"
pws["A4"].font = HFONT
pws["E4"] = '=SUMPRODUCT((tSales[Status]="Valid")*(COUNTIFS(tPayout[EmployeeID],tSales[EmployeeID],tPayout[Period],tSales[Period])=0))'
pws["E4"].font = HFONT
pws.column_dimensions["P"].width = 28

S = '=IF([@QuotaEUR]="","",LET(sales,[@SalesEUR],quota,[@QuotaEUR],ROUND(SUMPRODUCT((tPlan[PlanVersion]=[@PlanVersion])*tPlan[Rate]*(sales>tPlan[LoPct]*quota)*(sales-(sales>tPlan[HiPct]*quota)*(sales-tPlan[HiPct]*quota)-tPlan[LoPct]*quota)),2)))'
sanity = [
    ["SC1", "v1", 350000, 300000, 3200, "Q=300k. T2: 2% x (300k-240k)=1,200. T3: 4% x (350k-300k)=2,000. Total 3,200"],
    ["SC2", "v1", 180000, 200000, 400, "T2: 2% x (180k-160k)=400. T3: 0"],
    ["SC3", "v1", 250000, 200000, 2800, "T2: 2% x 40k=800. T3: 4% x 50k=2,000"],
    ["SC4", "v1", 120000, 100000, 1200, "T2: 2% x 20k=400. T3: 4% x 20k=800"],
    ["BD1", "v1", 79990, 100000, 0, "79.99%: below 80k threshold, T2 = 0"],
    ["BD2", "v1", 80000, 100000, 0, "80.00%: MIN(80k,100k)-80k = 0"],
    ["BD3", "v1", 80010, 100000, 0.2, "80.01%: 2% x 10 = 0.20"],
    ["BD4", "v1", 100000, 100000, 400, "100.00%: T2 max 2% x 20k=400, T3 = 0"],
    ["BD5", "v1", 100010, 100000, 400.4, "100.01%: 400 + 4% x 10 = 400.40"],
    ["PV1", "v2", 120000, 100000, 1400, "v2: T2 400 + T3 5% x 20k=1,000"],
]
sws = new_table("Sanity_Checks", "tSanity", 4, ["CaseID", "PlanVersion", "SalesEUR", "QuotaEUR", "ExpectedPayout", "Working"],
                [("EnginePayout", S), ("Diff", "=[@EnginePayout]-[@ExpectedPayout]"), ("Result", '=IF(ABS([@Diff])<0.005,"PASS","FAIL")')],
                sanity, {"SalesEUR": "#,##0", "QuotaEUR": "#,##0", "ExpectedPayout": money, "EnginePayout": money, "Diff": money},
                "ExpectedPayout is typed, hand-derived (see Working). EnginePayout uses the same formula text as tPayout[CalcPayoutEUR].")
sws.column_dimensions["F"].width = 70
sws["A3"] = "Failed sanity checks:"
sws["A3"].font = HFONT
sws["C3"] = '=ROWS(tSanity[Result])-COUNTIF(tSanity[Result],"PASS")'

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
