from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether)
from reportlab.lib.enums import TA_CENTER

OUT = "/tmp/sic/SIC_Intelligence_Hub_Guide.pdf"
NAVY = colors.HexColor("#1F3864")
LIGHT = colors.HexColor("#EAF1FB")
YEL = colors.HexColor("#FFF2CC")
GREEN = colors.HexColor("#E2F0D9")
RED = colors.HexColor("#C00000")

ss = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss["Heading1"], fontName="Helvetica-Bold", fontSize=17, textColor=NAVY, spaceBefore=6, spaceAfter=8)
H2 = ParagraphStyle("H2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, textColor=NAVY, spaceBefore=10, spaceAfter=4)
B = ParagraphStyle("B", parent=ss["BodyText"], fontName="Helvetica", fontSize=10, leading=14.5, spaceAfter=5)
SM = ParagraphStyle("SM", parent=B, fontSize=8.8, leading=12)
CELL = ParagraphStyle("CELL", parent=B, fontSize=9, leading=12, spaceAfter=0)
CELLB = ParagraphStyle("CELLB", parent=CELL, fontName="Helvetica-Bold")
QUOTE = ParagraphStyle("Q", parent=B, backColor=LIGHT, borderPadding=7, leftIndent=6, rightIndent=6, spaceBefore=4, spaceAfter=10)
TITLE = ParagraphStyle("T", parent=H1, fontSize=26, leading=32, alignment=TA_CENTER, spaceAfter=6)
SUB = ParagraphStyle("S", parent=B, fontSize=12, alignment=TA_CENTER, textColor=colors.HexColor("#444444"))


def p(t, s=B):
    return Paragraph(t, s)


def bullets(items, s=B):
    return [Paragraph("&bull;&nbsp;&nbsp;" + i, ParagraphStyle("bl", parent=s, leftIndent=12, firstLineIndent=-9)) for i in items]


def table(rows, widths, head=True, shade=None):
    data = [[Paragraph(str(c), CELLB if (head and i == 0) else CELL) for c in r] for i, r in enumerate(rows)]
    t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1 if head else 0)
    st = [("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
          ("VALIGN", (0, 0), (-1, -1), "TOP"),
          ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]
    if head:
        st += [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white)]
        data[0] = [Paragraph(f'<font color="white"><b>{r}</b></font>', CELL) for r in rows[0]]
        t = Table(data, colWidths=[w * mm for w in widths], repeatRows=1)
    if shade:
        st.append(("BACKGROUND", (0, 1), (-1, -1), shade))
    t.setStyle(TableStyle(st))
    return t


def box(text, color=LIGHT):
    t = Table([[Paragraph(text, B)]], colWidths=[170 * mm])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), color), ("BOX", (0, 0), (-1, -1), 0.5, NAVY),
                           ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                           ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def footer(c, d):
    c.saveState()
    c.setFont("Helvetica", 7.5)
    c.setFillColor(colors.HexColor("#777777"))
    c.drawString(20 * mm, 10 * mm, "SIC Intelligence Hub - beginner's guide. Synthetic prototype, not GE HealthCare data or systems.")
    c.drawRightString(190 * mm, 10 * mm, f"Page {d.page}")
    c.restoreState()


s = []
s += [Spacer(1, 45 * mm), p("SIC Intelligence Hub", TITLE), p("A beginner's guide: what it is, how it works, how to use it, and how to sell it in your interview", SUB),
      Spacer(1, 14 * mm)]
s.append(box("<b>Read this first.</b> Everything in this project is <b>synthetic</b> (made-up) data. It is a prototype <i>inspired by</i> the "
             "Sales Incentive Compensation (SIC) Intern role at GE HealthCare Budapest. It does not use or copy GE HealthCare's real plans, systems, "
             "people or numbers. Say this out loud at the start of any demo.", YEL))
s.append(Spacer(1, 10 * mm))
s.append(table([["Where things are", ""],
                ["Project folder (your Mac)", "Documents / SIC_Intelligence_Hub"],
                ["Main file", "SIC_Intelligence_Hub.xlsx (open it in Excel, start on the <b>Controls</b> sheet)"],
                ["Power BI report", "app.powerbi.com > My workspace > <b>SIC Intelligence Hub - Dashboard (synthetic)</b>"],
                ["Demo file", "demo_hold_back / Sales_2026-11.csv (keep it OUT of the data folder until the demo)"],
                ["One-slide summary", "JD_Mapping.pptx"],
                ["Deeper notes", "docs folder (formulas, Power Query code, interview Q&amp;A)"]], [50, 120]))
s.append(PageBreak())

s.append(p("1. The project in one minute", H1))
s.append(p("Sales people earn a bonus (an <b>incentive</b>) based on how much they sell compared with a target (their <b>quota</b>). "
           "Calculating this sounds simple, but in real life the sales data is messy: missing IDs, people who left the company, "
           "sales in the wrong region, foreign currency, late files, and people changing territory in the middle of a month. "
           "A mistake means someone is paid wrongly, which hurts trust and costs money."))
s.append(p("This project is a small, complete <b>incentive calculation system</b> built in Excel, with Power Query for automation, "
           "Power BI for the dashboard, and a carefully limited AI helper. It does four jobs:"))
s += bullets(["<b>Cleans and checks</b> the sales data, and flags every problem.",
              "<b>Calculates</b> each person's incentive with a clear, fair formula.",
              "<b>Proves</b> the numbers are right with independent checks and tests.",
              "<b>Stops</b> the payment (status HOLD) until problems are fixed and a <b>human</b> approves."])
s.append(Spacer(1, 4))
s.append(box("<b>The golden rule of the project:</b> AI explains, Excel calculates, a human approves.<br/>"
             "AI never changes a number, never changes a status, and never approves anything."))

s.append(p("2. Key words, explained simply", H1))
s.append(table([["Word", "What it means", "Example from the project"],
                ["SIC", "Sales Incentive Compensation: the bonus paid to sales people for selling.", "The whole project."],
                ["Recipient", "A person who can earn an incentive.", "8 people, E101 to E108."],
                ["Quota", "The sales target for a month.", "E105's quota is 200,000 EUR."],
                ["Eligible sales", "Sales that are allowed to count (clean, valid ones).", "E105 sold 250,000 EUR in August."],
                ["Attainment", "Sales divided by quota, as a %.", "250,000 / 200,000 = 125%."],
                ["Tier", "A band of attainment with its own rate.", "0-80% pays 0%, 80-100% pays 2%, over 100% pays 4%."],
                ["Marginal tiers", "Each rate applies only to the sales inside its band (like income tax bands). No sudden jumps.", "See the worked example on the next page."],
                ["Plan version", "The set of rules for a period. Rules can change.", "v1 for Aug-Sep, v2 from Oct (top rate 4% to 5%)."],
                ["Adjustment", "A manual extra or deduction. Only counts if approved.", "+750 for E101 is Pending, so not paid."],
                ["Exception", "A problem found in the data that someone must fix.", "A sale with no employee ID."],
                ["Reconciliation", "Checking that two independent numbers agree.", "Rows in the files = rows after cleaning."],
                ["Control", "A check that stops mistakes reaching the payment.", "12 checks on the Controls sheet."],
                ["Gate / status", "The overall traffic light for the month.", "HOLD, READY FOR REVIEW, APPROVED."],
                ["FX", "Foreign exchange: converting another currency to EUR.", "HUF sale x 0.0025 = EUR."]], [28, 78, 64]))
s.append(PageBreak())

s.append(p("3. How it works: the flow", H1))
s.append(p("Think of it as a factory line with five stations. Data goes in on the left, a checked and approved payout comes out on the right."))
flow = [["1. Data in", "2. Check", "3. Calculate", "4. Prove", "5. Approve &amp; report"],
        ["Monthly sales CSV files land in the <b>data</b> folder. Power Query reads them.",
         "Every sale gets a status: Valid, Held, Duplicate or Excluded, with a reason.",
         "Valid sales are converted to EUR, added up per person and month, and run through the tier formula.",
         "Four reconciliations, 18 tests and boundary checks must all pass.",
         "A person types name, date and the approved total. Power BI and the AI summary report the result."]]
s.append(table(flow, [34, 34, 34, 34, 34]))
s.append(Spacer(1, 6))

s.append(p("The tier formula, with a worked example", H2))
s.append(p("<b>E105 in August:</b> sales 250,000 EUR, quota 200,000 EUR (attainment 125%), plan v1."))
s.append(table([["Band", "Sales inside this band", "Rate", "Pays"],
                ["0% - 80% (0 to 160,000)", "160,000", "0%", "0"],
                ["80% - 100% (160,000 to 200,000)", "40,000", "2%", "800"],
                ["Above 100% (above 200,000)", "50,000", "4%", "2,000"],
                ["<b>Total incentive</b>", "", "", "<b>2,800 EUR</b>"]], [70, 45, 20, 35]))
s.append(Spacer(1, 4))
s.append(p("In Excel this is one formula: <b>Rate x MAX(0, MIN(Sales, Top of band) - Bottom of band)</b>, added up for each tier. "
           "Why marginal? Because there is no cliff: selling one more euro never causes a big jump in pay. At exactly 80% the payout is 0; at 80.01% it is 0.20 EUR. "
           "That is fair, easy to explain and hard to game."))
s.append(PageBreak())

s.append(p("4. The main components", H1))
s.append(p("Each block below says <b>what it is</b>, <b>where it is</b>, and <b>what to say</b> about it."))
comps = [
    ("Synthetic data", "data folder + sheets FX, Recipients, Assignments, Quotas, Plan, Adjustments",
     "3 months (Aug-Oct 2026), 8 people, 59 sales, with 11 <b>deliberate errors</b> planted so the checks have something to catch (list on sheet Planted_Errors).",
     "\"I planted realistic errors on purpose, so I could prove each control actually catches something.\""),
    ("Validation", "Sales_Raw sheet, columns Status and Reason",
     "Each sale is checked: duplicate? missing ID? unknown person? person active on that date? right region on that date? FX rate available? Then it gets one status.",
     "\"Every row ends up in exactly one bucket, so nothing can disappear silently.\""),
    ("Calculation engine", "Calc_Payout sheet (table tPayout)",
     "One row per person per month: eligible sales, quota, attainment, the three tier amounts, approved adjustments, final payout, and a review status.",
     "\"Rates, dates and thresholds come from tables, nothing is typed into formulas.\""),
    ("Controls and the gate", "Controls sheet (start here)",
     "12 checks that must all show PASS. If any fails, the month is on HOLD. APPROVED needs a person's name, date and the approved total. If a number changes after approval, it drops back to review.",
     "\"Excel can never approve itself. And a broken formula shows FAIL, not a false green.\""),
    ("Four reconciliations", "Control_Waterfall, Control_CrossCut, Control_Recalc, Control_RollForward",
     "(1) Rows/amounts in the files = rows/amounts after cleaning. (2) Totals by region/product = totals by person. (3) A second, different calculation gives the same payout. (4) Last month + changes = this month.",
     "\"No check compares a number with itself; each uses an independent path.\""),
    ("Exceptions workflow", "Exceptions sheet",
     "Every problem gets an owner (Sales Ops, Finance, Commercial), a status (Open / In progress / Closed), an age in days, and a resolution note.",
     "\"Closing an exception needs a written resolution, so there is an audit trail.\""),
    ("Test suite", "Tests sheet + Sanity_Checks sheet",
     "18 tests with answers worked out by hand (boundaries, currency, plan change, territory change, late sale...). All pass.",
     "\"I know the calculation is right because hand-worked answers match the engine.\""),
    ("Power Query automation", "Queries qSalesFiles, qRawCapture, qSchemaCheck (sheets with the same names)",
     "Reads every Sales_*.csv in the data folder, counts rows and amounts before any cleaning, and checks every file has the right columns. Data > Refresh All updates it.",
     "\"A new month is one click. If a file arrives that isn't processed, the waterfall check catches it.\""),
    ("Power BI dashboard", "app.powerbi.com > My workspace",
     "One page: total incentive (13.93K EUR), eligible sales (2.50M EUR), monthly trend, incentive by region, and a detail table with each person's review status.",
     "\"The dashboard reads the same verified numbers; it never recalculates them.\""),
    ("AI layer", "AI_Prompt, Claim_Verification, AI_Eval sheets",
     "Excel builds a short fact list (totals only, no names) to give to an approved AI tool. The AI drafts commentary. Every number it writes is checked against its source cell; one mismatch blocks publication.",
     "\"AI only explains. It is fed aggregates, and its claims are verified before anyone reads them.\""),
    ("Stakeholder pack and process", "Stakeholder_Pack, SOP, Change_Log, Run_Log, Recipient_Changes",
     "Ready-made notice for Commercial, a Finance summary, an HR change report, a 10-step monthly process, and logs of changes and runs.",
     "\"It's not just a calculator; it supports the monthly cycle with the people involved.\""),
]
for name, where, what, say in comps:
    s.append(KeepTogether([table([[f"<b>{name}</b>", f"<i>{where}</i>"], ["What it does", what], ["What to say", say]], [32, 138], head=False),
                           Spacer(1, 5)]))
s.append(PageBreak())

s.append(p("5. How to use it, step by step", H1))
s.append(p("A. Look at this month's status", H2))
s += bullets(["Open <b>Documents / SIC_Intelligence_Hub / SIC_Intelligence_Hub.xlsx</b> in Excel.",
              "Go to the <b>Controls</b> sheet. You will see 11 checks PASS and 1 FAIL (C8: 4 open High exceptions). RUN STATUS = <b>HOLD</b>. That is correct: there are problems to fix.",
              "Go to <b>Calc_Payout</b> to see each person's payout and review status."])
s.append(p("B. Fix a problem (exception)", H2))
s += bullets(["Go to <b>Exceptions</b>. Find an Open row, for example EX009 (missing HUF rate for October).",
              "Change Status to <b>Closed</b>, type a ResolvedDate and a Resolution note.",
              "Back on Controls, the open-High count goes down. When all checks pass, the status becomes <b>READY FOR REVIEW</b>."])
s.append(p("C. Approve (only a human can)", H2))
s += bullets(["On Controls, a reviewer types their <b>name</b> (D20), the <b>date</b> (D21) and the <b>approved total</b> shown next to it (D22).",
              "The status becomes <b>APPROVED</b>. If anyone changes a number afterwards, it automatically drops back to review.",
              "If someone tries to approve while on HOLD, a warning appears and the approval is ignored."])
s.append(p("D. Load a new month (the demo)", H2))
s += bullets(["In Finder, move <b>demo_hold_back / Sales_2026-11.csv</b> into the <b>data</b> folder.",
              "In Excel: <b>Data > Refresh All</b>.",
              "Run_Log now shows 4 files and 73 rows. The waterfall shows <b>14 rows not processed</b>, 3 checks fail, status stays HOLD. "
              "Message: <i>\"New data arrived, the system noticed, and it will not pay until the data is processed and checked.\"</i>",
              "Afterwards move the file back to demo_hold_back and press Refresh All again (back to 3 files, 59 rows)."])
s.append(p("E. Open the dashboard", H2))
s += bullets(["Go to app.powerbi.com, sign in with your university account, My workspace, open <b>SIC Intelligence Hub - Dashboard (synthetic)</b>.",
              "Point at the total (13,930 EUR) and say it matches Excel exactly; the trend shows Aug 5,120, Sep 3,580, Oct 5,230."])
s.append(p("F. Use the AI safely", H2))
s += bullets(["Copy cell <b>AI_Prompt!A19</b> (a short list of verified totals) into an approved AI tool together with the prompt in docs/04_AI_Layer.md.",
              "For every number in the AI's answer, type it into <b>Claim_Verification</b>. All must be Y, otherwise publication is BLOCKED."])
s.append(PageBreak())

s.append(p("6. How to sell it in the interview", H1))
s.append(p("The 30-second pitch", H2))
s.append(box("\"I built a small end-to-end incentive compensation prototype on synthetic data. It takes raw monthly sales files, cleans and validates them, "
             "calculates marginal-tier payouts with effective-dated plans and FX, and then proves the result with four independent reconciliations and "
             "18 hand-checked tests. Nothing is paid until a person approves, and Excel can never approve itself. Power Query automates the monthly load, "
             "Power BI shows the results, and AI is used only to draft commentary, with every number it writes verified against the source.\""))
s.append(p("Three messages to repeat", H2))
s += bullets(["<b>Accuracy you can prove:</b> independent checks and tests, not just a formula.",
              "<b>Control and integrity:</b> HOLD until clean; human approval; audit trail. This matches GE HealthCare's stated values of transparency and integrity.",
              "<b>Practical automation and AI:</b> saves manual effort, but AI never touches the payout."])
s.append(p("5-minute demo order", H2))
s.append(table([["Time", "Show", "Say"],
                ["0:00", "Controls sheet", "The problem, synthetic data, the HOLD status and why."],
                ["0:40", "Calc_Payout, E105 August", "Marginal tiers: 800 + 2,000 = 2,800. Boundaries are tested."],
                ["1:30", "Sales_Raw T1035 + Exceptions", "Mid-month territory change: a wrong-region sale is held for Sales Ops."],
                ["2:15", "Drop the November file, Refresh All", "Waterfall catches 14 unprocessed rows, still HOLD."],
                ["3:15", "Power BI report", "Same verified numbers, by month and region."],
                ["4:00", "AI_Prompt + Claim_Verification", "Aggregates in, every claim checked."],
                ["4:40", "Type an approver while on HOLD", "Warning appears: Excel can't self-approve."]], [16, 54, 100]))
s.append(Spacer(1, 4))
s.append(p("If something fails live: show screenshots you took while rehearsing, and explain what should happen. Staying calm and honest scores well.", SM))
s.append(PageBreak())

s.append(p("7. Likely questions and short answers", H1))
qa = [
    ("How do you know the calculation is correct?", "Three layers: hand-worked tests (18, all pass), a second independent calculation that must match every row, and reconciliations from the files to the payout and month to month."),
    ("What happens when someone changes territory mid-month?", "Assignments have start and end dates. Each sale is checked against the region valid on its own date. E104's East sale after moving to West is held for confirmation."),
    ("Why marginal tiers?", "No cliffs. One more euro never causes a jump, so it's fairer and harder to game. You can see it at 80.00% (pays 0) vs 80.01% (pays 0.20)."),
    ("Incentive expense jumped. How would you investigate?", "Start with the roll-forward: split the change into existing people, new joiners, leavers and adjustments. October's +46% is mostly E101 at 148% under the new 5% top rate."),
    ("How would you use AI safely with pay data?", "Only approved tools, only totals (no names), AI never calculates or approves, every claim is checked against the source, and a human signs off."),
    ("What would you automate first in a real team?", "Data intake and validation: loading files, schema checks, standard flags and the exception log. It is repetitive, error-prone and comes before everything else."),
    ("Finance and Commercial disagree on a figure. What do you do?", "Go back to the source record and the policy definition, show the reconciliation trail, hold the item rather than change it for either side, and record who decided."),
    ("What are the limitations?", "Real plans are more complex (caps, multiple measures, clawbacks), there's no proration for part-month joiners, and the data is synthetic. I'd say that openly."),
    ("Why Excel and not Python or SQL?", "Excel is transparent and what a reviewer can audit today. The logic is table-driven, so the same rules and tests could move to SQL or Python for scale."),
    ("Tell me about a mistake you caught.", "My first failed-check counter ignored error values, so a broken formula showed 0 failures, a false green. I changed it to count anything that isn't PASS. Also, real Excel found a file-structure issue my test tool missed; I fixed it and re-tested in Excel."),
]
for q, a in qa:
    s.append(KeepTogether([p(f"<b>Q: {q}</b>"), p(a), Spacer(1, 2)]))

s.append(p("Good questions to ask them", H2))
s += bullets(["How did this year's sales reorganisation change territories or plans?",
              "Do you credit on orders, shipment or revenue recognition?",
              "Which tools does the team use today, and where is most manual effort?",
              "Which AI tools are approved for finance data, and what review applies?"])
s.append(PageBreak())

s.append(p("8. Honest limitations (say these before they ask)", H1))
s += bullets(["All data is synthetic; GE HealthCare's real plans are unknown and not copied.",
              "Plans are simplified: no caps, no multiple measures, no clawbacks, no part-month proration.",
              "Power Query reads and counts the files; the sales table itself is still loaded into Excel separately (kept simple on purpose). The waterfall check makes any gap visible.",
              "The Power BI report is one basic page built in the browser; row-level security and a full star schema are documented in docs/03 but not built.",
              "AI pass rates and time savings are not measured yet. Only quote numbers you have actually measured."])
s.append(p("9. Cheat sheet: numbers to remember", H1))
s.append(table([["What", "Value"],
                ["Months / people / sales rows", "Aug-Oct 2026 / 8 / 59 (73 with the November demo file)"],
                ["Planted errors", "11 types (duplicate, blank ID, unknown person, leaver, wrong region x2, missing FX, credit note, late sale, missing quota, unapproved adjustment)"],
                ["Total incentive (all months)", "13,930 EUR (Aug 5,120 / Sep 3,580 / Oct 5,230)"],
                ["Eligible sales (all months)", "2,497,000 EUR"],
                ["October weighted attainment", "102.6%"],
                ["Checks on Controls", "12 (11 PASS, 1 FAIL as delivered: 4 open High exceptions), status HOLD"],
                ["Tests", "18 of 18 PASS"],
                ["Sanity examples", "350k/300k = 3,200; 180k/200k = 400; 250k/200k = 2,800; 120k/100k = 1,200"]], [55, 115]))

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=18 * mm, bottomMargin=18 * mm,
                        title="SIC Intelligence Hub - Beginner's Guide", author="Synthetic prototype")
doc.build(s, onFirstPage=footer, onLaterPages=footer)
print("ok")
