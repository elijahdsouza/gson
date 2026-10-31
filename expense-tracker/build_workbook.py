import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

OUT = "Job_Ledger_Workbook.xlsx"
N = 500  # ledger rows prepared with formulas
J = 60   # job rows prepared with formulas

F = "Arial"
BLUE = Font(name=F, color="0000FF")
BLACK = Font(name=F)
GREEN = Font(name=F, color="008000")
BOLD = Font(name=F, bold=True)
H1 = Font(name=F, bold=True, size=16, color="0D5F5A")
H2 = Font(name=F, bold=True, size=12, color="0D5F5A")
HDR = Font(name=F, bold=True, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="0D5F5A")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SOFT = PatternFill("solid", fgColor="E6F1EF")
NOTE = Font(name=F, italic=True, color="5B6A65", size=9)
thin = Side(style="thin", color="D3DBD8")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
CUR = '$#,##0;($#,##0);"-"'
CUR2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
DATE = "dd-mmm-yy"

EXPENSE_CATS = ["Venue & hire", "Catering", "Staff & contractors", "Talent & entertainment", "Equipment & AV",
                "Production & print", "Marketing & ads", "Travel & transport", "Materials & supplies",
                "Software & subscriptions", "Insurance & permits", "Fees & bank charges", "Other expense"]
INCOME_CATS = ["Client fee", "Ticket sales", "Sponsorship", "Merch & bar sales", "Grants", "Other income"]

JOBS = [
    ("EV-001", "Harbourside Spring Gala", "Event", "Coastal Arts Foundation", "Complete", "2026-09-12", "2026-09-12", 48000, 30000, 220),
    ("EV-002", "Product Launch Night", "Event", "Brightline Studio", "Complete", "2026-08-21", "2026-08-21", 22000, 15000, 120),
    ("PR-001", "Brand Refresh & Website", "Project", "Northgate Cafés", "Active", "2026-07-15", "2026-11-30", 36000, 14000, None),
    ("EV-003", "Summer Market Pop-up", "Event", "City Laneways Collective", "Quoted", "2026-10-18", "2026-10-19", 15000, 11000, 600),
]
TX = [
    ("2026-08-01", "Income", "EV-001", "Client fee", "Coastal Arts Foundation", "Deposit 50%", 26400, "Y"),
    ("2026-09-20", "Income", "EV-001", "Client fee", "Coastal Arts Foundation", "Final invoice", 26400, "Y"),
    ("2026-09-12", "Income", "EV-001", "Sponsorship", "Harbour Wines", "", 6600, "Y"),
    ("2026-08-10", "Expense", "EV-001", "Venue & hire", "Pier 3 Events", "", 9900, "Y"),
    ("2026-09-12", "Expense", "EV-001", "Catering", "Salt & Ember Catering", "220 covers", 12100, "Y"),
    ("2026-09-12", "Expense", "EV-001", "Equipment & AV", "StageLight AV", "", 4950, "Y"),
    ("2026-09-13", "Expense", "EV-001", "Staff & contractors", "Event crew x8", "", 3800, "N"),
    ("2026-09-05", "Expense", "EV-001", "Talent & entertainment", "The Tidal Quartet", "", 2420, "Y"),
    ("2026-08-28", "Expense", "EV-001", "Production & print", "Paper Kite Print", "", 1650, "Y"),
    ("2026-08-25", "Income", "EV-002", "Client fee", "Brightline Studio", "", 23100, "Y"),
    ("2026-08-05", "Expense", "EV-002", "Venue & hire", "The Warehouse Loft", "", 6050, "Y"),
    ("2026-08-21", "Expense", "EV-002", "Catering", "Grazing Co.", "", 5500, "Y"),
    ("2026-08-21", "Expense", "EV-002", "Equipment & AV", "StageLight AV", "", 3850, "Y"),
    ("2026-08-14", "Expense", "EV-002", "Marketing & ads", "Meta Ads", "", 2200, "Y"),
    ("2026-08-22", "Expense", "EV-002", "Staff & contractors", "Event crew x4", "", 1900, "N"),
    ("2026-07-20", "Income", "PR-001", "Client fee", "Northgate Cafés", "Stage 1", 13200, "Y"),
    ("2026-09-30", "Income", "PR-001", "Client fee", "Northgate Cafés", "Stage 2", 13200, "Y"),
    ("2026-08-12", "Expense", "PR-001", "Staff & contractors", "Freelance designer", "", 5500, "Y"),
    ("2026-09-03", "Expense", "PR-001", "Staff & contractors", "Web developer", "", 3300, "Y"),
    ("2026-08-20", "Expense", "PR-001", "Production & print", "Paper Kite Print", "Menus", 1320, "Y"),
    ("2026-09-15", "Expense", "PR-001", "Software & subscriptions", "Webflow + fonts", "", 880, "Y"),
    ("2026-10-02", "Expense", "EV-003", "Insurance & permits", "City Council", "Street permit", 3300, "Y"),
    ("2026-10-05", "Expense", "EV-003", "Marketing & ads", "Meta Ads", "", 2750, "Y"),
    ("2026-10-08", "Income", "EV-003", "Sponsorship", "Laneway Brewing", "", 4400, "Y"),
    ("2026-07-01", "Expense", "", "Software & subscriptions", "Xero + Adobe", "Annual", 1650, "Y"),
    ("2026-07-03", "Expense", "", "Insurance & permits", "Public liability", "Annual", 2860, "Y"),
    ("2026-08-01", "Expense", "", "Fees & bank charges", "Bank", "", 440, "Y"),
    ("2026-09-01", "Expense", "", "Travel & transport", "Fuel & parking", "", 660, "Y"),
    ("2026-10-01", "Expense", "", "Travel & transport", "Fuel & parking", "", 550, "Y"),
]
d = lambda s: dt.datetime.strptime(s, "%Y-%m-%d")

wb = Workbook()


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.alignment = HDR, HDR_FILL, Alignment(wrap_text=True, vertical="center")
        c.border = BOX
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[row].height = 32


# ---------------- Read Me ----------------
rm = wb.active
rm.title = "Read Me"
rm.column_dimensions["A"].width = 3
rm.column_dimensions["B"].width = 110
lines = [
    ("Job Ledger Workbook", H1),
    ("Expense, profit and cash analysis by event and project. All report figures are ex GST, in AUD.", BLACK),
    ("", None),
    ("How to use it", H2),
    ("1. Settings: set the GST rate, financial-year start, cash on hand, and your target margin (yellow cells).", BLACK),
    ("2. Jobs: add one row per event or project in the blue columns. Give each a unique Job ID (e.g. EV-004, PR-002).", BLACK),
    ("3. Ledger: record every income and expense line in the blue columns. Leave Job ID blank for overheads.", BLACK),
    ("4. Monthly, Categories and Dashboard update automatically. Charts are on the Dashboard tab.", BLACK),
    ("", None),
    ("Colour legend", H2),
    ("Blue text = an input you type.   Black text = a formula (don't overwrite).   Green text = pulled from another tab.", BLACK),
    ("Yellow fill = key assumption to review.", BLACK),
    ("", None),
    ("Example data", H2),
    ("Rows marked 'Example' in the Jobs and Ledger tabs are sample figures so the reports have something to show.", BLACK),
    ("Delete them (clear the blue cells) when you enter your own. The formulas stay in place for 500 ledger rows and 60 jobs.", BLACK),
    ("", None),
    ("Definitions", H2),
    ("Job contribution = job income less its direct costs. Overheads = expenses with no Job ID.", BLACK),
    ("Overhead allocation spreads overheads across jobs by share of income, giving a fully loaded net profit per job.", BLACK),
    ("Break-even income = overheads ÷ contribution margin %. Runway = cash on hand ÷ average monthly net burn.", BLACK),
]
for i, (t, f) in enumerate(lines, 2):
    c = rm.cell(row=i, column=2, value=t)
    if f: c.font = f

# ---------------- Settings ----------------
st = wb.create_sheet("Settings")
st.column_dimensions["A"].width = 34
st.column_dimensions["B"].width = 16
st.column_dimensions["C"].width = 60
st["A1"] = "Settings"; st["A1"].font = H1
rows = [
    ("GST rate", 0.10, PCT, "Australian GST. Source: ATO standard rate."),
    ("Financial year start", d("2026-07-01"), DATE, "AU financial year runs 1 July – 30 June."),
    ("Cash on hand", 25000, CUR, "Assumption: example figure. Replace with your bank balance."),
    ("Target job contribution margin", 0.35, PCT, "Assumption: common benchmark for event and creative services."),
    ("Target net margin", 0.15, PCT, "Assumption: healthy SME net margin target."),
]
for i, (lab, val, fmt, note) in enumerate(rows, 3):
    st.cell(row=i, column=1, value=lab).font = BOLD
    c = st.cell(row=i, column=2, value=val); c.font = BLUE; c.fill = INPUT_FILL; c.number_format = fmt; c.border = BOX
    st.cell(row=i, column=3, value=note).font = NOTE
GST, FYS, CASH, TGT_CM, TGT_NM = "Settings!$B$3", "Settings!$B$4", "Settings!$B$5", "Settings!$B$6", "Settings!$B$7"

# lists for validation
st["E2"] = "Expense categories"; st["E2"].font = BOLD
st["F2"] = "Income categories"; st["F2"].font = BOLD
st.column_dimensions["E"].width = 26; st.column_dimensions["F"].width = 22
for i, c in enumerate(EXPENSE_CATS, 3): st.cell(row=i, column=5, value=c).font = BLUE
for i, c in enumerate(INCOME_CATS, 3): st.cell(row=i, column=6, value=c).font = BLUE
st["G2"] = "All categories"; st["G2"].font = BOLD; st.column_dimensions["G"].width = 26
for i, c in enumerate(EXPENSE_CATS + INCOME_CATS, 3): st.cell(row=i, column=7, value=c).font = GREEN

# ---------------- Jobs ----------------
jb = wb.create_sheet("Jobs")
cols = ["Job ID", "Name", "Type", "Client", "Status", "Start", "End", "Quoted income (ex GST)", "Cost budget (ex GST)", "Attendees / units", "Note",
        "Income (ex GST)", "Direct costs (ex GST)", "Contribution", "Contribution margin", "Budget used", "Income vs quote",
        "Contribution per unit", "Overhead allocated", "Net profit (loaded)", "Net margin", "Flag"]
header(jb, 1, cols, [9, 28, 9, 24, 10, 11, 11, 14, 14, 11, 10, 14, 14, 14, 12, 10, 13, 13, 13, 14, 10, 22])
jb.freeze_panes = "C2"
for r in range(2, J + 2):
    if r - 2 < len(JOBS):
        jid, name, typ, client, status, s, e, q, b, u = JOBS[r - 2]
        vals = [jid, name, typ, client, status, d(s), d(e), q, b, u, "Example"]
    else:
        vals = [None] * 11
    for ci, v in enumerate(vals, 1):
        c = jb.cell(row=r, column=ci, value=v); c.font = BLUE
        if ci in (6, 7): c.number_format = DATE
        if ci in (8, 9): c.number_format = CUR
    A = f"$A{r}"
    f = {
        12: f'=IF({A}="","",SUMIFS(Ledger!$K:$K,Ledger!$C:$C,{A},Ledger!$B:$B,"Income"))',
        13: f'=IF({A}="","",SUMIFS(Ledger!$K:$K,Ledger!$C:$C,{A},Ledger!$B:$B,"Expense"))',
        14: f'=IF({A}="","",L{r}-M{r})',
        15: f'=IF(OR({A}="",N(L{r})=0),"",N{r}/L{r})',
        16: f'=IF(OR({A}="",N(I{r})=0),"",M{r}/I{r})',
        17: f'=IF(OR({A}="",N(H{r})=0),"",L{r}-H{r})',
        18: f'=IF(OR({A}="",N(J{r})=0),"",N{r}/J{r})',
        19: f'=IF({A}="","",IF(Dashboard!$C$6=0,0,L{r}/Dashboard!$C$6*Dashboard!$C$9))',
        20: f'=IF({A}="","",N{r}-S{r})',
        21: f'=IF(OR({A}="",N(L{r})=0),"",T{r}/L{r})',
        22: f'=IF({A}="","",IF(N(P{r})>1,"Over budget",IF(N(N{r})<0,"Loss so far",IF(AND(L{r}>0,N(O{r})<{TGT_CM}),"Below target margin","OK"))))',
    }
    fmts = {12: CUR, 13: CUR, 14: CUR, 15: PCT, 16: PCT, 17: CUR, 18: CUR2, 19: CUR, 20: CUR, 21: PCT}
    for ci, fx in f.items():
        c = jb.cell(row=r, column=ci, value=fx); c.font = BLACK
        if ci in fmts: c.number_format = fmts[ci]
dv = DataValidation(type="list", formula1='"Event,Project"', allow_blank=True); jb.add_data_validation(dv); dv.add(f"C2:C{J+1}")
dv = DataValidation(type="list", formula1='"Quoted,Active,Complete,Cancelled"', allow_blank=True); jb.add_data_validation(dv); dv.add(f"E2:E{J+1}")
jb.conditional_formatting.add(f"V2:V{J+1}", CellIsRule(operator="equal", formula=['"Over budget"'], font=Font(name=F, color="B2452F", bold=True)))
jb.conditional_formatting.add(f"V2:V{J+1}", CellIsRule(operator="equal", formula=['"Below target margin"'], font=Font(name=F, color="9A6A00", bold=True)))
jb.conditional_formatting.add(f"V2:V{J+1}", CellIsRule(operator="equal", formula=['"Loss so far"'], font=Font(name=F, color="B2452F", bold=True)))
jb.conditional_formatting.add(f"N2:N{J+1}", DataBarRule(start_type="min", end_type="max", color="2B8A6A"))
jb.cell(row=1, column=19).comment = Comment("Overheads (expenses with no Job ID) spread across jobs by share of total job income.", "Job Ledger")

# ---------------- Ledger ----------------
lg = wb.create_sheet("Ledger")
cols = ["Date", "Type", "Job ID", "Category", "Supplier / payer", "Note", "Amount (incl GST)", "GST included? (Y/N)",
        "Job type", "GST", "Amount (ex GST)", "Month", "Signed (ex GST)"]
header(lg, 1, cols, [11, 10, 9, 24, 26, 18, 14, 10, 11, 11, 14, 10, 14])
lg.freeze_panes = "A2"
for r in range(2, N + 2):
    if r - 2 < len(TX):
        ds, ty, jid, cat, party, note, amt, g = TX[r - 2]
        vals = [d(ds), ty, jid or None, cat, party, ("Example" + (f" – {note}" if note else "")), amt, g]
    else:
        vals = [None] * 8
    for ci, v in enumerate(vals, 1):
        c = lg.cell(row=r, column=ci, value=v); c.font = BLUE
        if ci == 1: c.number_format = DATE
        if ci == 7: c.number_format = CUR2
    f = {
        9: f'=IF(A{r}="","",IF(C{r}="","Overhead",IFERROR(INDEX(Jobs!$C:$C,MATCH(C{r},Jobs!$A:$A,0)),"Unknown job")))',
        10: f'=IF(A{r}="","",IF(UPPER(H{r})="N",0,G{r}*{GST}/(1+{GST})))',
        11: f'=IF(A{r}="","",G{r}-J{r})',
        12: f'=IF(A{r}="","",DATE(YEAR(A{r}),MONTH(A{r}),1))',
        13: f'=IF(A{r}="","",IF(B{r}="Income",K{r},-K{r}))',
    }
    for ci, fx in f.items():
        c = lg.cell(row=r, column=ci, value=fx); c.font = BLACK
        c.number_format = {9: "General", 10: CUR2, 11: CUR2, 12: "mmm-yy", 13: CUR2}[ci]
dv = DataValidation(type="list", formula1='"Income,Expense"', allow_blank=True); lg.add_data_validation(dv); dv.add(f"B2:B{N+1}")
dv = DataValidation(type="list", formula1=f"=Jobs!$A$2:$A${J+1}", allow_blank=True); lg.add_data_validation(dv); dv.add(f"C2:C{N+1}")
dv = DataValidation(type="list", formula1=f"=Settings!$G$3:$G${2+len(EXPENSE_CATS)+len(INCOME_CATS)}", allow_blank=True); lg.add_data_validation(dv); dv.add(f"D2:D{N+1}")
dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True); lg.add_data_validation(dv); dv.add(f"H2:H{N+1}")
lg.conditional_formatting.add(f"I2:I{N+1}", CellIsRule(operator="equal", formula=['"Unknown job"'], font=Font(name=F, color="B2452F", bold=True)))
lg.auto_filter.ref = f"A1:M{N+1}"

# ---------------- Monthly ----------------
mo = wb.create_sheet("Monthly")
mo["A1"] = "Monthly profit & loss and GST"; mo["A1"].font = H1
mo["A2"] = "Financial year starting"; mo["C2"] = f"={FYS}"; mo["C2"].font = GREEN; mo["C2"].number_format = DATE
cols = ["Month", "Income", "Direct job costs", "Overheads", "Total expenses", "Net profit", "Cumulative profit", "Net margin",
        "GST collected", "GST credits", "Net GST (owing)"]
header(mo, 4, cols, [12, 13, 14, 13, 14, 13, 15, 11, 13, 13, 14])
for i in range(12):
    r = 5 + i
    mo.cell(row=r, column=1, value=f"=EDATE($C$2,{i})").number_format = "mmm-yy"
    M = f"$A{r}"
    f = {
        2: f'=SUMIFS(Ledger!$K:$K,Ledger!$L:$L,{M},Ledger!$B:$B,"Income")',
        3: f'=SUMIFS(Ledger!$K:$K,Ledger!$L:$L,{M},Ledger!$B:$B,"Expense",Ledger!$I:$I,"<>Overhead")',
        4: f'=SUMIFS(Ledger!$K:$K,Ledger!$L:$L,{M},Ledger!$B:$B,"Expense",Ledger!$I:$I,"Overhead")',
        5: f"=C{r}+D{r}",
        6: f"=B{r}-E{r}",
        7: f"=F{r}" if i == 0 else f"=G{r-1}+F{r}",
        8: f'=IF(B{r}=0,"",F{r}/B{r})',
        9: f'=SUMIFS(Ledger!$J:$J,Ledger!$L:$L,{M},Ledger!$B:$B,"Income")',
        10: f'=SUMIFS(Ledger!$J:$J,Ledger!$L:$L,{M},Ledger!$B:$B,"Expense")',
        11: f"=I{r}-J{r}",
    }
    for ci, fx in f.items():
        c = mo.cell(row=r, column=ci, value=fx); c.number_format = PCT if ci == 8 else CUR
tr = 17
mo.cell(row=tr, column=1, value="FY total").font = BOLD
for ci in range(2, 12):
    col = L(ci)
    if ci == 7: fx = f"=G16"
    elif ci == 8: fx = f'=IF(B{tr}=0,"",F{tr}/B{tr})'
    else: fx = f"=SUM({col}5:{col}16)"
    c = mo.cell(row=tr, column=ci, value=fx); c.font = BOLD; c.number_format = PCT if ci == 8 else CUR; c.fill = SOFT
mo.cell(row=19, column=1, value="BAS quarters").font = H2
header_r = 20
for ci, lab in enumerate(["Quarter", "GST collected (1A)", "GST credits (1B)", "Net GST payable"], 1):
    c = mo.cell(row=header_r, column=ci, value=lab); c.font = HDR; c.fill = HDR_FILL
for q in range(4):
    r = 21 + q
    a, b = 5 + q * 3, 7 + q * 3
    mo.cell(row=r, column=1, value=f'="Q{q+1} "&TEXT(A{a},"mmm")&"–"&TEXT(A{b},"mmm yy")')
    mo.cell(row=r, column=2, value=f"=SUM(I{a}:I{b})").number_format = CUR
    mo.cell(row=r, column=3, value=f"=SUM(J{a}:J{b})").number_format = CUR
    mo.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = CUR
mo.cell(row=25, column=1, value="Labels G1A/1B follow the ATO BAS form. Check with your accountant before lodging.").font = NOTE

# ---------------- Categories ----------------
ca = wb.create_sheet("Categories")
ca["A1"] = "Spend and income by category"; ca["A1"].font = H1
header(ca, 3, ["Expense category", "Total (ex GST)", "Share of spend", "On jobs", "Overhead"], [26, 14, 13, 13, 13])
for i, c in enumerate(EXPENSE_CATS):
    r = 4 + i
    ca.cell(row=r, column=1, value=f"=Settings!E{3+i}").font = GREEN
    ca.cell(row=r, column=2, value=f'=SUMIFS(Ledger!$K:$K,Ledger!$D:$D,A{r},Ledger!$B:$B,"Expense")').number_format = CUR
    ca.cell(row=r, column=3, value=f'=IF($B${4+len(EXPENSE_CATS)}=0,"",B{r}/$B${4+len(EXPENSE_CATS)})').number_format = PCT
    ca.cell(row=r, column=4, value=f'=SUMIFS(Ledger!$K:$K,Ledger!$D:$D,A{r},Ledger!$B:$B,"Expense",Ledger!$I:$I,"<>Overhead")').number_format = CUR
    ca.cell(row=r, column=5, value=f"=B{r}-D{r}").number_format = CUR
etot = 4 + len(EXPENSE_CATS)
ca.cell(row=etot, column=1, value="Total expenses").font = BOLD
for ci in (2, 4, 5):
    c = ca.cell(row=etot, column=ci, value=f"=SUM({L(ci)}4:{L(ci)}{etot-1})"); c.font = BOLD; c.number_format = CUR; c.fill = SOFT
ca.conditional_formatting.add(f"B4:B{etot-1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="D0795F"))
ist = etot + 3
header(ca, ist, ["Income category", "Total (ex GST)", "Share of income"])
for i, c in enumerate(INCOME_CATS):
    r = ist + 1 + i
    ca.cell(row=r, column=1, value=f"=Settings!F{3+i}").font = GREEN
    ca.cell(row=r, column=2, value=f'=SUMIFS(Ledger!$K:$K,Ledger!$D:$D,A{r},Ledger!$B:$B,"Income")').number_format = CUR
    ca.cell(row=r, column=3, value=f'=IF($B${ist+1+len(INCOME_CATS)}=0,"",B{r}/$B${ist+1+len(INCOME_CATS)})').number_format = PCT
itot = ist + 1 + len(INCOME_CATS)
ca.cell(row=itot, column=1, value="Total income").font = BOLD
c = ca.cell(row=itot, column=2, value=f"=SUM(B{ist+1}:B{itot-1})"); c.font = BOLD; c.number_format = CUR; c.fill = SOFT

# ---------------- Dashboard ----------------
db = wb.create_sheet("Dashboard", 1)
db.column_dimensions["A"].width = 2
db.column_dimensions["B"].width = 34
db.column_dimensions["C"].width = 15
db.column_dimensions["D"].width = 15
db.column_dimensions["E"].width = 15
db.column_dimensions["F"].width = 46
db["B1"] = "CFO dashboard"; db["B1"].font = H1
db["B2"] = "All figures ex GST, AUD, from the Ledger and Jobs tabs. Charts are below and to the right."; db["B2"].font = NOTE
db["B4"] = "Profit & loss"; db["B4"].font = H2
JL = f"Jobs!$L$2:$L${J+1}"; JM = f"Jobs!$M$2:$M${J+1}"; JN = f"Jobs!$N$2:$N${J+1}"
kp = [
    (6, "Total income", '=SUMIFS(Ledger!$K:$K,Ledger!$B:$B,"Income")', CUR, ""),
    (7, "Direct job costs", '=SUMIFS(Ledger!$K:$K,Ledger!$B:$B,"Expense",Ledger!$I:$I,"<>Overhead")', CUR, ""),
    (8, "Job contribution", "=C6-C7", CUR, ""),
    (9, "Overheads", '=SUMIFS(Ledger!$K:$K,Ledger!$B:$B,"Expense",Ledger!$I:$I,"Overhead")', CUR, ""),
    (10, "Net profit", "=C8-C9", CUR, ""),
    (11, "Contribution margin", '=IF(C6=0,"",C8/C6)', PCT, f'=IF(C11="","",IF(C11>={TGT_CM},"On target","Below target of "&TEXT({TGT_CM},"0%")))'),
    (12, "Net margin", '=IF(C6=0,"",C10/C6)', PCT, f'=IF(C12="","",IF(C12>={TGT_NM},"On target","Below target of "&TEXT({TGT_NM},"0%")))'),
]
db["C5"] = "Value"; db["F5"] = "Read-out"
for c in ("C5", "F5"): db[c].font = BOLD
for r, lab, fx, fmt, note in kp:
    db.cell(row=r, column=2, value=lab).font = BOLD if r in (8, 10) else BLACK
    c = db.cell(row=r, column=3, value=fx); c.number_format = fmt; c.border = BOX
    if r in (8, 10): c.font = BOLD; c.fill = SOFT
    if note: db.cell(row=r, column=6, value=note)

db["B14"] = "Health metrics"; db["B14"].font = H2
hm = [
    (15, "Overhead coverage (contribution ÷ overheads)", '=IF(C9=0,"",C8/C9)', '0.0"x"',
     '=IF(C15="","",IF(C15>=3,"Strong: jobs cover overheads 3x+",IF(C15>=1.5,"Adequate","Thin: overheads eat most job profit")))'),
    (16, "Break-even income", '=IF(OR(C11="",C11<=0),"",C9/C11)', CUR, '="Income needed to cover overheads at current margin"'),
    (17, "Months of data", '=MAX(1,SUMPRODUCT(--(((Monthly!$B$5:$B$16<>0)+(Monthly!$E$5:$E$16<>0))>0)))', "0", ""),
    (18, "Average monthly overhead", "=C9/C17", CUR, ""),
    (19, "Average monthly net profit", "=C10/C17", CUR, ""),
    (20, "Cash on hand", f"={CASH}", CUR, '="From Settings"'),
    (21, "Months of overhead covered by cash", '=IF(C18=0,"",C20/C18)', '0.0" mo"',
     '=IF(C21="","",IF(C21>=6,"Healthy buffer (6+ months)",IF(C21>=3,"OK: aim for 6 months","Low: build a cash reserve")))'),
    (22, "Largest client share of income", f'=IF(C6=0,"",MAX(Clients!$C$4:$C$63)/C6)', PCT,
     '=IF(C22="","",IF(C22>0.4,"Concentration risk: one client > 40%","Diversified"))'),
    (23, "Jobs over budget", f'=COUNTIF(Jobs!$V$2:$V${J+1},"Over budget")', "0", ""),
    (24, "Jobs below target margin or at a loss", f'=COUNTIF(Jobs!$V$2:$V${J+1},"Below target margin")+COUNTIF(Jobs!$V$2:$V${J+1},"Loss so far")', "0", ""),
    (25, "GST owing (net, year to date)", "=Monthly!K17", CUR, '="Set aside in a separate account for BAS"'),
]
for r, lab, fx, fmt, note in hm:
    db.cell(row=r, column=2, value=lab)
    c = db.cell(row=r, column=3, value=fx); c.number_format = fmt; c.border = BOX
    if r == 20 or r == 25: c.font = GREEN
    if note: db.cell(row=r, column=6, value=note)
db["C20"].fill = INPUT_FILL

db["B27"] = "Events vs projects"; db["B27"].font = H2
for ci, lab in enumerate(["", "Events", "Projects", "Overheads"], 2):
    c = db.cell(row=28, column=ci, value=lab or "Metric"); c.font = HDR; c.fill = HDR_FILL
evp = [
    ("Income", '=SUMIFS(Ledger!$K:$K,Ledger!$B:$B,"Income",Ledger!$I:$I,"{t}")'),
    ("Direct costs", '=SUMIFS(Ledger!$K:$K,Ledger!$B:$B,"Expense",Ledger!$I:$I,"{t}")'),
]
for i, (lab, tpl) in enumerate(evp):
    r = 29 + i
    db.cell(row=r, column=2, value=lab)
    for ci, t in ((3, "Event"), (4, "Project"), (5, "Overhead")):
        db.cell(row=r, column=ci, value=tpl.format(t=t)).number_format = CUR
db.cell(row=31, column=2, value="Contribution").font = BOLD
for ci in (3, 4, 5):
    col = L(ci)
    c = db.cell(row=31, column=ci, value=f"={col}29-{col}30"); c.number_format = CUR; c.font = BOLD; c.fill = SOFT
db.cell(row=32, column=2, value="Margin")
for ci in (3, 4):
    col = L(ci); db.cell(row=32, column=ci, value=f'=IF({col}29=0,"",{col}31/{col}29)').number_format = PCT
db.cell(row=33, column=2, value="Number of jobs")
db.cell(row=33, column=3, value=f'=COUNTIF(Jobs!$C$2:$C${J+1},"Event")')
db.cell(row=33, column=4, value=f'=COUNTIF(Jobs!$C$2:$C${J+1},"Project")')
db.cell(row=34, column=2, value="Average contribution per job")
for ci in (3, 4):
    col = L(ci); db.cell(row=34, column=ci, value=f'=IF({col}33=0,"",{col}31/{col}33)').number_format = CUR
db.cell(row=35, column=2, value="Contribution per attendee (events)")
db.cell(row=35, column=3, value=f'=IF(SUMIFS(Jobs!$J$2:$J${J+1},Jobs!$C$2:$C${J+1},"Event")=0,"",C31/SUMIFS(Jobs!$J$2:$J${J+1},Jobs!$C$2:$C${J+1},"Event"))').number_format = CUR2

db["B37"] = "CFO notes"; db["B37"].font = H2
notes = [
    '=IF(C11="","Add income to see margin commentary.",IF(C11<'+TGT_CM+',"Contribution margin is "&TEXT(C11,"0%")&", under the "&TEXT('+TGT_CM+',"0%")&" target. Review pricing or supplier costs on the jobs flagged in the Jobs tab.","Contribution margin of "&TEXT(C11,"0%")&" is on target. Protect it by quoting from actual cost data."))',
    '=IF(OR(C32="",C33="",D32=""),"Compare events and projects once both have income.",IF(C32>D32,"Events earn a higher margin ("&TEXT(C32,"0%")&" vs "&TEXT(D32,"0%")&"). Consider leading with events and pricing projects up.","Projects earn a higher margin ("&TEXT(D32,"0%")&" vs "&TEXT(C32,"0%")&"). Recurring project work may steady cash flow."))',
    '=IF(C22="","",IF(C22>0.4,"One client supplies "&TEXT(C22,"0%")&" of income. Losing them would hurt: widen the client base or secure a retainer.","No single client dominates income ("&TEXT(C22,"0%")&" max)."))',
    '=IF(C21="","",IF(C21<3,"Cash covers only "&TEXT(C21,"0.0")&" months of overhead. Ask for 50% deposits up front and chase final invoices quickly.","Cash covers "&TEXT(C21,"0.0")&" months of overhead."))',
    '="Hold back "&TEXT(C25,"$#,##0")&" for GST so it is not mistaken for profit."',
]
for i, n in enumerate(notes):
    c = db.cell(row=38 + i, column=2, value=n); c.alignment = Alignment(wrap_text=True, vertical="top")
    db.merge_cells(start_row=38 + i, start_column=2, end_row=38 + i, end_column=6)
    db.row_dimensions[38 + i].height = 32

# ---------------- Clients (helper) ----------------
cl = wb.create_sheet("Clients")
cl["A1"] = "Income by client"; cl["A1"].font = H1
cl["A2"] = "Clients are taken from the Jobs tab; income comes from the Ledger via Job ID."; cl["A2"].font = NOTE
header(cl, 3, ["Job ID", "Client", "Income (ex GST)", "Share"], [10, 30, 15, 10])
for i in range(J):
    r = 4 + i
    jr = 2 + i
    cl.cell(row=r, column=1, value=f'=IF(Jobs!A{jr}="","",Jobs!A{jr})').font = GREEN
    cl.cell(row=r, column=2, value=f'=IF(Jobs!A{jr}="","",Jobs!D{jr})').font = GREEN
    # income across all jobs for the same client, counted only on the first job row of that client
    cl.cell(row=r, column=3, value=f'=IF(B{r}="","",IF(COUNTIF($B$4:B{r},B{r})>1,0,SUMIFS(Jobs!$L$2:$L${J+1},Jobs!$D$2:$D${J+1},B{r})))').number_format = CUR
    cl.cell(row=r, column=4, value=f'=IF(OR(B{r}="",N(C{r})=0,Dashboard!$C$6=0),"",C{r}/Dashboard!$C$6)').number_format = PCT

# ---------------- Charts on Dashboard ----------------
ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.title = "Monthly income vs expenses"
ch.y_axis.title = "AUD ex GST"; ch.y_axis.numFmt = "$#,##0"; ch.height = 8; ch.width = 18
ch.add_data(Reference(mo, min_col=2, min_row=4, max_row=16), titles_from_data=True)
ch.add_data(Reference(mo, min_col=5, min_row=4, max_row=16), titles_from_data=True)
ch.set_categories(Reference(mo, min_col=1, min_row=5, max_row=16))
ch.series[0].graphicalProperties.solidFill = "2B8A6A"; ch.series[1].graphicalProperties.solidFill = "D0795F"
ln = LineChart(); ln.add_data(Reference(mo, min_col=7, min_row=4, max_row=16), titles_from_data=True)
ln.series[0].graphicalProperties.line.solidFill = "17221F"; ln.series[0].graphicalProperties.line.width = 28000
ln.y_axis.axId = 200; ln.y_axis.title = "Cumulative profit"; ln.y_axis.crosses = "max"; ln.y_axis.numFmt = "$#,##0"
ch += ln
db.add_chart(ch, "H4")

jc = BarChart(); jc.type = "bar"; jc.title = "Contribution by job"; jc.height = 8; jc.width = 18
jc.add_data(Reference(jb, min_col=14, min_row=1, max_row=1 + len(JOBS)), titles_from_data=True)
jc.add_data(Reference(jb, min_col=20, min_row=1, max_row=1 + len(JOBS)), titles_from_data=True)
jc.set_categories(Reference(jb, min_col=2, min_row=2, max_row=1 + len(JOBS)))
jc.series[0].graphicalProperties.solidFill = "0D5F5A"; jc.series[1].graphicalProperties.solidFill = "9AC9C2"
jc.x_axis.numFmt = "$#,##0"
db.add_chart(jc, "H21")

pc = PieChart(); pc.title = "Where the money goes"; pc.height = 8; pc.width = 12
pc.add_data(Reference(ca, min_col=2, min_row=3, max_row=etot - 1), titles_from_data=True)
pc.set_categories(Reference(ca, min_col=1, min_row=4, max_row=etot - 1))
pc.dataLabels = DataLabelList(); pc.dataLabels.showPercent = True
db.add_chart(pc, "H38")

ec = BarChart(); ec.type = "col"; ec.title = "Events vs projects"; ec.height = 8; ec.width = 12
ec.add_data(Reference(db, min_col=3, max_col=4, min_row=28, max_row=31), titles_from_data=True, from_rows=False)
ec.set_categories(Reference(db, min_col=2, min_row=29, max_row=31))
ec.y_axis.numFmt = "$#,##0"
db.add_chart(ec, "B45")

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.font and c.font.name != F:
                c.font = Font(name=F, bold=c.font.bold, italic=c.font.italic, size=c.font.size, color=c.font.color)
    ws.sheet_view.showGridLines = ws.title not in ("Dashboard", "Read Me")
wb.move_sheet("Clients", offset=0)
wb.active = 1
wb.save(OUT)
print("saved", OUT)
