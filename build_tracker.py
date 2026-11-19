"""
Build the SYNTHETIC Client Delivery Tracker (data + formula-driven Excel workbook).

10 invented client accounts, their projects, milestones and a RAID log (Risks, Assumptions, Issues, Dependencies),
as tracked by a project coordinator across a content-services delivery portfolio. All data is random (seed 81).

Run:  python build_tracker.py
Output: Data/*.csv and Client_Delivery_Tracker.xlsx (all KPIs are Excel formulas against the 'Status date' input)
"""
import os

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

rng = np.random.default_rng(81)
HERE = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(HERE, "Data"), exist_ok=True)
STATUS_DATE = pd.Timestamp("2025-01-31")

clients = ["Northbridge Press", "Larkfield Academic", "Bluewater Journals", "Meridian Books", "Oakline Publishing",
           "Harbourview Media", "Silverpine Education", "Redstone Scientific", "Westgate Law Reports", "Kestrel Medical"]
ptypes = ["Journal production", "Book conversion (ePub)", "Backlist digitisation", "Platform migration", "Editorial services"]
phases = ["Kick-off & requirements", "Sample / pilot approved", "Workflow set-up", "Batch 1 delivered", "Batch 2 delivered",
          "Final delivery", "Client sign-off"]
owners = ["Balaji R", "Priya S", "Karthik M", "Divya N"]
risk = {c: float(rng.uniform(0, 1)) for c in clients}          # hidden account difficulty

projects, ms, raid = [], [], []
pid = mid = rid = 1
for c in clients:
    for _ in range(int(rng.integers(1, 3))):
        start = pd.Timestamp("2024-08-15") + pd.Timedelta(days=int(rng.integers(0, 120)))
        pr = f"PRJ{pid:03d}"
        projects.append({"ProjectID": pr, "Client": c, "ProjectName": f"{c.split()[0]} {rng.choice(ptypes)}", "Owner": rng.choice(owners),
                         "StartDate": start.date(), "BudgetHours": int(rng.integers(600, 4000))})
        d = start
        slip = 0
        for ph in phases:
            d = d + pd.Timedelta(days=int(rng.integers(12, 30)))
            slip += max(0, int(rng.normal(risk[c] * 5 - 3, 2.5)))
            forecast = d + pd.Timedelta(days=slip)
            actual = forecast + pd.Timedelta(days=int(rng.integers(-2, 3))) if forecast <= STATUS_DATE - pd.Timedelta(days=2) else None
            ms.append({"MilestoneID": f"MS{mid:04d}", "ProjectID": pr, "Client": c, "Milestone": ph, "Owner": projects[-1]["Owner"],
                       "PlannedDate": d.date(), "ForecastDate": forecast.date(), "ActualDate": actual.date() if actual is not None else None})
            mid += 1
        pid += 1
    for _ in range(int(rng.integers(2, 6))):
        typ = rng.choice(["Risk", "Issue", "Dependency", "Assumption"], p=[0.45, 0.3, 0.15, 0.1])
        sev = rng.choice(["High", "Medium", "Low"], p=[0.15 + 0.35 * risk[c], 0.45, 0.40 - 0.35 * risk[c]])
        raised = pd.Timestamp("2024-08-01") + pd.Timedelta(days=int(rng.integers(0, 170)))
        closed = raised + pd.Timedelta(days=int(rng.integers(5, 60)))
        raid.append({"RAIDID": f"R{rid:03d}", "Client": c, "Type": typ, "Description": rng.choice([
            "Client source files arriving late", "Style guide changes after sample approval", "Key typesetter on leave in peak month",
            "Platform API access not yet granted", "Volume higher than contracted", "Client reviewer turnaround over 5 days",
            "XML schema version not confirmed", "Assumes client provides final artwork by batch 2"]), "Severity": sev, "Owner": rng.choice(owners),
            "RaisedDate": raised.date(), "DueDate": (raised + pd.Timedelta(days=21)).date(),
            "Status": "Closed" if closed <= STATUS_DATE and rng.random() < 0.65 else "Open", "Mitigation": "See weekly status call notes"})
        rid += 1
P, M, RD = pd.DataFrame(projects), pd.DataFrame(ms), pd.DataFrame(raid)
for n, d in (("Projects", P), ("Milestones", M), ("RAID_Log", RD)):
    d.to_csv(os.path.join(HERE, "Data", f"{n}.csv"), index=False)

# ------------------------------------------------------------------ workbook
F = "Arial"
H1, B, BB = Font(name=F, size=16, bold=True), Font(name=F, size=10), Font(name=F, size=10, bold=True)
HD, BLUE, GREY = Font(name=F, size=10, bold=True, color="FFFFFF"), Font(name=F, size=10, color="0000FF"), Font(name=F, size=9, italic=True, color="666666")
HF, YEL = PatternFill("solid", fgColor="2A78D6"), PatternFill("solid", fgColor="FFFF00")
G, A, Rr = (PatternFill("solid", fgColor=c) for c in ("C6EFCE", "FFEB9C", "FFC7CE"))
BOX = Border(*(Side(style="thin", color="D9D9D9"),) * 4)
wb = Workbook()


def table(ws, r0, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), start=1):
        x = ws.cell(row=r0, column=i, value=c)
        x.font, x.fill, x.border, x.alignment = HD, HF, BOX, Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w


def rag(ws, rng_, first):
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'{first}="Green"'], fill=G))
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'{first}="Amber"'], fill=A))
    ws.conditional_formatting.add(rng_, FormulaRule(formula=[f'{first}="Red"'], fill=Rr))


ps = wb.active
ps.title = "Portfolio Summary"
ws_m = wb.create_sheet("Milestones")
ws_r = wb.create_sheet("RAID Log")
ws_p = wb.create_sheet("Projects")
ws_s = wb.create_sheet("Client Status Report")
ws_h = wb.create_sheet("How To Use")
NM, NR = len(M) + 1, len(RD) + 1

# Milestones
table(ws_m, 1, ["MilestoneID", "ProjectID", "Client", "Milestone", "Owner", "Planned", "Forecast", "Actual", "Done?", "Slip (days)",
                "Overdue?", "Due in next 14 days?", "RAG"], [11, 9, 22, 24, 11, 11, 11, 11, 7, 9, 9, 11, 8])
for r, x in enumerate(M.itertuples(index=False), start=2):
    for c, v in enumerate([x.MilestoneID, x.ProjectID, x.Client, x.Milestone, x.Owner], start=1):
        ws_m.cell(row=r, column=c, value=v).font = B
    for c, v in ((6, x.PlannedDate), (7, x.ForecastDate), (8, x.ActualDate)):
        if v is not None and not pd.isna(v):
            ws_m.cell(row=r, column=c, value=v).number_format = "yyyy-mm-dd"
    ws_m.cell(row=r, column=9, value=f'=IF(H{r}<>"","Yes","No")')
    ws_m.cell(row=r, column=10, value=f'=IF(H{r}<>"",H{r}-F{r},G{r}-F{r})')
    ws_m.cell(row=r, column=11, value=f'=IF(AND(H{r}="",F{r}<StatusDate),"Yes","No")')
    ws_m.cell(row=r, column=12, value=f'=IF(AND(H{r}="",G{r}>=StatusDate,G{r}<=StatusDate+14),"Yes","No")')
    ws_m.cell(row=r, column=13, value=f'=IF(K{r}="Yes",IF(StatusDate-F{r}>7,"Red","Amber"),IF(J{r}>5,"Amber","Green"))')
rag(ws_m, f"M2:M{NM}", "M2")
ws_m.freeze_panes, ws_m.auto_filter.ref = "A2", f"A1:M{NM}"

# RAID
table(ws_r, 1, ["RAIDID", "Client", "Type", "Description", "Severity", "Owner", "Raised", "Due", "Status", "Mitigation / next step", "Days open",
                "Past due?"], [8, 22, 11, 44, 9, 11, 11, 11, 9, 32, 9, 9])
for r, x in enumerate(RD.itertuples(index=False), start=2):
    for c, v in enumerate([x.RAIDID, x.Client, x.Type, x.Description, x.Severity, x.Owner, x.RaisedDate, x.DueDate, x.Status, x.Mitigation], start=1):
        cell = ws_r.cell(row=r, column=c, value=v)
        cell.font = BLUE if c in (9, 10) else B
        if c in (7, 8):
            cell.number_format = "yyyy-mm-dd"
    ws_r.cell(row=r, column=11, value=f'=IF(I{r}="Open",StatusDate-G{r},"")')
    ws_r.cell(row=r, column=12, value=f'=IF(AND(I{r}="Open",H{r}<StatusDate),"Yes","No")')
ws_r.conditional_formatting.add(f"E2:E{NR}", FormulaRule(formula=['E2="High"'], fill=Rr))
ws_r.conditional_formatting.add(f"L2:L{NR}", FormulaRule(formula=['L2="Yes"'], fill=Rr))
dv = DataValidation(type="list", formula1='"Open,Closed"')
ws_r.add_data_validation(dv)
dv.add(f"I2:I{NR + 50}")
ws_r.freeze_panes = "A2"

# Projects
table(ws_p, 1, ["ProjectID", "Client", "Project", "Owner", "Start", "Budget hours"], [9, 22, 36, 11, 11, 11])
for r, x in enumerate(P.itertuples(index=False), start=2):
    for c, v in enumerate(x, start=1):
        cell = ws_p.cell(row=r, column=c, value=v)
        cell.font = B
        if c == 5:
            cell.number_format = "yyyy-mm-dd"

# Portfolio summary
ps["A1"], ps["A1"].font = "Client Delivery Portfolio - Summary", H1
ps["A3"], ps["A3"].font = "Status date:", BB
ps["B3"] = STATUS_DATE.to_pydatetime()
ps["B3"].number_format, ps["B3"].font, ps["B3"].fill = "yyyy-mm-dd", BLUE, YEL
wb.defined_names["StatusDate"] = __import__("openpyxl").workbook.defined_name.DefinedName("StatusDate", attr_text="'Portfolio Summary'!$B$3")
ps["C3"], ps["C3"].font = "<- change to re-run every status on a new date", GREY
cols = ["Client", "Projects", "Milestones", "Completed", "% complete", "Completed on time % (2-day tolerance)", "Overdue", "Due next 14 days",
        "Avg slip (days)", "Open risks/issues", "High severity open", "Overall RAG"]
table(ps, 5, cols, [22, 9, 10, 10, 10, 12, 9, 10, 10, 11, 11, 10])
mr = lambda col: f"Milestones!${col}$2:${col}${NM}"
rr = lambda col: f"'RAID Log'!${col}$2:${col}${NR}"
for i, c in enumerate(clients):
    r = 6 + i
    ps.cell(row=r, column=1, value=c).font = B
    ps.cell(row=r, column=2, value=f"=COUNTIF(Projects!$B$2:$B${len(P) + 1},A{r})")
    ps.cell(row=r, column=3, value=f"=COUNTIF({mr('C')},A{r})")
    ps.cell(row=r, column=4, value=f'=COUNTIFS({mr("C")},A{r},{mr("I")},"Yes")')
    ps.cell(row=r, column=5, value=f"=IFERROR(D{r}/C{r},0)").number_format = "0%"
    ps.cell(row=r, column=6, value=f'=IFERROR(COUNTIFS({mr("C")},A{r},{mr("I")},"Yes",{mr("J")},"<=2")/D{r},"")').number_format = "0%"
    ps.cell(row=r, column=7, value=f'=COUNTIFS({mr("C")},A{r},{mr("K")},"Yes")')
    ps.cell(row=r, column=8, value=f'=COUNTIFS({mr("C")},A{r},{mr("L")},"Yes")')
    ps.cell(row=r, column=9, value=f"=AVERAGEIF({mr('C')},A{r},{mr('J')})").number_format = "0.0"
    ps.cell(row=r, column=10, value=f'=COUNTIFS({rr("B")},A{r},{rr("I")},"Open",{rr("C")},"<>Assumption")')
    ps.cell(row=r, column=11, value=f'=COUNTIFS({rr("B")},A{r},{rr("I")},"Open",{rr("E")},"High")')
    ps.cell(row=r, column=12, value=f'=IF(OR(COUNTIFS({mr("C")},A{r},{mr("M")},"Red")>0,K{r}>=2),"Red",IF(OR(G{r}>0,I{r}>5,K{r}=1),"Amber","Green"))')
    for cc in range(1, 13):
        ps.cell(row=r, column=cc).border = BOX
tr = 6 + len(clients)
ps.cell(row=tr, column=1, value="Portfolio").font = BB
for cc, f in ((2, "SUM"), (3, "SUM"), (4, "SUM"), (7, "SUM"), (8, "SUM"), (10, "SUM"), (11, "SUM")):
    ps.cell(row=tr, column=cc, value=f"={f}({get_column_letter(cc)}6:{get_column_letter(cc)}{tr - 1})").font = BB
ps.cell(row=tr, column=5, value=f"=D{tr}/C{tr}").number_format = "0%"
ps.cell(row=tr, column=6, value=f'=COUNTIFS({mr("I")},"Yes",{mr("J")},"<=2")/D{tr}').number_format = "0%"
rag(ps, f"L6:L{tr - 1}", "L6")
ps.cell(row=tr + 2, column=1, value="RAG rule: Red = any milestone overdue by more than 7 days or 2+ high-severity open items; "
        "Amber = any overdue milestone, average slip over 5 days or 1 high-severity item; otherwise Green.").font = GREY

# Client status report
ws_s["A1"], ws_s["A1"].font = "Weekly Client Status Report", H1
ws_s["A3"], ws_s["A3"].font, ws_s["B3"] = "Client:", BB, "Redstone Scientific"
ws_s["B3"].fill, ws_s["B3"].font = YEL, BLUE
dv2 = DataValidation(type="list", formula1=f"='Portfolio Summary'!$A$6:$A${tr - 1}")
ws_s.add_data_validation(dv2)
dv2.add("B3")
ws_s["A4"], ws_s["B4"] = "Status date:", "=StatusDate"
ws_s["B4"].number_format = "yyyy-mm-dd"
items = [("Overall RAG", 12, None), ("Milestones completed", 4, None), ("% complete", 5, "0%"), ("Completed on time (2-day tolerance)", 6, "0%"),
         ("Overdue milestones", 7, None), ("Due in next 14 days", 8, None), ("Average slip (days)", 9, "0.0"), ("Open risks / issues", 10, None),
         ("High severity open", 11, None)]
for i, (lab, col, fmt) in enumerate(items, start=6):
    ws_s.cell(row=i, column=1, value=lab).font = BB
    c = ws_s.cell(row=i, column=2, value=f"=INDEX('Portfolio Summary'!${get_column_letter(col)}$6:${get_column_letter(col)}${tr - 1},"
                                          f"MATCH($B$3,'Portfolio Summary'!$A$6:$A${tr - 1},0))")
    if fmt:
        c.number_format = fmt
rag(ws_s, "B6", "B6")
ws_s["A17"], ws_s["A17"].font = "Commentary (fill in each week)", BB
for i, t in enumerate(["Progress this week:", "Plan for next week:", "Decisions needed from client:", "Escalations:"], start=18):
    ws_s.cell(row=i, column=1, value=t).font = B
    ws_s.cell(row=i, column=2, value="Example: Batch 1 delivered on 28 Jan; sample feedback received." if i == 18 else "").font = BLUE
ws_s.column_dimensions["A"].width, ws_s.column_dimensions["B"].width = 30, 70

ws_h["A1"], ws_h["A1"].font = "How to use this tracker", H1
for i, t in enumerate([
        "1. Update Actual and Forecast dates in 'Milestones' (blue = input) after each client call.",
        "2. Update Status and Mitigation in 'RAID Log'. Add new rows at the bottom and copy the formulas in K:L down.",
        "3. Set the Status date (yellow, 'Portfolio Summary' B3) to the reporting date. Every RAG and count recalculates.",
        "4. 'Client Status Report': pick a client (yellow cell) for the weekly status mail, then write the commentary.",
        "Colour legend: blue text = input; yellow fill = selector; green / amber / red = RAG.",
        "Synthetic data generated by build_tracker.py (seed 81). Client names are invented."], start=3):
    ws_h.cell(row=i, column=1, value=t).font = B
ws_h.column_dimensions["A"].width = 120
for w in wb.worksheets:
    w.sheet_view.showGridLines = w.title in ("Milestones", "RAID Log", "Projects")
wb.save(os.path.join(HERE, "Client_Delivery_Tracker.xlsx"))
print(f"projects {len(P)} milestones {len(M)} raid {len(RD)}")
