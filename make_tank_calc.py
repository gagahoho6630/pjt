# -*- coding: utf-8 -*-
"""
Storage Tank Sizing Calculation Sheet generator (5 tanks)
- 원본 Tank.pdf 와 동일 양식
- 모든 값은 셀 참조 수식으로 연동
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()

# ---------- 스타일 ----------
yellow = PatternFill("solid", fgColor="FFFF00")   # 입력값
orange = PatternFill("solid", fgColor="FFC000")    # 선택/주요 결과
red    = PatternFill("solid", fgColor="FF0000")    # Net capacity
blue   = PatternFill("solid", fgColor="D9E1F2")    # 라벨
grey   = PatternFill("solid", fgColor="F2F2F2")
F_title = Font(name="맑은 고딕", size=14, bold=True)
F_h     = Font(name="맑은 고딕", size=10, bold=True)
F_n     = Font(name="맑은 고딕", size=10)
F_b     = Font(name="맑은 고딕", size=10, bold=True)
F_red   = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
L = Alignment(horizontal="left",   vertical="center")
C = Alignment(horizontal="center", vertical="center")
R = Alignment(horizontal="right",  vertical="center")
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

def put(ws, coord, val, font=F_n, align=L, fill=None, fmt=None, border=False):
    c = ws[coord]
    c.value = val
    c.font = font
    c.alignment = align
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if border: c.border = box
    return c

def widths(ws):
    for col, w in {"A":5,"B":48,"C":7,"D":11,"E":12,"F":14,"G":12,"H":9,"I":6}.items():
        ws.column_dimensions[col].width = w

# ---------- 탱크 데이터 ----------
tanks = [
    dict(sheet="Condensate",   name="Condensate Storage Tank",
         D=11.6, Vw=568,  hwf=0.46, inch=12, Hs=7.75,  Ds=11.63, Vs=822.82),
    dict(sheet="Service&Fire", name="Service & Fire water Storage Tank",
         D=24.3, Vw=6400, hwf=0.57, inch=15, Hs=16.34, Ds=24.27, Vs=7556.38),
    dict(sheet="Demi",         name="Demi. water Storage Tank",
         D=17.7, Vw=2300, hwf=0.53, inch=17, Hs=11.87, Ds=17.68, Vs=2913.57),
    dict(sheet="FuelOil",      name="Fuel Oil Storage Tank",
         D=19.3, Vw=3000, hwf=0.53, inch=20, Hs=12.84, Ds=19.32, Vs=3761.78),
    dict(sheet="Seawater",     name="Filtered Seawater Storage Tank",
         D=11.2, Vw=500,  hwf=0.45, inch=10, Hs=7.45,  Ds=11.23, Vs=737.51),
]

VFMT = "#,##0.00"
NFMT = "#,##0.0"

def build(ws, t):
    widths(ws)

    # ---- Title ----
    ws.merge_cells("A1:I1")
    put(ws, "A1", f"Calculation for {t['name']}", F_title, C)

    put(ws, "A3", "1. Design Condition", F_h)
    put(ws, "B4", "1) Tank type : Cone Roof", F_n)
    put(ws, "B5", "3) Tank ID", F_n)
    put(ws, "E5", t["D"], F_b, C, yellow, NFMT, True)   # D nominal  (input)
    put(ws, "F5", "m", F_n)

    put(ws, "A7", "2. Calculation", F_h)
    put(ws, "B8", "∴ Working Capacity (Vw) =", F_b)
    put(ws, "E8", t["Vw"], F_b, C, yellow, NFMT, True)  # Vw required (input)
    put(ws, "F8", "m³", F_n)
    put(ws, "G8", "대당 =", F_n, R)
    put(ws, "H8", "=E8", F_b, C, None, NFMT)

    put(ws, "B10", "HH = Top level - High level", F_n)
    put(ws, "B11", "Hw = High level - Low level (Working Volume)", F_n)
    put(ws, "B12", "HL = Low level - Bottom level", F_n)

    put(ws, "B14", "2) Normal Capacity (Vn)", F_b)
    put(ws, "C15", "→ Working capacity + Dead capacity", F_n)
    put(ws, "C16", "Dead capacity (V0) = (HL + HH) x D²/4", F_n)

    put(ws, "A18", "- HW", F_b)
    put(ws, "B19", "→ Seismic Zone", F_n)
    put(ws, "E19", "HW", F_n, R); put(ws, "F19", "=", F_n, C)
    put(ws, "G19", t["hwf"], F_b, C, yellow, "0.00", True)  # seismic factor (input)
    put(ws, "H19", "x D", F_n)
    put(ws, "E20", "D", F_n, R); put(ws, "F20", "=", F_n, C)
    put(ws, "G20", "=E5", F_n, C, None, NFMT); put(ws, "H20", "m", F_n)
    put(ws, "E21", "HW", F_n, R); put(ws, "F21", "=", F_n, C)
    put(ws, "G21", "=G19*E5", F_b, C, None, NFMT); put(ws, "H21", "m", F_n)   # HW = factor x D
    put(ws, "E22", "Q", F_n, R); put(ws, "F22", "=", F_n, C)
    put(ws, "G22", "=E8", F_n, C, None, NFMT); put(ws, "H22", "m³", F_n)

    # ---- HL ----
    put(ws, "A24", "- HL", F_b)
    put(ws, "B25", "1) Tank Bottom to Outlet Nozzle Center Line dim. (Regular Type)", F_n)
    put(ws, "E25", t["inch"], F_b, C, yellow, "0", True)   # nozzle size inch (input)
    put(ws, "F25", '"', F_n, C); put(ws, "G25", "=", F_n, C)
    put(ws, "H25", "=E25*25.4", F_n, C, None, "0.0"); put(ws, "I25", "mm", F_n)
    put(ws, "B26", "2) Top outlet nozzle over", F_n)
    put(ws, "G26", "=", F_n, C); put(ws, "H26", 150, F_n, C, None, "0"); put(ws, "I26", "mm", F_n)
    put(ws, "B27", "3) Outlet nozzle radius", F_n)
    put(ws, "G27", "=", F_n, C); put(ws, "H27", 1000, F_n, C, None, "0"); put(ws, "I27", "mm", F_n)
    put(ws, "C28", "HL", F_b, R); put(ws, "D28", "=", F_n, C)
    put(ws, "E28", "=H25+H26+H27", F_b, C, None, "0.0"); put(ws, "F28", "mm", F_n)
    put(ws, "G28", "Application =", F_n, R)
    put(ws, "H28", "=ROUNDUP(E28/1000,1)", F_b, C, None, "0.0"); put(ws, "I28", "m", F_n)

    # ---- HH ----
    put(ws, "A30", "- HH", F_b)
    put(ws, "B31", "Cone roof case tank top to extinguishing agent discharge outlet", F_n)
    put(ws, "H31", 300, F_n, C, None, "0"); put(ws, "I31", "mm", F_n)
    put(ws, "B32", "HH level to extinguishing agent discharge", F_n)
    put(ws, "H32", 600, F_n, C, None, "0"); put(ws, "I32", "mm", F_n)
    put(ws, "C33", "HH", F_b, R); put(ws, "D33", "=", F_n, C)
    put(ws, "E33", "=(H31+H32)/1000", F_b, C, None, "0.0"); put(ws, "F33", "m", F_n)

    # ---- Dead / Nominal / HT ----
    put(ws, "C35", "Dead Capacity =", F_n, R)
    put(ws, "E35", "=D46+F46", F_b, C, None, VFMT); put(ws, "F35", "m³", F_n)
    put(ws, "C36", "So, Nominal capacity =", F_n, R)
    put(ws, "E36", "=E8+E35", F_b, C, None, VFMT); put(ws, "F36", "m³", F_n)
    put(ws, "C37", "HT =", F_n, R)
    put(ws, "E37", "=E40", F_b, C, None, VFMT); put(ws, "F37", "m", F_n)

    # ---- Selected tank table (attach#1) ----
    put(ws, "B39", "H = selected (attach#1)", F_b)
    put(ws, "F39", "H/D :", F_b, R)
    put(ws, "G39", "=E40/E41", F_b, C, orange, "0.00", True)
    put(ws, "C40", "H", F_b, C, blue, None, True); put(ws, "D40", "=", F_n, C)
    put(ws, "E40", t["Hs"], F_b, C, yellow, VFMT, True); put(ws, "F40", "m", F_n)   # selected H
    put(ws, "C41", "D", F_b, C, blue, None, True); put(ws, "D41", "=", F_n, C)
    put(ws, "E41", t["Ds"], F_b, C, yellow, VFMT, True); put(ws, "F41", "m", F_n)   # selected D
    put(ws, "C42", "Volume", F_b, C, blue, None, True); put(ws, "D42", "=", F_n, C)
    put(ws, "E42", t["Vs"], F_b, C, orange, VFMT, True); put(ws, "F42", "m³", F_n)  # selected Volume

    # ---- HH / HL dead capacity & Net ----
    put(ws, "D44", "HH", F_b, C, blue, None, True); put(ws, "F44", "HL", F_b, C, blue, None, True)
    put(ws, "D45", "=E33", F_n, C, yellow, "0.0", True); put(ws, "E45", "m", F_n)
    put(ws, "F45", "=H28", F_n, C, yellow, "0.0", True); put(ws, "G45", "m", F_n)
    put(ws, "B46", "Dead Capacity :", F_b, C, blue, None, True)
    put(ws, "D46", "=E42*E33/E40", F_b, C, None, VFMT); put(ws, "E46", "m³", F_n)   # HH dead = V*HH/H
    put(ws, "F46", "=E42*H28/E40", F_b, C, None, VFMT); put(ws, "G46", "m³", F_n)   # HL dead = V*HL/H
    put(ws, "B47", "Design Capacity (Net Capacity) :", F_b, C, blue, None, True)
    ws.merge_cells("D47:F47")
    put(ws, "D47", "=E42-D46-F46", F_red, C, red, NFMT)   # Net = V - HHdead - HLdead
    for cc in ("D47","E47","F47"): ws[cc].border = box
    put(ws, "G47", "m³", F_n)

    # 검산: Net 이 요구 Vw 이상인지 표시
    put(ws, "B49", "Check (Net ≥ Required Vw) :", F_n, R)
    put(ws, "E49", '=IF(D47>=E8,"OK","NG")', F_b, C)


for t in tanks:
    ws = wb.active if t is tanks[0] else wb.create_sheet()
    ws.title = t["sheet"]
    build(ws, t)

# ---------- Summary 시트 ----------
ws = wb.create_sheet("Summary", 0)
widths(ws)
ws.merge_cells("A1:H1")
put(ws, "A1", "Storage Tank Sizing Summary", F_title, C)
hdr = ["No","Tank","Required Vw (m³)","Sel. D (m)","Sel. H (m)","H/D","Sel. Volume (m³)","Net Capacity (m³)","Result"]
for i, h in enumerate(hdr):
    col = chr(ord("A")+i)
    put(ws, f"{col}3", h, F_h, C, grey, None, True)
ws.column_dimensions["A"].width = 5
ws.column_dimensions["B"].width = 34
for col in "CDEFGH": ws.column_dimensions[col].width = 16
ws.column_dimensions["I"].width = 10
for j, t in enumerate(tanks):
    r = 4 + j
    s = t["sheet"]
    put(ws, f"A{r}", j+1, F_n, C, None, "0", True)
    put(ws, f"B{r}", t["name"], F_n, L, None, None, True)
    put(ws, f"C{r}", f"='{s}'!E8",  F_n, C, None, NFMT, True)
    put(ws, f"D{r}", f"='{s}'!E41", F_n, C, None, VFMT, True)
    put(ws, f"E{r}", f"='{s}'!E40", F_n, C, None, VFMT, True)
    put(ws, f"F{r}", f"='{s}'!G39", F_n, C, None, "0.00", True)
    put(ws, f"G{r}", f"='{s}'!E42", F_n, C, None, VFMT, True)
    put(ws, f"H{r}", f"='{s}'!D47", F_b, C, None, NFMT, True)
    put(ws, f"I{r}", f"='{s}'!E49", F_b, C, None, None, True)

out = r"d:\작업방\작업방\0.프로젝트\00.2026 pjt\aI 설계\Tank_Calculation.xlsx"
wb.save(out)
print("saved:", out)
