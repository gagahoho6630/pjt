# -*- coding: utf-8 -*-
"""
Boiler Feed Water Pump Calculation Sheet generator
- 원본 cal.pdf 와 동일 양식
- 모든 값은 셀 참조 수식으로 연동
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ---------- 공통 스타일 ----------
title_fill   = PatternFill("solid", fgColor="FFFFFF")
section_fill = PatternFill("solid", fgColor="C6E0B4")   # 연두색 섹션 헤더
header_fill  = PatternFill("solid", fgColor="F2F2F2")
result_fill  = PatternFill("solid", fgColor="FFF2CC")   # 결과(노란색)
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
title_font   = Font(name="맑은 고딕", size=11, bold=True)
section_font = Font(name="맑은 고딕", size=10, bold=True)
normal_font  = Font(name="맑은 고딕", size=9)
result_font  = Font(name="맑은 고딕", size=9, bold=True)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",  vertical="center", wrap_text=True)
right  = Alignment(horizontal="right", vertical="center")

NUMFMT = "#,##0.000000"

def style_sheet(ws):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 16

def put(ws, r, item, desc, unit, source, value_formula, is_result=False, numfmt=NUMFMT):
    a = ws.cell(r, 1, item); a.font = normal_font; a.alignment = center; a.border = border
    b = ws.cell(r, 2, desc); b.font = normal_font; b.alignment = left; b.border = border
    c = ws.cell(r, 3, unit); c.font = normal_font; c.alignment = center; c.border = border
    d = ws.cell(r, 4, source); d.font = normal_font; d.alignment = left; d.border = border
    e = ws.cell(r, 5, value_formula)
    e.font = result_font if is_result else normal_font
    e.alignment = right; e.border = border
    if numfmt:
        e.number_format = numfmt
    if is_result:
        e.fill = result_fill
    return r

def section(ws, r, text):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    cell = ws.cell(r, 1, text)
    cell.fill = section_fill
    cell.font = section_font
    cell.alignment = left
    for col in range(1, 6):
        ws.cell(r, col).border = border
    return r

def header_row(ws, r, with_unit=True):
    labels = ["Item", "", "unit" if with_unit else "", "Source", "Value"]
    for col, lab in enumerate(labels, start=1):
        cell = ws.cell(r, col, lab)
        cell.fill = header_fill
        cell.font = Font(name="맑은 고딕", size=9, bold=True)
        cell.alignment = center
        cell.border = border
    return r

def title(ws, r, name, sub, date):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    t = ws.cell(r, 1, name); t.font = title_font; t.alignment = left
    s = ws.cell(r, 4, sub);  s.font = Font(name="맑은 고딕", size=10); s.alignment = left
    dd = ws.cell(r, 5, date); dd.font = Font(name="맑은 고딕", size=9); dd.alignment = right
    for col in range(1, 6):
        ws.cell(r, col).border = border
    return r


# ============================================================
#  SHEET 1 : Turbine Driven BFP  100% BMCR
# ============================================================
ws1 = wb.active
ws1.title = "100% BMCR (TD-BFP)"
style_sheet(ws1)

r = 1
title(ws1, r, "Turbine Driven Boiler Feed Water Pump", "100% BMCR 기준 (11/25 HBD.)", "2013.11.29"); r += 1

# ---- 입력 파라미터(공통 기준값) 영역 ----
section(ws1, r, "Basis / Input Parameters"); r += 1
P = {}  # 파라미터 셀 주소 저장
def param(ws, r, key, label, value, unit="", fmt=NUMFMT):
    ws.cell(r,1,label).font = normal_font; ws.cell(r,1).alignment=left; ws.cell(r,1).border=border
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
    ws.cell(r,4,unit).font=normal_font; ws.cell(r,4).alignment=center; ws.cell(r,4).border=border
    v = ws.cell(r,5,value); v.font=normal_font; v.alignment=right; v.border=border
    if fmt: v.number_format=fmt
    P[key] = f"$E${r}"
    return r

r = param(ws1, r, "sh_steam", "Superheater outlet steam flow (BMCR)", 875.141, "kg/s"); r+=1
r = param(ws1, r, "sec",      "Seconds per hour", 3600, "s/h", "#,##0"); r+=1
r = param(ws1, r, "ratio",    "BFP share ratio (2 x 60%)", 0.6, "-", "0.00"); r+=1
r = param(ws1, r, "misc",     "Miscellaneous (Attemp. spray) ratio", 0.025, "-", "0.000"); r+=1
r = param(ws1, r, "margin",   "Pump capacity margin (Surge5%+Wear6%)", 1.11, "-", "0.00"); r+=1
r = param(ws1, r, "sg",       "Specific gravity @ design temp", 0.874178955, "-"); r+=1
r += 1

# ---------------- Design Flow ----------------
section(ws1, r, "Design Flow"); r += 1
header_row(ws1, r); r += 1
F = {}
F["f1"]=r; put(ws1,r,"1)","Flow to boiler per BFP (2 x 60%)","kg/h",
               "Superheater outlet 기준  (875.141*3600)*0.6",
               f"={P['sh_steam']}*{P['sec']}*{P['ratio']}"); r+=1
F["f2"]=r; put(ws1,r,"2)","Miscellaneous Flow","kg/h",
               "Attemperator spray water (2.5% of 1)",
               f"=$E${F['f1']}*{P['misc']}"); r+=1
F["f3"]=r; put(ws1,r,"3)","Total flow","kg/h","1)+2)",
               f"=$E${F['f1']}+$E${F['f2']}"); r+=1
F["f4"]=r; put(ws1,r,"4)","Pump capacity incl. margin","kg/h","111% of 3)",
               f"=$E${F['f3']}*{P['margin']}", is_result=True); r+=1
F["f5"]=r; put(ws1,r,"5)","Design temp","'C","BFP suction (HBD)",191.7,numfmt="0.0"); r+=1
F["f6"]=r; put(ws1,r,"6)","Specific gravity","","",f"={P['sg']}"); r+=1
F["f7"]=r; put(ws1,r,"7)","Pump Suction flow (Q1)","m3/h","4) divide 6)",
               f"=$E${F['f4']}/($E${F['f6']}*1000)"); r+=1
F["f9"]=r+1  # interstage placeholder
F["f8"]=r; put(ws1,r,"8)","Pump Discharge flow (Q2)","m3/h","7)-9)",
               f"=$E${F['f7']}-$E${F['f9']}", is_result=True); r+=1
put(ws1,r,"9)","Interstage Normal. flow","m3/h","2) x 1.11(margin 11%) divide 6)",
               f"=$E${F['f2']}*{P['margin']}/($E${F['f6']}*1000)"); r+=1
r += 1

# ---------------- Discharge Head ----------------
section(ws1, r, "Discharge Head"); r += 1
header_row(ws1, r, with_unit=False); r += 1
D = {}
D["d1"]=r; put(ws1,r,"1)","Steam Turbine MSV inlet pressure","bar.a","HBD, BMCR 100%",250,numfmt="0.000"); r+=1
D["d2"]=r; put(ws1,r,"2)","P.D economizer inlet htr → Boiler SH outlet hdr","bar.a","Boiler data 참조",55.125,numfmt="0.000"); r+=1
D["d3"]=r; put(ws1,r,"3)","P.D Boiler SH outlet hdr → ST MSV inlet","bar.a","Assume",10.47375,numfmt="0.00000"); r+=1
D["d4"]=r; put(ws1,r,"4)","Superheater outlet Elevation","M","Boiler G.A 참조",81,numfmt="0.00"); r+=1
D["d5"]=r; put(ws1,r,"5)","Pump Center Elevation","M","EBARA BFP 견적 참조",2.05,numfmt="0.00"); r+=1
D["d6"]=r; put(ws1,r,"6)","Static head","bar.a","(4)-5))xs.g/10.2",
               f"=($E${D['d4']}-$E${D['d5']})*$E${F['f6']}/10.2"); r+=1
D["d7"]=r; put(ws1,r,"7)","Pressure Drop flow nozzle","bar.a","KEPCO E&C Comment",0.6,numfmt="0.00"); r+=1
D["d8"]=r; put(ws1,r,"8)","Pressure Drop heater 6","bar.a","Heater 견적 참조 (less than 1.7)",1.7,numfmt="0.00"); r+=1
D["d9"]=r; put(ws1,r,"9)","Pressure Drop heater 7","bar.a","Heater 견적 참조 (less than 1.7)",1.7,numfmt="0.00"); r+=1
D["d10"]=r; put(ws1,r,"10)","Pressure Drop heater 8","bar.a","Heater 견적 참조 (less than 1.7)",1.7,numfmt="0.00"); r+=1
D["d11"]=r; put(ws1,r,"11)","Pressure Drop Piping","bar.a","KEPCO E&C Comment",10,numfmt="0.00"); r+=1
D["d12"]=r; put(ws1,r,"12)","Calculated discharge P (PD)","bar.a","{1)+2)+3)+6)+7)+8)+9)+10)+11)}",
               f"=$E${D['d1']}+$E${D['d2']}+$E${D['d3']}+$E${D['d6']}+$E${D['d7']}+$E${D['d8']}+$E${D['d9']}+$E${D['d10']}+$E${D['d11']}",
               is_result=True); r+=1
r += 1

# ---------------- Suction Head ----------------
section(ws1, r, "Suction Head"); r += 1
header_row(ws1, r, with_unit=False); r += 1
S = {}
S["s1"]=r; put(ws1,r,"1)","Deaerator pressure (P1)","bar.a","HBD at BMCR",13.03,numfmt="0.000"); r+=1
S["s2"]=r; put(ws1,r,"2)","Deae. pres. fluct (PD1)","bar.a","P1 x 0.02",
               f"=$E${S['s1']}*0.02"); r+=1
S["s3"]=r; put(ws1,r,"3)","Piping & valve PD incl. suction strainer (PD2)","bar.a","Assume",0.5,numfmt="0.000"); r+=1
S["s4"]=r; put(ws1,r,"4)","Deaerator EL (E1)","m","G.A",23,numfmt="0.00"); r+=1
S["s5"]=r; put(ws1,r,"5)","Pump center EL (E2)","m","EBARA BFP 견적 참조",2.05,numfmt="0.00"); r+=1
S["s6"]=r; put(ws1,r,"6)","Static head (PH)","bar.a","(E1-E2)xS.G/10.2",
               f"=($E${S['s4']}-$E${S['s5']})*$E${F['f6']}/10.2"); r+=1
S["s7"]=r; put(ws1,r,"7)","Suction Pressure (PS)","bar.a","P1-PD1-PD2+PH",
               f"=$E${S['s1']}-$E${S['s2']}-$E${S['s3']}+$E${S['s6']}", is_result=True); r+=1
S["psat"]=r; put(ws1,r,"","포화증기압 (Sat. vapor pressure @191.7'C)","bar.a","Steam table",13.02468783,numfmt="0.0000000"); r+=1
S["s8"]=r; put(ws1,r,"8)","NPSHa","bar.a","PS - 포화증기압",
               f"=$E${S['s7']}-$E${S['psat']}"); r+=1
S["s9"]=r; put(ws1,r,"9)","NPSHa","m","bar.a to m",
               f"=$E${S['s8']}*10.2/$E${F['f6']}"); r+=1
S["s10"]=r; put(ws1,r,"10)","selected NPSHa","m","round down",
               f"=ROUNDDOWN($E${S['s9']},0)", is_result=True, numfmt="0"); r+=1
r += 1

# ---------------- Pump Head and Power ----------------
section(ws1, r, "Pump Head and Power"); r += 1
header_row(ws1, r, with_unit=False); r += 1
put(ws1,r,"1)","Total development head","bar.a","PD-PS",
    f"=$E${D['d12']}-$E${S['s7']}", is_result=True); h1=r; r+=1
put(ws1,r,"2)","Total development head","m","(1)x10.2/S.G)*1.05 (Head margin 5%, MSF)",
    f"=($E${h1}*10.2/$E${F['f6']})*1.05", is_result=True); r+=1


# ============================================================
#  공통 함수 : Start-up BFP 시트 작성
# ============================================================
def startup_sheet(ws, sub, date, df_inputs, dh_inputs, sh_inputs, has_interstage=False):
    style_sheet(ws)
    r = 1
    title(ws, r, "Start-up Boiler Feed Water Pump Calculation", sub, date); r += 1

    section(ws, r, "Basis / Input Parameters"); r += 1
    Pl = {}
    def lparam(r, key, label, value, unit="", fmt=NUMFMT):
        ws.cell(r,1,label).font=normal_font; ws.cell(r,1).alignment=left; ws.cell(r,1).border=border
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3)
        ws.cell(r,4,unit).font=normal_font; ws.cell(r,4).alignment=center; ws.cell(r,4).border=border
        v=ws.cell(r,5,value); v.font=normal_font; v.alignment=right; v.border=border
        if fmt: v.number_format=fmt
        Pl[key]=f"$E${r}"
        return r
    for key, label, val, unit, fmt in df_inputs["params"]:
        r = lparam(r, key, label, val, unit, fmt); r+=1
    r += 1

    # Design Flow
    section(ws, r, "Design Flow"); r += 1
    header_row(ws, r); r += 1
    f1=r; put(ws,r,"1)","Flow to boiler from one (1x100%) Start-up BFP","kg/h",
             df_inputs["f1_src"], df_inputs["f1_formula"](Pl)); r+=1
    f2=r; put(ws,r,"2)","Total flow","kg/h","1)", f"=$E${f1}"); r+=1
    f3=r; put(ws,r,"3)","Pump capacity incl. margin (Surge5%+Wear6%)","kg/h","111% of 2)",
             f"=$E${f2}*{Pl['margin']}", is_result=True); r+=1
    f4=r; put(ws,r,"4)","Design temp","'C",df_inputs["temp_src"], df_inputs["temp"], numfmt="0.0"); r+=1
    f5=r; put(ws,r,"5)","Specific gravity","","",f"={Pl['sg']}"); r+=1
    f6=r; put(ws,r,"6)","Pump Suction flow (Q1)","m3/h","3) divide 5)",
             f"=$E${f3}/($E${f5}*1000)"); r+=1
    f7=r; put(ws,r,"7)","Pump Discharge flow (Q2)","m3/h","6)", f"=$E${f6}", is_result=True); r+=1
    sg_cell=f"$E${f5}"
    r += 1

    # Discharge Head
    section(ws, r, "Discharge Head"); r += 1
    header_row(ws, r, with_unit=False); r += 1
    d={}
    d[1]=r; put(ws,r,"1)","Steam Turbine MSV inlet pressure","bar.a",dh_inputs["d1_src"],dh_inputs["d1"],numfmt="0.000"); r+=1
    d[2]=r; put(ws,r,"2)","P.D economizer inlet htr → Boiler SH outlet hdr","bar.a",dh_inputs["d2_src"],dh_inputs["d2"],numfmt="0.000"); r+=1
    d[3]=r; put(ws,r,"3)","P.D Boiler SH outlet hdr → ST MSV inlet","bar.a",dh_inputs["d3_src"],dh_inputs["d3"],numfmt="0.00000"); r+=1
    d[4]=r; put(ws,r,"4)","Superheater outlet Elevation","M","Boiler G.A 참조",81,numfmt="0.00"); r+=1
    d[5]=r; put(ws,r,"5)","Pump Center Elevation","M","EBARA BFP 견적 참조",2.05,numfmt="0.00"); r+=1
    d[6]=r; put(ws,r,"6)","Static head","bar.a","(4)-5))xs.g/10.2",
             f"=($E${d[4]}-$E${d[5]})*{sg_cell}/10.2"); r+=1
    d[7]=r; put(ws,r,"7)","Pressure Drop flow nozzle","bar.a",dh_inputs["d7_src"],dh_inputs["d7"],numfmt="0.000"); r+=1
    d[8]=r; put(ws,r,"8)","Pressure Drop heater 6","bar.a",dh_inputs["d8_src"],dh_inputs["d8"],numfmt="0.000"); r+=1
    d[9]=r; put(ws,r,"9)","Pressure Drop heater 7","bar.a",dh_inputs["d9_src"],dh_inputs["d9"],numfmt="0.000"); r+=1
    d[10]=r; put(ws,r,"10)","Pressure Drop heater 8","bar.a",dh_inputs["d10_src"],dh_inputs["d10"],numfmt="0.000"); r+=1
    d[11]=r; put(ws,r,"11)","Pressure Drop Piping","bar.a",dh_inputs["d11_src"],dh_inputs["d11"],numfmt="0.000"); r+=1
    d[12]=r; put(ws,r,"12)","Calculated discharge P (PD)","bar.a","{1)+2)+3)+6)+7)+8)+9)+10)+11)}",
             f"=$E${d[1]}+$E${d[2]}+$E${d[3]}+$E${d[6]}+$E${d[7]}+$E${d[8]}+$E${d[9]}+$E${d[10]}+$E${d[11]}",
             is_result=True); r+=1
    r += 1

    # Suction Head
    section(ws, r, "Suction Head"); r += 1
    header_row(ws, r, with_unit=False); r += 1
    s={}
    s[1]=r; put(ws,r,"1)","Deaerator pressure (P1)","bar.a",sh_inputs["p1_src"],sh_inputs["p1"],numfmt="0.000"); r+=1
    s[2]=r; put(ws,r,"2)","Deae. pres. fluct (PD1)","bar.a","P1 x 0.02",f"=$E${s[1]}*0.02"); r+=1
    s[3]=r; put(ws,r,"3)","Piping & valve PD incl. suction strainer (PD2)","bar.a","Assume",0.5,numfmt="0.000"); r+=1
    s[4]=r; put(ws,r,"4)","Deaerator EL (E1)","m","G.A",23,numfmt="0.00"); r+=1
    s[5]=r; put(ws,r,"5)","Pump center EL (E2)","m","EBARA BFP 견적 참조",2.05,numfmt="0.00"); r+=1
    s[6]=r; put(ws,r,"6)","Static head (PH)","bar.a","(E1-E2)xS.G/10.2",
             f"=($E${s[4]}-$E${s[5]})*{sg_cell}/10.2"); r+=1
    s[7]=r; put(ws,r,"7)","Suction Pressure (PS)","bar.a","P1-PD1-PD2+PH",
             f"=$E${s[1]}-$E${s[2]}-$E${s[3]}+$E${s[6]}", is_result=True); r+=1
    s["psat"]=r; put(ws,r,"","포화증기압 (Sat. vapor pressure)","bar.a","Steam table",sh_inputs["psat"],numfmt="0.0000000"); r+=1
    s[8]=r; put(ws,r,"8)","NPSHa","bar.a","PS - 포화증기압",f"=$E${s[7]}-$E${s['psat']}"); r+=1
    s[9]=r; put(ws,r,"9)","NPSHa","m","bar.a to m",f"=$E${s[8]}*10.2/{sg_cell}"); r+=1
    s[10]=r; put(ws,r,"10)","selected NPSHa","m","round down",
             f"=ROUNDDOWN($E${s[9]},0)", is_result=True, numfmt="0"); r+=1
    r += 1

    # Pump Head and Power
    section(ws, r, "Pump Head and Power"); r += 1
    header_row(ws, r, with_unit=False); r += 1
    put(ws,r,"1)","Total development head","bar.a","PD-PS",
        f"=$E${d[12]}-$E${s[7]}", is_result=True); h1=r; r+=1
    put(ws,r,"2)","Total development head","m","(1)x10.2/S.G)x1.05 (Head margin 5%, MSF)",
        f"=($E${h1}*10.2/{sg_cell})*1.05", is_result=True); r+=1


# ----- Sheet 2 : 95% Load parallel -----
ws2 = wb.create_sheet("95% Load (Start-up)")
df2 = {
    "params": [
        ("sh_steam","Superheater outlet steam flow (BMCR)",875.141,"kg/s",NUMFMT),
        ("sec","Seconds per hour",3600,"s/h","#,##0"),
        ("load","Start-up load ratio (35% at BMCR)",0.35,"-","0.00"),
        ("margin","Pump capacity margin (Surge5%+Wear6%)",1.11,"-","0.00"),
        ("sg","Specific gravity @ design temp",0.874178955,"-",NUMFMT),
    ],
    "f1_src":"Superheater outlet 기준(35%) at BMCR (875.141*3600*0.35)",
    "f1_formula": lambda Pl: f"={Pl['sh_steam']}*{Pl['sec']}*{Pl['load']}",
    "temp_src":"BFP suction (BMCR 100% 기준)", "temp":191.7,
}
dh2 = {
    "d1_src":"HBD","d1":250, "d2_src":"BMCR 100% P.D기준, 95%유량시 P.D","d2":45.125,
    "d3_src":"BMCR 100% P.D기준, 95%유량시 P.D","d3":8.57375,
    "d7_src":"BMCR 100% P.D기준","d7":0.6,
    "d8_src":"Heater 견적 참조 (less than 1.7)","d8":1.7,
    "d9_src":"Heater 견적 참조 (less than 1.7)","d9":1.7,
    "d10_src":"Heater 견적 참조 (less than 1.7)","d10":1.7,
    "d11_src":"BMCR 100% P.D기준, 95%유량시 P.D","d11":9.025,
}
sh2 = {"p1_src":"HBD","p1":13.03,"psat":13.02468783}
startup_sheet(ws2, "95% Load, parallel operation", "2013.11.29", df2, dh2, sh2)

# ----- Sheet 3 : 30% Load single -----
ws3 = wb.create_sheet("30% Load (Start-up)")
df3 = {
    "params": [
        ("flow_kgps","TMCR 30% steam flow",240.163,"kg/s",NUMFMT),
        ("sec","Seconds per hour",3600,"s/h","#,##0"),
        ("margin","Pump capacity margin (Surge5%+Wear6%)",1.11,"-","0.00"),
        ("sg","Specific gravity @ design temp",0.917938395,"-",NUMFMT),
    ],
    "f1_src":"HBD TMCR 30% load 기준 (240.163*3600)",
    "f1_formula": lambda Pl: f"={Pl['flow_kgps']}*{Pl['sec']}",
    "temp_src":"BFP suction (HBD MCR 30%)", "temp":149,
}
dh3 = {
    "d1_src":"HBD 30%","d1":100, "d2_src":"BMCR 100% P.D기준, 30%유량시 P.D","d2":4.5,
    "d3_src":"BMCR 100% P.D기준, 30%유량시 P.D","d3":0.855,
    "d7_src":"BMCR 100% P.D기준","d7":0.6,
    "d8_src":"BMCR 100% P.D기준, 30%유량시 P.D","d8":0.153,
    "d9_src":"BMCR 100% P.D기준, 30%유량시 P.D","d9":0.153,
    "d10_src":"BMCR 100% P.D기준, 30%유량시 P.D","d10":0.153,
    "d11_src":"BMCR 100% P.D기준, 30%유량시 P.D","d11":0.9,
}
sh3 = {"p1_src":"30% Load HBD","p1":4.64,"psat":4.634767973}
startup_sheet(ws3, "30% Load, single operation", "2013.11.29", df3, dh3, sh3)

out = r"d:\작업방\작업방\0.프로젝트\00.2026 pjt\aI 설계\BFP_Calculation.xlsx"
wb.save(out)
print("saved:", out)
