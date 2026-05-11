# -*- coding: utf-8 -*-
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── 페이지 여백 ──────────────────────────────────────────
section = doc.sections[0]
section.page_width  = Cm(21.0)
section.page_height = Cm(29.7)
section.top_margin    = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin   = Cm(2.5)
section.right_margin  = Cm(2.5)

# ── 기본 스타일 ──────────────────────────────────────────
normal_style = doc.styles['Normal']
normal_style.font.name = '맑은 고딕'
normal_style.font.size = Pt(10)

def set_font(run, bold=False, size=10, color=None):
    run.font.name = '맑은 고딕'
    run.font.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)

def heading1(text):
    p = doc.add_heading(text, level=1)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = '맑은 고딕'
        run.font.size = Pt(14)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
    return p

def heading2(text):
    p = doc.add_heading(text, level=2)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = '맑은 고딕'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    return p

def heading3(text):
    p = doc.add_heading(text, level=3)
    for run in p.runs:
        run.font.name = '맑은 고딕'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0x20, 0x60, 0x40)
    return p

def para(text, bold=False, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run(text)
    set_font(run, bold=bold)
    return p

def code_block(lines):
    """회색 배경 코드 박스"""
    for line in lines:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent  = Cm(0.5)
        p.paragraph_format.right_indent = Cm(0.5)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(0)
        run = p.add_run(line)
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        # 배경색 (XML 직접)
        pPr = p._p.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'F2F2F2')
        pPr.append(shd)

def add_table(headers, rows, col_widths=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # 헤더 행
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell.paragraphs[0].runs[0]
        run.font.name = '맑은 고딕'
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1F497D')
        tcPr.append(shd)

    # 데이터 행
    for ri, row_data in enumerate(rows):
        row = table.rows[ri + 1]
        fill = 'FFFFFF' if ri % 2 == 0 else 'EBF3FB'
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = str(val)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = cell.paragraphs[0].runs[0]
            run.font.name = '맑은 고딕'
            run.font.size = Pt(9)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill)
            tcPr.append(shd)

    # 열 너비
    if col_widths:
        for row in table.rows:
            for ci, w in enumerate(col_widths):
                row.cells[ci].width = Cm(w)
    return table

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.8)
    run = p.add_run('※ ' + text)
    run.font.name = '맑은 고딕'
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x80, 0x00, 0x00)
    return p

def warning_box(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run('▶ ' + text)
    run.font.name = '맑은 고딕'
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'FFE0E0')
    pPr.append(shd)

# ════════════════════════════════════════════════════════════
#  표지
# ════════════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()

title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title_p.add_run('수봉변(Water Seal) 설계 계산서')
run.font.name = '맑은 고딕'
run.font.bold = True
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

doc.add_paragraph()

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = sub_p.add_run('제철소 COG (Coke Oven Gas) 주배관 적용')
run.font.name = '맑은 고딕'
run.font.size = Pt(13)
run.font.color.rgb = RGBColor(0x40, 0x40, 0x40)

doc.add_paragraph()

info_table = doc.add_table(rows=5, cols=2)
info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
info_table.style = 'Table Grid'
info_data = [
    ('문서번호', 'WS-COG-2024-001'),
    ('작성일', '2026-05-11'),
    ('적용 설비', '제철소 COG 주배관 수봉변'),
    ('적용 기준', 'API 521, 사내 기준'),
    ('작성', 'Process / Piping Engineer'),
]
for i, (k, v) in enumerate(info_data):
    row = info_table.rows[i]
    row.cells[0].text = k
    row.cells[1].text = v
    for ci in range(2):
        run = row.cells[ci].paragraphs[0].runs[0]
        run.font.name = '맑은 고딕'
        run.font.size = Pt(10)
        run.font.bold = (ci == 0)
        row.cells[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

doc.add_page_break()

# ════════════════════════════════════════════════════════════
#  1. 설계 기본 조건
# ════════════════════════════════════════════════════════════
heading1('1. 설계 기본 조건 (Design Basis)')
add_table(
    ['항목', '값', '단위'],
    [
        ('유체', 'COG (Coke Oven Gas)', '—'),
        ('운전 온도', '70', '°C'),
        ('운전 압력', '1,400', 'mmAq (g)'),
        ('운전 압력 (환산)', '13,734', 'Pa (g)'),
        ('배관 내경', '2,400', 'mm'),
        ('설계 유량', '80,000', 'Nm³/hr'),
        ('설치 목적', '역화 방지 / 압력 차단', '—'),
        ('Seal Liquid', '공업용수 (Water)', '—'),
        ('적용 기준', 'API 521, 사내 기준', '—'),
    ],
    col_widths=[5, 7, 4]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  2. COG 물성
# ════════════════════════════════════════════════════════════
heading1('2. COG 물성 계산')
heading2('2-1. COG 표준 조성')
add_table(
    ['성분', '조성 (vol%)', '비고'],
    [
        ('H₂', '55.0', '주성분, 폭발 위험'),
        ('CH₄', '25.0', ''),
        ('CO', '6.0', '독성 (TLV 25 ppm)'),
        ('N₂', '5.0', ''),
        ('CₓHᵧ', '3.0', ''),
        ('CO₂', '2.0', ''),
        ('기타', '4.0', 'NH₃, H₂S, Benzene 포함'),
    ],
    col_widths=[4, 5, 7]
)

doc.add_paragraph()
heading2('2-2. 평균 분자량 및 밀도 계산')
code_block([
    '평균 분자량:',
    '  M = 0.55×2 + 0.25×16 + 0.06×28 + 0.05×28 + 0.03×30 + 0.02×44',
    '    = 1.10 + 4.00 + 1.68 + 1.40 + 0.90 + 0.88 = 10.8 g/mol',
    '',
    '표준 밀도 (0°C, 1atm):',
    '  ρ_N = M / 22.4 = 10.8 / 22.4 = 0.482 kg/Nm³  →  0.48 kg/Nm³ 적용',
    '',
    '운전 조건 환산:',
    '  T_op = 70 + 273.15 = 343.15 K',
    '  P_op = 101,325 + 13,734 = 115,059 Pa (abs)',
    '',
    '  ρ_op = 0.48 × (115,059/101,325) × (273.15/343.15)',
    '       = 0.48 × 1.1356 × 0.7961',
    '       = 0.434 kg/m³',
])
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  3. 유량 환산
# ════════════════════════════════════════════════════════════
heading1('3. 유량 환산 (Normal → Actual)')
code_block([
    'Q_op = Q_N × (P_N / P_op) × (T_op / T_N)',
    '     = 80,000 × (101,325 / 115,059) × (343.15 / 273.15)',
    '     = 80,000 × 0.8807 × 1.2563',
    '     = 88,580 m³/hr (actual)',
    '     = 24.61 m³/s',
])
doc.add_paragraph()
add_table(
    ['항목', '값', '단위'],
    [
        ('표준 유량 (Q_N)', '80,000', 'Nm³/hr'),
        ('실제 유량 (Q_op)', '88,580', 'm³/hr'),
        ('실제 유량 (Q_op)', '24.61', 'm³/s'),
    ],
    col_widths=[6, 6, 4]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  4. 배관 유속
# ════════════════════════════════════════════════════════════
heading1('4. 배관 유속 검토')
code_block([
    '배관 단면적:',
    '  A_pipe = π/4 × 2.4² = 4.524 m²',
    '',
    '배관 내 유속:',
    '  v = Q_op / A_pipe = 24.61 / 4.524 = 5.44 m/s',
    '',
    '동압:',
    '  q = 0.5 × ρ_op × v² = 0.5 × 0.434 × 5.44² = 6.4 Pa = 0.65 mmAq',
])
doc.add_paragraph()
add_table(
    ['항목', '계산값', '기준', '판정'],
    [
        ('배관 유속', '5.44 m/s', '≤ 8 m/s (COG 주배관)', '✅ OK'),
        ('동압 (q)', '6.4 Pa = 0.65 mmAq', '참고값', '—'),
    ],
    col_widths=[4.5, 4, 5, 2.5]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  5. 수봉 높이 계산
# ════════════════════════════════════════════════════════════
heading1('5. 수봉 높이 계산 (핵심 계산)')
heading2('5-1. Seal Liquid 물성 (70°C 물)')
add_table(
    ['항목', '값', '비고'],
    [
        ('물 밀도 (70°C)', '978 kg/m³', '표준값'),
        ('물 증기압 (70°C)', '31.2 kPa', '증발 모니터링 필요'),
        ('기준 밀도 (4°C)', '1,000 kg/m³', 'mmAq 정의 기준'),
    ],
    col_widths=[5, 5, 6]
)
note('mmAq는 4°C 물(1,000 kg/m³) 기준. 실제 수봉액(70°C, 978 kg/m³) 보정 필수.')
doc.add_paragraph()

heading2('5-2. 최소 수봉 높이')
code_block([
    '수봉 계산식:  h = ΔP / (ρ_water × g)',
    '',
    '운전 압력:',
    '  ΔP = 1,400 mmAq = 1,400 × 9.81 = 13,734 Pa',
    '',
    '최소 수봉 높이 (70°C 물 기준):',
    '  h_min = 13,734 / (978 × 9.81)',
    '        = 13,734 / 9,594',
    '        = 1,431 mm',
    '',
    '  ※ 4°C 물(1,000 kg/m³) 기준: 1,400 mm',
    '     70°C 물(978 kg/m³) 보정:  1,431 mm  (+2.2%)  → 온도 보정 필수',
])
doc.add_paragraph()

heading2('5-3. 설계 수봉 높이 (안전율 적용)')
code_block([
    '안전율 (SF) = 1.5  (역화방지 목적, COG 폭발 위험성 고려)',
    '',
    'h_design = h_min × SF = 1,431 × 1.5 = 2,147 mm  →  2,200 mm 적용',
])
doc.add_paragraph()
add_table(
    ['조건', '계산값', '비고'],
    [
        ('최소 수봉 높이 (h_min)', '1,431 mm', '70°C 물 보정 적용'),
        ('설계 수봉 높이 (h_design)', '2,200 mm', 'SF = 1.5, 올림 적용'),
        ('등가 유지 압력', '2,152 mmAq (21,103 Pa)', '978×9.81×2.2/9.81×1000'),
    ],
    col_widths=[6, 4.5, 5.5]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  6. 수봉변 형식 선정
# ════════════════════════════════════════════════════════════
heading1('6. 수봉변 형식 선정')
para('선정 형식: Drum Type Water Seal (대구경 강재 드럼)', bold=True)
doc.add_paragraph()
add_table(
    ['형식', '적합성', '판단'],
    [
        ('U-Type (매립형)', '배관 DN2400으로 공간 과다 소요', '△'),
        ('콘크리트 피트형', '대형 가스관에 가능, 유지관리 어려움', '△'),
        ('Drum Type (강재)', 'DN2400 COG 계통 표준, LG/LT 설치 용이, API 521 준용', '✅ 채택'),
        ('Loop Seal', '초대형 배관망 적합, 현 설계 과다', '✕'),
    ],
    col_widths=[4, 9, 3]
)
doc.add_paragraph()
para('채택 근거:', bold=True)
for item in [
    'DN2400 대구경: 충분한 액체 보유량(Liquid Inventory) 확보 필수',
    'COG 특성상 연속 Level 감시 필수 → 계장 설치 용이한 Drum Type',
    '콘크리트 피트 대비 유지보수성 우수, 누기 검지 용이',
    'API 521 4.3절 Gas Seal 설계 기준 준용',
]:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(item)
    run.font.name = '맑은 고딕'
    run.font.size = Pt(10)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  7. Seal Drum 사이즈 계산
# ════════════════════════════════════════════════════════════
heading1('7. Seal Drum 사이즈 계산')

heading2('7-1. Drum 직경 산정')
code_block([
    '기준: 드럼 내 가스 상부 유속 ≤ 2.0 m/s (Liquid Carry-over 방지)',
    '',
    '필요 드럼 단면적:',
    '  A_drum = Q_op / v_lim = 24.61 / 2.0 = 12.31 m²',
    '',
    '필요 직경:',
    '  D = √(4 × 12.31 / π) = 3.96 m  →  D_drum = 4,500 mm 적용',
    '',
    '실제 유효 가스 통과 면적 (Dip Pipe 제외):',
    '  A_eff = π/4 × 4.5² - π/4 × 2.4² = 15.90 - 4.52 = 11.38 m²',
    '',
    '  v_gas = 24.61 / 11.38 = 2.16 m/s  ← OK (≤ 3.0 m/s)',
])
doc.add_paragraph()

heading2('7-2. Drum 높이 및 Level 설정')
code_block([
    '┌───────────────────────────┐ ← 상부 탄젠트: 3,800 mm',
    '│    Top Freeboard  300 mm  │',
    '├───────────────────────────┤ ← HHL (High High Level): 3,500 mm',
    '│    HL ~ HHL:    300 mm    │',
    '├───────────────────────────┤ ← HL  (High Level):      3,200 mm',
    '│    NL ~ HL:     500 mm    │',
    '├───────────────────────────┤ ← NL  (Normal Level):    2,700 mm',
    '│                           │',
    '│    수봉 구간  2,200 mm    │  ← 설계 수봉 높이 확보 ✅',
    '│                           │',
    '├───────────────────────────┤ ← LL  (Low Level Alarm): 2,200 mm',
    '│    LLL ~ LL:    200 mm    │',
    '├───────────────────────────┤ ← LLL (ESD 트립):        2,000 mm',
    '│   Bottom Clearance 500 mm │',
    '└───────────────────────────┘ ← Dip Pipe 하단:          500 mm / 드럼 바닥: 0 mm',
])
doc.add_paragraph()

add_table(
    ['Level 항목', '드럼 바닥 기준 높이', '유효 수봉 높이', '수봉 SF'],
    [
        ('Dip Pipe 하단', '500 mm', '—', '—'),
        ('LLL (ESD 트립)', '2,000 mm', '1,500 mm', '1.05'),
        ('LL (저레벨 경보)', '2,200 mm', '1,700 mm', '1.19'),
        ('NL (정상 운전)', '2,700 mm', '2,200 mm ✅', '1.54'),
        ('HL (고레벨 경보)', '3,200 mm', '2,700 mm', '1.89'),
        ('HHL (고고레벨)', '3,500 mm', '3,000 mm', '2.10'),
    ],
    col_widths=[4.5, 4.5, 4, 3]
)
doc.add_paragraph()

heading2('7-3. Drum 기본 사이즈')
add_table(
    ['항목', '값'],
    [
        ('내경', '4,500 mm'),
        ('전체 높이 (T/T)', '3,800 mm'),
        ('설계 압력', '1,600 mmAq(g)  (운전의 1.15배)'),
        ('설계 온도', '90°C'),
        ('재질', 'SS400 + 내부 에폭시 코팅 (COG 부식 대응)'),
    ],
    col_widths=[6, 10]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  8. 보유 액량
# ════════════════════════════════════════════════════════════
heading1('8. 보유 액량 계산')
code_block([
    '드럼 단면적: A = π/4 × 4.5² = 15.90 m²',
    '',
    'NL 기준 보유 액량:',
    '  V_NL = 15.90 × 2.70 = 42.9 m³',
    '',
    'LL 기준 보유 액량:',
    '  V_LL = 15.90 × 2.20 = 35.0 m³',
    '',
    '유효 Liquid Reserve (NL → LL):',
    '  ΔV = 42.9 - 35.0 = 7.9 m³',
    '',
    '보수적 증발 추정 (2 m³/hr 가정):',
    '  Reserve 유지 시간 = 7.9 / 2.0 ≈ 3.9 hr → LL 경보 후 충분한 조치 시간 확보',
])
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  9. 설계 검증
# ════════════════════════════════════════════════════════════
heading1('9. 설계 검증 계산')

heading2('9-1. Gas Blow-through 검토')
code_block([
    '수봉 정압 압력 (NL 기준):',
    '  P_seal = ρ_w × g × h_NL = 978 × 9.81 × 2.20 = 21,103 Pa',
    '',
    '가스 동압:',
    '  q_gas = 0.5 × ρ_op × v² = 0.5 × 0.434 × 5.44² = 6.4 Pa',
    '',
    '동압/정압 비율:',
    '  q / P_seal = 6.4 / 21,103 = 0.0003 (0.03%)',
    '',
    '  → 가스 동압이 수봉 정압의 0.03% 수준 → Blow-through 위험 없음 ✅',
])
doc.add_paragraph()

heading2('9-2. Water Carry-over 검토 (500 μm 액적 기준)')
code_block([
    '종말 침강 속도 (C_D = 0.44):',
    '  v_t = √[4 × g × d_p × (ρ_L - ρ_G) / (3 × C_D × ρ_G)]',
    '      = √[4 × 9.81 × 0.0005 × 977.57 / (3 × 0.44 × 0.434)]',
    '      = √[19.18 / 0.573]',
    '      = √33.47 = 5.79 m/s',
    '',
    '드럼 내 가스 상부 유속: 2.16 m/s',
    '',
    '  v_gas (2.16) < v_t (5.79) → 500 μm 이상 액적 침강 ✅',
    '  ※ 200 μm 이하 미세 액적은 Demister Pad 설치로 포집 권장',
])
doc.add_paragraph()

heading2('9-3. Seal Break Margin 검토')
code_block([
    'NL 기준 수봉 최대 유지 압력:',
    '  P_seal_max = 978 × 9.81 × 2.20 = 21,103 Pa = 2,152 mmAq',
    '',
    'Seal Break Margin:',
    '  Margin = (2,152 - 1,400) / 1,400 × 100 = 53.7%',
    '',
    '  → 정상 운전 중 수봉 파손 여유 53.7% 확보 ✅',
    '  → LLL ESD 설정 시 1,500 mm 수봉 = 1,470 mmAq 유지',
])
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  10. Dip Pipe
# ════════════════════════════════════════════════════════════
heading1('10. Dip Pipe (침강관) 설계')
add_table(
    ['항목', '값', '비고'],
    [
        ('직경', 'DN2400', 'Main Pipe와 동일'),
        ('재질', 'SS400 + 내부 에폭시 코팅', 'COG 내 H₂S 대응'),
        ('드럼 바닥 기준 하단 위치', '500 mm', '침전물 흡입 방지'),
        ('하단 형상', '45° Bevel Cut', '가스 기포 분산'),
        ('유효 침강 깊이 (NL 기준)', '2,200 mm ✅', '설계 수봉 높이 확보'),
        ('Dip Pipe 내 가스 유속', '5.44 m/s', '배관 유속과 동일'),
    ],
    col_widths=[5, 5, 6]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  11. 노즐 목록
# ════════════════════════════════════════════════════════════
heading1('11. 노즐 목록 (Nozzle Schedule)')
add_table(
    ['번호', '명칭', '크기', '위치', '비고'],
    [
        ('N1', 'Gas Inlet (Dip Pipe)', 'DN2400', 'Top', 'Dip Pipe 포함'),
        ('N2', 'Gas Outlet', 'DN2400', 'Top/Side', 'NL + 500 mm 상부'),
        ('N3', 'Makeup Water Inlet', 'DN100 (4")', 'Side (HL 상부)', 'Check Valve 포함'),
        ('N4', 'Overflow', 'DN200 (8")', 'Side (HHL)', 'Free Overflow'),
        ('N5', 'Drain', 'DN100 (4")', 'Bottom', '수동 차단 밸브'),
        ('N6', 'Level Gauge 하부', 'DN50 (2")', 'Side', 'LLL 위치'),
        ('N7', 'Level Gauge 상부', 'DN50 (2")', 'Side', 'HL 위치'),
        ('N8', 'Level Transmitter (LT)', 'DN50 (2") × 2', 'Side', '이중화'),
        ('N9', 'Vent', 'DN50 (2")', 'Top', 'N₂ Purge 연결'),
        ('N10', 'Pressure Gauge (PI)', 'DN25 (1")', 'Top', ''),
        ('N11', 'Temperature Gauge (TI)', 'DN25 (1")', 'Side', ''),
        ('N12', 'Sample', 'DN25 (1")', 'Side', 'NL 위치'),
        ('N13', 'Manhole', 'DN600 × 2', 'Top, Side', '상부 + 측면'),
        ('N14', 'Inspection Nozzle', 'DN200 (8")', 'Side', '드럼 하부 점검'),
    ],
    col_widths=[1.2, 4.2, 3, 3, 4.6]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  12. 계장 및 안전장치
# ════════════════════════════════════════════════════════════
heading1('12. 계장 및 안전장치')
heading2('Level 연동 시퀀스')
code_block([
    'HHL (3,500 mm) ──→ 경보 + Overflow 자동 개방',
    'HL  (3,200 mm) ──→ 경보 + Makeup Water 차단',
    'NL  (2,700 mm) ──→ 정상 운전',
    'LL  (2,200 mm) ──→ 경보 + Makeup Water 자동 공급',
    'LLL (2,000 mm) ──→ 경보 + ESD (긴급 조치 요구)',
    '                   ※ 수봉 파손 임박: 즉각 조치',
])
doc.add_paragraph()
add_table(
    ['계장 항목', '사양', '목적'],
    [
        ('LT (Level Transmitter)', '2중화 (1oo2 표결)', '정확한 Level 감시'),
        ('LG (Level Gauge)', '투명 유리식, 조명 포함', '현장 확인'),
        ('LAH / LAL', 'Level 고/저 경보', 'DCS 연동'),
        ('LAHH / LALL', 'Level 고고/저저 경보', 'ESD 연동'),
        ('PI (Pressure Gauge)', '현장 지시계', '드럼 내압 확인'),
        ('TI (Temperature Gauge)', '현장 지시계', '수온 모니터링'),
        ('Makeup Water Auto-valve', '모터 구동 밸브 (MOV)', 'LL 시 자동 급수'),
        ('N₂ Purge', '드럼 Vent에 연결', '유지보수 시 가스 치환'),
    ],
    col_widths=[5, 6, 5]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  13. COG 특성 고려사항
# ════════════════════════════════════════════════════════════
heading1('13. COG 특성 고려사항')
add_table(
    ['위험 요인', '내용', '대응 방안'],
    [
        ('H₂S 함유', '배관/드럼 내면 황화수소 부식', '내면 에폭시 코팅 + SS재질 검토'),
        ('NH₃ 함유', '봉수 pH 상승 (알칼리화)', '봉수 정기 수질 분석 (pH 모니터링)'),
        ('벤젠/나프탈렌', '봉수 오염, 배출수 규제', '봉수 분리처리 (Oily Water 처리)'),
        ('CO 함유 (6%)', '누기 시 인체 독성 위험', 'COG 가스 검지기 주변 설치'),
        ('H₂ 함유 (55%)', '폭발 하한 4% — 최고 폭발 위험', '밀폐 공간 작업 시 N₂ 치환 필수'),
        ('70°C 운전', '봉수 증발 가속', 'Makeup Water 자동 연동'),
        ('봉수 증기압 (70°C)', '31.2 kPa → 증발량 고려', '드럼 온도 저하 시 응축 발생 가능'),
        ('Tar 성분', '드럼 하부 타르 침전', '정기 Drain (월 1회 이상)'),
    ],
    col_widths=[3.5, 5.5, 7]
)
doc.add_paragraph()

# ════════════════════════════════════════════════════════════
#  14. 최종 설계 요약
# ════════════════════════════════════════════════════════════
heading1('14. 최종 설계 요약')
heading2('설계 요약표')
add_table(
    ['항목', '결과', '단위'],
    [
        ('유체', 'COG (Coke Oven Gas)', '—'),
        ('운전 온도', '70', '°C'),
        ('운전 압력', '1,400', 'mmAq(g)'),
        ('설계 유량', '80,000 Nm³/hr / 88,580 m³/hr(act)', '—'),
        ('COG 밀도 (운전 조건)', '0.434', 'kg/m³'),
        ('배관 유속', '5.44', 'm/s'),
        ('최소 수봉 높이 (h_min)', '1,431', 'mm'),
        ('설계 수봉 높이 (h_design)', '2,200', 'mm'),
        ('적용 안전율 (SF)', '1.54', '—'),
        ('Seal 형식', 'Drum Type Water Seal', '—'),
        ('Drum 내경', '4,500', 'mm'),
        ('Drum 높이 (T/T)', '3,800', 'mm'),
        ('드럼 내 가스 유속', '2.16', 'm/s'),
        ('Dip Pipe 직경', '2,400', 'mm'),
        ('Normal Level (NL)', '2,700', 'mm (바닥 기준)'),
        ('Low Level Alarm (LL)', '2,200', 'mm (바닥 기준)'),
        ('Low Low Level (LLL / ESD)', '2,000', 'mm (바닥 기준)'),
        ('High Level (HL)', '3,200', 'mm (바닥 기준)'),
        ('NL 기준 보유 액량', '42.9', 'm³'),
        ('LL~NL 유효 Reserve', '7.9', 'm³'),
        ('Seal Break Margin (NL 기준)', '53.7', '%'),
    ],
    col_widths=[6.5, 6, 3.5]
)
doc.add_paragraph()

heading2('주의사항')
warning_box('LLL 도달 시 수봉 파손 → COG 역류/역화 발생 가능. 즉각 조치 (ESD 연동 권장)')
doc.add_paragraph()
warning_box('H₂S 함유 COG는 누기 시 독성 위험. 주변 가스 검지기(CO/H₂S) 연속 감시 필수')
doc.add_paragraph()
warning_box('봉수 오염 (타르, 벤젠) → 정기 수질 분석 및 Oily Water 처리 설비 연계')
doc.add_paragraph()

heading2('정기 유지관리')
for item in [
    '봉수 타르 침전물: 월 1회 이상 Drain',
    'Level Gauge 청소: 분기 1회',
    '봉수 pH 측정: 월 1회 (NH₃로 인한 알칼리화 확인)',
    'Dip Pipe 내면 코팅 상태: 연 1회 정기보수 시 점검',
    '가스 검지기 교정: 분기 1회',
]:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(item)
    run.font.name = '맑은 고딕'
    run.font.size = Pt(10)

doc.add_paragraph()
p = doc.add_paragraph()
run = p.add_run(
    '종합 의견: DN2400 / 80,000 Nm³/hr COG 대구경 배관에서 설계 수봉 높이 2,200 mm, '
    '드럼 직경 4,500 mm, 전체 높이 3,800 mm의 Drum Type Water Seal을 적용한다. '
    '온도 보정(70°C 물 밀도 978 kg/m³)을 반드시 적용하며, COG의 H₂, CO, H₂S '
    '복합 위험 특성상 Level 연속 감시 및 LLL ESD 연동을 필수로 요구한다.'
)
run.font.name = '맑은 고딕'
run.font.size = Pt(10)
run.font.bold = True
p.paragraph_format.left_indent = Cm(0.5)
pPr = p._p.get_or_add_pPr()
shd = OxmlElement('w:shd')
shd.set(qn('w:val'), 'clear')
shd.set(qn('w:color'), 'auto')
shd.set(qn('w:fill'), 'EBF3FB')
pPr.append(shd)

# 저장
out_path = r'C:\Users\yesksm\claude\COG_WaterSeal_계산서.docx'
doc.save(out_path)
print(f'저장 완료: {out_path}')
