# 펌프 / 탱크 설계 계산서

POSCO 플랜트 설계용 엑셀 계산서. 모든 값은 셀 수식으로 연동되어 입력값만 바꾸면 자동 재계산됩니다.

## 파일

| 파일 | 내용 |
|---|---|
| `BFP_Calculation.xlsx` | Boiler Feed Water Pump 계산서 (100% BMCR / 95% / 30% Load) |
| `Tank_Calculation.xlsx` | 저장탱크 사이징 계산서 (Condensate / Service&Fire / Demi / Fuel Oil / Seawater) |
| `make_bfp_calc.py` | BFP 계산서 생성 스크립트 (openpyxl) |
| `make_tank_calc.py` | 탱크 계산서 생성 스크립트 (openpyxl) |

## 재생성

```bash
pip install openpyxl
python make_bfp_calc.py
python make_tank_calc.py
```
