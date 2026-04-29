"""
코스피/코스닥 투자자별 매매동향 조회 및 Google Sheets 업로드
데이터 출처: KRX (pykrx)  |  단위: 억원  |  최근 7 영업일
claude-review 자동 리뷰 테스트
"""

import json
import sys
import warnings
from datetime import datetime, timedelta

# KRX self-signed 인증서 우회 (requests 세션 포함)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
import requests
_orig_send = requests.Session.send
def _send_no_verify(self, *args, **kwargs):
    kwargs["verify"] = False
    return _orig_send(self, *args, **kwargs)
requests.Session.send = _send_no_verify

import pandas as pd
from pykrx import stock


def get_business_days(n: int) -> tuple[str, str]:
    """최근 n 영업일의 시작일/종료일 반환 (YYYYMMDD)"""
    days = []
    d = datetime.today()
    while len(days) < n:
        if d.weekday() < 5:
            days.append(d)
        d -= timedelta(days=1)
    end_dt   = days[0].strftime("%Y%m%d")
    start_dt = days[-1].strftime("%Y%m%d")
    return start_dt, end_dt


def fetch_trend(market: str, start: str, end: str) -> pd.DataFrame:
    """투자자별 순매수 DataFrame (억원 단위)"""
    df = stock.get_market_trading_value_by_investor(start, end, market)
    # 순매수 컬럼만 추출 (백만원 → 억원)
    net_cols = [c for c in df.columns if "순매수" in c]
    df_net = df[net_cols].copy()
    df_net.columns = [c.replace("순매수", "").strip() for c in net_cols]
    df_net = (df_net / 100).round(1)   # 백만원 → 억원
    return df_net


def build_sheet_data(label: str, df: pd.DataFrame) -> list[list]:
    """Google Sheets에 넣을 2D 배열 생성"""
    rows = []
    rows.append([f"{label} 투자자별 순매수 (억원)"])
    rows.append([])

    # 헤더: 날짜 행
    date_row = ["투자자 \\ 날짜"] + [
        f"{str(idx)[:4]}-{str(idx)[4:6]}-{str(idx)[6:]}" for idx in df.index
    ]
    rows.append(date_row)

    # 투자자별 데이터 행
    display_order = ["외국인", "기관계", "금융투자", "보험", "투신", "은행", "기타금융", "연기금등", "개인"]
    for inv in display_order:
        if inv in df.columns:
            row = [inv] + [df.at[idx, inv] for idx in df.index]
            rows.append(row)

    rows.append([])
    return rows


def print_table(label: str, df: pd.DataFrame):
    """터미널 출력"""
    print(f"\n{'='*70}")
    print(f"  {label}  투자자별 순매수 동향  (단위: 억원)")
    print(f"{'='*70}")
    date_labels = [f"{str(i)[4:6]}/{str(i)[6:]}" for i in df.index]
    header = f"  {'투자자':<12}" + "".join(f"{d:>10}" for d in date_labels)
    print(header)
    print(f"  {'-'*65}")

    display_order = ["외국인", "기관계", "금융투자", "보험", "투신", "은행", "기타금융", "연기금등", "개인"]
    for inv in display_order:
        if inv not in df.columns:
            continue
        if inv in ("외국인", "기관계", "개인"):
            print(f"  {'-'*65}")
        vals = [df.at[idx, inv] for idx in df.index]
        val_str = "".join(
            f"{'▲' if v > 0 else '▼' if v < 0 else ' '}{abs(v):>8,.1f}" for v in vals
        )
        print(f"  {inv:<12}{val_str}")
    print()


def upload_to_gsheet(all_rows: list[list], title: str):
    """Google Sheets MCP를 통해 업로드"""
    # CSV 형태 문자열 생성
    csv_lines = []
    for row in all_rows:
        csv_lines.append(",".join(str(c) for c in row))
    csv_content = "\n".join(csv_lines)
    return csv_content


def main():
    start, end = get_business_days(7)
    print(f"\n▶ 조회 기간: {start[:4]}-{start[4:6]}-{start[6:]} ~ {end[:4]}-{end[4:6]}-{end[6:]}  (최근 7 영업일)\n")

    markets = [
        ("KOSPI",  "KOSPI (코스피)"),
        ("KOSDAQ", "KOSDAQ (코스닥)"),
    ]

    all_rows = [
        [f"코스피/코스닥 투자자별 매매동향"],
        [f"조회기간: {start[:4]}-{start[4:6]}-{start[6:]} ~ {end[:4]}-{end[4:6]}-{end[6:]}"],
        [],
    ]

    dfs = {}
    for code, name in markets:
        try:
            df = fetch_trend(code, start, end)
            dfs[code] = df
            print_table(name, df)
            all_rows.extend(build_sheet_data(name, df))
        except Exception as e:
            print(f"  [{name}] 오류: {e}")

    # JSON으로 저장 (업로드용)
    with open("C:/Users/yesksm/pjt/trend_data.json", "w", encoding="utf-8") as f:
        json.dump(all_rows, f, ensure_ascii=False, default=str)

    print("[완료] trend_data.json 저장")
    return all_rows


if __name__ == "__main__":
    main()
