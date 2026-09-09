import os
import math
import datetime
import openpyxl
import pandas as pd

# -------------------------------------------------------------
# 1. 영종도 천문 일출 / 일몰 계산기
# -------------------------------------------------------------
def get_solar_times(dt: datetime.date):
    lat, lon = 37.46, 126.44
    doy = dt.timetuple().tm_yday
    gamma = 2 * math.pi / 365 * (doy - 1)
    decl = 0.006918 - 0.399912 * math.cos(gamma) + 0.070257 * math.sin(gamma) \
           - 0.006758 * math.cos(2*gamma) + 0.000907 * math.sin(2*gamma)
    eqtime = 229.18 * (0.000075 + 0.001868 * math.cos(gamma) - 0.032077 * math.sin(gamma) \
             - 0.014615 * math.cos(2*gamma) - 0.040849 * math.sin(2*gamma))
    zenith = math.radians(90.833)
    lat_rad = math.radians(lat)
    cos_ha = (math.cos(zenith) / (math.cos(lat_rad) * math.cos(decl))) - (math.tan(lat_rad) * math.tan(decl))
    cos_ha = max(-1.0, min(1.0, cos_ha))
    ha = math.degrees(math.acos(cos_ha))
    solar_noon = 720 - 4 * lon - eqtime + 9 * 60
    return solar_noon - ha * 4, solar_noon + ha * 4

def classify_time_window(dt: datetime.date, time_str: str):
    if not time_str or str(time_str).strip() in ['None', '시간미상', '']:
        return 'Unknown'
    try:
        parts = str(time_str).strip().split(':')
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        t_min = h * 60 + m
    except:
        return 'Unknown'
    
    sunrise_min, sunset_min = get_solar_times(dt)
    if sunrise_min - 90 <= t_min <= sunrise_min + 90: return 'Dawn'
    elif sunset_min - 90 <= t_min <= sunset_min + 90: return 'Dusk'
    elif sunrise_min + 90 < t_min < sunset_min - 90: return 'Day'
    else: return 'Night'

# -------------------------------------------------------------
# 2. 활주로 ➜ 19개 표준 Zone 매핑
# -------------------------------------------------------------
def map_runway_to_zone(rw_str):
    rw = str(rw_str).strip().upper() if rw_str else ""
    if '34L' in rw: return 'ZONE-R4-S'
    elif '34R' in rw: return 'ZONE-R3-S'
    elif '16R' in rw: return 'ZONE-R4-N'
    elif '16L' in rw: return 'ZONE-R3-N'
    elif '33L' in rw: return 'ZONE-R2-S'
    elif '33R' in rw: return 'ZONE-R1-S'
    elif '15R' in rw: return 'ZONE-R2-N'
    elif '15L' in rw: return 'ZONE-R1-N'
    elif '34' in rw: return 'ZONE-R4-S'
    elif '16' in rw: return 'ZONE-R4-N'
    elif '33' in rw: return 'ZONE-R1-S'
    elif '15' in rw: return 'ZONE-R1-N'
    return 'ZONE-UNKNOWN'

# -------------------------------------------------------------
# 3. 25개년 장기 시계열 추이 엑셀 파싱
# -------------------------------------------------------------
def parse_annual_trend(excel_path, out_csv):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb['연간조류충돌 현황']
    rows = []
    for r in range(4, ws.max_row + 1):
        year = ws.cell(row=r, column=2).value
        on_ap = ws.cell(row=r, column=3).value
        near_ap = ws.cell(row=r, column=4).value
        off_ap = ws.cell(row=r, column=5).value
        unk = ws.cell(row=r, column=6).value
        tot = ws.cell(row=r, column=7).value
        if year:
            rows.append({
                "연도": int(year),
                "공항내부(On)": int(on_ap or 0),
                "공항인근(Near)": int(near_ap or 0),
                "공항외부(Off)": int(off_ap or 0),
                "구역미상(Unknown)": int(unk or 0),
                "총충돌건수": int(tot or 0)
            })
    df = pd.DataFrame(rows).sort_values("연도")
    df.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print(f"Annual trend parsed: {len(df)} years ➜ {out_csv}")
    return df

# -------------------------------------------------------------
# 4. 개별 연도 조류충돌 엑셀 파싱 (2024, 2025)
# -------------------------------------------------------------
def parse_birdstrike_incidents(excel_path, year_label):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    
    # 헤더 행 찾기
    h_row = 3
    for r in range(1, 10):
        row = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if any(c and '일자' in str(c) for c in row):
            h_row = r
            break
            
    records = []
    for r in range(h_row + 1, ws.max_row + 1):
        d_val = ws.cell(row=r, column=3).value # 일자
        if not d_val: continue
        d_str = str(d_val)[:10].strip()
        try:
            dt = datetime.date.fromisoformat(d_str)
        except:
            continue
            
        t_val = ws.cell(row=r, column=4).value # 시간
        t_str = str(t_val).strip() if t_val else "시간미상"
        species = ws.cell(row=r, column=5).value or "미상"
        airline = ws.cell(row=r, column=6).value or "항공사미상"
        flight = ws.cell(row=r, column=7).value or ""
        aircraft = ws.cell(row=r, column=8).value or ""
        route = ws.cell(row=r, column=9).value or ""
        runway = ws.cell(row=r, column=10).value or ""
        alt = ws.cell(row=r, column=11).value or ""
        phase = ws.cell(row=r, column=12).value or ""
        part = ws.cell(row=r, column=13).value or ""
        damage = ws.cell(row=r, column=14).value or "N"
        loc_desc = ws.cell(row=r, column=15).value or ""
        pilot_rep = ws.cell(row=r, column=16).value or ""
        dna_req = ws.cell(row=r, column=17).value or ""
        dna_res = ws.cell(row=r, column=18).value or ""
        
        time_window = classify_time_window(dt, t_str)
        zone_id = map_runway_to_zone(runway)
        
        # 기러기 위험 여부
        is_goose = any(k in str(species) for k in ['기러기', '큰기러기', '쇠기러기'])
        is_identified = (str(species).strip() not in ['미상', 'None', '', '미식별'])
        
        records.append({
            "연도": dt.year,
            "일자": d_str,
            "시간": t_str,
            "시간대_윈도우": time_window,
            "활주로": str(runway).strip(),
            "표준_Zone": zone_id,
            "조류종": str(species).strip(),
            "종식별여부": "식별성공" if is_identified else "미상",
            "기러기여부": "Y" if is_goose else "N",
            "항공사": str(airline).strip(),
            "편명": str(flight).strip(),
            "기종": str(aircraft).strip(),
            "노선": str(route).strip(),
            "고도": str(alt).strip(),
            "운항단계": str(phase).strip(),
            "충돌부위": str(part).strip(),
            "피해여부": str(damage).strip(),
            "DNA결과": str(dna_res).strip()
        })
    df = pd.DataFrame(records)
    print(f"Parsed {year_label}: {len(df)} records")
    return df

# -------------------------------------------------------------
# 5. 전체 실행 및 교차 검증 종합 보고서 생성
# -------------------------------------------------------------
def run_all():
    out_dir = r"C:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled"
    os.makedirs(out_dir, exist_ok=True)
    
    # 1) 25개년 추이
    trend_src = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2024년 자료\조류충돌 연간 횟수 현황.xlsx"
    trend_out = os.path.join(out_dir, "birdstrike_annual_trend_2001_2026.csv")
    df_trend = parse_annual_trend(trend_src, trend_out)
    
    # 2) 2024 & 2025 개별 충돌 건
    bs_2024_src = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2024년 자료\2024년 조류충돌 전체.xlsx"
    bs_2025_src = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2025년 자료\2025년 통계\2025년 조류충돌.xlsx"
    
    df_bs24 = parse_birdstrike_incidents(bs_2024_src, "2024년 조류충돌")
    df_bs25 = parse_birdstrike_incidents(bs_2025_src, "2025년 조류충돌")
    
    df_bs24.to_csv(os.path.join(out_dir, "2024_birdstrike_resampled.csv"), index=False, encoding="utf-8-sig")
    df_bs25.to_csv(os.path.join(out_dir, "2025_birdstrike_resampled.csv"), index=False, encoding="utf-8-sig")
    
    df_bs_total = pd.concat([df_bs24, df_bs25], ignore_index=True)
    df_bs_total.to_csv(os.path.join(out_dir, "2024_2025_birdstrike_matrix.csv"), index=False, encoding="utf-8-sig")
    
    # 3) 지상 일일점검(순찰)과의 교차 대조 분석
    patrol_master = r"C:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2024_2025_master_feature_matrix.csv"
    if os.path.exists(patrol_master):
        df_patrol = pd.read_csv(patrol_master)
        # 활주로 충돌 발생일과 순찰 발포량 대조
        strike_dates = set(df_bs_total['일자'])
        patrol_on_strike = df_patrol[df_patrol['일자'].isin(strike_dates)]
        print(f"Patrol events on strike days: {len(patrol_on_strike):,d} records")
        
    print("\nAll Bird Strike Resampling & Matrix Generation Done!")

if __name__ == "__main__":
    run_all()
