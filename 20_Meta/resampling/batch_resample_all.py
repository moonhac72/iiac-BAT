import os
import re
import glob
import math
import datetime
import openpyxl
import pandas as pd

# -------------------------------------------------------------
# 1. 영종도 천문 일출 / 일몰 계산기 (Yeongjongdo Solar Engine)
#    위도: 37.46 N, 경도: 126.44 E, KST (UTC+9)
# -------------------------------------------------------------
def get_solar_times(dt: datetime.date):
    lat = 37.46
    lon = 126.44
    day_of_year = dt.timetuple().tm_yday
    
    gamma = 2 * math.pi / 365 * (day_of_year - 1)
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
    
    sunrise_min = solar_noon - ha * 4
    sunset_min = solar_noon + ha * 4
    
    return sunrise_min, sunset_min

def classify_time_window(dt: datetime.date, time_str: str):
    try:
        parts = time_str.split(':')
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 else 0
        t_min = h * 60 + m
    except:
        return 'Night', 0, 0
    
    sunrise_min, sunset_min = get_solar_times(dt)
    
    dawn_start = sunrise_min - 90
    dawn_end = sunrise_min + 90
    
    dusk_start = sunset_min - 90
    dusk_end = sunset_min + 90
    
    if dawn_start <= t_min <= dawn_end:
        return 'Dawn', sunrise_min, sunset_min
    elif dusk_start <= t_min <= dusk_end:
        return 'Dusk', sunrise_min, sunset_min
    elif dawn_end < t_min < dusk_start:
        return 'Day', sunrise_min, sunset_min
    else:
        return 'Night', sunrise_min, sunset_min


# -------------------------------------------------------------
# 2. 19개 핵심 분석 Zone 매핑 엔진
# -------------------------------------------------------------
def map_to_19_zones(map_group, region, detail):
    mg = str(map_group).strip().upper() if map_group else ""
    rg = str(region).strip() if region else ""
    dt = str(detail).strip() if detail else ""
    
    is_ditch = ("배수로" in rg) or ("배수로" in dt)
    micro_habitat = "Warm_Water_Outflow" if is_ditch else "Normal"
    
    # 1) 배수로 지정
    if is_ditch:
        if "33" in mg or "33" in dt:
            return "ZONE-DR-33S", micro_habitat
        elif "15R" in dt or ("15" in mg and ("1R" in dt or "1R/W" in dt or "서" in dt)):
            return "ZONE-DR-15R-N", micro_habitat
        elif "15L" in dt or ("15" in mg and ("2R" in dt or "2R/W" in dt or "동" in dt)):
            return "ZONE-DR-15L-N", micro_habitat
        elif "34L" in dt or ("34" in mg and ("4R" in dt or "서" in dt)):
            return "ZONE-DR-34L-S", micro_habitat
        elif "34R" in dt or ("34" in mg and ("3R" in dt or "동" in dt)):
            return "ZONE-DR-34R-S", micro_habitat
        elif "16R" in dt or ("16" in mg and ("4R" in dt or "서" in dt)):
            return "ZONE-DR-16R-N", micro_habitat
        elif "16L" in dt or ("16" in mg and ("3R" in dt or "동" in dt)):
            return "ZONE-DR-16L-N", micro_habitat
        else:
            if "15" in mg: return "ZONE-DR-15L-N", micro_habitat
            if "16" in mg: return "ZONE-DR-16R-N", micro_habitat
            if "33" in mg: return "ZONE-DR-33S", micro_habitat
            if "34" in mg: return "ZONE-DR-34L-S", micro_habitat
    
    # 2) 활주로 및 녹지대 상세 매핑
    if "4활주로" in rg or "4R/W" in rg or "4RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R4-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R4-N", micro_habitat
        return "ZONE-R4-M", micro_habitat
    if "3활주로" in rg or "3R/W" in rg or "3RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R3-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R3-N", micro_habitat
        return "ZONE-R3-M", micro_habitat
    if "2활주로" in rg or "2R/W" in rg or "2RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R2-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R2-N", micro_habitat
        return "ZONE-R2-M", micro_habitat
    if "1활주로" in rg or "1R/W" in rg or "1RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R1-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R1-N", micro_habitat
        return "ZONE-R1-M", micro_habitat
        
    # 3) 맵영역그룹 기본 매핑
    if "AS-15" in mg:
        return "ZONE-R1-N", micro_habitat
    elif "AS-16" in mg:
        return "ZONE-R4-N", micro_habitat
    elif "AS-33" in mg:
        return "ZONE-R1-S", micro_habitat
    elif "AS-34" in mg:
        return "ZONE-R4-S", micro_habitat
    elif "LS" in mg or "랜드" in rg or "공사" in rg:
        return "ZONE-LANDSIDE", micro_habitat
    else:
        return "ZONE-PERI-SOUTH", micro_habitat


# -------------------------------------------------------------
# 3. 조류종 신뢰도 5단계 판정 엔진
# -------------------------------------------------------------
NOCTURNAL_SPECIES = {'수리부엉이', '올빼미', '쏙독새', '소쩍새', '솔부엉이'}
DUCK_SPECIES = {'흰뺨검둥오리', '청둥오리', '오리류', '쇠오리', '고방오리', '홍머리오리', '원앙', '가창오리', '오리'}
GEESE_SPECIES = {'큰기러기', '쇠기러기', '기러기', '기러기류'}

def evaluate_species_certainty(dt: datetime.date, time_window: str, species: str, activity: str, is_ditch: bool):
    sp = str(species).strip() if species else "미상"
    act = str(activity).strip() if activity else ""
    
    # Grade 1: 실물 검증
    if act in ['포획', '사체수거']:
        return "Grade 1 (Verified)", 1.0, sp, "실물/포획 검증"
    
    # Grade 2: 야행성 고유종
    if sp in NOCTURNAL_SPECIES:
        return "Grade 2 (Certain)", 1.0, sp, "야행성 고유종"
    
    # 광량 확보 시간대 (Day, Dawn, Dusk)
    if time_window in ['Day', 'Dawn', 'Dusk']:
        if sp in ['미상', 'None', '']:
            return "Grade 3 (High)", 0.9, "미식별", "주간/조석 미상 조류"
        return "Grade 3 (High)", 0.9, sp, f"{time_window} 육안 식별"
    
    # 이하 Time_Window == 'Night' (야간)
    # 2025년 4월 이전 기러기류
    is_pre_2025_04 = dt <= datetime.date(2025, 4, 30)
    if sp in GEESE_SPECIES and is_pre_2025_04:
        return "Grade 3 (High)", 0.9, sp, "25년 4월 이전 기러기 야간 잠자리 집중기"
    
    # 야간 오리류 (배수로 등 온수 서식지)
    if sp in DUCK_SPECIES:
        return "Grade 4 (Probabilistic)", 0.5, sp, "야간 배수로 오리류 (임의작성 경향 50% 가중치)"
    
    # 야간 일반 주행성 조류
    if sp not in ['미상', 'None', '']:
        return "Grade 5 (Doubtful)", 0.1, "미식별(야간_주행성_오인의심)", f"야간 식별 취약 주행성 조류({sp}) 노이즈 블라인드"
    
    return "Grade 5 (Doubtful)", 0.1, "미식별", "야간 미상"


# -------------------------------------------------------------
# 4. 동적 엑셀 파싱 및 리샘플링 단일 파일 처리
# -------------------------------------------------------------
def resample_file(file_path: str):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    
    # 1) 헤더 행 및 컬럼 인덱스 자동 감지
    header_row_idx = None
    col_map = {}
    
    for r_idx in range(1, 15):
        row_vals = [ws.cell(row=r_idx, column=c).value for c in range(1, ws.max_column + 1)]
        if any(v and '점검일자' in str(v) for v in row_vals):
            header_row_idx = r_idx
            for c_idx, val in enumerate(row_vals, 1):
                if not val: continue
                sval = str(val).strip()
                if '점검일자' in sval: col_map['date'] = c_idx
                elif '시간' in sval: col_map['time'] = c_idx
                elif '조류종' in sval: col_map['species'] = c_idx
                elif '개체수' in sval: col_map['count'] = c_idx
                elif '활동' in sval: col_map['activity'] = c_idx
                elif '장비' in sval: col_map['equip'] = c_idx
                elif '7,1/2' in sval or '7호탄' in sval: col_map['shot7'] = c_idx
                elif sval == '4' or '4호탄' in sval: col_map['shot4'] = c_idx
                elif '공포탄' in sval: col_map['blank'] = c_idx
                elif '대분류' in sval: col_map['major'] = c_idx
                elif '중분류' in sval: col_map['mid'] = c_idx
                elif '지역상세' in sval: col_map['detail'] = c_idx
                elif '지역' in sval: col_map['region'] = c_idx
                elif '맵영역그룹' in sval: col_map['map_group'] = c_idx
            break
            
    if not header_row_idx or 'date' not in col_map:
        print(f"[WARN] Header not found in {file_path}")
        return []

    records = []
    
    for r in range(header_row_idx + 1, ws.max_row + 1):
        raw_date = ws.cell(row=r, column=col_map['date']).value if 'date' in col_map else None
        if not raw_date:
            continue
            
        raw_time = ws.cell(row=r, column=col_map['time']).value if 'time' in col_map else None
        raw_species = ws.cell(row=r, column=col_map['species']).value if 'species' in col_map else None
        raw_count = ws.cell(row=r, column=col_map['count']).value if 'count' in col_map else None
        raw_activity = ws.cell(row=r, column=col_map['activity']).value if 'activity' in col_map else None
        raw_shot7 = ws.cell(row=r, column=col_map['shot7']).value if 'shot7' in col_map else 0
        raw_shot4 = ws.cell(row=r, column=col_map['shot4']).value if 'shot4' in col_map else 0
        raw_blank = ws.cell(row=r, column=col_map['blank']).value if 'blank' in col_map else 0
        raw_region = ws.cell(row=r, column=col_map['region']).value if 'region' in col_map else None
        raw_detail = ws.cell(row=r, column=col_map['detail']).value if 'detail' in col_map else None
        raw_map_group = ws.cell(row=r, column=col_map['map_group']).value if 'map_group' in col_map else None
        
        # 일자 파싱
        d_str = str(raw_date)[:10].strip()
        try:
            dt = datetime.date.fromisoformat(d_str)
        except:
            continue
            
        t_str = str(raw_time).strip() if raw_time else "00:00"
        
        time_window, sunrise_min, sunset_min = classify_time_window(dt, t_str)
        zone_id, micro_habitat = map_to_19_zones(raw_map_group, raw_region, raw_detail)
        
        count_val = 0
        try:
            count_val = int(raw_count) if raw_count is not None else 0
        except:
            count_val = 0
            
        act_str = str(raw_activity).strip() if raw_activity else ""
        if count_val == 0 or act_str == "서식확인":
            event_type = "Search_Recon"
        else:
            event_type = "Actual_Contact"
            
        is_ditch = (micro_habitat == "Warm_Water_Outflow")
        grade, cert_prob, corrected_species, reasoning = evaluate_species_certainty(
            dt, time_window, raw_species, act_str, is_ditch
        )
        
        try: shot7 = int(raw_shot7 or 0)
        except: shot7 = 0
        try: shot4 = int(raw_shot4 or 0)
        except: shot4 = 0
        try: blank = int(raw_blank or 0)
        except: blank = 0
        total_shots = shot7 + shot4 + blank
        
        recon_intensity = total_shots if event_type == "Search_Recon" else 0
        threat_score = round(count_val * cert_prob, 2) if event_type == "Actual_Contact" else 0.0
        
        record = {
            "일자": d_str,
            "시간": t_str,
            "시간대_윈도우": time_window,
            "분석_Zone": zone_id,
            "미세서식지": micro_habitat,
            "이벤트_유형": event_type,
            "원문_조류종": raw_species,
            "보정_조류종": corrected_species,
            "신뢰도_등급": grade,
            "신뢰도_확률": cert_prob,
            "개체수": count_val,
            "활동": act_str,
            "7호탄": shot7,
            "4호탄": shot4,
            "공포탄": blank,
            "총_발포발수": total_shots,
            "수색경계_강도": recon_intensity,
            "실질위협_점수": threat_score,
            "신뢰도_판정근거": reasoning,
            "원문_맵영역그룹": raw_map_group,
            "원문_지역": raw_region,
            "원문_지역상세": raw_detail
        }
        records.append(record)
        
    return records


# -------------------------------------------------------------
# 5. 메인 일괄 실행 함수 (2024년 1~12월 & 2025년 1~12월)
# -------------------------------------------------------------
def run_batch():
    out_dir = r"C:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled"
    os.makedirs(out_dir, exist_ok=True)
    
    # 2024 파일 목록
    folder_2024 = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2024년 자료\2024년 일일점검(1~12월)"
    # 2025 파일 목록
    folder_2025 = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2025년 자료\2025년 통계"
    
    all_2024_records = []
    all_2025_records = []
    
    print("==================================================")
    print("🚀 [BATCH START] 2024년 1월~12월 데이터 리샘플링")
    print("==================================================")
    for m in range(1, 13):
        # 2024년 m월 파일 찾기
        fpath = os.path.join(folder_2024, f"2024년 {m}월 일일점검.xlsx")
        if not os.path.exists(fpath):
            print(f"[MISSING] 2024년 {m}월 파일 없음: {fpath}")
            continue
            
        recs = resample_file(fpath)
        all_2024_records.extend(recs)
        df_m = pd.DataFrame(recs)
        out_m = os.path.join(out_dir, f"2024-{m:02d}_resampled.csv")
        df_m.to_csv(out_m, index=False, encoding="utf-8-sig")
        print(f"  ✅ 2024년 {m:02d}월 완료: {len(recs):,d} 건 ➜ {os.path.basename(out_m)}")

    print("\n==================================================")
    print("🚀 [BATCH START] 2025년 1월~12월 데이터 리샘플링")
    print("==================================================")
    for m in range(1, 13):
        # 2025년 m월 파일 찾기
        fpath = os.path.join(folder_2025, f"2025년{m}월 점검일지.xlsx")
        if not os.path.exists(fpath):
            print(f"[MISSING] 2025년 {m}월 파일 없음: {fpath}")
            continue
            
        recs = resample_file(fpath)
        all_2025_records.extend(recs)
        df_m = pd.DataFrame(recs)
        out_m = os.path.join(out_dir, f"2025-{m:02d}_resampled.csv")
        df_m.to_csv(out_m, index=False, encoding="utf-8-sig")
        print(f"  ✅ 2025년 {m:02d}월 완료: {len(recs):,d} 건 ➜ {os.path.basename(out_m)}")

    # 종합 저장
    df_2024 = pd.DataFrame(all_2024_records)
    df_2025 = pd.DataFrame(all_2025_records)
    df_total = pd.concat([df_2024, df_2025], ignore_index=True)
    
    path_2024_total = os.path.join(out_dir, "2024_total_resampled.csv")
    path_2025_total = os.path.join(out_dir, "2025_total_resampled.csv")
    path_master = os.path.join(out_dir, "2024_2025_master_feature_matrix.csv")
    
    df_2024.to_csv(path_2024_total, index=False, encoding="utf-8-sig")
    df_2025.to_csv(path_2025_total, index=False, encoding="utf-8-sig")
    df_total.to_csv(path_master, index=False, encoding="utf-8-sig")
    
    print("\n==================================================")
    print("🎉 [ALL BATCH COMPLETE] 전체 리샘플링 완료!")
    print("==================================================")
    print(f"총 2024년 데이터: {len(df_2024):,d} 건 ➜ {os.path.basename(path_2024_total)}")
    print(f"총 2025년 데이터: {len(df_2025):,d} 건 ➜ {os.path.basename(path_2025_total)}")
    print(f"통합 마스터 매트릭스: {len(df_total):,d} 건 ➜ {os.path.basename(path_master)}")
    
    # 기러기 출몰 월별 추이 통계
    print("\n📊 [핵심 검증] 2024 vs 2025 큰기러기 월별 개체수 추이:")
    geese_df = df_total[df_total['원문_조류종'].isin(['큰기러기', '쇠기러기', '기러기', '기러기류'])]
    geese_df['연월'] = geese_df['일자'].str[:7]
    geese_monthly = geese_df.groupby('연월')['개체수'].sum()
    print(geese_monthly)

if __name__ == "__main__":
    run_batch()
