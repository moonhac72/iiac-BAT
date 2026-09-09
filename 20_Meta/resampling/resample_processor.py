import os
import math
import datetime
import openpyxl
import pandas as pd

# -------------------------------------------------------------
# 1. 영종도 천문 일출 / 일몰 계산기 (Yeongjongdo Solar Engine)
#    위도: 37.46 N, 경도: 126.44 E, KST (UTC+9)
# -------------------------------------------------------------
def get_solar_times(dt: datetime.date):
    """
    영종도 천문학적 일출/일몰 시각(분 단위, 00:00 기준) 계산
    """
    lat = 37.46
    lon = 126.44
    day_of_year = dt.timetuple().tm_yday
    
    # 태양 적위 (declination)
    gamma = 2 * math.pi / 365 * (day_of_year - 1)
    decl = 0.006918 - 0.399912 * math.cos(gamma) + 0.070257 * math.sin(gamma) \
           - 0.006758 * math.cos(2*gamma) + 0.000907 * math.sin(2*gamma)
    
    # 균시차 (equation of time, in minutes)
    eqtime = 229.18 * (0.000075 + 0.001868 * math.cos(gamma) - 0.032077 * math.sin(gamma) \
             - 0.014615 * math.cos(2*gamma) - 0.040849 * math.sin(2*gamma))
    
    # 태양 시각각 (hour angle, zenith = 90.833 deg for standard sunrise/sunset)
    zenith = math.radians(90.833)
    lat_rad = math.radians(lat)
    
    cos_ha = (math.cos(zenith) / (math.cos(lat_rad) * math.cos(decl))) - (math.tan(lat_rad) * math.tan(decl))
    cos_ha = max(-1.0, min(1.0, cos_ha))
    ha = math.degrees(math.acos(cos_ha))
    
    # 태양 정오 (Solar Noon in KST minutes)
    solar_noon = 720 - 4 * lon - eqtime + 9 * 60
    
    sunrise_min = solar_noon - ha * 4
    sunset_min = solar_noon + ha * 4
    
    return sunrise_min, sunset_min

def classify_time_window(dt: datetime.date, time_str: str):
    """
    일자와 시간(HH:MM)을 바탕으로 4대 시간 윈도우 판정:
    - Dawn: 일출 전후 1.5시간 (90분)
    - Dusk: 일몰 전후 1.5시간 (90분)
    - Day: (일출 + 90분) ~ (일몰 - 90분)
    - Night: 그 외 야간
    """
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
    """
    맵영역그룹, 지역, 지역상세 정보를 기반으로 19개 분석 Zone 및 배수로 속성 매핑
    """
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
    # 제4활주로 (34L ↔ 16R)
    if "4활주로" in rg or "4R/W" in rg or "4RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R4-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R4-N", micro_habitat
        return "ZONE-R4-M", micro_habitat
    # 제3활주로 (34R ↔ 16L)
    if "3활주로" in rg or "3R/W" in rg or "3RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R3-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R3-N", micro_habitat
        return "ZONE-R3-M", micro_habitat
    # 제2활주로 (33L ↔ 15R)
    if "2활주로" in rg or "2R/W" in rg or "2RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R2-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R2-N", micro_habitat
        return "ZONE-R2-M", micro_habitat
    # 제1활주로 (33R ↔ 15L)
    if "1활주로" in rg or "1R/W" in rg or "1RW" in rg:
        if "남" in dt or "S" in dt: return "ZONE-R1-S", micro_habitat
        if "북" in dt or "N" in dt: return "ZONE-R1-N", micro_habitat
        return "ZONE-R1-M", micro_habitat
        
    # 3) 맵영역그룹 기본 매핑
    if "AS-15" in mg:
        # 북단 1/2활주로 구역
        return "ZONE-R1-N", micro_habitat
    elif "AS-16" in mg:
        # 북단 3/4활주로 구역
        return "ZONE-R4-N", micro_habitat
    elif "AS-33" in mg:
        # 남단 1/2활주로 구역
        return "ZONE-R1-S", micro_habitat
    elif "AS-34" in mg:
        # 남단 3/4활주로 구역
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
    """
    5단계 신뢰도 매트릭스 적용:
    - Grade 1 (Verified): 포획, 사체수거
    - Grade 2 (Certain): 야행성 고유종 (100%)
    - Grade 3 (High): 주간/일출/일몰 또는 2025년 4월 이전 기러기
    - Grade 4 (Probabilistic 0.5): 야간 배수로 오리류 (50% 가중치)
    - Grade 5 (Doubtful): 야간 식별 불가 일반 주행성 조류 -> 미식별 블라인드 처리
    """
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
        return "Grade 4 (Probabilistic)", 0.5, sp, "야간 배수로 오리류 (임의작성 경향 반영 50% 가중치)"
    
    # 야간 일반 주행성 조류 (까치, 까마귀, 멧비둘기, 백로 등)
    if sp not in ['미상', 'None', '']:
        return "Grade 5 (Doubtful)", 0.1, "미식별(야간_주행성_오인의심)", f"야간 식별 취약 주행성 조류({sp}) 노이즈 블라인드"
    
    return "Grade 5 (Doubtful)", 0.1, "미식별", "야간 미상"


# -------------------------------------------------------------
# 4. 엑셀 파일 리샘플링 실행 함수
# -------------------------------------------------------------
def resample_excel_file(file_path: str, output_csv_path: str):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb['Sheet1']
    
    records = []
    
    for r in range(2, ws.max_row + 1):
        raw_date = ws.cell(row=r, column=2).value
        raw_time = ws.cell(row=r, column=3).value
        raw_species = ws.cell(row=r, column=4).value
        raw_count = ws.cell(row=r, column=5).value
        raw_activity = ws.cell(row=r, column=6).value
        raw_equip = ws.cell(row=r, column=7).value
        raw_shot7 = ws.cell(row=r, column=8).value or 0
        raw_shot4 = ws.cell(row=r, column=9).value or 0
        raw_blank = ws.cell(row=r, column=10).value or 0
        raw_major = ws.cell(row=r, column=11).value
        raw_mid = ws.cell(row=r, column=12).value
        raw_region = ws.cell(row=r, column=13).value
        raw_detail = ws.cell(row=r, column=14).value
        raw_map_group = ws.cell(row=r, column=15).value
        
        if not raw_date:
            continue
            
        # 일자 파싱
        d_str = str(raw_date)[:10].strip()
        try:
            dt = datetime.date.fromisoformat(d_str)
        except:
            continue
            
        t_str = str(raw_time).strip() if raw_time else "00:00"
        
        # 1) 시간 윈도우 계산
        time_window, sunrise_min, sunset_min = classify_time_window(dt, t_str)
        
        # 2) 19개 Zone 매핑
        zone_id, micro_habitat = map_to_19_zones(raw_map_group, raw_region, raw_detail)
        
        # 3) 이벤트 유형 분리 (수색 사격 vs 실제 조우)
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
            
        # 4) 조류종 5단계 신뢰도 평가
        is_ditch = (micro_habitat == "Warm_Water_Outflow")
        grade, cert_prob, corrected_species, reasoning = evaluate_species_certainty(
            dt, time_window, raw_species, act_str, is_ditch
        )
        
        # 5) 탄약 및 강도 집계
        try: shot7 = int(raw_shot7)
        except: shot7 = 0
        try: shot4 = int(raw_shot4)
        except: shot4 = 0
        try: blank = int(raw_blank)
        except: blank = 0
        total_shots = shot7 + shot4 + blank
        
        # 순찰 경계 지수 (수색사격용) vs 충돌 위협 점수 (실제 조우용)
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
        
    df = pd.DataFrame(records)
    df.to_csv(output_csv_path, index=False, encoding="utf-8-sig")
    print(f"Resampling Complete! Total Processed: {len(df)} rows")
    return df

if __name__ == "__main__":
    src_file = r"C:\Users\moonh\Desktop\Anythingllm 자료폴더\2024년 자료\2024년 일일점검(1~12월)\2024년 1월 일일점검.xlsx"
    out_file = r"C:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2024-01_resampled.csv"
    resample_excel_file(src_file, out_file)
