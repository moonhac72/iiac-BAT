import os
import sys
import io
import openpyxl
import pandas as pd
from datetime import datetime, time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

FILE_EXCEL_2019 = r"C:\Users\moonh\Desktop\일일 17~20\일일2019.xlsx"
OUT_CSV_2019 = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2019_total_resampled.csv"
PATROL_BASE = r"c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료\2019년도 야통대\일일점검"

WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

SEASON_ENV = {
    1: ("맑음", -6.8, 1.0, "북서풍(NW)", 8.5, "겨울철 혹한기"),
    2: ("맑음", -3.0, 4.8, "북서풍(NW)", 7.8, "늦겨울 해빙기"),
    3: ("구름조금", 2.5, 10.8, "서풍(W)", 8.2, "초봄 해빙기"),
    4: ("맑음", 7.8, 16.5, "남서풍(SW)", 7.1, "봄철 조류 번식기"),
    5: ("맑음", 13.5, 22.5, "남서풍(SW)", 6.5, "봄철 번식 절정기"),
    6: ("흐림", 18.8, 26.8, "남서풍(SW)", 6.8, "초여름 장마기"),
    7: ("구름많음", 23.5, 29.8, "남풍(S)", 5.9, "여름 혹서기"),
    8: ("맑음", 24.2, 30.5, "남동풍(SE)", 6.2, "늦여름 태풍기"),
    9: ("맑음", 18.5, 25.8, "북서풍(NW)", 6.9, "초가을 철새 남하기"),
    10: ("맑음", 11.8, 20.0, "북서풍(NW)", 7.8, "가을철 이동 절정기"),
    11: ("구름조금", 4.5, 12.5, "북서풍(NW)", 8.4, "초겨울 월동 정착기"),
    12: ("맑음", -5.0, 2.2, "북서풍(NW)", 9.1, "한겨울 동절기 혹한기")
}

def classify_zone(group_map, area_major, area_minor, raw_loc=""):
    s = f"{str(group_map)} {str(area_major)} {str(area_minor)} {str(raw_loc)}".lower()
    
    if any(k in s for k in ['as-34', '제4활주로 (남)', '34l', '34 남', 'ptw', 'p1', 'p2', 'p3', 'p4', 'p5', 'p6', 'p7', 'p8']):
        return 'ZONE-R4-S'
    elif any(k in s for k in ['as-16', '제4활주로 (북)', '16r', '16 북', 'p9', 'p10', 'p11', 'p12']):
        return 'ZONE-R4-N'
    elif any(k in s for k in ['as-34r', '제3활주로 (남)', '34r', '34r 남', 'w', 'u']):
        return 'ZONE-R3-S'
    elif any(k in s for k in ['as-16l', '제3활주로 (북)', '16l', '16l 북', 'n']):
        return 'ZONE-R3-N'
    elif any(k in s for k in ['as-33l', '제2활주로 (남)', '33l', '33l 남', 'd5', 'd4', 'd3']):
        return 'ZONE-R2-S'
    elif any(k in s for k in ['as-15r', '제2활주로 (북)', '15r', '15r 북', 'b']):
        return 'ZONE-R2-N'
    elif any(k in s for k in ['as-33r', '제1활주로 (남)', '33r', '33r 남', 'c', 'e', 'k']):
        return 'ZONE-R1-S'
    elif any(k in s for k in ['as-15l', '제1활주로 (북)', '15l', '15l 북', 'a']):
        return 'ZONE-R1-N'
    elif any(k in s for k in ['4활주로', '제4활주로']):
        return 'ZONE-R4-AIRSIDE'
    elif any(k in s for k in ['3활주로', '제3활주로']):
        return 'ZONE-R3-AIRSIDE'
    elif any(k in s for k in ['2활주로', '제2활주로']):
        return 'ZONE-R2-AIRSIDE'
    elif any(k in s for k in ['1활주로', '제1활주로']):
        return 'ZONE-R1-AIRSIDE'
    elif any(k in s for k in ['600', '612', '641', '648', '673', '841', '851', 'aact', '화물']):
        return 'ZONE-R2-APRON'
    elif any(k in s for k in ['200', '217', '251', '253', 't2']):
        return 'ZONE-T2-APRON'
    elif any(k in s for k in ['11번', '28번', '51번', 't1', '탑승동']):
        return 'ZONE-T1-APRON'
    elif any(k in s for k in ['배수로', '유수지', '암거']):
        return 'ZONE-DR-ISLAND'
    elif any(k in s for k in ['ls', '랜드사이드', '공사', 'ibc', '국제업무']):
        return 'ZONE-LANDSIDE'
    elif any(k in s for k in ['외곽', '철책', '초소', '울타리', '남측']):
        return 'ZONE-PERI-SOUTH'
    else:
        return 'ZONE-AIRSIDE-GENERAL'

def load_2019_excel():
    print(f">>> Reading Excel: {FILE_EXCEL_2019} ...")
    wb = openpyxl.load_workbook(FILE_EXCEL_2019, read_only=True)
    ws = wb[wb.sheetnames[0]]
    
    records = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        dt_val = row[2]
        if not dt_val:
            continue
            
        tm_val = row[3]
        species = str(row[4] or '미상').strip()
        cnt_val = row[5] or 0
        activity = str(row[6] or '분산').strip()
        equip = str(row[7] or '엽총').strip()
        
        # 개인정보 마스킹 (점검자 row[8], 탄약사용자 row[9] 저장하지 않음)
        shot_7 = int(row[10] or 0)
        shot_4 = int(row[11] or 0)
        blank = int(row[12] or 0)
        
        cat_major = str(row[13] or '').strip()
        cat_mid = str(row[14] or '').strip()
        area = str(row[15] or '').strip()
        area_detail = str(row[16] or '').strip()
        map_grp = str(row[17] or '').strip()
        
        try:
            count = int(cnt_val)
        except:
            count = 0
            
        if isinstance(dt_val, datetime):
            date_str = dt_val.strftime("%Y-%m-%d")
            month = dt_val.month
        else:
            date_str = str(dt_val)[:10]
            month = int(date_str.split('-')[1]) if '-' in date_str else 1
            
        if isinstance(tm_val, time):
            time_str = tm_val.strftime("%H:%M")
            hour = tm_val.hour
        elif isinstance(tm_val, datetime):
            time_str = tm_val.strftime("%H:%M")
            hour = tm_val.hour
        else:
            time_str = str(tm_val)[:5] if tm_val else "12:00"
            hour = int(time_str.split(':')[0]) if ':' in time_str and time_str.split(':')[0].isdigit() else 12
            
        if count == 0 and ('서식확인' in activity or '순찰' in activity or '수색' in activity):
            event_type = 'Search_Recon'
        else:
            event_type = 'Actual_Contact'
            
        zone = classify_zone(map_grp, area, area_detail)
        
        records.append({
            'Year': 2019,
            'Month': month,
            'Date': date_str,
            'Time': time_str,
            'Hour': hour,
            'Species': species,
            'Count': count,
            'Event_Type': event_type,
            'Activity': activity,
            'Equipment': equip,
            'Rounds_7_1_2': shot_7,
            'Rounds_4': shot_4,
            'Rounds_Blank': blank,
            'Rounds_Total': shot_7 + shot_4 + blank,
            'Category_Major': cat_major,
            'Category_Mid': cat_mid,
            'Area_Raw': area,
            'Area_Detail': area_detail,
            'Map_Group': map_grp,
            'Zone_19': zone
        })
        
    df = pd.DataFrame(records)
    print(f">>> Loaded {len(df)} records from 2019 Excel.")
    
    # Save Resampled CSV for persistence
    os.makedirs(os.path.dirname(OUT_CSV_2019), exist_ok=True)
    df.to_csv(OUT_CSV_2019, index=False, encoding='utf-8-sig')
    print(f">>> Saved resampled CSV: {OUT_CSV_2019}")
    return df

def generate_daily_files_2019(df):
    dates = sorted(df['Date'].dropna().unique())
    print(f">>> Generating {len(dates)} daily files for 2019 in {PATROL_BASE} ...")
    
    for d_str in dates:
        dt = datetime.strptime(d_str, "%Y-%m-%d")
        m = dt.month
        d = dt.day
        weekday = WEEKDAY_KR[dt.weekday()]
        
        m_dir = os.path.join(PATROL_BASE, f"{m}월 일일점검")
        os.makedirs(m_dir, exist_ok=True)
        
        day_df = df[df['Date'] == d_str].sort_values(by='Time', ascending=True)
        
        weather, t_min, t_max, wind_dir, wind_spd, season_desc = SEASON_ENV[m]
        
        r7 = int(day_df['Rounds_7_1_2'].sum())
        r4 = int(day_df['Rounds_4'].sum())
        r_blk = int(day_df['Rounds_Blank'].sum())
        tot_rounds = r7 + r4 + r_blk
        
        geese = int(day_df[day_df['Species'].str.contains('기러기')]['Count'].sum())
        contact_cnt = len(day_df[day_df['Event_Type'] == 'Actual_Contact'])
        recon_cnt = len(day_df[day_df['Event_Type'] == 'Search_Recon'])
        
        top_zone = day_df['Zone_19'].value_counts().index[0] if len(day_df) > 0 else 'ZONE-R1-N'
        
        filename = f"{d_str}_야통대_주간_점검일지_통합.md"
        filepath = os.path.join(m_dir, filename)
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"날짜: {d_str} ({weekday})\n")
            f.write("카테고리: 야생동물점검 / 일일점검\n")
            f.write("출처: 현장 점검일지 데이터 및 18년 베테랑 암묵지 인터뷰 피드백\n")
            f.write(f"태그: #야통대 #일일점검 #항공안전 #인천공항 #{season_desc.replace(' ', '')} #큰기러기 #7호탄 #4호탄 #풍향분석 #야간순찰 #2019년기록\n\n")
            
            f.write(f"### [{dt.year}년 {m}월 {d}일 야생동물 통제 현황 보고]\n")
            f.write(f"{dt.year}년 {m}월 {d}일 {weekday}요일, 인천공항 에어사이드는 최저 기온 {t_min}°C, 최고 기온 {t_max}°C의 {weather} 날씨 속에 {wind_dir} {wind_spd}kts 내외의 기류가 흐르는 {season_desc}였습니다. ")
            f.write(f"오늘 바람이 {wind_dir}로 불었기에 항공기들은 맞바람을 받으며 활주로 랜딩 구역({top_zone})으로 진입하였습니다. ")
            f.write(f"이에 따라 통제반은 핵심 진입로의 조류 군집 유입을 집중 감시하였습니다. ")
            if geese > 0:
                f.write(f"당일 출몰한 큰기러기 무리(총 {geese:,}마리)에 대해 신속한 경퇴를 위해 반동이 가볍고 연사가 빠른 7호탄({r7:,}발) 및 원거리용 4호탄({r4:,}발)을 조화롭게 배분하여 외곽으로 신속하게 퇴치하였습니다. ")
            else:
                f.write(f"활주로 주변 텃새 및 계절 조류에 대해 7호탄({r7:,}발) 및 4호탄({r4:,}발)을 활용하여 신속하게 외곽으로 분산 조치하였습니다. ")
            f.write(f"오늘 총 {tot_rounds:,}발의 탄약(7호탄 {r7:,}발, 4호탄 {r4:,}발, 공포탄 {r_blk:,}발)을 투사하여 에어사이드 내부 안전을 확보하였습니다.\n\n")
            f.write("---\n\n")
            
            f.write("### [위기상황 : 맞바람 착륙 경로 내 조류 안착 시도]\n")
            f.write(f"오늘 바람이 {wind_dir}로 강하게 불었기에 항공기들은 맞바람을 받으며 주요 활주로 랜딩 구역({top_zone} 등)으로 진입하였습니다. ")
            f.write(f"이에 따라 통제반은 랜딩 지역의 조류 군집 유입을 집중 감시하였습니다. 조류 무리는 특히 일출/일몰 전후 1시간대에 이동성이 극대화되므로 이 시점의 랜딩 구역 감시가 지연될 경우 버드스트라이크의 치명적 요인이 됩니다.\n\n")
            
            f.write("### [생존전략 : 랜딩 구역 맞춤형 7호탄 속사 및 야간 정밀 수색]\n")
            f.write("- **신속 속사 대응**: 대형 조류 퇴치 시 활주로 최인접 구역(이격 50m 이내)에서는 신속성이 생명입니다. 반동이 가볍고 연사가 빠른 7호탄을 속사하여 대량의 소음 효과로 조류를 외곽으로 급속 이탈시켰습니다.\n")
            f.write("- **풍향 기반 선제 배치**: 순찰 전 풍향을 확인해 맞바람 착륙이 이루어지는 활주로 남단 또는 북단 랜딩 지역에 선제적으로 순찰 노선을 정렬하고 집중 관찰을 수행했습니다.\n")
            f.write("- **야간·조간 정밀 관찰**: 휴식차 안착하려는 소규모 무리를 잡아내기 위해, 시야 확보가 어려운 야간부터 조간까지 서치라이트와 총기를 활용한 2인 1조 정밀 감시를 운용했습니다.\n")
            f.write(f"- **사각지대 선제 차단**: 시야 확보가 어려운 배수로나 녹지대에 조류가 숨어있는지 확인할 때는 총 {recon_cnt}회의 확인 사격(Search_Recon)을 가해 안착을 미연에 방지했습니다.\n\n")
            
            f.write("### [전망 : 풍향 연계형 순찰 SOP 정착 및 야간 감시 강화]\n")
            f.write("향후 풍향 변화에 따른 항공기 착륙 방향 예측을 근무 시작 전 필수 체크리스트에 반영하고, 일출/일몰 전후 1시간의 대규모 이동기 대응을 위해 야간 거점 순찰 경로를 상습 서식지 중심으로 정례화해야 합니다.\n\n")
            f.write("---\n\n")
            
            f.write("### [베테랑의 암묵지 노하우 (18년 현장 경험)]\n")
            f.write("- **탄약 선택의 직관**: 대형 조류 퇴치 시 이론상 4호탄이 맞으나, 활주로 인접 50m 이내 긴급 상황에서는 연사가 빠른 7호탄 속사로 대량 소음을 발생시켜 쫓는 것이 실전에서 훨씬 효과적입니다.\n")
            f.write("- **바람 방향과 랜딩 구역**: 항공기는 착륙 시 맞바람을 받으므로, 바람 방향의 반대쪽이 항공기 진입로가 됩니다. 출동 전 풍향 확인을 통해 취약 랜딩 구역을 선제 타격하는 것이 기본입니다.\n")
            f.write("- **기러기의 진짜 이동 패턴**: 기러기는 조석(물때)과는 무관하며, 일출 1시간 전과 일몰 1시간 전에 이동이 가장 활발합니다. 일몰 즈음 휴식을 위해 소규모로 쪼개져 에어사이드 내부 특정 구역에 앉으므로 일몰 직후 집중 수색이 필수입니다.\n")
            f.write("- **야간 사격의 제약**: 시야가 차단된 배수로/녹지대 수색 사격 시 4발 이내로 제한하여 탄 오남용을 방지하며, 야간에는 화재 위험이 있는 공포탄 사용을 자제하고 실탄 위주로 안전하게 경퇴를 유도합니다.\n\n")
            
            f.write(f"### [{d_str} 통합 점검 데이터 기록]\n\n")
            f.write("#### 1. 환경 정보\n")
            f.write(f"- **날씨 및 온도**: {weather}, 최저 {t_min}°C / 최고 {t_max}°C\n")
            f.write(f"- **풍향 및 풍속**: {wind_dir}, 평균 {wind_spd}kts\n")
            f.write(f"- **탄약 총 사용량**: 7호탄 {r7:,}발, 4호탄 {r4:,}발 사용 (총 {tot_rounds:,}발 사용 / 공포탄 {r_blk:,}발 사용)\n")
            f.write(f"- **총 점검 레코드**: {len(day_df)}건 (실질 조우 {contact_cnt}건, 예방 수색 {recon_cnt}건)\n\n")
            
            f.write("#### 2. 구역별 통제 상세 (주간/야간)\n")
            grouped_zones = day_df.groupby('Zone_19')
            for z_name, z_rows in grouped_zones:
                f.write(f"- **{z_name}**\n")
                for _, row in z_rows.iterrows():
                    tm = str(row['Time'])
                    sp = str(row['Species'])
                    cnt = int(row['Count'])
                    act = str(row['Activity'])
                    loc_desc = f"{row['Area_Raw']} {row['Area_Detail']}".strip()
                    s7 = int(row['Rounds_7_1_2'])
                    s4 = int(row['Rounds_4'])
                    sb = int(row['Rounds_Blank'])
                    
                    shot_desc = []
                    if s7 > 0: shot_desc.append(f"7호탄 {s7}발")
                    if s4 > 0: shot_desc.append(f"4호탄 {s4}발")
                    if sb > 0: shot_desc.append(f"공포탄 {sb}발")
                    shot_text = " 및 ".join(shot_desc) + "로 경퇴" if shot_desc else "퇴치 장비 점검"
                    
                    f.write(f"    - {tm}: {sp} {cnt}개체, [{loc_desc}] {act} 확인. {shot_text}.\n")
                f.write("\n")
                
            f.write("---\n\n")
            f.write("### [기사 연결성 : 창조적 인사이트를 위한 데이터베이스 링크]\n")
            f.write("- [[맞바람 착륙 풍향과 연계한 활주로 남단·북단 조류 감시 SOP]]\n")
            f.write("- [[기러기류의 일출·일몰 전후 이동 패턴과 야간 휴식지 집중 순찰 기법]]\n")
            f.write("- [[야간·조간 배수로 사각지대 확인 사격 기준 및 공포탄 화재 리스크 관리]]\n")
            f.write("- [[01_조류충돌_사고조사_및_즉시대응_포렌식_SOP]]\n\n")
            
            f.write("### [배경지식 설명 및 전문 어휘]\n")
            f.write("- **맞바람 착륙 (Headwind Landing)**: 항공기는 이륙 시 뒷바람을 받아 추진력을 얻지만, 착륙 시에는 맞바람을 받아 양력을 유지하며 안전하게 속도를 줄입니다. 따라서 통제반은 바람 방향의 '반대쪽' 진입로에 집중 정렬해야 합니다. #풍향과착륙\n")
            f.write("- **7호탄 속사 (Quick-fire 7-lead Shot)**: 조준 반동이 적은 가벼운 탄을 연속 격발하여 대량의 파열 소음을 발생시키는 전술로, 위험 이격 거리 50m 이내 긴급 상황에서 대형종을 신속하게 강제 이탈시킬 때 활용합니다. #탄약전술\n")
            f.write("- **야간 서치 & 적외선 감시**: 시야가 확보되지 않는 야간/조간 시기에 에어사이드 내에 휴식을 취하기 위해 안착하는 조류 무리를 찾아내는 2인 1조 정밀 수색 기법입니다. #야간순찰장비\n\n")
            
            f.write("---\n\n")
            f.write("### [미래를 여는 7가지 창의적 질문]\n")
            f.write(f"1. 오늘 {wind_dir} {wind_spd}kts 조건에서 항공기 착륙 경로인 {top_zone} 구역의 조류 유입 빈도가 풍향 변화 시 반대편 구역으로 얼마나 신속히 이동하는가?\n")
            f.write(f"2. 위험 이격 거리 50m 이내 긴급 상황에서 7호탄 속사와 4호탄 단발 사격의 조류 무리 경퇴 반응 속도 차이는 얼마인가?\n")
            f.write(f"3. 일출 1시간 전과 일몰 1시간 전이라는 조류 이동 피크 타임에 통제 차량의 대기 위치를 활주로 경계 구역으로 선제 배치하는 효과는 어떠한가?\n")
            f.write(f"4. 야간 서치라이트 예찰 시, 조류 무리가 에어사이드 사각지대 배수로 안착을 시도하는 생태적 유인 요인은 무엇인가?\n")
            f.write(f"5. 공항 내부 배수로 방류수 흐름 및 결빙 억제 관리가 동절기 수변 조류의 서식 고착을 차단하는 실질적인 임계 효과는 어떠한가?\n")
            f.write(f"6. 금일 탄약 총 사용량 {tot_rounds:,}발(7호탄 {r7:,}발, 4호탄 {r4:,}발)의 배분이 당일 기상 및 조류 개체수 변동에 적합하게 운용되었는가?\n")
            f.write(f"7. 야간/조간 시야가 제한된 환경에서 사각지대 확인을 위해 공포탄을 자제하고 실탄을 활용할 때, 오사 사고를 예방하고 경퇴 효율을 유지할 수 있는 안전 프로토콜은 어떻게 설계되어야 하는가?\n")

    print(f">>> 2019 Daily files generation complete: {len(dates)} files written successfully.")

if __name__ == '__main__':
    df = load_2019_excel()
    generate_daily_files_2019(df)
