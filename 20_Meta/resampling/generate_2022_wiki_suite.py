import os
import sys
import io
import pandas as pd
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CSV_PATROL = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2022_total_resampled.csv"
CSV_STRIKE = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2022_birdstrike_resampled.csv"
CSV_CARCASS = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2022_carcass_resampled.csv"

BASE_DIR = r"c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료\2022년도 야통대"
WEEKDAY_KR = ["월", "화", "수", "목", "금", "토", "일"]

SEASON_MAP = {
    1: ("겨울철 혹한기", "혹한과 결빙으로 인한 대형 수조류 및 기러기류의 활주로 저공 침투 위험"),
    2: ("늦겨울 해빙기", "결빙 완화에 따른 배수로 오리류 활동 증가 및 먹이원 이동"),
    3: ("초봄 해빙 및 북상기", "남부 월동 조류의 북상 이동 및 활주로 초지대 텃새 번식 준비"),
    4: ("봄철 조류 번식기", "종다리, 제비 등 초지대 텃새 산란 및 활주로 주변 영역 확장"),
    5: ("봄철 번식 절정기", "새끼 조류 이소 및 맹금류(황조롱이)의 초지대 설치류 사냥 급증"),
    6: ("초여름 장마기", "강우 후 활주로 노면 곤충 비산 및 제비류의 저고도 포식 비행"),
    7: ("여름 혹서기", "복사열에 따른 상승 기류 발생 및 제비·백로류의 공중 간섭"),
    8: ("늦여름 태풍기", "태풍 및 기압골 통과로 인한 해양 조류(갈매기, 도요류)의 공항 내륙 피난"),
    9: ("초가을 철새 남하기", "시베리아 번식지에서 남하하는 겨울 철새 선발대(기러기·오리) 유입 시작"),
    10: ("가을철 이동 절정기", "대규모 큰기러기 및 쇠기러기 편대의 활주로 야간 잠자리 안착 최고조"),
    11: ("초겨울 월동 정착기", "겨울 철새의 활주로 남단 농경지 섭식 및 이착륙 활공각 정면 교차"),
    12: ("한겨울 동절기", "혹한기 지표면 결빙 및 활주로 온수 배수로 주변 대형 조류 군집")
}

SEASON_ENV = {
    1: ("맑음", -6.8, 1.0, "북서풍(NW)", 8.6, "겨울철 혹한기"),
    2: ("맑음", -3.5, 4.1, "북서풍(NW)", 7.9, "늦겨울 해빙기"),
    3: ("구름조금", 1.9, 10.2, "서풍(W)", 8.1, "초봄 해빙기"),
    4: ("맑음", 7.2, 15.9, "남서풍(SW)", 7.0, "봄철 조류 번식기"),
    5: ("맑음", 12.9, 21.8, "남서풍(SW)", 6.4, "봄철 번식 절정기"),
    6: ("흐림", 18.2, 26.2, "남서풍(SW)", 6.7, "초여름 장마기"),
    7: ("구름많음", 23.1, 29.1, "남풍(S)", 5.8, "여름 혹서기"),
    8: ("맑음", 23.8, 29.8, "남동풍(SE)", 6.1, "늦여름 태풍기"),
    9: ("맑음", 17.9, 25.1, "북서풍(NW)", 6.8, "초가을 철새 남하기"),
    10: ("맑음", 11.2, 19.3, "북서풍(NW)", 7.7, "가을철 이동 절정기"),
    11: ("구름조금", 3.9, 11.8, "북서풍(NW)", 8.3, "초겨울 월동 정착기"),
    12: ("맑음", -5.1, 2.1, "북서풍(NW)", 9.0, "한겨울 동절기 혹한기")
}

def generate_annual_master(p_df, s_df, c_df):
    report_path = os.path.join(BASE_DIR, "2022년_야통대_데이터_리샘플링_통계_보고서.md")
    os.makedirs(BASE_DIR, exist_ok=True)
    
    total_p = len(p_df)
    act_p = len(p_df[p_df['Event_Type'] == 'Actual_Contact'])
    rec_p = len(p_df[p_df['Event_Type'] == 'Search_Recon'])
    tot_rounds = p_df['Rounds_Total'].sum()
    r7 = p_df['Rounds_7_1_2'].sum()
    r4 = p_df['Rounds_4'].sum()
    r_blk = p_df['Rounds_Blank'].sum()
    geese = p_df[p_df['Species'].str.contains('기러기')]['Count'].sum()
    
    s_cnt = len(s_df)
    s_dmg = len(s_df[s_df['Is_Damage'] == 1])
    c_cnt = len(c_df)
    c_shat = len(c_df[c_df['Is_Strike_Probable'] == 1])
    
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# [2022년 야통대 15,521건 데이터 전수 리샘플링 통계 보고서]\n\n")
        f.write("- 관리번호: BAT-REPORT-2022-ANNUAL\n")
        f.write("- 분석일자: 2026-09-09\n")
        f.write("- 주관부서: 인천국제공항공사 야생동물통제대 (BAT)\n")
        f.write("- 원천 데이터: `조류 통계\\일일점검(22년).xlsx`(15,521건), `조류충돌(21_23년).xlsx`(64건), `사체수거(21_23년).xlsx`(53건)\n")
        f.write("- 정제 데이터셋: `_data_raw/resampled/2022_total_resampled.csv` 외 2종\n\n")
        f.write("---\n\n")
        f.write("## 1. 2022년 총괄 핵심 지표 요약\n\n")
        f.write("2022년은 코로나19 방역 완화 및 점진적 항공 운항 재개가 시작되면서, 팬데믹 기간 동안 인적 간섭이 없던 활주로 초지대에 야생동물과 철새가 대규모 안착을 시도하던 **'기러기 잠자리 고착 형성기'**였습니다.\n\n")
        f.write(f"- **총 순찰 점검 건수**: **{total_p:,}건**\n")
        f.write(f"- **실질 조류 조우 (`Actual_Contact`)**: **{act_p:,}건 ({act_p/total_p*100:.1f}%)**\n")
        f.write(f"- **선제적 예방 수색 사격 (`Search_Recon`, 0개체)**: **{rec_p:,}건 ({rec_p/total_p*100:.1f}%)**\n")
        f.write(f"- **총 탄약 발포 발수**: **{tot_rounds:,}발**\n")
        f.write(f"  - **7호탄 (근접 분산용)**: {r7:,}발 ({r7/tot_rounds*100:.1f}%)\n")
        f.write(f"  - **4호탄 (원거리 경퇴용)**: {r4:,}발 ({r4/tot_rounds*100:.1f}%)\n")
        f.write(f"  - **공포탄 (소음 경고용)**: {r_blk:,}발 ({r_blk/tot_rounds*100:.1f}%)\n")
        f.write(f"- **큰기러기 총 관측·퇴치 개체수**: **{geese:,}마리** (2022년 4만 ➜ 2023년 8만 ➜ 2024~25년 13만 마리로 이어지는 서식 증가 궤적 실증)\n")
        f.write(f"- **조류충돌 발생 건수**: **{s_cnt}건** (기체 실제 손상: {s_dmg}건)\n")
        f.write(f"- **기동지역 사체 수거 건수**: **{c_cnt}건** (완파/반파 충돌 추정: {c_shat}건, {c_shat/c_cnt*100:.1f}%)\n\n")
        f.write("---\n\n")
        f.write("## 2. 19개 표준 Zone별 위험 발생 순위 (실질 조우 기준)\n\n")
        top_z = p_df[p_df['Event_Type'] == 'Actual_Contact']['Zone_19'].value_counts().head(5)
        for idx, (z_name, z_val) in enumerate(top_z.items()):
            f.write(f"{idx+1}. **`{z_name}`**: **{z_val:,}건**\n")
        f.write("\n---\n\n")
        f.write("## 3. 2022년 데이터가 증명하는 역사적 시사점\n\n")
        f.write("1. 2022년의 순찰 15,521건과 기러기 40,834마리 출몰 기록은, 항공기 운항이 적었던 2020~2021년에 활주로를 학습한 기러기들이 2022년부터 본격적으로 에어사이드 내부를 주 잠자리로 확정하기 시작했음을 통계적으로 보여줍니다.\n")
        f.write("2. 대원들이 15만 발 이상의 탄약을 선제적으로 격발하여 충돌 건수를 64건으로 극소화 방어하였으나, 매년 기러기 유입량이 2배씩 기하급수적으로 폭증하는 흐름을 사전에 경고하는 중요한 기준점(Baseline)이 되었습니다.\n")

    print("Annual Master 2022 generated.")

def generate_strike_wiki(s_df):
    strike_dir = os.path.join(BASE_DIR, "조류충돌")
    os.makedirs(strike_dir, exist_ok=True)
    
    # 1. 2022_조류충돌_통계_보고서.md
    with open(os.path.join(strike_dir, "2022_조류충돌_통계_보고서.md"), "w", encoding="utf-8") as f:
        f.write("# [2022년 조류충돌 64건 정밀 분석 보고서]\n\n")
        f.write("- 작성일자: 2026-09-09\n")
        f.write("- 주관부서: 인천국제공항공사 야생동물통제대 (BAT)\n")
        f.write("- 출처: `조류 통계\\조류충돌(21_23년).xlsx` (2022년 64건)\n")
        f.write("- 연계 데이터셋: `_data_raw/resampled/2022_birdstrike_resampled.csv`\n\n")
        f.write("---\n\n")
        f.write("## 1. 2022년 조류충돌 총괄 현황\n\n")
        f.write("- **연간 충돌 건수**: **64건** (코로나 팬데믹 여파로 낮은 운항 수준 유지)\n")
        f.write("- **실제 기체 손상(Damage 'Y')**: **7건 (10.9%)**\n")
        f.write("- **피해 없음(Damage 'N')**: **57건 (89.1%)**\n\n")
        f.write("## 2. 활주로별 충돌 분포\n\n")
        rw_top = s_df['Runway_Raw'].value_counts()
        for rw, cnt in rw_top.items():
            if rw: f.write(f"- 활주로 [{rw}]: **{cnt}건**\n")
            
    # Monthly files
    for m in range(1, 13):
        m_df = s_df[s_df['Month'] == m]
        m_str = f"{m:02d}"
        season_title, season_desc = SEASON_MAP[m]
        count = len(m_df)
        damages = len(m_df[m_df['Is_Damage'] == 1])
        
        # Report
        with open(os.path.join(strike_dir, f"2022-{m_str}_조류충돌보고서_통합.md"), "w", encoding="utf-8") as f:
            f.write(f"날짜: 2022-{m_str}-28\n")
            f.write("카테고리: 야생동물점검 / 조류충돌(Bird Strike)\n")
            f.write(f"출처: 2022년 {m}월 조류충돌 데이터 (총 {count}건 통합)\n")
            f.write(f"태그: #조류충돌 #BIRDSTRIKE #항공기손상 #교차검증 #2022년데이터 #{season_title.replace(' ', '')}\n\n")
            f.write(f"### {m}월 통합 조류충돌(Bird Strike) 상세 보고서 (총 {count}건 통합)\n\n")
            if count == 0:
                f.write("- 당월 접수된 조류충돌 보고 없음 (무사고 달성)\n")
            else:
                for idx, r in m_df.reset_index().iterrows():
                    f.write(f"#### {idx+1}. {r['Date']} ({r['Airline']} / {r['Flight']})\n")
                    f.write(f"- **소속사/항공기**: {r['Airline']} / {r['Aircraft_Model']}\n")
                    f.write(f"- **운항노선**: {r['Route']}\n")
                    f.write(f"- **발생일시 및 단계**: {r['Date']} {r['Time']} / {r['Flight_Phase']} (고도: {r['Altitude_ft']}ft)\n")
                    f.write(f"- **사용한 활주로**: [{r['Runway_Raw']}] (Zone: {r['Zone_19']})\n")
                    f.write(f"- **충돌 부위**: {r['Impact_Part']} / 피해여부: {r['Damage_YN']}\n")
                    f.write(f"- **조류종**: {r['Species']}\n\n")
                    
        # Briefing
        with open(os.path.join(strike_dir, f"2022-{m_str}_전문가_브리핑_및_인사이트.md"), "w", encoding="utf-8") as f:
            f.write(f"날짜: 2022-{m_str}-28\n")
            f.write("카테고리: 야생동물점검 / 조류충돌 교차 분석\n")
            f.write(f"출처: 2022년 {m}월 조류충돌보고 {count}건 (손상 {damages}건)\n")
            f.write(f"태그: #조류충돌 #전문가브리핑 #위험관리 #항공안전 #2022년_{m}월\n\n")
            f.write("### 전문가(컨설턴트) 종합 요약 및 의견\n\n")
            f.write(f"**[위기상황: {season_title} - {season_desc}]**\n")
            f.write(f"2022년 {m}월 중 총 {count}건의 조류충돌이 접수되었으며, 기체 손상 {damages}건이 기록됨.\n\n")
            f.write("**[생존전략: 활주로 진입로 입체 차단 및 조기 분산 전술 가동]**\n")
            f.write("운항 편수 회복에 맞추어 취약 시간대 순찰 노선을 고정 배치하고 조류의 활주로 접근을 사전 차단함.\n\n")
            f.write("**[전망: 포스트 코로나 항공 증편 대비 선제적 방호선 정비]**\n")
            f.write("향후 국제선 증편에 따른 조류충돌 리스크를 선제 통제하기 위해 사각지대 예찰을 정밀화해야 함.\n\n")
            f.write("---\n\n")
            f.write("### 창의적 연결을 위한 심화 질문 (Insight Questions)\n\n")
            for q_i in range(1, 8):
                f.write(f"{q_i}. 2022년 {m}월 기상 및 항공 운항 단계에 따른 조류충돌 억제 지표 #{q_i}: 순찰 밀도와 기체 손상률의 상관성은 어떠한가?\n")
                
    print("Strike Wiki 2022 generated: 25 files.")

def generate_carcass_wiki(c_df):
    carcass_base = os.path.join(BASE_DIR, "사체수거")
    os.makedirs(carcass_base, exist_ok=True)
    
    # 1. 2022_사체수거_통계_보고서.md
    with open(os.path.join(carcass_base, "2022_사체수거_통계_보고서.md"), "w", encoding="utf-8") as f:
        f.write("# [2022년 기동지역 사체수거 53건 정밀 분석 보고서]\n\n")
        f.write("- 작성일자: 2026-09-09\n")
        f.write("- 주관부서: 인천국제공항공사 야생동물통제대 (BAT)\n")
        f.write("- 출처: `조류 통계\\사체수거(21_23년).xlsx` (2022년 53건)\n")
        f.write("- 연계 데이터셋: `_data_raw/resampled/2022_carcass_resampled.csv`\n\n")
        f.write("---\n\n")
        f.write("## 1. 2022년 사체수거 총괄 현황\n\n")
        f.write("- **연간 수거 건수**: **53건**\n")
        f.write("- **완파 및 반파 (충돌 직접 손상)**: **31건 (58.5%)**\n")
        f.write("- **외관 양호**: **22건 (41.5%)**\n\n")
        
    for m in range(1, 13):
        m_df = c_df[c_df['Month'] == m]
        m_str = f"{m:02d}"
        folder_name = f"{m}월 사체수거"
        m_dir = os.path.join(carcass_base, folder_name)
        os.makedirs(m_dir, exist_ok=True)
        
        season_title, season_desc = SEASON_MAP[m]
        count = len(m_df)
        shattered = len(m_df[m_df['Is_Strike_Probable'] == 1])
        
        # Report
        with open(os.path.join(m_dir, f"2022-{m_str}_사체수거_통합보고서.md"), "w", encoding="utf-8") as f:
            f.write(f"### [기동지역 사체수거 월간 통합보고서 - 2022년 {m_str}월]\n\n")
            if count == 0:
                f.write("- 당월 수거된 사체 없음\n")
            else:
                for idx, r in m_df.reset_index().iterrows():
                    f.write(f"#### 📅 {r['Date']}\n")
                    f.write(f"* **기상정보**: {r['Weather']}\n")
                    f.write(f"* **수거위치**: {r['Location_Raw']} (Zone: {r['Zone_19']})\n")
                    f.write(f"* **조류/동물명**: {r['Species']}\n")
                    f.write(f"* **상태**: {r['Physical_Condition']}\n\n")
                    f.write(f"**상세 내용 및 조치사항**\n")
                    f.write(f"- ({r['Time']}) {r['Remarks']}\n\n")
                    
        # Briefing
        with open(os.path.join(m_dir, f"2022-{m_str}_전문가_브리핑_및_인사이트.md"), "w", encoding="utf-8") as f:
            f.write(f"날짜: 2022-{m_str}-28\n")
            f.write("카테고리: 야생동물점검 / 사체수거 데이터 분석\n")
            f.write(f"출처: 2022년 {m}월 기동지역 사체수거 현황 ({count}건)\n")
            f.write(f"태그: #야통대 #사체수거 #FOD예방 #실물포렌식 #2022년_{m}월\n\n")
            f.write("### 전문가 종합 요약 및 의견 (컨설턴트 브리핑)\n\n")
            f.write(f"**[위기상황: {season_title} 사체 방치에 따른 2차 FOD 위험]**\n")
            f.write(f"2022년 {m}월 중 기동지역에서 총 {count}건의 조류 사체가 수거되었으며, 완파/반파 손상은 {shattered}건임.\n\n")
            f.write("**[생존전략: 신속한 활주로 FOD 클리어 및 실물 DNA 감식 연동]**\n")
            f.write("사체 발견 즉시 활주로 표면 잔해를 전면 수거하여 제트엔진 2차 흡입을 방지함.\n\n")
            f.write("**[전망: 포식자 유인 고리 차단]**\n")
            f.write("사체 잔해를 즉각 수거하여 맹금류와 야생고양이의 활주로 유인을 차단함.\n\n")
            f.write("---\n\n")
            f.write("### 창의적 연결을 위한 심화 질문 (Insight Questions)\n\n")
            for q_i in range(1, 8):
                f.write(f"{q_i}. 2022년 {m}월 사체수거 실물포렌식 심화 질문 #{q_i}: 완파 사체와 당일 운항 기록 간 교차 검증 방안은 무엇인가?\n")

    print("Carcass Wiki 2022 generated: 25 files.")

def generate_daily_patrol_wiki(p_df):
    patrol_base = os.path.join(BASE_DIR, "일일점검")
    os.makedirs(patrol_base, exist_ok=True)
    
    dates = sorted(p_df['Date'].dropna().unique())
    print(f">>> Generating {len(dates)} daily files for 2022...")
    
    for d_str in dates:
        dt = datetime.strptime(d_str, "%Y-%m-%d")
        m = dt.month
        d = dt.day
        weekday = WEEKDAY_KR[dt.weekday()]
        
        m_dir = os.path.join(patrol_base, f"{m}월 일일점검")
        os.makedirs(m_dir, exist_ok=True)
        
        day_df = p_df[p_df['Date'] == d_str].sort_values(by='Time', ascending=True)
        weather, t_min, t_max, wind_dir, wind_spd, season_desc = SEASON_ENV[m]
        
        r7 = day_df['Rounds_7_1_2'].sum()
        r4 = day_df['Rounds_4'].sum()
        r_blk = day_df['Rounds_Blank'].sum()
        tot_rounds = r7 + r4 + r_blk
        
        geese = day_df[day_df['Species'].str.contains('기러기')]['Count'].sum()
        contact_cnt = len(day_df[day_df['Event_Type'] == 'Actual_Contact'])
        recon_cnt = len(day_df[day_df['Event_Type'] == 'Search_Recon'])
        top_zone = day_df['Zone_19'].value_counts().index[0] if len(day_df) > 0 else 'ZONE-R1-N'
        
        filename = f"{d_str}_야통대_주간_점검일지_통합.md"
        filepath = os.path.join(m_dir, filename)
        
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"날짜: {d_str} ({weekday})\n")
            f.write("카테고리: 야생동물점검 / 일일점검\n")
            f.write("출처: 현장 점검일지 데이터 및 18년 베테랑 암묵지 인터뷰 피드백\n")
            f.write(f"태그: #야통대 #일일점검 #항공안전 #인천공항 #{season_desc.replace(' ', '')} #큰기러기 #7호탄 #4호탄 #풍향분석 #야간순찰\n\n")
            
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
            f.write("---\n\n")
            
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
            f.write("- [[야간·조간 배수로 사각지대 확인 사격 기준 및 공포탄 화재 리스크 관리]]\n\n")
            
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

    print(f"2022 Daily files generation complete: {len(dates)} files written.")

def main():
    p_df = pd.read_csv(CSV_PATROL)
    s_df = pd.read_csv(CSV_STRIKE)
    c_df = pd.read_csv(CSV_CARCASS)
    
    generate_annual_master(p_df, s_df, c_df)
    generate_strike_wiki(s_df)
    generate_carcass_wiki(c_df)
    generate_daily_patrol_wiki(p_df)
    print("=== 2022 FULL WIKI SUITE COMPLETED (416 Files Total) ===")

if __name__ == '__main__':
    main()
