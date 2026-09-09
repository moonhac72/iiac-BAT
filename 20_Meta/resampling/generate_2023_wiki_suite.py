import os
import sys
import io
import pandas as pd
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

CSV_PATROL = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2023_total_resampled.csv"
CSV_STRIKE = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2023_birdstrike_resampled.csv"
CSV_CARCASS = r"c:\Users\moonh\Documents\iiac-BAT\_data_raw\resampled\2023_carcass_resampled.csv"

BASE_DIR = r"c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료\2023년도 야통대"

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

def generate_strike_wiki(s_df):
    strike_dir = os.path.join(BASE_DIR, "조류충돌")
    os.makedirs(strike_dir, exist_ok=True)
    
    for m in range(1, 13):
        m_df = s_df[s_df['Month'] == m]
        m_str = f"{m:02d}"
        season_title, season_desc = SEASON_MAP[m]
        count = len(m_df)
        damages = len(m_df[m_df['Is_Damage'] == 1])
        
        # 1. 2023-XX_조류충돌보고서_통합.md
        report_path = os.path.join(strike_dir, f"2023-{m_str}_조류충돌보고서_통합.md")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"날짜: 2023-{m_str}-28\n")
            f.write(f"카테고리: 야생동물점검 / 조류충돌(Bird Strike)\n")
            f.write(f"출처: 2023년 {m}월 조류충돌 데이터 (총 {count}건 통합)\n")
            f.write(f"태그: #조류충돌 #BIRDSTRIKE #항공기손상 #교차검증 #2023년데이터 #{season_title.replace(' ', '')}\n\n")
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
                    f.write(f"- **조류종**: {r['Species']}\n")
                    f.write(f"- **기동지역 교차검증**: *{'실제 기체 피해 발생 중대 사건' if r['Damage_YN']=='Y' else '경미 또는 사후 정비창 확인 사건'}*\n\n")
        
        # 2. 2023-XX_전문가_브리핑_및_인사이트.md
        briefing_path = os.path.join(strike_dir, f"2023-{m_str}_전문가_브리핑_및_인사이트.md")
        with open(briefing_path, "w", encoding="utf-8") as f:
            f.write(f"날짜: 2023-{m_str}-28\n")
            f.write(f"카테고리: 야생동물점검 / 조류충돌 교차 분석\n")
            f.write(f"출처: 2023년 {m}월 조류충돌보고 {count}건 (손상 {damages}건)\n")
            f.write(f"태그: #조류충돌 #전문가브리핑 #위험관리 #항공안전 #2023년_{m}월\n\n")
            f.write("### 전문가(컨설턴트) 종합 요약 및 의견\n\n")
            f.write(f"**[위기상황: {season_title} - {season_desc}]**\n")
            f.write(f"2023년 {m}월 중 총 {count}건의 조류충돌이 접수되었으며, 이 중 실제 기체 손상을 동반한 중대 사건은 {damages}건으로 분석됨. ")
            if damages > 0:
                f.write(f"특히 엔진 및 날개 주요 부품에 타격을 입힌 사례가 포함되어 있어, 저고도 이착륙 활공각에서의 대형 조류 차단이 절실한 시점이었음.\n\n")
            else:
                f.write(f"대부분의 충돌이 피해 없는 경미한 사례로 종결되었으나 사후 정비창 확인 비중이 높아 잠재적 리스크 모니터링이 요구됨.\n\n")
            f.write(f"**[생존전략: 활주로 진입로 입체 차단 및 조기 분산 전술 가동]**\n")
            f.write(f"충돌 빈발 활주로에 순찰차량을 거점 배치하고, 일출 및 일몰 전후 취약 시간대에 집중 음향·화약 퇴치 장비를 연속 가동하여 조류 접근로를 사전 차단함.\n\n")
            f.write(f"**[전망: 계절성 생태 주기 예측 및 기체 정비 연계 강화]**\n")
            f.write(f"향후 계절 변화에 따른 조류 섭식 및 이동 패턴을 사전에 예측하여 활주로 초지대 제초 주기와 배수로 차단 스크린을 선제적으로 재정비해야 함.\n\n")
            f.write("---\n\n")
            f.write("### 창의적 연결을 위한 심화 질문 (Insight Questions)\n\n")
            f.write(f"1. **{m}월 주풍 방향과 항공기 이착륙 활주로 배정 간 조류충돌 상관성**: {m}월의 기상 특성이 조류의 양력 획득 방향과 항공기 이륙선 교차에 미친 물리적 영향은 무엇인가?\n")
            f.write(f"2. **손상 사건({damages}건) 발생 시점의 지상 순찰 대원 배치 간격**: 충돌 발생 직전 30분간 해당 활주로의 순찰 및 퇴치 사격 이력과의 정합성은 어떠한가?\n")
            f.write(f"3. **{m}월 주요 출몰 조류종의 비행 고도 특성과 조종사 시야 확보율**: 접근 및 상승 단계에서 조종사가 조류 무리를 육안으로 회피할 수 있는 최소 시간 한계는 얼마인가?\n")
            f.write(f"4. **사체 미수거 충돌의 궤적 역산 모델**: 항공기 충격 후 사체가 활주로 밖 초지대로 비산될 확률을 풍속과 비행 속도 기반으로 시뮬레이션할 수 있는가?\n")
            f.write(f"5. **공항 인근 농경지 및 유수지 생태 환경 변화의 영향**: {m}월 공항 외부 환경이 내부 조류 유입 압력(Pressure)에 미치는 기여도는 얼마인가?\n")
            f.write(f"6. **항공사 정비창 미세 혈흔 보고와 지상 야통대 관제 간의 실시간 데이터 핫라인 구축 방안**: 사후 발견 건수를 실시간 지상 예찰에 환류하는 시스템의 최적 구조는 무엇인가?\n")
            f.write(f"7. **{season_title} 맞춤형 탄약 소모 전략**: 7호탄(근접)과 4호탄(원거리)의 발포 비율을 {m}월 조류 행동 반경에 맞게 최적화하는 수리학적 모델은 어떻게 설계되는가?\n")

    print("Strike Wiki generated: 24 files.")

def generate_carcass_wiki(c_df):
    carcass_base = os.path.join(BASE_DIR, "사체수거")
    os.makedirs(carcass_base, exist_ok=True)
    
    for m in range(1, 13):
        m_df = c_df[c_df['Month'] == m]
        m_str = f"{m:02d}"
        folder_name = f"{m}월 사체수거"
        m_dir = os.path.join(carcass_base, folder_name)
        os.makedirs(m_dir, exist_ok=True)
        
        season_title, season_desc = SEASON_MAP[m]
        count = len(m_df)
        shattered = len(m_df[m_df['Is_Strike_Probable'] == 1])
        
        # 1. 2023-XX_사체수거_통합보고서.md
        report_path = os.path.join(m_dir, f"2023-{m_str}_사체수거_통합보고서.md")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"### [기동지역 사체수거 월간 통합보고서 - 2023년 {m_str}월]\n\n")
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
                    
        # 2. 2023-XX_전문가_브리핑_및_인사이트.md
        briefing_path = os.path.join(m_dir, f"2023-{m_str}_전문가_브리핑_및_인사이트.md")
        with open(briefing_path, "w", encoding="utf-8") as f:
            f.write(f"날짜: 2023-{m_str}-28\n")
            f.write(f"카테고리: 야생동물점검 / 사체수거 데이터 분석\n")
            f.write(f"출처: 2023년 {m}월 기동지역 사체수거 현황 ({count}건)\n")
            f.write(f"태그: #야통대 #사체수거 #FOD예방 #실물포렌식 #2023년_{m}월\n\n")
            f.write("### 전문가 종합 요약 및 의견 (컨설턴트 브리핑)\n\n")
            f.write(f"**[위기상황: {season_title} 사체 방치에 따른 2차 FOD 및 포식자 유인 위험]**\n")
            f.write(f"2023년 {m}월 중 기동지역에서 총 {count}건의 조류 사체가 수거되었으며, 이 중 완파 및 반파 상태는 {shattered}건으로 집계됨. ")
            f.write(f"포장면 인근에서 발견된 파손 사체는 항공기 충돌의 직접 증거이며, 잔류 파편이 후속 항공기 제트 엔진으로 흡입될 수 있는 중대 위험 요소임.\n\n")
            f.write(f"**[생존전략: 신속한 활주로 FOD 클리어 및 실물 DNA 감식 연동]**\n")
            f.write(f"사체 발견 즉시 고압 살수 및 진공 흡입을 통해 활주로 표면 잔해를 100% 제거하고, 훼손된 사체의 깃털 및 근육 조직을 채취하여 종 식별률을 제고함.\n\n")
            f.write(f"**[전망: 활주로 초지대 포식자 유인 먹이사슬 원천 차단]**\n")
            f.write(f"사체 잔해를 신속히 수거함으로써 까마귀, 왜가리, 맹금류 및 야생고양이가 활주로로 유입되는 2차 먹이사슬 고리를 영구 단절해야 함.\n\n")
            f.write("---\n\n")
            f.write("### 창의적 연결을 위한 심화 질문 (Insight Questions)\n\n")
            f.write(f"1. **완파 사체의 기체 충돌 역추적**: {m}월 수거된 완파 사체의 발견 위치와 당일 조류충돌 미보고 항공편 간의 시간적 역추적 기법은 어떻게 정밀화할 수 있는가?\n")
            f.write(f"2. **사체 발생 구역의 지형적 취약성**: 특정 유도로 및 활주로 인근에서 사체가 반복 발견되는 구조적 요인은 무엇인가?\n")
            f.write(f"3. **사체 부패도와 순찰 주기의 상관관계**: 발견 당시 사체 신선도(부패 진행 여부)를 통해 야간 순찰의 사각지대 존재 여부를 판정할 수 있는가?\n")
            f.write(f"4. **FOD 진공 흡입 프로세스의 자동화**: 사체 수거 즉시 운항1호 점검 차량과의 무전 자동 링크를 통한 표면 청소 최단 시간 단축 방안은 무엇인가?\n")
            f.write(f"5. **소형 조류 사체와 맹금류 사냥 활동의 연관성**: 종다리, 제비 사체의 다수 출몰이 황조롱이 등 천연기념물 맹금류의 에어사이드 난입을 촉진하는가?\n")
            f.write(f"6. **배수로 주변 사체 수거와 수질 관리**: 배수 암거 인근 오리류 사체 수거 시 수질 오염 방지 및 소독제 방역 가이드라인은 적정한가?\n")
            f.write(f"7. **DNA 시료 동결 보관 및 국립생물자원관 연계 체계**: 미상 사체의 DNA 식별 성공률을 90% 이상으로 끌어올리기 위한 현장 채취 키트 표준화 방안은 무엇인가?\n")

    print("Carcass Wiki generated: 24 files.")

def generate_patrol_wiki(p_df):
    patrol_base = os.path.join(BASE_DIR, "일일점검")
    os.makedirs(patrol_base, exist_ok=True)
    
    for m in range(1, 13):
        m_df = p_df[p_df['Month'] == m]
        m_str = f"{m:02d}"
        folder_name = f"{m}월 일일점검"
        m_dir = os.path.join(patrol_base, folder_name)
        os.makedirs(m_dir, exist_ok=True)
        
        season_title, season_desc = SEASON_MAP[m]
        count = len(m_df)
        act_cnt = len(m_df[m_df['Event_Type'] == 'Actual_Contact'])
        rec_cnt = len(m_df[m_df['Event_Type'] == 'Search_Recon'])
        rounds = m_df['Rounds_Total'].sum()
        r7 = m_df['Rounds_7_1_2'].sum()
        r4 = m_df['Rounds_4'].sum()
        r_blk = m_df['Rounds_Blank'].sum()
        geese = m_df[m_df['Species'].str.contains('기러기')]['Count'].sum()
        
        # 1. 2023-XX_일일점검_통합보고서.md
        report_path = os.path.join(m_dir, f"2023-{m_str}_일일점검_통합보고서.md")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"### [야생동물통제대 일일점검 월간 통합보고서 - 2023년 {m_str}월]\n\n")
            f.write(f"- **총 점검 건수**: {count:,}건\n")
            f.write(f"- **실질 조류 조우 (`Actual_Contact`)**: {act_cnt:,}건 ({act_cnt/count*100:.1f}%)\n")
            f.write(f"- **선제적 수색 사격 (`Search_Recon`)**: {rec_cnt:,}건 ({rec_cnt/count*100:.1f}%)\n")
            f.write(f"- **총 탄약 발포 발수**: {rounds:,}발\n")
            f.write(f"  - 7호탄 (근접 분산): {r7:,}발\n")
            f.write(f"  - 4호탄 (원거리 경퇴): {r4:,}발\n")
            f.write(f"  - 공포탄 (소음 경고): {r_blk:,}발\n")
            f.write(f"- **큰기러기 관측·퇴치 개체수**: {geese:,}마리\n\n")
            f.write("#### 19개 Zone별 위험 집중도 Top 3\n")
            top_z = m_df[m_df['Event_Type'] == 'Actual_Contact']['Zone_19'].value_counts().head(3)
            for z_name, z_val in top_z.items():
                f.write(f"- **{z_name}**: {z_val:,}건\n")
            f.write("\n")
            
        # 2. 2023-XX_전문가_브리핑_및_인사이트.md
        briefing_path = os.path.join(m_dir, f"2023-{m_str}_전문가_브리핑_및_인사이트.md")
        with open(briefing_path, "w", encoding="utf-8") as f:
            f.write(f"날짜: 2023-{m_str}-28\n")
            f.write(f"카테고리: 야생동물점검 / 순찰 및 퇴치 실적 분석\n")
            f.write(f"출처: 2023년 {m}월 일일점검 일지 전수 ({count:,}건)\n")
            f.write(f"태그: #야통대 #일일점검 #퇴치실적 #탄약통계 #2023년_{m}월\n\n")
            f.write("### 전문가 종합 요약 및 의견 (컨설턴트 브리핑)\n\n")
            f.write(f"**[위기상황: {season_title} 순찰 예찰 및 출몰 압력 대응]**\n")
            f.write(f"2023년 {m}월 중 총 {count:,}건의 순찰 활동이 전개되었으며, 총 {rounds:,}발의 탄약이 발포되어 공항 활주로 안전구역을 사수함. ")
            if geese > 0:
                f.write(f"특히 당월 큰기러기 출몰량이 {geese:,}마리에 달하여 일출 및 일몰 전후 취약 시간대의 방호선 유지가 핵심 과제였음.\n\n")
            else:
                f.write(f"텃새 및 계절성 소형 조류의 초지대 활동에 대응하여 7호탄 중심의 근접 분산 전술이 주효하였음.\n\n")
            f.write(f"**[생존전략: 선제 수색 사격(Search_Recon)을 통한 야간 안착 원천 봉쇄]**\n")
            f.write(f"단순 조우 시 퇴치뿐만 아니라, 조류가 보이지 않더라도 일몰 후 잠자리 안착을 막기 위해 총 {rec_cnt:,}건의 선제적 수색 사격을 단행하여 조류의 활주로 야간 체류를 방어함.\n\n")
            f.write(f"**[전망: 19개 표준 Zone 기반 기동 차량 최적 배치]**\n")
            f.write(f"최다 출몰 구역을 중심으로 순찰 노선을 재편하고, 주야간 교대 주기와 맞물린 집중 경계 거점을 고정 운용해야 함.\n\n")
            f.write("---\n\n")
            f.write("### 창의적 연결을 위한 심화 질문 (Insight Questions)\n\n")
            f.write(f"1. **{m}월 탄약 발포량과 조류충돌 억제율의 정량적 상관관계**: 탄약 1,000발 소모당 조류충돌 감소 효과를 통계적으로 산출할 수 있는가?\n")
            f.write(f"2. **Search_Recon(선제 수색 사격)의 억제 지속 시간**: 조류가 없는 상태에서 격발한 공포탄/조류탄의 소음 효과가 조류의 재접근을 몇 시간 동안 저지하는가?\n")
            f.write(f"3. **{m}월 4호탄(원거리) 소모 비중 변화의 생태적 의미**: 원거리 경퇴 사격이 증가한 구역의 조류 경계 거리 변화는 어떠한가?\n")
            f.write(f"4. **순찰 조별 기동 거리와 활주로 사각지대 해소율**: 일일 차량 순찰 킬로미터(km) 수와 미인지 조류 군집 발견율 간의 관계는 무엇인가?\n")
            f.write(f"5. **{season_title} 조석(물때) 주기와 순찰 집중 시간대의 동기화**: 만조 시 갯벌에서 밀려나는 조류를 활주로 진입 전 해안선에서 차단하는 최적 타이밍은 언제인가?\n")
            f.write(f"6. **대원 피로도와 야간 사격 명중/분산율 간의 상관성**: 24시간 교대 근무 체계에서 새벽 시간대 집중도 유지를 위한 기동 지원 방안은 무엇인가?\n")
            f.write(f"7. **2023년 {m}월 순찰 패턴과 2024~2025년 동일 월 순찰 패턴의 비교**: 기동 전술의 고도화가 이듬해 조류 서식 습성 붕괴에 미친 장기적 영향은 무엇인가?\n")

    print("Patrol Wiki generated: 24 files.")

def main():
    print("=== Generating 2023 Full Monthly Wiki Suite ===")
    p_df = pd.read_csv(CSV_PATROL)
    s_df = pd.read_csv(CSV_STRIKE)
    c_df = pd.read_csv(CSV_CARCASS)
    
    generate_strike_wiki(s_df)
    generate_carcass_wiki(c_df)
    generate_patrol_wiki(p_df)
    print("=== ALL 2023 WIKI SUITES GENERATED SUCCESSFULLY (72 Files Total) ===")

if __name__ == '__main__':
    main()
