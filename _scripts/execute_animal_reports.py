import os, glob, re, sys
sys.path.insert(0, os.path.abspath('.'))
from collections import Counter
from _scripts.reformat_animal_reports import parse_and_reformat_month

sys.stdout.reconfigure(encoding='utf-8')

def get_normalized_location_2024(loc):
    # AACT 및 648·649번 주기장 일대
    if any(k in loc for k in ['648', '649', 'AACT']):
        return "AACT 화물터미널 및 648·649번 주기장 일대"
    if '612' in loc:
        return "612번 주기장 및 조명탑 일대"
    if '601' in loc:
        return "A/S 화물터미널 601번 주기장"
    if '614' in loc:
        return "A/S 614번 주기장 및 ULD 적재장 인근"
    if '655' in loc:
        return "A/S 화물터미널 655번 주기장 인근"
    if '673' in loc:
        return "A/S 화물터미널 673번 주기장"
    if '841' in loc:
        return "841번 주기장 및 화물터미널 GSE 변압기 인근"
    if '851' in loc:
        return "A/S 851번 주기장"
    if '토목' in loc:
        return "A/S 토목정비고 및 액체탱크 담장 일대"
    if '32번' in loc:
        return "32번 주기장 R3 지역"
    if '6번주기장' in loc:
        return "A/S 6번 주기장 계류장"
    if '23번' in loc:
        return "23번 주기장 계류장"
    if '28번' in loc:
        return "28번 주기장 수화물 작업장"
    if '252' in loc:
        return "A/S 252번 주기장"
    if '망루' in loc or 'G10' in loc:
        return "외곽 감시망루 및 초지 일대 (망루 5·9·17번)"
    if '소방대' in loc:
        return "소방대 분소B 일대"
    if '초소' in loc or 'G2' in loc or 'FG1' in loc:
        return "외곽 보안초소 일대 (G2, G11, FG1초소)"
    if '교통센터' in loc:
        return "L/S 교통센터 구역"
    if '배수로' in loc:
        return "D2 간이배수로 및 33말단 EG7 배수로"
    return loc

def get_normalized_location_2025(loc):
    if '[주기장_K]' in loc:
        return "A/S [주기장_K] 구역"
    if '[주기장_Q]' in loc or 'BHS' in loc:
        return "A/S [주기장_Q] 지하 2층 수하물처리시설(BHS)"
    if '[주기장_E]' in loc or '수하물벨트' in loc:
        return "A/S [주기장_E] 수하물벨트 설치 구역"
    if '[유도로_AS]' in loc:
        return "A/S [유도로_AS] 초소 및 한국공항 GSE 정비시설"
    if '소방대' in loc:
        return "A/S 소방대 본대 뒷편 및 녹지대 배수로"
    if 'TG3' in loc:
        return "A/S TG3 우측 담장 포획틀 구역"
    if 'FG1' in loc:
        return "A/S FG1 초소 인근 우측벽 라인"
    if '기내' in loc or '아시아나' in loc:
        return "아시아나항공 기내 화장실 (계류장 주기 기준)"
    if '604' in loc or '605' in loc:
        return "A/S 화물터미널 604~605번 주기장 사이"
    if '[유도로_AP]' in loc and ('VIP' in loc or '6번출구' in loc or '1층' in loc):
        return "L/S [유도로_AP] 터미널 1층 및 VIP 주차장 화단"
    if '출국장' in loc or '게이트' in loc or '23번' in loc:
        return "여객터미널 3층 출국장 및 탑승구 구역"
    if '11번 탑승교' in loc:
        return "A/S 11번 탑승교 조업장 ULD박스 내부"
    if '항공등화' in loc:
        return "A/S 항공등화 작업장 건물 뒤편"
    if '[주기장_B]' in loc:
        return "A/S [주기장_B] 인근 시설"
    if '망루' in loc:
        return "A/S 망루 1번 및 6번 인근 GSE 도로"
    return loc

def build_annual_report(year, events):
    total_count = len(events)
    
    # 1. 월별 발생 추이
    month_counts = Counter()
    for e in events:
        month_counts[e['month']] += 1
    
    monthly_lines = []
    for m in range(1, 13):
        monthly_lines.append(f"- **{m}월**: {month_counts[m]}건")
    monthly_section = "\n".join(monthly_lines)
    
    # 2. 종별 현황 (세부 종명 추출)
    species_counter = Counter()
    for e in events:
        sp_meta = e['species_meta']
        if '/' in sp_meta:
            detail = sp_meta.split('/')[1].strip()
            detail = re.sub(r'\(항공사:.*?\)', '', detail).strip()
        else:
            detail = sp_meta.strip()
        species_counter[detail] += 1
        
    species_lines = []
    for sp_name, cnt in species_counter.most_common():
        pct = (cnt / total_count) * 100
        species_lines.append(f"- **{sp_name}**: {cnt}건 ({pct:.1f}%)")
    species_section = "\n".join(species_lines)
    
    # 3. 다발 장소 TOP 10 (정규화된 대표 거점별)
    loc_counter = Counter()
    for e in events:
        raw_loc = e['location'].strip()
        if year == 2024:
            norm_loc = get_normalized_location_2024(raw_loc)
        else:
            norm_loc = get_normalized_location_2025(raw_loc)
        loc_counter[norm_loc] += 1
        
    top_locs = loc_counter.most_common(10)
    top_loc_lines = []
    for loc_name, cnt in top_locs:
        top_loc_lines.append(f"- **{loc_name}**: {cnt}건")
    top_loc_section = "\n".join(top_loc_lines)
    
    # 4. 유형별 분포
    type_counter = Counter()
    for e in events:
        t = e['type'].strip()
        if '탈출' in t:
            if '기내' in t: main_type = '조업중 동물탈출(기내)'
            else: main_type = '조업중 동물탈출'
        elif '공항내 유입' in t or '유입' in t:
            main_type = '생동물 공항내 유입'
        elif '포획틀' in t:
            main_type = '생동물발견(포획틀 생포)'
        else:
            main_type = '생동물발견'
        type_counter[main_type] += 1
        
    type_lines = []
    for t_name, cnt in type_counter.most_common():
        type_lines.append(f"- **{t_name}**: {cnt}건")
    type_section = "\n".join(type_lines)
    
    # 5. 결론 및 종합 제언
    if year == 2024:
        conclusion = (
            "2024년 통계 분석 결과, AACT 화물터미널 및 계류장 주변(648·649번 주기장 일대)을 중심으로 야생고양이 출현 및 유입이 집중되었습니다. "
            "또한 고라니 유입 및 항공기 조업 중 반려견·반려묘 탈출 사례가 지속 발생함에 따라, 외곽 차단 울타리 및 배수로 점검 강화와 함께 조업 구역 내 동물 케이지 취급 관리 지침 준수가 필수적입니다."
        )
    else:
        conclusion = (
            "2025년 통계 분석 결과, 야생고양이 포획틀 상시 운용에 따른 선제적 생포 실적이 크게 증가하였으며, 여름철 파충류(뱀) 출현 및 조업 중 동물 탈출 사례가 주요 관리 요소로 확인되었습니다. "
            "수하물처리구역(BHS) 및 항공등화·정비시설 등 건물 틈새 유입을 원천 차단하기 위한 시설물 보강과 기동 포획 체계의 상시 유지가 요구됩니다."
        )

    report_md = f"""# {year}년 야통대 생동물(포유류 등) 발생 및 포획 종합 통계 보고서

## 1. 개요
{year}년 한 해 동안 인천국제공항 에어사이드(A/S) 및 랜드사이드(L/S) 경계 구역에서 발생한 생동물(포유류, 파충류 등) 침입·발견·탈출 이벤트는 **총 {total_count}건**이었습니다.
야생조수관리소(야통대)는 상시 기동 순찰과 포획틀 운용, 신속 출동 체계를 통해 비행장 안전을 확보했습니다.

## 2. 월별 발생 추이
{monthly_section}

## 3. 주요 침입·출현 종별 현황
{species_section}

## 4. 다발 장소 현황 (TOP 10)
{top_loc_section}

## 5. 이벤트 유형별 분포
{type_section}

## 6. 결론 및 종합 제언
{conclusion}
"""
    return report_md

def run():
    for year in [2024, 2025]:
        folder = f'10_Wiki/🛠️ Projects/야통대_자료/{year}년도 야통대/생동물'
        files = sorted(glob.glob(f'{folder}/{year}-*_생동물_통합보고서.md'))
        
        all_year_events = []
        for fpath in files:
            reformatted_md, events = parse_and_reformat_month(fpath, year)
            all_year_events.extend(events)
            with open(fpath, 'w', encoding='utf-8') as out_fp:
                out_fp.write(reformatted_md)
            print(f"Updated: {os.path.basename(fpath)} ({len(events)} events)")
            
        # 연간 통계 보고서 생성
        annual_md = build_annual_report(year, all_year_events)
        annual_path = f"{folder}/{year}_생동물_통계_보고서.md"
        with open(annual_path, 'w', encoding='utf-8') as out_fp:
            out_fp.write(annual_md)
        print(f"Generated Annual Report: {annual_path} (Total {len(all_year_events)} events)")

if __name__ == '__main__':
    run()
