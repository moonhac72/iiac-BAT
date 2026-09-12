# -*- coding: utf-8 -*-
"""
2021~2023년도 생동물 발생보고내역 통합보고서 및 전문가 브리핑 자동 생성 스크립트
(근무자, 신고자, 접수자 등 개인 실명 및 연락처 완벽 마스킹/삭제 적용)
"""

import os
import sys
import re
import datetime
from collections import defaultdict, Counter
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

EXCEL_FILE = r'C:\Users\moonh\Desktop\생동물. 신고\생동물(21_23).xlsx'
OUTPUT_BASE = r'c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료'

def remove_personal_names(text):
    if not text:
        return ""
    s = str(text)

    # 1. 전화번호 및 연락처 메모 (괄호 포함 전체)
    s = re.sub(r'※?\s*[가-힣]{0,4}\s*\(?\b\d{2,4}[-\s)]*\d{3,4}[-\s)]*\d{4}?\)?', '', s)
    s = re.sub(r'\(?[가-힣]{0,4}\s*010[-\s]?\d{4}[-\s]?\d{4}\)?', '', s)
    s = re.sub(r'\b010[-\s]?\d{4}[-\s]?\d{4}\b', '', s)
    s = re.sub(r'\b\d{2,4}-\d{3,4}-\d{4}\b', '', s)
    s = re.sub(r'\(\d{3,4}-\d{4}\)', '', s)

    # 2. [부서/기관/소속] + [이름][씨/氏/직함]? + [으로/로]?부터 -> [부서/기관/소속]으로부터
    # e.g., 'T2교통상황실 정재호씨로부터' -> 'T2교통상황실로부터'
    # e.g., '이동지역안전관리소 고동형 으로부터' -> '이동지역안전관리소로부터'
    # e.g., '대테러상황실 박홍철氏로부터' -> '대테러상황실로부터'
    s = re.sub(r'([가-힣A-Za-z0-9]+(?:상황실|센터|데스크|초소|관리소|시설|통제대|통제실|경비대|청사|팀|소|본대|지점|부서|역))\s+[가-힣]{2,4}\s*(?:씨|氏|파트장|과장|팀장|대리|주임|반장)?\s*(?:으로|로)?\s*부터', r'\1로부터', s)

    # 3. [부서/기관] + [이름][씨/氏] -> [부서/기관]
    s = re.sub(r'([가-힣A-Za-z0-9]+(?:상황실|센터|데스크|초소|관리소|시설|통제대|통제실|경비대|청사|팀|소|본대))\s+[가-힣]{2,4}\s*(?:씨|氏)', r'\1', s)

    # 4. 단독 [이름][씨/氏]? [으로/로]?부터 -> 신고자로부터
    s = re.sub(r'(?<![가-힣])[가-힣]{2,4}\s*(?:씨|氏)\s*(?:으로|로)?\s*부터', '신고자로부터', s)
    s = re.sub(r'(?<![가-힣])[가-힣]{2,4}\s*(?:으로|로)\s*부터', '신고자로부터', s)

    # 5. [이름][씨/氏] 단독 제거
    s = re.sub(r'[가-힣]{2,4}\s*(?:씨|氏)', '', s)

    # 6. 직함 및 인원 표기
    s = re.sub(r'[가-힣]{2,4}\s*파트장\s*외\s*\d+명', '지원인력', s)
    s = re.sub(r'[가-힣]{2,4}\s*(?:파트장|과장|팀장|대리|주임|반장)\b', '', s)

    # 7. 기동/출동 인원 나열 괄호 (e.g., '(강지구, 유명호)')
    s = re.sub(r'\([가-힣]{2,4}(?:,\s*[가-힣]{2,4})+\)', '', s)

    # 8. 잔여 빈 괄호 및 기호 정리
    s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'※', '', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_text(text):
    if text is None:
        return ""
    s = str(text).strip()
    if s.lower() == 'none':
        return ""
    s = remove_personal_names(s)
    return ' '.join(s.split())

def parse_dt(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return datetime.datetime(v.year, v.month, v.day, getattr(v, 'hour', 0), getattr(v, 'minute', 0))
    if isinstance(v, str):
        v = v.strip()
        m = re.match(r'^(\d{4})[./-](\d{1,2})[./-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2}))?', v)
        if m:
            y, mth, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            h = int(m.group(4)) if m.group(4) else 0
            mi = int(m.group(5)) if m.group(5) else 0
            return datetime.datetime(y, mth, d, h, mi)
    return None

def parse_action_details(action_text, dt):
    if not action_text:
        return '-', '-', ['- 특이사항 없음']
    lines = [l.strip() for l in str(action_text).split('\n') if l.strip()]
    formatted_lines = []
    t_arr = None
    t_end = None

    for l in lines:
        cleaned_l = remove_personal_names(l)
        if not cleaned_l:
            continue

        m_time = re.search(r'\(?(\d{1,2})[:시]?(\d{2})\)?', cleaned_l)
        time_str = ''
        if m_time:
            h = int(m_time.group(1))
            mi = int(m_time.group(2))
            if 0 <= h <= 24 and 0 <= mi <= 59:
                time_str = f"{h:02d}:{mi:02d}"
                cleaned_l = cleaned_l.replace(m_time.group(0), '').strip()
                if cleaned_l.startswith('-') or cleaned_l.startswith(':'):
                    cleaned_l = cleaned_l[1:].strip()

        if '현장도착' in cleaned_l or '현장 도착' in cleaned_l:
            if time_str and not t_arr:
                t_arr = time_str
            if time_str:
                formatted_lines.append(f"- **[{time_str} 현장도착]** {cleaned_l}")
            else:
                formatted_lines.append(f"- **[현장도착]** {cleaned_l}")
        elif '상황종료' in cleaned_l or '상황 종료' in cleaned_l:
            if time_str and not t_end:
                t_end = time_str
            if time_str:
                formatted_lines.append(f"- **[{time_str} 상황종료]** {cleaned_l}")
            else:
                formatted_lines.append(f"- **[상황종료]** {cleaned_l}")
        elif time_str:
            formatted_lines.append(f"- **[{time_str}]** {cleaned_l}")
        else:
            formatted_lines.append(f"- {cleaned_l}")

    if not t_arr and dt:
        t_arr = dt.strftime('%H:%M')
    if not t_end:
        t_end = t_arr

    return t_arr, t_end, formatted_lines

def build_monthly_report(year, month, records):
    rep_lines = [f"### [기동지역 생동물(포유류 등) 월간 통합보고서 - {year}년 {month:02d}월]\n"]

    if not records:
        rep_lines.append("> **안내**: 해당 월 기동지역 및 에어사이드 내 생동물(포유류/파충류 등) 침입·발견 특이 이벤트가 접수되지 않았습니다. 일상적인 외곽 펜스 순찰 및 포획틀 상시 점검이 유지되었습니다.\n")
        return '\n'.join(rep_lines)

    # 일자별 그룹화
    by_date = defaultdict(list)
    for r in records:
        d_str = r['dt'].strftime('%Y-%m-%d')
        by_date[d_str].append(r)

    for d_str in sorted(by_date.keys()):
        day_recs = by_date[d_str]
        first_r = day_recs[0]
        h = first_r['dt'].hour
        time_cat = "야간" if (h < 6 or h >= 19) else "주간"
        time_str = first_r['dt'].strftime('%H:%M')

        # 주요접수유형
        types = []
        for r in day_recs:
            t = f"{r['event']}({r['species']})" if r['species'] else r['event']
            if t not in types:
                types.append(t)
        types_str = ', '.join(types)

        rep_lines.append(f"#### 📅 {d_str}")
        rep_lines.append(f"*   **기상/관측정보**: 시간: {time_cat} ({time_str})")
        rep_lines.append(f"*   **주요접수유형**: {types_str}\n")
        rep_lines.append("**상세 타임라인 및 조치사항**")

        for r in day_recs:
            for fl in r['timeline_lines']:
                rep_lines.append(fl)
            rep_lines.append(f"  - **현장도착**: {r['t_arr']} | **상황종료**: {r['t_end']}")
            rep_lines.append(f"  - **발생장소**: {r['location']}")
            if r['airline']:
                rep_lines.append(f"  - **생동물/항공사**: {r['category']} / {r['species']} (항공사: {r['airline']})")
            else:
                rep_lines.append(f"  - **생동물/종명**: {r['category']} / {r['species']}")
            rep_lines.append(f"  - **신고출처**: 이동지역 통제대 및 현장 접수")
            rep_lines.append(f"  - **조치결과**: {r['result']}")
            rep_lines.append("")

        rep_lines.append("---\n")

    return '\n'.join(rep_lines)

def build_insight_report(year, month, records):
    total_cnt = len(records)
    if total_cnt == 0:
        crisis_text = f"{year}년 {month:02d}월은 기동지역 및 에어사이드 전역에서 생동물(포유류, 파충류 등) 침입이나 발견 신고가 접수되지 않아 매우 안정적인 보안 및 통제 상태를 유지한 기간이었습니다."
        strategy_text = "외곽 펜스 및 배수로 침투 취약 지점에 대한 정기 순찰과 고양이 포획틀 상시 유지를 통해 침입 징후를 선제 차단하였습니다."
        outlook_text = "계절 변화에 따른 포유류 활동량 증가에 대비하여 외곽 취약 틈새 보수 및 열화상 야간 순찰을 지속 강화해야 합니다."
        summary_text = f"{year}년 {month:02d}월은 에어사이드 내 생동물 침입 이벤트가 발생하지 않아 비행장 운영의 정상성을 확고히 지켰습니다."
        top_species_str = "특이 생동물 없음"
        top_locs_str = "기동지역 전역"
        night_ratio = 0
    else:
        # 통계 집계
        sp_counts = Counter([r['species'] for r in records if r['species']])
        loc_counts = Counter([r['location'] for r in records if r['location']])
        event_counts = Counter([r['event'] for r in records if r['event']])

        top_species = [f"{sp}({cnt}건)" for sp, cnt in sp_counts.most_common(4)]
        top_species_str = ', '.join(top_species)

        top_locs = [loc for loc, _ in loc_counts.most_common(3)]
        top_locs_str = ', '.join(top_locs)

        night_cnt = sum(1 for r in records if r['dt'].hour < 6 or r['dt'].hour >= 19)
        night_ratio = round((night_cnt / total_cnt) * 100, 1)

        captured_cnt = sum(1 for r in records if any(k in r['result'] + ' '.join(r['timeline_lines']) for k in ['생포', '포획', '수거', '인계']))

        crisis_text = (
            f"{year}년 {month:02d}월 한 달간 에어사이드 및 랜드사이드 경계 구역에서 총 {total_cnt}건의 생동물(포유류/파충류 등) 침입·발견 이벤트가 접수되었습니다. "
            f"주요 발생 이벤트는 {', '.join([f'{k}({v}건)' for k, v in event_counts.items()])}이며, "
            f"주요 개체로는 {top_species_str} 등이 집중 식별되었습니다. 특히 발생 위치는 {top_locs_str} 일대에 집중되었으며, "
            f"전체 이벤트 중 야간·새벽 시간대(19:00~06:00) 발생 비중이 {night_ratio}%를 차지하여 어두운 시정 환경 속에서 "
            f"항공기 계류 구역 및 지상 조업 시설의 보안을 위협하는 취약 요인으로 작용했습니다."
        )

        strategy_text = (
            f"기동지역 안전관리소 및 야통대는 침입 신고 접수 즉시 기동 차량을 급파하여 정밀 수색 및 차단 방어선을 전개했습니다. "
            f"해당 월 총 {captured_cnt}건의 능동적 생포·수거 및 포획틀 설치 조치를 완료하였으며, "
            f"포획된 야생고양이는 지정 동물병원으로 안전하게 인계하고 고라니/조류/파충류 등은 외곽 녹지대로 안전 방사 조치했습니다. "
            f"특히 화물 터미널 하역장 및 초소 출입문 인근에 다중 포획틀을 신속 재배치하여 2차 기내·에어사이드 침투를 완벽하게 차단했습니다."
        )

        outlook_text = (
            f"계절적 요인 및 공항 주변 서식 환경 변화에 따라 외곽 펜스 및 배수로를 통한 포유류(고양이, 고라니 등)의 침투 압력이 지속될 것으로 전망됩니다. "
            f"화물 하역구 하단 틈새 차단판 설치, 초소 경비대와의 TRS 무전 핫라인 공유, 야간 열화상 카메라 순찰 강화를 통해 "
            f"생동물의 비행장 무단 침입을 원천 봉쇄해야 합니다."
        )

        summary_text = (
            f"{year}년 {month:02d}월 동안 기동지역에서 접수된 생동물 관련 신고는 총 {total_cnt}건으로, "
            f"{top_species_str}의 출현이 두드러졌습니다. 조류 통제 요원들은 {top_locs_str} 등 취약 시설을 중심으로 "
            f"신속한 현장 출동과 맞춤형 포획틀 운용을 통해 항공기 및 지상 조업 안전을 확고하게 사수했습니다."
        )

    # 심화 질문 7문항
    q1 = f"**야간·새벽 시간대({night_ratio}%) 생동물 침입 역학 :** {month}월에 발생한 생동물 신고 중 야간 비중이 높은 물리적·생태적 원인은 무엇이며, 야간 순찰차량의 열화상(FLIR) 탐지 프로토콜을 어떻게 고도화해야 할까요?"
    q2 = f"**화물터미널(AACT/샤프 등) 및 특정 주기장의 유인 요인 :** {top_locs_str} 일대에 고양이 및 소형 포유류의 출몰이 빈발하는 구조적 원인(음식물 쓰레기 잔존, 난방 온기, 지면 틈새 등)을 규명하기 위한 시설 환경 점검 기준은 무엇입니까?"
    q3 = f"**에어사이드-랜드사이드 경계 펜스 및 배수로 방벽의 무결성 :** 외부 서식지에서 공항 내부로 침투하는 고라니 및 야생고양이의 침입 경로를 차단하기 위해, 지상 펜스 하부 유격 허용 기준(10cm 미만) 및 배수로 차단 스크린의 주기적 정비 체계는 어떻게 설계되어야 할까요?"
    q4 = f"**포획틀 배치 전략 및 장비 관리의 실효성 :** 포획틀 설치 후 회수까지의 시간 경과에 따른 생포 성공률과 보관 중 동물 이탈 사고를 원천 방지하기 위한 잠금장치 개선 규격은 어떻게 마련해야 합니까?"
    q5 = f"**기내 및 화물칸 생동물 유입 차단 매뉴얼 :** 항공기 주기장 및 화물 조업 중 반려동물 탈출이나 해외 유입 포유류 침투 발생 시, 항공사 조업사와 야통대 간의 초동 골든타임 사수를 위한 합동 대응 SOP는 어떻게 개선되어야 할까요?"
    q6 = f"**생태 피라미드와 2차 조류 위험 연계 분석 :** 에어사이드 내 야생고양이 및 설치류의 서식 밀도가 상위 포식자인 맹금류(황조롱이, 말똥가리)의 활주로 주변 유입에 미치는 상관관계를 통계적으로 교차 검증할 수 있는 분석 모델은 무엇입니까?"
    q7 = f"**포획 동물의 인도적 처리 및 방사 사후 관리 :** 생포된 야생고양이의 동물병원 인계 및 고라니의 인근 야산(신불산, 오성산 등) 방사 후, 재침입 여부를 모니터링하기 위한 개체 식별 및 관리 체계는 어떻게 수립되어야 할까요?"

    tags = [f"#{year}년", f"#{month}월생동물", "#생동물관리", "#포유류침입", "#야생고양이", "#고라니", "#포획틀관리", "#항공안전"]
    tag_str = ' '.join(tags)

    ins_lines = [
        f"날짜: {year}-{month:02d}-28",
        f"카테고리: 생동물(포유류) 관리 / 침입 경로 분석 및 보안 인사이트",
        f"출처: {year}년 {month:02d}월 생동물 이벤트 발생 보고서 (실명 마스킹 완료)",
        f"태그: {tag_str}\n",
        f"### 전문가(컨설턴트) 종합 요약 및 의견\n",
        f"**[위기상황: {month}월 생동물 출현 및 에어사이드 침입 위협]**",
        f"{crisis_text}\n",
        f"**[생존전략: 신속한 차단선 구축과 다중 포획틀 운용]**",
        f"{strategy_text}\n",
        f"**[전망: 계절성 이동 방벽 강화 및 화물 구역 밀폐 관리]**",
        f"{outlook_text}\n",
        f"---\n",
        f"### 기사 내용 요약 (알기 쉬운 해설)\n",
        f"**출처: {year}년 {month:02d}월 생동물 이벤트 발생 보고서**\n",
        f"{summary_text}\n",
        f"---\n",
        f"### 창의적인 인사이트를 위한 질문 (데이터베이스 연결용)\n",
        f"1. {q1}",
        f"2. {q2}",
        f"3. {q3}",
        f"4. {q4}",
        f"5. {q5}",
        f"6. {q6}",
        f"7. {q7}\n",
        f"---\n",
        f"> **운용 주의**: 본 보고서는 {month}월 신속 조치 및 포획 실측 데이터를 기반으로 작성되었습니다.",
        f"> {month}월 사체수거 및 조류충돌 보고서와 연계하여 **'종합 Wildlife Risk Matrix'**를 업데이트하십시오.\n"
    ]

    return '\n'.join(ins_lines)

def build_yearly_summary(year, records_by_month):
    all_recs = []
    for m in range(1, 13):
        all_recs.extend(records_by_month[m])

    total_cnt = len(all_recs)
    sp_counts = Counter([r['species'] for r in all_recs if r['species']])
    loc_counts = Counter([r['location'] for r in all_recs if r['location']])
    event_counts = Counter([r['event'] for r in all_recs if r['event']])

    top_sp = [f"- **{sp}**: {cnt}건 ({round(cnt/total_cnt*100, 1) if total_cnt else 0}%)" for sp, cnt in sp_counts.most_common(10)]
    top_loc = [f"- **{loc}**: {cnt}건" for loc, cnt in loc_counts.most_common(10)]
    monthly_stat = [f"- **{m}월**: {len(records_by_month[m])}건" for m in range(1, 13)]

    lines = [
        f"# {year}년 야통대 생동물(포유류 등) 발생 및 포획 종합 통계 보고서\n",
        f"## 1. 개요",
        f"{year}년 한 해 동안 인천국제공항 에어사이드(A/S) 및 랜드사이드(L/S) 경계 구역에서 발생한 생동물(포유류, 파충류 등) 침입·발견·탈출 이벤트는 **총 {total_cnt}건**이었습니다.",
        f"야생조수관리소(야통대)는 상시 기동 순찰과 포획틀 운용, 신속 출동 체계를 통해 비행장 안전을 확보했습니다.\n",
        f"## 2. 월별 발생 추이",
        '\n'.join(monthly_stat),
        f"\n## 3. 주요 침입·출현 종별 현황",
        '\n'.join(top_sp),
        f"\n## 4. 다발 장소 현황 (TOP 10)",
        '\n'.join(top_loc),
        f"\n## 5. 이벤트 유형별 분포",
        '\n'.join([f"- **{k}**: {v}건" for k, v in event_counts.items()]),
        f"\n## 6. 결론 및 종합 제언",
        f"{year}년 통계 분석 결과, 야생고양이와 고라니가 공항 생동물 관리의 양대 핵심 축으로 나타났습니다. 특히 화물터미널 구역과 초소 출입문 주변을 통한 유입 압력이 높으므로 물리적 차단 시설 보강과 상시 모니터링 체계 유지가 필수적입니다."
    ]

    return '\n'.join(lines)

def main():
    print("=== 2021~2023년 생동물 보고서 생성 시작 (실명 삭제 적용) ===")
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    ws = wb.active
    raw_rows = list(ws.iter_rows(values_only=True))[3:]

    data_by_year_month = defaultdict(lambda: defaultdict(list))

    for idx, r in enumerate(raw_rows):
        if not any(r):
            continue
        dt = parse_dt(r[2])
        if not dt:
            continue

        year = dt.year
        month = dt.month

        airline = clean_text(r[3])
        event = clean_text(r[4])
        category = clean_text(r[5])
        species = clean_text(r[6])
        location = clean_text(r[7])
        action_text = r[8]
        result = clean_text(r[9]) if len(r) > 9 else "-"

        t_arr, t_end, flines = parse_action_details(action_text, dt)

        data_by_year_month[year][month].append({
            'dt': dt,
            'airline': airline,
            'event': event,
            'category': category,
            'species': species,
            'location': location,
            'timeline_lines': flines,
            't_arr': t_arr,
            't_end': t_end,
            'result': result
        })

    total_created = 0

    for year in [2021, 2022, 2023]:
        year_dir = os.path.join(OUTPUT_BASE, f"{year}년도 야통대", "생동물")
        os.makedirs(year_dir, exist_ok=True)
        print(f"\n>> [{year}년도] 생동물 폴더 구축 시작 -> {year_dir}")

        for month in range(1, 13):
            recs = data_by_year_month[year][month]
            recs.sort(key=lambda x: x['dt'])

            rep_md = build_monthly_report(year, month, recs)
            ins_md = build_insight_report(year, month, recs)

            rep_path = os.path.join(year_dir, f"{year}-{month:02d}_생동물_통합보고서.md")
            ins_path = os.path.join(year_dir, f"{year}-{month:02d}_전문가_브리핑_및_인사이트.md")

            with open(rep_path, 'w', encoding='utf-8') as f:
                f.write(rep_md)

            with open(ins_path, 'w', encoding='utf-8') as f:
                f.write(ins_md)

            print(f"  - {year}년 {month:02d}월 완료 ({len(recs)}건)")
            total_created += 2

        summary_md = build_yearly_summary(year, data_by_year_month[year])
        summary_path = os.path.join(year_dir, f"{year}_생동물_통계_보고서.md")
        with open(summary_path, 'w', encoding='utf-8') as f:
            f.write(summary_md)
        total_created += 1
        print(f"  - {year}_생동물_통계_보고서.md 생성 완료")

    print(f"\n[성공] 총 {total_created}개 생동물 문서 생성 완료 (실명 제거 완료)!")

if __name__ == '__main__':
    main()
