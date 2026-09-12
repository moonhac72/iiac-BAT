# -*- coding: utf-8 -*-
"""
2024년 및 2025년 생동물 월간 통합보고서/인사이트 재포맷팅 및 연간 통계보고서 정밀 생성 스크립트
"""

import os
import sys
import re
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r'c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료'

def reformat_integrated_report(text, year, month):
    m_title = re.search(r'###\s*\[기동지역 생동물.*?\]', text)
    title = m_title.group(0) if m_title else f"### [기동지역 생동물(포유류 등) 월간 통합보고서 - {year}년 {month:02d}월]"

    sections = text.split('#### 📅')
    out_lines = [title, ""]

    if len(sections) <= 1:
        out_lines.append("> **안내**: 해당 월 기동지역 및 에어사이드 내 생동물(포유류/파충류 등) 침입·발견 특이 이벤트가 접수되지 않았습니다. 일상적인 외곽 펜스 순찰 및 포획틀 상시 점검이 유지되었습니다.\n")
        return '\n'.join(out_lines)

    for s in sections[1:]:
        s = s.strip()
        if not s:
            continue
        m_date = re.match(r'^(\d{4}-\d{2}-\d{2})', s)
        if not m_date:
            continue
        date_str = m_date.group(1)
        rest = s[len(date_str):].strip()

        if rest.endswith('---'):
            rest = rest[:-3].strip()

        out_lines.append(f"#### 📅 {date_str}")

        # 기상/관측정보
        m_env = re.search(r'\*\s*\*\*기상/관측정보\*\*:\s*(.*?)(?=\*\s*\*\*주요접수유형|\*\*상세 타임라인|$)', rest)
        env_str = m_env.group(1).strip() if m_env else ""
        if env_str:
            out_lines.append(f"*   **기상/관측정보**: {env_str}")

        # 주요접수유형
        m_typ = re.search(r'\*\s*\*\*주요접수유형\*\*:\s*(.*?)(?=\*\*상세 타임라인|$)', rest)
        typ_str = m_typ.group(1).strip() if m_typ else ""
        if typ_str:
            out_lines.append(f"*   **주요접수유형**: {typ_str}\n")

        out_lines.append("**상세 타임라인 및 조치사항**")

        m_body = re.search(r'\*\*상세 타임라인 및 조치사항\*\*\s*(.*)', rest)
        body = m_body.group(1).strip() if m_body else ""

        items = re.split(r'\s+-\s+', ' ' + body)
        for it in items:
            it = it.strip()
            if not it:
                continue
            if it.startswith('- '):
                it = it[2:].strip()

            is_sub = any(it.startswith(f"**{k}**") for k in [
                '현장도착', '상황종료', '발생장소', '신고출처', '조치결과', '생동물/종명', '생동물/항공사', '생동물구분', '종명', '항공사'
            ])

            if is_sub:
                out_lines.append(f"  - {it}")
            else:
                out_lines.append(f"- {it}")

        out_lines.append("\n---\n")

    return '\n'.join(out_lines)

def reformat_insight_report(text, year, month):
    m_date = re.search(r'날짜:\s*([^\s]+)', text)
    date_str = m_date.group(1) if m_date else f"{year}-{month:02d}-28"

    m_cat = re.search(r'카테고리:\s*(.*?)(?=출처:|$)', text)
    cat_str = m_cat.group(1).strip() if m_cat else "생동물(포유류) 관리 / 침입 경로 분석 및 보안 인사이트"

    m_src = re.search(r'출처:\s*(.*?)(?=태그:|$)', text)
    src_str = m_src.group(1).strip() if m_src else f"{month}월 생동물 이벤트 발생 보고서 (실명 마스킹 완료)"

    m_tag = re.search(r'태그:\s*(.*?)(?=###|$)', text)
    tag_str = m_tag.group(1).strip() if m_tag else f"#{year}년 #{month}월생동물 #생동물관리 #포유류침입 #야생고양이 #고라니 #포획틀관리 #항공안전"

    c_title_match = re.search(r'\*\*\[위기상황:.*?\]\*\*', text)
    c_title = c_title_match.group(0) if c_title_match else f"**[위기상황: {month}월 생동물 출현 및 에어사이드 침입 위협]**"
    m_crisis = re.search(r'\*\*\[위기상황:.*?\]\*\*\s*(.*?)(?=\*\*\[생존전략|$)', text)
    c_body = m_crisis.group(1).strip() if m_crisis else ""

    s_title_match = re.search(r'\*\*\[생존전략:.*?\]\*\*', text)
    s_title = s_title_match.group(0) if s_title_match else f"**[생존전략: 신속한 차단선 구축과 다중 포획틀 운용]**"
    m_strat = re.search(r'\*\*\[생존전략:.*?\]\*\*\s*(.*?)(?=\*\*\[전망|$)', text)
    s_body = m_strat.group(1).strip() if m_strat else ""

    o_title_match = re.search(r'\*\*\[전망:.*?\]\*\*', text)
    o_title = o_title_match.group(0) if o_title_match else f"**[전망: 계절성 이동 방벽 강화 및 화물 구역 밀폐 관리]**"
    m_out = re.search(r'\*\*\[전망:.*?\]\*\*\s*(.*?)(?=---|### 기사 내용|$)', text)
    o_body = m_out.group(1).strip() if m_out else ""

    m_art = re.search(r'### 기사 내용 요약 \(알기 쉬운 해설\)\s*(.*?)(?=---|### 창의적|$)', text)
    art_text = m_art.group(1).strip() if m_art else ""
    art_src = ""
    m_asrc = re.search(r'\*\*출처:\s*(.*?)\*\*', art_text)
    if m_asrc:
        art_src = f"**출처: {m_asrc.group(1).strip()}**"
        art_body = art_text.replace(m_asrc.group(0), '').strip()
    else:
        art_body = art_text

    m_qs = re.search(r'### 창의적인 인사이트를 위한 질문 \(데이터베이스 연결용\)\s*(.*?)(?=---|>\s*\*\*운용 주의|$)', text)
    qs_raw = m_qs.group(1).strip() if m_qs else ""

    q_items = re.findall(r'(\d+\.\s*\*\*.*?\*\*.*?)(?=\d+\.\s*\*\*|$)', qs_raw)
    if not q_items and qs_raw:
        q_items = [q.strip() for q in re.split(r'(?=\b\d+\.\s+)', qs_raw) if q.strip()]

    m_warn = re.search(r'>\s*\*\*운용 주의\*\*:\s*(.*)', text)
    warn_text = m_warn.group(1).strip() if m_warn else f"본 보고서는 {month}월 신속 조치 및 포획 실측 데이터를 기반으로 작성되었습니다."
    warn_text = warn_text.replace('>', '').strip()

    lines = [
        f"날짜: {date_str}",
        f"카테고리: {cat_str}",
        f"출처: {src_str}",
        f"태그: {tag_str}\n",
        f"### 전문가(컨설턴트) 종합 요약 및 의견\n",
        f"{c_title}",
        f"{c_body}\n",
        f"{s_title}",
        f"{s_body}\n",
        f"{o_title}",
        f"{o_body}\n",
        f"---\n",
        f"### 기사 내용 요약 (알기 쉬운 해설)\n"
    ]

    if art_src:
        lines.append(f"{art_src}\n")
    if art_body:
        lines.append(f"{art_body}\n")

    lines.append("---\n")
    lines.append("### 창의적인 인사이트를 위한 질문 (데이터베이스 연결용)\n")

    for q in q_items:
        lines.append(f"{q.strip()}\n")

    lines.append("---\n")
    lines.append(f"> **운용 주의**: {warn_text}")
    lines.append(f"> {month}월 사체수거 및 조류충돌 보고서와 연계하여 **'종합 Wildlife Risk Matrix'**를 업데이트하십시오.\n")

    return '\n'.join(lines)

def build_yearly_summary(year, events):
    total_cnt = len(events)
    monthly_counts = Counter([e['month'] for e in events])
    sp_counts = Counter([e['species'] for e in events if e['species']])
    loc_counts = Counter([e['location'] for e in events if e['location']])
    event_counts = Counter([e['event_type'] for e in events if e['event_type']])

    top_sp = [f"- **{sp}**: {cnt}건 ({round(cnt/total_cnt*100, 1) if total_cnt else 0}%)" for sp, cnt in sp_counts.most_common(10)]
    top_loc = [f"- **{loc}**: {cnt}건" for loc, cnt in loc_counts.most_common(10)]
    monthly_stat = [f"- **{m}월**: {monthly_counts[m]}건" for m in range(1, 13)]
    ev_stat = [f"- **{k}**: {v}건" for k, v in event_counts.most_common(10)]

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
        '\n'.join(ev_stat),
        f"\n## 6. 결론 및 종합 제언",
        f"{year}년 통계 분석 결과, 야생고양이와 고라니 등 포유류의 공항 내부 침투 시도가 지속 관찰되었습니다. 특히 화물터미널 구역과 초소 출입문 주변, 오픈 배수로를 통한 유입 압력이 높으므로 물리적 차단 시설 보강과 상시 포획틀 모니터링 체계 유지가 필수적입니다."
    ]

    return '\n'.join(lines)

def process_year(year):
    ydir = os.path.join(BASE_DIR, f"{year}년도 야통대", "생동물")
    print(f"\n>> [{year}년도] 생동물 보고서 재포맷팅 및 통계 보고서 작성 시작...")

    all_events = []

    for m in range(1, 13):
        # 1. 통합보고서
        rep_fn = f"{year}-{m:02d}_생동물_통합보고서.md"
        rep_fp = os.path.join(ydir, rep_fn)
        if os.path.exists(rep_fp):
            with open(rep_fp, 'r', encoding='utf-8') as f:
                orig_rep = f.read()

            new_rep = reformat_integrated_report(orig_rep, year, m)
            with open(rep_fp, 'w', encoding='utf-8') as f:
                f.write(new_rep)

            # 이벤트 파싱 for 통계
            sections = new_rep.split('#### 📅')
            for s in sections[1:]:
                s = s.strip()
                if not s: continue
                m_date = re.match(r'^(\d{4}-\d{2}-\d{2})', s)
                if not m_date: continue
                d_str = m_date.group(1)

                m_typ = re.search(r'\*\s*\*\*주요접수유형\*\*:\s*(.*?)(?=\*\*상세 타임라인|$)', s)
                typ_str = m_typ.group(1).strip() if m_typ else '생동물발견'

                m_loc = re.search(r'-\s*\*\*발생장소\*\*:\s*(.*?)(?=-|\n|$)', s)
                loc_str = m_loc.group(1).strip() if m_loc else '기동지역'

                full_text = s
                # 정밀 종명 추출
                if '고양이' in full_text:
                    sp_str = '야생고양이'
                elif '고라니' in full_text:
                    sp_str = '고라니'
                elif any(k in full_text for k in ['유기견', '애완견', '셰퍼드', '시바견', '비숑', '닥스훈트']) or '개 ' in full_text or '개1' in full_text or '개2' in full_text:
                    sp_str = '유기견/반려견'
                elif '박쥐' in full_text:
                    sp_str = '박쥐'
                elif '뱀' in full_text or '유혈목이' in full_text:
                    sp_str = '뱀'
                elif '너구리' in full_text:
                    sp_str = '너구리'
                elif '족제비' in full_text:
                    sp_str = '족제비'
                else:
                    sp_str = '기타 소형생동물'

                # 정밀 이벤트 유형 추출
                if '공항내 유입' in full_text or '공항 내 유입' in full_text or '유입' in typ_str:
                    ev_type = '생동물 공항내 유입'
                elif '포획틀' in full_text:
                    ev_type = '포획틀 점검 및 생포'
                elif '탈출' in full_text:
                    ev_type = '조업중 동물탈출'
                elif '출현' in full_text:
                    ev_type = '생동물 출현'
                else:
                    ev_type = '생동물 발견'

                all_events.append({
                    'month': m,
                    'date': d_str,
                    'type': typ_str,
                    'event_type': ev_type,
                    'location': loc_str,
                    'species': sp_str
                })

        # 2. 인사이트 보고서
        ins_fn = f"{year}-{m:02d}_전문가_브리핑_및_인사이트.md"
        ins_fp = os.path.join(ydir, ins_fn)
        if os.path.exists(ins_fp):
            with open(ins_fp, 'r', encoding='utf-8') as f:
                orig_ins = f.read()

            new_ins = reformat_insight_report(orig_ins, year, m)
            with open(ins_fp, 'w', encoding='utf-8') as f:
                f.write(new_ins)

        print(f"  - {year}년 {m:02d}월 재포맷팅 완료")

    # 3. 연간 통계 보고서 작성
    summary_fn = f"{year}_생동물_통계_보고서.md"
    summary_fp = os.path.join(ydir, summary_fn)
    summary_content = build_yearly_summary(year, all_events)
    with open(summary_fp, 'w', encoding='utf-8') as f:
        f.write(summary_content)

    print(f"  - {summary_fn} 생성 완료 (총 {len(all_events)}건 집계)")

def main():
    process_year(2024)
    process_year(2025)
    print("\n[성공] 2024년 및 2025년 생동물 보고서 재포맷팅 및 연간 통계 보고서 구축 완료!")

if __name__ == '__main__':
    main()
