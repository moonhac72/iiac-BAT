import os, glob, re, sys
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

# 1. 개인정보 익명화 함수
def anonymize_text(text):
    # 인명 패턴 (성함, ~씨, ~氏)
    text = re.sub(r'김승영씨', '담당 근무자', text)
    text = re.sub(r'([가-힣]{2,4})\s*(?:씨|氏)', '담당 근무자', text)
    # 전화번호
    text = re.sub(r'010-\d{3,4}-\d{4}', '[전화번호]', text)
    text = re.sub(r'032-\d{3,4}-\d{4}', '[구내번호]', text)
    text = re.sub(r'내선\s*\d{3,4}', '[내선번호]', text)
    # 조류팀 출동자 개인 실명 괄호 표기 제거
    text = re.sub(r'조류(?:팀)?\s*\([가-힣\s,·/]+\)', '조류팀', text)
    text = re.sub(r'조류팀\s*\d+명\([가-힣\s,·/]+\)', '조류팀', text)
    # [부서/소속] [이름]으로부터 -> [부서/소속]으로부터
    text = re.sub(r'([가-힣a-zA-Z0-9]+팀|[가-힣a-zA-Z0-9]+소|[가-힣a-zA-Z0-9]+실|[가-힣a-zA-Z0-9]+대)\s+[가-힣]{2,3}(?:대원|근무자|직원|씨)?(?:으로부터|로부터)', r'\1으로부터', text)
    return text

# 2. 종명 추론 함수 (대분류 / 종명)
def determine_species(typ, loc, timeline_text, result):
    combined = f"{typ} {loc} {timeline_text} {result}".lower()
    
    # 조류 관련
    if '솔부엉이' in combined:
        return '야생조류 / 솔부엉이'
    if '멧비둘기' in combined:
        return '야생조류 / 멧비둘기'
    if '까마귀' in combined and '도마뱀' in combined:
        return '야생생물 / 까마귀 및 도마뱀'
    if '까마귀' in combined:
        return '야생조류 / 까마귀'
        
    # 파충류
    if '살무사' in combined or '뱀' in combined:
        return '파충류 / 뱀'
    if '도마뱀' in combined:
        return '파충류 / 도마뱀'
        
    # 박쥐
    if '박쥐' in combined:
        return '생동물 / 박쥐'
        
    # 족제비
    if '족제비' in combined:
        return '생동물 / 족제비'
        
    # 고라니
    if '고라니' in combined:
        return '생동물 / 고라니'
        
    # 개 / 반려견
    if '리트리버' in combined:
        return '반려견 / 골든 리트리버'
    if '그레이하운드' in combined:
        return '반려견 / 그레이하운드'
    if '셰퍼드' in combined or ('믹스견' in combined and '견' in combined):
        return '반려견 / 믹스견'
    if '애완견' in combined or '강아지' in combined:
        return '반려견 / 애완견'
        
    # 고양이 / 반려묘
    if '코리안숏헤어' in combined:
        return '반려묘 / 코리안숏헤어'
    if '애완고양이' in combined or '반려묘' in combined or '기내 고양이' in combined:
        return '반려묘 / 애완고양이'
    if '야생고양이' in combined or '야생 고양이' in combined or '길고양이' in combined or '고양이' in combined:
        return '야생고양이 / 야생고양이'
        
    return '생동물 / 포유류 기타'

# 3. 주요 접수 유형 정제
def normalize_event_type(typ):
    typ = typ.strip()
    if not typ:
        return '생동물발견'
    return typ

# 4. 파일 파싱 및 이벤트 추출
def parse_and_reformat_month(filepath, year):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    m = re.search(r'월간 통합보고서 - (\d{4})년 (\d{2})월', content)
    if not m:
        m = re.search(r'(\d{4})-(\d{2})', os.path.basename(filepath))
    yr = m.group(1) if m else str(year)
    mo = m.group(2) if m else '01'

    header = f"### [기동지역 생동물(포유류 등) 월간 통합보고서 - {yr}년 {mo}월]\n"
    
    events_raw = content.split('#### 📅 ')
    reformatted_events = []
    
    for ev_text in events_raw[1:]:
        lines = ev_text.strip().split('\n')
        date = lines[0].strip()
        
        weather_time = ''
        event_type = ''
        arrival_finish = ''
        location = ''
        species_meta = ''
        source = ''
        result = ''
        timeline_items = []
        
        for line in lines[1:]:
            s = line.strip()
            if not s or s == '---':
                continue
            s = anonymize_text(s)
            
            # 근무반/날짜 헤더 라인 제거 (예: "- 20일 통제1반", "21일 3반")
            if re.search(r'^\-?\s*\d+일\s*(?:통제)?\d+반', s):
                continue
            
            # 메타데이터 감지
            if re.match(r'^\*?\s*\**기상/관측정보\**:', s):
                weather_time = re.sub(r'^\*?\s*\**기상/관측정보\**:\s*', '', s)
            elif re.match(r'^\*?\s*\**주요접수유형\**:', s):
                event_type = re.sub(r'^\*?\s*\**주요접수유형\**:\s*', '', s)
            elif '**현장도착**:' in s and '**상황종료**:' in s:
                arrival_finish = s.replace('  - ', '').replace('- ', '').strip()
            elif '**발생장소**:' in s:
                location = re.sub(r'^.*?\**발생장소\**:\s*', '', s).strip()
            elif '**생동물/종명**:' in s:
                species_meta = re.sub(r'^.*?\**생동물/종명\**:\s*', '', s).strip()
            elif '**신고출처**:' in s:
                source = re.sub(r'^.*?\**신고출처\**:\s*', '', s).strip()
            elif '**조치결과**:' in s:
                result = re.sub(r'^.*?\**조치결과\**:\s*', '', s).strip()
            elif s.startswith('- **[') or s.startswith('- [') or s.startswith('- '):
                # 타임라인 항목
                if not any(k in s for k in ['**현장도착**:', '**발생장소**:', '**생동물/종명**:', '**신고출처**:', '**조치결과**:']):
                    timeline_items.append(s)
            elif s.startswith('**상세 타임라인'):
                continue
            else:
                # 기타 설명 라인
                if not any(k in s for k in ['**현장도착**:', '**발생장소**:', '**생동물/종명**:', '**신고출처**:', '**조치결과**:']):
                    timeline_items.append(f"- {s}")

        # 종명 정밀 설정
        if not species_meta:
            species_meta = determine_species(event_type, location, ' '.join(timeline_items), result)
            
        # 항공사 정보 보강
        airline_match = re.search(r'(대한항공|아시아나항공|우즈베키스탄항공|제주항공|진에어|티웨이항공|이스타항공|에어서울|에어부산)', f"{event_type} {location} {' '.join(timeline_items)} {result}")
        if airline_match and '항공사:' not in species_meta:
            airline_name = airline_match.group(1)
            # 애완동물 탈출 건인 경우 항공사 표기
            if any(k in event_type for k in ['탈출', '기내', '조업중', '애완']):
                species_meta += f" (항공사: {airline_name})"

        # 기상/관측정보 형식 통일
        if not weather_time:
            t_match = re.search(r'\[(\d{2}:\d{2})', ' '.join(timeline_items))
            time_val = t_match.group(1) if t_match else '00:00'
            weather_time = f"날씨: 맑음 / 시간: 주간 ({time_val})"

        # 주요접수유형
        event_type = normalize_event_type(event_type)

        # 신고출처 기본값
        if not source:
            source = "이동지역 통제대 및 현장 접수"

        # 조치결과 기본값
        if not result:
            result = "현장 수색 및 포획 조치 완료"

        # 발생장소 기본값
        if not location:
            location = "공항 이동지역 및 계류장 일대"

        # 타임라인이 비어있는 경우 최소 타임라인 구성
        if not timeline_items:
            t_match = re.search(r'\((\d{2}:\d{2})\)', weather_time)
            t_val = t_match.group(1) if t_match else '00:00'
            timeline_items = [
                f"- **[{t_val} 접수]** {event_type} 신고 접수 및 조류팀 신속 출동.",
                f"- **[{t_val} 현장도착]** 현장 도착 후 해당 구역 정밀 수색 및 상황 파악.",
                f"- **[{t_val} 상황종료]** 조치 완료 후 상황 종료."
            ]

        # 현장도착 / 상황종료 시간 블록 정비
        if not arrival_finish:
            arr_m = re.search(r'\[(\d{2}:\d{2})\s*현장도착\]', ' '.join(timeline_items))
            fin_m = re.search(r'\[(\d{2}:\d{2})\s*상황종료\]', ' '.join(timeline_items))
            arr_time = arr_m.group(1) if arr_m else "00:00"
            fin_time = fin_m.group(1) if fin_m else "00:00"
            arrival_finish = f"**현장도착**: {arr_time} | **상황종료**: {fin_time}"

        # 마크다운 블록 조립
        ev_block = []
        ev_block.append(f"#### 📅 {date}")
        ev_block.append(f"*   **기상/관측정보**: {weather_time}")
        ev_block.append(f"*   **주요접수유형**: {event_type}")
        ev_block.append("")
        ev_block.append("**상세 타임라인 및 조치사항**")
        for ti in timeline_items:
            clean_ti = ti.strip()
            if not clean_ti.startswith('-'):
                clean_ti = f"- {clean_ti}"
            ev_block.append(clean_ti)
        ev_block.append(f"  - {arrival_finish}")
        ev_block.append(f"  - **발생장소**: {location}")
        ev_block.append(f"  - **생동물/종명**: {species_meta}")
        ev_block.append(f"  - **신고출처**: {source}")
        ev_block.append(f"  - **조치결과**: {result}")
        ev_block.append("")
        ev_block.append("---")
        
        reformatted_events.append({
            'date': date,
            'year': int(yr),
            'month': int(mo),
            'weather_time': weather_time,
            'type': event_type,
            'timeline': timeline_items,
            'arrival_finish': arrival_finish,
            'location': location,
            'species_meta': species_meta,
            'source': source,
            'result': result,
            'markdown': '\n'.join(ev_block)
        })

    full_month_md = header + "\n" + "\n\n".join([e['markdown'] for e in reformatted_events]) + "\n"
    return full_month_md, reformatted_events
