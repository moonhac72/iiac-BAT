import os, glob, re, sys
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

def parse_report_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # get month header
    m = re.search(r'월간 통합보고서 - (\d{4})년 (\d{2})월', content)
    year = m.group(1) if m else '2024'
    month = m.group(2) if m else '01'
    
    events = []
    chunks = content.split('#### 📅 ')
    for ch in chunks[1:]:
        lines = ch.strip().split('\n')
        date = lines[0].strip()
        
        info = {
            'date': date,
            'weather_time': '',
            'type': '',
            'timeline': [],
            'arrival_finish': '',
            'location': '',
            'species_meta': '',
            'source': '',
            'result': ''
        }
        
        timeline_mode = False
        for line in lines[1:]:
            s = line.strip()
            if not s or s == '---':
                continue
            if s.startswith('*   **기상/관측정보**:') or s.startswith('**기상/관측정보**:') or s.startswith('* **기상/관측정보**:') or s.startswith('* 기상/관측정보:'):
                info['weather_time'] = re.sub(r'^\*?\s*\**기상/관측정보\**:\s*', '', s)
            elif s.startswith('*   **주요접수유형**:') or s.startswith('**주요접수유형**:') or s.startswith('* **주요접수유형**:') or s.startswith('* 주요접수유형:'):
                info['type'] = re.sub(r'^\*?\s*\**주요접수유형\**:\s*', '', s)
            elif '**상세 타임라인 및 조치사항**' in s:
                timeline_mode = True
            elif s.startswith('- **현장도착**:') or s.startswith('**현장도착**:') or '- **현장도착**:' in s:
                info['arrival_finish'] = s.replace('  - ', '').replace('- ', '')
            elif s.startswith('- **발생장소**:') or s.startswith('**발생장소**:') or '- **발생장소**:' in s:
                info['location'] = re.sub(r'^.*?\**발생장소\**:\s*', '', s)
            elif s.startswith('- **생동물/종명**:') or s.startswith('**생동물/종명**:') or '- **생동물/종명**:' in s:
                info['species_meta'] = re.sub(r'^.*?\**생동물/종명\**:\s*', '', s)
            elif s.startswith('- **신고출처**:') or s.startswith('**신고출처**:') or '- **신고출처**:' in s:
                info['source'] = re.sub(r'^.*?\**신고출처\**:\s*', '', s)
            elif s.startswith('- **조치결과**:') or s.startswith('**조치결과**:') or '- **조치결과**:' in s:
                info['result'] = re.sub(r'^.*?\**조치결과\**:\s*', '', s)
            elif timeline_mode:
                if s.startswith('- **[') or s.startswith('- [') or s.startswith('- '):
                    info['timeline'].append(s)
        events.append(info)
    return year, month, events

for y in [2024, 2025]:
    folder = f'10_Wiki/🛠️ Projects/야통대_자료/{y}년도 야통대/생동물'
    files = sorted(glob.glob(f'{folder}/{y}-*_생동물_통합보고서.md'))
    all_events = []
    month_counts = Counter()
    locations = Counter()
    species_counter = Counter()
    type_counter = Counter()
    
    for f in files:
        yr, mo, evs = parse_report_file(f)
        month_counts[int(mo)] += len(evs)
        for ev in evs:
            all_events.append(ev)
            # location clean
            loc = ev['location']
            loc = loc.replace('A/S ', '').replace('L/S ', '').strip()
            locations[loc] += 1
            
            # infer species
            txt = (ev['type'] + ' ' + ev['result'] + ' ' + ' '.join(ev['timeline'])).lower()
            sp = '기타'
            if '코리안숏헤어' in txt: sp = '코리안숏헤어'
            elif '리트리버' in txt: sp = '골든 리트리버'
            elif '그레이하운드' in txt: sp = '그레이하운드'
            elif '믹스견' in txt or '셰퍼드' in txt: sp = '믹스견'
            elif '애완견' in txt or '강아지' in txt: sp = '애완견'
            elif '애완고양이' in txt or '반려묘' in txt: sp = '애완고양이'
            elif '야생고양이' in txt or '야생 고양이' in txt or '고양이' in txt: sp = '야생고양이'
            elif '고라니' in txt: sp = '고라니'
            elif '족제비' in txt: sp = '족제비'
            elif '살무사' in txt or '뱀' in txt: sp = '뱀'
            elif '도마뱀' in txt: sp = '도마뱀'
            elif '박쥐' in txt: sp = '박쥐'
            elif '솔부엉이' in txt: sp = '솔부엉이'
            elif '멧비둘기' in txt: sp = '멧비둘기'
            elif '까마귀' in txt: sp = '까마귀'
            species_counter[sp] += 1
            
            # types
            t = ev['type']
            main_t = '생동물발견'
            if '탈출' in t:
                if '기내' in t: main_t = '조업중 동물탈출(기내)'
                else: main_t = '조업중 동물탈출'
            elif '공항내 유입' in t or '유입' in t:
                main_t = '생동물 공항내 유입'
            elif '포획틀' in t:
                main_t = '생동물발견(포획틀 생포)'
            elif '발견' in t:
                main_t = '생동물발견'
            type_counter[main_t] += 1

    print(f'=== YEAR {y} STATS (Total: {len(all_events)}) ===')
    print('Monthly:', sorted(month_counts.items()))
    print('Top Locations:', locations.most_common(10))
    print('Species:', species_counter.most_common())
    print('Types:', type_counter.most_common())
    print()
