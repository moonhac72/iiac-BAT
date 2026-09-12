# -*- coding: utf-8 -*-
"""
야통대 전체 자료 문서에 대한 개인정보(근무자/신고자/접수자 이름, 전화번호 등) 전수 삭제/익명화 스크립트
"""

import os
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r'c:\Users\moonh\Documents\iiac-BAT\10_Wiki\🛠️ Projects\야통대_자료'

TARGET_NAMES = [
    '강영운', '강이중', '강지구', '고기철', '권혁락', '김광진', '김남황', '김동희', '김범수', '김부건',
    '김성찬', '김수경', '김영수', '김용선', '김용택', '김원기', '김윤식', '김은수', '김재민', '김정수',
    '김정현', '김지운', '김진혁', '김진현', '김태영', '김현철', '남중수', '남학수', '문정완', '문정환',
    '민찬홍', '박배원', '박세진', '박영준', '박정범', '배명현', '서우탁', '손병호', '신정현', '양민규',
    '여정현', '오만수', '유경환', '유명호', '유승범', '유재민', '윤의연', '윤종훈', '윤현식', '이기우',
    '이기태', '이무형', '이상현', '이상헌', '이성호', '이윤상', '이정모', '이종혁', '이창환', '정명운',
    '정상훈', '정성필', '정영일', '조두현', '최민수', '최성준', '한석범', '한종철', '한준원', '현종현',
    '김성수', '이상준', '김현수', '최현수', '정제면', '최종석', '김민우', '강민석', '김익종',
    '정재호', '서한민', '오경수', '김우람', '박진원', '곽홍철', '정운용', '조의수', '박안준', '박래빈',
    '송원석', '김민국', '김승영', '류강남', '김영기', '박진숙', '한기완', '이은지', '고동형', '최규재',
    '박홍철', '차국호', '권정률', '유의정', '김준희', '황효준', '송형구', '구본준', '최희영', '이승길',
    '박성필', '유현재', '서현석', '임승빈', '이주성', '여인황', '조남혁', '박만규', '김윤주', '엄강산',
    '안동은', '허지희', '최지용', '장윤빈', '허성봉', '권성일', '유승호', '장형석', '신정태', '양승필',
    '이인식', '김일호', '김양중', '오창식', '허택경', '이은정', '서봉현'
]

# 긴 이름 우선 정렬
TARGET_NAMES = sorted(list(set(TARGET_NAMES)), key=lambda x: -len(x))
NAMES_PATTERN = '|'.join(map(re.escape, TARGET_NAMES))

def anonymize_text(content):
    s = content

    # 1. 전화번호, 내선번호, 휴대전화번호 일체 삭제
    s = re.sub(r'※?\s*[가-힣]{0,4}\s*\(?\b\d{2,4}[-\s)]*\d{3,4}[-\s)]*\d{4}?\)?', '', s)
    s = re.sub(r'\(?[가-힣]{0,4}\s*010[-\s]?\d{4}[-\s]?\d{4}\)?', '', s)
    s = re.sub(r'\b010[-\s]?\d{4}[-\s]?\d{4}\b', '', s)
    s = re.sub(r'\b\d{2,4}-\d{3,4}-\d{4}\b', '', s)
    s = re.sub(r'\(\d{3,4}-\d{4}\)', '', s)

    # 2. [부서/기관/소속] + [이름][씨/氏/직함]? + [으로/로]?부터 -> [부서/기관/소속]으로부터
    s = re.sub(r'([가-힣A-Za-z0-9]+(?:상황실|센터|데스크|초소|관리소|시설|통제대|통제실|경비대|청사|팀|소|본대|지점|부서|역|파트))\s+[가-힣]{2,4}\s*(?:씨|氏|파트장|과장|팀장|대리|주임|반장)?\s*(?:으로|로)?\s*부터', r'\1로부터', s)

    # 3. [부서/기관] + [이름][씨/氏] -> [부서/기관]
    s = re.sub(r'([가-힣A-Za-z0-9]+(?:상황실|센터|데스크|초소|관리소|시설|통제대|통제실|경비대|청사|팀|소|본대))\s+[가-힣]{2,4}\s*(?:씨|氏)', r'\1', s)

    # 4. 단독 [이름][씨/氏]? [으로/로]?부터 -> 신고자로부터
    s = re.sub(r'(?<![가-힣])[가-힣]{2,4}\s*(?:씨|氏)\s*(?:으로|로)?\s*부터', '신고자로부터', s)
    s = re.sub(r'(?<![가-힣])[가-힣]{2,4}\s*(?:으로|로)\s*부터', '신고자로부터', s)

    # 5. 직함 및 인원 표기 정제
    s = re.sub(r'[가-힣]{2,4}\s*파트장\s*외\s*\d+명', '지원인력', s)
    s = re.sub(r'[가-힣]{2,4}\s*(?:파트장|과장|팀장|대리|주임|반장)\b', '', s)

    # 6. [이름][씨/氏] 단독 제거
    s = re.sub(r'[가-힣]{2,4}\s*(?:씨|氏)', '', s)

    # 7. 괄호 안 출동자/근무자 명단 제거
    # e.g., (한준원, 정영일) -> (통제반) 또는 조류팀(한준원, 정영일) -> 조류팀
    s = re.sub(rf'(조류\s*\d+호)\s*\((?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\)', r'\1', s)
    s = re.sub(rf'(조류\s*팀)\s*\((?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\)', r'\1', s)
    s = re.sub(rf'(통제\s*반|근무\s*조)\s*\((?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\)', r'\1', s)
    s = re.sub(rf'\((?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\)', '', s)

    # 8. 대원/요원/팀과 결합된 인명
    # e.g., '이창환, 정상훈 대원' -> '현장 대원', '김익종 대원' -> '현장 대원'
    s = re.sub(rf'(?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\s*대원', '현장 대원', s)
    s = re.sub(rf'(?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))*\s*요원', '현장 요원', s)
    s = re.sub(rf'(?:{NAMES_PATTERN})/(?:{NAMES_PATTERN})\s*팀', '통제팀', s)
    s = re.sub(rf'(?:{NAMES_PATTERN})/(?:{NAMES_PATTERN})\s*조', '통제조', s)
    s = re.sub(rf'(?:{NAMES_PATTERN})\s*등', '통제반 등', s)

    # 9. 인명 나열 표기 (쉼표나 슬래시로 연결된 이름들)
    # e.g., '김정수/최종석/김민우/강민석 통합' -> '통제반 통합'
    s = re.sub(rf'(?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))+\s*통합', '통제반 통합', s)
    s = re.sub(rf'(?:{NAMES_PATTERN})(?:[\s,./]+(?:{NAMES_PATTERN}))+', '통제반', s)

    # 10. 단독으로 남아있는 특정 타겟 인명 제거 (일반 단어 오인 방지를 위해 단어 경계 및 안전 패턴 적용)
    for name in TARGET_NAMES:
        # 단독 이름 뒤에 조사나 공백이 붙은 경우
        s = re.sub(rf'\b{re.escape(name)}(?:은|는|이|가|을|를|과|와|의|에게|한테)?\b', '', s)
        s = re.sub(rf'\s+{re.escape(name)}\s+', ' ', s)

    # 11. 잔여 빈 괄호 및 문장 부호 정리
    s = re.sub(r'\(\s*\)', '', s)
    s = re.sub(r'\[\s*\]', '', s)
    s = re.sub(r'※', '', s)
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\n{3,}', '\n\n', s)

    return s

def main():
    print("=== 야통대 전체 자료 개인정보(이름/전화번호) 전수 삭제 작업 시작 ===")
    processed_count = 0
    modified_count = 0

    for root, dirs, files in os.walk(BASE_DIR):
        for fname in files:
            if not fname.endswith('.md'):
                continue
            fpath = os.path.join(root, fname)
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    orig = f.read()

                cleaned = anonymize_text(orig)

                if cleaned != orig:
                    with open(fpath, 'w', encoding='utf-8') as f:
                        f.write(cleaned)
                    modified_count += 1

                processed_count += 1
            except Exception as e:
                print(f"Error processing {fpath}: {e}")

    print(f"\n[작업 완료]")
    print(f"- 검사된 전체 마크다운 파일 수: {processed_count}개")
    print(f"- 개인정보(이름/전화번호)가 삭제·정제된 파일 수: {modified_count}개")

if __name__ == '__main__':
    main()
