#!/usr/bin/env python3
"""
Naver Blog Automated Private Publisher (Playwright based)
P-Reinforce Architecture / 야생동물통제대 지식 창고 연동

Naver Blog Info:
- Blog ID: moonhac72
- Blog Name: 보물지도
- Write URL: https://blog.naver.com/moonhac72/postwrite
"""

import sys
import os
import re
import argparse
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

NAVER_ID = "moonhac72"
BLOG_NAME = "보물지도"
WRITE_URL = f"https://blog.naver.com/{NAVER_ID}/postwrite"
USER_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".naver_user_data")

def parse_markdown(file_path):
    """
    마크다운 파일에서 제목, 태그, 본문 및 카테고리를 추출합니다.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    title = Path(file_path).stem
    tags = []
    category = "야생동물통제대"
    
    # 1. YAML Frontmatter 추출
    frontmatter_match = re.match(r'^---\s*\n(.*?)\n---\s*\n(.*)$', content, re.DOTALL)
    if frontmatter_match:
        fm_text, body = frontmatter_match.groups()
        # 태그 추출
        tag_match = re.search(r'tags:\s*\n((?:\s*-\s*.*\n?)+)', fm_text)
        if tag_match:
            raw_tags = tag_match.group(1)
            tags = [t.strip().lstrip('-').strip() for t in raw_tags.strip().split('\n') if t.strip()]
        
        # 카테고리 추출
        cat_match = re.search(r'category:\s*(.+)', fm_text)
        if cat_match:
            category = cat_match.group(1).strip()
    else:
        body = content

    # 2. 본문 첫 번째 H1 제목이 있으면 제목으로 활용
    h1_match = re.search(r'^#\s+(.+)$', body, re.MULTILINE)
    if h1_match:
        title = h1_match.group(1).strip()
        body = re.sub(r'^#\s+.+\n?', '', body, count=1, flags=re.MULTILINE)

    # 마크다운 파일명 정리
    clean_title = re.sub(r'^\d{4}-\d{2}-\d{2}_', '', title)
    clean_title = clean_title.replace('_', ' ').replace('.md', '')

    return clean_title, body.strip(), tags, category

def upload_to_naver_blog(file_path, is_private=True, headless=False, category_name="야생동물통제대"):
    title, body, tags, file_category = parse_markdown(file_path)
    target_category = category_name or file_category

    print(f"🌐 블로그: {BLOG_NAME} (https://blog.naver.com/{NAVER_ID})")
    print(f"📄 파일: {os.path.basename(file_path)}")
    print(f"📌 제목: {title}")
    print(f"📂 카테고리: {target_category}")
    print(f"🏷️ 태그: {tags}")
    print(f"🔒 공개 여부: {'비공개' if is_private else '공개'}")
    print("-" * 60)

    os.makedirs(USER_DATA_DIR, exist_ok=True)

    with sync_playwright() as p:
        # Chrome 사용자 데이터 디렉터리로 실행하여 로그인 세션 유지
        context = p.chromium.launch_persistent_context(
            user_data_dir=USER_DATA_DIR,
            headless=headless,
            channel="chrome",
            args=["--no-sandbox", "--disable-setuid-sandbox"]
        )
        page = context.pages[0] if context.pages else context.new_page()

        # 네이버 블로그 에디터 페이지 접속
        print(f"🌐 [{BLOG_NAME}] 스마트에디터 접속 중... ({WRITE_URL})")
        page.goto(WRITE_URL, wait_until="networkidle")
        time.sleep(2)

        # 로그인 여부 확인
        if "nid.naver.com" in page.url or "login" in page.url:
            print("⚠️ 네이버 로그인이 필요합니다!")
            print(f"💡 네이버 ID [{NAVER_ID}]로 열린 브라우저 창에서 로그인 후 Enter 키를 눌러주세요...")
            input("로그인 완료 후 엔터를 입력하세요: ")

        print("📝 글 작성 페이지 진입 확인...")
        time.sleep(2)

        print("✅ 네이버 블로그 포스팅 준비가 완료되었습니다.")
        context.close()

def main():
    parser = argparse.ArgumentParser(description=f"네이버 블로그 '{BLOG_NAME}'({NAVER_ID}) 자동 비공개 업로더")
    parser.add_argument("filepath", nargs="?", help="업로드할 마크다운 파일 경로")
    parser.add_argument("--category", default="야생동물통제대", help="저장할 카테고리명 (기본값: 야생동물통제대)")
    parser.add_argument("--public", action="store_true", help="공개로 작성 (기본값: 비공개)")
    parser.add_argument("--headless", action="store_true", help="브라우저 창 숨기기")

    args = parser.parse_args()

    if not args.filepath:
        parser.print_help()
        sys.exit(1)

    if not os.path.exists(args.filepath):
        print(f"❌ 오류: 파일을 찾을 수 없습니다: {args.filepath}")
        sys.exit(1)

    upload_to_naver_blog(
        args.filepath,
        is_private=not args.public,
        headless=args.headless,
        category_name=args.category
    )

if __name__ == "__main__":
    main()
