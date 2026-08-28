# P-Reinforce RL Policy & Categorization Rules

## Reward Weights
$$R = w_1(\text{Categorization Accuracy: 0.4}) + w_2(\text{Graph Connectivity: 0.3}) + w_3(\text{User Satisfaction: 0.3})$$

## Classification Rules
1. **Raw Ingestion (`00_Raw/📥 Inbox`)**:
   - 사용자가 새롭게 수집/입력한 원시 데이터는 Inbox에 임시 보관 후 분석 및 정제 절차를 거친다.
   - 지식화가 완료된 파일은 `00_Raw/🗄️ Archive/YYYY/MM/`로 이동 저장한다.

2. **Wiki Layer (`10_Wiki/`)**:
   - `🛠️ Projects/`: 야통대 현장 일지, 포획, 신고, 사체 수거 등 목표 지향성 데이터 프로젝트 및 E-Book 집필 과제.
   - `💡 Topics/`: 주간브리핑, 월간브리핑, 주제별 분석 보고서 등 개념/뉴스/지식 요약 문서.
   - `⚖️ Decisions/`: 주요 정책 판단 및 조치 결정 내역.
   - `🚀 Skills/`: 운영 지침, 시스템 프롬프트, 분석 메커니즘 가이드라인.

3. **Flat Directory Rule (신문기사 및 브리핑)**:
   - 브리핑 및 주제별 분석 보고서는 대분류 폴더 직하위에 `YYYY-MM-DD_제목.md` 형태로 수록한다.
