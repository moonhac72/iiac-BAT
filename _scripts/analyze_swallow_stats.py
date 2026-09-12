import os, sys
import pandas as pd
from collections import Counter

sys.stdout.reconfigure(encoding='utf-8')

resampled_dir = r'_data_raw/resampled'

swallow_patrol = []

# 1. 2017 ~ 2023 영문 컬럼 파일
for yr in [2017, 2018, 2019, 2020, 2021, 2022, 2023]:
    pf = f"{yr}_total_resampled.csv"
    p_path = os.path.join(resampled_dir, pf)
    if os.path.exists(p_path):
        try:
            df = pd.read_csv(p_path, low_memory=False)
            if 'Month' not in df.columns and 'Date' in df.columns:
                df['Month'] = pd.to_datetime(df['Date'], errors='coerce').dt.month
            
            sub = df[df['Month'].isin([8, 9])].copy()
            is_swallow = sub['Species'].astype(str).str.contains('제비|귀제비', na=False) & ~sub['Species'].astype(str).str.contains('족제비', na=False)
            sub_swallow = sub[is_swallow].copy()
            
            # 표준화된 컬럼으로 통일
            std_df = pd.DataFrame({
                'Year': sub_swallow['Year'] if 'Year' in sub_swallow.columns else yr,
                'Month': sub_swallow['Month'],
                'Date': sub_swallow['Date'],
                'Time': sub_swallow['Time'],
                'Zone': sub_swallow['Zone_19'],
                'Species': sub_swallow['Species'],
                'Count': pd.to_numeric(sub_swallow['Count'], errors='coerce').fillna(0),
                'Area_Raw': sub_swallow['Area_Raw'].astype(str),
                'Area_Detail': sub_swallow['Area_Detail'].astype(str),
                'Map_Group': sub_swallow['Map_Group'].astype(str) if 'Map_Group' in sub_swallow.columns else ''
            })
            swallow_patrol.append(std_df)
        except Exception as e:
            print(f'Error reading {pf}: {e}')

# 2. 2024 ~ 2025 한글 컬럼 파일
for yr in [2024, 2025]:
    pf = f"{yr}_total_resampled.csv"
    p_path = os.path.join(resampled_dir, pf)
    if os.path.exists(p_path):
        try:
            df = pd.read_csv(p_path, encoding='utf-8', low_memory=False)
            df['Month'] = pd.to_datetime(df['일자'], errors='coerce').dt.month
            sub = df[df['Month'].isin([8, 9])].copy()
            
            sp_col = sub['보정_조류종'].fillna(sub['원문_조류종']).astype(str)
            is_swallow = sp_col.str.contains('제비|귀제비', na=False) & ~sp_col.str.contains('족제비', na=False)
            sub_swallow = sub[is_swallow].copy()
            
            std_df = pd.DataFrame({
                'Year': yr,
                'Month': sub_swallow['Month'],
                'Date': sub_swallow['일자'],
                'Time': sub_swallow['시간'],
                'Zone': sub_swallow['분석_Zone'],
                'Species': sp_col[is_swallow],
                'Count': pd.to_numeric(sub_swallow['개체수'], errors='coerce').fillna(0),
                'Area_Raw': sub_swallow['원문_지역'].astype(str),
                'Area_Detail': sub_swallow['원문_지역상세'].astype(str),
                'Map_Group': sub_swallow['원문_맵영역그룹'].astype(str)
            })
            swallow_patrol.append(std_df)
        except Exception as e:
            print(f'Error reading {pf}: {e}')

if swallow_patrol:
    all_sp = pd.concat(swallow_patrol, ignore_index=True)
    total_records = len(all_sp)
    total_count = all_sp['Count'].sum()
    print(f"================================================================================")
    print(f"   [인천국제공항 8~9월 제비(귀제비 포함) 9개년 전수 순찰 관측 빅데이터 통계]")
    print(f"   - 분석 대상 기간: 2017년 ~ 2025년 (8월~9월 늦여름·초가을 시즌 전수)")
    print(f"   - 총 관측(순찰 조우) 레코드: {total_records:,}건")
    print(f"   - 총 발견·통제 개체수: {total_count:,.0f}개체")
    print(f"================================================================================\n")

    print("▶ [1] 19개 표준 존(Zone_19)별 출현 빈도 및 개체수 랭킹:")
    z_stat = all_sp.groupby('Zone').agg(
        관측건수=('Count', 'count'),
        총개체수=('Count', 'sum'),
        평균개체수=('Count', 'mean')
    ).sort_values(by='관측건수', ascending=False)
    z_stat['건수비율(%)'] = (z_stat['관측건수'] / total_records) * 100
    z_stat['개체수비율(%)'] = (z_stat['총개체수'] / total_count) * 100
    print(z_stat[['관측건수', '건수비율(%)', '총개체수', '개체수비율(%)', '평균개체수']].to_string())

    print("\n▶ [2] 다발 출현 세부 녹지대 및 구역 TOP 20:")
    all_sp['Loc_Full'] = all_sp['Area_Raw'].str.strip() + ' ' + all_sp['Area_Detail'].str.strip()
    loc_stat = all_sp.groupby(['Loc_Full', 'Zone']).agg(
        관측건수=('Count', 'count'),
        총개체수=('Count', 'sum')
    ).reset_index().sort_values(by='관측건수', ascending=False).head(20)
    print(loc_stat.to_string(index=False))

    print("\n▶ [3] 맵영역그룹(Map_Group)별 집계 TOP 10:")
    mg_stat = all_sp[all_sp['Map_Group'] != ''].groupby('Map_Group').agg(
        관측건수=('Count', 'count'),
        총개체수=('Count', 'sum')
    ).sort_values(by='관측건수', ascending=False).head(10)
    print(mg_stat.to_string())

    print("\n▶ [4] 연도별 8~9월 제비 출현 추이:")
    y_stat = all_sp.groupby('Year').agg(
        관측건수=('Count', 'count'),
        총개체수=('Count', 'sum')
    )
    print(y_stat.to_string())

# 3. 조류충돌 및 사체수거 통합 집계
strike_files = [f for f in os.listdir(resampled_dir) if 'birdstrike' in f and f.endswith('.csv')]
swallow_strikes = []
for sf in strike_files:
    s_path = os.path.join(resampled_dir, sf)
    try:
        sdf = pd.read_csv(s_path, low_memory=False)
        date_col = 'Date' if 'Date' in sdf.columns else ('일자' if '일자' in sdf.columns else None)
        sp_col = 'Species' if 'Species' in sdf.columns else ('조류종' if '조류종' in sdf.columns else ('보정_조류종' if '보정_조류종' in sdf.columns else None))
        if date_col and sp_col:
            sdf['Month'] = pd.to_datetime(sdf[date_col], errors='coerce').dt.month
            sub_s = sdf[sdf['Month'].isin([8, 9])].copy()
            is_swallow_s = sub_s[sp_col].astype(str).str.contains('제비|귀제비', na=False) & ~sub_s[sp_col].astype(str).str.contains('족제비', na=False)
            swallow_strikes.append(sub_s[is_swallow_s])
    except Exception as e:
        pass

if swallow_strikes:
    all_strikes = pd.concat(swallow_strikes, ignore_index=True)
    print(f"\n▶ [5] 8~9월 제비 조류충돌(Bird Strike) 다발 활주로 및 구역 (총 {len(all_strikes)}건):")
    rw_col = 'Runway_Raw' if 'Runway_Raw' in all_strikes.columns else ('Runway' if 'Runway' in all_strikes.columns else '활주로')
    if rw_col in all_strikes.columns:
        print(all_strikes[rw_col].value_counts().head(10).to_string())

carcass_files = [f for f in os.listdir(resampled_dir) if 'carcass' in f and f.endswith('.csv')]
swallow_carcass = []
for cf in carcass_files:
    c_path = os.path.join(resampled_dir, cf)
    try:
        cdf = pd.read_csv(c_path, low_memory=False)
        date_col = 'Date' if 'Date' in cdf.columns else ('일자' if '일자' in cdf.columns else None)
        sp_col = 'Species' if 'Species' in cdf.columns else ('조류종' if '조류종' in cdf.columns else ('보정_조류종' if '보정_조류종' in cdf.columns else None))
        if date_col and sp_col:
            cdf['Month'] = pd.to_datetime(cdf[date_col], errors='coerce').dt.month
            sub_c = cdf[cdf['Month'].isin([8, 9])].copy()
            is_swallow_c = sub_c[sp_col].astype(str).str.contains('제비|귀제비', na=False) & ~sub_c[sp_col].astype(str).str.contains('족제비', na=False)
            swallow_carcass.append(sub_c[is_swallow_c])
    except Exception as e:
        pass

if swallow_carcass:
    all_carcass = pd.concat(swallow_carcass, ignore_index=True)
    print(f"\n▶ [6] 8~9월 제비 사체수거(Carcass/충돌흔) 다발 장소 (총 {len(all_carcass)}건):")
    loc_col = 'Location_Raw' if 'Location_Raw' in all_carcass.columns else ('수거위치' if '수거위치' in all_carcass.columns else '위치')
    if loc_col in all_carcass.columns:
        print(all_carcass[loc_col].value_counts().head(12).to_string())
