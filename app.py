import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="재무팀 전표 검토기", layout="wide")
st.title("📊 재무팀 ERP 전표 자동 검토 시스템")
st.markdown("매월 ERP에서 다운로드한 **분개장 엑셀 파일**을 업로드하시면, 사내 규정에 따른 이상 전표를 자동 검출합니다.")

# 1. ERP 종류 선택
erp_type = st.radio("사용하시는 ERP 종류를 선택하세요:", ["영림원 ERP", "더존 ERP (Smart A / iCube)"], horizontal=True)

# 2. 파일 업로드 창
uploaded_file = st.file_uploader("전표 엑셀 파일 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner('전표 데이터를 분석 중입니다...'):
        try:
            # ------------------------------------------------------------------
            # ERP별 데이터 로드 및 표준 규격 통일
            # 표준 컬럼: ['일자', '전표번호', '차변계정', '금액', '적요', '거래처명']
            # ------------------------------------------------------------------
            clean_rows = []

            if erp_type == "영림원 ERP":
                # 영림원: 2번째 행이 헤더
                df_raw = pd.read_excel(uploaded_file, header=1)
                df_raw.columns = ['일자', '전표번호', '차변금액', '차변계정', '대변계정', '대변금액', '적요', '거래처코드', '거래처명', '결의일자', '결의번호', 'PJT코드', 'PJT명']
                
                for _, r in df_raw.iterrows():
                    amt = pd.to_numeric(r['차변금액'], errors='coerce')
                    if pd.notna(amt) and amt > 0:
                        clean_rows.append({
                            '일자': r['일자'],
                            '전표번호': r['전표번호'],
                            '차변계정': str(r['차변계정']),
                            '금액': amt,
                            '적요': str(r['적요']) if pd.notna(r['적요']) else '',
                            '거래처명': str(r['거래처명']) if pd.notna(r['거래처명']) else ''
                        })

            else:
                # 더존 ERP: 상단 헤더 위치 자동 탐색 및 컬럼 매핑
                # 더존은 보통 헤더 행에 '계정과목' 또는 '계정명'이 들어있음
                df_preview = pd.read_excel(uploaded_file, header=None)
                header_idx = 0
                for i in range(min(10, len(df_preview))):
                    row_str = " ".join(df_preview.iloc[i].astype(str).tolist())
                    if any(k in row_str for k in ['계정과목', '계정명', '차변', '적요']):
                        header_idx = i
                        break
                
                df_raw = pd.read_excel(uploaded_file, header=header_idx)
                
                # 컬럼 매핑 (더존의 다양한 표기 흡수)
                col_map = {}
                for c in df_raw.columns:
                    c_clean = str(c).strip().replace(" ", "")
                    if '일자' in c_clean or '날짜' in c_clean: col_map['일자'] = c
                    elif '번호' in c_clean: col_map['전표번호'] = c
                    elif '계정' in c_clean and '명' in c_clean: col_map['차변계정'] = c
                    elif '계정과목' in c_clean: col_map['차변계정'] = c
                    elif '차변' in c_clean or '차변금액' in c_clean: col_map['금액'] = c
                    elif '적요' in c_clean: col_map['적요'] = c
                    elif '거래처' in c_clean and '명' in c_clean: col_map['거래처명'] = c
                    elif '거래처' in c_clean: col_map['거래처명'] = c

                for _, r in df_raw.iterrows():
                    amt = pd.to_numeric(r.get(col_map.get('금액')), errors='coerce')
                    if pd.notna(amt) and amt > 0:
                        clean_rows.append({
                            '일자': r.get(col_map.get('일자'), ''),
                            '전표번호': r.get(col_map.get('전표번호'), '-'),
                            '차변계정': str(r.get(col_map.get('차변계정'), '')),
                            '금액': amt,
                            '적요': str(r.get(col_map.get('적요'), '')),
                            '거래처명': str(r.get(col_map.get('거래처명'), ''))
                        })

            df = pd.DataFrame(clean_rows)
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
            df['요일'] = df['일자'].dt.dayofweek

            # ------------------------------------------------------------------
            # 공통 감사 규칙 엔진 (규정 검증)
            # ------------------------------------------------------------------
            audit_results = []

            # 1. 주말 집행 의심건 (회의비, 접대비, 복리후생비, 여비교통비)
            weekend_mask = (df['요일'] >= 5) & df['차변계정'].str.contains('회의비|접대비|복리후생비|여비교통비')
            for _, row in df[weekend_mask].iterrows():
                audit_results.append({
                    '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                    '전표번호': row['전표번호'],
                    '계정과목': row['차변계정'],
                    '금액': f"{row['금액']:,.0f}",
                    '거래처명': row['거래처명'],
                    '적요': row['적요'],
                    '위험도': 'HIGH',
                    '검토의견': '주말/공휴일 복리후생 및 접대성 경비 집행 (사적 유용 의심)'
                })

            # 2. 회의비/접대비 규정 위반 (10만원 기준)
            over_100k = df[df['차변계정'].str.contains('회의비') & (df['금액'] > 100000)]
            for _, row in over_100k.iterrows():
                audit_results.append({
                    '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                    '전표번호': row['전표번호'],
                    '계정과목': row['차변계정'],
                    '금액': f"{row['금액']:,.0f}",
                    '거래처명': row['거래처명'],
                    '적요': row['적요'],
                    '위험도': 'HIGH',
                    '검토의견': '10만원 초과 건으로 회의비가 아닌 접대비 계정 처리 필요'
                })

            under_100k = df[df['차변계정'].str.contains('접대비') & (df['금액'] <= 100000)]
            for _, row in under_100k.iterrows():
                audit_results.append({
                    '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                    '전표번호': row['전표번호'],
                    '계정과목': row['차변계정'],
                    '금액': f"{row['금액']:,.0f}",
                    '거래처명': row['거래처명'],
                    '적요': row['적요'],
                    '위험도': 'MEDIUM',
                    '검토의견': '10만원 이하 건으로 한도 시부인 유리한 회의비 계정 원복 권장'
                })

            # 3. 고액 소모품비 자산화 누락 의심 (100만원 이상)
            asset_sus = df[df['차변계정'].str.contains('소모품비') & (df['금액'] >= 1000000)]
            for _, row in asset_sus.iterrows():
                audit_results.append({
                    '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                    '전표번호': row['전표번호'],
                    '계정과목': row['차변계정'],
                    '금액': f"{row['금액']:,.0f}",
                    '거래처명': row['거래처명'],
                    '적요': row['적요'],
                    '위험도': 'HIGH',
                    '검토의견': '100만원 이상 고액 소모품비. 자산(기계장치/비품) 등록 누락 의심'
                })

            # 4. 동일 일자, 동일 거래처 분할(쪼개기) 결제 탐지
            grouped = df.groupby(['일자', '거래처명', '차변계정']).size().reset_index(name='결제건수')
            split_sus = grouped[(grouped['결제건수'] >= 2) & (grouped['거래처명'] != '') & (grouped['거래처명'] != 'nan')]
            for _, split in split_sus.iterrows():
                if any(k in split['차변계정'] for k in ['소모품비', '회의비', '접대비']):
                    audit_results.append({
                        '일자': split['일자'].strftime('%Y-%m-%d') if pd.notna(split['일자']) else '-',
                        '전표번호': '-',
                        '계정과목': split['차변계정'],
                        '금액': '-',
                        '거래처명': split['거래처명'],
                        '적요': f"동일 일자 {split['결제건수']}건 중복 결제",
                        '위험도': 'HIGH',
                        '검토의견': '한도 회피 또는 자산화 회피 목적의 분할(쪼개기) 결제 의심'
                    })

            # 결과 화면 출력
            st.success(f"검토 완료! 총 {len(df)}건 중 {len(audit_results)}건의 이상 징후가 발견되었습니다.")

            if len(audit_results) > 0:
                result_df = pd.DataFrame(audit_results)
                st.dataframe(result_df, use_container_width=True)

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, sheet_name='이상징후_리포트', index=False)

                st.download_button(
                    label="📥 이상 징후 리포트 엑셀 다운로드",
                    data=buffer.getvalue(),
                    file_name=f"{erp_type[:3]}_전표검토결과.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.info("선택하신 ERP 전표에서 규정 위반 의심 건이 발견되지 않았습니다.")

        except Exception as e:
            st.error(f"파일을 분석하는 중 오류가 발생했습니다. 양식이 올바른지 확인해주세요. (상세 에러: {e})")
