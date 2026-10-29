import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="재무팀 전표 & 부가세 통합 검토기", layout="wide")
st.title("📊 재무팀 ERP 전표 & 부가세 자동 검토 시스템")
st.markdown("매월 ERP에서 다운로드한 **분개장 엑셀 파일**을 업로드하시면 전표 규정 위반(작성자 포함)과 부가세 오류를 동시에 검출합니다.")

# 1. ERP 종류 선택
erp_type = st.radio("사용하시는 ERP 종류를 선택하세요:", ["영림원 ERP", "더존 ERP (Smart A / iCube)"], horizontal=True)

# 2. 파일 업로드 창
uploaded_file = st.file_uploader("전표 엑셀 파일 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])

# 작성자 추출 보조 함수
def extract_author(row, user_col, desc_val):
    if user_col and pd.notna(row.get(user_col)) and str(row.get(user_col)).strip() not in ['', 'nan', '-']:
        return str(row.get(user_col)).strip()
    if desc_val and '_' in desc_val:
        prefix = desc_val.split('_')[0].strip()
        # 한글 성명 (2~4글자) 패턴인 경우 작성자로 인식
        if re.match(r'^[가-힣]{2,4}$', prefix):
            return prefix
    return '-'

if uploaded_file is not None:
    with st.spinner('전표 및 부가세 데이터를 분석 중입니다...'):
        try:
            # 엑셀 헤더 탐색
            df_preview = pd.read_excel(uploaded_file, header=None)
            header_idx = 0
            for i in range(min(10, len(df_preview))):
                row_str = " ".join([str(val) for val in df_preview.iloc[i].tolist() if pd.notna(val)])
                if any(k in row_str for k in ['계정과목', '거래처명', '적요', '년/월/일']):
                    header_idx = i
                    break
            
            df_raw = pd.read_excel(uploaded_file, header=header_idx)
            
            # 컬럼 매핑
            date_col, slip_col, amt_col, acc_col, desc_col, cust_col, user_col = None, None, None, None, None, None, None
            for col in df_raw.columns:
                c_str = str(col).strip().replace(" ", "")
                if not date_col and any(k in c_str for k in ['일자', '년/월/일', '날짜']): date_col = col
                elif not slip_col and '번호' in c_str: slip_col = col
                elif not amt_col and any(k in c_str for k in ['차변금액', '금액', '차변']): amt_col = col
                elif not acc_col and any(k in c_str for k in ['차변계정', '계정과목', '계정명']): acc_col = col
                elif not desc_col and '적요' in c_str: desc_col = col
                elif not cust_col and ('거래처명' in c_str or '거래처' in c_str): cust_col = col
                elif not user_col and any(k in c_str for k in ['작성자', '기표자', '결의자', '사원명', '담당자', '입력자']): user_col = col

            clean_rows = []
            for _, r in df_raw.iterrows():
                amt = pd.to_numeric(r.get(amt_col), errors='coerce') if amt_col else 0
                if pd.notna(amt) and amt > 0:
                    desc_text = str(r.get(desc_col, '')) if pd.notna(r.get(desc_col)) else ''
                    author = extract_author(r, user_col, desc_text)
                    clean_rows.append({
                        '일자': r.get(date_col, ''),
                        '전표번호': r.get(slip_col, '-'),
                        '작성자': author,
                        '차변계정': str(r.get(acc_col, '')),
                        '금액': amt,
                        '적요': desc_text,
                        '거래처명': str(r.get(cust_col, '')) if pd.notna(r.get(cust_col)) else ''
                    })

            df = pd.DataFrame(clean_rows)
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
            df['요일'] = df['일자'].dt.dayofweek

            # ==========================================
            # 탭 구성: [전표 이상징후 검토] vs [부가세 대사 및 점검]
            # ==========================================
            tab1, tab2 = st.tabs(["📝 전표 이상징후 검토", "🔍 부가세 대사 및 불공제 점검"])

            # ------------------------------------------
            # TAB 1: 전표 이상징후 검토
            # ------------------------------------------
            with tab1:
                audit_results = []
                # 1. 주말 집행 의심건
                weekend_mask = (df['요일'] >= 5) & df['차변계정'].str.contains('회의비|접대비|복리후생비|여비교통비')
                for _, row in df[weekend_mask].iterrows():
                    audit_results.append({
                        '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                        '전표번호': row['전표번호'], '작성자': row['작성자'], '계정과목': row['차변계정'], 
                        '금액': f"{row['금액']:,.0f}", '거래처명': row['거래처명'], '적요': row['적요'], 
                        '위험도': 'HIGH', '검토의견': '주말/공휴일 복리후생 및 접대성 경비 집행 (사적 유용 의심)'
                    })

                # 2. 회의비/접대비 규정 위반 (10만원 기준)
                over_100k = df[df['차변계정'].str.contains('회의비') & (df['금액'] > 100000)]
                for _, row in over_100k.iterrows():
                    audit_results.append({
                        '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                        '전표번호': row['전표번호'], '작성자': row['작성자'], '계정과목': row['차변계정'], 
                        '금액': f"{row['금액']:,.0f}", '거래처명': row['거래처명'], '적요': row['적요'], 
                        '위험도': 'HIGH', '검토의견': '1회 10만원 초과 건으로 회의비가 아닌 접대비 계정 처리 필요'
                    })

                under_100k = df[df['차변계정'].str.contains('접대비') & (df['금액'] <= 100000)]
                for _, row in under_100k.iterrows():
                    audit_results.append({
                        '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                        '전표번호': row['전표번호'], '작성자': row['작성자'], '계정과목': row['차변계정'], 
                        '금액': f"{row['금액']:,.0f}", '거래처명': row['거래처명'], '적요': row['적요'], 
                        '위험도': 'MEDIUM', '검토의견': '10만원 이하 건으로 한도 시부인 유리한 회의비 계정 원복 권장'
                    })

                # 3. 고액 소모품비 자산화 누락 의심 (100만원 이상)
                asset_sus = df[df['차변계정'].str.contains('소모품비') & (df['금액'] >= 1000000)]
                for _, row in asset_sus.iterrows():
                    audit_results.append({
                        '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                        '전표번호': row['전표번호'], '작성자': row['작성자'], '계정과목': row['차변계정'], 
                        '금액': f"{row['금액']:,.0f}", '거래처명': row['거래처명'], '적요': row['적요'], 
                        '위험도': 'HIGH', '검토의견': '100만원 이상 고액 소모품비. 자산(기계장치/비품) 등록 누락 의심'
                    })

                # 4. 동일 일자, 동일 거래처 분할(쪼개기) 결제
                grouped = df.groupby(['일자', '거래처명', '차변계정']).agg(
                    결제건수=('금액', 'count'),
                    작성자목록=('작성자', lambda x: ", ".join(sorted(list(set([v for v in x if v != '-'])))) or '-')
                ).reset_index()
                split_sus = grouped[(grouped['결제건수'] >= 2) & (grouped['거래처명'] != '') & (grouped['거래처명'] != 'nan')]
                for _, split in split_sus.iterrows():
                    if any(k in split['차변계정'] for k in ['소모품비', '회의비', '접대비']):
                        audit_results.append({
                            '일자': split['일자'].strftime('%Y-%m-%d') if pd.notna(split['일자']) else '-',
                            '전표번호': '-', '작성자': split['작성자목록'], '계정과목': split['차변계정'], 
                            '금액': '-', '거래처명': split['거래처명'], '적요': f"동일 일자 {split['결제건수']}건 중복 결제",
                            '위험도': 'HIGH', '검토의견': '한도 회피 또는 자산화 회피 목적의 분할(쪼개기) 결제 의심'
                        })

                st.subheader(f"전표 규정 위반 검토 결과 ({len(audit_results)}건)")
                if len(audit_results) > 0:
                    res_df1 = pd.DataFrame(audit_results)
                    st.dataframe(res_df1, use_container_width=True)
                else:
                    st.info("전표 규정 위반 의심 건이 없습니다.")

            # ------------------------------------------
            # TAB 2: 부가세 대사 및 불공제 점검
            # ------------------------------------------
            with tab2:
                vat_results = []
                vat_df = df[df['차변계정'].str.contains('부가세대급금|부가세')]
                
                # 1. 비영업용 소형승용차 관련 불공제 의심건
                car_keywords = ['G80', '아반떼', '쏘렌토', '그랜저', '제네시스', 'K5', 'K7', 'K8', '소나타', '개인차량', '주차비', '주유비']
                for _, row in vat_df.iterrows():
                    desc = row['적요']
                    is_allowed_van = ('스타리아' in desc and '공무' in desc) or ('3.5톤' in desc) or ('화물' in desc)
                    if not is_allowed_van and any(k in desc for k in car_keywords):
                        vat_results.append({
                            '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                            '전표번호': row['전표번호'], '작성자': row['작성자'], '부가세금액': f"{row['금액']:,.0f}",
                            '거래처명': row['거래처명'], '적요': row['적요'],
                            '점검항목': '차량 부가세 불공제 대상 매입세액 공제 의심',
                            '조치사항': '비영업용 소형승용차 관련 비용 여부 확인 후 불공제 처리 필요'
                        })

                # 2. 병원, 약국, 면세 관련 거래처 부가세대급금 계상 오류
                tax_free_keywords = ['약국', '병원', '의원', '의료재단', '면세점']
                for _, row in vat_df.iterrows():
                    if any(k in row['거래처명'] or k in row['적요'] for k in tax_free_keywords):
                        vat_results.append({
                            '일자': row['일자'].strftime('%Y-%m-%d') if pd.notna(row['일자']) else '-',
                            '전표번호': row['전표번호'], '작성자': row['작성자'], '부가세금액': f"{row['금액']:,.0f}",
                            '거래처명': row['거래처명'], '적요': row['적요'],
                            '점검항목': '면세/의료비/면세점 부가세대급금 기표 의심',
                            '조치사항': '면세 거래처에 대한 부가세 대급금 계상 오류 확인'
                        })

                st.subheader(f"부가세 매입세액 불공제 및 오류 점검 결과 ({len(vat_results)}건)")
                if len(vat_results) > 0:
                    res_df2 = pd.DataFrame(vat_results)
                    st.dataframe(res_df2, use_container_width=True)
                else:
                    st.info("부가세 계정 관련 특이사항이 발견되지 않았습니다.")

            # ------------------------------------------
            # 통합 다운로드 (하나의 엑셀에 시트 2개로 저장)
            # ------------------------------------------
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                if len(audit_results) > 0:
                    pd.DataFrame(audit_results).to_excel(writer, sheet_name='전표_이상징후', index=False)
                if len(vat_results) > 0:
                    pd.DataFrame(vat_results).to_excel(writer, sheet_name='부가세_불공제의심', index=False)
            
            st.divider()
            st.download_button(
                label="📥 전표 & 부가세 검토 결과 통합 엑셀 다운로드 (작성자 포함)",
                data=buffer.getvalue(),
                file_name="전표_부가세_통합검토결과.xlsx",
                mime="application/vnd.ms-excel"
            )

        except Exception as e:
            st.error(f"파일을 분석하는 중 오류가 발생했습니다. (상세 에러: {e})")
