import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="재무팀 전표 검토기", layout="wide")
st.title("📊 재무팀 영림원 ERP 전표 자동 검토 시스템")

uploaded_file = st.file_uploader("이번 달 전표 엑셀 파일 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])

if uploaded_file is not None:
    with st.spinner('전표 데이터를 분석 중입니다...'):
        try:
            # 영림원 양식에 맞춰 첫 번째 행을 컬럼명으로 인식하여 읽기
            df = pd.read_excel(uploaded_file, header=1)
            df.columns = ['일자', '전표번호', '차변금액', '차변계정', '대변계정', '대변금액', '적요', '거래처코드', '거래처명', '결의일자', '결의번호', 'PJT코드', 'PJT명']
            
            df['차변금액'] = pd.to_numeric(df['차변금액'], errors='coerce').fillna(0)
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce')
            df['차변계정'] = df['차변계정'].fillna('')
            df['적요'] = df['적요'].fillna('')
            df['거래처명'] = df['거래처명'].fillna('')
            
            audit_results = []

            # 1. 주말/공휴일 집행 의심건
            df['요일'] = df['일자'].dt.dayofweek
            weekend_mask = (df['요일'] >= 5) & df['차변계정'].str.contains('회의비|접대비|복리후생비|여비교통비') & (df['차변금액'] > 0)
            for _, row in df[weekend_mask].iterrows():
                audit_results.append({'일자': row['일자'].strftime('%Y-%m-%d'), '계정과목': row['차변계정'], '금액': row['차변금액'], '거래처명': row['거래처명'], '위험도': 'HIGH', '검토의견': '주말/공휴일 복리후생 및 접대성 경비 집행'})

            # 2. 회의비/접대비 규정 위반
            over_100k = df[df['차변계정'].str.contains('회의비') & (df['차변금액'] > 100000)]
            for _, row in over_100k.iterrows():
                audit_results.append({'일자': row['일자'].strftime('%Y-%m-%d'), '계정과목': row['차변계정'], '금액': row['차변금액'], '거래처명': row['거래처명'], '위험도': 'HIGH', '검토의견': '10만원 초과로 접대비 처리 필요'})
            
            # 3. 자산화 누락 의심
            asset_sus = df[df['차변계정'].str.contains('소모품비') & (df['차변금액'] >= 1000000)]
            for _, row in asset_sus.iterrows():
                audit_results.append({'일자': row['일자'].strftime('%Y-%m-%d'), '계정과목': row['차변계정'], '금액': row['차변금액'], '거래처명': row['거래처명'], '위험도': 'HIGH', '검토의견': '100만원 이상 소모품비. 자산 등록 누락 의심'})

            # 4. 쪼개기 결제 의심
            grouped = df[df['차변금액'] > 0].groupby(['일자', '거래처명', '차변계정']).size().reset_index(name='결제건수')
            split_sus = grouped[(grouped['결제건수'] >= 2) & (grouped['거래처명'] != '')]
            for _, split in split_sus.iterrows():
                if any(x in split['차변계정'] for x in ['소모품비', '회의비', '접대비']):
                    audit_results.append({'일자': split['일자'].strftime('%Y-%m-%d'), '계정과목': split['차변계정'], '금액': '-', '거래처명': split['거래처명'], '위험도': 'HIGH', '검토의견': f"동일 일자 {split['결제건수']}건 중복 결제 (분할 의심)"})

            st.success(f"검토 완료! 총 {len(df)}건 중 {len(audit_results)}건의 이상 징후 발견.")
            
            if len(audit_results) > 0:
                result_df = pd.DataFrame(audit_results)
                st.dataframe(result_df, use_container_width=True)
                
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, sheet_name='이상징후', index=False)
                
                st.download_button(label="📥 이상 징후 리포트 엑셀 다운로드", data=buffer.getvalue(), file_name="전표검토결과.xlsx", mime="application/vnd.ms-excel")
            else:
                st.info("규정 위반 의심 건이 없습니다.")
        except Exception as e:
            st.error("파일을 읽는 중 오류가 발생했습니다. 영림원 분개장 원본 엑셀인지 확인해주세요.")
