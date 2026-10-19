import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

st.set_page_config(page_title="뽀득 초과근무 검증 시스템", page_icon="🐳", layout="wide")

st.title("🐳 뽀득 초과근무 자동 검증 시스템")
st.markdown("샤플(Shopl)과 BOS 시스템의 초과근무 내역을 단 1초 만에 비교합니다.")

with st.sidebar:
    st.header("📂 1. 엑셀 파일 업로드")
    bos_file = st.file_uploader("BOS 초과근무 파일 업로드", type=['xlsx'])
    shopl_file = st.file_uploader("샤플 근태 리포트 업로드", type=['xlsx'])
    
    st.markdown("---")
    st.header("📅 2. 공휴일 직접 입력")
    st.markdown("이번 달의 **평일 중 쉬는 날(법정공휴일)**만 쉼표(,)로 구분해서 적어주세요.")
    holidays_input = st.text_input("입력 예시", placeholder="2026-09-24, 2026-09-25")

if st.button("🚀 검증 시작하기", type="primary"):
    if not bos_file or not shopl_file:
        st.warning("👈 왼쪽 메뉴에서 두 개의 엑셀 파일을 모두 업로드해주세요!")
    else:
        with st.spinner('데이터를 비교하는 중입니다...'):
            try:
                bos_df = pd.read_excel(bos_file)
                shopl_df = pd.read_excel(shopl_file, header=3)
                
                kr_holidays = []
                if holidays_input.strip():
                    dates = holidays_input.replace(" ", "").split(",")
                    for d in dates:
                        kr_holidays.append(datetime.strptime(d.strip(), "%Y-%m-%d").date())
                
                bos_df = bos_df[bos_df['소속'] != '도급'].copy()
                bos_df['날짜'] = pd.to_datetime(bos_df['날짜'])
                bos_df['is_휴일'] = bos_df['날짜'].dt.dayofweek.isin([5, 6]) | bos_df['날짜'].dt.date.isin(kr_holidays)
                bos_df['is_야간'] = bos_df['주야'] == '야간'
                
                bos_df['BOS_연장'] = np.where(~bos_df['is_휴일'], bos_df['OT시간'], 0)
                bos_df['BOS_휴일총합'] = np.where(bos_df['is_휴일'], bos_df['OT시간'], 0)
                bos_df['BOS_야간'] = np.where(bos_df['is_야간'], bos_df['OT시간'], 0)
                
                bos_agg = bos_df.groupby('직원명')[['BOS_연장', 'BOS_휴일총합', 'BOS_야간']].sum().reset_index()
                bos_agg.rename(columns={'직원명': '이름'}, inplace=True)
                
                cols_to_use = ['이름', '연장 근로시간 -\n배정 (시간)', '야간 근로시간 -\n배정 (시간)', '휴일 근로시간 -\n배정 (시간)', '휴일연장 근로시간 -\n배정 (시간)']
                shopl_filtered = shopl_df[cols_to_use].copy()
                shopl_filtered.columns = ['이름', '샤플_연장', '샤플_야간', '샤플_휴일기본', '샤플_휴일연장']
                
                for col in ['샤플_연장', '샤플_야간', '샤플_휴일기본', '샤플_휴일연장']:
                    shopl_filtered[col] = pd.to_numeric(shopl_filtered[col], errors='coerce').fillna(0)
                shopl_filtered['샤플_휴일총합'] = shopl_filtered['샤플_휴일기본'] + shopl_filtered['샤플_휴일연장']
                
                merged_df = pd.merge(shopl_filtered, bos_agg, on='이름', how='outer').fillna(0)
                merged_df['Diff_연장오차'] = merged_df['샤플_연장'] - merged_df['BOS_연장']
                merged_df['Diff_야간오차'] = merged_df['샤플_야간'] - merged_df['BOS_야간']
                merged_df['Diff_휴일오차'] = merged_df['샤플_휴일총합'] - merged_df['BOS_휴일총합']
                
                merged_df['확인필요'] = (merged_df['Diff_연장오차'] != 0) | (merged_df['Diff_야간오차'] != 0) | (merged_df['Diff_휴일오차'] != 0)
                
                result_df = merged_df[merged_df['확인필요']].sort_values('이름')
                
                st.markdown("---")
                if result_df.empty:
                    st.success("🎉 완벽합니다! 모든 직원의 초과근무 내역이 일치합니다.")
                else:
                    st.error(f"🚨 총 {len(result_df)}명의 불일치 내역이 발견되었습니다.")
                    display_cols = ['이름', 'Diff_연장오차', 'Diff_야간오차', 'Diff_휴일오차', '샤플_연장', 'BOS_연장', '샤플_휴일총합', 'BOS_휴일총합']
                    st.dataframe(result_df[display_cols], use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='확인필요(오차발생)')
                        merged_df.sort_values(by=['확인필요', '이름'], ascending=[False, True]).to_excel(writer, index=False, sheet_name='전체데이터')
                    output.seek(0)
                    
                    target_month = bos_df['날짜'].dt.month.iloc[0]
                    st.download_button(
                        label="📥 [클릭] 검증 결과 엑셀 다운로드",
                        data=output,
                        file_name=f"초과근무_자동검증리포트_{target_month}월.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"오류가 발생했습니다. 파일 양식을 확인해주세요. (에러내용: {e})")