import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="요양원 청구서 자동화", layout="wide")

st.title("💊 요양원 청구서 자동화 시스템")

# 1. 파일 업로드
col1, col2 = st.columns(2)
with col1:
    file_info = st.file_uploader("요양원기본테이블.xlsx", type="xlsx", key="info")
with col2:
    file_data = st.file_uploader("처방데이터.xlsx", type="xlsx", key="data")

# 2. 처리 시작
if file_info and file_data:
    info_df = pd.read_excel(file_info)
    data_df = pd.read_excel(file_data)

    # 주민번호 앞6자리 추출
    data_df["주민번호앞6"] = data_df["주민등록번호"].astype(str).str[:6]
    info_df["주민번호앞6"] = info_df["주민등록번호"].astype(str).str[:6]

    # 요양원명 병합
    merged = pd.merge(data_df, info_df[["고객이름", "주민번호앞6", "요양원명"]], 
                      left_on=["고객이름", "주민번호앞6"],
                      right_on=["고객이름", "주민번호앞6"],
                      how="left")

    # 매칭되지 않은 행
    unmatched = merged[merged["요양원명"].isna()]
    if not unmatched.empty:
        st.warning("❗ 매칭되지 않은 환자가 있습니다. 수기로 요양원을 입력하세요.")
        for i, row in unmatched.iterrows():
            manual = st.text_input(f"{row['고객이름']} ({row['주민등록번호']}) 요양원명 입력", key=i)
            if manual:
                merged.at[i, "요양원명"] = manual

    # 매칭 완료 후 처리
    if merged["요양원명"].isna().sum() == 0:
        요양원목록 = merged["요양원명"].unique()
        선택된요양원 = st.selectbox("요양원 선택", 요양원목록)
        보기형식 = st.radio("형식 선택", ["기본형", "피벗형"])

        target = merged[merged["요양원명"] == 선택된요양원].copy()

        # 날짜 열 처리
        target["일자"] = pd.to_datetime(target["내방일"], errors="coerce").dt.day

        if 보기형식 == "기본형":
            st.dataframe(target[["고객이름", "주민등록번호", "요양급여액", "비급여액", "계", "내방일"]])
        else:
            pivot = target.pivot_table(index="고객이름", 
                                       columns="일자", 
                                       values="계", 
                                       aggfunc="sum", 
                                       fill_value=0)
            pivot["합계"] = pivot.sum(axis=1)
            st.dataframe(pivot)

        # 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if 보기형식 == "기본형":
                target.to_excel(writer, index=False, sheet_name="청구서")
            else:
                pivot.to_excel(writer, sheet_name="청구서")

        st.download_button(
            label="💾 청구서 Excel 다운로드",
            data=buffer.getvalue(),
            file_name=f"{선택된요양원}_청구서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
