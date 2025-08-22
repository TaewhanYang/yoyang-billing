// 업데이트된 청구서 자동화 시스템 with 개선사항 반영
// 주요 기능:
// - 내방일 날짜 형식 정제
// - 불필요한 컬럼 삭제
// - 동명이인 시 주민번호 앞 두자리 표기
// - 기본형 및 피벗형 청구서 구성

import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="요양원 청구서 자동화", layout="wide")
st.title("💊 요양원 청구서 자동화 시스템")

# 파일 업로드
col1, col2 = st.columns(2)
with col1:
    file_info = st.file_uploader("📁 요양원기본테이블.xlsx", type="xlsx", key="info")
with col2:
    file_data = st.file_uploader("📁 처방데이터.xlsx", type="xlsx", key="data")

if file_info and file_data:
    info_df = pd.read_excel(file_info)
    data_df = pd.read_excel(file_data)

    # 매칭키 생성
    data_df["고객이름"] = data_df["고객이름"].astype(str).str.strip()
    data_df["주민등록번호"] = data_df["주민등록번호"].astype(str)
    data_df["매칭키"] = data_df["고객이름"] + data_df["주민등록번호"].str[:6]

    info_df["고객이름"] = info_df["고객이름"].astype(str).str.strip()
    info_df["주민등록번호"] = info_df["주민등록번호"].astype(str)
    info_df["매칭키"] = info_df["고객이름"] + info_df["주민등록번호"].str[:6]

    merged = pd.merge(data_df, info_df[["매칭키", "요양원명"]], on="매칭키", how="left")
    unmatched = merged[merged["요양원명"].isna()].copy()

    if not unmatched.empty:
        st.warning(f"⚠️ 매칭되지 않은 환자가 {len(unmatched)}명 있습니다.")
        do_manual = st.radio("수기로 요양원명을 입력하시겠습니까?", ["아니요", "예"], index=0)

        if do_manual == "예":
            new_entries = []
            st.markdown("### ✏️ 수기 요양원 입력")
            for idx, row in unmatched.iterrows():
                col1, col2 = st.columns([2, 3])
                with col1:
                    st.text(row["고객이름"] + " / " + row["주민등록번호"])
                with col2:
                    val = st.text_input("요양원명 입력", key=f"manual_{idx}")
                    if val:
                        merged.at[idx, "요양원명"] = val
                        new_entries.append({
                            "고객이름": row["고객이름"],
                            "주민등록번호": row["주민등록번호"],
                            "요양원명": val
                        })
            if new_entries:
                new_df = pd.DataFrame(new_entries)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    new_df.to_excel(writer, index=False)
                date_str = datetime.now().strftime("%Y%m%d")
                st.download_button(
                    label="💾 신규 요양원 테이블 다운로드",
                    data=buffer.getvalue(),
                    file_name=f"신규요양원등록_{date_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.stop()

    if merged["요양원명"].isna().sum() == 0:
        st.success("✅ 모든 환자가 요양원과 성공적으로 매칭되었습니다.")

        # 불필요한 열 제거
        drop_cols = [col for col in merged.columns if col.lower().startswith("unnamed") or col in ["보험", "매칭키", "요양원명", "일자"]]
        merged.drop(columns=drop_cols, inplace=True, errors="ignore")

        # 내방일 포맷 변경 → "6월2일"
        merged["날짜"] = pd.to_datetime(merged["내방일"], errors="coerce").dt.strftime("%m월%-d일")

        # 동명이인 구분용 이름 변경
        merged["이름그룹"] = merged["고객이름"] + "_" + merged["주민등록번호"].str[:2]
        name_counts = merged["고객이름"].value_counts()
        merged["표시이름"] = merged.apply(
            lambda row: f"{row['고객이름']}({row['주민등록번호'][:2]})" if name_counts[row['고객이름']] > 1 else row['고객이름'],
            axis=1
        )

        요양원목록 = merged["요양원명"].unique()
        선택된요양원 = st.selectbox("📌 요양원 선택", 요양원목록)
        보기형식 = st.radio("📄 보기 형식", ["기본형", "피벗형"])

        target = merged[merged["요양원명"] == 선택된요양원].copy()

        if 보기형식 == "기본형":
            st.dataframe(target[["날짜", "표시이름", "주민등록번호", "계", "요양급여액", "비급여액"]])
        else:
            target["일"] = pd.to_datetime(target["날짜"], errors="coerce").dt.day
            pivot = target.pivot_table(index="표시이름", columns="일", values="계", aggfunc="sum", fill_value=0)
            pivot["합계"] = pivot.sum(axis=1)
            st.dataframe(pivot)

        # Excel 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if 보기형식 == "기본형":
                target[["날짜", "표시이름", "주민등록번호", "계", "요양급여액", "비급여액"]].to_excel(writer, index=False, sheet_name="청구서")
            else:
                pivot.to_excel(writer, sheet_name="청구서")

        st.download_button(
            label="💾 청구서 Excel 다운로드",
            data=buffer.getvalue(),
            file_name=f"{선택된요양원}_청구서.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
