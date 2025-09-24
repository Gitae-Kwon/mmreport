import pandas as pd
import streamlit as st

from compute import compute_report_frames
from formatter import to_formatted_excel_bytes, styled_preview_html

st.set_page_config(page_title="엑셀 → 양식 자동 변환기", layout="wide")
st.title("엑셀 → 회사 양식 자동 변환기")
st.caption("엑셀 업로드만 지원 (Google Sheets 없음)")

col1, col2 = st.columns([2, 1], gap="large")

with col1:
    st.subheader("1) 엑셀 업로드 및 탭(월) 선택")
    up = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="xlsx_up")
    df = None
    worksheet_name = None

    if up is not None:
        xl = pd.ExcelFile(up)
        sheet_tabs = xl.sheet_names
        worksheet_name = st.selectbox("월(탭) 선택", sheet_tabs)
        if worksheet_name:
            df = xl.parse(worksheet_name)
            st.success(f"탭 '{worksheet_name}' 읽기 완료 (행 {len(df)})")

    if df is not None:
        st.divider()
        st.subheader("2) 계산 및 양식 변환")

        with st.expander("원본 데이터 미리보기", expanded=False):
            st.dataframe(df.head(100), use_container_width=True)

        result = compute_report_frames(df)
        st.success("계산이 완료되었습니다.")

        html_preview = styled_preview_html(result)
        st.markdown("#### 미리보기(간단 스타일)")
        st.markdown(html_preview, unsafe_allow_html=True)

        # 다운로드 (포맷된 엑셀)
        bytes_xlsx = to_formatted_excel_bytes(result, month_label=worksheet_name or "선택월")
        st.download_button(
            label="포맷된 엑셀 다운로드 (.xlsx)",
            data=bytes_xlsx,
            file_name=f"MM_변환_{(worksheet_name or '선택월')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

with col2:
    st.subheader("옵션")
    st.caption("필요 시 확장 예정")
    st.checkbox("검증 모드(예외 상세 표시)", value=False, disabled=True)

st.divider()
st.markdown("""
### 사용법
1. 엑셀(.xlsx) 업로드 → 월(탭) 선택.
2. **계산 및 양식 변환** 섹션에서 결과 **미리보기** 확인 → **포맷된 엑셀 다운로드**.

> ⚠️ 실제 회사 양식과 계산 로직은 팀 규칙에 맞게 `compute.py`와 `formatter.py`를 수정하세요.
""")
