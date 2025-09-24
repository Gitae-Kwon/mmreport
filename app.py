
import io
import re
import time
import pandas as pd
import streamlit as st

from compute import list_month_tabs, compute_report_frames
from formatter import to_formatted_excel_bytes, styled_preview_html

try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

st.set_page_config(page_title="시트 → 양식 자동 변환기", layout="wide")

st.title("구글시트/엑셀 → 회사 양식 자동 변환기")
st.caption("Streamlit + Python")

tab1, tab2 = st.tabs(["데이터 불러오기", "도움말"])

with tab1:
    col1, col2 = st.columns([2,1], gap="large")
    with col1:
        st.subheader("1) 데이터 선택")
        mode = st.radio("불러오기 방법을 고르세요", ["Google Sheets", "로컬 엑셀 업로드"], horizontal=True)

        df = None
        sheet_tabs = []

        if mode == "Google Sheets":
            st.markdown("#### Google Sheets 설정")
            st.write("서비스 계정(JSON)과 **스프레드시트 ID**(또는 URL)를 입력해 주세요.")
            gs_json = st.file_uploader("Google Service Account JSON", type=["json"], key="gsjson")
            sheet_id_or_url = st.text_input("스프레드시트 ID 또는 URL", placeholder="https://docs.google.com/spreadsheets/d/... 또는 1AbCdEf...")
            worksheet_name = None

            if st.button("탭(월) 불러오기", type="primary", use_container_width=True, disabled=(gs_json is None or not sheet_id_or_url)):
                try:
                    if gspread is None:
                        st.error("gspread/Google 라이브러리가 설치되지 않았습니다. requirements.txt를 참고해 주세요.")
                    else:
                        creds = Credentials.from_service_account_info(
                            pd.read_json(gs_json).to_dict()
                        )
                        gc = gspread.authorize(creds)
                        # 지원: URL 전체 또는 ID
                        m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_id_or_url)
                        spreadsheet_key = m.group(1) if m else sheet_id_or_url.strip()
                        sh = gc.open_by_key(spreadsheet_key)
                        sheet_tabs = [ws.title for ws in sh.worksheets()]
                        st.success(f"탭 발견: {', '.join(sheet_tabs)}")
                        st.session_state["gs_spreadsheet_key"] = spreadsheet_key
                        st.session_state["gs_json_dict"] = pd.read_json(gs_json).to_dict()
                        st.session_state["sheet_tabs"] = sheet_tabs
                except Exception as e:
                    st.exception(e)

            if "sheet_tabs" in st.session_state and st.session_state.get("sheet_tabs"):
                worksheet_name = st.selectbox("월(탭) 선택", st.session_state["sheet_tabs"])
                if st.button("선택 탭 읽기", use_container_width=True):
                    try:
                        creds = Credentials.from_service_account_info(st.session_state["gs_json_dict"])
                        gc = gspread.authorize(creds)
                        sh = gc.open_by_key(st.session_state["gs_spreadsheet_key"])
                        ws = sh.worksheet(worksheet_name)
                        data = ws.get_all_values()
                        df = pd.DataFrame(data[1:], columns=data[0])
                        st.success(f"탭 '{worksheet_name}' 읽기 완료 (행 {len(df)}).")
                    except Exception as e:
                        st.exception(e)

        else:
            st.markdown("#### 엑셀 업로드")
            up = st.file_uploader("엑셀 파일(.xlsx) 업로드", type=["xlsx"], key="xlsx_up")
            if up is not None:
                xl = pd.ExcelFile(up)
                sheet_tabs = xl.sheet_names
                worksheet_name = st.selectbox("월(탭) 선택", sheet_tabs)
                if worksheet_name:
                    df = xl.parse(worksheet_name)
                    st.success(f"탭 '{worksheet_name}' 읽기 완료 (행 {len(df)}).")

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
        st.subheader("3) 옵션")
        st.caption("필요 시 확장 예정")
        st.checkbox("검증 모드(예외 상세 표시)", value=False, disabled=True)

with tab2:
    st.markdown("""
### 사용법
1. **Google Sheets** 모드: 서비스 계정 JSON과 문서 ID/URL을 입력 → *탭 불러오기* → *월(탭) 선택* → *선택 탭 읽기*.
2. **엑셀 업로드** 모드: xlsx 업로드 → 월(탭) 선택.
3. **계산 및 양식 변환** 섹션에서 결과 **미리보기** 확인 → **포맷된 엑셀 다운로드**.

> ⚠️ 실제 회사 양식과 계산 로직은 팀 규칙에 맞게 `compute.py`와 `formatter.py`를 수정하세요.
""")
