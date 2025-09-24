import io
import pandas as pd

def _hide_idx(styler: "pd.io.formats.style.Styler"):
    # pandas 버전 호환: hide_index() 우선, 없으면 hide(axis="index")
    if hasattr(styler, "hide_index"):
        return styler.hide_index()
    try:
        return styler.hide(axis="index")
    except Exception:
        return styler

def styled_preview_html(df: pd.DataFrame) -> str:
    """Streamlit에서 표시할 간단 HTML 미리보기."""
    if isinstance(df.columns, pd.MultiIndex):
        styler = (
            df.style
            .format(precision=1)
            .set_table_attributes('class="table"')
            .set_table_styles([
                {"selector": "th", "props": [("text-align", "center"), ("background-color", "#f6f8fa")]},
                {"selector": "td", "props": [("text-align", "right")]},
            ])
        )
        styler = _hide_idx(styler)
    else:
        styler = df.style.format(precision=1)
        styler = _hide_idx(styler)

    return styler.to_html()

def to_formatted_excel_bytes(df: pd.DataFrame, month_label: str = "선택월") -> bytes:
    """
    xlsxwriter로 포맷된 엑셀을 메모리 바이트로 반환.
    - 제목 행
    - 2행 헤더(멀티헤더 병합)
    - 숫자/퍼센트 서식
    - 테두리/열너비
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = f"{month_label}"
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3, header=False)

        wb = writer.book
        ws = writer.sheets[sheet_name]

        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "left"})
        hdr_fmt   = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 1, "bg_color": "#F2F2F2"})
        cell_fmt  = wb.add_format({"align": "right", "border": 1, "num_format": "#,##0.0"})
        pct_fmt   = wb.add_format({"align": "right", "border": 1, "num_format": "0.0%"})
        left_fmt  = wb.add_format({"align": "left", "border": 1})

        # 제목
        ws.write(0, 0, f"플랫폼 기술본부 M/M 산정표 ({month_label})", title_fmt)

        # 헤더(멀티인덱스 지원)
        if isinstance(df.columns, pd.MultiIndex):
            top = [c[0] for c in df.columns]
            bot = [c[1] for c in df.columns]
            for j, v in enumerate(top):
                ws.write(2, j, v, hdr_fmt)
            for j, v in enumerate(bot):
                ws.write(3, j, v, hdr_fmt)
            # 같은 상단값 연속 병합
            start = 0
            for j in range(1, len(top) + 1):
                if j == len(top) or top[j] != top[start]:
                    if j - 1 > start:
                        ws.merge_range(2, start, 2, j - 1, top[start], hdr_fmt)
                    start = j
        else:
            for j, v in enumerate(df.columns):
                ws.write(2, j, v, hdr_fmt)

        # 데이터 서식
        nrows, ncols = df.shape
        for i in range(nrows):
            for j in range(ncols):
                val = df.iloc[i, j]
                fmt = left_fmt if j == 0 else cell_fmt

                # "비중(%)" 컬럼은 퍼센트 처리(0~100 → 0~1)
                colname = df.columns[j][1] if isinstance(df.columns, pd.MultiIndex) else str(df.columns[j])
                if "비중" in str(colname) and isinstance(val, (int, float)):
                    ws.write_number(4 + i, j, float(val) / 100.0, pct_fmt)
                    continue

                if isinstance(val, (int, float)):
                    ws.write_number(4 + i, j, float(val), fmt)
                else:
                    ws.write(4 + i, j, str(val), left_fmt)

        # 열 너비
        for j in range(ncols):
            ws.set_column(j, j, 14)

    return output.getvalue()
