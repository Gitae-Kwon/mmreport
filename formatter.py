
import io
import pandas as pd

def styled_preview_html(df: pd.DataFrame) -> str:
    """
    간단 HTML 미리보기 (Streamlit 표시용).
    """
    if isinstance(df.columns, pd.MultiIndex):
        styler = (df.style
                  .format(precision=1)
                  .set_table_attributes('class="table"')
                  .set_table_styles([
                      {"selector": "th", "props": [("text-align", "center"), ("background-color", "#f6f8fa")]},
                      {"selector": "td", "props": [("text-align", "right")]},
                  ])
                  .hide_indices())
    else:
        styler = df.style.format(precision=1).hide_indices()

    return styler.to_html()

def to_formatted_excel_bytes(df: pd.DataFrame, month_label: str = "선택월") -> bytes:
    """
    xlsxwriter로 포맷된 엑셀을 생성하여 바이트로 반환.
    실제 회사 양식(머지/굵게/테두리/퍼센트/소숫점 등)은 이 함수를 수정해서 맞추세요.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = f"{month_label}"
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3, header=False)

        wb  = writer.book
        ws  = writer.sheets[sheet_name]

        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "left"})
        hdr_fmt   = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "border":1, "bg_color": "#F2F2F2"})
        cell_fmt  = wb.add_format({"align": "right", "border":1, "num_format": "#,##0.0"})
        pct_fmt   = wb.add_format({"align": "right", "border":1, "num_format": "0.0%"})
        left_fmt  = wb.add_format({"align": "left", "border":1})

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
            # 상단 병합 (같은 값 연속 병합)
            start = 0
            for j in range(1, len(top)+1):
                if j == len(top) or top[j] != top[start]:
                    if j-1 > start:
                        ws.merge_range(2, start, 2, j-1, top[start], hdr_fmt)
                    start = j
        else:
            for j, v in enumerate(df.columns):
                ws.write(2, j, v, hdr_fmt)
            # 한 줄 헤더 모드
            # 데이터는 startrow=3에 이미 써졌으므로 1줄 당겨서 보이게
            pass

        # 데이터 서식
        nrows, ncols = df.shape
        for i in range(nrows):
            for j in range(ncols):
                val = df.iloc[i, j]
                fmt = cell_fmt
                # 첫 컬럼은 좌측정렬
                if j == 0:
                    fmt = left_fmt
                # "비중(%)" 컬럼은 퍼센트
                colname = df.columns[j][1] if isinstance(df.columns, pd.MultiIndex) else str(df.columns[j])
                if "비중" in str(colname):
                    try:
                        # 0-100 값을 0-1 로 변환해 퍼센트 표현
                        if isinstance(val, (int, float)):
                            ws.write_number(4+i, j, float(val)/100.0, pct_fmt)
                            continue
                    except:
                        pass
                # 숫자/문자 구분
                if isinstance(val, (int, float)):
                    ws.write_number(4+i, j, float(val), cell_fmt)
                else:
                    ws.write(4+i, j, str(val), left_fmt)

        # 폭 자동(대충)
        for j in range(ncols):
            ws.set_column(j, j, 14)

    return output.getvalue()
