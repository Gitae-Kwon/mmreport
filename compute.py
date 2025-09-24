
import pandas as pd

def list_month_tabs(xl: pd.ExcelFile):
    """엑셀 파일의 시트명 리스트를 반환 (스트림릿에서 사용)."""
    return xl.sheet_names

def compute_report_frames(df: pd.DataFrame) -> pd.DataFrame:
    """
    ❗️여기를 실제 로직으로 교체하세요.
    현재는 데모용으로 다음을 수행합니다:
    - '플랫폼'과 '부서' 같은 열이 있다 가정하고 비율(%) 피벗 생성
    - 없는 경우, 임의로 첫 두 열을 키로 사용
    """
    if df.empty:
        return df

    # 컬럼 가드
    cols = list(df.columns)
    if len(cols) < 3:
        # 숫자열이 하나도 없으면 시연용 숫자 컬럼 추가
        out = df.copy()
        out["가중치"] = 1
        return out

    # 숫자열 선택
    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not num_cols:
        # 숫자열 없으면 1로 채운 임의 수치 생성
        df = df.copy()
        df["값"] = 1
        num_cols = ["값"]

    # 첫 두 열을 그룹키로 가정
    keys = [c for c in df.columns if c not in num_cols][:2]
    if not keys:
        keys = [df.columns[0]]

    # 합계 및 비중
    g = df.groupby(keys)[num_cols].sum().reset_index()
    total = g[num_cols].sum().sum()
    if total == 0:
        g["비중(%)"] = 0.0
    else:
        g["비중(%)"] = g[num_cols].sum(axis=1) / total * 100

    # 정렬
    g = g.sort_values("비중(%)", ascending=False).reset_index(drop=True)

    # 멀티헤더 샘플 (회사 양식 느낌)
    g.columns = pd.MultiIndex.from_tuples(
        [(keys[0], "")] + ([(keys[1], "")] if len(keys) > 1 else []) +
        [("합계", c) for c in num_cols] + [("합계", "비중(%)")]
    )

    return g
