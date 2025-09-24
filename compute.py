import pandas as pd

def compute_report_frames(df: pd.DataFrame) -> pd.DataFrame:
    """
    데모 로직:
    - 숫자열 합계를 기준으로 그룹 비중(%) 계산
    - 첫 두 열을 그룹키로 가정
    - MultiIndex 컬럼(회사 양식 느낌)으로 반환
    """
    if df.empty:
        return df

    cols = list(df.columns)
    if len(cols) < 3:
        out = df.copy()
        out["가중치"] = 1
        return out

    num_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    if not num_cols:
        df = df.copy()
        df["값"] = 1
        num_cols = ["값"]

    keys = [c for c in df.columns if c not in num_cols][:2]
    if not keys:
        keys = [df.columns[0]]

    g = df.groupby(keys)[num_cols].sum().reset_index()
    total = g[num_cols].sum().sum()
    g["비중(%)"] = 0.0 if total == 0 else g[num_cols].sum(axis=1) / total * 100
    g = g.sort_values("비중(%)", ascending=False).reset_index(drop=True)

    # MultiIndex 헤더 구성 (예: (부문,""), (팀,""), (합계,매출), (합계,비중(%)) ...)
    new_cols = [(keys[0], "")]
    if len(keys) > 1:
        new_cols.append((keys[1], ""))
    new_cols += [("합계", c) for c in num_cols]
    new_cols += [("합계", "비중(%)")]

    g.columns = pd.MultiIndex.from_tuples(new_cols)
    return g
