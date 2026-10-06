"""パイプラインExcel(MOC Pipeline)の読み込み。

見出し行と実データの列がずれているため、'Sales month' 列の位置を起点に
位置ベースで列名を割り当てる(担当列は 'Sales month' の1つ左)。
"""
import pandas as pd

# 'Sales month' 列からの相対位置 → 正規化列名(ユーザー定義の A〜O 列に対応)
OFFSETS = {
    -1: "担当",
    0: "時期",
    2: "モデルイニシャル",
    5: "モデル詳細",
    6: "施設",
    7: "案件確度",
    10: "台数",
    11: "区分",          # FCT / backup / nFCT(=n月FCTで成約)
    12: "顧客ジャンル",  # AC=アカデミア, BT=バイオ企業, PHM=製薬, BTL
    13: "納期確度",      # A/B/C
    14: "価格",
    15: "代理店",
    16: "課題",
    17: "アクション",
    18: "アクション時期",
}


def load_pipeline(path, sheet=None):
    raw = pd.read_excel(path, sheet_name=sheet or 0, header=None)
    hdr_row, s = None, None
    for i in range(min(10, len(raw))):
        for j, v in enumerate(raw.iloc[i]):
            if isinstance(v, str) and v.strip() == "Sales month":
                hdr_row, s = i, j
                break
        if hdr_row is not None:
            break
    if hdr_row is None:
        raise ValueError(f"'Sales month' 見出しが見つかりません: {path}")
    body = raw.iloc[hdr_row + 1:]
    # 代理店〜アクション時期の列は、見出しに '代理店' がある版(10/5版以降)のみ存在する
    has_agent = any(isinstance(v, str) and "代理店" in v for v in raw.iloc[hdr_row])
    out = pd.DataFrame(index=body.index)
    for off, name in OFFSETS.items():
        j = s + off
        if 0 <= j < raw.shape[1] and (off < 15 or has_agent):
            out[name] = body.iloc[:, j]
        else:
            out[name] = pd.NA
    out["時期"] = pd.to_numeric(out["時期"], errors="coerce")
    out = out[out["時期"].notna()].copy()
    out["時期"] = out["時期"].astype(int)
    out["案件確度"] = pd.to_numeric(out["案件確度"], errors="coerce").fillna(0)
    out["台数"] = pd.to_numeric(out["台数"], errors="coerce").fillna(1)
    out["区分"] = out["区分"].astype(str).str.strip()
    out["モデルイニシャル"] = out["モデルイニシャル"].astype(str).str.strip()
    out["Excel行"] = out.index + 1
    return out.reset_index(drop=True)

