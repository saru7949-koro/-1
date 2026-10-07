"""役員報告スライド用の数字を計算し JSON に書き出す。

使い方: python3 scripts/deck_metrics.py 2026-10-07 [出力先.json]
build_exec_deck.js はこの JSON を読んでスライドを組み立てる。
"""
import json
import math
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from analyze import PLAN_MODELS, compute  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent

# FY20-25 商談総括「分析サマリー」の記載値(勝敗の総括は年次で変わらないため固定)
WINLOSS = {"rate": 0.74, "win": 335, "loss": 118,
           "models": [["SH", 0.79], ["MA", 0.77], ["SA", 0.67], ["ID", 0.58]]}
SHORT_NAME = {"富士フイルム和光純薬": "富士フイルム和光"}


def r1(x):
    return round(float(x), 1)


def main(day, out_path):
    s, cfg = compute(day)
    a, b, ag, sc = s["a"], s["b"], s["agents"], s["sc"]
    p, q = cfg["period"], cfg["q3"]
    wr = {m: v["win_rate"] for m, v in cfg["models"].items()}
    plan = list(PLAN_MODELS)

    t9 = s["land9"].query("モデル == '計'").iloc[0]
    tl = s["land10"].query("モデル == '計'").iloc[0]
    target = sum(v["target_9f"] for v in cfg["models"].values())

    mv = s["moved"]
    mvp = mv[mv["モデル"].isin(plan)]
    cnt = lambda k: int((mvp["変化"] == k).sum())  # noqa: E731
    h2 = b[b["open_FCT"] & (b["時期"] >= p["h2_start"])]
    mt = s["month_t"]

    # 代理店(下期)上位5社: 着地寄与順
    known = ag.drop(index=[i for i in ("(未記入)", "未定") if i in ag.index])
    top5 = known.sort_values("着地寄与(台)", ascending=False).head(5)
    agents = [[SHORT_NAME.get(n, n), int(r["下期FCT"]), int(r["下期Backup"]), r1(r["着地寄与(台)"]),
               int(r["9F以降の期外移動"]), None if pd.isna(r["FY20-25勝率"]) else round(float(r["FY20-25勝率"]), 2)]
              for n, r in top5.iterrows()]
    w = ag.loc["富士フイルム和光純薬"] if "富士フイルム和光純薬" in ag.index else None

    # 年間の挽回
    row = lambda name: float(sc[sc["項目"].str.startswith(name)].iloc[0]["台"])  # noqa: E731
    up, pb, tm, gl, ex = row("①"), row("②"), row("③"), row("④"), row("⑤")
    short = float(sc.iloc[0]["台"])
    gap = row("残ギャップ")
    k = (short - (tm + gl + ex)) / (up + pb) if (up + pb) > 0 else 0
    rates = cfg["scenario"]

    # Q3
    def q3(d):
        x = d[d["時期"].between(q["start"], q["end"])]
        return x, x[x["open_FCT"] & x["モデルイニシャル"].isin(plan)], x[x["open_BK"] & x["モデルイニシャル"].isin(plan)]
    q9, q9f, q9b = q3(a)
    ql, qlf, qlb = q3(b)
    q3_land = lambda f, bk, x: float(x.loc[x["成約"], "台数"].sum() + (f["台数"] * f["勝率"]).sum()  # noqa: E731
                                     + bk["台数"].sum() * cfg["backup_rate"])
    land9_q3, landl_q3 = q3_land(q9f, q9b, q9), q3_land(qlf, qlb, ql)
    won_q3 = int(ql.loc[ql["成約"], "台数"].sum())
    y, mo = divmod(q["end"], 100)
    q4_start = (y + 1) * 100 + 1 if mo == 12 else q["end"] + 1
    q4 = b[b["時期"].between(q4_start, p["fy_end"]) & b["open_FCT"] & b["モデルイニシャル"].isin(plan)]
    q4c = q4[q4["納期確度"].isin(["A", "B"]) | (q4["案件確度"] >= q["pf_min_prob"])]
    pf_units = math.floor(len(q4c) * q["pull_forward_share"] + 0.5)
    close = max(0, min(len(qlf), q["bgt"] - won_q3 - q["backup_upgrade"] - pf_units))
    qa = qlf.groupby("一次店").agg(FCT=("台数", "size"), C=("納期確度", lambda v: int((v == "C").sum())),
                                    見込=("着地寄与", "sum"))
    qa["BK"] = qlb.groupby("一次店").size()
    qa = qa.fillna(0).sort_values(["FCT", "見込"], ascending=False)
    main_ag = [i for i in qa.index if i not in ("(未記入)", "未定")][:4]
    rows = [[SHORT_NAME.get(n, n), int(qa.loc[n, "FCT"]), int(qa.loc[n, "BK"]), int(qa.loc[n, "C"]), r1(qa.loc[n, "見込"])]
            for n in main_ag]
    rest = qa.drop(index=main_ag)
    if len(rest):
        rows.append(["その他・未記入", int(rest["FCT"].sum()), int(rest["BK"].sum()), int(rest["C"].sum()), r1(rest["見込"].sum())])
    rows.append(["計", len(qlf), len(qlb), int((qlf["納期確度"] == "C").sum()), r1(landl_q3 - won_q3)])
    top2 = qa.head(2)
    m = {
        "label": cfg["source"].get("latest_label", "最新"),
        "as_of": p["as_of"],
        "annual": {
            "target": target, "h1": cfg["carryover"]["h1_actual"], "h2": target - cfg["carryover"]["h1_actual"],
            "land9": r1(t9["着地理論値"]), "land": r1(tl["着地理論値"]),
            "short9": r1(-t9["不足(台)"]), "short": r1(-tl["不足(台)"]),
            "need9_deck": 21, "need": int(round(float(tl["必要リード(件)"]))),
            "fct9": int(t9["FCT案件"]), "fct": int(tl["FCT案件"]), "bk9": int(t9["Backup"]), "bk": int(tl["Backup"]),
            "moved": cnt("期外移動(来期以降)"),
            "moved_bk": int(((mvp["変化"] == "期外移動(来期以降)") & (mvp["9F区分"] == "backup")).sum()),
            "new": cnt("新規追加"), "lost": cnt("0%化(LOST等)"),
            "h2_fct": int(mt.loc["計", "計"]), "mar": int(mt.loc[202703, "計"]) if 202703 in mt.index else 0,
            "mar_c": int(mt.loc[202703, "C"]) if (202703 in mt.index and "C" in mt.columns) else 0,
        },
        "winloss": WINLOSS,
        "agents": {"rows": agents,
                   "share": round(float(top5["着地寄与(台)"].sum() / ag["着地寄与(台)"].sum()), 2),
                   "wako_fct": int(w["下期FCT"]) if w is not None else 0,
                   "wako_moved": int(w["9F以降の期外移動"]) if w is not None else 0,
                   "wako_open": int(w["期日超過"] + w["アクション未記入"]) if w is not None else 0},
        "recovery": {
            "short": r1(short), "up": r1(up), "pb": r1(pb), "tm": r1(tm), "lead": r1(gl + ex),
            "agent": r1(up + pb), "inhouse": r1(tm + gl + ex), "filled": r1(up + pb + tm + gl + ex), "gap": r1(gap),
            "land_after": r1(float(tl["着地理論値"]) + up + pb + tm + gl + ex),
            "rate_up": rates["upgrade_rate"], "rate_pb": rates["pullback_rate"], "rate_lead": rates["lead_meeting_rate"],
            "stretch_up": round(rates["upgrade_rate"] * k, 2), "stretch_pb": round(rates["pullback_rate"] * k, 2),
        },
        "q3": {
            "bgt": q["bgt"], "land9": r1(land9_q3), "land": r1(landl_q3), "gap": r1(q["bgt"] - landl_q3),
            "won": won_q3, "fct9": len(q9f), "bk9": len(q9b), "fct": len(qlf), "bk": len(qlb),
            "fct_c": int((qlf["納期確度"] == "C").sum()),
            "mto": int(qlf["MTO"].astype(str).str.contains("MTO").sum()),
            "models": {mm: int((qlf["モデルイニシャル"] == mm).sum()) for mm in plan if (qlf["モデルイニシャル"] == mm).any()},
            "genre": {g: int((qlf["顧客ジャンル"] == g).sum()) for g in ("AC", "BT", "PHM") if (qlf["顧客ジャンル"] == g).any()},
            "agents": rows,
            "top2": [SHORT_NAME.get(n, n) for n in top2.index], "top2_fct": int(top2["FCT"].sum()),
            "close": close, "close_rate": round(close / len(qlf), 2) if len(qlf) else 0,
            "close_expected": int(round(float(qlf["勝率"].sum()))),
            "bk_up": q["backup_upgrade"], "pf": pf_units, "pf_cand": len(q4c),
            "max_all": len(qlf) + won_q3,
        },
        "actions": {
            "a_cases": int((s["acts"]["優先度"] == "A").sum()),
            "bk_cases": int((s["op"]["open_BK"] & s["op"]["モデルイニシャル"].isin(plan)).sum()),
            "pull_cases": len(set(mvp.loc[mvp["変化"] == "期外移動(来期以降)", "Excel行"].dropna().astype(int))
                              | set(s["pf"].loc[s["pf"]["モデルイニシャル"].isin(plan), "Excel行"].astype(int))),
            "prospecting": int(s["lr"]["Prospecting"].sum()),
        },
    }
    Path(out_path).write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(m, ensure_ascii=False))


if __name__ == "__main__":
    day = sys.argv[1] if len(sys.argv) > 1 else "2026-10-07"
    main(day, sys.argv[2] if len(sys.argv) > 2 else ROOT / "出力" / f"deck_metrics_{day.replace('-', '')}.json")
