"""Phase 4: 定量項目のみで「優先調査対象（最大30）」を抽出する（定性評価は Phase 5 で実施）。

    python3 -I fi_screening/scripts/select_priority.py

- 対象：FY2025 の外国証券（VERIFIED/ESTIMATED）が取得できた先のみ。取得できていない先は選ばない（無理に件数を満たさない）。
- 事前スコア＝ブックと同じ配点・同じ業態内パーセンタイル帯で A1+A2+B4+C1+C3 を計算（満点 39）。
- 全国偏り防止：セグメント別に上限（地銀・第二地銀 16、信金・大 8、信金・中 5、信金・小 1）を置き、合計30以内。
  信金中金は単独プロファイル対象として別枠（stage="別枠"）。
- 外国証券がゼロ（外債保有なし）の先は Type D 候補として理由欄に記載するが、優先30からは除外。
出力: data/selection.csv（stage, reason, exclusion）。理由欄は事実（数値）と推論を分けて記述。
"""
import csv
import pathlib
from collections import defaultdict

DATA = pathlib.Path(__file__).resolve().parents[1] / "data"
P0, P1, P2 = "FY2025(2026/3期)", "FY2024(2025/3期)", "FY2023(2024/3期)"
QUOTA = {"地銀・第二地銀": 16, "信金・大": 8, "信金・中": 5, "信金・小": 1}
W = dict(A1=10, A2=10, B4=5, C1=8, C3=6)


def read(n):
    p = DATA / n
    return list(csv.DictReader(open(p, encoding="utf-8"))) if p.exists() else []


def band(p):
    return 1 if p >= 0.8 else 0.7 if p >= 0.6 else 0.5 if p >= 0.4 else 0.3 if p >= 0.2 else 0.1


def pct(vals, v, reverse=False):
    xs = [x for x in vals if x is not None]
    if v is None or len(xs) < 2:
        return None
    n = sum(1 for x in xs if (x > v if reverse else x < v))
    return n / (len(xs) - 1)


def main():
    uni = {u["fi_code"]: u for u in read("universe.csv")}
    vals = {}
    status = {}
    for r in read("manual_financial.csv") + read("financial_data.csv") + read("auto_financial.csv"):
        k = (r["fi_code"], r["metric"], r["period"])
        if k in vals or r["value"] == "":
            continue
        vals[k] = float(r["value"])
        status[k] = r["status"]
    g = lambda c, m, p=P0: vals.get((c, m, p))  # noqa: E731

    def seg(c):
        u = uni[c]
        if u["category"] == "信金中央金庫":
            return "信金中金"
        if u["category"] == "信用金庫":
            d = g(c, "預金")
            return "信金（規模未判定）" if d is None else "信金・大" if d >= 1_000_000 else "信金・中" if d >= 300_000 else "信金・小"
        return "地銀・第二地銀"

    rows = []
    for c in uni:
        fx, sec = g(c, "外国証券"), g(c, "有価証券")
        rows.append(dict(code=c, seg=seg(c), fx=fx, sec=sec, ratio=(fx / sec if fx is not None and sec else None),
                         oci=(g(c, "その他有価証券評価差額") / g(c, "純資産") if g(c, "その他有価証券評価差額") is not None and g(c, "純資産") else None),
                         ltd=(g(c, "貸出金") / (g(c, "預金") + (g(c, "譲渡性預金") or 0)) if g(c, "貸出金") is not None and g(c, "預金") else None),
                         fx1=g(c, "外国証券", P1), fx2=g(c, "外国証券", P2)))
    byseg = defaultdict(list)
    for r in rows:
        byseg[r["seg"]].append(r)
    for s, rs in byseg.items():
        for r in rs:
            p = dict(A1=pct([x["fx"] for x in rs], r["fx"]), A2=pct([x["ratio"] for x in rs], r["ratio"]),
                     B4=pct([x["oci"] for x in rs], r["oci"], reverse=True), C1=pct([x["sec"] for x in rs], r["sec"]),
                     C3=pct([x["ltd"] for x in rs], r["ltd"], reverse=True))
            r["pts"] = {k: (W[k] * band(v) if v is not None else None) for k, v in p.items()}
            got = [v for v in r["pts"].values() if v is not None]
            mx = sum(W[k] for k, v in r["pts"].items() if v is not None)
            r["pre"] = sum(got) / mx * 100 if mx else None
            r["pct"] = p
    out = []
    picked = []
    for s, q in QUOTA.items():
        cands = [r for r in byseg.get(s, []) if r["fx"] is not None and r["fx"] > 0 and r["pre"] is not None]
        cands.sort(key=lambda r: -r["pre"])
        picked += cands[:q]
    picked = sorted(picked, key=lambda r: -r["pre"])[:30]
    pick_codes = {r["code"] for r in picked}
    for r in rows:
        c = r["code"]
        u = uni[c]
        if r["seg"] == "信金中金":
            out.append(dict(fi_code=c, stage="別枠", reason="信金中央金庫は個別信用金庫と別の投資主体として単独で詳細分析（業態内ランキング対象外）", exclusion=""))
            continue
        if c in pick_codes:
            chg = ""
            if r["fx1"]:
                chg = f"、前期比{(r['fx'] / r['fx1'] - 1) * 100:+.1f}%"
            facts = (f"【事実】FY2025 外国証券 {r['fx']:,.0f}百万円（有価証券比 {r['ratio'] * 100:.1f}%{chg}）、有価証券 {r['sec']:,.0f}百万円"
                     f"（出典は 3_Financial Data 参照、ステータス {status.get((c, '外国証券', P0), '')}）。")
            inf = (f"【推論】{r['seg']}内で外国証券規模 上位{(1 - r['pct']['A1']) * 100:.0f}%・比率 上位{(1 - r['pct']['A2']) * 100:.0f}%"
                   if r["pct"]["A1"] is not None and r["pct"]["A2"] is not None else "【推論】")
            inf += f"。定量事前スコア {r['pre']:.1f}（A1+A2+B4+C1+C3、100換算）。外債運用方針・ヘッジ状況は Phase5 で原文確認。"
            out.append(dict(fi_code=c, stage="30", reason=facts + inf, exclusion=""))
        elif r["fx"] is None:
            out.append(dict(fi_code=c, stage="", reason="", exclusion="FY2025 外国証券が未取得（NOT YET CHECKED）のため選定保留"))
        elif r["fx"] == 0:
            out.append(dict(fi_code=c, stage="", reason="",
                            exclusion="外国証券残高ゼロ（開示値）。外債投資の制約・方針を確認する Type D 的論点として別途検討"))
        else:
            out.append(dict(fi_code=c, stage="", reason="", exclusion=f"定量事前スコア {r['pre']:.1f}：セグメント枠内の上位に入らず" if r["pre"] is not None else "比較可能な指標不足"))
    with open(DATA / "selection.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["fi_code", "stage", "reason", "exclusion"])
        w.writeheader()
        w.writerows(out)
    print("picked", len(picked))
    for r in picked:
        print(r["code"], uni[r["code"]]["short_name"], r["seg"], f"{r['fx']:,.0f}", f"{r['ratio'] * 100:.1f}%", f"{r['pre']:.1f}")


if __name__ == "__main__":
    main()
