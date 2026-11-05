"""extract_securities.py / extract_shinkin.py の抽出結果を、データ・インテグリティ列付きの縦持ちデータに変換する。

    python3 -I fi_screening/scripts/ingest_extracted.py

入力: data/extracted_foreign.csv, data/extracted_shinkin.csv（あれば）, data/financial_data.csv
出力: data/auto_financial.csv（機械抽出値。手入力は manual_financial.csv に）, data/auto_sources.csv
      data/financial_data.csv の全銀協計数のうち、一次資料と一致した値の status を VERIFIED に更新
"""
import csv
import hashlib
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
ACCESSED = "2026-11-05"
FY = {2025: "FY2025(2026/3期)", 2024: "FY2024(2025/3期)", 2023: "FY2023(2024/3期)", 2022: "FY2022(2023/3期)"}


def read(n):
    p = DATA / n
    return list(csv.DictReader(open(p, encoding="utf-8"))) if p.exists() else []


def write(n, rows, fields):
    with open(DATA / n, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def sid(url):
    return "S-DISC-" + hashlib.md5(url.encode()).hexdigest()[:8].upper()


def main():
    uni = {u["fi_code"]: u for u in read("universe.csv")}
    ext = read("extracted_foreign.csv") + read("extracted_shinkin.csv")
    out, srcs, seen_src = [], [], set()
    verified_keys = set()
    for e in ext:
        code, y = e["fi_code"], int(e["year"])
        if y not in FY:
            continue
        period = FY[y]
        s = sid(e["url"])
        if s not in seen_src:
            seen_src.add(s)
            srcs.append(dict(source_id=s, fi_code=code, doc_type="ディスクロージャー誌／決算資料（PDF）",
                             title=f"{uni.get(code, {}).get('short_name', code)} 公表資料（{e['url'].rsplit('/', 1)[-1]}）",
                             period="", publisher=uni.get(code, {}).get("legal_name", ""), url=e["url"],
                             published_date="", accessed_date=ACCESSED, tier="Tier1", basis="単体",
                             notes="クローラで取得。資料名・公表日は原資料表紙で要確認"))
        checks = e["checks"]
        strong = checks != "内訳合計" and len(checks.split("+")) >= 2
        status = "VERIFIED" if (strong or checks == "内訳合計") else "SECONDARY"
        how = (f"{e['method']}／全銀協計数と一致: {checks}" if checks != "内訳合計"
               else f"{e['method']}／内訳の合計が表の合計と一致。年度は見出し位置で判定")
        base = dict(fi_code=code, fi_name=uni.get(code, {}).get("short_name", e.get("name", "")), unit="百万円", period=period,
                    basis="単体", source_id=s, page=f"p.{e['page']}（PDF通しページ）", accessed_date=ACCESSED)
        present = {"外国債券": e.get("fb_row") == "True", "外国株式": e.get("fe_row") == "True"}
        vals = {"外国債券": e.get("foreign_bond", ""), "外国株式": e.get("foreign_equity", ""), "外国証券": e.get("foreign_sec", "")}
        for m in ("外国債券", "外国株式", "外国証券"):
            if vals[m] != "":
                out.append(dict(base, metric=m, metric_group="B", value=vals[m], status=status, formula="", note=how))
        # 外国証券（スコア用）：開示が無ければ 外国債券＋外国株式 で推計
        if vals["外国証券"] == "" and vals["外国債券"] != "":
            fe = float(vals["外国株式"]) if vals["外国株式"] != "" else 0.0
            out.append(dict(base, metric="外国証券", metric_group="B", value=float(vals["外国債券"]) + fe, status="ESTIMATED",
                            formula="外国債券＋外国株式（同表の開示値。外国株式行が無い場合は外国債券のみ）",
                            note=how + ("" if present["外国株式"] else "／外国株式の行なし")))
        if vals["外国債券"] == "" and vals["外国証券"] == "":
            out.append(dict(base, metric="外国証券", metric_group="B", value="", status="NOT YET CHECKED", formula="",
                            note=f"{how}／一致した表に外国証券の内訳行なし。他資料（有価証券の時価等）で要確認"))
        # 信金：基礎計数も抽出している場合
        for m in ("国債", "地方債", "短期社債", "社債", "株式", "その他の証券", "有価証券", "預金", "貸出金", "総資産", "自己資本比率", "純資産"):
            if e.get(m, "") != "" and e.get(m) is not None:
                out.append(dict(base, metric=m, metric_group="A" if m in ("預金", "貸出金", "総資産", "自己資本比率", "純資産") else "B",
                                value=e[m], unit="%" if m == "自己資本比率" else "百万円", status=status, formula="", note=how))
        if checks != "内訳合計":
            for m in checks.split("+"):
                verified_keys.add((code, m, period))
            verified_keys.add((code, "有価証券", period)) if e.get("total_matched") else None

    fields = ["fi_code", "fi_name", "metric", "metric_group", "value", "unit", "period", "basis", "source_id", "page", "status",
              "accessed_date", "formula", "note"]
    write("auto_financial.csv", out, fields)
    write("auto_sources.csv", srcs, ["source_id", "fi_code", "doc_type", "title", "period", "publisher", "url", "published_date",
                                     "accessed_date", "tier", "basis", "notes"])
    fin = read("financial_data.csv")
    n = 0
    for r in fin:
        if (r["fi_code"], r["metric"], r["period"]) in verified_keys and r["status"] == "SECONDARY":
            r["status"] = "VERIFIED"
            r["note"] = "全銀協計数（Tier2）。各行ディスクロージャー誌（Tier1）の同値を確認済み"
            n += 1
    write("financial_data.csv", fin, list(fin[0].keys()))
    print(f"auto rows={len(out)} sources={len(srcs)} upgraded_to_verified={n}")


if __name__ == "__main__":
    main()
