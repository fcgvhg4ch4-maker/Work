"""Batch 3-5: 信用金庫・信金中金のディスクロージャー誌テキストから基礎計数と有価証券の種類別残高を抽出する。

    python3 -I fi_screening/scripts/extract_shinkin.py <raw_dir> [code ...]

- 有価証券：種類別期末残高表で「内訳の合計＝表の合計」となる列のみ採用（extract_securities.method3）。
- 基礎計数：「主要な経営指標の推移」表の 総資産額・預金積金残高・貸出金残高・有価証券残高・純資産額・自己資本比率。
  年度は見出し行の年度表記の桁位置で対応付け、経営指標表の有価証券残高＝種類別表の合計 の一致を検証条件とする。
- 単位（千円／百万円／億円）は表の直前の「単位」表記で判定し、百万円に換算。
出力: data/extracted_shinkin.csv
"""
import csv
import importlib.util
import json
import pathlib
import re
import sys

RAW = pathlib.Path(sys.argv[1]).resolve()
ONLY = set(sys.argv[2:])
ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
spec = importlib.util.spec_from_file_location("xs", pathlib.Path(__file__).with_name("extract_securities.py"))
xs = importlib.util.module_from_spec(spec)
sys.argv, _argv = [sys.argv[0], str(RAW)], sys.argv
spec.loader.exec_module(xs)
sys.argv = _argv

UNIT = re.compile(r"単位[：:\s]*(千円|百万円|億円|円)")
SCALE = {"千円": 0.001, "百万円": 1.0, "億円": 100.0, "円": 1e-6}
KPI = [("総資産", r"^総資産(額)?$"), ("預金", r"^預金積金残高$|^預金(積金)?残高$"), ("貸出金", r"^貸出金残高$"),
       ("有価証券", r"^有価証券残高$"), ("純資産", r"^純資産(額|の部合計)?$"),
       ("自己資本比率", r"^(単体)?自己資本比率(（単体）)?$")]
HDR = re.compile(r"(20[12]\d)\s*年\s*(?:度|3\s*月)|令和\s*(\d|元)\s*年\s*(?:度|3\s*月)")


def unit_before(text, idx):
    m = list(UNIT.finditer(text[:idx]))
    return m[-1].group(1) if m else None


def year_of(m):
    y = int(m.group(1)) if m.group(1) else 2018 + (1 if m.group(2) == "元" else int(m.group(2)))
    if "月" in m.group(0):
        y -= 1
    return y


def securities(pages):
    """method3（種類別期末残高）→ 無ければ method4（残存期間別・年度ブロック）。単位換算付き {year: dict}"""
    res = {}
    found = xs.method4("", pages, {})
    found.update(xs.method3("", pages, {}))
    for (y,), v in found.items():
        pg = pages[v["page"] - 1]
        # 合計値が出現する位置の直前の単位表記
        i = pg.find(f"{int(v['total']):,}")
        u = unit_before(pg, i if i >= 0 else len(pg)) or "百万円"
        k = SCALE[u]
        parts = {m: (None if val is None else val * k) for m, val in v["parts"].items()}
        res[y] = dict(year=y, page=v["page"], unit=u, total=v["total"] * k, parts=parts, method=v["method"],
                      fb_row=v["row_present"]["外国債券"], fe_row=v["row_present"]["外国株式"])
    return res


def kpis(pages):
    out = {}
    for pi, pg in enumerate(pages, 1):
        if not re.search(r"主要な?経営指標", pg):
            continue
        lines = pg.split("\n")
        hdr = None
        for li, line in enumerate(lines):
            if len(HDR.findall(line)) >= 3:
                hdr = [(m.end(), year_of(m)) for m in HDR.finditer(line)]
                hdr_idx = sum(len(x) + 1 for x in lines[:li])
                continue
            if not hdr:
                continue
            head = re.split(r"[\d△▲]", line, maxsplit=1)[0]
            key = re.sub(r"[\s　（）()百万円千%％]", "", head)
            for m, rx in KPI:
                if re.search(rx, key):
                    u = unit_before(pg, hdr_idx + 1) or "百万円"
                    for pos, v, _ in xs.toks(line[len(head):]):
                        pos += len(head)
                        y = min(hdr, key=lambda h: abs(h[0] - pos))
                        if abs(y[0] - pos) > 8:
                            continue
                        val = v if m == "自己資本比率" else v * SCALE[u]
                        out.setdefault((y[1], m), (val, pi, u))
    return out


BS = [("総資産", r"^資産の部合計$"), ("預金", r"^預金積金$"), ("貸出金", r"^貸出金$"), ("有価証券", r"^有価証券$"),
      ("純資産", r"^純資産の部合計$")]
BS_HDR = re.compile(r"(20[12]\d)\s*年\s*3\s*月\s*末?|令和\s*(\d)\s*年\s*3\s*月\s*末?|(20[12]\d)\s*年度末?")


def balance_sheet(pages):
    """単体貸借対照表（2期比較）から 総資産・預金・貸出金・有価証券・純資産。連結のページは除外。"""
    out = {}
    for pi, pg in enumerate(pages, 1):
        if "貸借対照表" not in pg or "連結" in pg:
            continue
        lines = pg.split("\n")
        hdr = None
        for li, line in enumerate(lines):
            ms = list(BS_HDR.finditer(line))
            if len(ms) >= 2:
                hdr = []
                for m in ms:
                    if m.group(1):
                        y = int(m.group(1)) - 1
                    elif m.group(2):
                        y = 2018 + int(m.group(2)) - 1
                    else:
                        y = int(m.group(3))
                    hdr.append((m.end(), y))
                hdr_idx = sum(len(x) + 1 for x in lines[:li])
                continue
            if not hdr:
                continue
            # 左右2段組の B/S に対応：行内の各ラベル→直後の数値2つ
            for seg in re.finditer(r"([^\d△▲\s][^\d△▲]*?)\s+([△▲]?[\d,]+)\s+([△▲]?[\d,]+)", line):
                key = re.sub(r"[\s　（）()]", "", seg.group(1))
                for m, rx in BS:
                    if re.search(rx, key):
                        u = unit_before(pg, hdr_idx + 1) or "百万円"
                        for gi in (2, 3):
                            pos = seg.end(gi)
                            y = min(hdr, key=lambda h: abs(h[0] - pos))
                            v = xs.toks(" " + seg.group(gi))
                            if v:
                                out.setdefault((y[1], m), (v[0][1] * SCALE[u], pi, u))
    return out


KPI_LINE = [("預金", r"預金積金残高|預金残高"), ("貸出金", r"貸出金残高"), ("有価証券", r"有価証券残高"), ("総資産", r"総資産額"),
            ("純資産", r"純資産額"), ("自己資本比率", r"自己資本比率")]


def kpi_lines(pages, sec_by_year):
    """経営指標の推移行（ラベル＋複数年の数値、最終列＝最新期、左ほど過去）。
    有価証券残高行の末尾から k 番目が、種類別合計の (最新年−k) 年と一致（同単位/千倍）することを検証条件とする。
    返り値 {(year, metric): (value_百万円, page)}"""
    cand = {}
    for pi, pg in enumerate(pages, 1):
        for line in pg.split("\n"):
            flat = re.sub(r"[\s　]", "", line)
            for m, rx in KPI_LINE:
                mm = re.search(rx, flat)
                if not mm:
                    continue
                last_ch = mm.group(0)[-1]
                after = line.split(last_ch, 1)[-1]
                tk = [t[1] for t in xs.toks(" " + after) if t[2] not in ("ー", "―", "－", "-", "—")]
                if len(tk) >= 2:
                    cand.setdefault(m, []).append((tk, pi))
    ylatest = max(sec_by_year)
    for tk, pi in cand.get("有価証券", []):
        for scale in (1.0, 0.001):
            if abs(tk[-1] * scale - sec_by_year[ylatest]) > max(5.0, sec_by_year[ylatest] * 0.0005):
                continue
            out = {}
            for k in range(min(3, len(tk))):
                y = ylatest - k
                if y in sec_by_year and abs(tk[-1 - k] * scale - sec_by_year[y]) > max(5.0, sec_by_year[y] * 0.0005):
                    break
                out[(y, "有価証券")] = (tk[-1 - k] * scale, pi)
                for m in ("預金", "貸出金", "総資産", "純資産"):
                    if cand.get(m) and len(cand[m][0][0]) > k:
                        out[(y, m)] = (cand[m][0][0][-1 - k] * scale, cand[m][0][1])
                r = [c for c in cand.get("自己資本比率", []) if len(c[0]) > k and all(2 <= x <= 60 for x in c[0])]
                if r:
                    out[(y, "自己資本比率")] = (r[0][0][-1 - k], r[0][1])
            return out
    return {}


def main():
    uni = {r["fi_code"]: r for r in csv.DictReader(open(DATA / "universe.csv", encoding="utf-8"))}
    rows = []
    for d in sorted((RAW / "disc").iterdir()):
        code = d.name
        if (ONLY and code not in ONLY) or uni.get(code, {}).get("category") not in ("信用金庫", "信金中央金庫"):
            continue
        if not (d / "result.json").exists():
            continue
        res = json.loads((d / "result.json").read_text())
        best = {}
        docs = []
        for c in res.get("checked", []):
            tx = d / c["file"].replace(".pdf", ".txt")
            if tx.exists():
                docs.append((c, tx.read_text(encoding="utf-8", errors="replace").split("\f")))
        # 分冊（同一フォルダ）の結合文書を追加：表と経営指標が別ファイルでも突合できるようにする
        groups = {}
        for c, pg in docs:
            groups.setdefault(c["url"].rsplit("/", 1)[0], []).append((c, pg))
        for folder, items in groups.items():
            if len(items) >= 2:
                items.sort(key=lambda x: [int(n) if n.isdigit() else n for n in re.split(r"(\d+)", x[0]["url"])])
                pmap, allp = [], []
                for c, pg in items:
                    for i, p_ in enumerate(pg, 1):
                        pmap.append((c["url"], i))
                        allp.append(p_)
                docs.append((dict(url=folder + "/（分冊結合）", file="", pmap=pmap), allp))
        for c, pages in docs:
            sec, kp = securities(pages), kpis(pages)
            for k, v in balance_sheet(pages).items():
                kp.setdefault(k, v)
            for y, s in sec.items():
                if (y, "有価証券") not in kp:
                    for key, val in kpi_lines(pages, {yy: ss["total"] for yy, ss in sec.items()}).items():
                        kp.setdefault(key, (val[0], val[1], "kpi-line"))
                kpi_sec = kp.get((y, "有価証券"))
                cross = kpi_sec is not None and abs(kpi_sec[0] - s["total"]) <= max(5.0, s["total"] * 0.0005)
                if s["method"].startswith("残存") and not cross:
                    continue  # 残存期間別表は有価証券残高との一致を必須とする
                score = 2 if cross else 1
                if y in best and best[y][0] >= score:
                    continue
                src_url, src_page = (c["pmap"][s["page"] - 1] if c.get("pmap") else (c["url"], s["page"]))
                rec = dict(fi_code=code, name=res["name"], year=y, method=s["method"], page=src_page,
                           checks=("内訳合計" if s["method"].startswith("種類別") else "残存期間別の種類計") + ("+有価証券残高（経営指標/BS）" if cross else ""), total_matched=s["total"],
                           foreign_bond=s["parts"].get("外国債券") if s["parts"].get("外国債券") is not None else "",
                           foreign_equity=s["parts"].get("外国株式") if s["parts"].get("外国株式") is not None else "",
                           foreign_sec=s["parts"].get("外国証券") if s["parts"].get("外国証券") is not None else "",
                           other_sec=s["parts"].get("その他の証券") if s["parts"].get("その他の証券") is not None else "",
                           fb_row=s["fb_row"], fe_row=s["fe_row"], url=src_url, file=c["file"], unit_src=s["unit"])
                rec["有価証券"] = s["total"]
                for m in ("国債", "地方債", "短期社債", "社債", "株式", "その他の証券"):
                    rec[m] = "" if s["parts"].get(m) is None else s["parts"][m]
                for m in ("総資産", "預金", "貸出金", "純資産", "自己資本比率"):
                    v = kp.get((y, m)) if cross else None
                    rec[m] = "" if v is None else v[0]
                best[y] = (score, rec)
        for y, (_, rec) in sorted(best.items()):
            if rec["checks"] == "内訳合計":
                rec["checks"] = "内訳合計"
            rows.append(rec)
    fields = ["fi_code", "name", "year", "method", "page", "checks", "total_matched", "foreign_bond", "foreign_equity", "foreign_sec",
              "other_sec", "fb_row", "fe_row", "url", "file", "unit_src", "有価証券", "国債", "地方債", "短期社債", "社債", "株式",
              "その他の証券", "総資産", "預金", "貸出金", "純資産", "自己資本比率"]
    with open(DATA / "extracted_shinkin.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)
    print(f"shinkin rows={len(rows)} institutions={len({r['fi_code'] for r in rows})}")
    for r in rows[:40]:
        print(r["fi_code"], r["name"], r["year"], r["checks"], "有価証券", r["有価証券"], "外証", r["foreign_sec"], r["foreign_bond"], "預金", r["預金"], r["unit_src"])


if __name__ == "__main__":
    main()
