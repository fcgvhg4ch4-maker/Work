"""Batch 2b: ディスクロージャー誌テキストから「有価証券残高（期末）」表を読み、
外国債券・外国株式（・外国証券）を期別に抽出する。全銀協の行別計数と突合して列と期を確定する。

    python3 -I fi_screening/scripts/extract_securities.py <raw_dir> [code ...]

判定ロジック（誤読防止のため保守的）:
  1. pdftotext -layout の各ページで「合計」行の数値と、全銀協の 有価証券（期末・単体）を照合し、
     一致（±1百万円）した数値の桁位置＝その期の「合計」列とする（平均残高表は一致しないので除外される）。
  2. 同じ表の 国債・地方債・社債・株式 行の同じ列位置の値が全銀協計数と一致するかで列を検証。
  3. 同じ列位置の「うち外国債券」「うち外国株式」「外国証券」の値を採用。ダッシュ（―）は 0。
  4. 方式2（残存期間別残高表）：年度ラベル付き行の最終列（合計）で 国債 を全銀協と照合し、同年度の外国債券行を採用。
出力: data/extracted_foreign.csv（抽出結果。status=VERIFIED は 合計＋国債 の二重一致時のみ）
"""
import csv
import json
import pathlib
import re
import sys

RAW = pathlib.Path(sys.argv[1]).resolve()
ONLY = set(sys.argv[2:])
ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
PERIODS = {"FY2025(2026/3期)": 2025, "FY2024(2025/3期)": 2024, "FY2023(2024/3期)": 2023}

TOK = re.compile(r"[△▲]\s?[\d,]+|-?[\d][\d,]*(?:\.\d+)?|(?<![ァ-ヶ])[―－–—ー─━]+(?![ァ-ヶ])|(?<=\s)-(?=\s|$)")
LABELS = [  # (正規化キー, 正規表現: 空白除去後の行頭)
    ("外国債券", r"^(うち|う ち)?外国債券"), ("外国株式", r"^(うち)?外国株式"), ("外国証券", r"^(うち)?外国証券"),
    ("その他の証券", r"^その他の?証券"), ("短期社債", r"^短期社債"), ("国債", r"^国債"), ("地方債", r"^地方債"),
    ("社債", r"^社債"), ("株式", r"^株式"), ("合計", r"^(合計|計)$|^合計"),
]


def toks(line):
    out = []
    for m in TOK.finditer(line):
        s = m.group(0).replace(" ", "")
        if re.fullmatch(r"[―－–—ー─━\-]+", s):
            v = 0.0
        else:
            neg = s[0] in "△▲-"
            try:
                v = float(s.lstrip("△▲-").replace(",", ""))
            except ValueError:
                continue
            v = -v if neg else v
        out.append((m.end(), v, s))
    return out


def label_of(line):
    head = re.split(r"[\d△▲―－–—─━]|(?<![ァ-ヶ])ー(?![ァ-ヶ])", line, maxsplit=1)[0]
    key = re.sub(r"[\s　]", "", head)
    key = re.sub(r"^(有価証券|区分|種類)", "", key)
    for cand in (key, key[1:]):  # 左余白の縦書き文字（1字）を除いて再試行
        for k, rx in LABELS:
            if re.search(rx, cand):
                return k, len(head)
    return None, 0


def load_ref():
    ref = {}
    with open(DATA / "financial_data.csv", encoding="utf-8") as f:
        for r in csv.DictReader(f):
            if r["value"] and r["metric"] in ("有価証券", "国債", "地方債", "社債", "株式"):
                ref[(r["fi_code"], r["metric"], PERIODS[r["period"]])] = float(r["value"])
    return ref


def near(a, b, tol=1.0):
    return a is not None and b is not None and abs(a - b) <= tol


def method1(code, pages, ref):
    """期末残高表（国内/国際/合計 × 年度）"""
    found = {}
    for pi, pg in enumerate(pages, 1):
        if "有価証券" not in pg:
            continue
        lines = pg.split("\n")
        for li, line in enumerate(lines):
            lab, _ = label_of(line)
            if lab != "合計":
                continue
            tk = toks(line)
            for (pos, v, _), scale in [(t, sc) for t in tk for sc in (1.0, 100.0)]:
                tol = 1.0 if scale == 1.0 else 100.0
                for y in (2025, 2024, 2023):
                    tot = ref.get((code, "有価証券", y))
                    if not near(v * scale, tot, tol):
                        continue
                    # 同じ表の上方 18 行から各ラベル行を拾う
                    rows = {}
                    for lj in range(max(0, li - 18), li):
                        l2, _ = label_of(lines[lj])
                        if l2 and l2 != "合計":
                            cand = sorted([t for t in toks(lines[lj]) if abs(t[0] - pos) <= 7], key=lambda t: abs(t[0] - pos))
                            rows.setdefault(l2, cand[0][1] * scale if cand else None)
                    checks = [m for m in ("国債", "地方債", "株式") if near(rows.get(m), ref.get((code, m, y)), tol)]
                    if not checks:
                        continue
                    key = (y,)
                    if key in found and len(found[key]["checks"]) >= len(checks):
                        continue
                    found[key] = dict(year=y, page=pi, method="期末残高表" + ("（原資料単位: 億円→百万円換算）" if scale == 100.0 else ""), checks=checks, total=v * scale,
                                      外国債券=rows.get("外国債券"), 外国株式=rows.get("外国株式"), 外国証券=rows.get("外国証券"),
                                      その他の証券=rows.get("その他の証券"),
                                      row_present={k: k in rows for k in ("外国債券", "外国株式", "外国証券")})
    return found


YEAR_RX = re.compile(r"(20[12]\d)\s*年\s*度?|令和\s*(\d)\s*年\s*度?")


def method2(code, pages, ref):
    """残存期間別残高表（種類×年度の行、最終列=合計）"""
    found = {}
    for pi, pg in enumerate(pages, 1):
        if "残存期間" not in pg:
            continue
        lines = pg.split("\n")
        cur = None
        recs = []
        for li, line in enumerate(lines):
            lab, _ = label_of(line)
            ym = YEAR_RX.search(line)
            if lab and lab != "合計":
                cur = lab
            elif re.sub(r"\s", "", line).startswith(("うち", "うち")) and not lab:
                cur = "うち?"
            if not ym:
                continue
            y = int(ym.group(1)) if ym.group(1) else 2018 + int(ym.group(2))
            tk = [t for t in toks(line[ym.end():])]
            if not tk:
                continue
            recs.append((li, cur, y, tk[-1][1], lab))
        # 「うち?」は直後/直前の行のラベル（外国債券/外国株式）で補正
        fixed = []
        for i, (li, cur, y, v, lab) in enumerate(recs):
            if cur == "うち?":
                nb = [r for r in recs if abs(r[0] - li) <= 2 and r[4] in ("外国債券", "外国株式")]
                cur = nb[0][4] if nb else None
            fixed.append((li, cur, y, v))
        for y in (2025, 2024, 2023):
            g = [v for li, c, yy, v in fixed if c == "国債" and yy == y]
            if not g or not near(g[0], ref.get((code, "国債", y))):
                continue
            checks = ["国債"] + [m for m in ("地方債", "株式") if any(near(v, ref.get((code, m, y))) for li, c, yy, v in fixed if c == m and yy == y)]
            fb = [v for li, c, yy, v in fixed if c == "外国債券" and yy == y]
            fs = [v for li, c, yy, v in fixed if c == "外国株式" and yy == y]
            if (y,) not in found:
                found[(y,)] = dict(year=y, page=pi, method="残存期間別残高表", checks=checks, total=None,
                                   外国債券=fb[0] if fb else None, 外国株式=fs[0] if fs else None, 外国証券=None, その他の証券=None,
                                   row_present={"外国債券": bool(fb), "外国株式": bool(fs), "外国証券": False})
    return found


HDR_YEAR = re.compile(r"(20[12]\d)\s*年\s*(?:度|3\s*月)|令和\s*(\d)\s*年\s*(?:度|3\s*月)")


def method3(code, pages, ref):
    """内部整合方式（全銀協データが無い信金等）：種類別期末残高表で 内訳合計＝合計 となる列を採用。
    年度は表見出し行の年度表記の桁位置で対応付け（年度末=3月期。'2026年3月末'→FY2025）。"""
    found = {}
    for pi, pg in enumerate(pages, 1):
        if not re.search(r"有価証券", pg) or re.search(r"平均残高", pg) and not re.search(r"期末残高|種類別", pg):
            continue
        lines = pg.split("\n")
        for li, line in enumerate(lines):
            lab, _ = label_of(line)
            if lab != "合計":
                continue
            rows, hdr = {}, None
            for lj in range(li - 1, max(-1, li - 16), -1):
                l2, _ = label_of(lines[lj])
                if l2 and l2 != "合計" and l2 not in rows:
                    rows[l2] = toks(lines[lj])
                if HDR_YEAR.search(lines[lj]) and hdr is None and lj < li - 2:
                    hdr = lines[lj]
                if "平均残高" in lines[lj]:
                    rows = {}
                    break
            comp = [k for k in ("国債", "地方債", "短期社債", "社債", "株式", "外国証券", "その他の証券") if k in rows]
            if len(comp) < 4 or not hdr:
                continue
            for pos, tot, _ in toks(line):
                if tot < 100:
                    continue
                vals = {}
                for k in comp + ["外国債券", "外国株式"]:
                    c = sorted([t for t in rows.get(k, []) if abs(t[0] - pos) <= 7], key=lambda t: abs(t[0] - pos))
                    vals[k] = c[0][1] if c else None
                ssum = sum(v for k, v in vals.items() if k in comp and v is not None)
                if not near(ssum, tot, 2.0):
                    continue
                ys = [(m.end(), m) for m in HDR_YEAR.finditer(hdr)]
                if not ys:
                    continue
                endpos, m = min(ys, key=lambda x: abs(x[0] - pos))
                y = int(m.group(1)) if m.group(1) else 2018 + int(m.group(2))
                if "3" in m.group(0) and "月" in m.group(0):
                    y -= 1  # 2026年3月末 → FY2025
                if (y,) in found:
                    continue
                found[(y,)] = dict(year=y, page=pi, method="種類別期末残高表（内訳合計一致）", checks=["内訳合計"], total=tot,
                                   外国債券=vals.get("外国債券"), 外国株式=vals.get("外国株式"), 外国証券=vals.get("外国証券"),
                                   その他の証券=vals.get("その他の証券"), parts=vals,
                                   row_present={k: k in rows for k in ("外国債券", "外国株式", "外国証券")})
    return found



SEC_YEAR = re.compile(r"(20[12]\d)\s*年\s*3\s*月\s*末|令和\s*(\d)\s*年\s*3\s*月\s*末|(20[12]\d)\s*年度末")


def method4(code, pages, ref):
    """残存期間別残高表（年度ブロック型：『2026年3月末』見出しの下に種類別行、最終列=合計）。
    返り値の parts は各種類の合計列。検証は呼び出し側（全銀協 or 経営指標の有価証券残高）で行う。"""
    found = {}
    for pi, pg in enumerate(pages, 1):
        if "残存期間" not in pg:
            continue
        cur_y, blocks = None, {}
        for line in pg.split("\n"):
            ym = SEC_YEAR.search(line)
            if ym and len(xs_toks(line[ym.end():])) <= 1:
                cur_y = (int(ym.group(1)) - 1) if ym.group(1) else ((2018 + int(ym.group(2)) - 1) if ym.group(2) else int(ym.group(3)))
                continue
            if cur_y is None:
                continue
            lab, hl = label_of(line)
            if not lab or lab == "合計":
                continue
            tk = xs_toks(line[hl:])
            if len(tk) < 3:
                continue
            blocks.setdefault(cur_y, {}).setdefault(lab, tk[-1][1])
        for y, parts in blocks.items():
            comp = [k for k in ("国債", "地方債", "短期社債", "社債", "株式", "外国証券", "その他の証券") if k in parts]
            if len(comp) < 4 or (y,) in found:
                continue
            tot = sum(parts[k] for k in comp)
            found[(y,)] = dict(year=y, page=pi, method="残存期間別残高表（年度ブロック）", checks=[], total=tot,
                               外国債券=parts.get("外国債券"), 外国株式=parts.get("外国株式"), 外国証券=parts.get("外国証券"),
                               その他の証券=parts.get("その他の証券"), parts=parts,
                               row_present={k: k in parts for k in ("外国債券", "外国株式", "外国証券")})
    return found


xs_toks = toks


def main():
    ref = load_ref()
    out = []
    disc = RAW / "disc"
    for d in sorted(disc.iterdir()):
        code = d.name
        if ONLY and code not in ONLY or not (d / "result.json").exists():
            continue
        res = json.loads((d / "result.json").read_text())
        best = {}
        for c in res.get("checked", []):
            tx = d / c["file"].replace(".pdf", ".txt")
            if not tx.exists():
                continue
            pages = tx.read_text(encoding="utf-8", errors="replace").split("\f")
            m4 = {}
            for (y,), v in method4(code, pages, ref).items():
                chk = [m for m in ("国債", "地方債", "株式") if near(v["parts"].get(m), ref.get((code, m, y)))]
                if len(chk) >= 2:
                    v["checks"] = chk
                    m4[(y,)] = v
            m3 = {}
            for (y,), v in method3(code, pages, ref).items():
                if (code, "有価証券", y) in ref:  # 全銀協計数のある銀行は合計一致を必須に
                    if not near(v["total"], ref[(code, "有価証券", y)]):
                        continue
                    v["checks"] = ["合計"] + [m for m in ("国債", "地方債", "株式") if near(v["parts"].get(m), ref.get((code, m, y)))]
                m3[(y,)] = v
            for fnd in (method1(code, pages, ref), method2(code, pages, ref), m4, m3):
                for (y,), v in fnd.items():
                    v["url"] = c["url"]
                    v["file"] = c["file"]
                    score = len(v["checks"]) + (2 if v["method"] == "期末残高表" else 0) + (1 if v["外国債券"] is not None else 0) - (2 if v["checks"] == ["内訳合計"] else 0)
                    if y not in best or score > best[y][0]:
                        best[y] = (score, v)
        for y, (_, v) in sorted(best.items()):
            out.append(dict(fi_code=code, name=res["name"], year=y, method=v["method"], page=v["page"],
                            checks="+".join(v["checks"]), total_matched=v["total"] or "",
                            foreign_bond="" if v["外国債券"] is None else v["外国債券"],
                            foreign_equity="" if v["外国株式"] is None else v["外国株式"],
                            foreign_sec="" if v["外国証券"] is None else v["外国証券"],
                            other_sec="" if v["その他の証券"] is None else v["その他の証券"],
                            fb_row=v["row_present"]["外国債券"], fe_row=v["row_present"]["外国株式"],
                            url=v["url"], file=v["file"]))
    with open(DATA / "extracted_foreign.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()) if out else ["fi_code"])
        w.writeheader()
        w.writerows(out)
    codes = {o["fi_code"] for o in out}
    print(f"institutions with ≥1 period matched: {len(codes)} / rows {len(out)}")
    for o in out:
        print(o["fi_code"], o["name"], o["year"], o["method"], "p", o["page"], o["checks"], "FB=", o["foreign_bond"], "FE=", o["foreign_equity"])


if __name__ == "__main__":
    main()
