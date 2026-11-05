"""Batch 1 + 2a: 母集団（Universe）と全銀協 行別財務データ（縦持ち）を CSV に正規化する。

    python3 -I fi_screening/scripts/build_data.py <raw_dir>

出力（fi_screening/data/）:
  universe.csv          1行=1金融機関（主キー: fi_code）
  sources.csv           1行=1資料（主キー: source_id）
  financial_data.csv    1行=1機関×1指標×1期（データ・インテグリティ列をすべて持つ）
  batch_log.csv         バッチ実行記録（追記）
既存の手入力ファイル（policy_evidence.csv / foreign_bond.csv / sales_coverage.csv 等）には触れない。
"""
import csv
import datetime as dt
import html
import json
import pathlib
import re
import sys

RAW = pathlib.Path(sys.argv[1]).resolve()
ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
ACCESSED = "2026-11-05"
UNIVERSE_ASOF = "2026-10-01"

REGION10 = {
    "北海道": ["北海道"],
    "東北": ["青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県"],
    "関東": ["茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県"],
    "甲信越": ["新潟県", "山梨県", "長野県"],
    "北陸": ["富山県", "石川県", "福井県"],
    "東海": ["岐阜県", "静岡県", "愛知県", "三重県"],
    "近畿": ["滋賀県", "京都府", "大阪府", "兵庫県", "奈良県", "和歌山県"],
    "中国": ["鳥取県", "島根県", "岡山県", "広島県", "山口県"],
    "四国": ["徳島県", "香川県", "愛媛県", "高知県"],
    "九州・沖縄": ["福岡県", "佐賀県", "長崎県", "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"],
}
PREF2REGION = {p: r for r, ps in REGION10.items() for p in ps}
PREF_RE = re.compile("(" + "|".join(PREF2REGION) + ")")
# 地銀協の住所に都道府県名が無い行（札幌市）用の地域見出し→都道府県
CITY_PREF = {"札幌市": "北海道"}

REORG_MAP = {
    "0117": ["旧0118みちのく銀行を統合（FY2023まで別行・FY2024から統合後計数）"],
    "0143": ["旧0533長野銀行を統合（FY2024まで別行・FY2025から統合後計数）"],
    "0542": ["旧0544中京銀行を統合（FY2023まで別行・FY2024から統合後計数）"],
    "0147": ["旧0537福邦銀行を統合（FY2025まで別行。統合はFY2025期末後・要確認）"],
}

SHINKIN_REGION_FILES = {
    "hokkaido": "北海道", "tohoku": "東北", "kantou": "関東", "kosinetu": "甲信越",
    "hokuriku": "北陸", "tokai": "東海", "kinki": "近畿", "cyugoku": "中国",
    "sikoku": "四国", "kyusyu": "九州・沖縄",
}

# 全銀協 行別 B/S・P/L から取り込む勘定（Ref. Code → 指標名, 区分）
ZGK_ITEMS = {
    "D990": ("総資産", "A"),
    "A015": ("預金", "A"),
    "A135": ("譲渡性預金", "A"),
    "D395": ("貸出金", "A"),
    "D275": ("有価証券", "B"),
    "D290": ("国債", "B"),
    "D305": ("地方債", "B"),
    "D320": ("短期社債", "B"),
    "D335": ("社債", "B"),
    "D350": ("株式", "B"),
    "D260": ("金銭の信託", "B"),
    "H500": ("その他有価証券評価差額", "B"),
    "B700": ("その他有価証券評価差額金", "B"),
    "B750": ("繰延ヘッジ損益", "B"),
    "B950": ("純資産", "A"),
    "E040": ("有価証券利息配当金", "B"),
    "E230": ("国債等債券売却益", "B"),
    "E240": ("国債等債券償還益", "B"),
    "E610": ("国債等債券売却損", "B"),
    "E620": ("国債等債券償還損", "B"),
    "E630": ("国債等債券償却", "B"),
}
ZGK_SHEETS = {"地方銀行１（単体）": "地方銀行", "地方銀行２（単体）": "地方銀行", "地方銀行Ⅱ（単体）": "第二地方銀行"}
FY_OF = {"2025": "FY2025(2026/3期)", "2024": "FY2024(2025/3期)", "2023": "FY2023(2024/3期)"}


def text(s):
    return html.unescape(re.sub(r"<[^>]+>|\s", "", s))


def parse_chigin():
    s = (RAW / "chigin_list.html").read_text(encoding="utf-8", errors="replace")
    out = []
    for m in re.finditer(r"<dl>(.*?)</dl>", s, re.S):
        dt_ = re.search(r"<dt>(.*?)</dt>", m.group(1), re.S)
        dds = re.findall(r"<dd>(.*?)</dd>", m.group(1), re.S)
        if not dt_ or not dds:
            continue
        name = text(dt_.group(1))
        if not name.endswith("銀行"):
            continue
        addr = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", dds[0]))).strip()
        url = re.search(r'href="([^"]+)"', dt_.group(1))
        pm = PREF_RE.search(addr)
        pref = pm.group(1) if pm else next((p for c, p in CITY_PREF.items() if c in addr), "")
        out.append(dict(name=name, addr=addr, pref=pref, url=url.group(1) if url else ""))
    return out


def parse_dainichi():
    s = (RAW / "dainichi_list.html").read_text(encoding="utf-8", errors="replace")
    out = []
    for m in re.finditer(r"<tr.*?</tr>", s, re.S):
        tds = re.findall(r"<td[^>]*>(.*?)</td>", m.group(0), re.S)
        if len(tds) < 4:
            continue
        name = text(tds[0])
        if not name.endswith("銀行"):
            continue
        addr = re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", tds[3]))).strip()
        url = re.search(r'href="([^"]+)"', tds[0])
        pm = PREF_RE.search(addr)
        out.append(dict(name=name, addr=addr, pref=pm.group(1) if pm else "", url=url.group(1) if url else ""))
    return out


def parse_shinkin():
    out = []
    for f, region in SHINKIN_REGION_FILES.items():
        s = (RAW / f"shinkin_{f}.htm").read_text(encoding="utf-8-sig", errors="replace")
        pref = None
        for m in re.finditer(r"<tr.*?</tr>", s, re.S):
            row = m.group(0)
            th = re.findall(r"<th[^>]*>(.*?)</th>", row, re.S)
            if len(th) == 1 and PREF_RE.fullmatch(text(th[0])):
                pref = text(th[0])
                continue
            tds = [text(x) for x in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
            if tds and re.fullmatch(r"\d{4}", tds[0]):
                url = re.search(r'href="([^"]+)"', row)
                out.append(dict(code=tds[0], name=tds[1], kana=tds[2] if len(tds) > 2 else "",
                                pref=pref, region=region, url=url.group(1) if url else ""))
    return out


def load_zgk(year):
    """全銀協 行別財務諸表（単体）を {code: {name, category, items{ref: value}}} で返す。"""
    p_xlsx, p_xls = RAW / f"zgk_kobetsu_{year}.xlsx", RAW / f"zgk_kobetsu_{year}.xls"
    sheets = {}
    if p_xlsx.exists():
        import openpyxl
        wb = openpyxl.load_workbook(p_xlsx, read_only=True, data_only=True)
        for sn in ZGK_SHEETS:
            sheets[sn] = [list(r) for r in wb[sn].iter_rows(values_only=True)]
    else:
        import xlrd
        wb = xlrd.open_workbook(str(p_xls))
        for sn in ZGK_SHEETS:
            sh = wb.sheet_by_name(sn)
            sheets[sn] = [sh.row_values(i) for i in range(sh.nrows)]
    res = {}
    for sn, cat in ZGK_SHEETS.items():
        rows = sheets[sn]
        hdr = rows[4]
        banks = []
        for ci, v in enumerate(hdr):
            m = re.match(r"^(\d{4})[\s　]+(.+)$", str(v or "").strip())
            if m:
                banks.append((ci, m.group(1), m.group(2).strip()))
        for ci, code, name in banks:
            items = {}
            for r in rows:
                ref = str(r[1] or "").strip()
                if ref in ZGK_ITEMS:
                    v = r[ci + 1] if ci + 1 < len(r) else None
                    if v in ("－", "-", "", None):
                        v = 0 if v in ("－", "-") else None
                    items[ref] = float(v) if v is not None else None
            res[code] = dict(name=name, category=cat, sheet=sn, col=ci + 1, items=items)
    return res


def colname(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    DATA.mkdir(parents=True, exist_ok=True)
    chigin, dainichi, shinkin = parse_chigin(), parse_dainichi(), parse_shinkin()
    zgk = {y: load_zgk(y) for y in FY_OF}

    sources = [
        dict(source_id="S-CHIGIN-LIST", fi_code="ALL", doc_type="協会会員名簿", title="全国地方銀行協会 地方銀行一覧",
             period="現在", publisher="全国地方銀行協会", url="https://www.chiginkyo.or.jp/regional_banks/list/",
             published_date="", accessed_date=ACCESSED, tier="Tier2", basis="", notes=f"{len(chigin)}行掲載"),
        dict(source_id="S-DAINICHI-LIST", fi_code="ALL", doc_type="協会会員名簿", title="第二地方銀行協会 加盟地方銀行一覧",
             period="現在", publisher="第二地方銀行協会", url="https://www.dainichiginkyo.or.jp/membership/member_list.html",
             published_date="", accessed_date=ACCESSED, tier="Tier2", basis="", notes=f"{len(dainichi)}行掲載"),
        dict(source_id="S-ZENSHINKYO-LIST", fi_code="ALL", doc_type="協会会員名簿", title="全国信用金庫協会 全国の信用金庫のご紹介（地区別）",
             period="現在", publisher="全国信用金庫協会", url="https://www.shinkin.co.jp/tikubetu/",
             published_date="", accessed_date=ACCESSED, tier="Tier2", basis="", notes=f"{len(shinkin)}金庫掲載（金融機関コード付）"),
    ]
    zgk_pub = {"2025": "2026-08-07", "2024": "2025-08-01", "2023": "2024-08-02"}
    for y in FY_OF:
        sources.append(dict(
            source_id=f"S-ZGK-KOBETSU-{y}", fi_code="ALL", doc_type="業界団体統計",
            title=f"全国銀行協会 全国銀行財務諸表分析 {y}年度決算 貸借対照表・損益計算書（行別・単体）",
            period=FY_OF[y], publisher="全国銀行協会",
            url=(RAW / f"zgk_kobetsu_{y}.url").read_text().strip(),
            published_date=zgk_pub[y], accessed_date=ACCESSED, tier="Tier2", basis="単体",
            notes="単位: 百万円。行別の値は各行公表の財務諸表を全銀協が様式統一したもの"))

    # ---------- Universe ----------
    uni = []
    z25 = zgk["2025"]
    name2code = {v["name"]: k for k, v in z25.items()}
    for lst, cat, assoc, src in [(chigin, "地方銀行", "全国地方銀行協会", "S-CHIGIN-LIST"),
                                 (dainichi, "第二地方銀行", "第二地方銀行協会", "S-DAINICHI-LIST")]:
        for b in lst:
            code = name2code.get(b["name"], "")
            zc = z25.get(code, {}).get("category")
            note = ""
            if not code:
                note = "全銀協FY2025行別データに同名行なし（要確認）"
            elif zc != cat:
                note = f"全銀協の業態区分={zc}（協会所属と不一致・要確認）"
            uni.append(dict(
                fi_code=code, legal_name=f"株式会社{b['name']}", short_name=b["name"], category=cat,
                size_band="", association=assoc, hq_pref=b["pref"], region=PREF2REGION.get(b["pref"], ""),
                hq_address=b["addr"], website=b["url"], corporate_number="", corporate_number_status="NOT YET CHECKED",
                holdco="", holdco_status="NOT YET CHECKED", reorg_history="", reorg_planned="",
                universe_asof=UNIVERSE_ASOF, code_source="S-ZGK-KOBETSU-2025" if code else "",
                member_source=src, research_status="Batch2a完了（全銀協基礎計数）", note=note, continuity_break=""))
    for s in shinkin:
        uni.append(dict(
            fi_code=s["code"], legal_name=f"{s['name']}信用金庫", short_name=s["name"], category="信用金庫",
            size_band="（預金残高取得後に判定）", association="全国信用金庫協会", hq_pref=s["pref"] or "",
            region=SHINKIN_REGION_FILES_REGION(s), hq_address="", website=s["url"], corporate_number="",
            corporate_number_status="NOT YET CHECKED", holdco="", holdco_status="該当なし（協同組織）",
            reorg_history="", reorg_planned="", universe_asof=UNIVERSE_ASOF, code_source="S-ZENSHINKYO-LIST",
            member_source="S-ZENSHINKYO-LIST", research_status="Batch1完了（母集団のみ）", note="", continuity_break=""))
    uni.append(dict(
        fi_code="1000", legal_name="信金中央金庫", short_name="信金中金", category="信金中央金庫", size_band="—",
        association="全国信用金庫協会", hq_pref="東京都", region="関東", hq_address="東京都中央区八重洲1-3-7",
        website="https://www.shinkin-central-bank.jp/", corporate_number="", corporate_number_status="NOT YET CHECKED",
        holdco="", holdco_status="該当なし", reorg_history="", reorg_planned="", universe_asof=UNIVERSE_ASOF,
        code_source="金融機関コード1000（全銀協統一コード。要一次確認）", member_source="",
        research_status="Batch1完了（母集団のみ）", note="個別信用金庫とは別の投資主体として扱う。業態内ランキング対象外", continuity_break=""))

    # 再編による非連続：FY2023/FY2024 に存在し FY2025 に無いコード、FY2025 に存在し現協会名簿に無いコード
    for y in ("2024", "2023"):
        for code, v in zgk[y].items():
            if code not in z25:
                sources_note.append(f"{FY_OF[y]}の旧行 {code} {v['name']} はFY2025の行別データに存在しない（合併等）")
    listed = {u["fi_code"] for u in uni}
    for code, v in z25.items():
        if code not in listed:
            sources_note.append(f"FY2025行別データの {code} {v['name']}（{v['category']}）は現協会名簿に無い（基準日までの合併等）")
    # 合併の存続行へのひも付け（全銀協行別データでの旧行コード消滅と現名簿・名称から判定。効力発生日は各行の適時開示で要確認）
    for survivor, absorbed in REORG_MAP.items():
        for u in uni:
            if u["fi_code"] == survivor:
                u["reorg_history"] += "; ".join(absorbed) + "; "
                u["continuity_break"] = "あり（合併前の期は存続行単体の計数。前期比は非連続）"
    # 名称変化（同一コードで名称が異なる）を検出
    for code, v in z25.items():
        for y in ("2024", "2023"):
            old = zgk[y].get(code)
            if old and old["name"] != v["name"]:
                for u in uni:
                    if u["fi_code"] == code:
                        u["reorg_history"] += f"{FY_OF[y]}名称={old['name']}; "

    uni.sort(key=lambda u: ({"地方銀行": 0, "第二地方銀行": 1, "信金中央金庫": 2, "信用金庫": 3}[u["category"]], u["fi_code"]))
    write_csv(DATA / "universe.csv", uni)

    # ---------- Financial data (long) ----------
    fin = []
    uni_codes = {u["fi_code"] for u in uni}
    for y, banks in zgk.items():
        for code, b in banks.items():
            if code not in uni_codes:
                continue
            for ref, (metric, grp) in ZGK_ITEMS.items():
                v = b["items"].get(ref)
                fin.append(dict(
                    fi_code=code, fi_name=b["name"], metric=metric, metric_group=grp, value="" if v is None else v,
                    unit="百万円", period=FY_OF[y], basis="単体", source_id=f"S-ZGK-KOBETSU-{y}",
                    page=f"シート「{b['sheet']}」列{colname(b['col'])}／Ref.{ref}",
                    status="SECONDARY" if v is not None else "NOT YET CHECKED",
                    accessed_date=ACCESSED, formula="",
                    note="全銀協が各行公表の単体財務諸表を集計（Tier2）。各行の一次資料との突合はBatch2bで実施"))
            it = b["items"]
            parts = [it.get(k) for k in ("D290", "D305", "D320", "D335", "D350")]
            if it.get("D275") is not None and all(p is not None for p in parts):
                fin.append(dict(
                    fi_code=code, fi_name=b["name"], metric="その他の証券（推計・外国証券等を含む上位集合）", metric_group="B",
                    value=it["D275"] - sum(parts), unit="百万円", period=FY_OF[y], basis="単体",
                    source_id=f"S-ZGK-KOBETSU-{y}", page=f"シート「{b['sheet']}」列{colname(b['col'])}",
                    status="ESTIMATED", accessed_date=ACCESSED,
                    formula="有価証券(D275)−国債(D290)−地方債(D305)−短期社債(D320)−社債(D335)−株式(D350)",
                    note="外国証券・投資信託・その他を含む。外国証券残高とはみなさない"))
    write_csv(DATA / "financial_data.csv", fin)
    write_csv(DATA / "sources.csv", sources)

    # ---------- Batch log ----------
    n_bank = sum(1 for u in uni if u["category"] in ("地方銀行", "第二地方銀行"))
    n_bank_ok = sum(1 for u in uni if u["category"] in ("地方銀行", "第二地方銀行") and u["fi_code"] in z25)
    log = [
        dict(batch="Batch1", run_date=ACCESSED, scope="母集団（全業態）", target=len(uni), completed=len(uni),
             not_started=0, data_ok=len(uni), data_short=sum(1 for u in uni if u["note"].startswith("全銀協")),
             next_target="法人番号の付与（国税庁法人番号公表サイト）・持株会社ひも付け",
             note=f"地銀{len(chigin)}・第二地銀{len(dainichi)}・信金{len(shinkin)}・信金中金1"),
        dict(batch="Batch2a", run_date=ACCESSED, scope="地銀・第二地銀 基礎計数（全銀協 行別単体 FY2023-FY2025）",
             target=n_bank, completed=n_bank_ok, not_started=n_bank - n_bank_ok, data_ok=n_bank_ok,
             data_short=n_bank - n_bank_ok,
             next_target="Batch2b: 各行ディスクロージャー誌・決算説明資料で外国証券残高・自己資本比率・外債方針を取得",
             note="外国証券は全銀協様式に区分がないため未取得（その他の証券を推計値として保持）"),
    ]
    write_csv(DATA / "batch_log.csv", log)
    print(f"universe={len(uni)} (地銀{len(chigin)} 第二{len(dainichi)} 信金{len(shinkin)} +信金中金) "
          f"fin_rows={len(fin)} unmatched={[u['short_name'] for u in uni if u['note']]}")
    for n in sources_note:
        print("REORG:", n)


sources_note = []


def SHINKIN_REGION_FILES_REGION(s):
    return PREF2REGION.get(s["pref"] or "", s["region"])


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)


if __name__ == "__main__":
    main()
