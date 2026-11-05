"""Phase 4-5: 取得済み PDF テキストから運用方針に関する記述（原文）を抽出し、証拠候補として保存する。

    python3 -I fi_screening/scripts/extract_policy.py <raw_dir> [code ...]

出力: data/policy_candidates.csv（原文抜粋・ページ・URL・カテゴリ）。
これは「候補」であり、分析者が原文を確認して 6_Investment Policy Evidence（data/policy_evidence.csv）へ採用する。
"""
import csv
import json
import pathlib
import re
import sys

RAW = pathlib.Path(sys.argv[1]).resolve()
ONLY = set(sys.argv[2:])
DATA = pathlib.Path(__file__).resolve().parents[1] / "data"

CATS = [
    ("為替ヘッジ方針", r"為替ヘッジ|ヘッジ付|ヘッジ外債|ヘッジなし|オープン外債|為替リスク(?:を|の)?(?:ヘッジ|回避)"),
    ("ヘッジコスト", r"ヘッジコスト|ヘッジ費用|調達コスト.{0,10}外貨|外貨調達"),
    ("外債方針", r"外債|外国債券|外国証券|米国債|海外債券|外貨建(?:債券|資産|有価証券)"),
    ("米国金利見通し", r"米国(?:の)?(?:長期)?金利|米金利|FRB|ＦＲＢ|米国債利回り"),
    ("ポートフォリオ再構築", r"ポートフォリオ(?:の)?(?:再構築|リバランス|入替|入れ替え|見直し)|損切り|含み損(?:の)?処理|評価損(?:の)?処理|入替(?:え)?売買|ロスカット"),
    ("投信・ファンド活用", r"投資信託(?:を|の)?(?:活用|積み増し|運用)|ファンド(?:投資|運用|を活用)|オルタナティブ|プライベート|私募投信|外部運用|運用委託|投資一任"),
    ("債券リスク管理", r"金利リスク(?:量|管理)|ＢＰＶ|BPV|ベーシス・?ポイント・?バリュー|デュレーション|VaR|ＶａＲ|アウトライヤー|IRRBB|ＩＲＲＢＢ"),
]
SENT = re.compile(r"[^。\n]{0,220}(?:。|\n)")


def norm(t):
    # 段組で分断された行を軽く連結（行末の日本語→次行頭の日本語）
    t = re.sub(r"(?<=[ぁ-んァ-ヶ一-龠、])\n\s*(?=[ぁ-んァ-ヶ一-龠])", "", t)
    return re.sub(r"[ \t　]{2,}", " ", t)


def main():
    rows = []
    for d in sorted((RAW / "disc").iterdir()):
        code = d.name
        if (ONLY and code not in ONLY) or not (d / "result.json").exists():
            continue
        res = json.loads((d / "result.json").read_text())
        seen = set()
        for c in res.get("checked", []):
            tx = d / c["file"].replace(".pdf", ".txt")
            if not tx.exists():
                continue
            for pi, pg in enumerate(tx.read_text(encoding="utf-8", errors="replace").split("\f"), 1):
                pg = norm(pg)
                for sm in SENT.finditer(pg):
                    sent = sm.group(0).strip()
                    if len(sent) < 25 or len(re.findall(r"[\d,]{4,}", sent)) > 3:  # 表の行は除外
                        continue
                    if re.search(r"会計|換算|引当|注記|決算日|ヘッジ会計|繰延ヘッジ|とは、|定義|算出方法|計測方法|監査|－ －|− −", sent):
                        continue  # 会計方針・用語定義などの定型文は除外
                    for cat, rx in CATS:
                        if re.search(rx, sent):
                            key = (cat, sent[:60])
                            if key in seen:
                                break
                            seen.add(key)
                            rows.append(dict(fi_code=code, fi_name=res["name"], category=cat, quote=sent[:300],
                                             url=c["url"], page=pi, file=c["file"]))
                            break
    with open(DATA / "policy_candidates.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["fi_code", "fi_name", "category", "quote", "url", "page", "file"])
        w.writeheader()
        w.writerows(rows)
    from collections import Counter
    print(len(rows), Counter(r["category"] for r in rows))


if __name__ == "__main__":
    main()
