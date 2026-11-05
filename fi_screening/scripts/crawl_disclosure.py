"""Batch 2b/3+: 各金融機関の公式サイトからディスクロージャー誌（PDF）候補を探索・取得し、
「有価証券残高」表の外国債券・外国株式・外国証券を抽出する。

    python3 -I fi_screening/scripts/crawl_disclosure.py <raw_dir> <category> [code ...]

    category: 地方銀行 / 第二地方銀行 / 信用金庫 / 信金中央金庫 / all

出力: <raw_dir>/disc/<code>/ に PDF とテキスト、<raw_dir>/disc/<code>/result.json に抽出結果。
抽出値は必ず人手（または二次チェック）で表の所在ページを確認してから data/manual_financial.csv に転記する。
（本スクリプトはあくまで候補探索と一次抽出の補助）
"""
import concurrent.futures as cf
import csv
import hashlib
import html
import json
import pathlib
import re
import subprocess
import sys
import urllib.parse as up

RAW = pathlib.Path(sys.argv[1]).resolve()
CAT = sys.argv[2]
ONLY = set(sys.argv[3:])
ROOT = pathlib.Path(__file__).resolve().parents[1]
OUT = RAW / "disc"
MAXDL = 6
MAXPAGES = 50
UA = "Mozilla/5.0 (research crawler; financial-institution disclosure survey)"

NAV_KW = re.compile(r"ディスクロ|disclosure|IR|ＩＲ|投資家|株主|決算|財務|経営情報|会社情報|企業情報|銀行概要|当行について|当金庫について|"
                    r"金庫について|銀行について|会社案内|about|corporate|ir/|/ir|kinko|zaimu|keiei|library|ライブラリ|経営内容|開示", re.I)
PDF_KW = re.compile(r"ディスクロ|disclosure|disclo|資料編|データ編|財務|決算|dis|data|annual|統合報告|ミニディスク|現況|経営内容|事業の概況", re.I)
YEAR_KW = re.compile(r"2026|2025|令和[78]|R[78]|r0?[78]")


def fetch(url, dest=None, timeout=40, tries=3):
    """curl で取得。ネットワーク瞬断（DNS失敗・トンネル切断）に備えて指数バックオフで再試行。"""
    import time
    cmd = ["curl", "-sSL", "--fail", "-m", str(timeout), "-A", UA, "--max-filesize", "80000000", url]
    if dest:
        cmd += ["-o", str(dest)]
    for i in range(tries):
        r = subprocess.run(cmd, capture_output=True)
        if r.returncode == 0:
            break
        if r.returncode == 22:  # HTTP 4xx/5xx は再試行しない
            return False if dest else None
        time.sleep(2 * (2 ** i))
    else:
        return False if dest else None
    if dest:
        return True
    b = r.stdout
    for enc in ("utf-8", "cp932", "euc_jp"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", "replace")


def render(url):
    """headless Chromium で描画後の HTML を取得（scripts/render.js）。失敗時 None。"""
    js = pathlib.Path(__file__).with_name("render.js")
    npm_root = subprocess.run(["npm", "root", "-g"], capture_output=True, text=True).stdout.strip()
    r = subprocess.run(["node", str(js), url], capture_output=True, timeout=120,
                       env={**__import__("os").environ, "NODE_PATH": npm_root})
    return r.stdout.decode("utf-8", "replace") if r.returncode == 0 and r.stdout else None


def links(base, page):
    for m in re.finditer(r'<a\b[^>]*href="([^"#]+)"[^>]*>(.*?)</a>', page, re.S | re.I):
        href = html.unescape(m.group(1)).strip()
        if href.startswith(("mailto:", "tel:", "javascript:")):
            continue
        txt = re.sub(r"<[^>]+>|\s+", "", html.unescape(m.group(2)))
        alt = " ".join(re.findall(r'alt="([^"]*)"', m.group(2)))
        yield up.urljoin(base, href), (txt + " " + alt).strip()


def site_key(u):
    h = up.urlparse(u).netloc.lower()
    parts = h.split(".")
    return ".".join(parts[-3:]) if h.endswith((".co.jp", ".or.jp", ".ne.jp")) else ".".join(parts[-2:])


def load_seeds():
    p = ROOT / "data" / "disclosure_seeds.csv"
    seeds = {}
    if p.exists():
        with open(p, encoding="utf-8") as f:
            for r in csv.DictReader(f):
                seeds.setdefault(r["fi_code"], []).append(r["url"])
    return seeds


SEEDS = None


def crawl(code, home):
    """ホーム（＋data/disclosure_seeds.csv の起点URL）から最大 depth3 で探索し PDF 候補を返す。"""
    global SEEDS
    if SEEDS is None:
        SEEDS = load_seeds()
    seeds = SEEDS.get(code, [])
    direct = [dict(url=u, text="seed ディスクロージャー", found_on="seed") for u in seeds if u.lower().split("?")[0].endswith(".pdf")]
    pages = [u for u in seeds if not u.lower().split("?")[0].endswith(".pdf")]
    if pages:
        home = pages[0]
    if not home:
        return direct
    home_key = site_key(home)
    path_prefix = ""
    if "shinkin.co.jp" in home:  # 共同ドメイン配下の金庫サイトはパス接頭辞で範囲を限定
        path_prefix = "/" + up.urlparse(home).path.strip("/").split("/")[0] + "/"
    seen, queue, pdfs = set(), [(u, 1 if u != home else 0, "ディスクロ seed") for u in ([home] + pages[1:])], {}
    for dct in direct:
        pdfs.setdefault(dct["url"], (dct["text"], "seed"))
    while queue and len(seen) < MAXPAGES:
        url, d, via = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        page = fetch(url, timeout=15, tries=2)
        if not page:
            continue
        if re.search(r"disclo|ディスクロ", url + via, re.I):
            via = via + " ディスクロ"
        lk = list(links(url, page))
        if d <= 1 and re.search(r"disclo|ディスクロ|ir/|library", url + via, re.I) and not any(
                up.urlparse(u).path.lower().endswith(".pdf") for u, _ in lk):
            rp = render(url)  # JS で一覧を描画する IR ページ（IRPOCKET 等）
            if rp:
                lk = list(links(url, rp))
        for u, t in lk:
            p = up.urlparse(u)
            if p.scheme not in ("http", "https"):
                continue
            same = site_key(u) == home_key and (not path_prefix or p.path.startswith(path_prefix))
            if p.path.lower().endswith(".pdf"):
                if PDF_KW.search(t + " " + p.path) or "ディスクロ" in via:
                    pdfs.setdefault(u, (t, url))
                continue
            strong = re.search(r"ディスクロ|IR情報|ＩＲ情報|投資家|株主・投資家|IR$|^IR|決算", t)
            cross_ok = (not same) and strong and d <= 2 and p.netloc.endswith(".jp") and "shinkin.co.jp" not in p.netloc
            if d < 3 and (NAV_KW.search(t) or NAV_KW.search(p.path)) and (same or cross_ok or "ディスクロ" in t):
                queue.append((u.split("#")[0], d + 1, t))
    return [dict(url=u, text=t, found_on=f) for u, (t, f) in pdfs.items()]


def max_year(blob):
    ys = [int(y) for y in re.findall(r"(20[12]\d)", blob)]
    ys += [2018 + int(y) for y in re.findall(r"令和\s*(\d)", blob)]
    ys += [2018 + int(y) for y in re.findall(r"(?<![a-z])[rR]0?(\d)(?!\d)", blob)]
    return max(ys) if ys else 0


def rank(c):
    s = 0
    blob = c["text"] + " " + c["url"]
    if re.search(r"ディスクロ|disclo", blob, re.I):
        s += 5
    y = max_year(blob)
    s += {2026: 6, 2025: 4}.get(y, -6 if 2000 < y < 2024 else (-2 if y == 2024 else 0))
    if re.search(r"資料編|データ編|財務|data", blob, re.I):
        s += 2
    if re.search(r"中間|半期|interim|mini|ミニ|_h\.|hosoku|補足|アンケート|questionnaire", blob, re.I):
        s -= 3
    return s


NUM = r"[△▲\-－]?[\d,，]+|[―－\-–—]"


def to_num(x):
    x = x.replace(",", "").replace("，", "")
    if re.fullmatch(r"[―－\-–—]", x):
        return 0.0
    neg = x[:1] in "△▲-－"
    x = x.lstrip("△▲-－")
    try:
        v = float(x)
    except ValueError:
        return None
    return -v if neg else v


def extract(txt):
    """「有価証券残高」等の表から 期末残高 の うち外国債券/外国株式/外国証券 の行を抽出（ページ番号付き）。"""
    hits = []
    pages = txt.split("\f")
    for pi, pg in enumerate(pages, 1):
        if not re.search(r"有価証券", pg):
            continue
        for m in re.finditer(r"^.*(うち外国債券|うち外国株式|外国債券|外国株式|外国証券).*$", pg, re.M):
            line = m.group(0)
            label = m.group(1)
            nums = [to_num(x) for x in re.findall(NUM, line.split(label, 1)[1])]
            nums = [n for n in nums if n is not None]
            if not nums:
                continue
            ctx_before = pg[max(0, m.start() - 1500):m.start()]
            hits.append(dict(page=pi, label=label, nums=nums, line=line.strip()[:200],
                             table_hint="期末残高" if re.search(r"期末残高|有価証券残高|種類別", ctx_before) and not re.search(r"平均残高[^\n]*\n(?:(?!期末).)*$", ctx_before, re.S) else "",
                             ctx_head=re.sub(r"\s+", " ", ctx_before[-300:])))
    return hits


def process(u):
    code, home = u["fi_code"], u["website"]
    d = OUT / code
    if (d / "result.json").exists() and not __import__("os").environ.get("FORCE"):
        r = json.loads((d / "result.json").read_text())
        return code, u["short_name"], r["n_candidates"], len(r["checked"]), any(g.get("has_foreign") for g in r["checked"])
    d.mkdir(parents=True, exist_ok=True)
    cands = crawl(code, home)
    nm = u["short_name"].replace("銀行", "")
    for c in cands:
        c["_bonus"] = 2 if nm and nm in c["text"] else 0
    cands.sort(key=lambda c: rank(c) + c["_bonus"], reverse=True)
    # 分冊形式（同一フォルダに章別PDF）の場合は、最上位候補と同じフォルダの PDF をすべて優先取得
    if cands:
        top_dir = cands[0]["url"].rsplit("/", 1)[0] + "/"
        same = [c for c in cands if c["url"].startswith(top_dir)]
        stem = re.sub(r"[_\-]?[\d\-_]*(sup)?[\d\-_]*\.pdf$", "", cands[0]["url"])
        sib = [c for c in cands if c["url"].startswith(stem)]
        if 3 <= len(sib) < len(same):
            same = sib
        if len(same) >= 3 and len(same) <= 60:
            rest = [c for c in cands if not c["url"].startswith(top_dir)]
            cands = same[:40] + rest
    maxdl = max(MAXDL, min(30, sum(1 for c in cands if cands and c["url"].startswith(cands[0]["url"].rsplit("/", 1)[0] + "/"))))
    got = []
    target = re.compile(r"(残存期間|種類別|有価証券残高)[\s\S]{0,6000}(外国債券|外国証券)")
    hit = False
    for i, c in enumerate(cands[:max(maxdl, 12)]):
        if i >= maxdl and hit:
            break
        if rank(c) < 0 and got and hit:
            break
        fn = d / (hashlib.md5(c["url"].encode()).hexdigest()[:10] + ".pdf")
        if not fn.exists() and not fetch(c["url"], fn, timeout=150):
            continue
        tx = fn.with_suffix(".txt")
        if not tx.exists():
            subprocess.run(["pdftotext", "-layout", str(fn), str(tx)], capture_output=True)
        if not tx.exists():
            continue
        t = tx.read_text(encoding="utf-8", errors="replace")
        got.append(dict(url=c["url"], text=c["text"], found_on=c["found_on"], file=fn.name, rank=rank(c),
                        n_pages=t.count("\f") + 1, has_foreign=bool(re.search(r"外国債券|外国証券", t))))
        hit = hit or bool(target.search(t))
    res = dict(fi_code=code, name=u["short_name"], website=home, n_candidates=len(cands),
               candidates=[c["url"] for c in cands[:40]], checked=got)
    (d / "result.json").write_text(json.dumps(res, ensure_ascii=False, indent=1))
    ok = any(g["has_foreign"] for g in got)
    return code, u["short_name"], len(cands), len(got), ok


def main():
    with open(ROOT / "data" / "universe.csv", encoding="utf-8") as f:
        uni = [r for r in csv.DictReader(f) if (CAT == "all" or r["category"] == CAT) and (not ONLY or r["fi_code"] in ONLY)]
    OUT.mkdir(parents=True, exist_ok=True)
    with cf.ThreadPoolExecutor(10) as ex:
        for code, name, nc, ng, ok in ex.map(process, uni):
            print(f"{code} {name}: candidates={nc} checked={ng} foreign_hit={'Y' if ok else '-'}", flush=True)


if __name__ == "__main__":
    main()
