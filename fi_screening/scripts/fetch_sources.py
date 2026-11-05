"""Batch 0/1/2a: 公式ソースの取得。

取得物はすべて引数で指定した raw ディレクトリに保存する（リポジトリには入れない）。
    python3 -I fi_screening/scripts/fetch_sources.py <raw_dir>
"""
import pathlib
import re
import subprocess
import sys
import zipfile

RAW = pathlib.Path(sys.argv[1]).resolve()

PAGES = {
    "chigin_list.html": "https://www.chiginkyo.or.jp/regional_banks/list/",
    "dainichi_list.html": "https://www.dainichiginkyo.or.jp/membership/member_list.html",
}
SHINKIN_REGIONS = ["hokkaido", "tohoku", "kantou", "kosinetu", "hokuriku",
                   "tokai", "kinki", "cyugoku", "sikoku", "kyusyu"]
for r in SHINKIN_REGIONS:
    PAGES[f"shinkin_{r}.htm"] = f"https://www.shinkin.co.jp/tikubetu/{r}.htm"
ZGK_YEARS = ["2025", "2024", "2023"]
for y in ZGK_YEARS:
    PAGES[f"zgk_fs_{y}.html"] = f"https://www.zenginkyo.or.jp/stats/year2-02/account{y}-terminal/"


def get(url, dest):
    dest.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["curl", "-sSL", "--fail", "-m", "180", url, "-o", str(dest)], check=True)


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    for name, url in PAGES.items():
        get(url, RAW / name)
        print("ok", name)
    for y in ZGK_YEARS:
        page = (RAW / f"zgk_fs_{y}.html").read_text(encoding="utf-8", errors="replace")
        m = re.search(r'/fileadmin/[^"]*kobetsu\d+\.zip', page)
        if not m:
            raise SystemExit(f"kobetsu zip not found for {y}")
        z = RAW / f"zgk_kobetsu_{y}.zip"
        get("https://www.zenginkyo.or.jp" + m.group(0), z)
        with zipfile.ZipFile(z) as zf:
            xl = [n for n in zf.namelist() if n.lower().endswith((".xlsx", ".xls"))]
            ext = ".xlsx" if xl[0].lower().endswith(".xlsx") else ".xls"
            (RAW / f"zgk_kobetsu_{y}{ext}").write_bytes(zf.read(xl[0]))
        (RAW / f"zgk_kobetsu_{y}.url").write_text("https://www.zenginkyo.or.jp" + m.group(0))
        print("ok kobetsu", y)


if __name__ == "__main__":
    main()
