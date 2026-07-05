# -*- coding: utf-8 -*-
"""JPX公式の上場銘柄一覧(data_j.xls)から銘柄辞書を再生成する。

使い方:
  python tools/update_stocks.py            # ダウンロードして stocks.js と app.html を更新
  python tools/update_stocks.py path.xls   # ダウンロード済みのxlsを使う

更新対象(二重実装のため必ず両方):
  - stocks.js               (window.STOCKS=... 全体を書き換え)
  - app.html                (インラインの window.STOCKS={...}; 行を書き換え)

JPXは毎月末時点の一覧を翌月頭に公開する。月1回これを回せば新規上場(186A等の
英字コード含む)に追従できる。PRO Market(一般個人が売買不可)は除外する。
"""
import json, re, sys, urllib.request
from pathlib import Path

JPX_URL = "https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls"
EXCLUDE = {"PRO Market"}
ROOT = Path(__file__).resolve().parent.parent


def load_rows(xls_path):
    import xlrd  # pip install xlrd
    sh = xlrd.open_workbook(xls_path).sheet_by_index(0)
    date = int(sh.cell_value(1, 0))
    stocks = {}
    for r in range(1, sh.nrows):
        code, name, cat = sh.cell_value(r, 1), sh.cell_value(r, 2), sh.cell_value(r, 3)
        if cat in EXCLUDE:
            continue
        code = str(int(code)).zfill(4) if isinstance(code, float) else str(code).strip().upper()
        stocks[code] = str(name).replace("　", " ").strip()
    return date, stocks


def main():
    if len(sys.argv) > 1:
        xls = sys.argv[1]
    else:
        xls = str(ROOT / "tools" / "data_j.xls")
        print(f"downloading {JPX_URL} ...")
        req = urllib.request.Request(JPX_URL, headers={"User-Agent": "Mozilla/5.0"})
        Path(xls).write_bytes(urllib.request.urlopen(req, timeout=60).read())

    date, stocks = load_rows(xls)
    js = "window.STOCKS=" + json.dumps(stocks, ensure_ascii=False, separators=(",", ":")) + ";"
    print(f"JPX {date} 時点 / {len(stocks)} 銘柄 (英字コード {sum(1 for c in stocks if not c.isdigit())} 件)")

    (ROOT / "stocks.js").write_text(js + "\n", encoding="utf-8")
    print("stocks.js updated")

    app = ROOT / "app.html"
    html = app.read_text(encoding="utf-8")
    new_html, n = re.subn(r"^window\.STOCKS=\{.*?\};?$", js, html, count=1, flags=re.M)
    if n != 1:
        sys.exit("ERROR: app.html 内の window.STOCKS 行が見つからない/複数ある")
    app.write_text(new_html, encoding="utf-8")
    print("app.html updated")


if __name__ == "__main__":
    main()
