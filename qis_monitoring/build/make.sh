#!/usr/bin/env bash
# サンプル CSV → ブック生成 → LibreOffice 再計算で検証 → テスト結果を埋め込んで再生成 → ドキュメント生成
set -euo pipefail
cd "$(dirname "$0")"
python3 sample_data.py
python3 build_workbook.py --results /nonexistent
python3 verify_workbook.py || echo "!! verify reported failures (see docs/07_test_results.md)"
python3 build_workbook.py            # dist/test_results.json を 22_Test_Results に埋め込む
python3 gen_docs.py
