#!/usr/bin/env python3
"""コードと同期させる必要があるドキュメントを生成する（docs/05, docs/08）。"""
from pathlib import Path

import build_workbook as B
from preprod import PREPROD

ROOT = Path(__file__).resolve().parent.parent
DOCS = ROOT / "docs"


def esc(x):
    return str(x).replace("|", "\\|").replace("\n", " ")


def alerts_md():
    pv = {k: v for k, v, _, _ in B.PARAMS}
    out = ["# リスクアラート設定表", "",
           "すべて `19_Settings` で変更可能（A 表＝閾値パラメータ、B 表＝ルール）。初期値は仮の値であり、商品特性に合わせて調整すること。",
           "判定は `13_Risk_Alerts` の C 表（Fund-Index Mapping 行 × ルール）で行い、発生したものだけを A 表に重要度順で並べる。", "",
           "- **Scope=FUND**：ファンド単位（各ファンドの最初の Mapping 行でのみ評価し、重複させない）",
           "- **Scope=INDEX**：参照指数単位（1 ファンドが複数指数を持つ場合は指数ごとに評価）",
           "- **Suppressed By**：上位ルールが発生している場合は表示しない（例：Strike 3% 以内が出ていれば 5% 以内は出さない）",
           "- 説明文の `<v>`＝値、`<t>`＝閾値、`<x>`＝詳細（銘柄名など）に置換", "",
           "## ルール一覧", "",
           "| ID | Alert Type | Scope | Strategy | 指標 | 条件 | 閾値（初期値） | 重要度 | 抑制 | Recommended Sales Action | Data Source |",
           "|---|---|---|---|---|---|---|---|---|---|---|"]
    for r in B.RULES:
        rid, name, scope, strat, metric, op, thr, unit, sev, supp, extra, expl, act, src = r[:14]
        tv = pv.get(thr, thr) if isinstance(thr, str) else thr
        tl = f"{thr} = {tv}" if isinstance(thr, str) else str(thr)
        out.append(f"| {rid} | {esc(name)} | {scope} | {strat} | `{metric}` | {esc(op)} | {esc(tl)} | {sev} | {supp or '—'} | {esc(act)} | {esc(src)} |")
    out += ["", "## 閾値・スコアのパラメータ（19_Settings A 表）", "", "| Key | 初期値 | 単位 | 説明 |", "|---|---|---|---|"]
    for k, v, u, dsc in B.PARAMS:
        out.append(f"| `S_{k}` | {esc(v)} | {esc(u)} | {esc(dsc)} |")
    out += ["", "## 依頼仕様の初期閾値との対応", "",
            "| 依頼仕様 | 実装 |", "|---|---|",
            "| Weekly NAV Decline -2% | R01（S_WeeklyNavDecline） |",
            "| Monthly NAV Decline -5% | R02（S_MonthlyNavDecline） |",
            "| Drawdown -10% | R03（TR 指数の高値からの現在ドローダウン） |",
            "| Option Strike Distance 5% 以内 / 3% 以内 | R11（Medium）/ R10（High） |",
            "| Maturity 30 日以内 / 3 営業日以内 | R14（Low）/ R13（Medium） |",
            "| Exposure Change 前週比 20% 以上 | R20（縮小 High）/ R21（拡大 Low） |",
            "| NAV Delay 公表予定日超過 | R07（評価基準日＋公表ラグ営業日 < 作成日 かつ 基準日 NAV 未取得） |",
            "| Position Data Delay 更新予定日超過 | R19（ポジション基準日が評価基準日より 3 営業日超古い） |",
            "", "## 営業優先度スコア（14_Sales_Priority）", "",
            "`Score = min(High件数,S_SevCap)×S_SevScoreHigh + min(Medium件数,S_SevCap)×S_SevScoreMed + min(Low件数,S_SevCap)×S_SevScoreLow`",
            "`+ 1（残高 ≥ S_LargeBalance） + 2（償還・満期まで ≤ S_RedemptionDays 日） + 1（前回接触から > S_ContactOverdueDays 日、または記録なし） + 1（次回フォロー予定が作成日から S_FollowupWindowDays 日以内）`", "",
            "Score ≥ S_PriorityHigh → High、≥ S_PriorityMed → Medium、それ以外 Low。Override 列に入力すると手動値が優先され、Reason に「（手動変更）」と表示される。", "",
            "## データ品質チェック（18_Data_Quality）", "",
            "Critical が 1 件でも OPEN の場合、`DQ_Overall`＝要確認 となり、メール本文冒頭に【要確認】、Outlook 下書きの件名にも【要確認】が付く。確認のうえ問題なしと判断した場合は Ack 欄に Y と確認者・確認内容を記入する。", "",
            "| ID | チェック | 区分 | 重大度 |", "|---|---|---|---|"]
    for a, b, c, d, _, _ in B.DQ_CHECKS:
        out.append(f"| {a} | {esc(b)} | {c} | {d} |")
    return "\n".join(out) + "\n"


def gaps_md():
    out = ["# 実装ステータス・未実装機能・本番導入前の確認事項", "",
           "実装できていない機能を実装済みと表現しないため、区分を明示する。ブック内では `21_Impl_Status` に同じ内容を収録。", "",
           "## 実装ステータス", "", "| 区分 | 手段 | 内容 |", "|---|---|---|"]
    for a, b, c in B.IMPL:
        out.append(f"| **{a}** | {b} | {esc(c)} |")
    out += ["", "## 未実装機能・制約（Future Enhancement を含む）", "",
            "1. **Power Query はブックに埋め込まれていない。** openpyxl では Power Query（DataMashup）を生成できないため、M コードを `powerquery/` に同梱。ローカル Excel でクエリを作成する必要がある。",
            "2. **VBA はブックに埋め込まれていない。** `vba/modQIS.bas` をインポートして .xlsm で保存する。",
            "3. **Bloomberg 連携は未検証。** 開発環境に Bloomberg 端末がないため、BDH/BDP 数式は `NO_CONN` を返す状態でのみ確認。フィールド名・Ticker・単位・Fill オプションは端末での確認が必要。",
            "4. **Excel 365 での再計算は未検証。** 検証は LibreOffice で実施。Excel 固有の差異（動的配列化による暗黙の交差 `@` の付与など）がないか、初回オープン時に `22_Test_Results` の値と照合する。配列が必要な数式はすべて従来型の配列数式（CSE）で保存している。",
            "5. **実際のポジションファイルは未受領。** 構造分析・テストは追加仕様 A の 17 列定義に基づく架空データで実施（docs/01）。",
            "6. **Greeks の合計は参考値。** 単位・符号の確認前のため、Σ(Greek×Notional) を『参考値（単位未確認）』として表示し、正式なポートフォリオ感応度としては扱わない。",
            "7. **寄与度分析は運用会社データ依存。** 正式データがない場合は Observed（観測）表示に留め、要因を断定しない。Estimated は Vol Target の原資産寄与（前週 Exposure×原指数騰落率）と為替ヘッジなしファンドの為替寄与のみ。",
            "8. **ロール判定は推定。** 同一指数・同一原資産・同一 Put/Call の解消（または満期）と、より遠い満期の新規が同時に起きた場合に Possible Roll とする。取引識別子が提供されれば確定判定に置き換え可能。",
            "9. **Leverage Change／Exposure Change（オプション）**：オプション指数については Gross Notional の前回比として 10 に表示するが、専用アラートは未設定（必要に応じて 19 にルール追加）。",
            "10. **アセットクラス別の損益寄与変化アラート**（Trend/Carry）：寄与データの入力欄はあるが、アラートは未実装（Future Enhancement）。",
            "11. **高度な自然言語生成**：Excel 関数による定型コメントのみ。社内承認済み AI 環境が整備された後に検討（顧客情報を未承認サービスへ送らない）。",
            "12. **容量**：ファンド 30／指数 30／Mapping 30／保有 100／ポジション履歴 1,500 行／NAV 2,000 行／アラート一覧 200 件／メール 1 通あたり 10 保有先。超える場合は build_workbook.py の定数を変更して再生成。",
            "", "## 本番導入前の確認事項", "", "| # | 区分 | 確認事項 |", "|---|---|---|"]
    for i, (c, t) in enumerate(PREPROD, 1):
        out.append(f"| {i} | {c} | {esc(t)} |")
    return "\n".join(out) + "\n"


if __name__ == "__main__":
    DOCS.mkdir(exist_ok=True)
    (DOCS / "05_alert_settings.md").write_text(alerts_md(), encoding="utf-8")
    (DOCS / "08_gaps_and_preprod.md").write_text(gaps_md(), encoding="utf-8")
    print("wrote docs/05, docs/08")
