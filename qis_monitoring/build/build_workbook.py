#!/usr/bin/env python3
"""QIS Fund Monitoring & Weekly Sales Intelligence Workbook を生成する。

  python3 build_workbook.py [--results test_results.json]

計算はすべて Excel 数式で書く（値のハードコードなし）。サンプルデータは sample_data.py。
Excel 2016 以降の関数（MAXIFS / MINIFS / TEXTJOIN）は使うが、XLOOKUP / FILTER / LET /
動的配列は使わない（LibreOffice での検証と、古い Excel 365 チャネルでの互換性のため）。
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.table import Table as XLTable, TableStyleInfo

import sample_data as SD
from xlkit import (BORDER, CENTER, F_BIG, F_BODY, F_H2, F_HEAD, F_INPUT, F_SMALL, F_SUB, FILL_GREY,
                   FILL_HEAD, FILL_INPUT, FILL_RED, FMT_4, FMT_6, FMT_DATE, FMT_INT, FMT_NUM, FMT_PCT,
                   FMT_PCT1, WRAP, Table, add_name, dv_list, section, set_formula, severity_cf, style_value,
                   text_cf, title)

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "dist" / "QIS_Fund_Monitoring_Workbook.xlsx"

CAP_FUND, CAP_IDX, CAP_MAP = 30, 30, 30
CAP_CLIENT, CAP_REP, CAP_HOLD = 50, 20, 100
CAP_POSRAW, CAP_PH, CAP_NAV, CAP_DIST = 500, 1500, 2000, 100
CAP_MD, CAP_SER, CAP_VT, CAP_TC = 80, 2000, 500, 1000
CAP_CH, CAP_CUR, CAP_RULE, CAP_ALERT = 300, 300, 40, 200
CAP_TCX, CAP_SUM, CAP_ATTR, CAP_EMAIL_ITEMS = 100, 20, 300, 10
CAP_GRID = CAP_MAP * CAP_RULE

SHEETS = [
    ("00_Control", "更新操作・設定・更新ステータス"),
    ("01_Dashboard", "ファンド全体の週次サマリー"),
    ("02_Fund_Master", "ファンド属性・指数マスタ・Fund-Index Mapping"),
    ("03_Client_Master", "地銀・社内営業担当者"),
    ("04_Client_Holdings", "地銀別保有ファンド"),
    ("05_Position_Raw", "最新ポジション（Power Query 出力先）"),
    ("06_Position_History", "ポジション履歴（スナップショット）＋標準化"),
    ("07_Bloomberg_Data", "Bloomberg 市場データ・フィールドマッピング"),
    ("08_NAV_History", "NAV 履歴・分配金・NAV 照合"),
    ("09_Performance", "パフォーマンス・寄与度"),
    ("10_Option_Analytics", "オプション系分析（指数単位・ポジション比較）"),
    ("11_VolTarget_Analytics", "Vol Target 分析"),
    ("12_TrendCarry_Analytics", "Trend / Carry 分析"),
    ("13_Risk_Alerts", "リスクアラート"),
    ("14_Sales_Priority", "営業フォローアップ優先順位"),
    ("15_Weekly_Commentary", "週次分析コメント"),
    ("16_Email_Generator", "営業担当者別メール生成"),
    ("17_Email_History", "メール生成履歴"),
    ("18_Data_Quality", "データ品質・整合性チェック"),
    ("19_Settings", "閾値・ルール・マッピング設定"),
    ("20_Instructions", "操作マニュアル"),
    ("21_Impl_Status", "実装ステータス・未実装機能・本番前確認事項"),
    ("22_Test_Results", "テスト結果"),
]

T: dict[str, Table] = {}


def C(key, head, kind="calc", w=11, fmt=None, f=None, arr=False, wrap=False):
    return dict(key=key, head=head, kind=kind, w=w, fmt=fmt, f=f, arr=arr, wrap=wrap)


def G(ex):
    """空行ガード付き数式"""
    return ex


# =====================================================================================
# 19_Settings
# =====================================================================================
PARAMS = [
    ("DataMode", "MANUAL", "MANUAL/BLOOMBERG", "市場データの採用元。BLOOMBERG＝BDH 取得値のみ採用／MANUAL＝手入力・CSV 値を採用（サンプル配布時は MANUAL）"),
    ("NavMatchTol", 0.0005, "比率", "NAV 照合：差が この値以内なら MATCH"),
    ("NavWarnTol", 0.003, "比率", "NAV 照合：差が この値以内なら WARNING、超えると MISMATCH"),
    ("NavDateTol", 0, "日", "評価基準日と NAV 日付の許容差（0＝同一日のみ）"),
    ("AtmBand", 0.005, "比率", "ATM 判定の許容幅（|Strike Distance| ≤ この値なら ATM）"),
    ("NearStrike1", 0.05, "比率", "Strike 接近（Medium）"),
    ("NearStrike2", 0.03, "比率", "Strike 接近（High）"),
    ("MatBizDays", 3, "営業日", "満期接近（営業日）"),
    ("MatDays", 30, "日", "満期接近（暦日）"),
    ("IvJump", 0.03, "Vol（小数）", "加重平均 IV の前回比上昇幅（0.03＝3pt）"),
    ("GammaChg", 0.5, "比率", "Gamma×Notional 合計の前回比変化率（絶対値）"),
    ("ConcUp", 0.10, "比率", "最大 Weight 構成比の上昇幅（0.10＝10pt）"),
    ("NotionalTol", 0.001, "比率", "Notional 変化がこの比率以内なら Unchanged"),
    ("WeeklyNavDecline", -0.02, "比率", "週次 NAV 下落"),
    ("MonthlyNavDecline", -0.05, "比率", "月次 NAV 下落"),
    ("Drawdown", -0.10, "比率", "現在ドローダウン"),
    ("ExposureDrop", -0.20, "比率", "Vol Target Exposure 急縮小（前週比）"),
    ("ExposureChg", 0.20, "比率", "Vol Target Exposure 急拡大（前週比）"),
    ("LevCapRatio", 0.95, "比率", "Exposure ÷ 上限 がこの値以上で上限接近"),
    ("RvJump", 0.03, "Vol（小数）", "実現 Vol 前週比上昇幅"),
    ("TrackDiff", 0.02, "比率", "Strategy と原指数の週次騰落率差（絶対値）"),
    ("CarryDrop", -0.01, "比率", "加重平均 Carry の前週比低下幅（-0.01＝-1pt）"),
    ("PosConc", 0.40, "比率", "Trend/Carry 単一銘柄の |Weight| 構成比"),
    ("PosDelayBizDays", 3, "営業日", "ポジション基準日が評価基準日よりこの営業日数を超えて古い場合に遅延"),
    ("MktStaleDays", 0, "日", "市場データ As-of が評価基準日よりこの日数を超えて古い場合 STALE"),
    ("AbnReturn", 0.10, "比率", "週次リターンの絶対値がこの値以上なら異常値として確認"),
    ("VolWeeks", 26, "週", "実現ボラティリティの計測期間"),
    ("IdxLvlTol", 0.001, "比率", "ポジションファイルの Index Level と Bloomberg 指数値の許容差"),
    ("SevScoreHigh", 3, "点", "営業優先度：High アラート 1 件あたり"),
    ("SevScoreMed", 2, "点", "営業優先度：Medium アラート 1 件あたり"),
    ("SevScoreLow", 1, "点", "営業優先度：Low アラート 1 件あたり"),
    ("LargeBalance", 1_000_000_000, "円", "残高大（+1 点）"),
    ("RedemptionDays", 90, "日", "償還・満期予定がこの日数以内（+2 点）"),
    ("ContactOverdueDays", 30, "日", "前回コンタクトからの経過日数超過（+1 点）"),
    ("FollowupWindowDays", 7, "日", "次回フォロー予定がこの日数以内（+1 点）"),
    ("SevCap", 2, "件", "重要度ごとに加点するアラート件数の上限（同一事象の指数別重複による過大評価を防ぐ）"),
    ("PriorityHigh", 8, "点", "合計点がこの値以上で High"),
    ("PriorityMed", 4, "点", "合計点がこの値以上で Medium"),
    ("ReportPrefix", "【QIS Weekly Update】担当先ファンド状況｜", "文字列", "メール件名の接頭辞"),
    ("WorkbookVersion", "v1.0.0", "文字列", "ブックのバージョン（メール履歴に記録）"),
]

# id, name, scope, strategy, metric, op, thr(param or number), unit, sev, supp, extra, expl, action, source, tp, q, a
RULES = [
    ("R01", "週次NAV下落", "FUND", "ALL", "WEEKLY_RET", "<=", "WeeklyNavDecline", "PCT", "High", "", "",
     "公式NAVの前週比が<v>（閾値<t>）。", "パフォーマンス説明（下落要因は確認済み情報と推測を区別して説明）", "08_NAV_History（運用会社公式NAV）",
     "下落幅と参照指数・原資産の動きを並べて示し、確認できた要因と未確認事項を分けて説明する。",
     "今週NAVが下落した理由は？", "参照指数・原資産の下落と同方向の動きが観測されています。正式な要因分解は運用会社の寄与度データで確認のうえ回答します。"),
    ("R02", "月次NAV下落", "FUND", "ALL", "MONTHLY_RET", "<=", "MonthlyNavDecline", "PCT", "High", "", "",
     "前月末比が<v>（閾値<t>）。", "月次での下落局面説明と今後の見通し・リスク要因の共有", "08_NAV_History",
     "月次での下落が閾値を超過。保有継続判断に関わる質問が想定されるため、商品性（下落時の挙動）を再確認して説明する。",
     "下落が続いているが、商品性どおりの動きか？", "参照指数の動きと概ね整合しているか確認中です。商品設計上の想定範囲かは運用会社資料に基づき回答します。"),
    ("R03", "ドローダウン拡大", "FUND", "ALL", "CUR_DD", "<=", "Drawdown", "PCT", "Medium", "", "",
     "TRベースの高値からの下落率が<v>（閾値<t>）。", "ドローダウン状況と回復条件の説明", "08_NAV_History（TR指数化）",
     "高値からの下落率が閾値を超過。戦略特性上の想定ドローダウンと比較して説明する。",
     "過去最大の下落局面と比べてどうか？", "現時点の下落率と過去の最大ドローダウンを09_Performanceで比較できます。将来の回復は保証されない点も併せて説明します。"),
    ("R04", "NAV不一致", "FUND", "ALL", "NAV_MISMATCH", ">=", 1, "FLAG", "High", "", "NAV_DESC",
     "公式NAVとBloomberg NAVが許容差を超えて乖離。<x>", "社内確認（顧客連絡前に運用会社・データ担当へ確認）", "08_NAV_History（公式 vs Bloomberg）",
     "NAVの不一致を解消するまで、当該ファンドの数値を顧客向けに使用しない。",
     "（社内）どちらのNAVが正しいか？", "運用会社公式NAVを優先します。Bloomberg側の更新遅延・定義差の可能性を確認中です。"),
    ("R05", "NAV照合 警告", "FUND", "ALL", "NAV_WARNING", ">=", 1, "FLAG", "Low", "", "NAV_DESC",
     "公式NAVとBloomberg NAVに軽微な差。<x>", "データ担当による確認", "08_NAV_History",
     "軽微な差。丸め・更新タイミングの差か確認する。", "（社内）差の原因は？", "丸め・更新タイミングの差の可能性。運用会社公式NAVを優先します。"),
    ("R06", "NAV照合不能", "FUND", "ALL", "NAV_UNCOMP", ">=", 1, "FLAG", "Medium", "", "NAV_DESC",
     "同一基準日・同一定義での照合ができない。<x>", "データ取得状況の確認", "08_NAV_History",
     "照合できないため、差額によるアラートは出していない。データの揃い次第再確認する。", "（社内）照合できない理由は？", "日付・通貨・NAV定義のいずれかが一致していないため。比較可能なデータの取得後に再照合します。"),
    ("R07", "NAV公表遅延", "FUND", "ALL", "NAV_DELAY", ">=", 1, "FLAG", "Medium", "", "",
     "公表予定日を過ぎても評価基準日の公式NAVが未取得。", "運用会社へNAV公表状況を確認", "08_NAV_History／02_Fund_Master（公表ラグ）",
     "NAV未取得のため、週次リターンは算出していない（前回値の代用はしない）。", "今週のNAVはいつ分かるか？", "運用会社に公表予定を確認中です。判明次第共有します。"),
    ("R08", "異常リターン（要確認）", "FUND", "ALL", "ABN_RET", "ABS>=", "AbnReturn", "PCT", "Medium", "", "",
     "週次リターン<v>が異常値閾値（<t>）を超過。データ誤りの可能性も含め確認。", "データ確認後に顧客対応を判断", "09_Performance",
     "異常値の可能性があるため、確認が取れるまで顧客説明に使用しない。", "（社内）データは正しいか？", "運用会社の公式NAVと突合中です。"),
    ("R10", "Strike接近（3%以内）", "INDEX", "OPTION", "STRIKE_DIST", "<=", "NearStrike2", "PCTABS", "High", "", "NEAREST_DESC",
     "最もStrikeに近いポジションの距離が<v>（閾値<t>）：<x>。", "リスク状況確認・原資産下落時の挙動説明", "06_Position_History（運用会社ポジションファイル）",
     "Strike近傍のポジションがあり、原資産が更に下落した場合の損益感応度が高まり得る（ポジションからの推定。NAV影響は未確定）。",
     "原資産が更に下落した場合の影響は？", "【推定】Strike近傍のポジションがあり、下落時は損失が拡大し得る構造です。正式な感応度は運用会社のリスク数値で確認のうえ回答します。"),
    ("R11", "Strike接近（5%以内）", "INDEX", "OPTION", "STRIKE_DIST", "<=", "NearStrike1", "PCTABS", "Medium", "R10", "NEAREST_DESC",
     "最もStrikeに近いポジションの距離が<v>（閾値<t>）：<x>。", "リスク状況確認", "06_Position_History",
     "Strikeまでの距離が縮小。原資産の動向次第で感応度が高まる点を共有（推定）。",
     "Strikeまでの余裕はどの程度か？", "【事実】最接近ポジションのStrike距離は<v>です（ポジション基準日時点）。"),
    ("R12", "ITMへの変化", "INDEX", "OPTION", "ITM_TRANS", ">=", 1, "COUNT", "High", "", "ITM_DESC",
     "前回OTM/ATMだったポジションがITMに変化：<x>。", "リスク状況確認・損益影響の確認（運用会社照会）", "06_Position_History",
     "ITM化したポジションがある。満期時の損益への影響は運用会社の評価で確認が必要（推定と事実を区別）。",
     "ITMになったオプションはどうなるのか？", "【推定】満期までに原資産が回復しなければ、当該ポジションは損失要因となり得ます。運用会社のロール・ヘッジ方針を確認のうえ回答します。"),
    ("R13", "満期接近（営業日）", "INDEX", "OPTION", "MAT_BDAYS", "<=", "MatBizDays", "BDAYS", "Medium", "", "MAT_DESC",
     "最短満期まで<v>（閾値<t>）：<x>。", "満期・ロール予定の確認", "06_Position_History",
     "近く満期を迎えるポジションがあり、ロールに伴うポジション変化が想定される。",
     "満期を迎えるポジションはどうなるか？", "運用ルールに基づき新しい満期へロールされる見込みです（運用会社に確認）。"),
    ("R14", "満期接近（30日以内）", "INDEX", "OPTION", "MAT_DAYS", "<=", "MatDays", "DAYS", "Low", "R13", "MAT_DESC",
     "最短満期まで<v>（閾値<t>）：<x>。", "満期スケジュールの把握", "06_Position_History",
     "30日以内に満期を迎えるポジションあり。", "次のロールはいつか？", "最短満期は<x>です。"),
    ("R15", "IV急上昇", "INDEX", "OPTION", "IV_CHG", ">=", "IvJump", "PCTPT", "Medium", "", "",
     "加重平均IVが前回比<v>（閾値<t>）。", "ボラティリティ環境の説明", "06_Position_History（ファイル上のImplied Vol）",
     "IVが上昇。オプション売り戦略では評価損要因となる一方、新規ポジションのプレミアム水準は上昇し得る（一般論）。",
     "ボラティリティ上昇の影響は？", "【推定】売りポジションには評価上の逆風となり得ますが、ロール時に受け取るプレミアムは増える傾向にあります。実際の影響は運用会社の寄与度で確認します。"),
    ("R16", "Gamma急変", "INDEX", "OPTION", "GAMMA_CHG", "ABS>=", "GammaChg", "PCT", "Low", "", "",
     "Gamma×Notional 合計（参考値・単位未確認）が前回比<v>（閾値<t>）。", "感応度変化の把握（社内）", "06_Position_History",
     "感応度が変化（参考値。Greeksの単位確認前のため正式数値ではない）。", "（社内）感応度は上がっているか？", "参考値では上昇。正式値は運用会社に確認します。"),
    ("R17", "Weight集中度上昇", "INDEX", "OPTION", "CONC_CHG", ">=", "ConcUp", "PCTPT", "Low", "", "",
     "単一ポジションの|Weight|構成比が前回比<v>（閾値<t>）。", "ポジション構成の把握", "06_Position_History",
     "ポジション構成の集中度が上昇。", "（社内）集中度の変化は？", "最大構成比が上昇しています（Weight 合計は100%を前提としない）。"),
    ("R18", "オプションロール（推定）", "INDEX", "OPTION", "ROLL_CNT", ">=", 1, "COUNT", "Low", "", "ROLL_DESC",
     "解消と新規の組合せからロールの可能性を検出（推定）：<x>。", "ポジション変化の説明（ロールは推定である旨を明記）", "06_Position_History（前回比較）",
     "ロールと推定されるポジション入替えあり（確定ではない）。", "ポジションが入れ替わったのはなぜか？", "【推定】満期接近・解消に伴うロールと見られます。確定情報は運用会社に確認します。"),
    ("R19", "ポジションデータ遅延", "INDEX", "OPTION", "POS_STALE", ">", "PosDelayBizDays", "BDAYS", "Medium", "", "",
     "ポジション基準日が評価基準日より<v>古い（閾値<t>）。", "最新ポジションファイルの受領確認", "05_Position_Raw／06_Position_History",
     "ポジション情報が古いため、ポジション分析は参考扱い。", "（社内）最新ポジションは？", "運用会社に最新ファイルを依頼中です。"),
    ("R20", "Exposure急縮小", "INDEX", "VOLTARGET", "EXP_CHG", "<=", "ExposureDrop", "PCT", "High", "", "VT_DESC",
     "Exposureが前週比<v>（閾値<t>）。<x>", "戦略動向説明（Vol Targetの仕組み）", "11_VolTarget_Analytics（運用会社週次データ）",
     "ボラティリティ上昇局面でExposureが縮小。Vol Targetの仕組みと整合的な動きとして説明できる（観測）。",
     "なぜ株価下落時に組入比率が下がったのか？", "【観測】実現ボラティリティの上昇に応じて、目標ボラティリティを維持するため組入比率を引き下げる仕組みです。反発局面では追随が遅れる可能性もあります。"),
    ("R21", "Exposure急拡大", "INDEX", "VOLTARGET", "EXP_CHG", ">=", "ExposureChg", "PCT", "Low", "", "VT_DESC",
     "Exposureが前週比<v>（閾値<t>）。<x>", "戦略動向説明", "11_VolTarget_Analytics",
     "ボラティリティ低下に伴いExposureが拡大。", "組入比率が上がった理由は？", "【観測】実現ボラティリティの低下に伴い組入比率が引き上げられています。"),
    ("R22", "レバレッジ上限接近", "INDEX", "VOLTARGET", "EXP_RATIO", ">=", "LevCapRatio", "RATIO", "Medium", "", "",
     "Exposure÷上限 が<v>（閾値<t>）。", "上限到達時の挙動説明", "11_VolTarget_Analytics",
     "Exposureが上限付近。ボラティリティ低下局面でもこれ以上は増えない。", "上限に達するとどうなるか？", "上限に達した場合、目標ボラティリティを下回ってもExposureは上限で頭打ちになります。"),
    ("R23", "実現Vol急上昇", "INDEX", "VOLTARGET", "RV_CHG", ">=", "RvJump", "PCTPT", "Medium", "", "VT_DESC",
     "実現Volが前週比<v>（閾値<t>）。<x>", "市場環境説明", "11_VolTarget_Analytics",
     "実現Volが上昇し、Exposure調整が想定される。", "今後さらに組入比率は下がるか？", "【推定】実現ボラティリティが高止まりする場合、追加の縮小が想定されます。"),
    ("R24", "リバランス発生", "INDEX", "VOLTARGET", "REBAL", ">=", 1, "COUNT", "Low", "", "",
     "評価期間内にリバランスが<v>発生。", "戦略動向の把握", "11_VolTarget_Analytics",
     "リバランス実施。", "（社内）リバランス頻度は？", "運用会社データで確認できます。"),
    ("R25", "原指数との乖離", "INDEX", "VOLTARGET", "TRACK", "ABS>=", "TrackDiff", "PCT", "Low", "", "",
     "Strategy Indexと原指数の週次騰落率差が<v>（閾値<t>）。", "乖離要因の説明", "11_VolTarget_Analytics",
     "原指数との乖離が拡大（Exposure調整による一般的な差）。", "株価指数と動きが違うのはなぜか？", "【観測】組入比率の調整により、原指数と異なる騰落となります。"),
    ("R26", "Trendシグナル反転", "INDEX", "TRENDCARRY", "SIG_REV", ">=", 1, "COUNT", "Medium", "", "SIGREV_DESC",
     "シグナルの符号が反転した銘柄：<x>。", "戦略動向説明（ポジション転換）", "12_TrendCarry_Analytics（運用会社週次ポジション）",
     "トレンドシグナルが反転した資産あり。反転局面ではトレンド戦略の損益が振れやすい（一般論）。",
     "ポジションが反転したのはなぜか？", "【事実】シグナルの符号が反転しました。【推定】価格トレンドの変化を受けたルールに基づく転換です。"),
    ("R27", "Long/Short反転", "INDEX", "TRENDCARRY", "LS_REV", ">=", 1, "COUNT", "Medium", "", "LSREV_DESC",
     "Long/Shortが反転した銘柄：<x>。", "戦略動向説明", "12_TrendCarry_Analytics",
     "保有方向が反転した資産あり。", "どの資産で売り買いが入れ替わったか？", "【事実】<x>で保有方向が反転しています。"),
    ("R28", "Carry低下", "INDEX", "TRENDCARRY", "CARRY_CHG", "<=", "CarryDrop", "PCTPT", "Low", "", "",
     "加重平均Carry推定値が前週比<v>（閾値<t>）。", "収益源の変化説明", "12_TrendCarry_Analytics",
     "Carry推定値が低下（推定値）。", "インカム的な収益は減っているか？", "【推定】Carry推定値は低下しています。実績寄与は運用会社データで確認します。"),
    ("R29", "ポジション集中", "INDEX", "TRENDCARRY", "TC_CONC", ">=", "PosConc", "RATIO", "Low", "", "",
     "単一銘柄の|Weight|構成比が<v>（閾値<t>）。", "集中リスクの説明", "12_TrendCarry_Analytics",
     "特定資産への集中度が上昇。", "特定資産に偏っていないか？", "最大構成比は<v>です。"),
    ("R30", "市場データ欠損", "INDEX", "ALL", "MKT_MISSING", ">=", 1, "COUNT", "Medium", "", "MKT_DESC",
     "関連する市場データが欠損または古い：<x>。", "データ取得の確認（前回値で代用しない）", "07_Bloomberg_Data",
     "市場データ欠損のため、関連する市場コメントは未算出。", "（社内）データは揃っているか？", "欠損データの取得後に更新します。"),
]

STRATS = [
    ("OPTION", "Option-Based Strategies",
     "原資産の水準とStrikeとの距離・満期構成を確認し、オプション戦略の損益の出方（推定）を説明する。",
     "オプション戦略はどのような局面で収益が出るか？",
     "一般にオプション売り戦略は相場が安定・緩やかな上昇局面でプレミアム収入を得やすく、急落局面では損失が拡大し得る構造です。"),
    ("VOLTARGET", "Volatility Target",
     "実現ボラティリティとExposureの関係を示し、Vol Targetの仕組みを説明する。",
     "Vol Target戦略はどのようにリスクを調整しているか？",
     "実現ボラティリティが上昇すると組入比率を下げ、低下すると引き上げることで、目標ボラティリティ近辺を維持する仕組みです。"),
    ("TRENDCARRY", "Trend Following / Carry",
     "資産クラス別のシグナル・保有方向とCarry推定値の変化を説明する。",
     "トレンド戦略はどのような局面に弱いか？",
     "トレンドが反転を繰り返すレンジ相場ではシグナル転換が多くなり、収益が出にくい傾向があります。"),
]

COLMAP = [  # SourceColumn（受領ファイルの列名）→ StandardColumn（05 の列名）。列名変更時は SourceColumn のみ修正
    (c, c, t, req) for c, t, req in [
        ("Position Date", "date", "Y"), ("Index Name", "text", "Y"), ("Name", "text", "N"), ("BloombergTicker", "text", "Y"),
        ("Option Mid Price", "number", "N"), ("Strike", "number", "Y"), ("Maturity Date", "date", "Y"),
        ("Notional(USD)", "number", "Y"), ("Delta", "number", "N"), ("Gamma", "number", "N"), ("Vega", "number", "N"),
        ("Theta", "number", "N"), ("Implied Vol", "number", "N"), ("NDX Spot", "number", "Y"), ("Index Level", "number", "N"),
        ("Weight", "number", "N"), ("open_date", "date", "N")]
]


def build_settings(wb, ws):
    title(ws, "19_Settings　閾値・ルール・マッピング設定", "閾値・アラートルール・リスト・Power Query 列マッピング・休日・イベントを一元管理。数式はここを参照する。", 18)
    section(ws, 4, "A. パラメータ（名前定義 S_<Key> で参照）")
    t = Table(wb, ws, "SP0", 5, [C("Key", "Key", "aux", 22), C("Value", "Value", "in", 16), C("Unit", "単位", "aux", 12),
                                 C("Desc", "説明", "aux", 70, wrap=False)], len(PARAMS), name_cols=False)
    t.write([dict(Key=k, Value=v, Unit=u, Desc=d) for k, v, u, d in PARAMS])
    for i, (k, v, u, d) in enumerate(PARAMS):
        r = t.first + i
        add_name(wb, f"S_{k}", ws.title, f"$B${r}")
        if isinstance(v, float) and abs(v) < 1:
            ws[f"B{r}"].number_format = "0.00%" if k not in ("IvJump", "RvJump") else "0.000"
    dv_list(ws, f"B{t.first}", '"MANUAL,BLOOMBERG"')

    r0 = 60
    section(ws, r0 - 1, "B. アラートルール（有効化・閾値・重要度・説明文・営業アクションを変更可。<v>=値、<t>=閾値、<x>=詳細 に置換）")
    rl = Table(wb, ws, "RL", r0, [
        C("ID", "Rule ID", "in", 8), C("Name", "Alert Type", "in", 20), C("Scope", "Scope", "in", 8),
        C("Strategy", "Strategy", "in", 12), C("Metric", "Metric Key", "in", 14), C("Op", "Op", "in", 6),
        C("Thr", "Threshold", "in", 10), C("Unit", "Unit", "in", 8), C("Sev", "Severity", "in", 9),
        C("Supp", "Suppressed By", "in", 9), C("Extra", "Extra Key", "in", 13), C("Expl", "Explanation", "in", 40),
        C("Action", "Recommended Sales Action", "in", 32), C("Source", "Data Source", "in", 24),
        C("TP", "Talking Point", "in", 40), C("Q", "想定質問", "in", 28), C("A", "回答案", "in", 40),
        C("Enabled", "Enabled", "in", 8)], CAP_RULE)
    rows = []
    for (rid, name, scope, strat, metric, op, thr, unit, sev, supp, extra, expl, act, src, tp, qq, aa) in RULES:
        rows.append(dict(ID=rid, Name=name, Scope=scope, Strategy=strat, Metric=metric, Op=op,
                         Thr=(f"=S_{thr}" if isinstance(thr, str) else thr), Unit=unit, Sev=sev, Supp=supp,
                         Extra=extra, Expl=expl, Action=act, Source=src, TP=tp, Q=qq, A=aa, Enabled="Y"))
    rl.write(rows)
    for i, row in enumerate(rows):
        if isinstance(row["Thr"], str):
            ws[f"{rl.col('Thr')}{rl.first + i}"] = row["Thr"]
    dv_list(ws, f"{rl.col('Sev')}{rl.first}:{rl.col('Sev')}{rl.last}", '"High,Medium,Low"')
    dv_list(ws, f"{rl.col('Enabled')}{rl.first}:{rl.col('Enabled')}{rl.last}", '"Y,N"')
    dv_list(ws, f"{rl.col('Op')}{rl.first}:{rl.col('Op')}{rl.last}", '"<=,>=,<,>,ABS>="')
    severity_cf(ws, f"{rl.col('Sev')}{rl.first}:{rl.col('Sev')}{rl.last}")
    T["RL"] = rl

    r1 = 105
    section(ws, r1 - 1, "C. 戦略別の既定コメント（アラートがない場合に使用）")
    st = Table(wb, ws, "ST", r1, [C("Code", "Strategy", "in", 12), C("Name", "名称", "in", 22), C("TP", "Talking Point", "in", 40),
                                   C("Q", "想定質問", "in", 30), C("A", "回答案", "in", 40)], 8)
    st.write([dict(Code=a, Name=b, TP=c, Q=d, A=e) for a, b, c, d, e in STRATS])
    T["ST"] = st

    r2 = 118
    section(ws, r2 - 1, "D. 入力リスト・ダッシュボード絞込みリスト")
    lists = {
        "L_Sev": ["High", "Medium", "Low"], "L_YN": ["Y", "N"], "L_Strategy": ["OPTION", "VOLTARGET", "TRENDCARRY"],
        "L_Source": ["OFFICIAL", "BLOOMBERG"], "L_Basis": ["TR", "PRICE"], "L_Status": ["ACTIVE", "PIPELINE", "CLOSED"],
        "L_Prio": ["High", "Medium", "Low"], "L_Review": ["未確認", "対応中", "対応済", "対応不要"],
    }
    for j, (nm, vals) in enumerate(lists.items()):
        L = chr(ord("A") + j)
        ws[f"{L}{r2}"] = nm
        ws[f"{L}{r2}"].font = F_HEAD
        ws[f"{L}{r2}"].fill = FILL_HEAD
        for i, v in enumerate(vals):
            ws[f"{L}{r2 + 1 + i}"] = v
            style_value(ws[f"{L}{r2 + 1 + i}"], "in")
        add_name(wb, nm, ws.title, f"${L}${r2 + 1}:${L}${r2 + len(vals)}")
    r3 = 130
    heads = [("L_FBank", "投資家（絞込）", "03_Client_Master", "B", 6, CAP_CLIENT),
             ("L_FRep", "営業担当（絞込）", "03_Client_Master", "B", 61, CAP_REP),
             ("L_FStrategy", "戦略（絞込）", None, None, None, None),
             ("L_FFund", "ファンドID（チャート）", "02_Fund_Master", "A", 6, CAP_FUND),
             ("L_Rep", "営業担当（メール）", "03_Client_Master", "B", 61, CAP_REP)]
    for j, (nm, head, sh, col, start, cap) in enumerate(heads):
        L = chr(ord("A") + j)
        ws[f"{L}{r3}"] = head
        ws[f"{L}{r3}"].font = F_HEAD
        ws[f"{L}{r3}"].fill = FILL_HEAD
        off = 0
        if nm not in ("L_FFund", "L_Rep"):
            ws[f"{L}{r3 + 1}"] = "(全て)"
            off = 1
        if nm == "L_FStrategy":
            for i, v in enumerate(lists["L_Strategy"]):
                ws[f"{L}{r3 + 2 + i}"] = v
            n = 4
        else:
            for i in range(cap):
                set_formula(ws, f"{L}{r3 + 1 + off + i}", f"=IF('{sh}'!{col}{start + i}=\"\",\"\",'{sh}'!{col}{start + i})")
            n = cap + off
        add_name(wb, nm, ws.title, f"${L}${r3 + 1}:${L}${r3 + n}")

    r4 = 190
    section(ws, r4 - 1, "E. Power Query 列マッピング（tblColMap：ソース列名 → 標準列名。運用会社の列名変更時はここを修正）")
    for j, h in enumerate(["SourceColumn", "StandardColumn", "DataType", "Required"]):
        c = ws.cell(r4, 1 + j, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
    for i, row in enumerate(COLMAP):
        for j, v in enumerate(row):
            c = ws.cell(r4 + 1 + i, 1 + j, v)
            style_value(c, "in")
    tab = XLTable(displayName="tblColMap", ref=f"A{r4}:D{r4 + len(COLMAP)}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tab)

    r5 = 215
    section(ws, r5 - 1, "F. 休日（NETWORKDAYS 用。米国市場休場日のサンプル。要更新）")
    ws[f"A{r5}"] = "Holiday"
    ws[f"A{r5}"].font = F_HEAD
    ws[f"A{r5}"].fill = FILL_HEAD
    for i in range(50):
        c = ws[f"A{r5 + 1 + i}"]
        c.value = SD.HOLIDAYS[i] if i < len(SD.HOLIDAYS) else None
        style_value(c, "in", FMT_DATE)
    add_name(wb, "Holidays", ws.title, f"$A${r5 + 1}:$A${r5 + 50}")

    r6 = 270
    section(ws, r6 - 1, "G. 市場イベントカレンダー（翌週の注目事項に使用。サンプル：要更新）")
    ev = Table(wb, ws, "EV", r6, [C("Date", "Date", "in", 12, FMT_DATE), C("Event", "Event", "in", 40),
                                   C("Region", "Region", "in", 8), C("Importance", "Importance", "in", 10)], 50)
    ev.write([dict(Date=a, Event=b, Region=c, Importance=d) for a, b, c, d in SD.EVENTS])
    ws.freeze_panes = "A5"


# =====================================================================================
# 00_Control
# =====================================================================================
def build_control(wb, ws):
    title(ws, "00_Control　更新操作・設定・更新ステータス", "毎週月曜日の更新はこのシートの手順 1→12 の順に実行。入力セル（薄黄色）以外は数式。", 12)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 46
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 52
    for c in "GHI":
        ws.column_dimensions[c].width = 14
    section(ws, 4, "A. 基準日・作成日")
    items = [
        (5, "評価基準日（自動：直近金曜日）", "=TODAY()-MOD(WEEKDAY(TODAY(),2)-5,7)", "calc", FMT_DATE, "月曜実行時は前週金曜日"),
        (6, "評価基準日（手入力で上書き）", SD.EVAL_DATE, "in", FMT_DATE, "空欄なら自動値を使用。祝日等で変更する場合に入力"),
        (7, "評価基準日（使用）", "=IF(ISNUMBER(C6),C6,C5)", "calc", FMT_DATE, "名前 EvalDate"),
        (8, "前週基準日", "=EvalDate-7", "calc", FMT_DATE, "名前 PrevEvalDate"),
        (9, "前月末基準日", "=EOMONTH(EvalDate,-1)", "calc", FMT_DATE, "名前 MonthBaseDate（これ以前の直近 NAV を基準に使用）"),
        (10, "前年末基準日", "=DATE(YEAR(EvalDate)-1,12,31)", "calc", FMT_DATE, "名前 YearBaseDate"),
        (11, "レポート作成日（自動）", "=TODAY()", "calc", FMT_DATE, ""),
        (12, "レポート作成日（手入力で上書き）", SD.RUN_DATE, "in", FMT_DATE, "再現・監査時は当時の作成日を入力"),
        (13, "レポート作成日（使用）", "=IF(ISNUMBER(C12),C12,C11)", "calc", FMT_DATE, "名前 RunDate（件名の日付・NAV 遅延判定に使用）"),
        (14, "市場データ採用モード", "=S_DataMode", "calc", None, "19_Settings で変更（MANUAL / BLOOMBERG）"),
        (15, "Bloomberg 接続テスト（BDP）", '=IFERROR(BDP("NDX Index","PX_LAST"),"NO_CONN")', "calc", FMT_NUM, "NO_CONN＝Add-in 未接続。値が出れば接続 OK（※参考値。レポートには使用しない）"),
        (16, "Bloomberg 接続状態", '=IF(ISNUMBER(C15),"CONNECTED","NOT CONNECTED")', "calc", None, ""),
    ]
    for r, label, val, kind, fmt, note in items:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = F_BODY
        if isinstance(val, str) and val.startswith("="):
            set_formula(ws, f"C{r}", val)
        else:
            ws[f"C{r}"] = val
        style_value(ws[f"C{r}"], kind, fmt)
        ws[f"D{r}"] = note
        ws[f"D{r}"].font = F_SMALL
    add_name(wb, "EvalDate", ws.title, "$C$7")
    add_name(wb, "PrevEvalDate", ws.title, "$C$8")
    add_name(wb, "MonthBaseDate", ws.title, "$C$9")
    add_name(wb, "YearBaseDate", ws.title, "$C$10")
    add_name(wb, "RunDate", ws.title, "$C$13")
    add_name(wb, "BbgConn", ws.title, "$C$16")

    section(ws, 4, "B. Power Query パラメータ（tblPQParams）", col=5)
    pq = [("PositionFolder", r"C:\QIS\Positions\Inbox"), ("PositionFilePattern", "positions_*.csv;*.xlsx"),
          ("PositionSheetName", "Positions"), ("ArchiveFolder", r"C:\QIS\Positions\Archive"),
          ("NavFolder", r"C:\QIS\NAV\Inbox"), ("MarketCsvPath", r"C:\QIS\Market\market_data_manual.csv")]
    ws["E5"], ws["F5"] = "Param", "Value"
    for c in ("E5", "F5"):
        ws[c].font = F_HEAD
        ws[c].fill = FILL_HEAD
    for i, (k, v) in enumerate(pq):
        ws[f"E{6 + i}"] = k
        ws[f"F{6 + i}"] = v
        style_value(ws[f"E{6 + i}"], "aux")
        style_value(ws[f"F{6 + i}"], "in")
    tab = XLTable(displayName="tblPQParams", ref=f"E5:F{5 + len(pq)}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(tab)

    section(ws, 18, "C. 更新ステータス")
    status = [
        (19, "最新ポジション基準日（オプション指数）", "=MAX(OA_CurDate)", FMT_DATE),
        (20, "ポジション基準日 ≠ 評価基準日 の指数数", '=SUMPRODUCT((OA_ID<>"")*(OA_CurDate<>EvalDate))', FMT_INT),
        (21, "公式 NAV 取得済ファンド（評価基準日）", '=COUNTIF(PF_OnEval,"Y")&" / "&COUNTIF(FM_Status,"ACTIVE")', None),
        (22, "NAV 照合 MISMATCH 件数", '=COUNTIF(RC_Status,"MISMATCH")', FMT_INT),
        (23, "市場データ 欠損・STALE 件数", '=COUNTIF(MD_Status,"MISSING")+COUNTIF(MD_Status,"STALE")', FMT_INT),
        (24, "データ品質（Critical 未解決）", "=DQ_Overall", None),
        (25, "アラート件数 High / Medium / Low", '=COUNTIF(AL_Sev,"High")&" / "&COUNTIF(AL_Sev,"Medium")&" / "&COUNTIF(AL_Sev,"Low")', None),
        (26, "営業優先度 High の保有行数", '=COUNTIF(SP_Final,"High")', FMT_INT),
        (27, "メール生成ステータス", "=EM_Status", None),
        (28, "ブックのバージョン", "=S_WorkbookVersion", None),
    ]
    for r, label, f, fmt in status:
        ws[f"B{r}"] = label
        ws[f"B{r}"].font = F_BODY
        set_formula(ws, f"C{r}", f)
        style_value(ws[f"C{r}"], "calc", fmt)
    text_cf(ws, "C24", ["要確認"])

    section(ws, 31, "D. 週次オペレーション（毎週月曜日）")
    heads = ["No", "手順", "操作（ボタン／マクロ／手作業）", "自動ステータス", "完了(Y)", "担当", "備考"]
    for j, h in enumerate(heads):
        c = ws.cell(32, 1 + j, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = CENTER
    steps = [
        ("最新ポジションファイルを指定フォルダに保存", "手作業：tblPQParams の PositionFolder へ保存", '=IF(MAX(OA_CurDate)>=EvalDate,"OK：最新","要確認：基準日 "&TEXT(MAX(OA_CurDate),"yyyy/mm/dd"))'),
        ("Excel を開く（編集を有効化）", "手作業", '="OK"'),
        ("データ更新（Power Query）", "データ＞すべて更新、またはマクロ QIS_RefreshAll", '=IF(COUNTA(PR_PositionDate)>0,"取込済 "&COUNTA(PR_PositionDate)&"行","未取込")'),
        ("ポジションを履歴へ保存", "マクロ QIS_ArchivePositions（同一ファイルの重複保存は自動でスキップ）", '=IF(COUNTIF(PH_Dup,">1")=0,"OK：重複なし","要確認：重複 "&COUNTIF(PH_Dup,">1")&"行")'),
        ("Bloomberg 市場データ更新", "Bloomberg＞Refresh、またはマクロ QIS_RefreshBloomberg", '=IF(S_DataMode="BLOOMBERG",BbgConn,"MANUAL モード（手入力値を使用）")'),
        ("NAV 照合", "08_NAV_History（自動計算）", '="MISMATCH "&COUNTIF(RC_Status,"MISMATCH")&"件／照合不能 "&(COUNTIF(RC_Status,"DATE DIFFERENCE")+COUNTIF(RC_Status,"MISSING DATA")+COUNTIF(RC_Status,"NOT COMPARABLE"))&"件"'),
        ("データ品質チェック", "18_Data_Quality（自動計算）。重大項目は確認のうえ承認欄に記入", "=DQ_Overall"),
        ("Performance / Risk Analytics 更新", "自動計算（09〜13）", '="アラート "&COUNT(GR_RankKey)&"件"'),
        ("営業優先順位の更新", "14_Sales_Priority（手動変更は Override 列）", '="High "&COUNTIF(SP_Final,"High")&"件"'),
        ("Weekly Commentary 生成", "15_Weekly_Commentary（修正は Override 列）", '="OK"'),
        ("営業担当者別メール生成", "16_Email_Generator で担当者を選択", "=EM_Status"),
        ("レビュー → Outlook 下書き作成", "マクロ QIS_CreateOutlookDrafts（自動送信はしない）／QIS_LogEmail", '=IF(COUNTA(EH_LogID)>0,"履歴 "&COUNTA(EH_LogID)&"件","未作成")'),
    ]
    for i, (s, op, f) in enumerate(steps):
        r = 33 + i
        ws[f"A{r}"] = i + 1
        ws[f"B{r}"] = s
        ws[f"C{r}"] = op
        set_formula(ws, f"D{r}", f)
        for col in "ABCD":
            style_value(ws[f"{col}{r}"], "calc")
            ws[f"{col}{r}"].alignment = WRAP
        for col in "EFG":
            style_value(ws[f"{col}{r}"], "in")
        ws.row_dimensions[r].height = 30
    dv_list(ws, "E33:E44", '"Y,N"')
    text_cf(ws, "D33:D44", ["要確認"])


# =====================================================================================
# 02 / 03 / 04 masters
# =====================================================================================
def build_masters(wb):
    ws = wb["02_Fund_Master"]
    title(ws, "02_Fund_Master　ファンド・指数・Fund-Index Mapping", "ファンドと参照指数（Index Name）は 1 対 1 とは限らないため Mapping 表で関連付ける。新規ファンド・指数は空行に追加。", 20)
    section(ws, 4, "A. ファンドマスタ")
    fm = Table(wb, ws, "FM", 5, [
        C("ID", "Fund ID", "in", 8), C("Name", "Fund Name", "in", 36), C("Short", "略称", "in", 11), C("Class", "Share Class", "in", 12),
        C("Strategy", "Strategy", "in", 12), C("Manager", "運用会社", "in", 16), C("Ccy", "通貨", "in", 6),
        C("Inception", "設定日", "in", 11, FMT_DATE), C("IncNAV", "設定時NAV", "in", 10, FMT_NUM), C("Basis", "NAV Basis", "in", 8),
        C("Dist", "分配型", "in", 6), C("Hedged", "為替ヘッジ", "in", 8), C("Lag", "NAV公表ラグ(営業日)", "in", 9),
        C("BbgTicker", "BBG NAV Ticker", "in", 18), C("Maturity", "償還日", "in", 11, FMT_DATE), C("Status", "Status", "in", 9),
        C("FxKey", "FX Key", "in", 11), C("Notes", "備考", "in", 16),
        C("PrimaryIdx", "Primary Index", "calc", 13, f='=IF({ID}="","",IFERROR(INDEX(MP_Index,MATCH({ID}&"|Y",MP_FundPrim,0)),""))'),
        C("MapCount", "指数数", "calc", 6, f='=IF({ID}="","",COUNTIF(MP_Fund,{ID}))'),
    ], CAP_FUND)
    fm.write([dict(zip(["ID", "Name", "Short", "Class", "Strategy", "Manager", "Ccy", "Inception", "IncNAV", "Basis", "Dist",
                        "Hedged", "Lag", "BbgTicker", "Maturity", "Status", "FxKey"], f)) for f in SD.FUNDS])
    dv_list(ws, f"E6:E{fm.last}", "=L_Strategy")
    dv_list(ws, f"P6:P{fm.last}", "=L_Status")
    T["FM"] = fm

    section(ws, 39, "B. 指数マスタ（Index Name はポジションファイルの表記と完全一致させる）")
    im = Table(wb, ws, "IM", 40, [
        C("ID", "Index ID", "in", 12), C("Name", "Index Name（ファイル表記）", "in", 36), C("Strategy", "Strategy", "in", 12),
        C("Underlying", "原資産", "in", 14), C("UndKey", "原資産 Mkt Key", "in", 12), C("QisKey", "QIS指数 Mkt Key", "in", 13),
        C("VolKey", "Vol Mkt Key", "in", 10), C("Ccy", "通貨", "in", 6), C("GreeksOK", "Greeks単位確認済", "in", 9),
        C("SpotValid", "NDX Spot列 有効", "in", 9), C("Notes", "備考", "in", 14),
        C("OptSeq", "OPT#", "aux", 5, f='=IF(AND({ID}<>"",{Strategy}="OPTION"),COUNTIFS({Strategy@1}:{Strategy},"OPTION",{ID@1}:{ID},"<>"),"")'),
        C("VtSeq", "VT#", "aux", 5, f='=IF(AND({ID}<>"",{Strategy}="VOLTARGET"),COUNTIFS({Strategy@1}:{Strategy},"VOLTARGET",{ID@1}:{ID},"<>"),"")'),
        C("TcSeq", "TC#", "aux", 5, f='=IF(AND({ID}<>"",{Strategy}="TRENDCARRY"),COUNTIFS({Strategy@1}:{Strategy},"TRENDCARRY",{ID@1}:{ID},"<>"),"")'),
    ], CAP_IDX)
    im.write([dict(zip(["ID", "Name", "Strategy", "Underlying", "UndKey", "QisKey", "VolKey", "Ccy", "GreeksOK", "SpotValid", "Notes"], x)) for x in SD.INDICES])
    T["IM"] = im

    section(ws, 74, "C. Fund-Index Mapping（1 ファンド複数指数可。Primary=Y の指数を参照指数として表示）")
    mp = Table(wb, ws, "MP", 75, [
        C("ID", "Map ID", "in", 8), C("Fund", "Fund ID", "in", 8), C("Index", "Index ID", "in", 12), C("Alloc", "配分", "in", 7, "0%"),
        C("Primary", "Primary", "in", 7), C("Notes", "備考", "in", 20),
        C("Strategy", "Strategy", "calc", 12, f='=IF({Fund}="","",IFERROR(INDEX(IM_Strategy,MATCH({Index},IM_ID,0)),"UNKNOWN"))'),
        C("First", "First", "aux", 6, f='=IF({Fund}="","",IF(COUNTIF({Fund@1}:{Fund},{Fund})=1,1,0))'),
        C("Valid", "Valid", "aux", 6, f='=IF({Fund}="","",IF(AND(COUNTIF(FM_ID,{Fund})>0,COUNTIF(IM_ID,{Index})>0),1,0))'),
        C("FundPrim", "Fund|Primary", "aux", 12, f='=IF({Fund}="","",{Fund}&"|"&{Primary})'),
        C("IdxPrim", "Index|Primary", "aux", 14, f='=IF({Fund}="","",{Index}&"|"&{Primary})'),
    ], CAP_MAP)
    mp.write([dict(zip(["ID", "Fund", "Index", "Alloc", "Primary", "Notes"], m)) for m in SD.MAPPINGS])
    T["MP"] = mp
    ws.freeze_panes = "B6"

    ws = wb["03_Client_Master"]
    title(ws, "03_Client_Master　地銀・社内営業担当者", "顧客情報は社内限り。外部サービスへ送信しない。", 10)
    section(ws, 4, "A. 地銀（投資家）")
    cl = Table(wb, ws, "CL", 5, [
        C("ID", "Client ID", "in", 9), C("Bank", "銀行名", "in", 14), C("Type", "業態", "in", 12), C("Region", "地域", "in", 8),
        C("Rep", "担当 Rep ID", "in", 9), C("LastContact", "前回コンタクト", "in", 12, FMT_DATE),
        C("NextFollow", "次回フォロー予定", "in", 12, FMT_DATE), C("Notes", "備考", "in", 20),
        C("RepName", "担当者名", "calc", 12, f='=IF({ID}="","",IFERROR(INDEX(RP_Name,MATCH({Rep},RP_ID,0)),"#未登録"))'),
    ], CAP_CLIENT)
    cl.write([dict(zip(["ID", "Bank", "Type", "Region", "Rep", "LastContact", "NextFollow", "Notes"], c)) for c in SD.CLIENTS])
    T["CL"] = cl
    section(ws, 59, "B. 社内営業担当者")
    rp = Table(wb, ws, "RP", 60, [C("ID", "Rep ID", "in", 9), C("Name", "氏名", "in", 14), C("Email", "メール", "in", 28),
                                   C("Dept", "部署", "in", 18), C("Active", "Active", "in", 7)], CAP_REP)
    rp.write([dict(zip(["ID", "Name", "Email", "Dept", "Active"], r)) for r in SD.REPS])
    T["RP"] = rp

    ws = wb["04_Client_Holdings"]
    title(ws, "04_Client_Holdings　地銀別保有ファンド", "保有口数・簿価・取得日を入力。時価・取得来リターンは数式。", 22)
    hd = Table(wb, ws, "HD", 5, [
        C("ID", "Holding ID", "in", 9), C("Client", "Client ID", "in", 8), C("Fund", "Fund ID", "in", 7),
        C("Units", "保有口数", "in", 12, FMT_NUM), C("Book", "簿価（円）", "in", 15, FMT_INT), C("EntryDate", "取得日", "in", 11, FMT_DATE),
        C("EntryNAV", "取得時NAV", "in", 10, FMT_NUM), C("Redemption", "償還・満期予定", "in", 11, FMT_DATE), C("Notes", "備考", "in", 12),
        C("Bank", "銀行名", "calc", 10, f='=IF({Client}="","",IFERROR(INDEX(CL_Bank,MATCH({Client},CL_ID,0)),"#未登録"))'),
        C("RepID", "Rep ID", "calc", 7, f='=IF({Client}="","",IFERROR(INDEX(CL_Rep,MATCH({Client},CL_ID,0)),""))'),
        C("RepName", "担当者", "calc", 10, f='=IF({Client}="","",IFERROR(INDEX(RP_Name,MATCH({RepID},RP_ID,0)),"#未登録"))'),
        C("FundName", "ファンド", "calc", 11, f='=IF({Fund}="","",IFERROR(INDEX(FM_Short,MATCH({Fund},FM_ID,0)),"#未登録"))'),
        C("Strategy", "Strategy", "calc", 11, f='=IF({Fund}="","",IFERROR(INDEX(FM_Strategy,MATCH({Fund},FM_ID,0)),""))'),
        C("Ccy", "通貨", "calc", 5, f='=IF({Fund}="","",IFERROR(INDEX(FM_Ccy,MATCH({Fund},FM_ID,0)),""))'),
        C("LatestNAV", "最新NAV", "calc", 10, FMT_NUM, f='=IF({Fund}="","",IFERROR(INDEX(PF_NAV,MATCH({Fund},PF_ID,0)),""))'),
        C("MV", "時価（ファンド通貨）", "calc", 15, FMT_INT, f='=IF(AND(ISNUMBER({Units}),ISNUMBER({LatestNAV})),{Units}*{LatestNAV},"")'),
        C("EntryTRI", "取得時TRI", "aux", 9, FMT_4, f='=IF({Fund}="","",IFERROR(INDEX(NV_TRI,MATCH({Fund}&"|OFFICIAL|"&MAXIFS(NV_Date,NV_Fund,{Fund},NV_Source,"OFFICIAL",NV_Date,"<="&{EntryDate}),NV_Key,0)),""))'),
        C("SinceEntry", "取得来リターン", "calc", 10, FMT_PCT, f='=IF({Fund}="","",IFERROR(IF(AND(ISNUMBER({EntryTRI}),ISNUMBER(INDEX(PF_TRI,MATCH({Fund},PF_ID,0)))),INDEX(PF_TRI,MATCH({Fund},PF_ID,0))/{EntryTRI}-1,IF(AND(ISNUMBER({LatestNAV}),ISNUMBER({EntryNAV})),{LatestNAV}/{EntryNAV}-1,"")),""))'),
        C("WeeklyRet", "週次", "calc", 8, FMT_PCT, f='=IF({Fund}="","",IFERROR(INDEX(PF_WeeklyRet,MATCH({Fund},PF_ID,0)),""))'),
        C("DaysToRed", "償還まで日数", "calc", 8, FMT_INT, f='=IF(ISNUMBER({Redemption}),{Redemption}-EvalDate,"")'),
        C("Valid", "Valid", "aux", 5, f='=IF({ID}="","",IF(AND(COUNTIF(CL_ID,{Client})>0,COUNTIF(FM_ID,{Fund})>0),1,0))'),
    ], CAP_HOLD)
    hd.write([dict(zip(["ID", "Client", "Fund", "Units", "Book", "EntryDate", "EntryNAV", "Redemption", "Notes"], h)) for h in SD.holdings()])
    T["HD"] = hd
    ws.freeze_panes = "B6"


# =====================================================================================
# 05 / 06 positions
# =====================================================================================
RAW_KEYS = ["PositionDate", "IndexName", "Name", "Ticker", "Mid", "Strike", "Maturity", "Notional", "Delta", "Gamma",
            "Vega", "Theta", "IV", "Spot", "IndexLevel", "Weight", "OpenDate"]
RAW_FMT = [FMT_DATE, None, None, None, FMT_NUM, FMT_NUM, FMT_DATE, FMT_INT, FMT_4, FMT_6, FMT_NUM, FMT_NUM, FMT_PCT1,
           FMT_NUM, FMT_NUM, FMT_PCT1, FMT_DATE]
RAW_W = [11, 30, 20, 28, 9, 9, 11, 13, 7, 9, 7, 7, 7, 10, 9, 7, 11]


def pos_row(d):
    return {k: d[c] for k, c in zip(RAW_KEYS, SD.POS_COLUMNS)}


def build_positions(wb):
    ws = wb["05_Position_Raw"]
    title(ws, "05_Position_Raw　最新ポジション（Power Query 出力先）",
          "Power Query（qPositionLatest）の読込先。列名は 19_Settings の tblColMap で標準化。手貼り付けも可（値のみ）。取込後に QIS_ArchivePositions で 06 へ保存。", 22)
    cols = [C(k, h, "in", w, f) for k, h, w, f in zip(RAW_KEYS, SD.POS_COLUMNS, RAW_W, RAW_FMT)]
    cols += [C("SourceFile", "Source File", "in", 22), C("ImportedAt", "Imported At", "in", 16, "yyyy/mm/dd hh:mm"),
             C("Check", "取込チェック", "calc", 18, f='=IF({PositionDate}="","",IF({Ticker}="","Ticker欠損",IF(COUNTIF(IM_Name,{IndexName})=0,"Index未登録",IF(COUNTIFS(PR_PositionDate,{PositionDate},PR_IndexName,{IndexName},PR_Ticker,{Ticker},PR_OpenDate,{OpenDate},PR_Notional,{Notional})>1,"重複行","OK"))))')]
    pr = Table(wb, ws, "PR", 5, cols, CAP_POSRAW)
    (fname, latest) = SD.positions()[-1]
    imported = dt.datetime(2026, 11, 2, 8, 30)
    pr.write([dict(pos_row(d), SourceFile=fname, ImportedAt=imported) for d in latest])
    T["PR"] = pr
    ws.freeze_panes = "E6"

    ws = wb["06_Position_History"]
    title(ws, "06_Position_History　ポジション履歴（値で保存）＋標準化（数式）",
          "A〜T 列は取込時点の値を保存（元ファイルが後で更新されても再現可能）。U 列以降は標準化・判定の数式。同一 Snapshot ID の再保存は QIS_ArchivePositions がスキップ。", 30)
    cols = [C("SnapshotID", "Snapshot ID", "in", 30), C("ArchivedAt", "Archived At", "in", 15, "yyyy/mm/dd hh:mm"),
            C("SourceFile", "Source File", "in", 20)]
    cols += [C(k, h, "in", w, f) for k, h, w, f in zip(RAW_KEYS, SD.POS_COLUMNS, RAW_W, RAW_FMT)]
    PT = 'FIND(" ",{Ticker},FIND("/",{Ticker}))'
    calc = [
        C("IndexID", "Index ID", "calc", 12, f='=IF({PositionDate}="","",IFERROR(INDEX(IM_ID,MATCH({IndexName},IM_Name,0)),"UNMAPPED"))'),
        C("Underlying", "Underlying", "calc", 8, f='=IF({Ticker}="","",IFERROR(LEFT({Ticker},FIND(" ",{Ticker})-1),{Ticker}))'),
        C("OptType", "Option Type", "calc", 8, f=('=IF({PositionDate}="","",IFERROR(IF(AND(MID({Ticker},P+1,1)="P",ISNUMBER(--MID({Ticker},P+2,1))),"PUT",'
                                                'IF(AND(MID({Ticker},P+1,1)="C",ISNUMBER(--MID({Ticker},P+2,1))),"CALL",1/0)),'
                                                'IF(ISNUMBER(SEARCH(" PUT "," "&{Name}&" ")),"PUT",IF(ISNUMBER(SEARCH(" CALL "," "&{Name}&" ")),"CALL","UNKNOWN"))))').replace("P+", PT + "+")),
        C("AggKey", "Position Key（集約）", "calc", 34, f='=IF({PositionDate}="","",{IndexID}&"|"&{Ticker})'),
        C("TradeSeq", "Trade Seq", "aux", 6, f='=IF({PositionDate}="","",COUNTIFS({PositionDate@1}:{PositionDate},{PositionDate},{AggKey@1}:{AggKey},{AggKey}))'),
        C("TradeKey", "Trade Key（個別取引）", "aux", 40, f='=IF({PositionDate}="","",{AggKey}&"|"&TEXT({OpenDate},"yyyymmdd")&"|"&{TradeSeq})'),
        C("DTE", "Days to Expiry", "calc", 7, FMT_INT, f='=IF({PositionDate}="","",{Maturity}-{PositionDate})'),
        C("BizDTE", "営業日", "calc", 6, FMT_INT, f='=IF({PositionDate}="","",IF({Maturity}<{PositionDate},"",NETWORKDAYS({PositionDate},{Maturity},Holidays)-1))'),
        C("Dist", "Strike Distance", "calc", 9, FMT_PCT, f='=IF({PositionDate}="","",IF(N({Spot})<=0,"",({Spot}-{Strike})/{Spot}))'),
        C("AbsDist", "|Dist|", "aux", 8, FMT_PCT, f='=IF({Dist}="","",ABS({Dist}))'),
        C("Mny", "Moneyness", "calc", 9, f='=IF({PositionDate}="","",IF(OR({Dist}="",{OptType}="UNKNOWN"),"UNKNOWN",IF({AbsDist}<=S_AtmBand,"ATM",IF({OptType}="PUT",IF({Dist}>0,"OTM","ITM"),IF({Dist}<0,"OTM","ITM")))))'),
        C("Flag", "Snapshot", "calc", 8, f='=IF({PositionDate}="","",IFERROR(IF({PositionDate}=INDEX(OA_CurDate,MATCH({IndexID},OA_ID,0)),"CUR",IF({PositionDate}=INDEX(OA_PrevDate,MATCH({IndexID},OA_ID,0)),"PRV","OLD")),"UNMAPPED"))'),
        C("Use", "Position Status", "calc", 11, f='=IF({Flag}="CUR",IF({Maturity}>=EvalDate,"CUR","CUR_EXPIRED"),IF({Flag}="PRV",IF({Maturity}>={PositionDate},"PRV","PRV_EXPIRED"),""))'),
        C("KeyFlag", "Key|Status", "aux", 20, f='=IF({PositionDate}="","",{AggKey}&"|"&{Use})'),
        C("NearKey", "Nearest", "aux", 8, f='=IF({Use}<>"CUR","",IF({AbsDist}=INDEX(OA_Nearest,MATCH({IndexID},OA_ID,0)),{IndexID}&"|N",""))'),
        C("MatKey", "NearMat", "aux", 8, f='=IF({Use}<>"CUR","",IF({Maturity}=INDEX(OA_NearMat,MATCH({IndexID},OA_ID,0)),{IndexID}&"|M",""))'),
        C("AbsN", "|Notional|", "aux", 12, FMT_INT, f='=IF({PositionDate}="","",ABS({Notional}))'),
        C("AbsW", "|Weight|", "aux", 7, FMT_PCT1, f='=IF({PositionDate}="","",ABS({Weight}))'),
        C("IVxN", "IV×|N|", "aux", 10, FMT_INT, f='=IF({PositionDate}="","",{IV}*{AbsN})'),
        C("DistxN", "Dist×|N|", "aux", 10, FMT_INT, f='=IF({Dist}="","",{Dist}*{AbsN})'),
        C("DeltaN", "Δ×N(参考)", "aux", 11, FMT_INT, f='=IF({PositionDate}="","",{Delta}*{Notional})'),
        C("GammaN", "Γ×N(参考)", "aux", 9, FMT_NUM, f='=IF({PositionDate}="","",{Gamma}*{Notional})'),
        C("VegaN", "V×N(参考)", "aux", 11, FMT_INT, f='=IF({PositionDate}="","",{Vega}*{Notional})'),
        C("ThetaN", "Θ×N(参考)", "aux", 11, FMT_INT, f='=IF({PositionDate}="","",{Theta}*{Notional})'),
        C("Univ", "比較対象", "aux", 5, f='=IF(OR({Use}="CUR",{Use}="PRV"),1,0)'),
        C("First", "初出", "aux", 5, f='=IF({Univ}=1,IF(COUNTIFS({AggKey@1}:{AggKey},{AggKey},{Univ@1}:{Univ},1)=1,1,0),0)'),
        C("ChgSeq", "比較#", "aux", 5, f='=IF({First}=1,COUNTIF({First@1}:{First},1),"")'),
        C("CurSeq", "現在#", "aux", 5, f='=IF({Use}="CUR",COUNTIF({Use@1}:{Use},"CUR"),"")'),
        C("Dup", "重複数", "aux", 5, f='=IF({PositionDate}="","",COUNTIFS(PH_PositionDate,{PositionDate},PH_IndexName,{IndexName},PH_Ticker,{Ticker},PH_OpenDate,{OpenDate},PH_Notional,{Notional}))'),
        C("Multi", "同一銘柄取引数", "aux", 6, f='=IF({PositionDate}="","",COUNTIFS(PH_PositionDate,{PositionDate},PH_AggKey,{AggKey}))'),
        C("SrcConflict", "別ファイル同日", "aux", 6, f='=IF({PositionDate}="","",IF(COUNTIFS(PH_PositionDate,{PositionDate},PH_IndexName,{IndexName},PH_SourceFile,"<>"&{SourceFile})>0,1,0))'),
    ]
    ph = Table(wb, ws, "PH", 5, cols + calc, CAP_PH)
    rows = []
    for fname, lst in SD.positions():
        pdate = lst[0]["Position Date"]
        archived = dt.datetime.combine(pdate + dt.timedelta(days=3), dt.time(8, 45))
        for d in lst:
            rows.append(dict(pos_row(d), SnapshotID=f"{pdate:%Y%m%d}|{d['Index Name']}|{fname}", ArchivedAt=archived, SourceFile=fname))
    ph.write(rows)
    T["PH"] = ph
    ws.freeze_panes = "H6"


# =====================================================================================
# 07 Bloomberg
# =====================================================================================
FIELD_MAP = [
    ("指数・株価・為替・金利 終値", "PX_LAST", "BDH／BDP", "評価基準日の値は BDH（Fill=N）。BDP はライブ値のため参考のみ", "確認要（端末 FLDS）"),
    ("ファンドNAV", "FUND_NET_ASSET_VAL／PX_LAST", "BDH", "ファンド Ticker によりフィールドが異なる。公式 NAV を優先し照合用に使用", "確認要"),
    ("NAV 日付", "NAV_DATE／LAST_UPDATE_DT", "BDP", "基準日の確認用", "確認要"),
    ("分配金履歴", "DVD_HIST_ALL", "BDS", "分配型ファンドの分配金調整（Ex-Date・金額）", "確認要"),
    ("オプション Mid", "PX_MID", "BDH／BDP", "ポジションファイルの Option Mid Price と照合（基準日一致時のみ）", "確認要"),
    ("Delta / Gamma / Vega / Theta", "OPT_DELTA_MID／OPT_GAMMA_MID／OPT_VEGA_MID／OPT_THETA_MID", "BDP", "多くはライブ値のみ。過去時点の Greeks は取得できない可能性", "確認要"),
    ("Implied Vol", "IVOL_MID／OPT_IMPLIED_VOLATILITY_MID", "BDP", "単位（%表示か小数か）を確認", "確認要"),
    ("オプション属性", "OPT_STRIKE_PX／OPT_EXPIRE_DT／OPT_PUT_CALL／OPT_UNDL_TICKER", "BDP", "Ticker 解析結果（Put/Call・Strike・満期）の検証", "確認要"),
    ("実現ボラティリティ", "VOLATILITY_30D／VOLATILITY_260D", "BDP／BDH", "年率 % 表示", "確認要"),
    ("ATM インプライドVol", "30DAY_IMPVOL_100.0%MNY_DF", "BDP／BDH", "指数オプションの ATM IV", "確認要"),
    ("QIS 指数レベル", "PX_LAST", "BDH", "QIS 指数の Bloomberg Ticker を運用会社・指数提供者に確認", "確認要（架空 Ticker を使用中）"),
    ("指数構成", "INDX_MWEIGHT", "BDS", "必要な場合のみ", "将来拡張"),
]


def build_bloomberg(wb, ws):
    title(ws, "07_Bloomberg_Data　市場データ（Bloomberg Excel Add-in）",
          "H〜J 列は BDH 数式（Add-in 未接続時は NO_CONN）。K〜N 列は手入力・CSV 代替。採用値(P〜R)は 19_Settings の DataMode に従い、欠損・古い値は最新値として扱わない。Bloomberg 由来データの顧客向け転送は不可。", 26)
    bdh = lambda d, fill: f'=IF({{Ticker}}="","",IFERROR(BDH({{Ticker}},{{Field}},{d},{d},"Dates=H","Fill={fill}"),"NO_CONN"))'
    md = Table(wb, ws, "MD", 5, [
        C("Key", "Key", "in", 11), C("Desc", "説明", "in", 20), C("Class", "Asset Class", "in", 10), C("Ticker", "Bloomberg Ticker", "in", 22),
        C("Field", "Field", "in", 14), C("Critical", "重要", "in", 5), C("Unit", "Unit", "in", 7),
        C("BbgCur", "BBG 評価基準日", "calc", 11, FMT_NUM, f=bdh("EvalDate", "N")),
        C("BbgPrev", "BBG 前週", "calc", 11, FMT_NUM, f=bdh("PrevEvalDate", "P")),
        C("BbgMEnd", "BBG 前月末", "calc", 11, FMT_NUM, f=bdh("MonthBaseDate", "P")),
        C("ManCur", "手入力 評価基準日", "in", 11, FMT_NUM), C("ManPrev", "手入力 前週", "in", 11, FMT_NUM),
        C("ManMEnd", "手入力 前月末", "in", 11, FMT_NUM), C("ManAsOf", "手入力 As-of", "in", 11, FMT_DATE), C("ManNote", "手入力 備考", "in", 16),
        C("UsedCur", "採用 評価基準日", "calc", 11, FMT_NUM, f='=IF({Key}="","",IF(S_DataMode="BLOOMBERG",IF(ISNUMBER({BbgCur}),{BbgCur},""),IF(ISNUMBER({ManCur}),{ManCur},"")))'),
        C("UsedPrev", "採用 前週", "calc", 11, FMT_NUM, f='=IF({Key}="","",IF(S_DataMode="BLOOMBERG",IF(ISNUMBER({BbgPrev}),{BbgPrev},""),IF(ISNUMBER({ManPrev}),{ManPrev},"")))'),
        C("UsedMEnd", "採用 前月末", "calc", 11, FMT_NUM, f='=IF({Key}="","",IF(S_DataMode="BLOOMBERG",IF(ISNUMBER({BbgMEnd}),{BbgMEnd},""),IF(ISNUMBER({ManMEnd}),{ManMEnd},"")))'),
        C("AsOf", "As-of", "calc", 11, FMT_DATE, f='=IF({Key}="","",IF(S_DataMode="BLOOMBERG",IF(ISNUMBER({BbgCur}),EvalDate,""),{ManAsOf}))'),
        C("Status", "Status", "calc", 9, f='=IF({Key}="","",IF({UsedCur}="","MISSING",IF(OR({AsOf}="",{AsOf}<EvalDate-S_MktStaleDays),"STALE","OK")))'),
        C("Value", "検証済 値", "calc", 11, FMT_NUM, f='=IF({Status}="OK",{UsedCur},"")'),
        C("WoW", "前週比", "calc", 9, "0.00%;-0.00%", f='=IF(AND(ISNUMBER({Value}),ISNUMBER({UsedPrev})),IF(OR({Unit}="PRICE",{Unit}="FX"),{Value}/{UsedPrev}-1,{Value}-{UsedPrev}),"")'),
        C("MoM", "前月末比", "calc", 9, "0.00%;-0.00%", f='=IF(AND(ISNUMBER({Value}),ISNUMBER({UsedMEnd})),IF(OR({Unit}="PRICE",{Unit}="FX"),{Value}/{UsedMEnd}-1,{Value}-{UsedMEnd}),"")'),
        C("Display", "表示文", "calc", 34, f=('=IF({Key}="","",IF({Value}="",{Desc}&"：データ"&IF({Status}="STALE","更新なし(STALE)","欠損"),{Desc}&" "&TEXT({Value},IF({Unit}="FX","#,##0.00##","#,##0.00"))&"（"&'
                                             'IF({WoW}="","前週比 不明",IF(OR({Unit}="PRICE",{Unit}="FX"),TEXT({WoW},"+0.00%;-0.00%"),IF(OR({Unit}="YIELD"),TEXT({WoW}*100,"+0.0;-0.0")&"bp",IF({Unit}="SPREAD",TEXT({WoW},"+0.0;-0.0")&"bp",TEXT({WoW},"+0.0;-0.0")&"pt"))))&"）"))')),
        C("Live", "BDP ライブ値（参考・レポート不使用）", "aux", 12, FMT_NUM, f='=IF({Ticker}="","",IFERROR(BDP({Ticker},{Field}),"NO_CONN"))'),
    ], CAP_MD)
    md.write([dict(Key=m[0], Desc=m[1], Class=m[2], Ticker=m[3], Field=m[4], Critical=m[5], Unit=m[6], ManCur=m[7],
                   ManPrev=m[8], ManMEnd=m[9], ManAsOf=m[10], ManNote=m[11]) for m in SD.MARKET])
    text_cf(ws, f"{md.col('Status')}{md.first}:{md.col('Status')}{md.last}", ["MISSING", "STALE"], strong=False)
    T["MD"] = md
    ws.freeze_panes = "B6"

    # Field map (right side)
    c0 = 28  # AB
    section(ws, 4, "B. Bloomberg フィールドマッピング（端末の FLDS で要確認）", col=c0)
    fmap = Table(wb, ws, "BF", 5, [C("Item", "データ項目", "in", 22), C("Field", "Field Mnemonic", "in", 30), C("Func", "関数", "in", 9),
                                   C("Note", "用途・注意", "in", 40), C("Verify", "確認状況", "in", 16)], 20, first_col=c0)
    fmap.write([dict(zip(["Item", "Field", "Func", "Note", "Verify"], f)) for f in FIELD_MAP])

    # Series (long table)
    c1 = c0 + 6
    section(ws, 30, "C. 時系列（週次。BDH 出力を値で保存、または CSV。チャート・参照指数用）", col=c0)
    sr = Table(wb, ws, "SR", 31, [C("Key", "Key", "in", 11), C("Date", "Date", "in", 11, FMT_DATE), C("Value", "Value", "in", 11, FMT_NUM),
                                  C("Source", "Source", "in", 10),
                                  C("KD", "Key|Date", "aux", 18, f='=IF({Key}="","",{Key}&"|"&{Date})')], CAP_SER, first_col=c0)
    sr.write([dict(Key=a, Date=b, Value=c, Source=d) for a, b, c, d in SD.series()])
    T["SR"] = sr

    # Option BBG reconciliation (requires Bloomberg)
    section(ws, 4, "D. オプション照合（Bloomberg 接続時のみ。時点が異なる場合は『時点不一致』と表示）", col=c1)
    orc = Table(wb, ws, "OB", 5, [
        C("n", "#", "aux", 4),
        C("Ticker", "Ticker", "calc", 28, f='=IFERROR(INDEX(PH_Ticker,MATCH({n},PH_CurSeq,0)),"")'),
        C("PosDate", "Position Date", "calc", 11, FMT_DATE, f='=IF({Ticker}="","",INDEX(PH_PositionDate,MATCH({n},PH_CurSeq,0)))'),
        C("FileMid", "File Mid", "calc", 9, FMT_NUM, f='=IF({Ticker}="","",INDEX(PH_Mid,MATCH({n},PH_CurSeq,0)))'),
        C("BbgMid", "BBG PX_MID(基準日)", "calc", 10, FMT_NUM, f='=IF({Ticker}="","",IFERROR(BDH({Ticker},"PX_MID",{PosDate},{PosDate},"Dates=H","Fill=N"),"NO_CONN"))'),
        C("MidDiff", "差", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({FileMid}),ISNUMBER({BbgMid})),{BbgMid}/{FileMid}-1,"")'),
        C("FileIV", "File IV", "calc", 7, FMT_PCT1, f='=IF({Ticker}="","",INDEX(PH_IV,MATCH({n},PH_CurSeq,0)))'),
        C("BbgIV", "BBG IVOL_MID(ライブ)", "calc", 9, FMT_NUM, f='=IF({Ticker}="","",IFERROR(BDP({Ticker},"IVOL_MID"),"NO_CONN"))'),
        C("FileDelta", "File Delta", "calc", 7, FMT_4, f='=IF({Ticker}="","",INDEX(PH_Delta,MATCH({n},PH_CurSeq,0)))'),
        C("BbgDelta", "BBG Delta(ライブ)", "calc", 8, FMT_4, f='=IF({Ticker}="","",IFERROR(BDP({Ticker},"OPT_DELTA_MID"),"NO_CONN"))'),
        C("Timing", "時点", "calc", 12, f='=IF({Ticker}="","",IF(RunDate>{PosDate},"時点不一致（Greeks/IVはライブ値）","同時点"))'),
    ], 60, first_col=c1)
    orc.write([dict(n=i + 1) for i in range(60)])


# =====================================================================================
# 08 NAV
# =====================================================================================
def build_nav(wb, ws):
    title(ws, "08_NAV_History　NAV 履歴・分配金・NAV 照合",
          "公式 NAV（OFFICIAL）を優先し、Bloomberg NAV（BLOOMBERG）は照合用。TRI＝分配金を足し戻したトータルリターン指数（初回=100）。", 30)
    nv = Table(wb, ws, "NV", 5, [
        C("Fund", "Fund ID", "in", 7), C("Date", "NAV Date", "in", 11, FMT_DATE), C("NAV", "NAV", "in", 11, FMT_NUM),
        C("Ccy", "通貨", "in", 5), C("Source", "Source", "in", 10), C("Type", "NAV Type", "in", 9), C("Basis", "Basis", "in", 6),
        C("Class", "Share Class", "in", 11), C("LoadedAt", "Loaded At", "in", 14, "yyyy/mm/dd hh:mm"),
        C("Key", "Key", "aux", 22, f='=IF({Fund}="","",{Fund}&"|"&{Source}&"|"&{Date})'),
        C("PrevDate", "前回日付", "aux", 11, FMT_DATE, f='=IF({Fund}="","",MAXIFS(NV_Date,NV_Fund,{Fund},NV_Source,{Source},NV_Date,"<"&{Date}))'),
        C("PrevNAV", "前回NAV", "aux", 11, FMT_NUM, f='=IF(OR({Fund}="",N({PrevDate})=0),"",SUMIFS(NV_NAV,NV_Key,{Fund}&"|"&{Source}&"|"&{PrevDate}))'),
        C("Dist", "期間分配金", "aux", 8, FMT_NUM, f='=IF(OR({Fund}="",N({PrevDate})=0),"",SUMIFS(DS_Amount,DS_Fund,{Fund},DS_ExDate,">"&{PrevDate},DS_ExDate,"<="&{Date}))'),
        C("Ret", "期間TR", "calc", 8, FMT_PCT, f='=IF(OR({Fund}="",N({PrevDate})=0),"",({NAV}+{Dist})/{PrevNAV}-1)'),
        C("TRI", "TRI", "calc", 10, FMT_4, f='=IF({Fund}="","",IF(N({PrevDate})=0,100,INDEX(NV_TRI,MATCH({Fund}&"|"&{Source}&"|"&{PrevDate},NV_Key,0))*(1+{Ret})))'),
        C("Peak", "TRI高値", "aux", 10, FMT_4, f='=IF({Fund}="","",MAXIFS(NV_TRI,NV_Fund,{Fund},NV_Source,{Source},NV_Date,"<="&{Date}))'),
        C("DD", "Drawdown", "calc", 8, FMT_PCT, f='=IF({Fund}="","",{TRI}/{Peak}-1)'),
        C("Win", "Vol窓", "aux", 5, f='=IF({Fund}="",0,IF(AND({Source}="OFFICIAL",{Date}>EvalDate-7*S_VolWeeks,{Date}<=EvalDate,ISNUMBER({Ret})),1,0))'),
        C("RW", "r(窓)", "aux", 8, FMT_PCT, f='=IF({Win}=1,{Ret},"")'),
        C("R2", "r²(窓)", "aux", 9, FMT_6, f='=IF({Win}=1,{Ret}^2,"")'),
        C("Dup", "重複", "aux", 5, f='=IF({Fund}="","",COUNTIF(NV_Key,{Key}))'),
    ], CAP_NAV)
    off, bbg, dists = SD.nav_history()
    loaded = dt.datetime(2026, 11, 2, 8, 40)
    rows = [dict(Fund=r[0], Date=r[1], NAV=r[2], Ccy=r[3], Source=r[4], Type=r[5], Basis=r[6], Class=r[7], LoadedAt=loaded) for r in off + bbg]
    nv.write(rows)
    T["NV"] = nv
    ws.freeze_panes = "C6"

    c0 = 24  # X
    section(ws, 4, "B. 分配金（運用会社通知。Ex-Date 基準で TR に加算）", col=c0)
    ds = Table(wb, ws, "DS", 5, [C("Fund", "Fund ID", "in", 7), C("ExDate", "Ex-Date", "in", 11, FMT_DATE), C("Amount", "分配金/口", "in", 9, FMT_NUM),
                                 C("Ccy", "通貨", "in", 5), C("Source", "出所", "in", 24)], CAP_DIST, first_col=c0)
    ds.write([dict(Fund=a, ExDate=b, Amount=c, Ccy=d, Source=e) for a, b, c, d, e in dists])
    T["DS"] = ds

    c1 = c0 + 6
    section(ws, 4, "C. NAV 照合（評価基準日。同一基準日・同一通貨・同一 Basis の場合のみ差額判定）", col=c1)
    rc = Table(wb, ws, "RC", 5, [
        C("Fund", "Fund ID", "calc", 7, f="=IF(INDEX(FM_ID,{n})=\"\",\"\",INDEX(FM_ID,{n}))"),
        C("OffDate", "公式NAV日付", "calc", 11, FMT_DATE, f='=IF({Fund}="","",MAXIFS(NV_Date,NV_Fund,{Fund},NV_Source,"OFFICIAL",NV_Date,"<="&EvalDate))'),
        C("OffNAV", "公式NAV", "calc", 10, FMT_NUM, f='=IF(N({OffDate})=0,"",SUMIFS(NV_NAV,NV_Key,{Fund}&"|OFFICIAL|"&{OffDate}))'),
        C("OffCcy", "公式通貨", "calc", 6, f='=IF(N({OffDate})=0,"",INDEX(NV_Ccy,MATCH({Fund}&"|OFFICIAL|"&{OffDate},NV_Key,0)))'),
        C("OffBasis", "公式Basis", "calc", 6, f='=IF(N({OffDate})=0,"",INDEX(NV_Basis,MATCH({Fund}&"|OFFICIAL|"&{OffDate},NV_Key,0)))'),
        C("BbgDate", "BBG NAV日付", "calc", 11, FMT_DATE, f='=IF({Fund}="","",MAXIFS(NV_Date,NV_Fund,{Fund},NV_Source,"BLOOMBERG",NV_Date,"<="&EvalDate))'),
        C("BbgNAV", "BBG NAV(最新)", "calc", 10, FMT_NUM, f='=IF(N({BbgDate})=0,"",SUMIFS(NV_NAV,NV_Key,{Fund}&"|BLOOMBERG|"&{BbgDate}))'),
        C("BbgCcy", "BBG通貨", "calc", 6, f='=IF(N({BbgDate})=0,"",INDEX(NV_Ccy,MATCH({Fund}&"|BLOOMBERG|"&{BbgDate},NV_Key,0)))'),
        C("BbgBasis", "BBG Basis", "calc", 6, f='=IF(N({BbgDate})=0,"",INDEX(NV_Basis,MATCH({Fund}&"|BLOOMBERG|"&{BbgDate},NV_Key,0)))'),
        C("BbgSame", "BBG NAV(公式と同日)", "calc", 10, FMT_NUM, f='=IF(N({OffDate})=0,"",IF(COUNTIF(NV_Key,{Fund}&"|BLOOMBERG|"&{OffDate})=0,"",SUMIFS(NV_NAV,NV_Key,{Fund}&"|BLOOMBERG|"&{OffDate})))'),
        C("Diff", "差(BBG/公式-1)", "calc", 9, "0.000%", f='=IF(AND(ISNUMBER({OffNAV}),ISNUMBER({BbgSame})),{BbgSame}/{OffNAV}-1,"")'),
        C("Status", "照合結果", "calc", 15, f=('=IF({Fund}="","",IF(OR(N({OffDate})=0,N({BbgDate})=0),"MISSING DATA",IF(OR({OffCcy}<>{BbgCcy},{OffBasis}<>{BbgBasis}),"NOT COMPARABLE",'
                                               'IF({BbgSame}="","DATE DIFFERENCE",IF(ABS({Diff})<=S_NavMatchTol,"MATCH",IF(ABS({Diff})<=S_NavWarnTol,"WARNING","MISMATCH"))))))')),
        C("Comment", "コメント", "calc", 40, f=('=IF({Fund}="","",IF({Status}="MISSING DATA",IF(N({OffDate})=0,"公式NAVなし","Bloomberg NAVなし")&"（比較不能）",IF({Status}="NOT COMPARABLE","通貨/Basis不一致（公式 "&{OffCcy}&"/"&{OffBasis}&"、BBG "&{BbgCcy}&"/"&{BbgBasis}&"）",'
                                                'IF({Status}="DATE DIFFERENCE","日付不一致（公式 "&TEXT({OffDate},"mm/dd")&"、BBG "&TEXT({BbgDate},"mm/dd")&"）。差額判定せず",'
                                                '"公式 "&TEXT({OffNAV},"#,##0.00")&" vs BBG "&TEXT({BbgSame},"#,##0.00")&"（"&TEXT({Diff},"+0.00%;-0.00%")&"、"&TEXT({OffDate},"mm/dd")&"）"))))')),
        C("ExpPub", "公表予定日", "calc", 11, FMT_DATE, f='=IF({Fund}="","",WORKDAY(EvalDate,INDEX(FM_Lag,{n}),Holidays))'),
        C("Delay", "NAV遅延", "calc", 6, f='=IF({Fund}="","",IF(INDEX(FM_Status,{n})<>"ACTIVE",0,IF(AND(N({OffDate})<EvalDate-S_NavDateTol,RunDate>{ExpPub}),1,0)))'),
        C("CcyOK", "通貨=マスタ", "calc", 6, f='=IF(OR({Fund}="",{OffCcy}=""),"",IF({OffCcy}=INDEX(FM_Ccy,{n}),1,0))'),
    ], CAP_FUND, first_col=c1)
    rc.write()
    text_cf(ws, f"{rc.col('Status')}{rc.first}:{rc.col('Status')}{rc.last}", ["MISMATCH"])
    text_cf(ws, f"{rc.col('Status')}{rc.first}:{rc.col('Status')}{rc.last}", ["WARNING", "DATE DIFFERENCE", "MISSING DATA", "NOT COMPARABLE"], strong=False)
    T["RC"] = rc


# =====================================================================================
# 09 Performance
# =====================================================================================
def build_perf(wb, ws):
    title(ws, "09_Performance　パフォーマンス・寄与度",
          "リターンは TRI（分配金調整後）の比率。週次は評価基準日と前週基準日の 2 時点の公式 NAV がそろう場合のみ算出（欠損時は空欄＝前回値で代用しない）。", 32)
    tri = lambda d: f'INDEX(NV_TRI,MATCH({{ID}}&"|OFFICIAL|"&{d},NV_Key,0))'
    pf = Table(wb, ws, "PF", 5, [
        C("ID", "Fund ID", "calc", 7, f='=IF(INDEX(FM_ID,{n})="","",INDEX(FM_ID,{n}))'),
        C("Short", "ファンド", "calc", 10, f='=IF({ID}="","",INDEX(FM_Short,{n}))'),
        C("Strategy", "Strategy", "calc", 11, f='=IF({ID}="","",INDEX(FM_Strategy,{n}))'),
        C("Ccy", "通貨", "calc", 5, f='=IF({ID}="","",INDEX(FM_Ccy,{n}))'),
        C("NavDate", "最新NAV日付", "calc", 11, FMT_DATE, f='=IF({ID}="","",IF(N(INDEX(RC_OffDate,{n}))=0,"",INDEX(RC_OffDate,{n})))'),
        C("NAV", "Latest NAV", "calc", 10, FMT_NUM, f='=IF({NavDate}="","",INDEX(RC_OffNAV,{n}))'),
        C("OnEval", "基準日NAV", "calc", 6, f='=IF({ID}="","",IF({NavDate}="","N",IF({NavDate}>=EvalDate-S_NavDateTol,"Y","N")))'),
        C("TRI", "TRI", "aux", 9, FMT_4, f='=IF({NavDate}="","",' + tri("{NavDate}") + ")"),
        C("WBase", "前週基準NAV日", "aux", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(NV_Date,NV_Fund,{ID},NV_Source,"OFFICIAL",NV_Date,"<="&PrevEvalDate))'),
        C("WeeklyRet", "Weekly Return", "calc", 9, FMT_PCT, f='=IF({ID}="","",IF(AND({OnEval}="Y",N({WBase})>=PrevEvalDate-S_NavDateTol),{TRI}/' + tri("{WBase}") + '-1,""))'),
        C("MBase", "前月末基準NAV日", "aux", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(NV_Date,NV_Fund,{ID},NV_Source,"OFFICIAL",NV_Date,"<="&MonthBaseDate))'),
        C("MonthlyRet", "Monthly Return", "calc", 9, FMT_PCT, f='=IF({ID}="","",IF(AND({OnEval}="Y",N({MBase})>0),{TRI}/' + tri("{MBase}") + '-1,""))'),
        C("YBase", "前年末基準NAV日", "aux", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(NV_Date,NV_Fund,{ID},NV_Source,"OFFICIAL",NV_Date,"<="&YearBaseDate))'),
        C("YTD", "YTD Return", "calc", 9, FMT_PCT, f='=IF({ID}="","",IF(AND({OnEval}="Y",N({YBase})>0),{TRI}/' + tri("{YBase}") + '-1,""))'),
        C("FirstDate", "履歴開始日", "aux", 10, FMT_DATE, f='=IF({ID}="","",MINIFS(NV_Date,NV_Fund,{ID},NV_Source,"OFFICIAL"))'),
        C("SI", "Since Inception", "calc", 9, FMT_PCT, f='=IF(OR({ID}="",{TRI}=""),"",{TRI}/100-1)'),
        C("SINote", "SI基準", "calc", 9, f='=IF({ID}="","",IF({FirstDate}=INDEX(FM_Inception,{n}),"設定来","履歴開始来"))'),
        C("IdxKey", "参照指数Key", "aux", 11, f='=IF({ID}="","",IFERROR(INDEX(IM_QisKey,MATCH(INDEX(FM_PrimaryIdx,{n}),IM_ID,0)),""))'),
        C("IdxWoW", "Linked Index Weekly", "calc", 9, FMT_PCT, f='=IF({IdxKey}="","",IFERROR(INDEX(MD_WoW,MATCH({IdxKey},MD_Key,0)),""))'),
        C("IdxMoM", "Linked Index Monthly", "calc", 9, FMT_PCT, f='=IF({IdxKey}="","",IFERROR(INDEX(MD_MoM,MATCH({IdxKey},MD_Key,0)),""))'),
        C("MaxDD", "Maximum Drawdown", "calc", 9, FMT_PCT, f='=IF({ID}="","",MINIFS(NV_DD,NV_Fund,{ID},NV_Source,"OFFICIAL",NV_Date,"<="&EvalDate))'),
        C("CurDD", "現在DD", "calc", 8, FMT_PCT, f='=IF({NavDate}="","",INDEX(NV_DD,MATCH({ID}&"|OFFICIAL|"&{NavDate},NV_Key,0)))'),
        C("NObs", "Vol観測数", "aux", 6, f='=IF({ID}="","",COUNTIFS(NV_Fund,{ID},NV_Win,1))'),
        C("Vol", "Realized Vol(年率)", "calc", 9, FMT_PCT, f='=IF(OR({ID}="",N({NObs})<4),"",SQRT(MAX(0,(SUMIFS(NV_R2,NV_Fund,{ID},NV_Win,1)-SUMIFS(NV_RW,NV_Fund,{ID},NV_Win,1)^2/{NObs})/({NObs}-1))*52))'),
        C("Exposure", "Current Exposure", "calc", 13, "#,##0.00", f=('=IF({ID}="","",IFERROR(IF({Strategy}="OPTION",INDEX(OA_Gross,MATCH(INDEX(FM_PrimaryIdx,{n}),OA_ID,0)),'
                                                                      'IF({Strategy}="VOLTARGET",INDEX(VS_Exp,MATCH(INDEX(FM_PrimaryIdx,{n}),VS_ID,0)),IF({Strategy}="TRENDCARRY",INDEX(TS_Gross,MATCH(INDEX(FM_PrimaryIdx,{n}),TS_ID,0)),""))),""))')),
        C("ExpLabel", "Exposure単位", "calc", 14, f='=IF({ID}="","",IF({Strategy}="OPTION","Gross Notional(USD)",IF({Strategy}="VOLTARGET","Exposure(倍)","Gross |Weight|")))'),
        C("PosChg", "Weekly Position Change", "calc", 50, f=('=IF({ID}="","",IFERROR(IF({Strategy}="OPTION",INDEX(OA_ChgText,MATCH(INDEX(FM_PrimaryIdx,{n}),OA_ID,0)),'
                                                               'IF({Strategy}="VOLTARGET",INDEX(VS_Text,MATCH(INDEX(FM_PrimaryIdx,{n}),VS_ID,0)),INDEX(TS_Text,MATCH(INDEX(FM_PrimaryIdx,{n}),TS_ID,0)))),"データなし"))')),
        C("Abn", "異常値", "aux", 5, f='=IF(ISNUMBER({WeeklyRet}),IF(ABS({WeeklyRet})>=S_AbnReturn,1,0),"")'),
        C("Recon", "NAV照合", "calc", 14, f='=IF({ID}="","",INDEX(RC_Status,{n}))'),
        C("DataStatus", "データ状況", "calc", 14, f='=IF({ID}="","",IF({OnEval}<>"Y","NAV未更新",IF({IdxWoW}="","参照指数欠損","OK")))'),
    ], CAP_FUND)
    pf.write()
    T["PF"] = pf
    ws.freeze_panes = "C6"

    r0 = 40
    section(ws, r0 - 1, "B. 寄与度（Confirmed＝運用会社の正式データ／Estimated＝モデル推定／Observed＝市場との関連が観測されるのみ）")
    a_cols = [
        C("ID", "Fund ID", "calc", 7, f='=IF(INDEX(FM_ID,{n})="","",INDEX(FM_ID,{n}))'),
        C("Ret", "Weekly Return", "calc", 9, FMT_PCT, f='=IF({ID}="","",INDEX(PF_WeeklyRet,{n}))'),
        C("HasConf", "正式データ", "calc", 7, f='=IF({ID}="","",IF(COUNTIFS(AT_Fund,{ID},AT_Date,EvalDate,AT_Level,"CONFIRMED")>0,"あり","なし"))'),
    ]
    comps = ["Strategy Contribution", "Underlying Market", "FX", "Carry/Income", "Option Premium", "Fees", "Hedging Costs"]
    for i, comp in enumerate(comps):
        a_cols.append(C(f"C{i}", comp, "calc", 10, FMT_PCT, f=f'=IF({{ID}}="","",IF(COUNTIFS(AT_Fund,{{ID}},AT_Date,EvalDate,AT_Component,"{comp}")=0,"",SUMIFS(AT_Value,AT_Fund,{{ID}},AT_Date,EvalDate,AT_Component,"{comp}")))'))
    a_cols += [
        C("Residual", "Residual", "calc", 9, FMT_PCT, f='=IF(OR({HasConf}<>"あり",NOT(ISNUMBER({Ret}))),"",{Ret}-SUM(' + "{C0}:{C6}" + '))'),
        C("EstFX", "FX(Estimated)", "calc", 9, FMT_PCT, f='=IF({ID}="","",IF(AND(INDEX(FM_Ccy,{n})="JPY",INDEX(FM_Hedged,{n})="N"),IFERROR(INDEX(MD_WoW,MATCH(INDEX(FM_FxKey,{n}),MD_Key,0)),""),""))'),
        C("EstUnd", "原資産(Estimated)", "calc", 10, FMT_PCT, f='=IF({ID}="","",IF(INDEX(FM_Strategy,{n})="VOLTARGET",IFERROR(INDEX(VS_EstContrib,MATCH(INDEX(FM_PrimaryIdx,{n}),VS_ID,0)),""),""))'),
        C("Level", "寄与度レベル", "calc", 11, f='=IF({ID}="","",IF({HasConf}="あり","Confirmed",IF(OR(ISNUMBER({EstFX}),ISNUMBER({EstUnd})),"Estimated","Observed")))'),
        C("Text", "寄与度コメント", "calc", 70, f=('=IF({ID}="","",IF({Level}="Confirmed","【確認済】運用会社の寄与度データあり（Residual "&TEXT({Residual},"+0.00%;-0.00%")&"）。",'
                                                     'IF({Level}="Estimated","【推定】"&IF(ISNUMBER({EstUnd}),"前週Exposure×原指数騰落率による原資産寄与の推定 "&TEXT({EstUnd},"+0.00%;-0.00%")&"。","")&IF(ISNUMBER({EstFX}),"為替（ヘッジなし）の推定寄与 "&TEXT({EstFX},"+0.00%;-0.00%")&"。",""),'
                                                     '"【観測】正式な寄与度データなし。参照指数 "&IF(ISNUMBER(INDEX(PF_IdxWoW,{n})),TEXT(INDEX(PF_IdxWoW,{n}),"+0.00%;-0.00%"),"（欠損）")&" の変動が観測されるのみで、要因は断定しない。")))'), wrap=True),
    ]
    at = Table(wb, ws, "AS", r0, a_cols, CAP_FUND)
    at.write()
    r_in = r0 + CAP_FUND + 5
    section(ws, r_in - 1, "C. 寄与度データ入力（運用会社提供。Level=CONFIRMED / ESTIMATED）")
    ain = Table(wb, ws, "AT", r_in, [C("Fund", "Fund ID", "in", 7), C("Date", "評価日", "in", 11, FMT_DATE),
                                     C("Component", "Component", "in", 20), C("Value", "寄与", "in", 9, FMT_PCT),
                                     C("Level", "Level", "in", 11), C("Source", "出所", "in", 32)], CAP_ATTR)
    ain.write([dict(Fund=a, Date=b, Component=c, Value=d, Level=e, Source=f) for a, b, c, d, e, f in SD.ATTRIBUTION])
    dv_list(ws, f"{ain.col('Component')}{ain.first}:{ain.col('Component')}{ain.last}", '"' + ",".join(comps) + '"')


# =====================================================================================
# 10 Option analytics
# =====================================================================================
def build_option(wb, ws):
    title(ws, "10_Option_Analytics　オプション系分析（Index Name 単位）",
          "A: 指数サマリー　B: 前回比較（New/Closed/Expired/Increased/Reduced/Possible Roll）　C: 現在の有効ポジション明細。Greeks の合計は単位確認前は参考値。", 40)
    S = lambda use, extra="": f'PH_IndexID,{{ID}},PH_Use,"{use}"{extra}'
    oa = Table(wb, ws, "OA", 5, [
        C("n", "#", "aux", 4),
        C("ID", "Index ID", "calc", 12, f='=IFERROR(INDEX(IM_ID,MATCH({n},IM_OptSeq,0)),"")'),
        C("Name", "Index Name", "calc", 28, f='=IF({ID}="","",INDEX(IM_Name,MATCH({ID},IM_ID,0)))'),
        C("Und", "原資産", "calc", 11, f='=IF({ID}="","",INDEX(IM_Underlying,MATCH({ID},IM_ID,0)))'),
        C("CurDate", "今回基準日", "calc", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(PH_PositionDate,PH_IndexID,{ID},PH_PositionDate,"<="&EvalDate))'),
        C("PrevDate", "前回基準日", "calc", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(PH_PositionDate,PH_IndexID,{ID},PH_PositionDate,"<"&{CurDate}))'),
        C("SpotCur", "Spot", "calc", 10, FMT_NUM, f='=IF(N({CurDate})=0,"",AVERAGEIFS(PH_Spot,PH_IndexID,{ID},PH_PositionDate,{CurDate}))'),
        C("SpotPrev", "Spot(前回)", "calc", 10, FMT_NUM, f='=IF(N({PrevDate})=0,"",AVERAGEIFS(PH_Spot,PH_IndexID,{ID},PH_PositionDate,{PrevDate}))'),
        C("SpotChg", "Spot変化", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({SpotCur}),ISNUMBER({SpotPrev})),{SpotCur}/{SpotPrev}-1,"")'),
        C("LvlCur", "Index Level", "calc", 10, FMT_NUM, f='=IF(N({CurDate})=0,"",AVERAGEIFS(PH_IndexLevel,PH_IndexID,{ID},PH_PositionDate,{CurDate}))'),
        C("LvlPrev", "Level(前回)", "calc", 10, FMT_NUM, f='=IF(N({PrevDate})=0,"",AVERAGEIFS(PH_IndexLevel,PH_IndexID,{ID},PH_PositionDate,{PrevDate}))'),
        C("LvlChg", "Level変化", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({LvlCur}),ISNUMBER({LvlPrev})),{LvlCur}/{LvlPrev}-1,"")'),
        C("NPos", "ポジション数", "calc", 6, f='=IF({ID}="","",COUNTIFS(' + S("CUR") + "))"),
        C("NPosPrev", "前回数", "calc", 6, f='=IF({ID}="","",COUNTIFS(' + S("PRV") + "))"),
        C("Gross", "Gross Notional", "calc", 13, FMT_INT, f='=IF({ID}="","",SUMIFS(PH_AbsN,' + S("CUR") + "))"),
        C("Net", "Net Notional", "calc", 13, FMT_INT, f='=IF({ID}="","",SUMIFS(PH_Notional,' + S("CUR") + "))"),
        C("AbsW", "Σ|Weight|", "calc", 8, FMT_PCT1, f='=IF({ID}="","",SUMIFS(PH_AbsW,' + S("CUR") + "))"),
        C("MaxShare", "最大構成比", "calc", 8, FMT_PCT1, f='=IF(N({AbsW})=0,"",MAXIFS(PH_AbsW,' + S("CUR") + ")/{AbsW})"),
        C("MaxSharePrev", "最大構成比(前回)", "calc", 8, FMT_PCT1, f='=IF({ID}="","",IFERROR(MAXIFS(PH_AbsW,' + S("PRV") + ")/SUMIFS(PH_AbsW," + S("PRV") + '),""))'),
        C("ConcChg", "構成比変化", "calc", 8, FMT_PCT1, f='=IF(AND(ISNUMBER({MaxShare}),ISNUMBER({MaxSharePrev})),{MaxShare}-{MaxSharePrev},"")'),
        C("Nearest", "Nearest Strike Dist", "calc", 9, FMT_PCT, f='=IF(N({NPos})=0,"",MINIFS(PH_AbsDist,' + S("CUR") + "))"),
        C("NearestPrev", "前回Nearest", "calc", 9, FMT_PCT, f='=IF(N({NPosPrev})=0,"",MINIFS(PH_AbsDist,' + S("PRV") + "))"),
        C("NearTicker", "最接近Ticker", "calc", 26, f='=IF({Nearest}="","",IFERROR(INDEX(PH_Ticker,MATCH({ID}&"|N",PH_NearKey,0)),""))'),
        C("NearSigned", "最接近 Dist(符号付)", "calc", 9, FMT_PCT, f='=IF({Nearest}="","",IFERROR(INDEX(PH_Dist,MATCH({ID}&"|N",PH_NearKey,0)),""))'),
        C("NearMny", "最接近 Moneyness", "calc", 8, f='=IF({Nearest}="","",IFERROR(INDEX(PH_Mny,MATCH({ID}&"|N",PH_NearKey,0)),""))'),
        C("WtdDist", "加重平均 Dist", "calc", 9, FMT_PCT, f='=IF(N({Gross})=0,"",SUMIFS(PH_DistxN,' + S("CUR") + ")/{Gross})"),
        C("NearMat", "Nearest Maturity", "calc", 10, FMT_DATE, f='=IF(N({NPos})=0,"",MINIFS(PH_Maturity,' + S("CUR") + "))"),
        C("MatTicker", "最短満期Ticker", "calc", 26, f='=IF({NearMat}="","",IFERROR(INDEX(PH_Ticker,MATCH({ID}&"|M",PH_MatKey,0)),""))'),
        C("MatBdays", "満期まで営業日", "calc", 7, f='=IF({NearMat}="","",NETWORKDAYS(EvalDate,{NearMat},Holidays)-1)'),
        C("MatDays", "満期まで日数", "calc", 7, f='=IF({NearMat}="","",{NearMat}-EvalDate)'),
        C("ITM", "ITM数", "calc", 5, f='=IF({ID}="","",COUNTIFS(' + S("CUR", ',PH_Mny,"ITM"') + "))"),
        C("ITMPrev", "ITM数(前回)", "calc", 5, f='=IF({ID}="","",COUNTIFS(' + S("PRV", ',PH_Mny,"ITM"') + "))"),
        C("NearCnt", "Strike 5%以内", "calc", 6, f='=IF({ID}="","",COUNTIFS(' + S("CUR") + ',PH_AbsDist,"<="&S_NearStrike1))'),
        C("IV", "加重平均IV", "calc", 7, FMT_PCT1, f='=IF(N({Gross})=0,"",SUMIFS(PH_IVxN,' + S("CUR") + ")/{Gross})"),
        C("IVPrev", "IV(前回)", "calc", 7, FMT_PCT1, f='=IF({ID}="","",IFERROR(SUMIFS(PH_IVxN,' + S("PRV") + ")/SUMIFS(PH_AbsN," + S("PRV") + '),""))'),
        C("IVChg", "IV変化", "calc", 7, "+0.0%;-0.0%", f='=IF(AND(ISNUMBER({IV}),ISNUMBER({IVPrev})),{IV}-{IVPrev},"")'),
        C("DeltaN", "Σ Delta×N（参考）", "calc", 13, FMT_INT, f='=IF({ID}="","",SUMIFS(PH_DeltaN,' + S("CUR") + "))"),
        C("GammaN", "Σ Gamma×N（参考）", "calc", 10, FMT_NUM, f='=IF({ID}="","",SUMIFS(PH_GammaN,' + S("CUR") + "))"),
        C("GammaNPrev", "前回Γ（参考）", "calc", 10, FMT_NUM, f='=IF({ID}="","",SUMIFS(PH_GammaN,' + S("PRV") + "))"),
        C("GammaChg", "Γ変化率", "calc", 8, FMT_PCT1, f='=IF(AND(ISNUMBER({GammaN}),N({GammaNPrev})<>0),({GammaN}-{GammaNPrev})/ABS({GammaNPrev}),"")'),
        C("VegaN", "Σ Vega×N（参考）", "calc", 13, FMT_INT, f='=IF({ID}="","",SUMIFS(PH_VegaN,' + S("CUR") + "))"),
        C("ThetaN", "Σ Theta×N（参考）", "calc", 13, FMT_INT, f='=IF({ID}="","",SUMIFS(PH_ThetaN,' + S("CUR") + "))"),
        C("GreeksStatus", "Greeks区分", "calc", 14, f='=IF({ID}="","",IF(INDEX(IM_GreeksOK,MATCH({ID},IM_ID,0))="Y","単位確認済","参考値（単位未確認）"))'),
        C("New", "New", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Base,"New"))'),
        C("Closed", "Closed", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Base,"Closed"))'),
        C("Expired", "Expired", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Base,"Expired"))'),
        C("Inc", "Increased", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Base,"Increased"))'),
        C("Red", "Reduced", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Base,"Reduced"))'),
        C("RollNew", "Possible Roll", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_Roll,"Possible Roll (新規側)"))'),
        C("ItmTrans", "→ITM", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_ITM,"→ITM"))'),
        C("StrikeChg", "Strike Change", "calc", 5, f='=IF({ID}="","",COUNTIFS(CH_Index,{ID},CH_StrikeChg,"Strike Change"))'),
        C("Stale", "基準日遅れ(営業日)", "calc", 7, f='=IF(N({CurDate})=0,"",NETWORKDAYS({CurDate},EvalDate,Holidays)-1)'),
        C("LvlCheck", "Level vs BBG", "calc", 8, "0.00%", f='=IF(OR({LvlCur}="",{CurDate}<>EvalDate),"",IFERROR({LvlCur}/INDEX(MD_Value,MATCH(INDEX(IM_QisKey,MATCH({ID},IM_ID,0)),MD_Key,0))-1,""))'),
        C("ChgText", "ポジション変化（文章）", "calc", 60, wrap=True, f=('=IF({ID}="","",IF(N({CurDate})=0,"ポジションデータなし（要追加データ）","ポジション基準日"&TEXT({CurDate},"mm/dd")&"：有効"&{NPos}&"行（Gross "&TEXT({Gross}/1000000,"#,##0")&"百万USD）。前回("&TEXT({PrevDate},"mm/dd")&")比 新規"&{New}&"・解消"&{Closed}&"・満期"&{Expired}&"・増額"&{Inc}&"・減額"&{Red}&IF({RollNew}>0,"、うちロール候補"&{RollNew}&"（推定）","")&"。"))')),
        C("ItmDesc", "ITM化 銘柄", "calc", 30, arr=True, f='=IF({ID}="","",TEXTJOIN("、",TRUE,IF((CH_Index={ID})*(CH_ITM="→ITM"),CH_Ticker,"")))'),
        C("RollDesc", "ロール候補 銘柄", "calc", 30, arr=True, f='=IF({ID}="","",TEXTJOIN("、",TRUE,IF((CH_Index={ID})*(CH_Roll<>""),CH_Ticker&"("&CH_Base&")","")))'),
        C("NearDesc", "最接近（文章）", "calc", 40, f='=IF({NearTicker}="","",{NearTicker}&"（"&{NearMny}&"、Strike距離 "&TEXT({NearSigned},"+0.00%;-0.00%")&"）")'),
        C("MatDesc", "最短満期（文章）", "calc", 40, f='=IF({MatTicker}="","",{MatTicker}&"（満期 "&TEXT({NearMat},"yyyy/mm/dd")&"、残り"&{MatBdays}&"営業日）")'),
    ], CAP_IDX)
    oa.write([dict(n=i + 1) for i in range(CAP_IDX)])
    T["OA"] = oa

    r1 = 40
    section(ws, r1 - 1, "B. ポジション比較（前回 vs 今回。集約ポジション＝Index ID｜Ticker 単位。ロールは推定であり確定ではない）")
    k = lambda fld, fl: f'IFERROR(INDEX(PH_{fld},MATCH({{Key}}&"|{fl}",PH_KeyFlag,0)),"")'
    ch = Table(wb, ws, "CH", r1, [
        C("n", "#", "aux", 4),
        C("Key", "Position Key", "calc", 36, f='=IFERROR(INDEX(PH_AggKey,MATCH({n},PH_ChgSeq,0)),"")'),
        C("Index", "Index ID", "calc", 12, f='=IF({Key}="","",INDEX(PH_IndexID,MATCH({Key},PH_AggKey,0)))'),
        C("Und", "原資産", "calc", 6, f='=IF({Key}="","",INDEX(PH_Underlying,MATCH({Key},PH_AggKey,0)))'),
        C("Type", "Type", "calc", 6, f='=IF({Key}="","",INDEX(PH_OptType,MATCH({Key},PH_AggKey,0)))'),
        C("Ticker", "Ticker", "calc", 26, f='=IF({Key}="","",INDEX(PH_Ticker,MATCH({Key},PH_AggKey,0)))'),
        C("Group", "Group", "aux", 18, f='=IF({Key}="","",{Index}&"|"&{Und}&"|"&{Type})'),
        C("StrikePrev", "Strike(前回)", "calc", 8, FMT_NUM, f="=" + k("Strike", "PRV")),
        C("Strike", "Strike", "calc", 8, FMT_NUM, f='=IF({Key}="","",IFERROR(INDEX(PH_Strike,MATCH({Key}&"|CUR",PH_KeyFlag,0)),{StrikePrev}))'),
        C("Mat", "Maturity", "calc", 10, FMT_DATE, f='=IF({Key}="","",INDEX(PH_Maturity,MATCH({Key},PH_AggKey,0)))'),
        C("NPrev", "Notional(前回)", "calc", 12, FMT_INT, f='=IF({Key}="","",SUMIFS(PH_Notional,PH_AggKey,{Key},PH_Use,"PRV"))'),
        C("NCur", "Notional(今回)", "calc", 12, FMT_INT, f='=IF({Key}="","",SUMIFS(PH_Notional,PH_AggKey,{Key},PH_Use,"CUR"))'),
        C("RowsPrev", "行数(前回)", "aux", 5, f='=IF({Key}="","",COUNTIFS(PH_AggKey,{Key},PH_Use,"PRV"))'),
        C("RowsCur", "行数(今回)", "aux", 5, f='=IF({Key}="","",COUNTIFS(PH_AggKey,{Key},PH_Use,"CUR"))'),
        C("Chg", "変化", "calc", 12, FMT_INT, f='=IF({Key}="","",{NCur}-{NPrev})'),
        C("ChgPct", "変化率(|N|)", "calc", 8, FMT_PCT1, f='=IF(OR({Key}="",N({NPrev})=0),"",(ABS({NCur})-ABS({NPrev}))/ABS({NPrev}))'),
        C("MnyPrev", "Moneyness(前回)", "calc", 8, f="=" + k("Mny", "PRV")),
        C("MnyCur", "Moneyness(今回)", "calc", 8, f="=" + k("Mny", "CUR")),
        C("CurPosDate", "今回基準日", "aux", 10, FMT_DATE, f='=IF({Key}="","",INDEX(OA_CurDate,MATCH({Index},OA_ID,0)))'),
        C("Base", "基本分類", "calc", 10, f='=IF({Key}="","",IF({RowsPrev}=0,"New",IF({RowsCur}=0,IF({Mat}<={CurPosDate},"Expired","Closed"),IF(ABS(N({ChgPct}))<=S_NotionalTol,"Unchanged",IF({ChgPct}>0,"Increased","Reduced")))))'),
        C("StrikeChg", "Strike Change", "calc", 9, f='=IF({Key}="","",IF(AND({RowsPrev}>0,{RowsCur}>0,{StrikePrev}<>"",{Strike}<>{StrikePrev}),"Strike Change",""))'),
        C("Roll", "ロール判定（推定）", "calc", 18, f=('=IF({Key}="","",IF({Base}="New",IF(COUNTIFS(CH_Group,{Group},CH_Base,"Closed",CH_Mat,"<"&{Mat})+COUNTIFS(CH_Group,{Group},CH_Base,"Expired",CH_Mat,"<"&{Mat})>0,"Possible Roll (新規側)",""),'
                                                       'IF(OR({Base}="Closed",{Base}="Expired"),IF(COUNTIFS(CH_Group,{Group},CH_Base,"New",CH_Mat,">"&{Mat})>0,"Possible Roll (旧側)",""),"")))')),
        C("Final", "分類", "calc", 12, f='=IF({Key}="","",IF({Roll}<>"","Possible Roll",IF({StrikeChg}<>"","Strike Change",{Base})))'),
        C("ITM", "ITM変化", "calc", 7, f='=IF(OR({MnyPrev}="",{MnyCur}=""),"",IF(AND({MnyCur}="ITM",{MnyPrev}<>"ITM"),"→ITM",IF(AND({MnyPrev}="ITM",{MnyCur}<>"ITM"),"ITM→","")))'),
        C("Note", "要確認事項", "calc", 36, f='=IF({Key}="","",TRIM(IF({Type}="UNKNOWN","Put/Call判定不能（要確認） ","")&IF({Roll}<>"","ロールは推定（運用会社に確認） ","")&IF({StrikeChg}<>"","同一TickerでStrike値が変化（要確認）","")))'),
    ], CAP_CH)
    ch.write([dict(n=i + 1) for i in range(CAP_CH)])
    T["CH"] = ch
    text_cf(ws, f"{ch.col('ITM')}{ch.first}:{ch.col('ITM')}{ch.last}", ["→ITM"])

    r2 = r1 + CAP_CH + 5
    section(ws, r2 - 1, "C. 現在の有効ポジション明細（今回基準日・満期未到来。個別取引単位）")
    I = lambda fld: f'=IF({{Row}}="","",INDEX(PH_{fld},{{Row}}))'
    cp = Table(wb, ws, "CP", r2, [
        C("n", "#", "aux", 4), C("Row", "履歴行", "aux", 5, f='=IFERROR(MATCH({n},PH_CurSeq,0),"")'),
        C("Index", "Index ID", "calc", 12, f=I("IndexID")), C("Ticker", "Ticker", "calc", 26, f=I("Ticker")),
        C("Type", "Type", "calc", 6, f=I("OptType")), C("Strike", "Strike", "calc", 8, FMT_NUM, f=I("Strike")),
        C("Mat", "Maturity", "calc", 10, FMT_DATE, f=I("Maturity")), C("DTE", "DTE", "calc", 5, f=I("DTE")),
        C("BizDTE", "営業日", "calc", 5, f=I("BizDTE")), C("Notional", "Notional", "calc", 12, FMT_INT, f=I("Notional")),
        C("Weight", "Weight", "calc", 7, FMT_PCT1, f=I("Weight")), C("Spot", "Spot", "calc", 9, FMT_NUM, f=I("Spot")),
        C("Dist", "Strike Dist", "calc", 8, FMT_PCT, f=I("Dist")), C("Mny", "Moneyness", "calc", 8, f=I("Mny")),
        C("Delta", "Delta", "calc", 7, FMT_4, f=I("Delta")), C("Gamma", "Gamma", "calc", 9, FMT_6, f=I("Gamma")),
        C("Vega", "Vega", "calc", 7, FMT_NUM, f=I("Vega")), C("Theta", "Theta", "calc", 7, FMT_NUM, f=I("Theta")),
        C("IV", "IV", "calc", 7, FMT_PCT1, f=I("IV")), C("Mid", "Mid", "calc", 8, FMT_NUM, f=I("Mid")),
        C("Open", "open_date", "calc", 10, FMT_DATE, f=I("OpenDate")), C("TradeKey", "Trade Key", "calc", 40, f=I("TradeKey")),
        C("Multi", "同一銘柄取引数", "calc", 6, f=I("Multi")),
    ], CAP_CUR)
    cp.write([dict(n=i + 1) for i in range(CAP_CUR)])
    text_cf(ws, f"{cp.col('Mny')}{cp.first}:{cp.col('Mny')}{cp.last}", ["ITM"], strong=False)
    ws.freeze_panes = "C6"


# =====================================================================================
# 11 Vol Target / 12 Trend Carry
# =====================================================================================
def build_vt(wb, ws):
    title(ws, "11_VolTarget_Analytics　Vol Target 分析",
          "B 表（運用会社の週次データ）を入力。ボラティリティ上昇時の Exposure 縮小を検出。※ポジションファイルには Vol Target 項目がないため追加データが必要。", 30)
    r_in = 40
    key = lambda d: f'{{ID}}&"|"&{d}'
    val = lambda fld, d: f'=IF(N({{{d}}})=0,"",IF(COUNTIF(VI_Key,{key("{" + d + "}")})=0,"",SUMIFS(VI_{fld},VI_Key,{key("{" + d + "}")})))'
    vs = Table(wb, ws, "VS", 5, [
        C("n", "#", "aux", 4),
        C("ID", "Index ID", "calc", 11, f='=IFERROR(INDEX(IM_ID,MATCH({n},IM_VtSeq,0)),"")'),
        C("Name", "Index Name", "calc", 28, f='=IF({ID}="","",INDEX(IM_Name,MATCH({ID},IM_ID,0)))'),
        C("CurDate", "今回", "calc", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(VI_Date,VI_Index,{ID},VI_Date,"<="&EvalDate))'),
        C("PrevDate", "前回", "calc", 10, FMT_DATE, f='=IF(N({CurDate})=0,"",MAXIFS(VI_Date,VI_Index,{ID},VI_Date,"<"&{CurDate}))'),
        C("Target", "Target Vol", "calc", 7, FMT_PCT1, f=val("Target", "CurDate")),
        C("RV", "Realized Vol", "calc", 7, FMT_PCT1, f=val("RV", "CurDate")),
        C("RVPrev", "RV(前回)", "calc", 7, FMT_PCT1, f=val("RV", "PrevDate")),
        C("RVChg", "RV変化", "calc", 7, "+0.0%;-0.0%", f='=IF(AND(ISNUMBER({RV}),ISNUMBER({RVPrev})),{RV}-{RVPrev},"")'),
        C("Exp", "Exposure", "calc", 8, FMT_PCT1, f=val("Exp", "CurDate")),
        C("ExpPrev", "Exposure(前回)", "calc", 8, FMT_PCT1, f=val("Exp", "PrevDate")),
        C("ExpChg", "Exposure変化率", "calc", 8, FMT_PCT1, f='=IF(AND(ISNUMBER({Exp}),N({ExpPrev})<>0),{Exp}/{ExpPrev}-1,"")'),
        C("Min", "Min Exp", "calc", 7, FMT_PCT1, f=val("Min", "CurDate")),
        C("Max", "Max Exp", "calc", 7, FMT_PCT1, f=val("Max", "CurDate")),
        C("Ratio", "Exp/Max", "calc", 7, FMT_PCT1, f='=IF(AND(ISNUMBER({Exp}),N({Max})>0),{Exp}/{Max},"")'),
        C("Lev", "Leverage", "calc", 7, "0.00", f=val("Lev", "CurDate")),
        C("Rebal", "Rebalance回数", "calc", 6, f='=IF(N({PrevDate})=0,"",COUNTIFS(VI_Index,{ID},VI_Date,">"&{PrevDate},VI_Date,"<="&{CurDate},VI_Rebal,"Y"))'),
        C("UndCur", "原指数", "calc", 9, FMT_NUM, f=val("Und", "CurDate")),
        C("UndPrev", "原指数(前回)", "calc", 9, FMT_NUM, f=val("Und", "PrevDate")),
        C("UndWoW", "原指数騰落", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({UndCur}),N({UndPrev})>0),{UndCur}/{UndPrev}-1,"")'),
        C("StratCur", "Strategy", "calc", 9, FMT_NUM, f=val("Strat", "CurDate")),
        C("StratPrev", "Strategy(前回)", "calc", 9, FMT_NUM, f=val("Strat", "PrevDate")),
        C("StratWoW", "Strategy騰落", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({StratCur}),N({StratPrev})>0),{StratCur}/{StratPrev}-1,"")'),
        C("Track", "乖離(S-U)", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({StratWoW}),ISNUMBER({UndWoW})),{StratWoW}-{UndWoW},"")'),
        C("EstContrib", "原資産寄与(推定)", "calc", 8, FMT_PCT, f='=IF(AND(ISNUMBER({ExpPrev}),ISNUMBER({UndWoW})),{ExpPrev}*{UndWoW},"")'),
        C("Stale", "基準日遅れ(営業日)", "calc", 7, f='=IF(N({CurDate})=0,"",NETWORKDAYS({CurDate},EvalDate,Holidays)-1)'),
        C("Text", "コメント", "calc", 70, wrap=True, f=('=IF({ID}="","",IF(N({CurDate})=0,"Vol Targetデータなし（要追加データ）","【事実】"&TEXT({CurDate},"mm/dd")&"時点 Exposure "&TEXT({Exp},"0%")&"（前週 "&TEXT({ExpPrev},"0%")&"、"&TEXT({ExpChg},"+0.0%;-0.0%")&"）、実現Vol "&TEXT({RV},"0.0%")&"（目標 "&TEXT({Target},"0.0%")&"）。"'
                                                           '&IF(AND(ISNUMBER({ExpChg}),ISNUMBER({RVChg})),IF(AND({ExpChg}<0,{RVChg}>0),"【観測】実現Vol上昇に伴いExposureが縮小（Vol Targetの仕組みと整合的）。",""),"")))')),
        C("Desc", "詳細（アラート用）", "calc", 40, f='=IF(N({CurDate})=0,"","RV "&TEXT({RVPrev},"0.0%")&"→"&TEXT({RV},"0.0%")&"、Exposure "&TEXT({ExpPrev},"0%")&"→"&TEXT({Exp},"0%"))'),
    ], CAP_SUM)
    vs.write([dict(n=i + 1) for i in range(CAP_SUM)])
    T["VS"] = vs

    section(ws, r_in - 1, "B. Vol Target 入力（運用会社 週次データ。Requires Additional Data）")
    vi = Table(wb, ws, "VI", r_in, [
        C("Date", "Date", "in", 11, FMT_DATE), C("Index", "Index ID", "in", 11), C("Target", "Target Vol", "in", 8, FMT_PCT1),
        C("RV", "Realized Vol", "in", 8, FMT_PCT1), C("Exp", "Exposure", "in", 8, FMT_PCT1), C("Min", "Min Exp", "in", 8, FMT_PCT1),
        C("Max", "Max Exp", "in", 8, FMT_PCT1), C("Lev", "Leverage", "in", 8, "0.00"), C("Und", "Underlying Level", "in", 10, FMT_NUM),
        C("Strat", "Strategy Level", "in", 10, FMT_NUM), C("Rebal", "Rebalance", "in", 7), C("Source", "出所", "in", 28),
        C("Key", "Key", "aux", 18, f='=IF({Date}="","",{Index}&"|"&{Date})'),
    ], CAP_VT)
    vi.write([dict(zip(["Date", "Index", "Target", "RV", "Exp", "Min", "Max", "Lev", "Und", "Strat", "Rebal", "Source"], v)) for v in SD.VT_INPUT])
    ws.freeze_panes = "C6"


def build_tc(wb, ws):
    title(ws, "12_TrendCarry_Analytics　Trend / Carry 分析",
          "C 表（運用会社の週次ポジション）を入力。シグナル反転・Long/Short 反転・Carry 低下・集中度を検出。※ポジションファイル外の追加データが必要。", 30)
    classes = ["Equity", "Rates", "FX", "Commodity"]
    cols = [
        C("n", "#", "aux", 4),
        C("ID", "Index ID", "calc", 11, f='=IFERROR(INDEX(IM_ID,MATCH({n},IM_TcSeq,0)),"")'),
        C("Name", "Index Name", "calc", 28, f='=IF({ID}="","",INDEX(IM_Name,MATCH({ID},IM_ID,0)))'),
        C("CurDate", "今回", "calc", 10, FMT_DATE, f='=IF({ID}="","",MAXIFS(TI_Date,TI_Index,{ID},TI_Date,"<="&EvalDate))'),
        C("PrevDate", "前回", "calc", 10, FMT_DATE, f='=IF(N({CurDate})=0,"",MAXIFS(TI_Date,TI_Index,{ID},TI_Date,"<"&{CurDate}))'),
        C("Gross", "Gross |W|", "calc", 8, FMT_PCT1, f='=IF(N({CurDate})=0,"",SUMIFS(TI_AbsW,TI_Index,{ID},TI_Date,{CurDate}))'),
        C("GrossPrev", "Gross(前回)", "calc", 8, FMT_PCT1, f='=IF(N({PrevDate})=0,"",SUMIFS(TI_AbsW,TI_Index,{ID},TI_Date,{PrevDate}))'),
        C("Net", "Net W", "calc", 8, FMT_PCT1, f='=IF(N({CurDate})=0,"",SUMIFS(TI_Weight,TI_Index,{ID},TI_Date,{CurDate}))'),
        C("NLong", "Long数", "calc", 5, f='=IF(N({CurDate})=0,"",COUNTIFS(TI_Index,{ID},TI_Date,{CurDate},TI_Weight,">0"))'),
        C("NShort", "Short数", "calc", 5, f='=IF(N({CurDate})=0,"",COUNTIFS(TI_Index,{ID},TI_Date,{CurDate},TI_Weight,"<0"))'),
        C("SigRev", "Signal反転", "calc", 5, f='=IF({ID}="","",COUNTIFS(TX_Index,{ID},TX_SigRev,"Y"))'),
        C("LsRev", "L/S反転", "calc", 5, f='=IF({ID}="","",COUNTIFS(TX_Index,{ID},TX_LsRev,"Y"))'),
        C("Carry", "加重Carry", "calc", 7, "0.00%", f='=IF(N({Gross})=0,"",SUMIFS(TI_CarryW,TI_Index,{ID},TI_Date,{CurDate})/{Gross})'),
        C("CarryPrev", "Carry(前回)", "calc", 7, "0.00%", f='=IF(N({GrossPrev})=0,"",SUMIFS(TI_CarryW,TI_Index,{ID},TI_Date,{PrevDate})/{GrossPrev})'),
        C("CarryChg", "Carry変化", "calc", 7, "+0.00%;-0.00%", f='=IF(AND(ISNUMBER({Carry}),ISNUMBER({CarryPrev})),{Carry}-{CarryPrev},"")'),
        C("Conc", "最大構成比", "calc", 7, FMT_PCT1, f='=IF(N({Gross})=0,"",MAXIFS(TI_AbsW,TI_Index,{ID},TI_Date,{CurDate})/{Gross})'),
    ]
    for c in classes:
        cols.append(C(f"N_{c}", f"Net {c}", "calc", 7, FMT_PCT1, f=f'=IF(N({{CurDate}})=0,"",SUMIFS(TI_Weight,TI_Index,{{ID}},TI_Date,{{CurDate}},TI_Asset,"{c}"))'))
        cols.append(C(f"P_{c}", f"前回 {c}", "calc", 7, FMT_PCT1, f=f'=IF(N({{PrevDate}})=0,"",SUMIFS(TI_Weight,TI_Index,{{ID}},TI_Date,{{PrevDate}},TI_Asset,"{c}"))'))
    cols += [
        C("Contrib", "週次寄与(提供時)", "calc", 8, FMT_PCT, f='=IF(N({CurDate})=0,"",IF(COUNTIFS(TI_Index,{ID},TI_Date,{CurDate},TI_Contrib,"<>")=0,"",SUMIFS(TI_Contrib,TI_Index,{ID},TI_Date,{CurDate})))'),
        C("Stale", "基準日遅れ(営業日)", "calc", 7, f='=IF(N({CurDate})=0,"",NETWORKDAYS({CurDate},EvalDate,Holidays)-1)'),
        C("SigRevDesc", "シグナル反転 銘柄", "calc", 30, arr=True, f='=IF({ID}="","",TEXTJOIN("、",TRUE,IF((TX_Index={ID})*(TX_SigRev="Y"),TX_Asset&":"&TX_Instrument&"("&TEXT(TX_SigPrev,"+0.00;-0.00")&"→"&TEXT(TX_SigCur,"+0.00;-0.00")&")","")))'),
        C("LsRevDesc", "L/S反転 銘柄", "calc", 30, arr=True, f='=IF({ID}="","",TEXTJOIN("、",TRUE,IF((TX_Index={ID})*(TX_LsRev="Y"),TX_Instrument&"("&TX_LSPrev&"→"&TX_LSCur&")","")))'),
        C("Text", "コメント", "calc", 70, wrap=True, f=('=IF({ID}="","",IF(N({CurDate})=0,"Trend/Carryデータなし（要追加データ）","【事実】"&TEXT({CurDate},"mm/dd")&"時点 Gross "&TEXT({Gross},"0%")&"・Net "&TEXT({Net},"+0%;-0%")&"（Long "&{NLong}&"／Short "&{NShort}&"）。シグナル反転"&{SigRev}&"件、L/S反転"&{LsRev}&"件"'
                                                           '&IF({SigRevDesc}<>"","（"&{SigRevDesc}&"）","")&"。加重Carry推定 "&TEXT({Carry},"0.0%")&"（前週比 "&TEXT({CarryChg}*100,"+0.0;-0.0")&"pt、推定値）。"))')),
    ]
    ts = Table(wb, ws, "TS", 5, cols, CAP_SUM)
    ts.write([dict(n=i + 1) for i in range(CAP_SUM)])
    T["TS"] = ts

    r1 = 30
    section(ws, r1 - 1, "B. 銘柄別 前週比較")
    kf = lambda fld, fl: f'=IFERROR(INDEX(TI_{fld},MATCH({{InstKey}}&"|{fl}",TI_KeyFlag,0)),"")'
    tx = Table(wb, ws, "TX", r1, [
        C("n", "#", "aux", 4),
        C("InstKey", "Key", "calc", 26, f='=IFERROR(INDEX(TI_InstKey,MATCH({n},TI_Seq,0)),"")'),
        C("Index", "Index ID", "calc", 10, f='=IF({InstKey}="","",INDEX(TI_Index,MATCH({InstKey},TI_InstKey,0)))'),
        C("Asset", "Asset Class", "calc", 10, f='=IF({InstKey}="","",INDEX(TI_Asset,MATCH({InstKey},TI_InstKey,0)))'),
        C("Instrument", "Instrument", "calc", 16, f='=IF({InstKey}="","",INDEX(TI_Instrument,MATCH({InstKey},TI_InstKey,0)))'),
        C("SigPrev", "Signal(前回)", "calc", 7, "0.00", f=kf("Signal", "PRV")),
        C("SigCur", "Signal(今回)", "calc", 7, "0.00", f=kf("Signal", "CUR")),
        C("SigRev", "Signal反転", "calc", 6, f='=IF(AND(ISNUMBER({SigPrev}),ISNUMBER({SigCur})),IF({SigPrev}*{SigCur}<0,"Y",""),"")'),
        C("WPrev", "Weight(前回)", "calc", 7, FMT_PCT1, f=kf("Weight", "PRV")),
        C("WCur", "Weight(今回)", "calc", 7, FMT_PCT1, f=kf("Weight", "CUR")),
        C("LSPrev", "L/S(前回)", "calc", 6, f='=IF(ISNUMBER({WPrev}),IF({WPrev}>0,"Long",IF({WPrev}<0,"Short","Flat")),"")'),
        C("LSCur", "L/S(今回)", "calc", 6, f='=IF(ISNUMBER({WCur}),IF({WCur}>0,"Long",IF({WCur}<0,"Short","Flat")),"")'),
        C("LsRev", "L/S反転", "calc", 6, f='=IF(AND({LSPrev}<>"",{LSCur}<>"",{LSPrev}<>"Flat",{LSCur}<>"Flat",{LSPrev}<>{LSCur}),"Y","")'),
        C("CarryPrev", "Carry(前回)", "calc", 7, "0.00%", f=kf("Carry", "PRV")),
        C("CarryCur", "Carry(今回)", "calc", 7, "0.00%", f=kf("Carry", "CUR")),
        C("CarryChg", "Carry変化", "calc", 7, "+0.00%;-0.00%", f='=IF(AND(ISNUMBER({CarryPrev}),ISNUMBER({CarryCur})),{CarryCur}-{CarryPrev},"")'),
        C("Contrib", "寄与(今回)", "calc", 7, FMT_PCT, f=kf("Contrib", "CUR")),
    ], CAP_TCX)
    tx.write([dict(n=i + 1) for i in range(CAP_TCX)])
    T["TX"] = tx
    text_cf(ws, f"{tx.col('SigRev')}{tx.first}:{tx.col('SigRev')}{tx.last}", ["Y"], strong=False)
    text_cf(ws, f"{tx.col('LsRev')}{tx.first}:{tx.col('LsRev')}{tx.last}", ["Y"], strong=False)

    r2 = r1 + CAP_TCX + 10
    section(ws, r2 - 1, "C. Trend / Carry 入力（運用会社 週次ポジション。Requires Additional Data）")
    ti = Table(wb, ws, "TI", r2, [
        C("Date", "Date", "in", 11, FMT_DATE), C("Index", "Index ID", "in", 10), C("Asset", "Asset Class", "in", 10),
        C("Instrument", "Instrument", "in", 16), C("Ticker", "BBG Ticker", "in", 14), C("Signal", "Signal", "in", 7, "0.00"),
        C("Weight", "Weight(符号付)", "in", 8, FMT_PCT1), C("Carry", "Carry推定(年率)", "in", 8, "0.00%"),
        C("Roll", "Roll Yield", "in", 7, "0.00%"), C("RateDiff", "金利差", "in", 7, "0.00%"), C("Contrib", "週次寄与", "in", 7, FMT_PCT),
        C("ContribLevel", "寄与Level", "in", 10), C("Source", "出所", "in", 26),
        C("InstKey", "Key", "aux", 22, f='=IF({Date}="","",{Index}&"|"&{Instrument})'),
        C("Flag", "Snapshot", "aux", 6, f='=IF({Date}="","",IFERROR(IF({Date}=INDEX(TS_CurDate,MATCH({Index},TS_ID,0)),"CUR",IF({Date}=INDEX(TS_PrevDate,MATCH({Index},TS_ID,0)),"PRV","")),""))'),
        C("KeyFlag", "Key|Flag", "aux", 24, f='=IF({Date}="","",{InstKey}&"|"&{Flag})'),
        C("Univ", "対象", "aux", 5, f='=IF(OR({Flag}="CUR",{Flag}="PRV"),1,0)'),
        C("First", "初出", "aux", 5, f='=IF({Univ}=1,IF(COUNTIFS({InstKey@1}:{InstKey},{InstKey},{Univ@1}:{Univ},1)=1,1,0),0)'),
        C("Seq", "#", "aux", 5, f='=IF({First}=1,COUNTIF({First@1}:{First},1),"")'),
        C("AbsW", "|W|", "aux", 7, FMT_PCT1, f='=IF({Date}="","",ABS({Weight}))'),
        C("CarryW", "Carry×|W|", "aux", 8, "0.0000", f='=IF({Date}="","",N({Carry})*{AbsW})'),
    ], CAP_TC)
    ti.write([dict(zip(["Date", "Index", "Asset", "Instrument", "Ticker", "Signal", "Weight", "Carry", "Roll", "RateDiff", "Contrib", "ContribLevel", "Source"], r)) for r in SD.tc_input()])
    T["TI"] = ti
    ws.freeze_panes = "C6"


# =====================================================================================
# 13 Risk Alerts
# =====================================================================================
METRICS = [
    # key, formula (tokens: {pf} {oa} {vt} {tc} {Index} {Strategy})
    ("WEEKLY_RET", '=IF({pf}="","",INDEX(PF_WeeklyRet,{pf}))'),
    ("MONTHLY_RET", '=IF({pf}="","",INDEX(PF_MonthlyRet,{pf}))'),
    ("CUR_DD", '=IF({pf}="","",INDEX(PF_CurDD,{pf}))'),
    ("NAV_MISMATCH", '=IF({pf}="","",IF(INDEX(RC_Status,{pf})="MISMATCH",1,0))'),
    ("NAV_WARNING", '=IF({pf}="","",IF(INDEX(RC_Status,{pf})="WARNING",1,0))'),
    ("NAV_UNCOMP", '=IF({pf}="","",IF(OR(INDEX(RC_Status,{pf})="DATE DIFFERENCE",INDEX(RC_Status,{pf})="NOT COMPARABLE",INDEX(RC_Status,{pf})="MISSING DATA"),1,0))'),
    ("NAV_DELAY", '=IF({pf}="","",INDEX(RC_Delay,{pf}))'),
    ("ABN_RET", '=IF({pf}="","",INDEX(PF_WeeklyRet,{pf}))'),
    ("POS_STALE", '=IF({oa}="","",INDEX(OA_Stale,{oa}))'),
    ("STRIKE_DIST", '=IF({oa}="","",INDEX(OA_Nearest,{oa}))'),
    ("ITM_TRANS", '=IF({oa}="","",INDEX(OA_ItmTrans,{oa}))'),
    ("MAT_BDAYS", '=IF({oa}="","",INDEX(OA_MatBdays,{oa}))'),
    ("MAT_DAYS", '=IF({oa}="","",INDEX(OA_MatDays,{oa}))'),
    ("IV_CHG", '=IF({oa}="","",INDEX(OA_IVChg,{oa}))'),
    ("GAMMA_CHG", '=IF({oa}="","",INDEX(OA_GammaChg,{oa}))'),
    ("CONC_CHG", '=IF({oa}="","",INDEX(OA_ConcChg,{oa}))'),
    ("ROLL_CNT", '=IF({oa}="","",INDEX(OA_RollNew,{oa}))'),
    ("EXP_CHG", '=IF({vt}="","",INDEX(VS_ExpChg,{vt}))'),
    ("EXP_RATIO", '=IF({vt}="","",INDEX(VS_Ratio,{vt}))'),
    ("RV_CHG", '=IF({vt}="","",INDEX(VS_RVChg,{vt}))'),
    ("REBAL", '=IF({vt}="","",INDEX(VS_Rebal,{vt}))'),
    ("TRACK", '=IF({vt}="","",INDEX(VS_Track,{vt}))'),
    ("SIG_REV", '=IF({tc}="","",INDEX(TS_SigRev,{tc}))'),
    ("LS_REV", '=IF({tc}="","",INDEX(TS_LsRev,{tc}))'),
    ("CARRY_CHG", '=IF({tc}="","",INDEX(TS_CarryChg,{tc}))'),
    ("TC_CONC", '=IF({tc}="","",INDEX(TS_Conc,{tc}))'),
    ("MKT_MISSING", '=IF({Index}="","",IF({UndSt}<>"OK",1,0)+IF({QisSt}<>"OK",1,0)+IF({VolSt}<>"OK",1,0))'),
    ("NEAREST_DESC", '=IF({oa}="","",INDEX(OA_NearDesc,{oa}))'),
    ("ITM_DESC", '=IF({oa}="","",INDEX(OA_ItmDesc,{oa}))'),
    ("MAT_DESC", '=IF({oa}="","",INDEX(OA_MatDesc,{oa}))'),
    ("ROLL_DESC", '=IF({oa}="","",INDEX(OA_RollDesc,{oa}))'),
    ("VT_DESC", '=IF({vt}="","",INDEX(VS_Desc,{vt}))'),
    ("SIGREV_DESC", '=IF({tc}="","",INDEX(TS_SigRevDesc,{tc}))'),
    ("LSREV_DESC", '=IF({tc}="","",INDEX(TS_LsRevDesc,{tc}))'),
    ("NAV_DESC", '=IF({pf}="","",INDEX(RC_Comment,{pf}))'),
    ("MKT_DESC", '=IF({Index}="","",TRIM(IF({UndSt}<>"OK",{UndKey}&"("&{UndSt}&") ","")&IF({QisSt}<>"OK",{QisKey}&"("&{QisSt}&") ","")&IF({VolSt}<>"OK",{VolKey}&"("&{VolSt}&")","")))'),
    ("OBS_FUND", '=IF({pf}="","",IF(N(INDEX(RC_OffDate,{pf}))=0,"",INDEX(RC_OffDate,{pf})))'),
    ("OBS_INDEX", '=IF({Index}="","",IFERROR(IF({Strategy}="OPTION",INDEX(OA_CurDate,{oa}),IF({Strategy}="VOLTARGET",INDEX(VS_CurDate,{vt}),IF({Strategy}="TRENDCARRY",INDEX(TS_CurDate,{tc}),EvalDate))),EvalDate))'),
]


def build_alerts(wb, ws):
    title(ws, "13_Risk_Alerts　リスクアラート",
          "A: アラート一覧（重要度順）。B: 指標マトリクス（Fund-Index Mapping 1 行ごと）。C: ルール評価グリッド（Mapping × ルール）。ルール・閾値は 19_Settings で変更。確認状況は右側の Review Log に記録。", 22)
    r_mx, r_gr = 215, 255
    # ---- B. matrix
    section(ws, r_mx - 1, "B. 指標マトリクス（Mapping 行 × 指標。アラート判定の入力値）")
    st = lambda k: f'IF({{{k}}}="","OK",IFERROR(INDEX(MD_Status,MATCH({{{k}}},MD_Key,0)),"未設定"))'
    mcols = [
        C("m", "#", "aux", 4),
        C("MapID", "Map ID", "calc", 7, f='=IF(INDEX(MP_Fund,{m})="","",INDEX(MP_ID,{m}))'),
        C("Fund", "Fund", "calc", 7, f='=IF({MapID}="","",INDEX(MP_Fund,{m}))'),
        C("Index", "Index", "calc", 11, f='=IF({MapID}="","",INDEX(MP_Index,{m}))'),
        C("Strategy", "Strategy", "calc", 10, f='=IF({MapID}="","",INDEX(MP_Strategy,{m}))'),
        C("First", "First", "aux", 5, f='=IF({MapID}="","",INDEX(MP_First,{m}))'),
        C("pf", "pf", "aux", 4, f='=IF({Fund}="","",IFERROR(MATCH({Fund},PF_ID,0),""))'),
        C("oa", "oa", "aux", 4, f='=IF({Index}="","",IFERROR(MATCH({Index},OA_ID,0),""))'),
        C("vt", "vt", "aux", 4, f='=IF({Index}="","",IFERROR(MATCH({Index},VS_ID,0),""))'),
        C("tc", "tc", "aux", 4, f='=IF({Index}="","",IFERROR(MATCH({Index},TS_ID,0),""))'),
        C("UndKey", "UndKey", "aux", 8, f='=IF({Index}="","",IFERROR(INDEX(IM_UndKey,MATCH({Index},IM_ID,0))&"",""))'),
        C("QisKey", "QisKey", "aux", 9, f='=IF({Index}="","",IFERROR(INDEX(IM_QisKey,MATCH({Index},IM_ID,0))&"",""))'),
        C("VolKey", "VolKey", "aux", 8, f='=IF({Index}="","",IFERROR(INDEX(IM_VolKey,MATCH({Index},IM_ID,0))&"",""))'),
        C("UndSt", "Und St", "aux", 7, f="=" + st("UndKey")),
        C("QisSt", "Qis St", "aux", 7, f="=" + st("QisKey")),
        C("VolSt", "Vol St", "aux", 7, f="=" + st("VolKey")),
    ]
    first_metric = len(mcols)
    for key, f in METRICS:
        fmt = FMT_PCT if key in ("WEEKLY_RET", "MONTHLY_RET", "CUR_DD", "ABN_RET", "STRIKE_DIST", "IV_CHG", "GAMMA_CHG", "CONC_CHG", "EXP_CHG", "EXP_RATIO", "RV_CHG", "TRACK", "CARRY_CHG", "TC_CONC") else None
        fmt = FMT_DATE if key.startswith("OBS_") else fmt
        mcols.append(C(key, key, "calc", 12 if "DESC" not in key else 26, fmt, f=f))
    mx = Table(wb, ws, "MX", r_mx, mcols, CAP_MAP)
    mx.write([dict(m=i + 1) for i in range(CAP_MAP)])
    L0 = mx.col(METRICS[0][0])
    L1 = mx.col(METRICS[-1][0])
    add_name(wb, "MX_Data", ws.title, f"${L0}${mx.first}:${L1}${mx.last}")
    add_name(wb, "MX_Header", ws.title, f"${L0}${mx.hdr}:${L1}${mx.hdr}")
    T["MX"] = mx

    # ---- C. grid
    section(ws, r_gr - 1, f"C. ルール評価グリッド（Mapping {CAP_MAP} 行 × ルール {CAP_RULE} 行。Scope=FUND のルールは各ファンドの最初の Mapping 行でのみ評価）")
    unit_fmt = lambda v, u: (f'IF({u}="PCT",TEXT({v},"+0.00%;-0.00%"),IF({u}="PCTABS",TEXT({v},"0.00%"),IF({u}="PCTPT",TEXT({v}*100,"+0.0;-0.0")&"pt",'
                             f'IF({u}="RATIO",TEXT({v},"0.0%"),IF({u}="BDAYS",{v}&"営業日",IF({u}="DAYS",{v}&"日",IF({u}="COUNT",{v}&"件","該当")))))))')
    gr = Table(wb, ws, "GR", r_gr, [
        C("g", "#", "aux", 5), C("m", "Map#", "aux", 5, f=f"=INT(({{g}}-1)/{CAP_RULE})+1"), C("k", "Rule#", "aux", 5, f=f"=MOD({{g}}-1,{CAP_RULE})+1"),
        C("RuleID", "Rule", "calc", 6, f='=IF(INDEX(RL_ID,{k})="","",INDEX(RL_ID,{k}))'),
        C("Fund", "Fund", "calc", 7, f='=IF(OR({RuleID}="",INDEX(MX_Fund,{m})=""),"",INDEX(MX_Fund,{m}))'),
        C("Index", "Index", "calc", 11, f='=IF({Fund}="","",INDEX(MX_Index,{m}))'),
        C("Strategy", "Strategy", "calc", 10, f='=IF({Fund}="","",INDEX(MX_Strategy,{m}))'),
        C("Scope", "Scope", "aux", 6, f='=IF({Fund}="","",INDEX(RL_Scope,{k}))'),
        C("App", "適用", "aux", 4, f='=IF({Fund}="",0,IF(AND(INDEX(RL_Enabled,{k})="Y",OR(INDEX(RL_Strategy,{k})="ALL",INDEX(RL_Strategy,{k})={Strategy}),OR({Scope}="INDEX",INDEX(MX_First,{m})=1)),1,0))'),
        C("Val", "Value", "calc", 9, f='=IF({App}=1,INDEX(MX_Data,{m},MATCH(INDEX(RL_Metric,{k}),MX_Header,0)),"")'),
        C("Thr", "Threshold", "calc", 8, f='=IF({App}=1,INDEX(RL_Thr,{k}),"")'),
        C("Op", "Op", "aux", 5, f='=IF({App}=1,INDEX(RL_Op,{k}),"")'),
        C("Raw", "Raw", "aux", 4, f='=IF(AND({App}=1,ISNUMBER({Val})),IF(OR(AND({Op}="<=",{Val}<={Thr}),AND({Op}=">=",{Val}>={Thr}),AND({Op}="<",{Val}<{Thr}),AND({Op}=">",{Val}>{Thr}),AND({Op}="ABS>=",ABS({Val})>={Thr})),1,0),0)'),
        C("Supp", "抑制", "aux", 4, f=f'=IF(OR({{Raw}}=0,INDEX(RL_Supp,{{k}})=""),0,IFERROR(INDEX(GR_Raw,({{m}}-1)*{CAP_RULE}+MATCH(INDEX(RL_Supp,{{k}}),RL_ID,0)),0))'),
        C("Final", "発生", "calc", 4, f='=IF(AND({Raw}=1,{Supp}=0),1,0)'),
        C("Sev", "Severity", "calc", 7, f='=IF({Final}=1,INDEX(RL_Sev,{k}),"")'),
        C("RankKey", "RankKey", "aux", 8, f='=IF({Final}=1,MATCH({Sev},{"Low","Medium","High"},0)*100000-{g},"")'),
        C("FundSev", "Fund|Sev", "aux", 12, f='=IF({Final}=1,{Fund}&"|"&{Sev},"")'),
        C("ObsDate", "Obs Date", "calc", 10, FMT_DATE, f='=IF({Final}=1,IF({Scope}="FUND",INDEX(MX_OBS_FUND,{m}),INDEX(MX_OBS_INDEX,{m})),"")'),
        C("Extra", "Extra", "calc", 20, f='=IF({Final}=1,IF(INDEX(RL_Extra,{k})="","",IFERROR(INDEX(MX_Data,{m},MATCH(INDEX(RL_Extra,{k}),MX_Header,0))&"","")),"")'),
        C("ValTxt", "Value(表示)", "calc", 10, f='=IF({Final}=0,"",' + unit_fmt("{Val}", "INDEX(RL_Unit,{k})") + ")"),
        C("ThrTxt", "Threshold(表示)", "calc", 10, f='=IF({Final}=0,"",IF(INDEX(RL_Unit,{k})="FLAG","—",IF({Op}="ABS>=","±"&TEXT({Thr},"0.00%"),' + unit_fmt("{Thr}", "INDEX(RL_Unit,{k})") + ")))"),
    ], CAP_GRID)
    gr.write([dict(g=i + 1) for i in range(CAP_GRID)])
    T["GR"] = gr

    # ---- A. list
    section(ws, 4, "A. アラート一覧（High → Medium → Low）")
    G_ = lambda fld: f'=IF({{g}}="","",INDEX(GR_{fld},{{g}}))'
    al = Table(wb, ws, "AL", 5, [
        C("n", "#", "aux", 4),
        C("g", "grid", "aux", 5, f='=IFERROR(INDEX(GR_g,MATCH(LARGE(GR_RankKey,{n}),GR_RankKey,0)),"")'),
        C("Sev", "Severity", "calc", 8, f=G_("Sev")),
        C("Fund", "Fund", "calc", 6, f=G_("Fund")),
        C("FundName", "Fund Name", "calc", 10, f='=IF({g}="","",IFERROR(INDEX(FM_Short,MATCH({Fund},FM_ID,0)),""))'),
        C("Investors", "Investor（保有地銀）", "calc", 16, arr=True, f='=IF({g}="","",TEXTJOIN("、",TRUE,IF(HD_Fund={Fund},HD_Bank,"")))'),
        C("Index", "Index", "calc", 11, f='=IF({g}="","",IF(INDEX(GR_Scope,{g})="FUND","",INDEX(GR_Index,{g})))'),
        C("Type", "Alert Type", "calc", 18, f='=IF({g}="","",INDEX(RL_Name,INDEX(GR_k,{g})))'),
        C("Value", "Trigger Value", "calc", 10, f=G_("ValTxt")),
        C("Thr", "Threshold", "calc", 9, f=G_("ThrTxt")),
        C("Obs", "Observation Date", "calc", 10, FMT_DATE, f=G_("ObsDate")),
        C("Expl", "Explanation", "calc", 50, wrap=True, f='=IF({g}="","",SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(INDEX(RL_Expl,INDEX(GR_k,{g})),"<v>",{Value}),"<t>",{Thr}),"<x>",INDEX(GR_Extra,{g})))'),
        C("Action", "Recommended Sales Action", "calc", 30, wrap=True, f='=IF({g}="","",INDEX(RL_Action,INDEX(GR_k,{g})))'),
        C("Source", "Data Source", "calc", 22, f='=IF({g}="","",INDEX(RL_Source,INDEX(GR_k,{g})))'),
        C("Key", "Alert Key", "aux", 30, f='=IF({g}="","",TEXT(EvalDate,"yyyymmdd")&"|"&{Fund}&"|"&{Index}&"|"&INDEX(GR_RuleID,{g}))'),
        C("Review", "確認状況", "calc", 8, f='=IF({g}="","",IFERROR(INDEX(RV_Status,MATCH({Key},RV_Key,0)),"未確認"))'),
        C("Strategy", "Strategy", "aux", 9, f=G_("Strategy")),
        C("k", "Rule#", "aux", 4, f=G_("k")),
        C("DashOK", "Dash", "aux", 4, f='=IF({g}="",0,IF(AND(OR(Dash_Strategy="(全て)",{Strategy}=Dash_Strategy),OR(Dash_Bank="(全て)",COUNTIFS(HD_Fund,{Fund},HD_Bank,Dash_Bank)>0),OR(Dash_Rep="(全て)",COUNTIFS(HD_Fund,{Fund},HD_RepName,Dash_Rep)>0)),1,0))'),
        C("DashSeq", "Dash#", "aux", 4, f='=IF({DashOK}=1,COUNTIF({DashOK@1}:{DashOK},1),"")'),
        C("RepOK", "Rep", "aux", 4, f='=IF({g}="",0,IF(COUNTIFS(HD_Fund,{Fund},HD_RepName,EM_Rep)>0,1,0))'),
        C("RepSeq", "Rep#", "aux", 4, f='=IF({RepOK}=1,COUNTIF({RepOK@1}:{RepOK},1),"")'),
        C("FundRank", "Fund内順位", "aux", 4, f='=IF({g}="","",COUNTIF({Fund@1}:{Fund},{Fund}))'),
        C("FundQ", "顧客向けQ", "aux", 9, f='=IF({g}="","",{Fund}&"|"&IF(LEFT(INDEX(RL_Q,{k}),4)="（社内）",0,1))'),
    ], CAP_ALERT)
    al.write([dict(n=i + 1) for i in range(CAP_ALERT)])
    severity_cf(ws, f"{al.col('Sev')}{al.first}:{al.col('Sev')}{al.last}")
    T["AL"] = al

    c_rv = 25
    section(ws, 4, "D. Review Log（Alert Key ごとの確認状況。監査用に保存）", col=c_rv)
    rv = Table(wb, ws, "RV", 5, [C("Key", "Alert Key", "in", 30), C("Status", "確認状況", "in", 9), C("Comment", "コメント", "in", 30),
                                 C("Reviewer", "確認者", "in", 9), C("Date", "確認日", "in", 10, FMT_DATE)], CAP_ALERT, first_col=c_rv)
    rv.write([dict(Key="20261030|F002||R04", Status="対応中", Comment="（サンプル）BBG側の更新タイミング差の可能性。運用会社に照会中", Reviewer="データ管理", Date=SD.RUN_DATE)])
    dv_list(ws, f"{rv.col('Status')}{rv.first}:{rv.col('Status')}{rv.last}", "=L_Review")
    ws.freeze_panes = "D6"


# =====================================================================================
# 14 Sales priority
# =====================================================================================
def build_priority(wb, ws):
    title(ws, "14_Sales_Priority　営業フォローアップ優先順位",
          "ルールベースのスコア（内訳は Reason 列）。手動で変更する場合は Override 列に High/Medium/Low を入力（理由は備考に記録）。下段は優先度順の一覧。", 24)
    H = lambda fld: f'=IF(INDEX(HD_ID,{{n}})="","",INDEX(HD_{fld},{{n}}))'
    sp = Table(wb, ws, "SP", 5, [
        C("ID", "Holding ID", "calc", 8, f=H("ID")), C("Bank", "Bank", "calc", 9, f=H("Bank")), C("Rep", "営業担当", "calc", 10, f=H("RepName")),
        C("Fund", "Fund", "calc", 6, f=H("Fund")), C("FundName", "Fund Name", "calc", 10, f=H("FundName")),
        C("Balance", "保有残高(簿価円)", "calc", 14, FMT_INT, f=H("Book")),
        C("nH", "High", "calc", 5, f='=IF({ID}="","",COUNTIF(GR_FundSev,{Fund}&"|High"))'),
        C("nM", "Medium", "calc", 5, f='=IF({ID}="","",COUNTIF(GR_FundSev,{Fund}&"|Medium"))'),
        C("nL", "Low", "calc", 5, f='=IF({ID}="","",COUNTIF(GR_FundSev,{Fund}&"|Low"))'),
        C("TopAlert", "主なアラート", "calc", 16, f='=IF({ID}="","",IFERROR(INDEX(AL_Type,MATCH({Fund},AL_Fund,0)),"—"))'),
        C("TopSev", "重要度", "calc", 7, f='=IF({ID}="","",IFERROR(INDEX(AL_Sev,MATCH({Fund},AL_Fund,0)),""))'),
        C("Weekly", "NAV週次", "calc", 8, FMT_PCT, f=H("WeeklyRet")),
        C("DaysToRed", "償還まで日数", "calc", 7, f=H("DaysToRed")),
        C("LastContact", "前回コンタクト", "calc", 10, FMT_DATE, f='=IF({ID}="","",IFERROR(IF(INDEX(CL_LastContact,MATCH(INDEX(HD_Client,{n}),CL_ID,0))="","",INDEX(CL_LastContact,MATCH(INDEX(HD_Client,{n}),CL_ID,0))),""))'),
        C("DaysSince", "経過日数", "calc", 6, f='=IF(OR({ID}="",{LastContact}=""),"",RunDate-{LastContact})'),
        C("NextFollow", "次回フォロー", "calc", 10, FMT_DATE, f='=IF({ID}="","",IFERROR(IF(INDEX(CL_NextFollow,MATCH(INDEX(HD_Client,{n}),CL_ID,0))="","",INDEX(CL_NextFollow,MATCH(INDEX(HD_Client,{n}),CL_ID,0))),""))'),
        C("Score", "Score", "calc", 6, f=('=IF({ID}="","",MIN({nH},S_SevCap)*S_SevScoreHigh+MIN({nM},S_SevCap)*S_SevScoreMed+MIN({nL},S_SevCap)*S_SevScoreLow+IF(N({Balance})>=S_LargeBalance,1,0)'
                                         '+IF(AND(ISNUMBER({DaysToRed}),N({DaysToRed})<=S_RedemptionDays),2,0)+IF(OR({LastContact}="",N({DaysSince})>S_ContactOverdueDays),1,0)'
                                         '+IF(AND({NextFollow}<>"",N({NextFollow})<=RunDate+S_FollowupWindowDays),1,0))')),
        C("Auto", "Auto Priority", "calc", 8, f='=IF({ID}="","",IF({Score}>=S_PriorityHigh,"High",IF({Score}>=S_PriorityMed,"Medium","Low")))'),
        C("Override", "Override(手動)", "in", 8),
        C("Final", "Priority", "calc", 8, f='=IF({ID}="","",IF({Override}<>"",{Override},{Auto}))'),
        C("Reason", "Reason（スコア内訳）", "calc", 44, wrap=True, f=('=IF({ID}="","",TRIM(IF({nH}>0,"High"&{nH}&"件 ","")&IF({nM}>0,"Medium"&{nM}&"件 ","")&IF({nL}>0,"Low"&{nL}&"件 ","")&IF(N({Balance})>=S_LargeBalance,"残高大 ","")'
                                                                         '&IF(AND(ISNUMBER({DaysToRed}),N({DaysToRed})<=S_RedemptionDays),"償還/満期まで"&{DaysToRed}&"日 ","")&IF(OR({LastContact}="",N({DaysSince})>S_ContactOverdueDays),"前回接触から"&IF({LastContact}="","記録なし",{DaysSince}&"日")&" ","")'
                                                                         '&IF(AND({NextFollow}<>"",N({NextFollow})<=RunDate+S_FollowupWindowDays),"フォロー予定"&TEXT({NextFollow},"mm/dd")&" ","")&IF({Override}<>"","（手動変更）","")))')),
        C("Action", "Sales Action", "calc", 34, wrap=True, f='=IF({ID}="","",IFERROR(INDEX(AL_Action,MATCH({Fund},AL_Fund,0)),"定例の運用状況共有（大きな変化なし）"))'),
        C("Note", "備考(手動)", "in", 16),
        C("SortKey", "SortKey", "aux", 9, f='=IF({ID}="","",MATCH({Final},{"Low","Medium","High"},0)*100000+{Score}*100+(100-{n}))'),
        C("EmailKey", "EmailKey", "aux", 9, f='=IF(AND({ID}<>"",{Rep}=EM_Rep),{SortKey},"")'),
    ], CAP_HOLD)
    sp.write([dict() for _ in range(CAP_HOLD)])
    dv_list(ws, f"{sp.col('Override')}{sp.first}:{sp.col('Override')}{sp.last}", "=L_Prio")
    severity_cf(ws, f"{sp.col('Final')}{sp.first}:{sp.col('Final')}{sp.last}")
    T["SP"] = sp

    r1 = 5 + CAP_HOLD + 5
    section(ws, r1 - 1, "B. 今週のフォローアップ一覧（優先度順）")
    I = lambda fld: f'=IF({{Row}}="","",INDEX(SP_{fld},{{Row}}))'
    sv = Table(wb, ws, "SV", r1, [
        C("n", "#", "aux", 4), C("Row", "行", "aux", 4, f='=IFERROR(MATCH(LARGE(SP_SortKey,{n}),SP_SortKey,0),"")'),
        C("Prio", "Priority", "calc", 8, f=I("Final")), C("Bank", "Bank", "calc", 9, f=I("Bank")), C("Fund", "Fund", "calc", 10, f=I("FundName")),
        C("Alert", "Alert", "calc", 16, f=I("TopAlert")), C("Action", "Sales Action", "calc", 34, wrap=True, f=I("Action")),
        C("Rep", "営業担当", "calc", 10, f=I("Rep")), C("Score", "Score", "calc", 5, f=I("Score")), C("Reason", "Reason", "calc", 44, wrap=True, f=I("Reason")),
    ], CAP_HOLD)
    sv.write([dict(n=i + 1) for i in range(CAP_HOLD)])
    severity_cf(ws, f"{sv.col('Prio')}{sv.first}:{sv.col('Prio')}{sv.last}")
    ws.freeze_panes = "C6"


# =====================================================================================
# 15 Commentary
# =====================================================================================
def build_commentary(wb, ws):
    title(ws, "15_Weekly_Commentary　週次分析コメント（Excel 関数による定型コメント）",
          "【事実】＝確認済みデータ／【推定】＝ポジション等からの推定／【観測】＝市場との関連が観測されるのみ／【示唆】＝営業上の示唆。Auto を修正する場合は Override 列に入力（Final が採用）。", 30)
    P = lambda fld: f"INDEX(PF_{fld},{{n}})"
    pct = lambda x: f'TEXT({x},"+0.00%;-0.00%")'
    secs = ["Perf", "Mkt", "Pos", "Risk", "TP", "Q", "A", "Next"]
    heads = ["①Performance Summary", "②Market Context", "③Position Update", "④Risk Watch", "⑤Sales Talking Points",
             "⑥想定質問", "⑥回答案", "⑦Next Week Watch"]
    md = lambda k: f'IFERROR(INDEX(MD_Display,MATCH({{{k}}},MD_Key,0)),"")'
    cols = [
        C("Fund", "Fund ID", "calc", 7, f='=IF(INDEX(FM_ID,{n})="","",INDEX(FM_ID,{n}))'),
        C("Short", "ファンド", "calc", 10, f='=IF({Fund}="","",INDEX(FM_Short,{n}))'),
        C("Strategy", "Strategy", "calc", 10, f='=IF({Fund}="","",INDEX(FM_Strategy,{n}))'),
        C("Prim", "Primary Index", "aux", 11, f='=IF({Fund}="","",INDEX(FM_PrimaryIdx,{n}))'),
        C("oa", "oa", "aux", 4, f='=IF({Fund}="","",IFERROR(MATCH({Prim},OA_ID,0),""))'),
        C("vt", "vt", "aux", 4, f='=IF({Fund}="","",IFERROR(MATCH({Prim},VS_ID,0),""))'),
        C("tc", "tc", "aux", 4, f='=IF({Fund}="","",IFERROR(MATCH({Prim},TS_ID,0),""))'),
        C("UndKey", "UndKey", "aux", 8, f='=IF({Fund}="","",IFERROR(INDEX(IM_UndKey,MATCH({Prim},IM_ID,0))&"",""))'),
        C("VolKey", "VolKey", "aux", 8, f='=IF({Fund}="","",IFERROR(INDEX(IM_VolKey,MATCH({Prim},IM_ID,0))&"",""))'),
        C("FxKey", "FxKey", "aux", 9, f='=IF({Fund}="","",INDEX(FM_FxKey,{n})&"")'),
        C("IdxName", "参照指数", "aux", 20, f='=IF({Fund}="","",IFERROR(INDEX(IM_Name,MATCH({Prim},IM_ID,0)),"未設定"))'),
        C("TopK", "TopRule#", "aux", 5, f='=IF({Fund}="","",IFERROR(INDEX(AL_k,MATCH({Fund},AL_Fund,0)),""))'),
        C("QRow", "Q用Alert行", "aux", 5, f='=IF({Fund}="","",IFERROR(MATCH({Fund}&"|1",AL_FundQ,0),""))'),
        C("NAlerts", "Alert数", "aux", 5, f='=IF({Fund}="","",COUNTIF(AL_Fund,{Fund}))'),
        C("Auto_Perf", "Auto " + heads[0], "calc", 48, wrap=True, f=(
            '=IF({Fund}="","",IF(NOT(ISNUMBER(' + P("WeeklyRet") + ')),"【データ欠損】"&TEXT(EvalDate,"yyyy/mm/dd")&"基準の公式NAVが未取得のため、週次リターンは算出していません（前回値で代用しない）。",'
            '"【事実】"&TEXT(' + P("NavDate") + ',"yyyy/mm/dd")&"基準NAV "&TEXT(' + P("NAV") + ',"#,##0.00")&"（前週比 "&' + pct(P("WeeklyRet")) + '&IF(ISNUMBER(' + P("MonthlyRet") + '),"、前月末比 "&' + pct(P("MonthlyRet")) + ',"")&"）。")'
            '&IF(ISNUMBER(' + P("IdxWoW") + '),"参照指数（"&{IdxName}&"）前週比 "&' + pct(P("IdxWoW")) + '&"。","参照指数（"&{IdxName}&"）の前週比はデータ欠損のため未算出。")'
            '&IF(' + P("Recon") + '<>"MATCH","NAV照合："&' + P("Recon") + '&"（08_NAV_History参照）。",""))')),
        C("Auto_Mkt", "Auto " + heads[1], "calc", 44, wrap=True, f=(
            '=IF({Fund}="","","【事実】"&TEXT(EvalDate,"mm/dd")&"時点：" &TEXTJOIN("、",TRUE,' + md("UndKey") + "," + md("VolKey") + "," + md("FxKey") + ')&"。")')),
        C("Auto_Pos", "Auto " + heads[2], "calc", 48, wrap=True, f=(
            '=IF({Fund}="","",IF({Strategy}="OPTION","【事実】"&IFERROR(INDEX(OA_ChgText,{oa}),"ポジションデータなし"),IF({Strategy}="VOLTARGET",IFERROR(INDEX(VS_Text,{vt}),"—"),IF({Strategy}="TRENDCARRY",IFERROR(INDEX(TS_Text,{tc}),"—"),"—")))'
            '&IF(INDEX(FM_MapCount,{n})>1,"（他"&(INDEX(FM_MapCount,{n})-1)&"指数は10_Option_Analytics参照）",""))')),
        C("Auto_Risk", "Auto " + heads[3], "calc", 48, wrap=True, arr=True, f=(
            '=IF({Fund}="","",IF({NAlerts}=0,"【事実】今週、設定閾値に抵触したアラートはありません。","【アラート】"&TEXTJOIN(" / ",TRUE,IF(AL_Fund={Fund},"["&AL_Sev&"]"&AL_Type&"："&AL_Value&IF(AL_Index="","","（"&AL_Index&"）"),""))))')),
        C("Auto_TP", "Auto " + heads[4], "calc", 44, wrap=True, f=(
            '=IF({Fund}="","","【示唆】"&IF({TopK}="",IFERROR(INDEX(ST_TP,MATCH({Strategy},ST_Code,0)),""),INDEX(RL_TP,{TopK})))')),
        C("Auto_Q", "Auto " + heads[5], "calc", 28, wrap=True, f=(
            '=IF({Fund}="","",IF({QRow}="",IFERROR(INDEX(ST_Q,MATCH({Strategy},ST_Code,0)),""),INDEX(RL_Q,INDEX(AL_k,{QRow}))))')),
        C("Auto_A", "Auto " + heads[6], "calc", 44, wrap=True, f=(
            '=IF({Fund}="","",IF({QRow}="",IFERROR(INDEX(ST_A,MATCH({Strategy},ST_Code,0)),""),SUBSTITUTE(SUBSTITUTE(INDEX(RL_A,INDEX(AL_k,{QRow})),"<v>",INDEX(AL_Value,{QRow})),"<x>",INDEX(GR_Extra,INDEX(AL_g,{QRow})))))')),
        C("Auto_Next", "Auto " + heads[7], "calc", 44, wrap=True, arr=True, f=(
            '=IF({Fund}="","","【予定】"&IF({oa}<>"",IF(INDEX(OA_MatDesc,{oa})<>"","最短満期："&INDEX(OA_MatDesc,{oa})&"。",""),"")&IF({oa}<>"",IF(N(INDEX(OA_RollNew,{oa}))>0,"ロール候補の確認（運用会社へ照会）。",""),"")'
            '&"次回公式NAV公表予定："&TEXT(WORKDAY(EvalDate+7,INDEX(FM_Lag,{n}),Holidays),"mm/dd")&"。")')),
    ]
    for s, h in zip(secs, heads):
        cols.append(C(f"Ovr_{s}", "Override " + h, "in", 22, wrap=True))
    for s, h in zip(secs, heads):
        cols.append(C(f"Final_{s}", "Final " + h, "calc", 44, wrap=True, f=f'=IF({{Fund}}="","",IF({{Ovr_{s}}}<>"",{{Ovr_{s}}},{{Auto_{s}}}))'))
    cols.append(C("EmailRisk", "メール用 注目ポイント（上位3件）", "calc", 44, wrap=True, arr=True, f=(
        '=IF({Fund}="","",IF({Ovr_Risk}<>"",{Ovr_Risk},IF({NAlerts}=0,"設定閾値に抵触したアラートなし。",TEXTJOIN(" / ",TRUE,IF((AL_Fund={Fund})*(AL_FundRank<=3),"["&AL_Sev&"]"&AL_Type&"（"&AL_Value&"）",""))'
        '&IF({NAlerts}>3,"　ほか"&({NAlerts}-3)&"件（13_Risk_Alerts）",""))))')))
    cm = Table(wb, ws, "CM", 5, cols, CAP_FUND)
    cm.write()
    for r in range(cm.first, cm.first + len(SD.FUNDS)):
        ws.row_dimensions[r].height = 120
    T["CM"] = cm
    ws["A37"] = "翌2週の市場イベント"
    ws["A37"].font = F_H2
    set_formula(ws, "D37", '=IF(COUNTIFS(EV_Date,">"&EvalDate,EV_Date,"<="&EvalDate+14)=0,"（登録なし）",TEXTJOIN("、",TRUE,IF((EV_Date>EvalDate)*(EV_Date<=EvalDate+14),TEXT(EV_Date,"mm/dd")&" "&EV_Event,"")))', array=True)
    style_value(ws["D37"], "calc")
    add_name(wb, "CM_Events", ws.title, "$D$37")

    r1 = 5 + CAP_FUND + 5
    section(ws, r1 - 1, "B. 指数別コメント（Option 指数。ポジションファイル由来。運用実績として確認できた事項と、ポジションからの推測を区別）")
    O = lambda fld: f"INDEX(OA_{fld},{{n}})"
    ci = Table(wb, ws, "CI", r1, [
        C("ID", "Index ID", "calc", 11, f='=IF(' + O("ID") + '="","",' + O("ID") + ')'),
        C("Name", "Index Name", "calc", 20, f='=IF({ID}="","",' + O("Name") + ')'),
        C("FundOf", "Primary Fund", "aux", 7, f='=IF({ID}="","",IFERROR(INDEX(MP_Fund,MATCH({ID}&"|Y",MP_IdxPrim,0)),""))'),
        C("I1", "1 指数の前週比", "calc", 30, wrap=True, f='=IF({ID}="","","【事実】"&IF(' + O("LvlChg") + '="","指数水準の前回比はポジションファイルから算出不可。","ポジションファイル上の指数水準 "&TEXT(' + O("LvlCur") + ',"#,##0.00")&"（前回比 "&' + pct(O("LvlChg")) + '&"）。"))'),
        C("I2", "2 原資産市場", "calc", 30, wrap=True, f='=IF({ID}="","","【事実】"&' + O("Und") + '&" "&TEXT(' + O("SpotCur") + ',"#,##0.00")&"（前回比 "&' + pct(O("SpotChg")) + '&"、ポジションファイル基準）。")'),
        C("I3", "3 ポジション構成", "calc", 34, wrap=True, f='=IF({ID}="","","【事実】有効"&' + O("NPos") + '&"行、Gross "&TEXT(' + O("Gross") + '/1000000,"#,##0")&"百万USD・Net "&TEXT(' + O("Net") + '/1000000,"#,##0")&"百万USD、ITM "&' + O("ITM") + '&"行、Strike5%以内 "&' + O("NearCnt") + '&"行。Σ|Weight| "&TEXT(' + O("AbsW") + ',"0.0%")&"（100%を前提としない）。")'),
        C("I4", "4 最もStrikeに近いポジション", "calc", 30, wrap=True, f='=IF({ID}="","","【事実】"&' + O("NearDesc") + '&"。")'),
        C("I5", "5 主要なポジション変化", "calc", 34, wrap=True, f='=IF({ID}="","","【事実】"&' + O("ChgText") + ')'),
        C("I6", "6 IVの変化", "calc", 26, wrap=True, f='=IF({ID}="","","【事実】加重平均IV "&TEXT(' + O("IV") + ',"0.0%")&"（前回 "&TEXT(' + O("IVPrev") + ',"0.0%")&"、"&TEXT(' + O("IVChg") + '*100,"+0.0;-0.0")&"pt）。")'),
        C("I7", "7 満期・ロール", "calc", 34, wrap=True, f='=IF({ID}="","","【事実】最短満期 "&' + O("MatDesc") + '&"。"&IF(N(' + O("RollNew") + ')>0,"【推定】ロール候補："&' + O("RollDesc") + '&"（解消と新規の組合せからの推定。確定ではない）。",""))'),
        C("I8", "8 NAVへの影響（確認状況）", "calc", 30, wrap=True, f='=IF({ID}="","",IF(COUNTIFS(AT_Fund,{FundOf},AT_Date,EvalDate,AT_Level,"CONFIRMED")>0,"【確認済】運用会社の寄与度データあり（09_Performance参照）。","【未確認】正式な寄与度データなし。NAVへの影響はポジションからの推測にとどまり、断定しない。"))'),
        C("I9", "9 説明ポイント", "calc", 34, wrap=True, f='=IF({ID}="","","【示唆】"&IF(N(' + O("Nearest") + ')<=S_NearStrike1,"Strike近傍（最接近 "&TEXT(' + O("Nearest") + ',"0.0%")&"）のポジションがあり、原資産が一段下落した場合の損益感応度が高まり得る点を、ポジションに基づく推定として説明。","Strikeまでの距離は相対的に確保（最接近 "&TEXT(' + O("Nearest") + ',"0.0%")&"）。市場急変時の感応度変化は継続確認。"))'),
        C("I10", "10 想定質問と回答案", "calc", 40, wrap=True, f='=IF({ID}="","","Q：原資産が更に5%下落した場合の影響は？　A：【推定】Strike5%以内のポジションが"&' + O("NearCnt") + '&"行あり、下落時は損失が拡大し得る構造です。正式な感応度は運用会社のリスク数値を確認のうえ回答します。")'),
    ], CAP_IDX)
    ci.write()
    ws.freeze_panes = "C6"


# =====================================================================================
# 16 Email generator
# =====================================================================================
def build_email(wb, ws):
    title(ws, "16_Email_Generator　営業担当者別 週次メール生成",
          "C4 で営業担当者を選択 → 担当先の保有ファンドを優先度順に抽出し本文を生成。本文は B25 セル（コピー用）。Outlook 下書きはマクロ QIS_CreateOutlookDrafts（自動送信はしない）。", 14)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 18
    labels = [
        (4, "営業担当者（選択）", SD.REPS[0][1], "in"),
        (5, "宛先メールアドレス", '=IFERROR(INDEX(RP_Email,MATCH(EM_Rep,RP_Name,0)),"")', "calc"),
        (6, "件名", '=S_ReportPrefix&TEXT(RunDate,"yyyy/mm/dd")', "calc"),
        (7, "ステータス", '=IF(DQ_Overall="要確認","要確認（18_Data_Quality の重大項目を確認）",IF(COUNT(SP_EmailKey)=0,"対象なし","レビュー可"))', "calc"),
        (8, "追記（任意・本文冒頭に挿入）", None, "in"),
    ]
    for r, lab, v, kind in labels:
        ws[f"B{r}"] = lab
        ws[f"B{r}"].font = F_BODY
        ws.merge_cells(f"C{r}:H{r}")
        if isinstance(v, str) and v.startswith("="):
            set_formula(ws, f"C{r}", v)
        else:
            ws[f"C{r}"] = v
        style_value(ws[f"C{r}"], kind)
    add_name(wb, "EM_Rep", ws.title, "$C$4")
    add_name(wb, "EM_To", ws.title, "$C$5")
    add_name(wb, "EM_Subject", ws.title, "$C$6")
    add_name(wb, "EM_Status", ws.title, "$C$7")
    add_name(wb, "EM_Note", ws.title, "$C$8")
    dv_list(ws, "C4", "=L_Rep", allow_blank=False)
    text_cf(ws, "C7", ["要確認（18_Data_Quality の重大項目を確認）"])

    PF = lambda fld: f'INDEX(PF_{fld},MATCH({{Fund}},PF_ID,0))'
    el = Table(wb, ws, "EL", 10, [
        C("n", "#", "aux", 3),
        C("Row", "行", "aux", 4, f='=IFERROR(MATCH(LARGE(SP_EmailKey,{n}),SP_EmailKey,0),"")'),
        C("Bank", "銀行名", "calc", 10, f='=IF({Row}="","",INDEX(SP_Bank,{Row}))'),
        C("Fund", "Fund", "calc", 6, f='=IF({Row}="","",INDEX(SP_Fund,{Row}))'),
        C("FundName", "ファンド名", "calc", 30, f='=IF({Row}="","",INDEX(FM_Name,MATCH({Fund},FM_ID,0)))'),
        C("NAV", "最新NAV", "calc", 10, FMT_NUM, f='=IF({Row}="","",' + PF("NAV") + ")"),
        C("NavDate", "NAV日付", "calc", 10, FMT_DATE, f='=IF({Row}="","",' + PF("NavDate") + ")"),
        C("Ccy", "通貨", "calc", 5, f='=IF({Row}="","",' + PF("Ccy") + ")"),
        C("WoW", "NAV前週比", "calc", 8, FMT_PCT, f='=IF({Row}="","",' + PF("WeeklyRet") + ")"),
        C("IdxWoW", "参照指数前週比", "calc", 8, FMT_PCT, f='=IF({Row}="","",' + PF("IdxWoW") + ")"),
        C("Prio", "Priority", "calc", 7, f='=IF({Row}="","",INDEX(SP_Final,{Row}))'),
        C("Alert", "主なアラート", "calc", 14, f='=IF({Row}="","",INDEX(SP_TopAlert,{Row}))'),
        C("Action", "Sales Action", "calc", 30, f='=IF({Row}="","",INDEX(SP_Action,{Row}))'),
        C("cm", "cm", "aux", 4, f='=IF({Row}="","",MATCH({Fund},CM_Fund,0))'),
        C("FundFirst", "初出", "aux", 4, f='=IF({Fund}="",0,IF(COUNTIF({Fund@1}:{Fund},{Fund})=1,1,0))'),
        C("FundSeq", "Fund#", "aux", 4, f='=IF({FundFirst}=1,COUNTIF({FundFirst@1}:{FundFirst},1),"")'),
        C("BankHigh", "要フォロー地銀", "aux", 4, f='=IF(AND({Prio}="High",COUNTIFS({Bank@1}:{Bank},{Bank},{Prio@1}:{Prio},"High")=1),1,0)'),
    ], CAP_EMAIL_ITEMS)
    el.write([dict(n=i + 1) for i in range(CAP_EMAIL_ITEMS)])
    severity_cf(ws, f"{el.col('Prio')}{el.first}:{el.col('Prio')}{el.last}")
    T["EL"] = el

    # summary helpers (row 22)
    helpers = [
        ("P22", "EM_NFunds", '=COUNTIF(EL_FundFirst,1)'),
        ("Q22", "EM_NRows", '=COUNT(SP_EmailKey)'),
        ("R22", "EM_Up", '=SUMPRODUCT((EL_FundFirst=1)*ISNUMBER(EL_WoW)*(EL_WoW>0))'),
        ("S22", "EM_Down", '=SUMPRODUCT((EL_FundFirst=1)*ISNUMBER(EL_WoW)*(EL_WoW<0))'),
        ("T22", "EM_NA", '=SUMPRODUCT((EL_FundFirst=1)*NOT(ISNUMBER(EL_WoW)))'),
        ("U22", "EM_NAlerts", '=COUNTIF(AL_RepOK,1)'),
    ]
    ws["O22"] = "集計"
    for cell, nm, f in helpers:
        set_formula(ws, cell, f)
        style_value(ws[cell], "aux")
        add_name(wb, nm, ws.title, "$" + cell[0] + "$" + cell[1:])
    set_formula(ws, "V22", '=TEXTJOIN("、",TRUE,IF(EL_BankHigh=1,EL_Bank,""))', array=True)
    add_name(wb, "EM_FollowBanks", ws.title, "$V$22")
    style_value(ws["V22"], "aux")

    ws["B24"] = "メール本文（コピー用）"
    ws["B24"].font = F_H2
    LINE0, NLINES = 40, 135
    set_formula(ws, "B25", f'=TEXTJOIN(CHAR(10),TRUE,B{LINE0}:B{LINE0 + NLINES - 1})')
    ws.merge_cells("B25:N25")
    ws["B25"].alignment = WRAP
    ws["B25"].font = Font(name="Meiryo UI", size=9)
    ws.row_dimensions[25].height = 400
    add_name(wb, "EM_Body", ws.title, "$B$25")

    ws[f"B{LINE0 - 1}"] = "本文（行ごと・数式）"
    ws[f"B{LINE0 - 1}"].font = F_H2
    lines = []
    lines += ['=EM_Rep&"さん"', '=" "', '="お疲れ様です。"',
              '="今週の担当先QISファンドの状況を共有します（評価基準日："&TEXT(EvalDate,"yyyy/mm/dd")&"）。"',
              '=IF(DQ_Overall="要確認","【要確認】データ品質チェックに未解決の重大項目があります（18_Data_Quality）。該当数値は確認後に利用してください。","")',
              '=IF(EM_Note="","",EM_Note)', '=" "', '="【今週のサマリー】"',
              '="・対象ファンド数："&EM_NFunds&"本（担当先保有 "&EM_NRows&"件）"',
              '="・NAV上昇："&EM_Up&"本／下落："&EM_Down&"本"&IF(EM_NA>0,"／NAV未更新："&EM_NA&"本","")',
              '="・主なリスクイベント："&IF(EM_NAlerts=0,"設定閾値に抵触したアラートなし",EM_NAlerts&"件（上位3件）")']
    for i in range(1, 4):
        lines.append(f'=IFERROR("　["&INDEX(AL_Sev,MATCH({i},AL_RepSeq,0))&"] "&INDEX(AL_FundName,MATCH({i},AL_RepSeq,0))&"："&INDEX(AL_Type,MATCH({i},AL_RepSeq,0))&"（"&INDEX(AL_Value,MATCH({i},AL_RepSeq,0))&"）","")')
    lines += ['="・特にフォローアップが必要な地銀："&IF(EM_FollowBanks="","なし",EM_FollowBanks)', '=" "', '="【担当先別ファンド状況】"']
    CMF = lambda s, i: f'INDEX(CM_Final_{s},INDEX(EL_cm,{i}))'
    for i in range(1, CAP_EMAIL_ITEMS + 1):
        g = f'INDEX(EL_Row,{i})=""'
        lines += [
            f'=IF({g},"","■ 銀行名："&INDEX(EL_Bank,{i})&IF(INDEX(EL_Prio,{i})="High","　★要フォロー",""))',
            f'=IF({g},"","　ファンド名："&INDEX(EL_FundName,{i}))',
            f'=IF({g},"","　最新NAV："&IF(INDEX(EL_NAV,{i})="","データ欠損",TEXT(INDEX(EL_NAV,{i}),"#,##0.00")&" "&INDEX(EL_Ccy,{i})&"（"&TEXT(INDEX(EL_NavDate,{i}),"yyyy/mm/dd")&"基準）"))',
            f'=IF({g},"","　NAV前週比："&IF(ISNUMBER(INDEX(EL_WoW,{i})),TEXT(INDEX(EL_WoW,{i}),"+0.00%;-0.00%"),"算出不可（基準日NAV未取得）"))',
            f'=IF({g},"","　参照指数前週比："&IF(ISNUMBER(INDEX(EL_IdxWoW,{i})),TEXT(INDEX(EL_IdxWoW,{i}),"+0.00%;-0.00%"),"データ欠損"))',
            f'=IF({g},"","　今週の主な動き："&{CMF("Mkt", i)}&" "&{CMF("Pos", i)})',
            f'=IF({g},"","　注目ポイント："&INDEX(CM_EmailRisk,INDEX(EL_cm,{i})))',
            f'=IF({g},"","　営業上のアクション："&INDEX(EL_Action,{i})&"。"&{CMF("TP", i)})',
            f'=IF({g},""," ")',
        ]
    lines += ['="【想定される顧客質問】"']
    for j in range(1, 4):
        cm = f'INDEX(EL_cm,MATCH({j},EL_FundSeq,0))'
        lines += [f'=IFERROR("Q（"&INDEX(CM_Short,{cm})&"）："&INDEX(CM_Final_Q,{cm}),"")',
                  f'=IFERROR("A："&INDEX(CM_Final_A,{cm}),"")']
    lines += ['=" "', '="【来週の注目事項】"']
    for j in range(1, 6):
        cm = f'INDEX(EL_cm,MATCH({j},EL_FundSeq,0))'
        lines.append(f'=IFERROR("・"&INDEX(CM_Short,{cm})&"："&INDEX(CM_Final_Next,{cm}),"")')
    lines += ['="・市場イベント（翌2週）："&CM_Events',
              '=" "',
              '="※社内限り。Bloomberg由来のデータを含むため、顧客への転送・転載は行わないでください。"',
              '="※【事実】確認済みデータ／【推定】ポジション等からの推定／【観測】関連の観測のみ／【示唆】営業上の示唆。"',
              '=" "', '="以上、よろしくお願いします。"']
    assert len(lines) <= NLINES, len(lines)
    for i in range(NLINES):
        r = LINE0 + i
        if i < len(lines):
            set_formula(ws, f"B{r}", lines[i])
        ws.merge_cells(f"B{r}:N{r}")
        ws[f"B{r}"].font = F_BODY
    ws.freeze_panes = "A10"


# =====================================================================================
# 17 Email history / 18 DQ
# =====================================================================================
def build_history(wb, ws):
    title(ws, "17_Email_History　メール生成履歴（監査用）",
          "マクロ QIS_LogEmail が件名・本文・DQ 状態・基準日を値で追記する（上書きしない）。手作業の場合は本文を値で貼り付け。", 12)
    eh = Table(wb, ws, "EH", 5, [
        C("LogID", "Log ID", "in", 18), C("GeneratedAt", "生成日時", "in", 16, "yyyy/mm/dd hh:mm"), C("EvalDate", "評価基準日", "in", 11, FMT_DATE),
        C("Rep", "営業担当", "in", 10), C("To", "宛先", "in", 24), C("Subject", "件名", "in", 40), C("Body", "本文", "in", 60),
        C("DQ", "DQ状態", "in", 8), C("Status", "メール状態", "in", 16), C("By", "生成者", "in", 10), C("Version", "Version", "in", 9),
        C("Draft", "Outlook下書き", "in", 8), C("Note", "備考", "in", 16)], 500)
    eh.write()


DQ_CHECKS = [
    ("DQ01", "評価基準日の公式NAV欠損（ACTIVE）", "NAV", "Critical", '=SUMPRODUCT((FM_Status="ACTIVE")*(PF_OnEval="N"))',
     'TEXTJOIN("、",TRUE,IF((FM_Status="ACTIVE")*(PF_OnEval="N"),FM_ID,""))'),
    ("DQ02", "NAV照合 MISMATCH", "NAV", "Critical", '=COUNTIF(RC_Status,"MISMATCH")', 'TEXTJOIN("、",TRUE,IF(RC_Status="MISMATCH",RC_Fund,""))'),
    ("DQ03", "NAV照合 WARNING／照合不能", "NAV", "Warning", '=COUNTIF(RC_Status,"WARNING")+COUNTIF(RC_Status,"DATE DIFFERENCE")+COUNTIF(RC_Status,"NOT COMPARABLE")+COUNTIF(RC_Status,"MISSING DATA")',
     'TEXTJOIN("、",TRUE,IF((RC_Status<>"MATCH")*(RC_Status<>"MISMATCH")*(RC_Status<>""),RC_Fund&":"&RC_Status,""))'),
    ("DQ04", "NAV通貨がマスタと不一致（Currency Mismatch）", "NAV", "Critical", '=COUNTIF(RC_CcyOK,0)', 'TEXTJOIN("、",TRUE,IF(RC_CcyOK=0,RC_Fund,""))'),
    ("DQ05", "NAV履歴の重複キー", "NAV", "Critical", '=COUNTIF(NV_Dup,">1")', '""'),
    ("DQ06", "異常リターン（Abnormal Return）", "NAV", "Warning", '=COUNTIF(PF_Abn,1)', 'TEXTJOIN("、",TRUE,IF(PF_Abn=1,PF_ID,""))'),
    ("DQ07", "ポジション基準日の遅延（オプション指数）", "Position", "Critical", '=COUNTIF(OA_Stale,">"&S_PosDelayBizDays)+SUMPRODUCT((OA_ID<>"")*(OA_CurDate=0))',
     'TEXTJOIN("、",TRUE,IF((OA_ID<>"")*((OA_CurDate=0)+(IFERROR(OA_Stale*1,0)>S_PosDelayBizDays)),OA_ID,""))'),
    ("DQ08", "ポジション基準日 ≠ 評価基準日（時点不一致）", "Position", "Warning", '=SUMPRODUCT((OA_ID<>"")*(OA_CurDate<>EvalDate))', 'TEXTJOIN("、",TRUE,IF((OA_ID<>"")*(OA_CurDate<>EvalDate),OA_ID,""))'),
    ("DQ09", "Index Name 未登録（Fund-Index マッピング不能）", "Position", "Critical", '=COUNTIF(PH_IndexID,"UNMAPPED")+COUNTIF(PR_Check,"Index未登録")', '""'),
    ("DQ10", "Fund-Index Mapping 不整合／Primary 未設定", "Master", "Critical", '=COUNTIF(MP_Valid,0)+SUMPRODUCT((FM_Status="ACTIVE")*(FM_PrimaryIdx=""))', '""'),
    ("DQ11", "ポジション行の重複（Position ID Duplicate）", "Position", "Critical", '=COUNTIF(PH_Dup,">1")+COUNTIF(PR_Check,"重複行")', '""'),
    ("DQ12", "Ticker 欠損（Missing Ticker）", "Position", "Critical", '=COUNTIFS(PH_PositionDate,">0",PH_Ticker,"")+COUNTIF(PR_Check,"Ticker欠損")', '""'),
    ("DQ13", "Put/Call 判定不能（UNKNOWN）", "Position", "Warning", '=COUNTIF(PH_OptType,"UNKNOWN")', '""'),
    ("DQ14", "最新ファイル内の満期経過ポジション", "Position", "Warning", '=COUNTIF(PH_Use,"CUR_EXPIRED")', '""'),
    ("DQ15", "同一基準日・同一指数で複数ファイル（再取込の衝突）", "Position", "Warning", '=COUNTIF(PH_SrcConflict,1)', '""'),
    ("DQ16", "重要市場データの欠損・STALE", "Market", "Critical", '=COUNTIFS(MD_Critical,"Y",MD_Status,"MISSING")+COUNTIFS(MD_Critical,"Y",MD_Status,"STALE")',
     'TEXTJOIN("、",TRUE,IF((MD_Critical="Y")*(MD_Status<>"OK")*(MD_Status<>""),MD_Key&"("&MD_Status&")",""))'),
    ("DQ17", "その他市場データの欠損・STALE", "Market", "Warning", '=COUNTIFS(MD_Critical,"<>Y",MD_Status,"MISSING")+COUNTIFS(MD_Critical,"<>Y",MD_Status,"STALE")',
     'TEXTJOIN("、",TRUE,IF((MD_Critical<>"Y")*(MD_Status<>"OK")*(MD_Status<>""),MD_Key&"("&MD_Status&")",""))'),
    ("DQ18", "指数マスタの市場データKeyが07に未設定", "Market", "Warning", '=SUMPRODUCT((IM_UndKey<>"")*(COUNTIF(MD_Key,IM_UndKey)=0))+SUMPRODUCT((IM_QisKey<>"")*(COUNTIF(MD_Key,IM_QisKey)=0))', '""'),
    ("DQ19", "Index Level（ファイル）と Bloomberg 指数値の乖離", "Market", "Warning", '=SUMPRODUCT(ISNUMBER(OA_LvlCheck)*(ABS(IFERROR(OA_LvlCheck*1,0))>S_IdxLvlTol))', '""'),
    ("DQ20", "ACTIVE ファンドに保有先なし（Missing Client Holdings）", "Client", "Warning", '=SUMPRODUCT((FM_ID<>"")*(FM_Status="ACTIVE")*(COUNTIF(HD_Fund,FM_ID)=0))',
     'TEXTJOIN("、",TRUE,IF((FM_ID<>"")*(FM_Status="ACTIVE")*(COUNTIF(HD_Fund,FM_ID)=0),FM_ID,""))'),
    ("DQ21", "保有データの顧客/ファンド未登録", "Client", "Critical", '=COUNTIF(HD_Valid,0)', '""'),
    ("DQ22", "Greeks の単位・符号が未確認（参考値扱い）", "Position", "Warning", '=SUMPRODUCT((IM_Strategy="OPTION")*(IM_GreeksOK<>"Y"))', '""'),
    ("DQ23", "NDX Spot 列の有効性未確認（原資産が NDX 以外の可能性）", "Position", "Warning", '=SUMPRODUCT((IM_Strategy="OPTION")*(IM_SpotValid<>"Y"))', '""'),
    ("DQ24", "市場データが MANUAL モード（Bloomberg 未使用）", "Market", "Warning", '=IF(S_DataMode="MANUAL",1,0)', '""'),
]


def build_dq(wb, ws):
    title(ws, "18_Data_Quality　データ品質・整合性チェック",
          "Critical が未解決（OPEN）の場合、メール生成ステータスは『要確認』。確認のうえ問題なしと判断した場合は Ack 欄に Y・確認者・理由を記入（監査証跡）。", 12)
    ws["B4"] = "総合判定（Critical 未解決）"
    ws["B4"].font = F_H2
    dq = Table(wb, ws, "DQ", 7, [
        C("ID", "ID", "aux", 6), C("Check", "チェック項目", "aux", 44), C("Cat", "区分", "aux", 9), C("Sev", "重大度", "aux", 8),
        C("Count", "件数", "calc", 6), C("Status", "結果", "calc", 6, f='=IF({ID}="","",IF(N({Count})>0,"NG","OK"))'),
        C("Detail", "詳細", "calc", 36), C("Ack", "Ack(Y)", "in", 6), C("AckBy", "確認者", "in", 10), C("AckNote", "確認内容", "in", 30),
        C("Eff", "判定", "calc", 8, f='=IF({ID}="","",IF({Status}="NG",IF(AND({Ack}="Y",{AckBy}<>""),"ACKED","OPEN"),"OK"))'),
    ], 30)
    dq.write([dict(ID=a, Check=b, Cat=c, Sev=d) for a, b, c, d, _, _ in DQ_CHECKS])
    for i, (_, _, _, _, f, det) in enumerate(DQ_CHECKS):
        r = dq.first + i
        set_formula(ws, f"{dq.col('Count')}{r}", f, array=True)
        set_formula(ws, f"{dq.col('Detail')}{r}", "=" + det if det != '""' else '=""', array=det != '""')
    set_formula(ws, "C4", '=IF(COUNTIFS(DQ_Sev,"Critical",DQ_Eff,"OPEN")>0,"要確認","OK")')
    style_value(ws["C4"], "calc")
    ws["C4"].font = F_BIG
    set_formula(ws, "D4", '="Critical OPEN "&COUNTIFS(DQ_Sev,"Critical",DQ_Eff,"OPEN")&"件／Warning NG "&COUNTIFS(DQ_Sev,"Warning",DQ_Status,"NG")&"件"')
    add_name(wb, "DQ_Overall", ws.title, "$C$4")
    text_cf(ws, "C4", ["要確認"])
    text_cf(ws, f"{dq.col('Eff')}{dq.first}:{dq.col('Eff')}{dq.last}", ["OPEN"])
    dv_list(ws, f"{dq.col('Ack')}{dq.first}:{dq.col('Ack')}{dq.last}", '"Y,N"')
    T["DQ"] = dq


# =====================================================================================
# 01 Dashboard
# =====================================================================================
def build_dashboard(wb, ws):
    title(ws, "01_Dashboard　QIS Fund Monitoring Weekly Dashboard",
          "上部：更新状況　中段：ファンド別サマリー（絞込可）　下段：重要リスク・チャート。赤はリスクアラートのみ。", 17)
    widths = {"A": 2, "B": 4, "C": 11, "D": 11, "E": 11, "F": 10, "G": 9, "H": 9, "I": 9, "J": 10, "K": 11, "L": 9, "M": 15, "N": 10, "O": 22, "P": 22}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    cards = [("B", "C", "評価基準日", "=EvalDate", FMT_DATE), ("D", "E", "作成日", "=RunDate", FMT_DATE), ("F", "G", "データモード", "=S_DataMode", None),
             ("H", "I", "データ品質", "=DQ_Overall", None), ("J", "J", "High", '=COUNTIF(AL_Sev,"High")', FMT_INT),
             ("K", "K", "Medium", '=COUNTIF(AL_Sev,"Medium")', FMT_INT), ("L", "L", "Low", '=COUNTIF(AL_Sev,"Low")', FMT_INT),
             ("M", "N", "メール", "=EM_Status", None), ("O", "P", "最新ポジション基準日", "=MAX(OA_CurDate)", FMT_DATE)]
    for col, col2, lab, f, fmt in cards:
        ws[f"{col}4"] = lab
        ws[f"{col}4"].font = F_SMALL
        set_formula(ws, f"{col}5", f)
        ws[f"{col}5"].font = Font(name="Meiryo UI", size=11, bold=True, color="1F3864")
        ws[f"{col}5"].fill = FILL_GREY
        ws[f"{col}5"].alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        if col2 != col:
            ws.merge_cells(f"{col}5:{col2}5")
        if fmt:
            ws[f"{col}5"].number_format = fmt
    ws.row_dimensions[5].height = 22
    text_cf(ws, "H5", ["要確認"])
    ws.conditional_formatting.add("J5", FormulaRule(formula=["J5>0"], fill=FILL_RED, font=Font(name="Meiryo UI", size=12, bold=True, color="FFFFFF")))

    ws["B7"] = "絞込 →"
    for col, lab, nm, lst, default in [("C", "投資家", "Dash_Bank", "L_FBank", "(全て)"), ("E", "営業担当", "Dash_Rep", "L_FRep", "(全て)"),
                                        ("G", "戦略", "Dash_Strategy", "L_FStrategy", "(全て)"), ("J", "チャート対象 Fund ID", "Dash_FundID", "L_FFund", "F001")]:
        ws[f"{col}7"] = lab
        ws[f"{col}7"].font = F_SMALL
        nxt = chr(ord(col) + 1)
        ws[f"{nxt}7"] = default
        style_value(ws[f"{nxt}7"], "in")
        add_name(wb, nm, ws.title, f"${nxt}$7")
        dv_list(ws, f"{nxt}7", f"={lst}", allow_blank=False)

    # helper (hidden-ish) columns AA..AC aligned with FM
    for j, h in enumerate(["Fund", "表示", "#"]):
        ws.cell(9, 27 + j, h).font = F_SMALL
    for n in range(1, CAP_FUND + 1):
        r = 9 + n
        set_formula(ws, f"AA{r}", f'=IF(INDEX(FM_ID,{n})="","",INDEX(FM_ID,{n}))')
        set_formula(ws, f"AB{r}", (f'=IF(AA{r}="",0,IF(AND(INDEX(FM_Status,{n})="ACTIVE",OR(Dash_Strategy="(全て)",INDEX(FM_Strategy,{n})=Dash_Strategy),'
                                   f'OR(Dash_Bank="(全て)",COUNTIFS(HD_Fund,AA{r},HD_Bank,Dash_Bank)>0),OR(Dash_Rep="(全て)",COUNTIFS(HD_Fund,AA{r},HD_RepName,Dash_Rep)>0)),1,0))'))
        set_formula(ws, f"AC{r}", f'=IF(AB{r}=1,COUNTIF(AB$10:AB{r},1),"")')
        for c in ("AA", "AB", "AC"):
            style_value(ws[f"{c}{r}"], "aux")
    add_name(wb, "DB_Seq", ws.title, f"$AC$10:$AC${9 + CAP_FUND}")

    section(ws, 8, "ファンド別サマリー")
    heads = [("B", "#"), ("C", "Fund"), ("D", "Strategy"), ("E", "NAV日付"), ("F", "NAV"), ("G", "週次"), ("H", "月次"), ("I", "YTD"),
             ("J", "参照指数週次"), ("K", "現在DD"), ("L", "実現Vol"), ("M", "NAV照合"), ("N", "H/M/L"), ("O", "主なアラート"), ("P", "保有地銀")]
    for col, h in heads:
        c = ws[f"{col}9"]
        c.value = h
        c.font = F_HEAD
        c.fill = FILL_HEAD
        c.alignment = CENTER
    NROW = 15
    for n in range(1, NROW + 1):
        r = 9 + n
        ws[f"B{r}"] = n
        rr = f'MATCH({n},DB_Seq,0)'
        P = lambda fld: f'=IFERROR(INDEX(PF_{fld},{rr}),"")'
        fs = {"C": P("Short"), "D": P("Strategy"), "E": P("NavDate"), "F": P("NAV"), "G": P("WeeklyRet"), "H": P("MonthlyRet"),
              "I": P("YTD"), "J": P("IdxWoW"), "K": P("CurDD"), "L": P("Vol"), "M": P("Recon"),
              "N": f'=IFERROR(COUNTIF(GR_FundSev,INDEX(PF_ID,{rr})&"|High")&"/"&COUNTIF(GR_FundSev,INDEX(PF_ID,{rr})&"|Medium")&"/"&COUNTIF(GR_FundSev,INDEX(PF_ID,{rr})&"|Low"),"")',
              "O": f'=IFERROR(INDEX(AL_Type,MATCH(INDEX(PF_ID,{rr}),AL_Fund,0)),IF(ISNUMBER({rr}),"—",""))',
              }
        fmts = {"E": FMT_DATE, "F": FMT_NUM, "G": FMT_PCT, "H": FMT_PCT, "I": FMT_PCT, "J": FMT_PCT, "K": FMT_PCT, "L": FMT_PCT}
        for col, f in fs.items():
            set_formula(ws, f"{col}{r}", f)
            style_value(ws[f"{col}{r}"], "calc", fmts.get(col))
        set_formula(ws, f"P{r}", f'=IFERROR(TEXTJOIN("、",TRUE,IF(HD_Fund=INDEX(PF_ID,{rr}),HD_Bank,"")),"")', array=True)
        style_value(ws[f"P{r}"], "calc")
        style_value(ws[f"B{r}"], "aux")
    text_cf(ws, f"M10:M{9 + NROW}", ["MISMATCH"])

    r0 = 9 + NROW + 3
    section(ws, r0 - 1, "重要リスク（上位10件・絞込反映）")
    heads2 = [("B", "#"), ("C", "Severity"), ("D", "Fund"), ("E", "Index"), ("F", "Alert Type"), ("I", "Value"), ("J", "Threshold"), ("K", "Obs Date"),
              ("L", "保有地銀"), ("N", "Recommended Sales Action")]
    for col, h in heads2:
        c = ws[f"{col}{r0}"]
        c.value = h
        c.font = F_HEAD
        c.fill = FILL_HEAD
    for n in range(1, 11):
        r = r0 + n
        ws[f"B{r}"] = n
        rr = f'MATCH({n},AL_DashSeq,0)'
        A = lambda fld: f'=IFERROR(INDEX(AL_{fld},{rr}),"")'
        for col, fld in [("C", "Sev"), ("D", "FundName"), ("E", "Index"), ("F", "Type"), ("I", "Value"), ("J", "Thr"), ("K", "Obs"), ("L", "Investors"), ("N", "Action")]:
            set_formula(ws, f"{col}{r}", A(fld))
            style_value(ws[f"{col}{r}"], "calc", FMT_DATE if col == "K" else None)
        ws.merge_cells(f"F{r}:H{r}")
        ws.merge_cells(f"L{r}:M{r}")
        ws.merge_cells(f"N{r}:P{r}")
    severity_cf(ws, f"C{r0 + 1}:C{r0 + 10}")

    # chart helper
    rc0 = r0 + 14
    section(ws, rc0 - 1, "NAV（TRI）と参照指数の推移（直近26週・起点=100）／ファンド別 週次リターン")
    ws["AE" + str(rc0)] = "Date"
    ws["AF" + str(rc0)] = "NAV(TRI)"
    ws["AG" + str(rc0)] = "参照指数"
    ws["AH" + str(rc0)] = "NAV(起点100)"
    ws["AI" + str(rc0)] = "参照指数(起点100)"
    set_formula(ws, "AK" + str(rc0), '=IFERROR(INDEX(IM_QisKey,MATCH(INDEX(FM_PrimaryIdx,MATCH(Dash_FundID,FM_ID,0)),IM_ID,0)),"")')
    add_name(wb, "Dash_IdxKey", ws.title, f"$AK${rc0}")
    for k in range(27):
        r = rc0 + 1 + k
        set_formula(ws, f"AE{r}", f"=EvalDate-7*{26 - k}")
        ws[f"AE{r}"].number_format = "mm/dd"
        set_formula(ws, f"AF{r}", f'=IFERROR(INDEX(NV_TRI,MATCH(Dash_FundID&"|OFFICIAL|"&MAXIFS(NV_Date,NV_Fund,Dash_FundID,NV_Source,"OFFICIAL",NV_Date,"<="&AE{r}),NV_Key,0)),NA())')
        set_formula(ws, f"AG{r}", f'=IFERROR(IF(MAXIFS(SR_Date,SR_Key,Dash_IdxKey,SR_Date,"<="&AE{r})<AE{r}-3,NA(),INDEX(SR_Value,MATCH(Dash_IdxKey&"|"&MAXIFS(SR_Date,SR_Key,Dash_IdxKey,SR_Date,"<="&AE{r}),SR_KD,0))),NA())')
        set_formula(ws, f"AH{r}", f"=IFERROR(AF{r}/AF${rc0 + 1}*100,NA())")
        set_formula(ws, f"AI{r}", f"=IFERROR(AG{r}/AG${rc0 + 1}*100,NA())")
    lc = LineChart()
    lc.title = "NAV(TRI)と参照指数"
    lc.height, lc.width = 7.5, 16
    lc.y_axis.title = "起点=100"
    data = Reference(ws, min_col=34, max_col=35, min_row=rc0, max_row=rc0 + 27)
    lc.add_data(data, titles_from_data=True)
    lc.set_categories(Reference(ws, min_col=31, min_row=rc0 + 1, max_row=rc0 + 27))
    lc.series[0].graphicalProperties.line.solidFill = "1F3864"
    lc.series[1].graphicalProperties.line.solidFill = "A6A6A6"
    lc.legend.position = "b"
    ws.add_chart(lc, f"B{rc0 + 1}")
    # weekly return bar (by fund master order)
    ws["AM" + str(rc0)] = "Fund"
    ws["AN" + str(rc0)] = "週次リターン"
    for n in range(1, 11):
        r = rc0 + n
        set_formula(ws, f"AM{r}", f'=IF(INDEX(PF_Short,{n})="","",INDEX(PF_Short,{n}))')
        set_formula(ws, f"AN{r}", f'=IF(ISNUMBER(INDEX(PF_WeeklyRet,{n})),INDEX(PF_WeeklyRet,{n}),NA())')
        ws[f"AN{r}"].number_format = FMT_PCT
    bc = BarChart()
    bc.title = "週次リターン（公式NAV・TR）"
    bc.height, bc.width = 7.5, 13
    bc.add_data(Reference(ws, min_col=40, min_row=rc0, max_row=rc0 + len(SD.FUNDS)), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=39, min_row=rc0 + 1, max_row=rc0 + len(SD.FUNDS)))
    bc.series[0].graphicalProperties.solidFill = "2F5496"
    bc.legend = None
    bc.x_axis.tickLblPos = "low"
    ws.add_chart(bc, f"J{rc0 + 1}")
    ws.freeze_panes = "A8"


# =====================================================================================
# 20 Instructions / 21 status / 22 tests
# =====================================================================================
INSTRUCTIONS = [
    ("1. 目的", "地銀担当営業向け週次 QIS モニタリング（Sales Intelligence）。NAV・参照指数・ポジション・市場データを統合し、アラート・優先度・コメント・メール本文を生成する。"),
    ("2. シート構成", "Raw Data：05/06/07/08（＋11・12・09 の入力表）｜Master：02/03/04/19｜Analytics：09〜14｜Output：01/15/16｜Audit：17/18/21/22。"),
    ("3. 毎週の手順", "00_Control の D 表（手順 1〜12）の順に実行。自動ステータス列で各工程の結果を確認し、完了欄に Y を入力。"),
    ("3-1 ポジション取込", "最新ファイルを tblPQParams.PositionFolder に保存 → データ＞すべて更新（qPositionLatest が 05 に出力）→ QIS_ArchivePositions で 06 に値で追記。同一 Snapshot ID（基準日｜Index Name｜ファイル名）は追記しない。"),
    ("3-2 市場データ", "Bloomberg 接続時：19_Settings の DataMode を BLOOMBERG にし、Bloomberg リボンの Refresh（または QIS_RefreshBloomberg）。未接続時：MANUAL のまま 07 の K〜N 列に値と As-of を入力（CSV からの貼付け可）。欠損・古い値は自動的に採用されない。"),
    ("3-3 NAV", "運用会社の公式 NAV を 08 の A〜I 列（Source=OFFICIAL）に追記。Bloomberg NAV は Source=BLOOMBERG で追記（照合用）。分配金は分配金表へ。"),
    ("3-4 確認", "18_Data_Quality の OPEN（Critical）を解消、または確認のうえ Ack 欄に記入。13_Risk_Alerts の Review Log に確認状況を記録。"),
    ("3-5 メール", "16_Email_Generator で担当者を選択 → 本文をレビュー（修正は 15 の Override 列、または 16 の追記欄）→ QIS_CreateOutlookDrafts で下書き作成（送信は手動）→ QIS_LogEmail で履歴保存。"),
    ("4. 入力セル", "薄黄色＋青字のみ入力。白セルは数式（上書き禁止）。灰色は補助計算。"),
    ("5. 新規追加", "ファンド：02 A 表に行追加＋C 表に Mapping。指数：02 B 表（Index Name はファイル表記と一致）。地銀：03。保有：04。ルール：19 B 表（Metric Key は 13 B 表の見出しから選択）。容量：ファンド30・指数30・Mapping30・保有100・ルール40・ポジション履歴1,500行。"),
    ("6. 判定ルールの要点", "Strike Distance＝(Spot−Strike)/Spot。Put は正で OTM、Call は負で OTM、|距離|≤ATM 幅で ATM。満期日＜評価基準日は有効ポジションから除外（履歴には残る）。ロールは同一指数・同一原資産・同一 Put/Call で旧満期の解消/満期と新満期の新規が同時に起きた場合の『推定』。"),
    ("7. 区分表示", "【事実】確認済みデータ／【推定】ポジション等からの推定／【観測】関連の観測のみ／【示唆】営業上の示唆。寄与度は Confirmed／Estimated／Observed。"),
    ("8. 禁止事項", "Bloomberg 由来データの顧客向け転送・再配布をしない（契約条件に従う）。顧客情報を未承認の外部サービスへ送信しない。欠損値を前回値で代用しない。メールは自動送信しない。"),
    ("9. マクロ", "vba/modQIS.bas を VBE でインポートし .xlsm で保存。QIS_RefreshAll／QIS_RefreshBloomberg／QIS_ArchivePositions／QIS_LogEmail／QIS_CreateOutlookDrafts／QIS_CopyEmailBody。"),
    ("10. Power Query", "powerquery/*.pq を『データ＞データの取得＞その他＞空のクエリ』の詳細エディターに貼付け。qPositionLatest の読込先を 05_Position_Raw の A6 に設定（または既存範囲を置換）。仕様は docs/03_power_query_spec.md。"),
]

IMPL = [
    ("Fully Implemented", "Excel 数式", "ポジション標準化（Index ID・Put/Call 判定・Strike Distance・Moneyness・DTE・Position Status）／前回比較（New・Closed・Expired・Increased・Reduced・Possible Roll・Strike Change・ITM 変化）／指数単位集計／NAV 照合 5 区分＋比較不能／TR リターン（週次・月次・YTD・設定来・取得来）／MDD・現在 DD・実現 Vol／Vol Target・Trend/Carry 分析／ルールベースのアラート（29 ルール、閾値は Settings）／営業優先度（スコア内訳表示・手動 Override）／定型コメント（Fund 7 項目・Option 指数 10 項目）／営業担当者別メール本文／データ品質 24 項目／ダッシュボード（絞込・チャート）"),
    ("Requires Bloomberg Connection", "BDH／BDP", "07 の BDH 列（評価基準日・前週・前月末）・BDP ライブ値・オプション照合（PX_MID・IVOL_MID・OPT_DELTA_MID）・接続テスト。未接続時は NO_CONN となり、MANUAL モードで手入力値を使用。フィールド名・Ticker・単位は端末で要確認。"),
    ("Requires Local Excel Configuration", "Power Query", "qPositionLatest／qPositionAll／qNavOfficial／qMarketCsv の M コードは powerquery/ に同梱。クエリの作成と 05 への読込設定はローカルの Excel で実施（openpyxl では Power Query を埋め込めないため、ブックには含まれていない）。"),
    ("Requires VBA / Outlook Permission", "VBA", "vba/modQIS.bas（更新一括実行・ポジション履歴追記（重複防止）・メール履歴保存・Outlook 下書き作成・本文コピー）。.xlsm への保存とマクロ有効化、Outlook オブジェクトモデルへのアクセス許可が必要。"),
    ("Requires Additional Data", "データ", "Vol Target 項目（Target/Realized Vol・Exposure・上下限・リバランス）と Trend/Carry 項目（Signal・Weight・Carry・Roll Yield・寄与）はポジションファイルにないため、運用会社の週次データが必要（11・12 の入力表）。正式な寄与度データ（09 C 表）。公式 NAV・分配金。"),
    ("Future Enhancement", "—", "高度な自然言語生成（社内承認済み AI 環境が前提）／アセットクラス別寄与変化アラート／Greeks の公式ポートフォリオ感応度（単位確認後）／BDS による分配金自動取得／複数シェアクラスの自動照合／ポジション明細の Trade ID（取引識別子）が提供された場合の取引単位比較。"),
]


def build_docs_sheets(wb, results):
    ws = wb["20_Instructions"]
    title(ws, "20_Instructions　操作マニュアル（要約）", "詳細は docs/06_operation_manual.md。", 4)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 120
    for i, (h, t) in enumerate(INSTRUCTIONS):
        r = 5 + i
        ws[f"A{r}"] = h
        ws[f"A{r}"].font = F_H2
        ws[f"B{r}"] = t
        ws[f"B{r}"].alignment = WRAP
        ws[f"B{r}"].font = F_BODY
        ws.row_dimensions[r].height = 46

    ws = wb["21_Impl_Status"]
    title(ws, "21_Impl_Status　実装ステータス／未実装機能／本番導入前の確認事項", "実装できていない機能を実装済みと表現しないための一覧。", 4)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 120
    for j, h in enumerate(["区分", "手段", "内容"]):
        c = ws.cell(5, 1 + j, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
    for i, row in enumerate(IMPL):
        for j, v in enumerate(row):
            c = ws.cell(6 + i, 1 + j, v)
            c.alignment = WRAP
            c.font = F_BODY
            c.border = BORDER
        ws.row_dimensions[6 + i].height = 70
    from preprod import PREPROD
    r0 = 6 + len(IMPL) + 2
    ws[f"A{r0}"] = "本番導入前の確認事項"
    ws[f"A{r0}"].font = F_H2
    for j, h in enumerate(["#", "区分", "確認事項"]):
        c = ws.cell(r0 + 1, 1 + j, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
    for i, (cat, item) in enumerate(PREPROD):
        r = r0 + 2 + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, cat)
        ws.cell(r, 3, item).alignment = WRAP
        for j in range(1, 4):
            ws.cell(r, j).border = BORDER
            ws.cell(r, j).font = F_BODY

    ws = wb["22_Test_Results"]
    title(ws, "22_Test_Results　テスト結果（架空データ）", "verify_workbook.py が LibreOffice で再計算した値を、Python で独立に計算した期待値と照合した結果。", 6)
    widths = [8, 30, 60, 28, 28, 8]
    for j, w in enumerate(widths):
        ws.column_dimensions[chr(ord("A") + j)].width = w
    if not results:
        ws["A5"] = "未実行（verify_workbook.py を実行するとここに結果が入る）"
        return
    ws["A4"] = f"実行日時：{results['run_at']}　合計 {results['total']} 件　合格 {results['passed']} 件　不合格 {results['failed']} 件"
    ws["A4"].font = F_H2
    for j, h in enumerate(["ID", "テスト", "確認内容", "期待値", "結果値", "判定"]):
        c = ws.cell(5, 1 + j, h)
        c.font = F_HEAD
        c.fill = FILL_HEAD
    for i, t in enumerate(results["tests"]):
        r = 6 + i
        for j, k in enumerate(["id", "scenario", "check", "expected", "actual", "result"]):
            c = ws.cell(r, 1 + j, str(t[k]))
            c.font = F_BODY
            c.alignment = WRAP
            c.border = BORDER
        if t["result"] != "PASS":
            ws.cell(r, 6).fill = FILL_RED


# =====================================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--results", default=str(ROOT / "dist" / "test_results.json"))
    ap.add_argument("--out", default=str(OUT))
    a = ap.parse_args()
    results = None
    if a.results and Path(a.results).exists():
        results = json.loads(Path(a.results).read_text(encoding="utf-8"))

    wb = Workbook()
    wb.remove(wb.active)
    for name, _ in SHEETS:
        wb.create_sheet(name)
    build_settings(wb, wb["19_Settings"])
    build_control(wb, wb["00_Control"])
    build_masters(wb)
    build_positions(wb)
    build_bloomberg(wb, wb["07_Bloomberg_Data"])
    build_nav(wb, wb["08_NAV_History"])
    build_perf(wb, wb["09_Performance"])
    build_option(wb, wb["10_Option_Analytics"])
    build_vt(wb, wb["11_VolTarget_Analytics"])
    build_tc(wb, wb["12_TrendCarry_Analytics"])
    build_alerts(wb, wb["13_Risk_Alerts"])
    build_priority(wb, wb["14_Sales_Priority"])
    build_commentary(wb, wb["15_Weekly_Commentary"])
    build_email(wb, wb["16_Email_Generator"])
    build_history(wb, wb["17_Email_History"])
    build_dq(wb, wb["18_Data_Quality"])
    build_dashboard(wb, wb["01_Dashboard"])
    build_docs_sheets(wb, results)

    tabs = {"00": "1F3864", "01": "1F3864", "02": "808080", "03": "808080", "04": "808080", "05": "A6A6A6", "06": "A6A6A6",
            "07": "A6A6A6", "08": "A6A6A6", "09": "2F5496", "10": "2F5496", "11": "2F5496", "12": "2F5496", "13": "2F5496",
            "14": "2F5496", "15": "1F3864", "16": "1F3864", "17": "595959", "18": "595959", "19": "808080", "20": "595959",
            "21": "595959", "22": "595959"}
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = tabs[ws.title[:2]]
    wb.calculation.fullCalcOnLoad = True
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    wb.save(a.out)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
