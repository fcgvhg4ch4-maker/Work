"""SYNTHETIC TEST FIXTURES — NOT REAL DATA.

Used only to build dist/QIS_Fund_Monitoring_TESTDATA.xlsx and to run verify.py.
Every fund / investor / price below is invented for testing formula behaviour
(scenario coverage), and must never be used for client communication.
"""
import math
from datetime import date, datetime, timedelta

EVAL = date(2026, 10, 30)
HOLIDAYS = [date(2026, 10, 12)]
TAG = "TEST_FIXTURE"


def bdays(start, end):
    d = start
    while d <= end:
        if d.weekday() < 5 and d not in HOLIDAYS:
            yield d
        d += timedelta(days=1)


def nav_series():
    rows = []
    # F001: rises then falls ~6% in the last weeks; last NAV 2026-10-28 (stale 2bd)
    t = 0
    nav = 10000.0
    for d in bdays(date(2025, 6, 2), date(2026, 10, 28)):
        if t:
            r = 0.0004 + 0.003 * math.sin(t / 9.0)
            if d >= date(2026, 9, 21):
                r = -0.0045 + 0.001 * math.cos(t)
            nav *= 1 + r
        rows.append({"fund_code": "F001", "nav_date": d, "nav": round(nav, 4),
                     "aum": round(nav * (10000 if d < date(2026, 10, 26) else 11500), 0),
                     "aum_currency": "JPY", "source": TAG, "source_file": "nav_F001_TEST.csv"})
        t += 1
    # F002: steady, with one distribution and a confirmed cumulative TR index
    t = 0
    for d in bdays(date(2025, 9, 1), date(2026, 10, 30)):
        nav = 10000 * (1.0003 ** t) * (1 + 0.002 * math.sin(t / 5.0))
        factor = 1.02 if d >= date(2026, 3, 31) else 1.0
        rows.append({"fund_code": "F002", "nav_date": d, "nav": round(nav, 4),
                     "nav_with_distribution": round(nav * factor, 4),
                     "distribution_amount": 200 if d == date(2026, 3, 31) else None,
                     "aum": round(nav * 20000, 0), "aum_currency": "JPY", "source": TAG,
                     "source_file": "nav_F002_TEST.csv"})
        t += 1
    # F003: short history + one duplicated date (2026-09-15)
    t = 0
    for d in bdays(date(2026, 8, 3), date(2026, 10, 30)):
        nav = 1000 * (1 + 0.0002 * t)
        rows.append({"fund_code": "F003", "nav_date": d, "nav": round(nav, 4), "aum": None,
                     "source": TAG, "source_file": "nav_F003_TEST.csv"})
        if d == date(2026, 9, 15):
            rows.append({"fund_code": "F003", "nav_date": d, "nav": round(nav * 1.01, 4),
                         "source": TAG, "source_file": "nav_F003_TEST_dup.csv"})
        t += 1
    return rows


def market_series():
    rows = []
    t = 0
    for d in bdays(date(2025, 6, 2), date(2026, 10, 30)):
        nky = 39000 + 900 * math.sin(t / 20.0)
        spx = 5600 + 150 * math.sin(t / 15.0)
        if d == date(2026, 10, 23):
            nky = 39500
        if d == date(2026, 10, 30):
            nky, spx = 39000, 5800
        rows.append({"series_id": "NKY Index", "series_name": "Nikkei 225 (TEST)", "data_date": d,
                     "value": round(nky, 2), "currency": "JPY", "hedge_type": "Unhedged", "series_type": "Spot",
                     "source": TAG, "source_file": "mkt_TEST.csv"})
        rows.append({"series_id": "SPX Index", "series_name": "S&P 500 (TEST)", "data_date": d,
                     "value": round(spx, 2), "currency": "USD", "hedge_type": "Unhedged", "series_type": "Spot",
                     "source": TAG, "source_file": "mkt_TEST.csv"})
        t += 1
    return rows


def leg(src, key, pdate, recv, fname, seq, ticker, entry, weight, **kw):
    r = {"source_type": src, "strategy_data_key": key, "position_date": pdate, "received_datetime": recv,
         "file_name": fname, "file_seq": seq, "BloombergTicker": ticker, "date_of_entry": entry, "weight": weight,
         "greeks_basis": kw.pop("greeks_basis", "Unknown"), "data_quality_raw": TAG}
    r.update(kw)
    return r


def positions():
    P, A, B = [], "STEP_NKY_A", "STEP_NKY_B"
    d0, d1 = date(2026, 10, 23), date(2026, 10, 30)
    r0 = datetime(2026, 10, 23, 18, 0)
    r1a, r1b = datetime(2026, 10, 30, 17, 0), datetime(2026, 10, 30, 19, 0)
    # ---- STEP_NKY_A previous snapshot
    for tk, en, w in [("NKY 12/11/26 P38000 Index", date(2026, 9, 11), 0.25),
                      ("NKY 11/13/26 P42000 Index", date(2026, 8, 14), 0.25),
                      ("NKY 12/11/26 C45000 Index", date(2026, 9, 11), 0.10),
                      ("NKY 10/23/26 P39000 Index", date(2026, 7, 24), 0.20),
                      ("NKY 12/11/26 P35000 Index", date(2026, 9, 11), 0.20)]:
        P.append(leg("Composition", A, d0, r0, "comp_A_20261023_TEST.csv", 1, tk, en, w, delta_reported=-0.3))
    # ---- STEP_NKY_A current, superseded version v1 (must NOT be used)
    P.append(leg("Composition", A, d1, r1a, "comp_A_20261030_v1_TEST.csv", 1, "NKY 12/11/26 P38000 Index",
                 date(2026, 9, 11), 0.99))
    # ---- STEP_NKY_A current, latest version v2
    cur = [("NKY 12/11/26 P38000 Index", date(2026, 9, 11), 0.25, {}),
           ("NKY 11/13/26 P42000 Index", date(2026, 8, 14), 0.25, {}),
           ("NKY 01/15/27 P36000 Index", date(2026, 10, 28), 0.20, {}),
           ("NKY 12/11/26 C45000 Index", date(2026, 9, 11), 0.10, {}),
           ("NKY 12/11/26 P37000 Index", date(2026, 9, 11), 0.10, {"spot_raw": 5800, "spot_underlying_raw": "SPX Index"}),
           ("BADTICKER", date(2026, 9, 11), 0.05, {}),
           ("NKY 12/11/26 P38000 Index", date(2026, 9, 11), 0.05, {})]   # duplicate key without leg_id
    for tk, en, w, kw in cur:
        P.append(leg("Composition", A, d1, r1b, "comp_A_20261030_v2_TEST.csv", 2, tk, en, w, delta_reported=-0.3, **kw))
    for tk, en, w in [("NKY 12/11/26 P38000 Index", date(2026, 9, 11), 0.25),
                      ("NKY 11/13/26 P42000 Index", date(2026, 8, 14), 0.25),
                      ("NKY 12/11/26 C45000 Index", date(2026, 9, 11), 0.10)]:
        P.append(leg("Portfolio", A, d1, r1b, "port_A_20261030_TEST.csv", 1, tk, en, w))
    # ---- STEP_NKY_B (clean)
    g = "Per unit (weight-scalable)"
    for tk, en, w, dl in [("NKY 12/11/26 P36000 Index", date(2026, 9, 11), 0.50, -0.20),
                          ("NKY 11/13/26 P35000 Index", date(2026, 8, 14), 0.50, -0.10)]:
        P.append(leg("Composition", B, d0, r0, "comp_B_20261023_TEST.csv", 1, tk, en, w, delta_reported=dl, greeks_basis=g))
    for tk, en, w, dl in [("NKY 12/11/26 P36000 Index", date(2026, 9, 11), 0.55, -0.25),
                          ("NKY 11/13/26 P35000 Index", date(2026, 8, 14), 0.30, -0.10),
                          ("SPX US 12/18/26 P5000 Index", date(2026, 10, 29), 0.15, -0.05)]:
        P.append(leg("Composition", B, d1, r1b, "comp_B_20261030_TEST.csv", 1, tk, en, w, delta_reported=dl, greeks_basis=g))
    return P


def fixture_data(**over):
    funds = [
        dict(fund_code="F001", fund_name="TEST STEP Fund A", fund_status="Live", strategy_type="STEP",
             index_name="TEST STEP NKY Index", currency="JPY", inception_date=date(2025, 6, 2)),
        dict(fund_code="F002", fund_name="TEST STEP Fund B", fund_status="Live", strategy_type="STEP",
             index_name="TEST STEP NKY Index", currency="JPY", inception_date=date(2025, 9, 1)),
        dict(fund_code="F003", fund_name="TEST Target Return C", fund_status="Live", strategy_type="Target Return",
             index_name="TEST TR Index", currency="JPY", inception_date=date(2026, 8, 1)),
        dict(fund_code="F004", fund_name="TEST CTA D", fund_status="Live", strategy_type="CTA",
             index_name="TEST CTA Index", currency="USD", inception_date=date(2026, 1, 5)),
        dict(fund_code="F005", fund_name="TEST Rates E", fund_status="Live", strategy_type="Rates",
             index_name="TEST Rates Index", currency="JPY"),
        dict(fund_code="F005", fund_name="TEST Rates E (dup row)", fund_status="Live", strategy_type="Rates",
             index_name="TEST Rates Index", currency="JPY"),
        dict(fund_code="F006", fund_name="TEST STEP Fund F", fund_status="Live", strategy_type="STEP",
             index_name="TEST STEP NKY Index", currency="JPY"),
    ]
    for f in funds:
        f.update(source_file="live_fund_list_TEST.xlsx", load_date=EVAL)
    hold = [("B001", "TEST Bank Alpha", "F001", 100_000_000), ("B001", "TEST Bank Alpha", "F001", 50_000_000),
            ("B001", "TEST Bank Alpha", "F002", 200_000_000), ("B002", "TEST Bank Beta", "F001", 80_000_000),
            ("B003", "TEST Bank Gamma", "F003", 30_000_000), ("B003", "TEST Bank Gamma", "F003", "n/a"),
            ("B004", "TEST Bank Delta", "F999", 10_000_000), ("B002", "TEST Bank Beta", "F002", 25_000_000)]
    holdings = [dict(as_of_date=EVAL, end_investor_bdr_id=i, end_investor_name=n, fund_code=f, amount=a, currency="JPY",
                     sales_rep="TEST Rep", source_file="end_investor_holdings_view_TEST.xlsx", source_row=k + 2)
                for k, (i, n, f, a) in enumerate(hold)]
    fundmap = [
        dict(fund_code="F001", explicit_strategy_data_key="STEP_NKY_A", benchmark_series_id="SPX Index",
             benchmark_basis_confirmed="No", nav_dist_definition="Unknown"),
        dict(fund_code="F002", explicit_strategy_data_key="STEP_NKY_B", fund_exposure_multiplier=1.5,
             multiplier_basis="TEST term sheet", multiplier_confirmed="Yes", benchmark_series_id="NKY Index",
             benchmark_basis_confirmed="Yes", nav_dist_definition="Cumulative TR index", nav_dist_def_confirmed="Yes"),
        dict(fund_code="F003"), dict(fund_code="F005"), dict(fund_code="F006"),
    ]
    indexmap = [dict(index_name="TEST STEP NKY Index", strategy_data_key="STEP_NKY_A", approved="Yes"),
                dict(index_name="TEST STEP NKY Index", strategy_data_key="STEP_NKY_B", approved="Yes"),
                dict(index_name="TEST TR Index", strategy_data_key="TR_C_KEY", approved="Yes")]
    dist = [dict(end_investor_bdr_id="B001", to_email="alpha.test@example.invalid", distribution_mode="Combined",
                 distribution_status="Not sent", review_status="Reviewed"),
            dict(end_investor_bdr_id="B002", distribution_mode="Individual"),
            dict(end_investor_bdr_id="B003", to_email="gamma.test@example.invalid", distribution_mode="Individual")]
    cid = "B001|F002|20261030|N20261030|P20261030"
    commentreview = [dict(comment_id=cid, reviewed_text="TEST reviewed text for B001 x F002", reviewer="TEST",
                          review_status=over.get("review_status", "Approved"), approver="TEST Approver",
                          approved_date=EVAL)]
    control = dict(banner="TEST DATA — 架空データ（検証専用・顧客利用不可）", eval_date=EVAL, sel_investor="B001",
                   sel_fund="F002", owner="TEST Owner", thr_confirmed="Yes", max_gap=5,
                   thresholds=over.get("thresholds", {}))
    if "sel_fund" in over:
        control["sel_fund"] = over["sel_fund"]
    if "sel_investor" in over:
        control["sel_investor"] = over["sel_investor"]
    fx = dict(control=control, holidays=[dict(holiday_date=d, holiday_name="TEST holiday", market="JP") for d in HOLIDAYS],
              funds=funds, holdings=holdings, fundmap=fundmap, indexmap=indexmap, nav=nav_series(), market=market_series(),
              undlmap=[dict(ticker_underlying="SPX US Index", spot_series_id="SPX Index", confirmed="Yes", confirmed_by="TEST")],
              positions=positions(), dist=dist, commentreview=commentreview, alertreview=[])
    for k in ("holdings", "positions", "market", "nav"):
        if k + "_extra" in over:
            fx[k] = fx[k] + over[k + "_extra"]
    if "drop_market" in over:
        fx["market"] = [r for r in fx["market"] if not (r["series_id"] == over["drop_market"])]
    return fx
