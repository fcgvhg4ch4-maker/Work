# QIS Fund Monitoring & Sales Update Workbook

This folder holds an Excel 365 workbook (17 sheets, no VBA) for the regional-bank sales team. It brings together
live funds, investor holdings, strategy mapping, NAV, market data and STEP positions. It also contains the scripts
that generate the workbook and test it.

| Path | Contents |
|---|---|
| `dist/QIS_Fund_Monitoring.xlsx` | **Production template.** Every input table is empty; no fund, investor or market data is included |
| `dist/QIS_Fund_Monitoring_TESTDATA.xlsx` | The same workbook filled with **synthetic TEST data**, for verification only |
| `build_workbook.py` | Generates the workbook: sheets, tables, formulas, named ranges, formatting and the chart |
| `xlhelpers.py` | Helpers for writing tables, expanding formulas and styling |
| `fixtures.py` | Synthetic test scenarios (made-up data) |
| `recalc.py` | Recalculates a workbook with headless LibreOffice |
| `verify.py` | Runs the 16 required tests and writes `docs/test_results.md` |
| `docs/operation_guide.md` | Deliverable 2: operating guide for sales staff (Japanese) |
| `docs/verification_report.md` | Deliverable 3: verification report (Japanese) |

```sh
python3 build_workbook.py              # -> dist/QIS_Fund_Monitoring.xlsx
python3 build_workbook.py --fixtures   # -> dist/QIS_Fund_Monitoring_TESTDATA.xlsx
python3 verify.py                      # recalculates in LibreOffice, then runs the 16 tests
```
