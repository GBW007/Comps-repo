# BFSI Comps & Quarterly Results Tracker (Python)

A small, reusable Python tool for Indian bank research:
1. **Comps:** ranks 20 listed private banks on P/E, P/B, ROE and profit growth, with medians by size bucket (large > Rs1 lakh cr, mid Rs35,000 cr–1 lakh cr, small < Rs35,000 cr).
2. **P/E vs growth:** OLS fit after removing outliers by the IQR rule. Result on 30-Sep-2026 data: R² ≈ 0 – one-quarter growth does not explain bank P/Es, which is why banks are valued on P/B vs ROE.
3. **P/B vs ROE:** the standard bank valuation lens. Switches on automatically once 5+ banks have book value and ROE in `data/peers.csv` (3 filled today – see Exercise 1).
4. **Karur Vysya Bank tracker:** 13 quarters of NII, other income, PBT, profit and NPAs, with YoY/QoQ growth and charts.
5. **Results-day sheet:** model-implied Q2 FY27 estimates; type the actuals on 15-Oct-2026 and variance and Beat/Miss update by formula.

## Run
```bash
pip install -r requirements.txt
python bfsi_toolkit.py
```
Outputs go to `output/`: `BFSI_Comps_Tracker.xlsx`, `summary.md` and `charts/*.png`.

## Inputs (edit the CSVs, then re-run)
| File | Columns |
|---|---|
| `data/peers.csv` | bank, cmp, pe_ttm, mcap_cr, div_yield_pct, pat_q1fy27_cr, pat_yoy_pct, int_earned_q1fy27_cr, int_earned_yoy_pct, bvps, roe_pct |
| `data/kvb_quarterly.csv` | quarter, interest_earned, interest_expended, opex_and_provisions, other_income, net_profit, gnpa_pct, nnpa_pct, nii, pbt |
| `data/kvb_estimates.csv` | metric, fy27e_model (from the KVB initiation model) |

Data: Screener.in industry table and company pages, 30-Sep-2026. Screener's "sales" for banks is interest earned.


## Limitations
Twenty banks is a small sample; P/E uses trailing earnings; the Q2 estimate simply spreads the model's FY27E remainder over three quarters.
