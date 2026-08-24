"""
Generates the synthetic source data for a $50M PE fund:
  - LP register (15 LPs, commitments summing to $50,000,000)
  - Portfolio company register (10 investments made during the 3-year
    investment period)
  - Quarterly fair-value marks and exit events for each portfolio company,
    built from a hand-set drift per company archetype (winner / steady /
    laggard / write-down) plus small seeded noise

Writes:
  lp_register.csv
  portfolio_companies.csv
  portfolio_quarterly_marks.csv
  exit_events.csv
  quarters.csv

This is a synthetic fund. The return drivers (which company wins, which is
written down, when exits happen) are authored in this file, not observed.
What downstream files (calculations.py, the Excel model, the dashboard)
demonstrate is the calculation machinery — capital calls, the waterfall,
NAV rollforward, capital accounts, IRR/MOIC, and reconciliation — applied
to those cash flows.
"""
import csv
import numpy as np
from datetime import date

RNG_SEED = 42
rng = np.random.default_rng(RNG_SEED)

FUND_NAME = "Meridian Growth Partners III, L.P."
FUND_SIZE = 50_000_000
VINTAGE_YEAR = 2022
N_QUARTERS = 18            # Q1 2022 through Q2 2026
INVESTMENT_PERIOD_QUARTERS = 12   # Q1 2022 through Q4 2024
MGMT_FEE_RATE = 0.02        # 2% annual
PREFERRED_RETURN_RATE = 0.08  # 8% annual, compounded quarterly
GP_CARRY_RATE = 0.20        # 20% carried interest
CATCHUP_GP_SHARE = 1.00     # 100% catch-up to GP until caught up to 20% of profit


def build_quarters():
    quarters = []
    y, q = VINTAGE_YEAR, 1
    end_month_day = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}
    for i in range(1, N_QUARTERS + 1):
        m, d = end_month_day[q]
        quarters.append({
            "quarter_index": i,
            "year": y,
            "quarter": q,
            "label": f"Q{q} {y}",
            "quarter_end": date(y, m, d).isoformat(),
            "in_investment_period": i <= INVESTMENT_PERIOD_QUARTERS,
        })
        q += 1
        if q == 5:
            q = 1
            y += 1
    return quarters


LPS = [
    ("LP-01", "Northbridge State Pension Fund", "Pension Fund", 8_000_000),
    ("LP-02", "Meridian University Endowment", "Endowment", 6_000_000),
    ("LP-03", "Harborview Insurance Co.", "Insurance Company", 5_500_000),
    ("LP-04", "Cascade Partners Fund of Funds", "Fund of Funds", 5_000_000),
    ("LP-05", "Elmwood Family Office", "Family Office", 4_000_000),
    ("LP-06", "Tanager Sovereign Partners", "Sovereign Wealth Fund", 4_000_000),
    ("LP-07", "Birchwood College Endowment", "Endowment", 3_000_000),
    ("LP-08", "Kestrel Insurance Group", "Insurance Company", 2_500_000),
    ("LP-09", "Wrenfield Family Office", "Family Office", 2_000_000),
    ("LP-10", "Alder Street Foundation", "Foundation", 2_000_000),
    ("LP-11", "Pinecrest Pension Trust", "Pension Fund", 1_800_000),
    ("LP-12", "Ivywood Capital Partners (FoF)", "Fund of Funds", 1_500_000),
    ("LP-13", "Grayling Family Office", "Family Office", 1_500_000),
    ("LP-14", "Redstone Family Office", "Family Office", 1_600_000),
    ("LP-15", "Sable Foundation", "Foundation", 1_600_000),
]


def build_lp_register():
    total = sum(r[3] for r in LPS)
    assert total == FUND_SIZE, f"LP commitments sum to {total:,}, expected {FUND_SIZE:,}"
    rows = []
    for lp_id, name, lp_type, commitment in LPS:
        rows.append({
            "lp_id": lp_id,
            "lp_name": name,
            "lp_type": lp_type,
            "commitment": commitment,
            "commitment_pct": commitment / FUND_SIZE,
        })
    return rows


# (company_id, name, sector, invest_quarter, cost_basis, archetype)
PORTFOLIO_COMPANIES = [
    ("PC-01", "Vantage Robotics",        "Industrial Tech",     1,  6_000_000, "winner"),
    ("PC-02", "Coastal Foods Group",     "Consumer",            2,  5_500_000, "steady"),
    ("PC-03", "Nimbus Data Systems",     "Software",            3,  4_500_000, "writedown"),
    ("PC-04", "Harbor Dental Partners",  "Healthcare Services",  4,  5_000_000, "recap"),
    ("PC-05", "Trailhead Outdoor Co.",   "Consumer",            5,  4_000_000, "laggard"),
    ("PC-06", "Summit Logistics",        "Transportation",      6,  5_000_000, "winner"),
    ("PC-07", "BluePeak Software",       "Software",            7,  4_500_000, "steady"),
    ("PC-08", "Ridgeline Materials",     "Industrials",         8,  3_500_000, "laggard"),
    ("PC-09", "Fairmont Health Services","Healthcare Services",  9,  4_000_000, "steady"),
    ("PC-10", "Alpine Consumer Brands",  "Consumer",           10,  3_500_000, "steady"),
]

# archetype -> (quarterly drift, noise sigma, exit_quarter or None, exit_multiple_of_last_mark)
ARCHETYPE_PARAMS = {
    "winner":    dict(drift=0.105, sigma=0.015),
    "steady":    dict(drift=0.054, sigma=0.012),
    "laggard":   dict(drift=0.014, sigma=0.010),
    "writedown": dict(drift=-0.045, sigma=0.015, floor_quarter=11),
    "recap":     dict(drift=0.054, sigma=0.012),
}

# Explicit exit / partial-realization events: (company_id, quarter_index, event_type, pct_of_position)
# event_type: "full_exit" | "partial_recap"
EXIT_EVENTS = [
    ("PC-01", 13, "full_exit", 1.00),      # Vantage Robotics — strong exit, Q1 2025
    ("PC-02", 15, "full_exit", 1.00),      # Coastal Foods Group — Q3 2025
    ("PC-06", 17, "full_exit", 1.00),      # Summit Logistics — Q1 2026
    ("PC-07", 18, "full_exit", 1.00),      # BluePeak Software — Q2 2026
    ("PC-09", 18, "full_exit", 1.00),      # Fairmont Health Services — Q2 2026
    ("PC-04", 18, "partial_recap", 0.60),  # Harbor Dental Partners — dividend recap, Q2 2026
]


def simulate_marks():
    """Quarterly fair value for every company from its invest quarter through
    N_QUARTERS (or its exit quarter). Returns list of dict rows and a dict of
    exit records with realized proceeds."""
    exit_map = {(cid, q): (etype, pct) for cid, q, etype, pct in EXIT_EVENTS}
    marks = []
    exits_out = []

    for company_id, name, sector, invest_q, cost, archetype in PORTFOLIO_COMPANIES:
        params = ARCHETYPE_PARAMS[archetype]
        drift = params["drift"]
        sigma = params["sigma"]
        floor_q = params.get("floor_quarter")

        fv = cost
        remaining_cost = cost  # shrinks on partial realizations
        for q in range(invest_q, N_QUARTERS + 1):
            if q == invest_q:
                fv = cost
            else:
                noise = rng.normal(0, sigma)
                step = drift + noise
                if floor_q and q >= floor_q:
                    step = rng.normal(0.0, 0.006)  # stabilizes after the write-down
                fv = max(fv * (1 + step), 0.05 * cost)

            marks.append({
                "company_id": company_id,
                "company_name": name,
                "sector": sector,
                "quarter_index": q,
                "cost_basis": round(remaining_cost, 2),
                "fair_value": round(fv, 2),
            })

            key = (company_id, q)
            if key in exit_map:
                etype, pct = exit_map[key]
                proceeds = round(fv * pct, 2)
                realized_cost = round(remaining_cost * pct, 2)
                exits_out.append({
                    "company_id": company_id,
                    "company_name": name,
                    "quarter_index": q,
                    "event_type": etype,
                    "pct_realized": pct,
                    "proceeds": proceeds,
                    "realized_cost_basis": realized_cost,
                    "realized_gain": round(proceeds - realized_cost, 2),
                })
                remaining_cost = round(remaining_cost - realized_cost, 2)
                fv = round(fv - proceeds, 2)
                if etype == "full_exit":
                    break  # no further marks after a full exit

    return marks, exits_out


def write_csv(path, rows, fieldnames):
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow(r)


if __name__ == "__main__":
    quarters = build_quarters()
    lps = build_lp_register()
    marks, exits = simulate_marks()

    write_csv("quarters.csv", quarters,
               ["quarter_index", "year", "quarter", "label", "quarter_end", "in_investment_period"])
    write_csv("lp_register.csv", lps,
               ["lp_id", "lp_name", "lp_type", "commitment", "commitment_pct"])
    write_csv("portfolio_companies.csv",
               [{"company_id": c[0], "company_name": c[1], "sector": c[2],
                 "invest_quarter": c[3], "cost_basis": c[4], "archetype": c[5]} for c in PORTFOLIO_COMPANIES],
               ["company_id", "company_name", "sector", "invest_quarter", "cost_basis", "archetype"])
    write_csv("portfolio_quarterly_marks.csv", marks,
               ["company_id", "company_name", "sector", "quarter_index", "cost_basis", "fair_value"])
    write_csv("exit_events.csv", exits,
               ["company_id", "company_name", "quarter_index", "event_type", "pct_realized",
                "proceeds", "realized_cost_basis", "realized_gain"])

    print(f"Fund: {FUND_NAME}  |  Size: ${FUND_SIZE:,}  |  {len(lps)} LPs  |  {len(PORTFOLIO_COMPANIES)} portfolio companies")
    print(f"Quarters modeled: {N_QUARTERS} ({quarters[0]['label']} - {quarters[-1]['label']})")
    print(f"Total invested at cost: ${sum(c[4] for c in PORTFOLIO_COMPANIES):,}")
    print(f"Exit events: {len(exits)}")
    for e in exits:
        print(f"  {e['company_name']:<24} Q{e['quarter_index']:>2}  {e['event_type']:<14} proceeds=${e['proceeds']:,.0f}  gain=${e['realized_gain']:,.0f}")
