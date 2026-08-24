"""
Runs the fund calculation engine on top of the synthetic data written by
generate_data.py:
  1. Capital call schedule (investment need + management fee, called pro-rata
     by LP commitment %)
  2. Management fee schedule (2% of committed capital during the investment
     period, stepping down to 2% of remaining invested cost thereafter)
  3. Distribution waterfall (European / whole-fund): return of capital ->
     8% preferred return (compounded quarterly) -> 100% GP catch-up to 20% ->
     80/20 carry split
  4. Quarterly NAV rollforward (Beginning NAV, contributions, fees, realized
     and unrealized gain, distributions, Ending NAV)
  5. Per-LP capital account statements
  6. Fund-level gross IRR/MOIC and net (to-LP) IRR/TVPI/DPI/RVPI
  7. A GP-ledger-vs-fund-admin variance reconciliation exercise

Writes model_output.json plus a handful of flat CSVs used by the Excel
model and the dashboard.
"""
import csv
import json
from collections import defaultdict
from datetime import date

from generate_data import (
    FUND_NAME, FUND_SIZE, N_QUARTERS, INVESTMENT_PERIOD_QUARTERS,
    MGMT_FEE_RATE, PREFERRED_RETURN_RATE, GP_CARRY_RATE, PORTFOLIO_COMPANIES,
)


def read_csv(path):
    with open(path) as f:
        return list(csv.DictReader(f))


def to_num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return x


# ---------------------------------------------------------------- load data
quarters = read_csv("quarters.csv")
for q in quarters:
    q["quarter_index"] = int(q["quarter_index"])
    q["year"] = int(q["year"])
    q["quarter"] = int(q["quarter"])
    q["in_investment_period"] = q["in_investment_period"] == "True"

lps = read_csv("lp_register.csv")
for lp in lps:
    lp["commitment"] = float(lp["commitment"])
    lp["commitment_pct"] = float(lp["commitment_pct"])

marks = read_csv("portfolio_quarterly_marks.csv")
for m in marks:
    m["quarter_index"] = int(m["quarter_index"])
    m["cost_basis"] = float(m["cost_basis"])
    m["fair_value"] = float(m["fair_value"])

exits = read_csv("exit_events.csv")
for e in exits:
    e["quarter_index"] = int(e["quarter_index"])
    e["pct_realized"] = float(e["pct_realized"])
    e["proceeds"] = float(e["proceeds"])
    e["realized_cost_basis"] = float(e["realized_cost_basis"])
    e["realized_gain"] = float(e["realized_gain"])

portfolio_companies = {c[0]: {"name": c[1], "sector": c[2], "invest_quarter": c[3], "cost": c[4]}
                        for c in PORTFOLIO_COMPANIES}

marks_by_q = defaultdict(list)
for m in marks:
    marks_by_q[m["quarter_index"]].append(m)

exits_by_q = defaultdict(list)
for e in exits:
    exits_by_q[e["quarter_index"]].append(e)

quarter_meta = {q["quarter_index"]: q for q in quarters}

# ------------------------------------------------------- capital call sched
investment_need_by_q = defaultdict(float)
for cid, meta in portfolio_companies.items():
    investment_need_by_q[meta["invest_quarter"]] += meta["cost"]

capital_calls = []          # fund-level, one row per quarter
lp_calls = []                # LP x quarter
mgmt_fee_rows = []
cum_invested_cost_by_q = {}  # cost basis on the books entering quarter q (for fee step-down)

running_invested_cost = 0.0
for q in quarters:
    qi = q["quarter_index"]
    # cost basis on the books at the *start* of this quarter = prior quarter's
    # marks total (companies still held), used only post-investment-period
    if qi == 1:
        invested_cost_start = 0.0
    else:
        invested_cost_start = sum(m["cost_basis"] for m in marks_by_q[qi - 1]) if marks_by_q[qi - 1] else \
                               sum(m["cost_basis"] for m in marks_by_q[qi]) if marks_by_q[qi] else 0.0
        # fallback above only triggers if no holdings existed at qi-1 but do at qi (never happens here)

    if q["in_investment_period"]:
        fee = FUND_SIZE * MGMT_FEE_RATE / 4
        fee_basis = FUND_SIZE
        fee_basis_desc = "committed capital"
    else:
        fee_basis = invested_cost_start
        fee = fee_basis * MGMT_FEE_RATE / 4
        fee_basis_desc = "invested cost (step-down)"

    inv_need = investment_need_by_q.get(qi, 0.0)
    total_call = inv_need + fee

    mgmt_fee_rows.append({
        "quarter_index": qi, "label": q["label"], "fee_basis": round(fee_basis, 2),
        "fee_basis_desc": fee_basis_desc, "mgmt_fee": round(fee, 2),
    })
    capital_calls.append({
        "quarter_index": qi, "label": q["label"],
        "investment_call": round(inv_need, 2), "mgmt_fee_call": round(fee, 2),
        "total_call": round(total_call, 2),
    })
    for lp in lps:
        lp_calls.append({
            "quarter_index": qi, "lp_id": lp["lp_id"], "lp_name": lp["lp_name"],
            "investment_call": round(inv_need * lp["commitment_pct"], 2),
            "mgmt_fee_call": round(fee * lp["commitment_pct"], 2),
            "total_call": round(total_call * lp["commitment_pct"], 2),
        })

cum_calls_total = sum(r["total_call"] for r in capital_calls)
unfunded_commitment = FUND_SIZE - cum_calls_total

# ------------------------------------------------------------- fund NAV path
# Ending NAV(q) = sum over companies with a mark this quarter of
#                 (fair_value - proceeds realized this quarter for that company)
ending_nav_by_q = {0: 0.0}
realized_gain_by_q = defaultdict(float)
distributions_total_by_q = defaultdict(float)

for q in quarters:
    qi = q["quarter_index"]
    proceeds_by_company = {e["company_id"]: e["proceeds"] for e in exits_by_q[qi]}
    nav = 0.0
    for m in marks_by_q[qi]:
        proceeds = proceeds_by_company.get(m["company_id"], 0.0)
        nav += m["fair_value"] - proceeds
    ending_nav_by_q[qi] = round(nav, 2)
    realized_gain_by_q[qi] = round(sum(e["realized_gain"] for e in exits_by_q[qi]), 2)
    distributions_total_by_q[qi] = round(sum(e["proceeds"] for e in exits_by_q[qi]), 2)

# ------------------------------------------------------------ waterfall run
# Whole-fund European waterfall, run cumulatively across the fund's life.
# Preferred return of 8%/yr accrues quarterly (2%/qtr) on the LP's
# outstanding (unreturned) capital balance.
def run_waterfall_tiers(distributable, cum_contributions, cum_roc_distributed,
                         cum_pref_accrued, cum_pref_paid, cum_catchup_paid):
    """Applies one quarter's distributable cash through the four tiers given
    the cumulative state entering the quarter. Returns tier amounts; does not
    mutate the inputs."""
    remaining = distributable

    unreturned = cum_contributions - cum_roc_distributed
    roc = min(remaining, max(unreturned, 0.0))
    remaining -= roc
    new_cum_roc = cum_roc_distributed + roc

    pref_owed = cum_pref_accrued - cum_pref_paid
    pref = min(remaining, max(pref_owed, 0.0))
    remaining -= pref
    new_cum_pref_paid = cum_pref_paid + pref

    target_catchup_total = 0.25 * new_cum_pref_paid
    catchup_owed = target_catchup_total - cum_catchup_paid
    catchup = min(remaining, max(catchup_owed, 0.0))
    remaining -= catchup

    carry = remaining
    return {
        "roc": roc, "pref": pref, "catchup": catchup, "carry_total": carry,
        "carry_lp": carry * (1 - GP_CARRY_RATE), "carry_gp": carry * GP_CARRY_RATE,
        "total_to_gp": catchup + carry * GP_CARRY_RATE,
        "total_to_lps": roc + pref + carry * (1 - GP_CARRY_RATE),
    }


cum_contributions = 0.0
cum_roc_distributed = 0.0
cum_pref_accrued = 0.0
cum_pref_paid = 0.0
cum_catchup_paid = 0.0
cum_carry_split_total = 0.0

waterfall_rows = []
nav_rollforward = []
lp_capital_accounts = []
waterfall_state_before = {}   # quarter_index -> cumulative state snapshot, for reconciliation

lp_running_balance = {lp["lp_id"]: 0.0 for lp in lps}

for q in quarters:
    qi = q["quarter_index"]
    call_row = next(c for c in capital_calls if c["quarter_index"] == qi)
    contributions_q = call_row["total_call"]
    fees_q = call_row["mgmt_fee_call"]

    # accrue preferred return on the balance outstanding *before* this
    # quarter's contribution/distribution activity
    outstanding_before = cum_contributions - cum_roc_distributed
    pref_accrual_q = outstanding_before * (PREFERRED_RETURN_RATE / 4)
    cum_pref_accrued += pref_accrual_q

    cum_contributions += contributions_q

    # ---- run the waterfall against this quarter's distributable cash, if any
    distributable = distributions_total_by_q[qi]
    waterfall_state_before[qi] = dict(
        cum_contributions=cum_contributions, cum_roc_distributed=cum_roc_distributed,
        cum_pref_accrued=cum_pref_accrued, cum_pref_paid=cum_pref_paid, cum_catchup_paid=cum_catchup_paid,
    )
    roc_q = pref_q = catchup_q = carry_q = 0.0
    if distributable > 0:
        tiers = run_waterfall_tiers(distributable, cum_contributions, cum_roc_distributed,
                                     cum_pref_accrued, cum_pref_paid, cum_catchup_paid)
        roc_q, pref_q, catchup_q, carry_q = tiers["roc"], tiers["pref"], tiers["catchup"], tiers["carry_total"]
        cum_roc_distributed += roc_q
        cum_pref_paid += pref_q
        cum_catchup_paid += catchup_q
        cum_carry_split_total += carry_q

    carry_split_lp_q = carry_q * (1 - GP_CARRY_RATE)
    carry_split_gp_q = carry_q * GP_CARRY_RATE
    gp_carry_earned_q = catchup_q + carry_split_gp_q
    lp_cash_dist_q = roc_q + pref_q + carry_split_lp_q

    waterfall_rows.append({
        "quarter_index": qi, "label": q["label"],
        "distributable": round(distributable, 2),
        "tier1_return_of_capital": round(roc_q, 2),
        "tier2_preferred_return": round(pref_q, 2),
        "tier3_gp_catchup": round(catchup_q, 2),
        "tier4_carry_split_lp_80": round(carry_split_lp_q, 2),
        "tier4_carry_split_gp_20": round(carry_split_gp_q, 2),
        "total_to_lps": round(lp_cash_dist_q, 2),
        "total_to_gp": round(gp_carry_earned_q, 2),
        "cum_pref_accrued": round(cum_pref_accrued, 2),
        "cum_pref_paid": round(cum_pref_paid, 2),
    })

    # ---------------------------------------------------------- NAV rollforward
    beg_nav = ending_nav_by_q[qi - 1]
    end_nav = ending_nav_by_q[qi]
    realized_g = realized_gain_by_q[qi]
    # unrealized gain is the plug that makes the rollforward tie exactly
    unrealized_g = round(end_nav - beg_nav - contributions_q + fees_q - realized_g + distributable, 2)

    nav_rollforward.append({
        "quarter_index": qi, "label": q["label"],
        "beginning_nav": round(beg_nav, 2),
        "contributions": round(contributions_q, 2),
        "management_fees": round(-fees_q, 2),
        "realized_gain_loss": round(realized_g, 2),
        "unrealized_gain_loss": unrealized_g,
        "distributions": round(-distributable, 2),
        "ending_nav": round(end_nav, 2),
    })

    net_income_q = realized_g + unrealized_g - fees_q

    # ------------------------------------------------------- LP capital accts
    for lp in lps:
        pct = lp["commitment_pct"]
        beg_bal = lp_running_balance[lp["lp_id"]]
        contrib = round(contributions_q * pct, 2)
        dist_roc = round(roc_q * pct, 2)
        dist_pref = round(pref_q * pct, 2)
        dist_carry_lp = round(carry_split_lp_q * pct, 2)
        total_cash_dist = round(dist_roc + dist_pref + dist_carry_lp, 2)
        carry_alloc_to_gp = round(gp_carry_earned_q * pct, 2)
        income_alloc = round(net_income_q * pct - carry_alloc_to_gp, 2)
        end_bal = round(beg_bal + contrib - total_cash_dist + income_alloc, 2)
        lp_running_balance[lp["lp_id"]] = end_bal

        lp_capital_accounts.append({
            "quarter_index": qi, "label": q["label"], "lp_id": lp["lp_id"], "lp_name": lp["lp_name"],
            "beginning_balance": beg_bal, "contributions": contrib,
            "dist_return_of_capital": dist_roc, "dist_preferred_return": dist_pref,
            "dist_carry_split": dist_carry_lp, "total_cash_distributions": total_cash_dist,
            "gain_loss_allocation_net_of_fees": round(net_income_q * pct, 2),
            "carried_interest_allocated_to_gp": -carry_alloc_to_gp,
            "ending_balance": end_bal,
        })

# sanity: sum of LP ending balances should equal fund ending NAV in the final quarter
final_q = N_QUARTERS
lp_sum_final = round(sum(r["ending_balance"] for r in lp_capital_accounts if r["quarter_index"] == final_q), 2)
assert abs(lp_sum_final - ending_nav_by_q[final_q]) < 1.0, \
    f"LP capital accounts ({lp_sum_final:,.2f}) don't tie to fund NAV ({ending_nav_by_q[final_q]:,.2f})"


# --------------------------------------------------------------------- IRR
def xirr(cashflows):
    """cashflows: list of (date, amount). Bisection on NPV(rate)=0."""
    d0 = cashflows[0][0]
    def npv(rate):
        return sum(cf / (1 + rate) ** ((d - d0).days / 365.0) for d, cf in cashflows)
    lo, hi = -0.9999, 10.0
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


q_end_date = {q["quarter_index"]: date.fromisoformat(q["quarter_end"]) for q in quarters}

# Gross (fund/portfolio level, pre-fee, pre-carry)
gross_cfs = []
for cid, meta in portfolio_companies.items():
    gross_cfs.append((q_end_date[meta["invest_quarter"]], -meta["cost"]))
for e in exits:
    gross_cfs.append((q_end_date[e["quarter_index"]], e["proceeds"]))
remaining_portfolio_fv = ending_nav_by_q[final_q]
gross_cfs.append((q_end_date[final_q], remaining_portfolio_fv))
gross_cfs.sort(key=lambda x: x[0])
gross_irr = xirr(gross_cfs)
total_invested = sum(meta["cost"] for meta in portfolio_companies.values())
total_realized_proceeds = sum(e["proceeds"] for e in exits)
gross_moic = (total_realized_proceeds + remaining_portfolio_fv) / total_invested

# Net (to LPs, post-fee, post-carry) — identical for every LP given pari-passu timing
net_cfs = []
for c in capital_calls:
    net_cfs.append((q_end_date[c["quarter_index"]], -c["total_call"]))
for w in waterfall_rows:
    if w["total_to_lps"] > 0:
        net_cfs.append((q_end_date[w["quarter_index"]], w["total_to_lps"]))
net_cfs.append((q_end_date[final_q], ending_nav_by_q[final_q]))
net_cfs.sort(key=lambda x: x[0])
net_irr = xirr(net_cfs)

cum_lp_contributions = cum_calls_total
cum_lp_distributions = sum(w["total_to_lps"] for w in waterfall_rows)
net_tvpi = (cum_lp_distributions + ending_nav_by_q[final_q]) / cum_lp_contributions
dpi = cum_lp_distributions / cum_lp_contributions
rvpi = ending_nav_by_q[final_q] / cum_lp_contributions

performance_summary = {
    "as_of_quarter": quarter_meta[final_q]["label"],
    "as_of_date": quarter_meta[final_q]["quarter_end"],
    "committed_capital": FUND_SIZE,
    "called_capital": round(cum_calls_total, 2),
    "unfunded_commitment": round(unfunded_commitment, 2),
    "pct_called": round(cum_calls_total / FUND_SIZE, 4),
    "total_invested_cost": round(total_invested, 2),
    "total_realized_proceeds": round(total_realized_proceeds, 2),
    "remaining_portfolio_fv": round(remaining_portfolio_fv, 2),
    "ending_nav": round(ending_nav_by_q[final_q], 2),
    "cum_lp_distributions": round(cum_lp_distributions, 2),
    "cum_gp_carry_earned": round(sum(w["total_to_gp"] for w in waterfall_rows), 2),
    "gross_irr": gross_irr, "gross_moic": round(gross_moic, 4),
    "net_irr": net_irr, "net_tvpi": round(net_tvpi, 4),
    "dpi": round(dpi, 4), "rvpi": round(rvpi, 4),
}

# --------------------------------------------------- reconciliation exercise
# GP ledger = the figures computed above (source of truth). Fund admin
# introduces three realistic, independently-explainable breaks.
lp_by_id = {lp["lp_id"]: lp for lp in lps}

# Break 1: settlement-timing lag on two LPs' Q9 capital call wires
timing_lps = ["LP-05", "LP-09"]      # Elmwood Family Office, Wrenfield Family Office
q9_call = next(c for c in capital_calls if c["quarter_index"] == 9)
timing_amt = round(sum(q9_call["total_call"] * lp_by_id[l]["commitment_pct"] for l in timing_lps), 2)

# Break 2: fund admin didn't step the mgmt fee basis down post-investment-period;
# kept billing 2% of committed capital for Q13-Q16 (caught and corrected starting Q17)
fee_break_quarters = [13, 14, 15, 16]
fee_breaks = []
admin_fee_error_total = 0.0
for qi in fee_break_quarters:
    correct = next(r["mgmt_fee"] for r in mgmt_fee_rows if r["quarter_index"] == qi)
    admin_amt = round(FUND_SIZE * MGMT_FEE_RATE / 4, 2)
    variance = round(admin_amt - correct, 2)
    admin_fee_error_total += variance
    fee_breaks.append({"quarter_index": qi, "label": quarter_meta[qi]["label"],
                        "gp_ledger_fee": correct, "fund_admin_fee": admin_amt, "variance": variance})

# Break 3: Harbor Dental partial recap (Q18) — fund admin's system flagged the
# recap as a return of capital event and remitted it outside the tiered
# waterfall, before running the tiered waterfall on the quarter's other two
# exits. Counterfactual: run the admin's (wrong) sequencing against the same
# cumulative state entering Q18 and compare GP carry earned.
q18_state = waterfall_state_before[18]
hd_recap = next(e for e in exits if e["company_id"] == "PC-04")
other_q18_distributable = distributions_total_by_q[18] - hd_recap["proceeds"]

admin_roc_from_recap = hd_recap["proceeds"]  # admin pays 100% of the recap as ROC, outside the tiers
admin_state_roc = q18_state["cum_roc_distributed"] + admin_roc_from_recap
admin_tiers = run_waterfall_tiers(other_q18_distributable, q18_state["cum_contributions"], admin_state_roc,
                                   q18_state["cum_pref_accrued"], q18_state["cum_pref_paid"], q18_state["cum_catchup_paid"])
admin_q18_gp_carry = admin_tiers["total_to_gp"]

gp_ledger_q18_carry = next(w for w in waterfall_rows if w["quarter_index"] == 18)["total_to_gp"]
carry_break_variance = round(admin_q18_gp_carry - gp_ledger_q18_carry, 2)

# Break 4: unreconciled wire fee netted against a Q15 LP distribution
wire_fee_break = -375.00

reconciliation = {
    "as_of": quarter_meta[final_q]["label"],
    "breaks": [
        {
            "break_id": "R-1", "category": "Timing difference",
            "description": ("Q9 (Q1 2024) capital call: GP ledger records LP contributions on trade "
                             "date (2024-01-15). Fund administrator books on wire settlement date; two "
                             "LPs' wires cleared after quarter-end and landed in the Q10 admin ledger."),
            "affected_lps": timing_lps,
            "gp_ledger_amount": q9_call["total_call"],
            "fund_admin_amount": round(q9_call["total_call"] - timing_amt, 2),
            "variance": round(-timing_amt, 2),
            "resolution": ("No adjusting entry required — timing-only difference. Confirmed both LP "
                            "wires received in full; Q10 fund admin capital call report shows the "
                            f"offsetting +${timing_amt:,.2f}, netting to zero across the two quarters."),
            "status": "Resolved — timing only, nets to zero by Q10",
        },
        {
            "break_id": "R-2", "category": "Management fee basis error",
            "description": ("Per the LPA, management fee steps down from 2% of committed capital to 2% "
                             "of remaining invested cost after the investment period ends (Q4 2024). "
                             "Fund administrator's fee calculation continued billing 2% of committed "
                             "capital for Q13-Q16 before the step-down was applied starting Q17."),
            "quarterly_detail": fee_breaks,
            "gp_ledger_amount": round(sum(b["gp_ledger_fee"] for b in fee_breaks), 2),
            "fund_admin_amount": round(sum(b["fund_admin_fee"] for b in fee_breaks), 2),
            "variance": round(admin_fee_error_total, 2),
            "resolution": (f"Adjusting entry required: credit LP capital accounts ${admin_fee_error_total:,.2f} "
                            "(pro rata by commitment %) for fees over-charged Q13-Q16; fund administrator "
                            "instructed to apply the step-down formula (2% x prior-quarter remaining "
                            "invested cost) going forward."),
            "status": "Open — adjusting entry pending LP notice",
        },
        {
            "break_id": "R-3", "category": "Distribution character misclassification",
            "description": (f"Harbor Dental Partners Q2 2026 dividend recap (${hd_recap['proceeds']:,.2f}, 60% "
                             "partial realization), paid alongside the BluePeak Software and Fairmont Health "
                             "full exits in the same quarter: fund administrator's system flagged the "
                             "recap as a recapitalization rather than a standard realization and remitted "
                             "it 100% as return of capital outside the tiered waterfall, before running "
                             "the preferred return / catch-up / carry tiers on the quarter's other two "
                             "exits. The LPA requires every cash distribution — recap or sale — to pass "
                             "through the same waterfall in the order it is received."),
            "gp_ledger_amount": gp_ledger_q18_carry,
            "fund_admin_amount": round(admin_q18_gp_carry, 2),
            "variance": carry_break_variance,
            "resolution": ("Adjusting entry required: re-sequence the Q2 2026 distribution through the "
                            "waterfall in receipt order (recap included), and remit the GP the additional "
                            f"${abs(carry_break_variance):,.2f} in carried interest the misclassification under-paid."),
            "status": "Open — reclassification and GP remittance pending",
        },
        {
            "break_id": "R-4", "category": "Unreconciled cash break",
            "description": ("Q15 (Q3 2025) LP distribution: fund administrator's bank netted a $375.00 "
                             "outgoing wire fee against one LP's distribution payment; the fee was not "
                             "recorded in the GP's internal ledger."),
            "gp_ledger_amount": 0.00,
            "fund_admin_amount": wire_fee_break,
            "variance": wire_fee_break,
            "resolution": "Booked to fund operating expenses; immaterial, no LP notice required.",
            "status": "Resolved — booked to fund expenses",
        },
    ],
}
reconciliation["total_open_variance_gross"] = round(
    sum(abs(b["variance"]) for b in reconciliation["breaks"] if b["status"].startswith("Open")), 2)
reconciliation["open_items_detail"] = [
    {"break_id": b["break_id"], "category": b["category"], "variance": b["variance"]}
    for b in reconciliation["breaks"] if b["status"].startswith("Open")
]

# ------------------------------------------------------------------- output
output = {
    "fund_name": FUND_NAME, "fund_size": FUND_SIZE,
    "n_quarters": N_QUARTERS, "investment_period_quarters": INVESTMENT_PERIOD_QUARTERS,
    "mgmt_fee_rate": MGMT_FEE_RATE, "preferred_return_rate": PREFERRED_RETURN_RATE,
    "gp_carry_rate": GP_CARRY_RATE,
    "quarters": quarters, "lps": lps,
    "capital_calls": capital_calls, "mgmt_fee_schedule": mgmt_fee_rows,
    "waterfall": waterfall_rows, "nav_rollforward": nav_rollforward,
    "performance_summary": performance_summary,
    "reconciliation": reconciliation,
    "exit_events": exits,
}
with open("model_output.json", "w") as f:
    json.dump(output, f, indent=2, default=str)

with open("lp_capital_accounts.csv", "w", newline="") as f:
    fieldnames = list(lp_capital_accounts[0].keys())
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(lp_capital_accounts)

with open("capital_calls.csv", "w", newline="") as f:
    fieldnames = list(capital_calls[0].keys())
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(capital_calls)

with open("lp_capital_calls.csv", "w", newline="") as f:
    fieldnames = list(lp_calls[0].keys())
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(lp_calls)

with open("nav_rollforward.csv", "w", newline="") as f:
    fieldnames = list(nav_rollforward[0].keys())
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(nav_rollforward)

with open("waterfall.csv", "w", newline="") as f:
    fieldnames = list(waterfall_rows[0].keys())
    w = csv.DictWriter(f, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(waterfall_rows)

print(f"{FUND_NAME}  —  as of {performance_summary['as_of_quarter']}")
print(f"Called: ${performance_summary['called_capital']:,.0f} ({performance_summary['pct_called']*100:.1f}% of commitments)")
print(f"Ending NAV: ${performance_summary['ending_nav']:,.0f}")
print(f"Cumulative LP distributions: ${performance_summary['cum_lp_distributions']:,.0f}")
print(f"Cumulative GP carry earned: ${performance_summary['cum_gp_carry_earned']:,.0f}")
print(f"Gross IRR: {gross_irr*100:.2f}%   Gross MOIC: {gross_moic:.2f}x")
print(f"Net IRR (to LPs): {net_irr*100:.2f}%   Net TVPI: {net_tvpi:.2f}x   DPI: {dpi:.2f}x   RVPI: {rvpi:.2f}x")
print(f"LP capital accounts tie to fund NAV: {lp_sum_final:,.2f} vs {ending_nav_by_q[final_q]:,.2f}")
print(f"Reconciliation: {len(reconciliation['breaks'])} breaks identified across {performance_summary['as_of_quarter']}'s tie-out")
for item in reconciliation["open_items_detail"]:
    direction = "owed to GP" if item["variance"] < 0 else "owed back to LPs"
    print(f"  {item['break_id']} {item['category']}: ${abs(item['variance']):,.2f} {direction}")
