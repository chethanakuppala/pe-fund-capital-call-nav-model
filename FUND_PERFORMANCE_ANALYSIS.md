# Fund Performance Analysis — Meridian Growth Partners III, L.P.

**Fund:** $50,000,000 committed capital, 15 LPs, 3-year investment period (Q1 2022 – Q4 2024)
**As of:** Q2 2026 (18 quarters modeled)
**Prepared from:** `calculations.py` against `generate_data.py`'s synthetic LP, portfolio, and marks data

> Synthetic fund — see the note in `README.md`. The waterfall mechanics, the
> NAV rollforward logic, and the reconciliation findings are the substance
> here; the underlying valuations and exit outcomes come from a documented
> generator in `generate_data.py`, not a real fund.

---

## 1. Executive summary

Meridian Growth Partners III called $49.7M of its $50.0M committed capital
(99.4%), deployed $45.5M at cost across 10 portfolio companies, and has
returned $62.1M to LPs against an ending NAV of $21.7M — a net TVPI of **1.69x**
(1.25x DPI + 0.44x RVPI) and a net IRR to LPs of **20.9%**. On the portfolio
itself, pre-fee and pre-carry, gross performance is **25.6% IRR / 1.91x MOIC**.

Three results are decision-useful:

1. **The fund crossed into carried interest territory only in the final
   quarter modeled.** Through Q1 2026, every dollar distributed was still
   inside the return-of-capital tier or working through the preferred return.
   The GP has earned $3.1M in carry to date — all of it booked in Q2 2026,
   once cumulative distributions cleared the $49.7M return-of-capital
   threshold and the $9.6M accrued preferred return.
2. **A four-break reconciliation between the GP's internal ledger and the
   fund administrator's records surfaced $478,903 in combined open variance**,
   of which $150,000 is owed back to LPs (a management fee basis error) and
   $328,903 is owed to the GP (a distribution misclassification that bypassed
   the waterfall entirely for one exit).
3. **The portfolio is bifurcated.** Two of ten companies (Vantage Robotics,
   Summit Logistics) drove the bulk of realized gains; one (Nimbus Data
   Systems) was written down 45% from cost and never recovered. The blended
   result is good, but it rests on a small number of winners — not a
   uniformly strong book.

---

## 2. Fund terms and capital structure

| Term | Value |
|---|---|
| Committed capital | $50,000,000 |
| Number of LPs | 15 |
| Investment period | 12 quarters (Q1 2022 – Q4 2024) |
| Management fee | 2.0% annual on committed capital during the investment period; steps down to 2.0% annual on remaining invested cost thereafter |
| Preferred return | 8.0% annual, compounded quarterly, on each LP's outstanding (unreturned) capital balance |
| GP catch-up | 100% to GP until GP has received 20% of (preferred return + catch-up) paid to date |
| Carried interest | 80% LP / 20% GP on all distributions above the catch-up |
| Waterfall structure | Whole-fund (European) — run cumulatively across the fund's life, not deal-by-deal |

Fifteen LPs committed capital ranging from $1.2M to $8.0M, led by Northbridge
State Pension Fund (16.0% of the fund) and Meridian University Endowment
(12.0%). The full register is in `lp_register.csv` and the **LP Register** tab
of the Excel model. Because every LP is called and distributed strictly
pro-rata to its commitment percentage, net IRR and TVPI are **identical for
every LP** — confirmed directly: Northbridge (16.0% of the fund) shows a
1.6865x MOIC on its own $7.95M of contributions, matching the fund-level net
TVPI of 1.6865x to four decimal places.

---

## 3. Capital calls and management fees

Total capital called through Q2 2026: **$49,690,000** (99.4% of commitments),
leaving **$310,000** unfunded. That splits into $45,500,000 for portfolio
investments and **$4,190,000** in cumulative management fees.

The fee mechanism has two regimes:

- **Investment period (Q1 2022 – Q4 2024):** 2.0% annual on the full $50.0M
  committed capital, called quarterly — a flat $250,000 per quarter regardless
  of how much capital had actually been deployed yet.
- **Post-investment-period (Q1 2025 onward):** 2.0% annual on the *prior
  quarter's* remaining invested cost — a shrinking base as portfolio companies
  are exited. The fee stepped from $250,000/quarter (Q4 2024, still on
  committed capital) to $227,500/quarter (Q1–Q2 2025, on $45.5M invested cost)
  and down further as exits reduced the base, reaching $170,000/quarter by
  Q1–Q2 2026.

This step-down is exactly the mechanism the fund administrator got wrong in
reconciliation break R-2 (§6).

---

## 4. Portfolio and NAV rollforward

Ten investments were made across the first ten quarters of the investment
period, totaling $45.5M at cost. Six realization events occurred during the
harvest period:

| Company | Exit quarter | Type | Proceeds | Realized gain |
|---|---|---|---:|---:|
| Vantage Robotics | Q1 2025 | Full exit | $19,411,887 | $13,411,887 |
| Coastal Foods Group | Q3 2025 | Full exit | $10,993,566 | $5,493,566 |
| Summit Logistics | Q1 2026 | Full exit | $14,803,879 | $9,803,879 |
| Harbor Dental Partners | Q2 2026 | Partial recap (60%) | $6,125,185 | $3,125,185 |
| BluePeak Software | Q2 2026 | Full exit | $7,670,711 | $3,170,711 |
| Fairmont Health Services | Q2 2026 | Full exit | $6,199,333 | $2,199,333 |

Four companies remain in the portfolio at Q2 2026, marked at:

| Company | Cost | Q2 2026 fair value | Unrealized gain |
|---|---:|---:|---:|
| Harbor Dental Partners (remaining 40%) | $2,000,000 | $4,083,457 | $2,083,457 |
| Alpine Consumer Brands | $3,500,000 | $5,473,487 | $1,973,487 |
| Trailhead Outdoor Co. | $4,000,000 | $4,763,542 | $763,542 |
| Ridgeline Materials | $3,500,000 | $3,931,043 | $431,043 |
| Nimbus Data Systems | $4,500,000 | $3,449,249 | ($1,050,751) |

Nimbus Data Systems is the fund's one clear disappointment — written down
roughly 23% from cost in the year following investment and never recovering,
consistent with the NAV rollforward's negative unrealized gain/loss line in
quarters 9–11.

**NAV rollforward mechanics.** Each quarter's Ending NAV is computed
independently as the portfolio's fair value net of that quarter's
distributions — not accumulated from the prior quarter's components. Beginning
NAV, contributions, management fees, realized gain, and distributions are all
directly observable each quarter; unrealized gain/loss is solved as the plug
that makes the rollforward tie exactly to that independently-computed Ending
NAV. This mirrors how a fund administrator actually closes a NAV rollforward:
the ending mark is the anchor, and unrealized gain/loss absorbs whatever the
other five lines don't explain.

Fund NAV peaked at $69.5M at the end of Q4 2024 — the close of the investment
period, with all ten companies deployed and marked up but nothing yet
realized — and has come down to $21.7M by Q2 2026 as realizations (which
convert marked value to cash paid out) outpaced new unrealized appreciation.

---

## 5. The distribution waterfall

The waterfall is whole-fund and cumulative — every dollar distributed, in
every quarter, is tested against the same four running balances (cumulative
contributions, cumulative return of capital paid, cumulative preferred return
accrued and paid, cumulative GP catch-up paid) rather than being evaluated
deal-by-deal.

**Through Q1 2026, every distribution stayed inside Tier 1.** Cumulative
distributions had not yet caught up to cumulative capital called, so 100% of
every payout — Vantage Robotics, Coastal Foods Group, and Summit Logistics —
went to LPs as return of capital.

**Q2 2026 crossed three tier boundaries in a single quarter.** With
$19,995,229 distributable (the Harbor Dental recap plus the BluePeak and
Fairmont Health exits) and only $4,480,669 of unreturned capital remaining,
the waterfall ran through all four tiers:

| Tier | Amount | Running total after |
|---|---:|---|
| 1 — Return of capital | $4,480,669 | Capital fully returned |
| 2 — Preferred return (8%, compounded quarterly since inception) | $9,617,220 | Preferred return fully paid — accrued exactly $9,617,220 since Q1 2022 |
| 3 — GP catch-up (100% to GP, to 20% of tier 2 + tier 3) | $2,404,305 | GP caught up to 20% of profit distributed so far |
| 4 — Carry split (80/20) | $3,493,036 total | $2,794,429 to LPs, $698,607 to GP |
| **Total** | **$19,995,229** | **$16,892,317 to LPs / $3,102,912 to GP** |

The GP's cumulative carried interest earned to date is **$3,102,912** — the
sum of the Q2 2026 catch-up ($2,404,305) and the GP's 20% share of the Q2 2026
carry split ($698,607). No carry was earned in any earlier quarter.

---

## 6. Fund performance: IRR, MOIC, TVPI, DPI, RVPI

| Metric | Gross (portfolio) | Net (to LPs) |
|---|---:|---:|
| IRR | 25.6% | 20.9% |
| MOIC / TVPI | 1.91x | 1.69x |
| DPI | — | 1.25x |
| RVPI | — | 0.44x |

The 470 bp gap between gross and net IRR, and the 0.22x gap between gross
MOIC and net TVPI, is the combined cost of the 2% management fee drag over
4.5 years and the carried interest paid in the final quarter. Because DPI
(1.25x) already exceeds 1.0x, LPs have received back more cash than they put
in, with 0.44x of additional value still marked in the portfolio — a fund
that has "returned the fund" on a cash basis and is still creating value.

---

## 7. Variance reconciliation: GP ledger vs. fund administrator

A four-item reconciliation between the GP's internal ledger and the fund
administrator's records, run at the Q2 2026 tie-out, surfaced two resolved,
timing-or-immaterial breaks and two open breaks requiring adjusting entries.

| Break | Category | Variance | Status |
|---|---|---:|---|
| R-1 | Timing difference — Q1 2024 capital call settlement lag on two LPs' wires | $510,000 (nets to $0 by the following quarter) | Resolved |
| R-2 | Management fee basis error — Q1 2025–Q4 2025 billed on committed capital instead of the stepped-down invested-cost basis | $150,000 owed back to LPs | **Open** |
| R-3 | Distribution character misclassification — the Harbor Dental recap booked 100% as return of capital, bypassing the waterfall | $328,903 owed to the GP | **Open** |
| R-4 | Unreconciled cash break — a $375 outgoing wire fee netted by the administrator's bank, not on the GP ledger | $375 (booked to fund expenses) | Resolved |

**R-2 in detail.** The LPA specifies that the management fee steps down from
2% of committed capital to 2% of remaining invested cost once the investment
period ends. The fund administrator's fee run continued billing the
committed-capital basis for four quarters (Q1 2025 – Q4 2025) before the
step-down was correctly applied starting Q1 2026 — overcharging the fund
$22,500–$52,500 per quarter as the invested-cost basis shrank further below
committed capital with each exit. The fix is a credit to LP capital accounts,
pro rata by commitment percentage.

**R-3 in detail, and why it matters more than R-2's dollar amount.** The
Harbor Dental Partners recap was flagged by the administrator's system as a
recapitalization rather than a standard sale, and was booked entirely as
return of capital — outside the waterfall altogether — rather than being run
through the same return-of-capital → preferred → catch-up → carry sequence
required for every distribution, recap or sale, under the LPA. Run as a
counterfactual against the same cumulative fund state entering Q2 2026 (same
contributions, same preferred return accrued, same catch-up paid to date),
misclassifying the recap changes not just its own tier allocation but *the
tier boundaries the quarter's other two exits are tested against* — because
the waterfall's tiers are cumulative and order-independent only at the total
level, not at the level of which specific distribution "uses up" a tier. The
result: the administrator's sequencing under-remits the GP by $328,903 versus
the correct calculation, more than double R-2's dollar impact despite
affecting a single distribution event rather than four quarters of fee runs.
This is the reconciliation's central finding — a classification error in how
one event is *tagged* can move real dollars across every downstream
distribution in the same period, which is why the fix requires re-sequencing
the full quarter, not just re-tagging the one line item.

**Net effect.** $150,000 flows back to LPs; $328,903 flows to the GP. These
are not netted against each other in the model — they involve different
counterparties and different LPA provisions, and each requires its own
adjusting entry and notice.

---

## 8. Recommendations

1. **Automate the fee basis step-down** rather than relying on the fund
   administrator to manually switch formulas at the investment period
   boundary — R-2 is exactly the kind of break a scheduled formula check would
   catch before it compounds across four quarters.
2. **Route every distribution — sale or recap — through the same waterfall
   calculation**, with no branch for "type of realization." R-3 shows that a
   classification shortcut, not a math error, was the root cause.
3. **Reconcile management fee runs and distribution tier allocations every
   quarter**, not only at annual audit — both open breaks here would have been
   caught one quarter after they occurred instead of accumulating.
4. **Watch portfolio concentration.** Two of ten companies account for the
   majority of realized gains to date; the fund's net performance is more
   dependent on Vantage Robotics and Summit Logistics than a ten-company book
   might suggest at a glance.
5. **Confirm the preferred return compounding convention (quarterly, on the
   declining unreturned balance) is documented identically in the LPA and the
   fund administrator's system** — this is the single assumption every
   downstream tier depends on, and it was not the source of any break here,
   but it is the one that would be hardest to catch if it drifted.

---

## 9. Limitations

- Synthetic data with an authored return path (see `generate_data.py`);
  effect sizes, exit multiples, and the reconciliation breaks are illustrative,
  not empirical.
- The waterfall models a whole-fund (European) structure with a single 8%
  hurdle and 100% GP catch-up; deal-by-deal (American) waterfalls, multiple
  hurdle tiers, or a clawback provision are not represented.
- No GP capital commitment is modeled — the GP's economics here are carried
  interest only, which simplifies the capital account structure but
  understates a real GP's own capital at risk.
- The reconciliation's four breaks are illustrative categories (timing,
  fee-basis, classification, and a small cash break); a real quarterly
  reconciliation would also cover position-level valuation differences and
  cross-custodian cash reconciliation, which are not modeled here.
- All LPs are called and distributed strictly pro rata with no side letters,
  excuse rights, or differing fee arrangements — a simplification real fund
  documents rarely permit uniformly across 15 LPs.
