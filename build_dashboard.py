"""
Builds dashboard.html from model_output.json — a self-contained, no-external-
assets fund performance dashboard: KPI tiles, the capital call / distribution
cash flow curve, quarterly NAV rollforward, the distribution waterfall by
tier, LP commitment concentration, and the GP-ledger-vs-fund-admin
reconciliation.

Usage:
    python3 build_dashboard.py
Reads:
    model_output.json
Writes:
    dashboard.html
"""
import json

with open("model_output.json") as f:
    M = json.load(f)

quarters = M["quarters"]
for q in quarters:
    q["year"] = int(q["year"])
    q["quarter"] = int(q["quarter"])
lps = sorted(M["lps"], key=lambda x: -x["commitment"])
capital_calls = M["capital_calls"]
waterfall = M["waterfall"]
nav_rf = M["nav_rollforward"]
perf = M["performance_summary"]
recon = M["reconciliation"]
exits = M["exit_events"]

# Palette: dataviz skill reference instance, validated adjacent-pairlist (stacks/bars/lines).
SLOTS_LIGHT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
SLOTS_DARK = ["#3987e5", "#d95926", "#199e70", "#c98500", "#d55181"]
TIER_NAMES = ["Return of Capital", "Preferred Return", "GP Catch-up", "Carry to LPs (80%)", "Carry to GP (20%)"]


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def money(x, dp=0):
    if x is None:
        return "-"
    sign = "-" if x < 0 else ""
    return f"{sign}${abs(x):,.{dp}f}"


def moneym(x):
    return f"${x/1e6:,.1f}M"


def pct(x, dp=1):
    return f"{x * 100:.{dp}f}%"


def mult(x):
    return f"{x:.2f}x"


# ---------------------------------------------------------------------------
# Components
# ---------------------------------------------------------------------------
def tile(label, value, sub):
    return f"""
      <div class="tile">
        <div class="tile-label">{esc(label)}</div>
        <div class="tile-value">{esc(value)}</div>
        <div class="tile-sub">{esc(sub)}</div>
      </div>"""


def bar_row(label, width_pct, value_text, color_var, sub=""):
    sub_html = f'<span class="row-sub">{esc(sub)}</span>' if sub else ""
    return f"""
        <div class="row" tabindex="0" data-tip="{esc(label)} — {esc(value_text)}{(' · ' + esc(sub)) if sub else ''}">
          <div class="row-label">{esc(label)}</div>
          <div class="row-track">
            <div class="row-fill" style="width:{max(width_pct,0.6):.2f}%;background:{color_var}"></div>
          </div>
          <div class="row-value">{esc(value_text)}{sub_html}</div>
        </div>"""


def card(title, note, body, wide=False):
    return f"""
      <section class="card{' wide' if wide else ''}">
        <h2>{esc(title)}</h2>
        <p class="note">{esc(note)}</p>
        {body}
      </section>"""


def table(caption, headers, rows):
    head = "".join(f"<th>{esc(h)}</th>" for h in headers)
    body = "".join("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>" for r in rows)
    return f"<table><caption>{esc(caption)}</caption><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>"


# ---------------------------------------------------------------------------
# KPI tiles
# ---------------------------------------------------------------------------
tiles = "".join([
    tile("Committed capital", moneym(perf["committed_capital"]), f"{len(lps)} LPs"),
    tile("Called capital", moneym(perf["called_capital"]), f"{pct(perf['pct_called'])} of commitments"),
    tile("Ending NAV", moneym(perf["ending_nav"]), perf["as_of_quarter"]),
    tile("Cumulative LP distributions", moneym(perf["cum_lp_distributions"]), f"DPI {mult(perf['dpi'])}"),
    tile("Net IRR (to LPs)", pct(perf["net_irr"], 2), f"Net TVPI {mult(perf['net_tvpi'])}"),
    tile("Gross IRR (portfolio)", pct(perf["gross_irr"], 2), f"Gross MOIC {mult(perf['gross_moic'])}"),
    tile("RVPI", mult(perf["rvpi"]), "Residual value / paid-in"),
    tile("GP carried interest earned", moneym(perf["cum_gp_carry_earned"]), "Cumulative, all tiers"),
])

# ---------------------------------------------------------------------------
# Cash flow curve: capital calls (down) vs distributions to LPs (up), by quarter
# ---------------------------------------------------------------------------
W, HGT = 900, 300
PAD_L, PAD_R, PAD_T, PAD_B = 54, 16, 16, 34
n = len(quarters)
call_by_q = {c["quarter_index"]: c["total_call"] for c in capital_calls}
dist_by_q = {w["quarter_index"]: w["total_to_lps"] + w["total_to_gp"] for w in waterfall}
max_val = max(max(call_by_q.values()), max(dist_by_q.values())) * 1.12
mid_y = PAD_T + (HGT - PAD_T - PAD_B) / 2
scale = (HGT - PAD_T - PAD_B) / 2 / max_val
bw = (W - PAD_L - PAD_R) / n * 0.62


def cx(i):
    return PAD_L + (i + 0.5) * (W - PAD_L - PAD_R) / n


cf_bars, cf_hits, cf_xlabels = "", "", ""
for i, q in enumerate(quarters):
    qi = q["quarter_index"]
    call = call_by_q.get(qi, 0)
    dist = dist_by_q.get(qi, 0)
    x = cx(i) - bw / 2
    if call > 0:
        h = call * scale
        cf_bars += f'<rect x="{x:.1f}" y="{mid_y:.1f}" width="{bw:.1f}" height="{h:.1f}" class="cf-call"/>'
    if dist > 0:
        h = dist * scale
        cf_bars += f'<rect x="{x:.1f}" y="{mid_y-h:.1f}" width="{bw:.1f}" height="{h:.1f}" class="cf-dist"/>'
    cf_hits += (f'<rect x="{cx(i)-bw/1.4:.1f}" y="{PAD_T}" width="{bw*1.4:.1f}" height="{HGT-PAD_T-PAD_B}" '
                f'class="hit" data-tip="{q["label"]} — Calls {money(call)} · Distributions {money(dist)}"/>')
    if qi % 2 == 1:
        cf_xlabels += f'<text x="{cx(i):.1f}" y="{HGT-14}" class="tick" text-anchor="middle">{q["label"][:2]}{q["year"]%100:02d}</text>'

cf_chart = f"""
        <svg viewBox="0 0 {W} {HGT}" class="linechart" role="img"
             aria-label="Quarterly capital calls versus distributions to LPs">
          <line x1="{PAD_L}" y1="{mid_y:.1f}" x2="{W-PAD_R}" y2="{mid_y:.1f}" class="axis"/>
          {cf_bars}{cf_xlabels}
          {cf_hits}
        </svg>
        <div class="legend-row">
          <span class="legend-item"><span class="swatch" style="background:var(--series-1)"></span>Capital calls</span>
          <span class="legend-item"><span class="swatch" style="background:var(--series-2)"></span>Distributions to LPs</span>
        </div>"""

# ---------------------------------------------------------------------------
# NAV rollforward line chart
# ---------------------------------------------------------------------------
nav_vals = [r["ending_nav"] for r in nav_rf]
lo, hi = 0, max(nav_vals) * 1.08


def hx(i):
    return PAD_L + i * (W - PAD_L - PAD_R) / (n - 1)


def hy(v):
    return PAD_T + (hi - v) / (hi - lo) * (HGT - PAD_T - PAD_B)


points = " ".join(f"{hx(i):.1f},{hy(v):.1f}" for i, v in enumerate(nav_vals))
area_points = f"{hx(0):.1f},{hy(0):.1f} " + points + f" {hx(n-1):.1f},{hy(0):.1f}"
gridlines, ticks = "", ""
for g in range(4):
    v = lo + (hi - lo) * g / 3
    y = hy(v)
    gridlines += f'<line x1="{PAD_L}" y1="{y:.1f}" x2="{W-PAD_R}" y2="{y:.1f}" class="grid"/>'
    ticks += f'<text x="{PAD_L-8}" y="{y+4:.1f}" class="tick" text-anchor="end">{moneym(v)}</text>'
nav_xlabels, nav_hits, markers = "", "", ""
for i, (q, v) in enumerate(zip(quarters, nav_vals)):
    x, y = hx(i), hy(v)
    if q["quarter_index"] % 2 == 1:
        nav_xlabels += f'<text x="{x:.1f}" y="{HGT-14}" class="tick" text-anchor="middle">{q["label"][:2]}{q["year"]%100:02d}</text>'
    nav_hits += (f'<rect x="{x-16:.1f}" y="{PAD_T}" width="32" height="{HGT-PAD_T-PAD_B}" '
                 f'class="hit" data-x="{x:.1f}" data-tip="{q["label"]} — Ending NAV {money(v)}"/>')
    if i in (0, n - 1) or q["quarter_index"] == M["investment_period_quarters"]:
        markers += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="4" class="dot"/>'
inv_end_x = hx(M["investment_period_quarters"] - 1)
nav_chart = f"""
        <svg viewBox="0 0 {W} {HGT}" class="linechart" role="img"
             aria-label="Fund ending NAV by quarter, {money(nav_vals[-1])} as of {quarters[-1]['label']}">
          {gridlines}
          <rect x="{PAD_L}" y="{PAD_T}" width="{inv_end_x-PAD_L:.1f}" height="{HGT-PAD_T-PAD_B}" class="period-shade"/>
          <line x1="{PAD_L}" y1="{HGT-PAD_B}" x2="{W-PAD_R}" y2="{HGT-PAD_B}" class="axis"/>
          {ticks}{nav_xlabels}
          <polygon points="{area_points}" class="area"/>
          <polyline points="{points}" class="series"/>
          {markers}
          <line class="crosshair" x1="0" y1="{PAD_T}" x2="0" y2="{HGT-PAD_B}" style="opacity:0"/>
          {nav_hits}
        </svg>
        <p class="note">Shaded region: the 3-year investment period (Q1 2022 – Q4 2024). NAV climbs through
        deployment, then rides realizations and remaining marks through the harvest period.</p>"""

# ---------------------------------------------------------------------------
# Distribution waterfall — stacked bars, quarters with a distribution event
# ---------------------------------------------------------------------------
dist_quarters = [w for w in waterfall if w["distributable"] > 0]
tier_keys = ["tier1_return_of_capital", "tier2_preferred_return", "tier3_gp_catchup",
             "tier4_carry_split_lp_80", "tier4_carry_split_gp_20"]
wf_max = max(w["distributable"] for w in dist_quarters)
wf_rows = ""
for w in dist_quarters:
    segs = ""
    x_cursor = 0.0
    seg_track_w = 100.0 * w["distributable"] / wf_max
    for tk, color in zip(tier_keys, SLOTS_LIGHT):
        val = w[tk]
        if val <= 0:
            continue
        seg_w = val / w["distributable"] * seg_track_w
        segs += f'<div class="seg" style="width:{seg_w:.3f}%;background:{color}" title="{tk}"></div>'
        x_cursor += seg_w
    label = next(q["label"] for q in quarters if q["quarter_index"] == w["quarter_index"])
    tip_parts = " · ".join(f"{n_}: {money(w[k])}" for n_, k in zip(TIER_NAMES, tier_keys) if w[k] > 0)
    wf_rows += f"""
        <div class="wf-row" tabindex="0" data-tip="{esc(label)} — {esc(tip_parts)}">
          <div class="row-label">{esc(label)}</div>
          <div class="wf-track" style="width:{seg_track_w:.2f}%">{segs}</div>
          <div class="row-value">{money(w['distributable'])}</div>
        </div>"""
wf_legend = "".join(
    f'<span class="legend-item"><span class="swatch" style="background:{c}"></span>{esc(t)}</span>'
    for t, c in zip(TIER_NAMES, SLOTS_LIGHT)
)

# ---------------------------------------------------------------------------
# LP commitment concentration
# ---------------------------------------------------------------------------
lp_max = lps[0]["commitment"]
lp_rows = "".join(
    bar_row(lp["lp_name"], lp["commitment"] / lp_max * 100, moneym(lp["commitment"]),
            "var(--series-1)", sub=f"{pct(lp['commitment_pct'])} of fund · {lp['lp_type']}")
    for lp in lps
)

# ---------------------------------------------------------------------------
# Reconciliation cards
# ---------------------------------------------------------------------------
def status_class(status):
    return "status-good" if status.startswith("Resolved") else "status-warn"


recon_cards = ""
for b in recon["breaks"]:
    cls = status_class(b["status"])
    var_txt = money(b["variance"])
    recon_cards += f"""
        <div class="recon-card">
          <div class="recon-head">
            <span class="recon-id">{esc(b['break_id'])}</span>
            <span class="recon-badge {cls}">{esc(b['status'].split('—')[0].strip())}</span>
          </div>
          <div class="recon-cat">{esc(b['category'])}</div>
          <p class="recon-desc">{esc(b['description'])}</p>
          <div class="recon-figs">
            <div><span class="fig-label">GP Ledger</span><span class="fig-val">{money(b['gp_ledger_amount'])}</span></div>
            <div><span class="fig-label">Fund Admin</span><span class="fig-val">{money(b['fund_admin_amount'])}</span></div>
            <div><span class="fig-label">Variance</span><span class="fig-val fig-variance">{var_txt}</span></div>
          </div>
          <p class="recon-res"><strong>Resolution:</strong> {esc(b['resolution'])}</p>
        </div>"""

open_total = sum(abs(item["variance"]) for item in recon["open_items_detail"])

# ---------------------------------------------------------------------------
# Data tables (accessibility fallback)
# ---------------------------------------------------------------------------
tables = "".join([
    table("Capital call schedule", ["Quarter", "Investment call", "Mgmt fee call", "Total call"],
          [(c["label"], money(c["investment_call"]), money(c["mgmt_fee_call"]), money(c["total_call"]))
           for c in capital_calls]),
    table("Distribution waterfall", ["Quarter", "Distributable", "ROC", "Pref", "GP Catch-up",
                                      "Carry (LP)", "Carry (GP)", "Total to LPs", "Total to GP"],
          [(w["label"], money(w["distributable"]), money(w["tier1_return_of_capital"]),
            money(w["tier2_preferred_return"]), money(w["tier3_gp_catchup"]),
            money(w["tier4_carry_split_lp_80"]), money(w["tier4_carry_split_gp_20"]),
            money(w["total_to_lps"]), money(w["total_to_gp"])) for w in dist_quarters]),
    table("NAV rollforward", ["Quarter", "Beginning NAV", "Contributions", "Mgmt fees",
                               "Realized G/L", "Unrealized G/L", "Distributions", "Ending NAV"],
          [(r["label"], money(r["beginning_nav"]), money(r["contributions"]), money(r["management_fees"]),
            money(r["realized_gain_loss"]), money(r["unrealized_gain_loss"]), money(r["distributions"]),
            money(r["ending_nav"])) for r in nav_rf]),
    table("LP register", ["LP", "Type", "Commitment", "Commitment %"],
          [(lp["lp_name"], lp["lp_type"], money(lp["commitment"]), pct(lp["commitment_pct"])) for lp in lps]),
    table("Exit events", ["Company", "Quarter", "Type", "Proceeds", "Realized gain"],
          [(e["company_name"], next(q["label"] for q in quarters if q["quarter_index"] == e["quarter_index"]),
            e["event_type"].replace("_", " "), money(e["proceeds"]), money(e["realized_gain"])) for e in exits]),
])

funnel_vars_light = "".join(f"  --series-{i+1}: {c};\n" for i, c in enumerate(SLOTS_LIGHT))
funnel_vars_dark = "".join(f"  --series-{i+1}: {c};\n" for i, c in enumerate(SLOTS_DARK))

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(M['fund_name'])} — Fund Dashboard</title>
<style>
:root {{
  color-scheme: light;
  --page: #f9f9f7;
  --surface-1: #fcfcfb;
  --text-primary: #0b0b0b;
  --text-secondary: #52514e;
  --muted: #898781;
  --grid: #e1e0d9;
  --axis: #c3c2b7;
  --border: rgba(11,11,11,0.10);
  --track: rgba(11,11,11,0.05);
  --period-shade: rgba(42,120,214,0.07);
  --status-good-bg: rgba(12,163,12,0.12);
  --status-good-fg: #006300;
  --status-warn-bg: rgba(250,178,25,0.16);
  --status-warn-fg: #8a5a00;
{funnel_vars_light}}}
@media (prefers-color-scheme: dark) {{
  :root:where(:not([data-theme="light"])) {{
    color-scheme: dark;
    --page: #0d0d0d;
    --surface-1: #1a1a19;
    --text-primary: #ffffff;
    --text-secondary: #c3c2b7;
    --muted: #898781;
    --grid: #2c2c2a;
    --axis: #383835;
    --border: rgba(255,255,255,0.10);
    --track: rgba(255,255,255,0.07);
    --period-shade: rgba(57,135,229,0.10);
    --status-good-bg: rgba(12,163,12,0.20);
    --status-good-fg: #4fd24f;
    --status-warn-bg: rgba(250,178,25,0.20);
    --status-warn-fg: #f0c05a;
{funnel_vars_dark}  }}
}}
:root[data-theme="dark"] {{
  color-scheme: dark;
  --page: #0d0d0d;
  --surface-1: #1a1a19;
  --text-primary: #ffffff;
  --text-secondary: #c3c2b7;
  --muted: #898781;
  --grid: #2c2c2a;
  --axis: #383835;
  --border: rgba(255,255,255,0.10);
  --track: rgba(255,255,255,0.07);
  --period-shade: rgba(57,135,229,0.10);
  --status-good-bg: rgba(12,163,12,0.20);
  --status-good-fg: #4fd24f;
  --status-warn-bg: rgba(250,178,25,0.20);
  --status-warn-fg: #f0c05a;
{funnel_vars_dark}}}

* {{ box-sizing: border-box; }}
body {{
  margin: 0; padding: 32px 20px 56px;
  background: var(--page); color: var(--text-primary);
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  font-size: 14px; line-height: 1.5;
}}
.wrap {{ max-width: 1160px; margin: 0 auto; }}
header {{ margin-bottom: 24px; }}
h1 {{ font-size: 23px; margin: 0 0 6px; letter-spacing: -0.01em; }}
.sub {{ color: var(--text-secondary); margin: 0; max-width: 82ch; }}
.banner {{
  margin: 16px 0 0; padding: 10px 14px; border-radius: 8px;
  border: 1px solid var(--border); background: var(--surface-1);
  color: var(--text-secondary); font-size: 13px;
}}
.tiles {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; margin: 20px 0; }}
.tile {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }}
.tile-label {{ color: var(--text-secondary); font-size: 12px; text-transform: uppercase; letter-spacing: 0.04em; }}
.tile-value {{ font-size: 26px; font-weight: 600; letter-spacing: -0.02em; margin: 4px 0 2px; font-variant-numeric: tabular-nums; }}
.tile-sub {{ color: var(--muted); font-size: 12px; }}
.grid-2 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(440px, 1fr)); gap: 16px; }}
.card {{ background: var(--surface-1); border: 1px solid var(--border); border-radius: 10px; padding: 18px 20px 20px; margin-bottom: 16px; }}
.card h2 {{ font-size: 15px; margin: 0 0 2px; }}
.note {{ color: var(--text-secondary); font-size: 12.5px; margin: 0 0 14px; }}
.bars {{ display: flex; flex-direction: column; gap: 8px; }}
.row {{ display: grid; grid-template-columns: 220px 1fr 150px; align-items: center; gap: 10px; border-radius: 6px; outline: none; }}
.row:hover, .row:focus-visible {{ background: var(--track); }}
.row-label {{ color: var(--text-secondary); font-size: 12.5px; }}
.row-track {{ background: var(--track); border-radius: 0 4px 4px 0; height: 16px; }}
.row-fill {{ height: 16px; border-radius: 0 4px 4px 0; min-width: 3px; }}
.row-value {{ font-size: 13px; font-weight: 600; font-variant-numeric: tabular-nums; text-align: right; }}
.row-sub {{ display: block; font-weight: 400; color: var(--muted); font-size: 11.5px; }}
.linechart {{ width: 100%; height: auto; overflow: visible; }}
.grid {{ stroke: var(--grid); stroke-width: 1; }}
.axis {{ stroke: var(--axis); stroke-width: 1; }}
.tick {{ fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }}
.series {{ fill: none; stroke: var(--series-1); stroke-width: 2; stroke-linejoin: round; stroke-linecap: round; }}
.area {{ fill: var(--series-1); opacity: 0.10; }}
.dot {{ fill: var(--series-1); stroke: var(--surface-1); stroke-width: 2; }}
.period-shade {{ fill: var(--period-shade); }}
.crosshair {{ stroke: var(--axis); stroke-width: 1; stroke-dasharray: 3 3; }}
.hit {{ fill: transparent; }}
.cf-call {{ fill: var(--series-1); }}
.cf-dist {{ fill: var(--series-2); }}
.legend-row {{ display: flex; gap: 18px; margin-top: 10px; }}
.legend-item {{ display: inline-flex; align-items: center; gap: 6px; color: var(--text-secondary); font-size: 12.5px; }}
.swatch {{ width: 10px; height: 10px; border-radius: 2px; display: inline-block; }}
.wf-row {{ display: grid; grid-template-columns: 90px 1fr 130px; align-items: center; gap: 10px; padding: 5px 0; border-radius: 6px; outline: none; }}
.wf-row:hover, .wf-row:focus-visible {{ background: var(--track); }}
.wf-track {{ display: flex; height: 20px; border-radius: 3px; overflow: hidden; background: var(--track); }}
.seg {{ height: 100%; }}
details {{ margin-top: 8px; }}
summary {{ cursor: pointer; color: var(--text-secondary); font-size: 13px; padding: 6px 0; }}
table {{ border-collapse: collapse; width: 100%; margin: 12px 0 22px; font-size: 12.5px; }}
caption {{ text-align: left; color: var(--text-secondary); font-weight: 600; padding: 6px 0; }}
th, td {{ text-align: right; padding: 6px 10px; border-bottom: 1px solid var(--grid); font-variant-numeric: tabular-nums; }}
th:first-child, td:first-child {{ text-align: left; font-variant-numeric: normal; }}
th {{ color: var(--muted); font-weight: 500; }}
.recon-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 14px; }}
.recon-card {{ background: var(--page); border: 1px solid var(--border); border-radius: 10px; padding: 14px 16px; }}
.recon-head {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }}
.recon-id {{ font-weight: 700; font-size: 13px; }}
.recon-badge {{ font-size: 11px; font-weight: 600; padding: 3px 8px; border-radius: 999px; }}
.status-good {{ background: var(--status-good-bg); color: var(--status-good-fg); }}
.status-warn {{ background: var(--status-warn-bg); color: var(--status-warn-fg); }}
.recon-cat {{ font-size: 12.5px; color: var(--text-secondary); font-weight: 600; margin-bottom: 6px; }}
.recon-desc {{ font-size: 12.5px; color: var(--text-secondary); margin: 0 0 10px; }}
.recon-figs {{ display: flex; gap: 16px; margin-bottom: 10px; padding: 8px 10px; background: var(--track); border-radius: 8px; }}
.fig-label {{ display: block; font-size: 10.5px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }}
.fig-val {{ display: block; font-size: 13px; font-weight: 600; font-variant-numeric: tabular-nums; }}
.fig-variance {{ color: var(--series-2); }}
.recon-res {{ font-size: 12px; color: var(--text-secondary); margin: 0; }}
.callout {{ margin-top: 14px; padding: 10px 12px; border-radius: 8px; background: var(--track); color: var(--text-secondary); font-size: 12.5px; }}
.callout strong {{ color: var(--text-primary); }}
footer {{ color: var(--muted); font-size: 12px; margin-top: 24px; }}
#tip {{
  position: fixed; pointer-events: none; opacity: 0; transition: opacity .1s;
  background: var(--text-primary); color: var(--surface-1);
  padding: 6px 9px; border-radius: 6px; font-size: 12px; white-space: nowrap; z-index: 10;
}}
@media (max-width: 640px) {{ .row {{ grid-template-columns: 130px 1fr 100px; }} .wf-row {{ grid-template-columns: 64px 1fr 90px; }} }}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>{esc(M['fund_name'])} — Fund Dashboard</h1>
    <p class="sub">$50M committed capital, {len(lps)} LPs, 3-year investment period. Capital call schedule,
      quarterly NAV rollforward, the return-of-capital → preferred return → GP catch-up → carry waterfall,
      and a GP-ledger-vs-fund-admin reconciliation — as of {perf['as_of_quarter']}.</p>
    <p class="banner"><strong>Synthetic fund.</strong> Generated by <code>generate_data.py</code> to
      demonstrate the calculation framework — capital calls, the waterfall, NAV rollforward, capital
      accounts, IRR/MOIC, and reconciliation. Not a real fund or real LPs.</p>
  </header>

  <div class="tiles">{tiles}</div>

  {card("Capital calls vs. distributions to LPs", "Quarterly cash flow — calls drawn down against commitments, distributions returned from realizations.", cf_chart)}

  {card("Quarterly NAV rollforward", f"Fund ending NAV by quarter, {money(nav_vals[-1])} as of {quarters[-1]['label']}.", nav_chart)}

  {card("Distribution waterfall, by quarter", "Each distribution event, split across the four tiers: return of capital, preferred return, GP catch-up, and the 80/20 carry split.", f'<div class="bars">{wf_rows}</div><div class="legend-row" style="margin-top:14px;flex-wrap:wrap">{wf_legend}</div>')}

  <div class="grid-2">
    {card("LP commitment concentration", "All 15 LPs, ranked by commitment.", f'<div class="bars">{lp_rows}</div>')}
    {card("Fund performance summary", "Gross vs. net, and value composition.", f'''
        <div class="callout"><strong>Gross:</strong> {pct(perf["gross_irr"],2)} IRR · {mult(perf["gross_moic"])} MOIC on the portfolio, pre-fee, pre-carry.</div>
        <div class="callout" style="margin-top:8px"><strong>Net (to LPs):</strong> {pct(perf["net_irr"],2)} IRR · {mult(perf["net_tvpi"])} TVPI = {mult(perf["dpi"])} DPI + {mult(perf["rvpi"])} RVPI.</div>
        <div class="callout" style="margin-top:8px">GP has earned <strong>{money(perf["cum_gp_carry_earned"])}</strong> in carried interest to date, all in the final two quarters once the preferred return hurdle cleared.</div>
    ''')}
  </div>

  <section class="card wide">
    <h2>Variance reconciliation — GP ledger vs. fund administrator</h2>
    <p class="note">Four breaks identified in the {perf['as_of_quarter']} tie-out. Two resolved (timing-only /
      immaterial); two open, totaling {money(open_total)} in adjusting entries pending.</p>
    <div class="recon-grid">{recon_cards}</div>
  </section>

  <section class="card">
    <h2>Data tables</h2>
    <p class="note">Every chart above, as text.</p>
    <details><summary>Show all figures</summary>{tables}</details>
  </section>

  <footer>Built from <code>model_output.json</code> by <code>build_dashboard.py</code> · underlying calculations
    in <code>calculations.py</code> · full model in <code>PE_Fund_Model.xlsx</code>.</footer>
</div>
<div id="tip" role="status"></div>
<script>
(function () {{
  var tip = document.getElementById('tip');
  function show(e, text) {{
    tip.textContent = text;
    tip.style.opacity = '1';
    var x = e.clientX + 14, y = e.clientY + 14;
    var r = tip.getBoundingClientRect();
    if (x + r.width > window.innerWidth - 8) x = e.clientX - r.width - 14;
    if (y + r.height > window.innerHeight - 8) y = e.clientY - r.height - 14;
    tip.style.left = x + 'px';
    tip.style.top = y + 'px';
  }}
  function hide() {{ tip.style.opacity = '0'; }}

  document.querySelectorAll('[data-tip]').forEach(function (el) {{
    el.addEventListener('mousemove', function (e) {{ show(e, el.getAttribute('data-tip')); }});
    el.addEventListener('mouseleave', hide);
    el.addEventListener('focus', function () {{
      var b = el.getBoundingClientRect();
      show({{ clientX: b.left + b.width / 2, clientY: b.top }}, el.getAttribute('data-tip'));
    }});
    el.addEventListener('blur', hide);
  }});

  document.querySelectorAll('.linechart').forEach(function (svg) {{
    var cross = svg.querySelector('.crosshair');
    if (!cross) return;
    svg.querySelectorAll('.hit[data-x]').forEach(function (h) {{
      h.addEventListener('mouseenter', function () {{
        var x = h.getAttribute('data-x');
        cross.setAttribute('x1', x); cross.setAttribute('x2', x);
        cross.style.opacity = '1';
      }});
    }});
    svg.addEventListener('mouseleave', function () {{ cross.style.opacity = '0'; }});
  }});
}})();
</script>
</body>
</html>
"""

with open("dashboard.html", "w") as f:
    f.write(html)

print("Wrote dashboard.html")
