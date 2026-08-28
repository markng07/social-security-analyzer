#!/usr/bin/env python3
"""
Streamlit web UI for the Social Security deferment & break-even calculator.

Run with:  streamlit run streamlit_app.py
"""

from __future__ import annotations

import io
import json
from datetime import date
from pathlib import Path
from typing import Any

import streamlit as st

import ss_deferment_calculator as calc


st.set_page_config(
    page_title="Social Security Claiming-Age Analyzer",
    page_icon="🏛️",
    layout="wide",
)

st.title("Social Security Claiming-Age Analyzer")
st.caption("Determine the optimal age to start Social Security retirement benefits.")

with st.expander("How to Use This App & Interpret Results", expanded=False):
    st.markdown("""
### Getting Started

1. **Fill in your information** in the left sidebar:
   - **Date of birth** and **Full Retirement Age (FRA)** — find these on your SSA statement.
   - **FRA monthly benefit (PIA)** — the monthly amount SSA says you'll receive if you claim at FRA.
   - **Sex** — selects the mortality table (male or female) for the actuarial analysis.

2. **Enter your salary** — either a flat annual amount, or a schedule by age (e.g., $71,537 until 67, then $0 for retirement). The salary determines how the **earnings test** affects your benefits before FRA.

3. **Set your expected end-of-life age** — this drives the fixed-lifespan analysis. The actuarial analysis (Expected Value tab) does not depend on this.

4. **Choose claim ages to compare** — select which ages (62–70) you want to analyze side by side.

5. **Optional: paste SSA statement benefits** — if your SSA statement lists specific monthly amounts by claim age, enter them to use those exact numbers instead of calculated estimates.

Results update **automatically** whenever you change any input.

---

### Understanding the Tabs

**Summary** — Start here. Shows the best claiming age under both methods:
- *Fixed lifespan*: assumes you live to exactly the age you entered.
- *Actuarial*: weights every possible lifespan by its probability — no guessing needed.
- The consecutive-pair verdicts (green = WAIT, red = CLAIM) walk through each adjacent pair of ages.

**Break-Even** — A matrix showing when waiting pays off:
- Each cell = the age at which waiting catches up to claiming earlier.
- **Green** = catches up before your expected lifespan (waiting is worthwhile).
- **Red** = catches up after, or never (claim earlier).
- The pairwise details table below shows exact dollar differences.

**Sensitivity** — How your total payout changes across different lifespans:
- Rows = claim ages, columns = checkpoint ages (70, 75, 80, 85, 90, 95).
- The **yellow column** is your expected lifespan.
- Read down a column to find the best claim age at that lifespan.
- Read across a row to see how one strategy performs if you live longer or shorter.

**Expected Value** — The most robust analysis (no lifespan guess):
- Uses SSA mortality tables to weight every possible death age by its probability.
- The **green row** is the best claim age by expected value.
- If the spread between best and worst is small (e.g., $12k), the decision barely matters financially — choose based on cash-flow needs instead.
- Survival probabilities help you gauge how realistic fixed-lifespan scenarios are.

**Projection** — The full year-by-year data behind all calculations:
- Filter by claim age using the dropdown.
- Key columns: `paid_social_security` (cash received that year), `withheld_deferred_due_to_work` (held back by earnings test — not lost), `cumulative_paid_social_security` (running total).

---

### Key Concepts

- **Earnings test**: If you claim before FRA and still work, SSA withholds some benefits. This money is **not lost** — your benefit is recalculated at FRA to credit the withheld months. But claiming early while earning a high salary means you take the permanent early-claiming reduction for little or no immediate cash.
- **Early claiming reduction**: Up to 30% permanent reduction for claiming at 62 (with FRA 67).
- **Delayed retirement credits**: 8% per year increase for each year you delay past FRA, up to age 70.
- **This model does NOT include**: COLA, inflation, taxes, Medicare premiums, or NPV discounting. All comparisons are in today's dollars.

---

### Downloads

Use the download buttons at the bottom to save:
- **Excel (.xlsx)** — the same 5-tab workbook, formatted for printing or sharing.
- **CSV** — raw projection data for all scenarios.
- **JSON** — your input settings, so you can reload or share your configuration.
""")


# ── Sidebar: inputs ──────────────────────────────────────────────────────────

with st.sidebar:
    st.header("Your Information")

    birth_date = st.date_input(
        "Date of birth",
        value=date(1962, 10, 3),
        min_value=date(1930, 1, 1),
        max_value=date.today(),
    )

    fra_years = st.number_input("Full Retirement Age (years)", 65, 67, 67)
    fra_months = st.number_input("FRA additional months", 0, 11, 0)

    pia = st.number_input(
        "FRA monthly benefit (PIA) $",
        min_value=0.0,
        value=3438.0,
        step=50.0,
        format="%.2f",
    )

    sex = st.selectbox("Sex (for life table)", ["Male", "Female"], index=0)

    st.divider()
    st.header("Work & Salary")

    salary_mode = st.radio(
        "Salary input",
        ["Flat salary", "Salary schedule (by age)"],
        index=0,
    )
    if salary_mode == "Flat salary":
        flat_salary = st.number_input(
            "Annual salary $", min_value=0.0, value=71537.0, step=1000.0, format="%.0f"
        )
        salary_config: Any = flat_salary
    else:
        st.caption("Enter one salary per line: `age salary` (e.g. `62 71537`). "
                   "Missing ages carry the previous value forward.")
        salary_text = st.text_area(
            "Age → Salary",
            value="62 71537\n67 0",
            height=120,
        )
        salary_config = {}
        for line in salary_text.strip().splitlines():
            parts = line.split()
            if len(parts) == 2:
                salary_config[parts[0]] = float(parts[1])

    st.divider()
    st.header("Analysis Settings")

    end_of_life_years = st.slider("Expected end-of-life age", 70, 100, 90)

    st.caption("Claim ages to compare (select all that apply)")
    claim_age_options = list(range(62, 71))
    selected_claim_ages = st.multiselect(
        "Claim ages",
        options=claim_age_options,
        default=claim_age_options,
        label_visibility="collapsed",
    )

    st.divider()
    st.header("SSA Statement Benefits")
    st.caption("Optional: paste monthly benefits from your SSA statement. "
               "Leave blank to derive from PIA.")
    benefits_text = st.text_area(
        "Age → Monthly Benefit",
        value="",
        height=120,
        placeholder="63 2695\n67 3438\n70 4268",
    )
    benefits_by_claim_age: dict[str, float] = {}
    for line in benefits_text.strip().splitlines():
        parts = line.split()
        if len(parts) == 2:
            benefits_by_claim_age[parts[0]] = float(parts[1])

    st.divider()
    st.header("Earnings Test Limits")
    lower_limit = st.number_input(
        "Lower limit (before FRA year) $", value=24480.0, step=100.0, format="%.0f"
    )
    higher_limit = st.number_input(
        "Higher limit (FRA year) $", value=65160.0, step=100.0, format="%.0f"
    )


# ── Run analysis ──────────────────────────────────────────────────────────────

if not selected_claim_ages:
    st.warning("Select at least one claim age in the sidebar.")
    st.stop()

config: dict[str, Any] = {
    "birth_date": birth_date.isoformat(),
    "full_retirement_age": {"years": fra_years, "months": fra_months},
    "end_of_life_age": {"years": end_of_life_years, "months": 0},
    "ssa_full_retirement_monthly_benefit": pia,
    "annual_salary": salary_config,
    "claim_ages": [{"years": a, "months": 0} for a in sorted(selected_claim_ages)],
    "adjust_benefit_at_fra_for_deferred_months": True,
    "default_earnings_test_limits": {
        "lower_limit": lower_limit,
        "higher_limit": higher_limit,
    },
    "earnings_test_limits_by_year": {},
}
if benefits_by_claim_age:
    config["benefits_by_claim_age"] = benefits_by_claim_age

claim_ages = [calc.parse_age(v) for v in config["claim_ages"]]
end_of_life_age = calc.parse_age(config["end_of_life_age"])
checkpoint_ages = calc.DEFAULT_LIFESPAN_CHECKPOINTS

projection_horizon = calc.Age(
    years=max(end_of_life_age.years, max(checkpoint_ages, default=0), calc.ACTUARIAL_MAX_AGE)
)

scenarios: list[tuple[calc.Age, list[dict[str, Any]]]] = []
all_rows: list[dict[str, Any]] = []
for ca in claim_ages:
    rows = calc.project_claim_scenario(config, ca, horizon_age=projection_horizon)
    scenarios.append((ca, rows))
    all_rows.extend(rows)

matrix, details = calc.compute_break_even_matrix_and_details(
    scenarios,
    min_age=62,
    max_age=end_of_life_age.years,
    expected_lifespan_age=end_of_life_age.years,
)
sensitivity_labels, sensitivity_rows = calc.compute_lifespan_sensitivity(
    scenarios,
    checkpoint_ages=checkpoint_ages,
    expected_lifespan_age=end_of_life_age.years,
)

bd = calc.parse_iso_date(config["birth_date"])
current_age = (date.today() - bd).days // 365
qx_table = calc.get_life_table(sex.lower())
ev_results = calc.compute_expected_value_analysis(
    scenarios, current_age, qx_table, max_age=calc.ACTUARIAL_MAX_AGE
)

summary = calc.build_summary(config, scenarios, end_of_life_age, ev_results=ev_results)


# ── Display results ──────────────────────────────────────────────────────────

tab_summary, tab_breakeven, tab_sensitivity, tab_ev, tab_projection = st.tabs(
    ["Summary", "Break-Even", "Sensitivity", "Expected Value", "Projection"]
)

# --- Summary tab ---
with tab_summary:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Inputs")
        st.markdown(f"""
| Field | Value |
|---|---|
| Birth date | {summary['birth_date']} |
| Full retirement age | {summary['full_retirement_age']} |
| Expected end-of-life | {summary['expected_end_of_life_age']} |
| FRA monthly benefit | ${summary['fra_monthly_benefit_pia']:,.2f} |
| Salary | {summary['salary_input']} |
| Benefits source | {summary['benefits_source']} |
""")

    with col2:
        best = summary["best_strategy_at_expected_lifespan"]
        st.subheader("Best Strategy (fixed lifespan)")
        st.metric(
            f"Claim at age {best['claim_age']}",
            f"${best['cumulative_paid']:,.0f}",
            help="Total cumulative Social Security received at expected lifespan",
        )

        if ev_results:
            best_ev = ev_results[0]
            st.subheader("Best Strategy (actuarial)")
            st.metric(
                f"Claim at age {best_ev['claim_age']}",
                f"${best_ev['expected_lifetime_benefit']:,.0f}",
                help="Mortality-weighted expected lifetime benefit",
            )
            st.caption(
                f"Life expectancy from age {current_age}: "
                f"{best_ev['life_expectancy_from_current_age']} years"
            )

    st.subheader("Consecutive-Pair Recommendations")
    for verdict in summary["consecutive_pair_verdicts"]:
        if "WAIT" in verdict:
            st.success(verdict)
        else:
            st.error(verdict)

    if ev_results:
        st.subheader("Expected Value Ranking")
        for i, ev_row in enumerate(ev_results):
            label = f"Claim {ev_row['claim_age']}: ${ev_row['expected_lifetime_benefit']:,.0f}"
            if i == 0:
                st.success(f"**{label}** — BEST")
            else:
                st.write(label)

    with st.expander("Model Caveats"):
        for caveat in summary["model_caveats"]:
            st.write(f"- {caveat}")


# --- Break-Even tab ---
with tab_breakeven:
    st.subheader("Break-Even Matrix")
    st.caption(
        f"Each cell = age at which waiting catches up. "
        f"Green = before expected lifespan ({end_of_life_years}). "
        f"Red = after or never."
    )

    import pandas as pd

    if len(matrix) > 1:
        header = matrix[0][1:]
        matrix_data = []
        original_values = []
        for row in matrix[1:]:
            matrix_data.append([str(v) for v in row[1:]])
            original_values.append(row[1:])
        df_matrix = pd.DataFrame(matrix_data, columns=header, index=[r[0] for r in matrix[1:]])

        def color_cell_by_pos(row_idx, col_idx):
            val = original_values[row_idx][col_idx]
            if isinstance(val, int):
                if val <= end_of_life_years:
                    return "background-color: #D6EFD6; color: #1a1a1a"
                return "background-color: #F7D6D6; color: #1a1a1a"
            if isinstance(val, str) and val.startswith(">"):
                return "background-color: #F7D6D6; color: #1a1a1a"
            return ""

        def color_matrix(df):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for r in range(len(df)):
                for c in range(len(df.columns)):
                    styles.iloc[r, c] = color_cell_by_pos(r, c)
            return styles

        st.dataframe(
            df_matrix.style.apply(lambda _: color_matrix(df_matrix), axis=None),
            use_container_width=True,
        )

    st.subheader("Pairwise Details")
    if details:
        df_details = pd.DataFrame(details)
        money_cols = [
            "cumulative_paid_earlier_at_expected_lifespan",
            "cumulative_paid_later_at_expected_lifespan",
            "lifetime_difference_later_minus_earlier",
        ]
        format_dict = {c: "${:,.0f}" for c in money_cols if c in df_details.columns}
        st.dataframe(
            df_details.style.format(format_dict),
            use_container_width=True,
            hide_index=True,
        )


# --- Sensitivity tab ---
with tab_sensitivity:
    st.subheader("Lifespan Sensitivity")
    st.caption("Cumulative Social Security paid at each checkpoint age.")

    if sensitivity_rows:
        df_sens = pd.DataFrame(sensitivity_rows)
        format_dict = {c: "${:,.0f}" for c in df_sens.columns if c != "claim_age"}

        def highlight_expected(col):
            if "expected_lifespan" in col.name:
                return ["background-color: #FFF3B8; color: #1a1a1a"] * len(col)
            return [""] * len(col)

        st.dataframe(
            df_sens.style.format(format_dict).apply(highlight_expected),
            use_container_width=True,
            hide_index=True,
        )


# --- Expected Value tab ---
with tab_ev:
    st.subheader("Actuarial Expected Value")
    st.caption(
        "Each claim age weighted by SSA 2021 Period Life Table mortality. "
        "No single lifespan guess needed."
    )

    if ev_results:
        df_ev = pd.DataFrame(ev_results)
        rename_map = {
            "claim_age": "Claim Age",
            "expected_lifetime_benefit": "Expected Lifetime Benefit",
            "life_expectancy_from_current_age": "Life Expectancy (yrs)",
            "cumulative_at_life_expectancy": "Cumulative at Life Exp.",
            "P_survive_to_75": "P(75) %",
            "P_survive_to_80": "P(80) %",
            "P_survive_to_85": "P(85) %",
            "P_survive_to_90": "P(90) %",
        }
        df_ev = df_ev.rename(columns=rename_map)

        money_cols_ev = ["Expected Lifetime Benefit", "Cumulative at Life Exp."]
        format_dict = {c: "${:,.0f}" for c in money_cols_ev}
        for c in df_ev.columns:
            if c.startswith("P("):
                format_dict[c] = "{:.1f}"

        def highlight_best(row):
            if row.name == 0:
                return ["background-color: #D6EFD6; color: #1a1a1a"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_ev.style.format(format_dict).apply(highlight_best, axis=1),
            use_container_width=True,
            hide_index=True,
        )


# --- Projection tab ---
with tab_projection:
    st.subheader("Year-by-Year Projection")

    filter_claim = st.selectbox(
        "Filter by claim age",
        ["All"] + [calc.age_to_label(ca) for ca in claim_ages],
    )

    df_proj = pd.DataFrame(all_rows)
    if filter_claim != "All":
        df_proj = df_proj[df_proj["claim_age"] == filter_claim]

    money_fields = [
        "salary", "monthly_benefit_at_year_end", "gross_social_security",
        "withheld_deferred_due_to_work", "paid_social_security",
        "cumulative_paid_social_security", "cumulative_deferred_due_to_work",
    ]
    format_dict = {c: "${:,.2f}" for c in money_fields if c in df_proj.columns}
    st.dataframe(
        df_proj.style.format(format_dict),
        use_container_width=True,
        hide_index=True,
        height=600,
    )


# ── Downloads ─────────────────────────────────────────────────────────────────

st.divider()
st.subheader("Download Results")

col_dl1, col_dl2, col_dl3 = st.columns(3)

with col_dl1:
    xlsx_buffer = io.BytesIO()
    xlsx_path = Path("/tmp/ss_analysis.xlsx")
    ok = calc.write_excel(
        xlsx_path,
        projection_rows=all_rows,
        matrix=matrix,
        break_even_details=details,
        sensitivity_labels=sensitivity_labels,
        sensitivity_rows=sensitivity_rows,
        summary=summary,
        expected_lifespan_age=end_of_life_years,
        ev_results=ev_results,
    )
    if ok:
        xlsx_buffer = xlsx_path.read_bytes()
        st.download_button(
            "Download Excel (.xlsx)",
            data=xlsx_buffer,
            file_name="social_security_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with col_dl2:
    csv_buffer = io.StringIO()
    if all_rows:
        import csv
        fieldnames = list(all_rows[0].keys())
        writer = csv.DictWriter(csv_buffer, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)
    st.download_button(
        "Download CSV",
        data=csv_buffer.getvalue(),
        file_name="social_security_projection.csv",
        mime="text/csv",
    )

with col_dl3:
    st.download_button(
        "Download Input (JSON)",
        data=json.dumps(config, indent=2),
        file_name="social_security_input.json",
        mime="application/json",
    )
