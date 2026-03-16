import os
import io
import re

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from openai import OpenAI
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference
from pptx import Presentation
from pptx.util import Inches as PPTInches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor


# ==================================================
# OPENAI CLIENT
# ==================================================
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


# ==================================================
# PAGE CONFIG
# ==================================================
st.set_page_config(page_title="TurboPitch AI", layout="wide")

# ==================================================
# CUSTOM CSS
# ==================================================
st.markdown("""
<style>
.kpi-card {
    padding: 18px;
    border-radius: 12px;
    color: white;
    margin-bottom: 10px;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
}
.kpi-blue {
    background: linear-gradient(135deg, #1f4e78, #2f75b5);
}
.kpi-green {
    background: linear-gradient(135deg, #2e7d32, #66bb6a);
}
.kpi-red {
    background: linear-gradient(135deg, #b71c1c, #ef5350);
}
.kpi-gold {
    background: linear-gradient(135deg, #b8860b, #f4c542);
    color: black;
}
.kpi-title {
    font-size: 14px;
    font-weight: 600;
    opacity: 0.9;
}
.kpi-value {
    font-size: 28px;
    font-weight: 700;
    margin-top: 6px;
}
.snapshot-card {
    padding: 16px;
    border-radius: 12px;
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.08);
}
</style>
""", unsafe_allow_html=True)

st.title("TurboPitch AI")
st.subheader("Turn an idea into investor-ready materials — and pressure-test it like a VC would.")


# ==================================================
# SESSION STATE DEFAULTS
# ==================================================
defaults = {
    "idea": "",
    "price_per_unit": 5.0,
    "year1_units": 80000,
    "growth_y2": 0.60,
    "growth_y3": 0.40,
    "cost_per_unit": 1.20,
    "opex_pct": 0.25,
    "fixed_overhead": 300000.0,
    "starting_cash": 600000.0,
    "pushback_pct": 25,
    "investor_feedback": "Investor says revenue assumptions are too aggressive and early customer acquisition may be unrealistic.",
    "sanity_output": "",
    "business_plan_output": "",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ==================================================
# HELPER FUNCTIONS
# ==================================================
def clean_ai_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("###", "")
    text = text.replace("##", "")
    text = text.replace("**", "")
    text = text.replace("*", "")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def create_revenue_chart_image(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    ax.plot(df["Year"], df["Revenue"], marker="o", linewidth=2.5)
    ax.set_title("Revenue Projection Summary", fontsize=16, fontweight="bold", pad=14)
    ax.set_xlabel("Year", fontsize=11)
    ax.set_ylabel("Revenue ($)", fontsize=11)
    ax.grid(True, alpha=0.25)
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    img_stream = io.BytesIO()
    fig.tight_layout()
    fig.savefig(img_stream, format="png", dpi=220, bbox_inches="tight")
    img_stream.seek(0)
    plt.close(fig)
    return img_stream


def create_ppt_financial_chart_image(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.plot(df["Year"], df["Revenue"], marker="o", label="Revenue")
    ax.plot(df["Year"], df["Net Income"], marker="o", label="Net Income")
    ax.plot(df["Year"], df["Ending Cash"], marker="o", label="Ending Cash")
    ax.set_title("Financial Projection Overview")
    ax.set_xlabel("Year")
    ax.set_ylabel("Dollars")
    ax.grid(True, alpha=0.3)
    ax.legend()

    img_stream = io.BytesIO()
    fig.tight_layout()
    fig.savefig(img_stream, format="png", dpi=200)
    img_stream.seek(0)
    plt.close(fig)
    return img_stream


def build_projection(price, year1_units, growth_y2, growth_y3, cost_per_unit, opex_pct, fixed_overhead, starting_cash):
    year2_units = int(year1_units * (1 + growth_y2))
    year3_units = int(year2_units * (1 + growth_y3))

    years = ["Year 1", "Year 2", "Year 3"]
    units = [year1_units, year2_units, year3_units]

    revenue = [u * price for u in units]
    cogs = [u * cost_per_unit for u in units]
    gross_profit = [r - c for r, c in zip(revenue, cogs)]
    gross_margin_pct = [(gp / r) if r != 0 else 0 for gp, r in zip(gross_profit, revenue)]

    operating_expenses = [(r * opex_pct) + fixed_overhead for r in revenue]
    operating_income = [gp - op for gp, op in zip(gross_profit, operating_expenses)]

    tax_rate = 0.21
    taxes = [max(0, oi * tax_rate) for oi in operating_income]
    net_income = [oi - tax for oi, tax in zip(operating_income, taxes)]

    ending_cash = []
    cash_balance = starting_cash
    for ni in net_income:
        cash_balance += ni
        ending_cash.append(cash_balance)

    return pd.DataFrame({
        "Year": years,
        "Units": units,
        "Revenue": revenue,
        "COGS": cogs,
        "Gross Profit": gross_profit,
        "Gross Margin %": gross_margin_pct,
        "Operating Expenses": operating_expenses,
        "Operating Income": operating_income,
        "Taxes": taxes,
        "Net Income": net_income,
        "Ending Cash": ending_cash,
    })


def build_pnl_view(projection_df: pd.DataFrame) -> pd.DataFrame:
    pnl_df = pd.DataFrame({
        "Line Item": [
            "Units",
            "Revenue",
            "COGS",
            "Gross Profit",
            "Gross Margin %",
            "Operating Expenses",
            "Operating Income",
            "Taxes",
            "Net Income",
            "Ending Cash",
        ],
        "Year 1": [
            projection_df.loc[0, "Units"],
            projection_df.loc[0, "Revenue"],
            projection_df.loc[0, "COGS"],
            projection_df.loc[0, "Gross Profit"],
            projection_df.loc[0, "Gross Margin %"],
            projection_df.loc[0, "Operating Expenses"],
            projection_df.loc[0, "Operating Income"],
            projection_df.loc[0, "Taxes"],
            projection_df.loc[0, "Net Income"],
            projection_df.loc[0, "Ending Cash"],
        ],
        "Year 2": [
            projection_df.loc[1, "Units"],
            projection_df.loc[1, "Revenue"],
            projection_df.loc[1, "COGS"],
            projection_df.loc[1, "Gross Profit"],
            projection_df.loc[1, "Gross Margin %"],
            projection_df.loc[1, "Operating Expenses"],
            projection_df.loc[1, "Operating Income"],
            projection_df.loc[1, "Taxes"],
            projection_df.loc[1, "Net Income"],
            projection_df.loc[1, "Ending Cash"],
        ],
        "Year 3": [
            projection_df.loc[2, "Units"],
            projection_df.loc[2, "Revenue"],
            projection_df.loc[2, "COGS"],
            projection_df.loc[2, "Gross Profit"],
            projection_df.loc[2, "Gross Margin %"],
            projection_df.loc[2, "Operating Expenses"],
            projection_df.loc[2, "Operating Income"],
            projection_df.loc[2, "Taxes"],
            projection_df.loc[2, "Net Income"],
            projection_df.loc[2, "Ending Cash"],
        ],
    })
    return pnl_df


def run_rule_based_sanity_check(price, year1_units, growth_y2, growth_y3, cost_per_unit, opex_pct):
    warnings = []

    if year1_units > 250000:
        warnings.append("🔴 Year 1 unit sales are very high for an early-stage company and may be unrealistic.")

    if price < cost_per_unit:
        warnings.append("🔴 Price per unit is below cost per unit, which creates a structurally unprofitable model.")

    if growth_y2 > 1.0:
        warnings.append("🟠 Year 2 growth exceeds 100%, which may be viewed as aggressive by investors.")

    if growth_y3 > 1.0:
        warnings.append("🟠 Year 3 growth exceeds 100%, which may be viewed as aggressive by investors.")

    if opex_pct > 0.60:
        warnings.append("🟠 Operating expense percentage is very high and may pressure long-term profitability.")

    if not warnings:
        warnings.append("🟢 No major rule-based assumption warnings triggered.")

    return warnings


def build_scorecard(projection_df, price, cost_per_unit, year1_units, growth_y2, growth_y3):
    ending_cash_final = projection_df["Ending Cash"].iloc[-1]
    gross_margin = ((price - cost_per_unit) / price) if price else 0

    pricing_status = "Green" if price >= cost_per_unit * 1.5 else "Yellow" if price >= cost_per_unit else "Red"
    sales_status = "Green" if year1_units <= 100000 else "Yellow" if year1_units <= 250000 else "Red"
    growth_status = "Green" if growth_y2 <= 0.75 and growth_y3 <= 0.60 else "Yellow" if growth_y2 <= 1.0 and growth_y3 <= 0.80 else "Red"
    margin_status = "Green" if gross_margin >= 0.50 else "Yellow" if gross_margin >= 0.25 else "Red"
    cash_status = "Green" if ending_cash_final > 0 else "Red"

    statuses = [pricing_status, sales_status, growth_status, margin_status, cash_status]

    if statuses.count("Red") >= 2:
        overall = "Red"
    elif "Red" in statuses or statuses.count("Yellow") >= 2:
        overall = "Yellow"
    else:
        overall = "Green"

    return {
        "Pricing Realism": pricing_status,
        "Sales Volume": sales_status,
        "Growth Assumptions": growth_status,
        "Margin Quality": margin_status,
        "Cash Viability": cash_status,
        "Overall Investor Readiness": overall,
    }


def score_metric(status):
    if status == "Green":
        return "🟢"
    if status == "Yellow":
        return "🟡"
    return "🔴"


def financial_summary_text(projection_df: pd.DataFrame) -> str:
    return "\n".join([
        f"{row['Year']}: Revenue ${row['Revenue']:,.0f}, "
        f"COGS ${row['COGS']:,.0f}, "
        f"Gross Profit ${row['Gross Profit']:,.0f}, "
        f"Gross Margin {row['Gross Margin %']:.1%}, "
        f"Operating Expenses ${row['Operating Expenses']:,.0f}, "
        f"Operating Income ${row['Operating Income']:,.0f}, "
        f"Taxes ${row['Taxes']:,.0f}, "
        f"Net Income ${row['Net Income']:,.0f}, "
        f"Ending Cash ${row['Ending Cash']:,.0f}"
        for _, row in projection_df.iterrows()
    ])


def run_ai_sanity_check(
    idea,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    projection_df
):
    fin_summary = financial_summary_text(projection_df)

    prompt = f"""
You are a skeptical but helpful venture capital analyst.

Review this startup idea and financial model like an investor would.

Write a SANITY CHECK REPORT with these exact sections:

Overall Impression
What Investors Will Like
What Investors Will Challenge
Biggest Risks
Suggestions to Improve Investor Appeal
How to Reframe This Pitch for Investors

Instructions:
- Be practical, concise, and honest.
- Focus on how the founder can make the startup look more credible and fundable.
- Include specific suggestions for improving assumptions, positioning, market framing, go-to-market logic, and investor confidence.
- If assumptions are too aggressive, explain how to make them more believable.
- If the idea feels weak, suggest a stronger framing or niche angle.
- Do NOT write a full business plan.
- Do NOT write a pitch deck.
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks.

Startup Idea:
{idea if idea else "No startup idea provided."}

Assumptions:
Price per unit: ${price_per_unit:,.2f}
Year 1 units sold: {year1_units:,}
Year 2 growth rate: {growth_y2:.2%}
Year 3 growth rate: {growth_y3:.2%}
Cost per unit: ${cost_per_unit:,.2f}
Operating expense percent of revenue: {opex_pct:.2%}
Fixed annual overhead: ${fixed_overhead:,.0f}

Projection Summary:
{fin_summary}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a practical VC analyst who critiques startup ideas and helps make them more investor-ready."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.6,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI sanity check unavailable.\n\nError: {e}"


def generate_business_plan_and_deck(
    idea,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    projection_df
):
    fin_summary = financial_summary_text(projection_df)

    prompt = f"""
You are a startup strategist helping a founder prepare investor-ready materials.

Write a FULL BUSINESS PLAN followed by a separate section called Pitch Deck Content.

Use these exact BUSINESS PLAN sections:

Executive Summary
Problem
Solution
Market Opportunity
Product Overview
Business Model
Go-To-Market Strategy
Competitive Landscape
Financial Overview
Funding Ask
Key Risks

Then create a separate section called:

Pitch Deck Content

For Pitch Deck Content:
- Use numbered slides
- Each slide should have a slide title
- Each slide should include 3 to 4 short, persuasive bullet points
- Make the slide bullets investor-friendly and realistic

Slides to generate:
1. Problem
2. Solution
3. Product
4. Market Opportunity
5. Business Model
6. Competitive Advantage
7. Go-To-Market Strategy
8. Financial Highlights
9. Funding Ask

Instructions:
- Write clearly and professionally
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks
- Keep the business plan polished and readable
- Make the deck bullets concise and presentation-ready

Startup Idea:
{idea if idea else "No startup idea provided."}

Assumptions:
Price per unit: ${price_per_unit:,.2f}
Year 1 units sold: {year1_units:,}
Year 2 growth rate: {growth_y2:.2%}
Year 3 growth rate: {growth_y3:.2%}
Cost per unit: ${cost_per_unit:,.2f}
Operating expense percent of revenue: {opex_pct:.2%}
Fixed annual overhead: ${fixed_overhead:,.0f}

Projection Summary:
{fin_summary}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a practical startup strategist who creates investor-facing business plans and pitch decks."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Business plan generation unavailable.\n\nError: {e}"


def extract_business_plan_section(text: str) -> str:
    cleaned = clean_ai_text(text)
    if not cleaned:
        return ""
    if "Pitch Deck Content" in cleaned:
        return cleaned.split("Pitch Deck Content", 1)[0].strip()
    return cleaned


def extract_pitch_deck_section(text: str):
    cleaned = clean_ai_text(text)
    if not cleaned:
        return [("Startup Overview", ["No pitch deck content available yet."])]

    if "Pitch Deck Content" not in cleaned:
        return [("Startup Overview", ["No pitch deck content available yet."])]

    section = cleaned.split("Pitch Deck Content", 1)[1].strip()
    lines = [line.strip() for line in section.splitlines() if line.strip()]

    slides = []
    current_title = None
    current_bullets = []

    for line in lines:
        if re.match(r"^\d+\.\s+", line):
            if current_title:
                slides.append((current_title, current_bullets))
            current_title = re.sub(r"^\d+\.\s*", "", line).strip()
            current_bullets = []
        else:
            current_bullets.append(line)

    if current_title:
        slides.append((current_title, current_bullets))

    if not slides:
        return [("Startup Overview", ["No pitch deck content available yet."])]

    return slides


# ==================================================
# MAIN IDEA INPUT
# ==================================================
st.session_state.idea = st.text_area(
    "Startup idea",
    value=st.session_state.idea,
    height=140,
    placeholder="Describe the startup idea here..."
)


# ==================================================
# SIDEBAR INPUTS
# ==================================================
with st.sidebar:
    st.header("Core Assumptions")

    st.session_state.price_per_unit = st.number_input(
        "Price per unit ($)",
        min_value=0.0,
        value=float(st.session_state.price_per_unit),
        step=0.25,
    )

    st.session_state.year1_units = st.number_input(
        "Year 1 units sold",
        min_value=0,
        value=int(st.session_state.year1_units),
        step=1000,
    )

    st.session_state.growth_y2 = st.slider(
        "Year 2 growth rate",
        min_value=0.0,
        max_value=2.0,
        value=float(st.session_state.growth_y2),
        step=0.01,
    )

    st.session_state.growth_y3 = st.slider(
        "Year 3 growth rate",
        min_value=0.0,
        max_value=2.0,
        value=float(st.session_state.growth_y3),
        step=0.01,
    )

    st.session_state.cost_per_unit = st.number_input(
        "Cost per unit ($)",
        min_value=0.0,
        value=float(st.session_state.cost_per_unit),
        step=0.25,
    )

    st.session_state.opex_pct = st.slider(
        "Operating expense % of revenue",
        min_value=0.0,
        max_value=1.0,
        value=float(st.session_state.opex_pct),
        step=0.01,
    )

    st.session_state.fixed_overhead = st.number_input(
        "Fixed annual overhead ($)",
        min_value=0.0,
        value=float(st.session_state.fixed_overhead),
        step=10000.0,
    )

    st.session_state.starting_cash = st.number_input(
        "Starting cash ($)",
        min_value=0.0,
        value=float(st.session_state.starting_cash),
        step=10000.0,
    )

    st.markdown("---")
    st.markdown("### Investor Pushback Simulator")

    st.session_state.pushback_pct = st.number_input(
        "Reduce revenue assumptions by %",
        min_value=0,
        max_value=100,
        value=int(st.session_state.pushback_pct),
        step=5,
    )

    st.session_state.investor_feedback = st.text_area(
        "Investor feedback",
        value=st.session_state.investor_feedback,
        height=100,
    )

    if st.button("Apply Investor Pushback"):
        factor = 1 - (st.session_state.pushback_pct / 100)
        st.session_state.year1_units = int(st.session_state.year1_units * factor)
        st.success(f"Applied investor pushback: reduced Year 1 units by {st.session_state.pushback_pct}%.")


# ==================================================
# BUILD MODEL
# ==================================================
projection_df = build_projection(
    st.session_state.price_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    st.session_state.cost_per_unit,
    st.session_state.opex_pct,
    st.session_state.fixed_overhead,
    st.session_state.starting_cash,
)

pnl_df = build_pnl_view(projection_df)

scorecard = build_scorecard(
    projection_df,
    st.session_state.price_per_unit,
    st.session_state.cost_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
)

warnings = run_rule_based_sanity_check(
    st.session_state.price_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    st.session_state.cost_per_unit,
    st.session_state.opex_pct,
)


# ==================================================
# TOP ACTION BUTTONS
# ==================================================
col_top1, col_top2, _ = st.columns([1.2, 1.6, 4])

with col_top1:
    if st.button("Run VC Sanity Check", key="run_vc_sanity_top"):
        with st.spinner("Running VC sanity check..."):
            sanity_text = run_ai_sanity_check(
                st.session_state.idea,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
            )
            st.session_state.sanity_output = sanity_text
        st.success("VC sanity check generated.")

with col_top2:
    if st.button("Generate Full Business Plan + Deck", key="generate_full_plan_top"):
        with st.spinner("Generating business plan and deck..."):
            plan_text = generate_business_plan_and_deck(
                st.session_state.idea,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
            )
            st.session_state.business_plan_output = plan_text
        st.success("Business plan and deck generated.")


# ==================================================
# DASHBOARD
# ==================================================
year1_revenue = projection_df.loc[0, "Revenue"]
year3_revenue = projection_df.loc[2, "Revenue"]
year3_gross_profit = projection_df.loc[2, "Gross Profit"]
year3_gross_margin = projection_df.loc[2, "Gross Margin %"]
year3_net_income = projection_df.loc[2, "Net Income"]
year3_cash = projection_df.loc[2, "Ending Cash"]

net_income_class = "kpi-green" if year3_net_income >= 0 else "kpi-red"
cash_class = "kpi-green" if year3_cash >= 0 else "kpi-red"

st.markdown("---")
st.subheader("Financial Dashboard")

kpi1, kpi2, kpi3 = st.columns(3)
kpi4, kpi5, kpi6 = st.columns(3)

with kpi1:
    st.markdown(f"""
    <div class="kpi-card kpi-blue">
        <div class="kpi-title">Year 1 Revenue</div>
        <div class="kpi-value">${year1_revenue:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi2:
    st.markdown(f"""
    <div class="kpi-card kpi-green">
        <div class="kpi-title">Year 3 Revenue</div>
        <div class="kpi-value">${year3_revenue:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi3:
    st.markdown(f"""
    <div class="kpi-card kpi-blue">
        <div class="kpi-title">Gross Profit</div>
        <div class="kpi-value">${year3_gross_profit:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi4:
    st.markdown(f"""
    <div class="kpi-card kpi-gold">
        <div class="kpi-title">Gross Margin</div>
        <div class="kpi-value">{year3_gross_margin:.1%}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi5:
    st.markdown(f"""
    <div class="kpi-card {net_income_class}">
        <div class="kpi-title">Net Income</div>
        <div class="kpi-value">${year3_net_income:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

with kpi6:
    st.markdown(f"""
    <div class="kpi-card {cash_class}">
        <div class="kpi-title">Ending Cash</div>
        <div class="kpi-value">${year3_cash:,.0f}</div>
    </div>
    """, unsafe_allow_html=True)

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown("#### Revenue / Net Income / Ending Cash")
    dashboard_chart_df = projection_df.set_index("Year")[["Revenue", "Net Income", "Ending Cash"]]
    st.line_chart(dashboard_chart_df, use_container_width=True)

with chart_col2:
    st.markdown("#### Revenue vs COGS vs Operating Expenses")
    compare_df = projection_df.set_index("Year")[["Revenue", "COGS", "Operating Expenses"]]
    st.bar_chart(compare_df, use_container_width=True)

summary_col1, summary_col2 = st.columns([2, 1])

with summary_col1:
    st.markdown("#### P&L Summary")

    display_pnl = pnl_df.copy()

    for col in ["Year 1", "Year 2", "Year 3"]:
        formatted_values = []
        for _, row in display_pnl.iterrows():
            value = row[col]
            line_item = row["Line Item"]

            if line_item == "Gross Margin %":
                formatted_values.append(f"{value:.1%}")
            elif line_item == "Units":
                formatted_values.append(f"{int(value):,}")
            else:
                formatted_values.append(f"${value:,.0f}")

        display_pnl[col] = formatted_values

    st.dataframe(display_pnl, use_container_width=True, hide_index=True)

with summary_col2:
    st.markdown("#### Investor Snapshot")
    st.markdown(f"""
    <div class="snapshot-card">
        <p><strong>Overall Readiness:</strong> {scorecard['Overall Investor Readiness']}</p>
        <p><strong>Pricing Realism:</strong> {scorecard['Pricing Realism']}</p>
        <p><strong>Sales Volume:</strong> {scorecard['Sales Volume']}</p>
        <p><strong>Growth Assumptions:</strong> {scorecard['Growth Assumptions']}</p>
        <p><strong>Margin Quality:</strong> {scorecard['Margin Quality']}</p>
        <p><strong>Cash Viability:</strong> {scorecard['Cash Viability']}</p>
    </div>
    """, unsafe_allow_html=True)


# ==================================================
# TABS
# ==================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Sanity Check",
    "Business Plan + Deck",
    "Financial Table",
    "Revenue Chart",
    "Downloads",
])

with tab1:
    st.markdown("## Investor Readiness Scorecard")
    for label, status in scorecard.items():
        st.write(f"{score_metric(status)} **{label}:** {status}")

    st.markdown("---")
    st.markdown("### Rule-Based Assumption Check")
    for warning in warnings:
        st.write(warning)

    st.markdown("---")
    st.markdown("### AI VC Sanity Check & Suggestions")

    if st.button("Run AI Sanity Check", key="run_ai_sanity_tab"):
        with st.spinner("Reviewing assumptions..."):
            sanity_text = run_ai_sanity_check(
                st.session_state.idea,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
            )
            st.session_state.sanity_output = sanity_text

    if st.session_state.get("sanity_output"):
        st.markdown(clean_ai_text(st.session_state.sanity_output))
    else:
        st.info("No sanity report yet. Click 'Run VC Sanity Check'.")

with tab2:
    st.markdown("### Full Business Plan + Pitch Deck Content")

    if st.button("Generate Business Plan + Deck", key="generate_plan_tab"):
        with st.spinner("Building investor materials..."):
            plan_text = generate_business_plan_and_deck(
                st.session_state.idea,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
            )
            st.session_state.business_plan_output = plan_text

    if st.session_state.get("business_plan_output"):
        st.markdown(clean_ai_text(st.session_state.business_plan_output))
    else:
        st.info("No business plan yet. Click 'Generate Full Business Plan + Deck'.")

with tab3:
    st.dataframe(display_pnl, use_container_width=True, hide_index=True)

with tab4:
    chart_df = projection_df.set_index("Year")[["Revenue", "Net Income", "Ending Cash"]]
    st.line_chart(chart_df, use_container_width=True)

with tab5:
    st.markdown("### Download Investor Materials")
    st.write("Word = full business plan, Excel = dashboard + financial model + assumptions, PowerPoint = formatted pitch deck with financial chart.")


# ==================================================
# DOWNLOAD MATERIALS
# ==================================================
plan_raw = st.session_state.get("business_plan_output", "")
business_plan_text = extract_business_plan_section(plan_raw)
deck_slides = extract_pitch_deck_section(plan_raw)

# --------------------------
# WORD: FULL BUSINESS PLAN + FINANCIALS + CHART
# --------------------------
doc = Document()
doc.add_heading("Investor Business Plan", 0)

if business_plan_text:
    for paragraph in business_plan_text.split("\n\n"):
        paragraph = paragraph.strip()
        if paragraph:
            doc.add_paragraph(paragraph)
else:
    doc.add_paragraph("No business plan available yet. Generate the full business plan and deck first.")

doc.add_heading("Financial Projections", level=1)

table = doc.add_table(rows=1, cols=len(pnl_df.columns))
table.style = "Table Grid"

hdr_cells = table.rows[0].cells
for i, col in enumerate(pnl_df.columns):
    hdr_cells[i].text = str(col)

for _, row in pnl_df.iterrows():
    row_cells = table.add_row().cells
    for i, value in enumerate(row):
        line_item = row["Line Item"] if "Line Item" in row else ""

        if i == 0:
            row_cells[i].text = str(value)
        else:
            if line_item == "Gross Margin %":
                row_cells[i].text = f"{value:.1%}"
            elif line_item == "Units":
                row_cells[i].text = f"{int(value):,}"
            else:
                row_cells[i].text = f"${value:,.0f}"

doc.add_heading("Revenue Chart", level=1)
chart_image = create_revenue_chart_image(projection_df)
doc.add_picture(chart_image, width=Inches(6.5))

doc_buffer = io.BytesIO()
doc.save(doc_buffer)
doc_buffer.seek(0)

# --------------------------
# EXCEL: DASHBOARD + FINANCIAL MODEL + ASSUMPTIONS
# --------------------------
wb = Workbook()

header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
header_font = Font(color="FFFFFF", bold=True)
section_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
title_font = Font(size=18, bold=True, color="1F1F1F")
section_font = Font(size=12, bold=True, color="1F1F1F")
thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

# Dashboard sheet
ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_dash.sheet_view.showGridLines = False
ws_dash.sheet_view.showRowColHeaders = False

ws_dash.merge_cells("A1:F1")
ws_dash["A1"] = "TurboPitch Financial Dashboard"
ws_dash["A1"].font = title_font
ws_dash["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_dash.merge_cells("A2:F2")
ws_dash["A2"] = "Investor-ready summary of projections, readiness, and performance outlook"
ws_dash["A2"].font = Font(size=10, italic=True, color="666666")
ws_dash["A2"].alignment = Alignment(horizontal="center", vertical="center")

ws_dash.merge_cells("A4:B4")
ws_dash["A4"] = "KPI Summary"
ws_dash["A4"].font = section_font
ws_dash["A4"].fill = section_fill

dashboard_metrics = [
    ("Year 1 Revenue", projection_df.loc[0, "Revenue"]),
    ("Year 3 Revenue", projection_df.loc[2, "Revenue"]),
    ("Gross Profit", projection_df.loc[2, "Gross Profit"]),
    ("Gross Margin %", projection_df.loc[2, "Gross Margin %"]),
    ("Net Income", projection_df.loc[2, "Net Income"]),
    ("Ending Cash", projection_df.loc[2, "Ending Cash"]),
]

for i, (label, value) in enumerate(dashboard_metrics, start=5):
    ws_dash[f"A{i}"] = label
    ws_dash[f"B{i}"] = value
    ws_dash[f"A{i}"].font = Font(bold=True)
    if label == "Gross Margin %":
        ws_dash[f"B{i}"].number_format = "0.0%"
    else:
        ws_dash[f"B{i}"].number_format = '$#,##0'

ws_dash.merge_cells("D4:E4")
ws_dash["D4"] = "Investor Snapshot"
ws_dash["D4"].font = section_font
ws_dash["D4"].fill = section_fill

snapshot_items = [
    ("Overall Readiness", scorecard["Overall Investor Readiness"]),
    ("Pricing Realism", scorecard["Pricing Realism"]),
    ("Sales Volume", scorecard["Sales Volume"]),
    ("Growth Assumptions", scorecard["Growth Assumptions"]),
    ("Margin Quality", scorecard["Margin Quality"]),
    ("Cash Viability", scorecard["Cash Viability"]),
]

for i, (label, value) in enumerate(snapshot_items, start=5):
    ws_dash[f"D{i}"] = label
    ws_dash[f"E{i}"] = value
    ws_dash[f"D{i}"].font = Font(bold=True)

ws_dash.merge_cells("A13:D13")
ws_dash["A13"] = "P&L Summary"
ws_dash["A13"].font = section_font
ws_dash["A13"].fill = section_fill

pnl_headers = list(pnl_df.columns)
for col_idx, col_name in enumerate(pnl_headers, start=1):
    cell = ws_dash.cell(row=14, column=col_idx)
    cell.value = col_name
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row_idx, (_, row) in enumerate(pnl_df.iterrows(), start=15):
    for col_idx, value in enumerate(row, start=1):
        cell = ws_dash.cell(row=row_idx, column=col_idx)
        cell.value = value
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col_idx > 1:
            line_item = row["Line Item"]
            if line_item == "Gross Margin %":
                cell.number_format = "0.0%"
            elif line_item == "Units":
                cell.number_format = '#,##0'
            else:
                cell.number_format = '$#,##0'

ws_dash["H2"] = "Year"
ws_dash["I2"] = "Revenue"
ws_dash["J2"] = "Net Income"
ws_dash["K2"] = "Ending Cash"
ws_dash["L2"] = "COGS"
ws_dash["M2"] = "Operating Expenses"

for idx, row in projection_df.iterrows():
    r = idx + 3
    ws_dash[f"H{r}"] = row["Year"]
    ws_dash[f"I{r}"] = row["Revenue"]
    ws_dash[f"J{r}"] = row["Net Income"]
    ws_dash[f"K{r}"] = row["Ending Cash"]
    ws_dash[f"L{r}"] = row["COGS"]
    ws_dash[f"M{r}"] = row["Operating Expenses"]

ws_dash.merge_cells("H8:M8")
ws_dash["H8"] = "Projection Charts"
ws_dash["H8"].font = section_font
ws_dash["H8"].fill = section_fill

line_chart = LineChart()
line_chart.title = "Revenue / Net Income / Ending Cash"
line_chart.style = 2
line_chart.y_axis.title = "USD"
line_chart.x_axis.title = "Year"
line_data = Reference(ws_dash, min_col=9, max_col=11, min_row=2, max_row=5)
line_cats = Reference(ws_dash, min_col=8, min_row=3, max_row=5)
line_chart.add_data(line_data, titles_from_data=True)
line_chart.set_categories(line_cats)
line_chart.height = 7
line_chart.width = 11
ws_dash.add_chart(line_chart, "H9")

bar_chart = BarChart()
bar_chart.title = "Revenue vs COGS vs Operating Expenses"
bar_chart.style = 10
bar_chart.y_axis.title = "USD"
bar_chart.x_axis.title = "Year"
bar_data = Reference(ws_dash, min_col=9, max_col=9, min_row=2, max_row=5)
bar_data2 = Reference(ws_dash, min_col=12, max_col=13, min_row=2, max_row=5)
bar_cats = Reference(ws_dash, min_col=8, min_row=3, max_row=5)
bar_chart.add_data(bar_data, titles_from_data=True)
bar_chart.add_data(bar_data2, titles_from_data=True)
bar_chart.set_categories(bar_cats)
bar_chart.height = 7
bar_chart.width = 11
ws_dash.add_chart(bar_chart, "H25")

for col in range(1, 14):
    col_letter = get_column_letter(col)
    max_length = 0
    for row in range(1, ws_dash.max_row + 1):
        value = ws_dash[f"{col_letter}{row}"].value
        if value is not None:
            max_length = max(max_length, len(str(value)))
    ws_dash.column_dimensions[col_letter].width = max_length + 3

# Financial Model sheet
ws_model = wb.create_sheet("Financial Model")
headers = list(pnl_df.columns)
ws_model.append(headers)

for _, row in pnl_df.iterrows():
    ws_model.append(list(row))

for cell in ws_model[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

for row in ws_model.iter_rows(min_row=2):
    for cell in row:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")

for row_idx in range(2, ws_model.max_row + 1):
    line_item = ws_model[f"A{row_idx}"].value
    for col_idx in range(2, ws_model.max_column + 1):
        cell = ws_model.cell(row=row_idx, column=col_idx)
        if line_item == "Gross Margin %":
            cell.number_format = '0.0%'
        elif line_item == "Units":
            cell.number_format = '#,##0'
        else:
            cell.number_format = '$#,##0'

for col_idx, col_name in enumerate(headers, start=1):
    col_letter = get_column_letter(col_idx)
    max_length = len(str(col_name))
    for row_idx in range(2, ws_model.max_row + 1):
        value = ws_model[f"{col_letter}{row_idx}"].value
        if value is not None:
            max_length = max(max_length, len(str(value)))
    ws_model.column_dimensions[col_letter].width = max_length + 2

ws_model.freeze_panes = "A2"

# Assumptions sheet
ws_assump = wb.create_sheet("Assumptions")
ws_assump.merge_cells("A1:B1")
ws_assump["A1"] = "Model Assumptions"
ws_assump["A1"].font = title_font
ws_assump["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_assump["A3"] = "Assumption"
ws_assump["B3"] = "Value"

for cell in ws_assump[3]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

assumptions_data = [
    ("Startup Idea", st.session_state.idea if st.session_state.idea else "No startup idea provided."),
    ("Price per Unit", st.session_state.price_per_unit),
    ("Year 1 Units Sold", st.session_state.year1_units),
    ("Year 2 Growth Rate", st.session_state.growth_y2),
    ("Year 3 Growth Rate", st.session_state.growth_y3),
    ("Cost per Unit", st.session_state.cost_per_unit),
    ("Operating Expense %", st.session_state.opex_pct),
    ("Fixed Annual Overhead", st.session_state.fixed_overhead),
    ("Starting Cash", st.session_state.starting_cash),
    ("Investor Pushback %", st.session_state.pushback_pct),
    ("Investor Feedback", st.session_state.investor_feedback),
]

for idx, (label, value) in enumerate(assumptions_data, start=4):
    ws_assump[f"A{idx}"] = label
    ws_assump[f"B{idx}"] = value
    ws_assump[f"A{idx}"].border = thin_border
    ws_assump[f"B{idx}"].border = thin_border
    ws_assump[f"A{idx}"].font = Font(bold=True)
    ws_assump[f"A{idx}"].alignment = Alignment(vertical="top")
    ws_assump[f"B{idx}"].alignment = Alignment(wrap_text=True, vertical="top")

    if label in {"Price per Unit", "Cost per Unit", "Fixed Annual Overhead", "Starting Cash"}:
        ws_assump[f"B{idx}"].number_format = '$#,##0.00'
    elif label in {"Year 2 Growth Rate", "Year 3 Growth Rate", "Operating Expense %"}:
        ws_assump[f"B{idx}"].number_format = '0.0%'
    elif label in {"Year 1 Units Sold"}:
        ws_assump[f"B{idx}"].number_format = '#,##0'

ws_assump.column_dimensions["A"].width = 28
ws_assump.column_dimensions["B"].width = 65
ws_assump.freeze_panes = "A4"

# Formula-driven Financial Model
line_item_rows = {}
for row_idx in range(2, ws_model.max_row + 1):
    line_item_rows[ws_model[f"A{row_idx}"].value] = row_idx

ws_model[f"B{line_item_rows['Units']}"] = "=Assumptions!B6"
ws_model[f"C{line_item_rows['Units']}"] = f"=B{line_item_rows['Units']}*(1+Assumptions!B7)"
ws_model[f"D{line_item_rows['Units']}"] = f"=C{line_item_rows['Units']}*(1+Assumptions!B8)"

ws_model[f"B{line_item_rows['Revenue']}"] = f"=B{line_item_rows['Units']}*Assumptions!B5"
ws_model[f"C{line_item_rows['Revenue']}"] = f"=C{line_item_rows['Units']}*Assumptions!B5"
ws_model[f"D{line_item_rows['Revenue']}"] = f"=D{line_item_rows['Units']}*Assumptions!B5"

ws_model[f"B{line_item_rows['COGS']}"] = f"=B{line_item_rows['Units']}*Assumptions!B9"
ws_model[f"C{line_item_rows['COGS']}"] = f"=C{line_item_rows['Units']}*Assumptions!B9"
ws_model[f"D{line_item_rows['COGS']}"] = f"=D{line_item_rows['Units']}*Assumptions!B9"

ws_model[f"B{line_item_rows['Gross Profit']}"] = f"=B{line_item_rows['Revenue']}-B{line_item_rows['COGS']}"
ws_model[f"C{line_item_rows['Gross Profit']}"] = f"=C{line_item_rows['Revenue']}-C{line_item_rows['COGS']}"
ws_model[f"D{line_item_rows['Gross Profit']}"] = f"=D{line_item_rows['Revenue']}-D{line_item_rows['COGS']}"

ws_model[f"B{line_item_rows['Gross Margin %']}"] = f"=IF(B{line_item_rows['Revenue']}=0,0,B{line_item_rows['Gross Profit']}/B{line_item_rows['Revenue']})"
ws_model[f"C{line_item_rows['Gross Margin %']}"] = f"=IF(C{line_item_rows['Revenue']}=0,0,C{line_item_rows['Gross Profit']}/C{line_item_rows['Revenue']})"
ws_model[f"D{line_item_rows['Gross Margin %']}"] = f"=IF(D{line_item_rows['Revenue']}=0,0,D{line_item_rows['Gross Profit']}/D{line_item_rows['Revenue']})"

ws_model[f"B{line_item_rows['Operating Expenses']}"] = f"=(B{line_item_rows['Revenue']}*Assumptions!B10)+Assumptions!B11"
ws_model[f"C{line_item_rows['Operating Expenses']}"] = f"=(C{line_item_rows['Revenue']}*Assumptions!B10)+Assumptions!B11"
ws_model[f"D{line_item_rows['Operating Expenses']}"] = f"=(D{line_item_rows['Revenue']}*Assumptions!B10)+Assumptions!B11"

ws_model[f"B{line_item_rows['Operating Income']}"] = f"=B{line_item_rows['Gross Profit']}-B{line_item_rows['Operating Expenses']}"
ws_model[f"C{line_item_rows['Operating Income']}"] = f"=C{line_item_rows['Gross Profit']}-C{line_item_rows['Operating Expenses']}"
ws_model[f"D{line_item_rows['Operating Income']}"] = f"=D{line_item_rows['Gross Profit']}-D{line_item_rows['Operating Expenses']}"

ws_model[f"B{line_item_rows['Taxes']}"] = f"=MAX(0,B{line_item_rows['Operating Income']}*21%)"
ws_model[f"C{line_item_rows['Taxes']}"] = f"=MAX(0,C{line_item_rows['Operating Income']}*21%)"
ws_model[f"D{line_item_rows['Taxes']}"] = f"=MAX(0,D{line_item_rows['Operating Income']}*21%)"

ws_model[f"B{line_item_rows['Net Income']}"] = f"=B{line_item_rows['Operating Income']}-B{line_item_rows['Taxes']}"
ws_model[f"C{line_item_rows['Net Income']}"] = f"=C{line_item_rows['Operating Income']}-C{line_item_rows['Taxes']}"
ws_model[f"D{line_item_rows['Net Income']}"] = f"=D{line_item_rows['Operating Income']}-D{line_item_rows['Taxes']}"

ws_model[f"B{line_item_rows['Ending Cash']}"] = f"=Assumptions!B12+B{line_item_rows['Net Income']}"
ws_model[f"C{line_item_rows['Ending Cash']}"] = f"=B{line_item_rows['Ending Cash']}+C{line_item_rows['Net Income']}"
ws_model[f"D{line_item_rows['Ending Cash']}"] = f"=C{line_item_rows['Ending Cash']}+D{line_item_rows['Net Income']}"

excel_buffer = io.BytesIO()
wb.save(excel_buffer)
excel_buffer.seek(0)

# PowerPoint
prs = Presentation()

slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "Investor Pitch Deck"
slide.placeholders[1].text = st.session_state.idea if st.session_state.idea else "Startup concept"

title = slide.shapes.title.text_frame.paragraphs[0]
title.font.size = Pt(28)
title.font.bold = True
title.font.color.rgb = RGBColor(31, 78, 121)

subtitle = slide.placeholders[1].text_frame.paragraphs[0]
subtitle.font.size = Pt(16)
subtitle.font.color.rgb = RGBColor(89, 89, 89)

for title_text, bullets in deck_slides:
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = clean_ai_text(title_text)

    title_paragraph = slide.shapes.title.text_frame.paragraphs[0]
    title_paragraph.font.size = Pt(24)
    title_paragraph.font.bold = True
    title_paragraph.font.color.rgb = RGBColor(31, 78, 121)

    tf = slide.placeholders[1].text_frame
    tf.clear()

    bullet_lines = [clean_ai_text(b) for b in bullets if clean_ai_text(b)]
    if not bullet_lines:
        bullet_lines = ["No content generated for this slide."]

    first = True
    for bullet in bullet_lines[:4]:
        if first:
            tf.text = bullet
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
            p.text = bullet

        p.level = 0
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(68, 68, 68)
        p.alignment = PP_ALIGN.LEFT

slide = prs.slides.add_slide(prs.slide_layouts[1])
slide.shapes.title.text = "Financial Highlights"

title_paragraph = slide.shapes.title.text_frame.paragraphs[0]
title_paragraph.font.size = Pt(24)
title_paragraph.font.bold = True
title_paragraph.font.color.rgb = RGBColor(31, 78, 121)

tf = slide.placeholders[1].text_frame
tf.clear()
tf.text = f"Year 1 Revenue: ${projection_df['Revenue'].iloc[0]:,.0f}"
p = tf.paragraphs[0]
p.font.size = Pt(20)
p.font.color.rgb = RGBColor(68, 68, 68)

for line in [
    f"Year 3 Revenue: ${projection_df['Revenue'].iloc[2]:,.0f}",
    f"Year 3 Net Income: ${projection_df['Net Income'].iloc[2]:,.0f}",
    f"Year 3 Ending Cash: ${projection_df['Ending Cash'].iloc[2]:,.0f}",
]:
    p = tf.add_paragraph()
    p.text = line
    p.level = 0
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(68, 68, 68)

slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "Financial Projection Chart"

title_paragraph = slide.shapes.title.text_frame.paragraphs[0]
title_paragraph.font.size = Pt(24)
title_paragraph.font.bold = True
title_paragraph.font.color.rgb = RGBColor(31, 78, 121)

chart_img = create_ppt_financial_chart_image(projection_df)
slide.shapes.add_picture(chart_img, PPTInches(0.7), PPTInches(1.4), width=PPTInches(8.2))

for i in range(1, len(prs.slides)):
    s = prs.slides[i]
    textbox = s.shapes.add_textbox(PPTInches(0.3), PPTInches(6.8), PPTInches(2.0), PPTInches(0.3))
    tf = textbox.text_frame
    tf.text = "TurboPitch"
    p = tf.paragraphs[0]
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(120, 120, 120)

ppt_buffer = io.BytesIO()
prs.save(ppt_buffer)
ppt_buffer.seek(0)

# Downloads
st.markdown("---")
st.subheader("Download Investor Materials")

col1, col2, col3 = st.columns(3)

with col1:
    st.download_button(
        label="Download Word Business Plan",
        data=doc_buffer,
        file_name="business_plan.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

with col2:
    st.download_button(
        label="Download Excel Financial Package",
        data=excel_buffer,
        file_name="financial_package.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col3:
    st.download_button(
        label="Download PowerPoint Deck",
        data=ppt_buffer,
        file_name="pitch_deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
