import os
import io
import re

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openai import OpenAI
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
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
st.set_page_config(page_title="TurboPitch", layout="wide")


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

st.title("TurboPitch")
st.subheader("Turn startup ideas into investor-ready materials — then pressure-test them through a VC lens.")
st.caption(
    "TurboPitch combines founder inputs, structured financial modeling, benchmark logic, market reality logic, "
    "and AI reasoning to evaluate startup assumptions through an investor-readiness lens."
)


# ==================================================
# SESSION STATE DEFAULTS
# ==================================================
defaults = {
    "idea": "",
    "industry": "SaaS",
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
    "interrogation_output": "",
    "answer_builder_output": "",
    "assumption_helper_output": "",
    "assumption_mode": "Manual",
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ==================================================
# BENCHMARK DATA
# ==================================================
INDUSTRY_BENCHMARKS = {
    "SaaS": {
        "gross_margin": (0.70, 0.90),
        "growth_y2": (0.50, 1.50),
        "growth_y3": (0.30, 1.00),
        "opex_pct": (0.40, 0.70),
        "year1_units": (1000, 50000),
    },
    "Marketplace": {
        "gross_margin": (0.40, 0.70),
        "growth_y2": (0.50, 1.20),
        "growth_y3": (0.30, 0.90),
        "opex_pct": (0.35, 0.60),
        "year1_units": (5000, 100000),
    },
    "Consumer Product": {
        "gross_margin": (0.40, 0.60),
        "growth_y2": (0.30, 0.70),
        "growth_y3": (0.20, 0.50),
        "opex_pct": (0.25, 0.50),
        "year1_units": (10000, 150000),
    },
    "Food / Delivery": {
        "gross_margin": (0.20, 0.35),
        "growth_y2": (0.30, 0.90),
        "growth_y3": (0.20, 0.60),
        "opex_pct": (0.30, 0.50),
        "year1_units": (10000, 120000),
    },
    "AI Startup": {
        "gross_margin": (0.50, 0.80),
        "growth_y2": (0.50, 1.50),
        "growth_y3": (0.30, 1.00),
        "opex_pct": (0.40, 0.75),
        "year1_units": (1000, 75000),
    },
}


# ==================================================
# TRUST / METHODOLOGY CONTENT
# ==================================================
TRUST_DATA_SOURCES = [
    {
        "category": "Founder Inputs",
        "examples": "Idea description, industry choice, business model hints, manual assumptions",
        "use_case": "Used as the base context for all generated outputs and assumption logic."
    },
    {
        "category": "Internal Benchmark Ranges",
        "examples": "Gross margin bands, growth ranges, opex ranges, Year 1 unit ranges by industry",
        "use_case": "Used to build directional starting assumptions and compare founder inputs to typical business model patterns."
    },
    {
        "category": "Pricing / Business Model Heuristics",
        "examples": "Higher pricing for B2B / enterprise / AI software, lower pricing for food, retail, consumer, and job-seeker concepts",
        "use_case": "Used to adapt suggested assumptions based on keywords, customer type, and business model."
    },
    {
        "category": "Reality Engine Logic",
        "examples": "Customer segment checks, pricing fit checks, adoption realism checks, enterprise sales friction checks",
        "use_case": "Used to challenge assumptions that may be mathematically clean but unrealistic in the real world."
    },
    {
        "category": "Financial Modeling Logic",
        "examples": "Revenue = price × units, COGS derived from target margin, opex as % of revenue plus fixed overhead, tax and ending cash flow logic",
        "use_case": "Used to translate assumptions into a structured 3-year model."
    },
    {
        "category": "AI Interpretation",
        "examples": "Investor-readiness analysis, founder answer prep, suggested assumption rationale",
        "use_case": "Used to explain what the numbers mean, where investors may push back, and what should be validated next."
    }
]

TRUST_LIMITATIONS = [
    "TurboPitch is a decision-support tool, not a guarantee of startup success or funding.",
    "Outputs depend heavily on the quality and realism of founder inputs.",
    "Benchmark ranges are directional and should support judgment, not replace it.",
    "Some niche industries may require more custom market research and expert review.",
    "Reality Engine checks are heuristics and are meant to improve realism, not replace live market research.",
    "AI commentary is generated from structured assumptions, benchmark logic, market reality heuristics, and financial reasoning, but should still be reviewed by the founder."
]

BENCHMARK_SOURCE_LABELS = {
    "gross_margin": "Benchmark basis: internal industry margin ranges and business model heuristics.",
    "growth": "Benchmark basis: internal startup growth ranges and investor-readiness scaling heuristics.",
    "opex": "Benchmark basis: internal cost structure ranges and early-stage operating model heuristics.",
    "units": "Benchmark basis: internal Year 1 traction ranges and go-to-market feasibility assumptions."
}


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


def currency_tick_formatter(x, pos):
    if abs(x) >= 1_000_000:
        return f"${x/1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"${x/1_000:.0f}K"
    return f"${x:,.0f}"


def create_revenue_chart_image(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    ax.plot(df["Year"], df["Revenue"], marker="o", linewidth=2.5)
    ax.set_title("Revenue Projection Summary", fontsize=16, fontweight="bold", pad=14)
    ax.set_xlabel("Year", fontsize=11)
    ax.set_ylabel("Revenue ($)", fontsize=11)
    ax.yaxis.set_major_formatter(FuncFormatter(currency_tick_formatter))
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
    ax.yaxis.set_major_formatter(FuncFormatter(currency_tick_formatter))
    ax.grid(True, alpha=0.3)
    ax.legend()

    img_stream = io.BytesIO()
    fig.tight_layout()
    fig.savefig(img_stream, format="png", dpi=200)
    img_stream.seek(0)
    plt.close(fig)
    return img_stream


def create_excel_projection_chart_image(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(9.5, 4.8))

    ax.plot(df["Year"], df["Revenue"], marker="o", linewidth=2.8, label="Revenue")
    ax.plot(df["Year"], df["Net Income"], marker="o", linewidth=2.8, label="Net Income")
    ax.plot(df["Year"], df["Ending Cash"], marker="o", linewidth=2.8, label="Ending Cash")

    ax.set_title("Revenue / Net Income / Ending Cash", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("Year", fontsize=10)
    ax.set_ylabel("USD", fontsize=10)
    ax.yaxis.set_major_formatter(FuncFormatter(currency_tick_formatter))
    ax.grid(True, axis="y", alpha=0.25)
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=3,
        frameon=False,
        fontsize=9
    )

    img_stream = io.BytesIO()
    fig.tight_layout(rect=[0.03, 0.08, 0.98, 0.95])
    fig.savefig(img_stream, format="png", dpi=240, bbox_inches="tight", facecolor="white")
    img_stream.seek(0)
    plt.close(fig)
    return img_stream


def create_excel_compare_chart_image(df: pd.DataFrame):
    fig, ax = plt.subplots(figsize=(9.5, 4.8))

    x = range(len(df["Year"]))
    width = 0.23

    ax.bar([i - width for i in x], df["Revenue"], width=width, label="Revenue")
    ax.bar(x, df["COGS"], width=width, label="COGS")
    ax.bar([i + width for i in x], df["Operating Expenses"], width=width, label="Operating Expenses")

    ax.set_xticks(list(x))
    ax.set_xticklabels(df["Year"])
    ax.set_title("Revenue vs COGS vs Operating Expenses", fontsize=13, fontweight="bold", pad=14)
    ax.set_xlabel("Year", fontsize=10)
    ax.set_ylabel("USD", fontsize=10)
    ax.yaxis.set_major_formatter(FuncFormatter(currency_tick_formatter))
    ax.grid(True, axis="y", alpha=0.25)
    ax.set_facecolor("white")
    fig.patch.set_facecolor("white")

    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    ax.legend(
        loc="upper center",
        bbox_to_anchor=(0.5, -0.18),
        ncol=3,
        frameon=False,
        fontsize=9
    )

    img_stream = io.BytesIO()
    fig.tight_layout(rect=[0.03, 0.08, 0.98, 0.95])
    fig.savefig(img_stream, format="png", dpi=240, bbox_inches="tight", facecolor="white")
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


def build_display_pnl(pnl_df: pd.DataFrame) -> pd.DataFrame:
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

    return display_pnl


# ==================================================
# REALITY ENGINE
# ==================================================
def detect_customer_segment(idea: str, industry: str) -> str:
    idea_lower = (idea or "").lower()

    enterprise_terms = [
        "enterprise", "b2b", "team", "company", "companies", "businesses",
        "workflow", "salesforce", "crm", "procurement", "operations",
        "analytics platform", "api", "infrastructure", "compliance", "hr software"
    ]
    consumer_terms = [
        "consumer", "job seeker", "resume", "dating", "fitness", "recipe", "pet owner",
        "student", "students", "creator", "freelancer", "marketplace app", "social app",
        "personal finance app", "subscription box", "food delivery", "local service"
    ]

    enterprise_hits = sum(1 for term in enterprise_terms if term in idea_lower)
    consumer_hits = sum(1 for term in consumer_terms if term in idea_lower)

    if enterprise_hits > consumer_hits:
        return "B2B / Enterprise"
    if consumer_hits > enterprise_hits:
        return "B2C / Consumer"

    if industry in ["Consumer Product", "Food / Delivery"]:
        return "B2C / Consumer"
    if industry in ["SaaS", "AI Startup"]:
        return "Mixed / Unclear"
    return "Mixed / Unclear"


def pricing_market_check(idea: str, industry: str, price: float) -> dict:
    idea_lower = (idea or "").lower()
    segment = detect_customer_segment(idea, industry)

    status = "Green"
    message = "Pricing appears broadly reasonable for the apparent customer segment."

    if segment == "B2C / Consumer":
        if any(word in idea_lower for word in ["resume", "job", "career", "job seeker"]):
            if price > 100:
                status = "Red"
                message = (
                    f"Price of ${price:,.2f} looks too high for a job-seeker / resume tool. "
                    "This audience is usually price-sensitive, so adoption may be far lower than the model assumes."
                )
            elif price > 50:
                status = "Yellow"
                message = (
                    f"Price of ${price:,.2f} may be high for a job-seeker / resume tool. "
                    "You may need stronger proof of premium value or a lower entry tier."
                )
        elif industry == "Food / Delivery":
            if price > 40:
                status = "Red"
                message = (
                    f"Price of ${price:,.2f} looks too high for a food / delivery concept unless it is premium or enterprise-driven."
                )
            elif price > 25:
                status = "Yellow"
                message = (
                    f"Price of ${price:,.2f} may be above what many food / delivery customers expect."
                )
        elif industry == "Consumer Product":
            if price > 150:
                status = "Red"
                message = (
                    f"Price of ${price:,.2f} may be too high for a consumer product unless you can clearly justify premium positioning."
                )
            elif price > 80:
                status = "Yellow"
                message = (
                    f"Price of ${price:,.2f} may limit adoption unless the product has strong perceived value."
                )
        else:
            if price > 100:
                status = "Red"
                message = (
                    f"Price of ${price:,.2f} looks high for a consumer-facing concept and may suppress demand."
                )
            elif price > 50:
                status = "Yellow"
                message = (
                    f"Price of ${price:,.2f} may be on the high side for a consumer-facing concept."
                )

    elif segment == "B2B / Enterprise":
        if price < 20 and industry in ["SaaS", "AI Startup"]:
            status = "Yellow"
            message = (
                f"Price of ${price:,.2f} may be too low for a B2B / enterprise software concept and could understate value."
            )
        elif price > 5000:
            status = "Yellow"
            message = (
                f"Price of ${price:,.2f} is very high and may require an enterprise sales motion, proof of ROI, and long sales cycles."
            )
    else:
        if industry in ["AI Startup", "SaaS"] and price > 300:
            status = "Yellow"
            message = (
                f"Price of ${price:,.2f} may be reasonable for some B2B software models, but the concept does not clearly prove enterprise willingness to pay."
            )

    return {
        "status": status,
        "message": message,
        "segment": segment,
    }


def volume_market_check(idea: str, industry: str, year1_units: int) -> dict:
    segment = detect_customer_segment(idea, industry)

    status = "Green"
    message = "Initial Year 1 adoption looks directionally reasonable."

    if segment == "B2C / Consumer":
        if year1_units > 250000:
            status = "Red"
            message = (
                f"Year 1 unit target of {year1_units:,} looks extremely aggressive for a consumer startup without proven distribution."
            )
        elif year1_units > 100000:
            status = "Yellow"
            message = (
                f"Year 1 unit target of {year1_units:,} is ambitious for a consumer startup and likely requires major marketing or partnerships."
            )
    elif segment == "B2B / Enterprise":
        if year1_units > 20000:
            status = "Red"
            message = (
                f"Year 1 unit target of {year1_units:,} looks too high for a B2B / enterprise concept unless units represent very small transactions rather than customers."
            )
        elif year1_units > 5000:
            status = "Yellow"
            message = (
                f"Year 1 unit target of {year1_units:,} may be aggressive for a B2B / enterprise concept, especially if each unit implies a customer contract."
            )
    else:
        if year1_units > 150000:
            status = "Yellow"
            message = (
                f"Year 1 unit target of {year1_units:,} may be too aggressive unless customer acquisition is unusually efficient."
            )

    return {
        "status": status,
        "message": message,
    }


def growth_market_check(idea: str, industry: str, growth_y2: float, growth_y3: float) -> dict:
    segment = detect_customer_segment(idea, industry)

    status = "Green"
    message = "Growth path appears directionally plausible for an early-stage startup."

    if segment == "B2B / Enterprise":
        if growth_y2 > 1.0 or growth_y3 > 0.8:
            status = "Yellow"
            message = (
                f"Growth rates of {growth_y2:.0%} in Year 2 and {growth_y3:.0%} in Year 3 may be aggressive for a B2B / enterprise motion with longer sales cycles."
            )
    else:
        if growth_y2 > 1.5 or growth_y3 > 1.0:
            status = "Red"
            message = (
                f"Growth rates of {growth_y2:.0%} in Year 2 and {growth_y3:.0%} in Year 3 look extremely aggressive and may be difficult to support."
            )
        elif growth_y2 > 1.0 or growth_y3 > 0.8:
            status = "Yellow"
            message = (
                f"Growth rates of {growth_y2:.0%} in Year 2 and {growth_y3:.0%} in Year 3 are ambitious and will require strong proof of traction."
            )

    return {
        "status": status,
        "message": message,
    }


def opex_reality_check(industry: str, opex_pct: float, fixed_overhead: float, segment: str) -> dict:
    status = "Green"
    message = "Operating expense assumptions appear directionally reasonable."

    if segment == "B2C / Consumer":
        if opex_pct < 0.20:
            status = "Yellow"
            message = (
                f"Opex ratio of {opex_pct:.0%} may be too low for a consumer startup that likely needs paid acquisition, support, and brand-building."
            )
    elif segment == "B2B / Enterprise":
        if opex_pct < 0.25:
            status = "Yellow"
            message = (
                f"Opex ratio of {opex_pct:.0%} may be too low for a B2B / enterprise model that may require sales, onboarding, and account support."
            )

    if fixed_overhead < 100000 and industry in ["AI Startup", "SaaS"]:
        status = "Yellow"
        message = (
            f"Fixed overhead of ${fixed_overhead:,.0f} may be too low for a software or AI startup once salaries, tooling, and infrastructure are considered."
        )

    return {
        "status": status,
        "message": message,
    }


def run_reality_engine(idea, industry, price, year1_units, growth_y2, growth_y3, opex_pct, fixed_overhead):
    pricing_check = pricing_market_check(idea, industry, price)
    volume_check = volume_market_check(idea, industry, year1_units)
    growth_check = growth_market_check(idea, industry, growth_y2, growth_y3)
    opex_check = opex_reality_check(industry, opex_pct, fixed_overhead, pricing_check["segment"])

    checks = {
        "Customer Segment Detection": {
            "status": "Green",
            "message": f"Detected customer segment: {pricing_check['segment']}"
        },
        "Pricing Market Fit": {
            "status": pricing_check["status"],
            "message": pricing_check["message"]
        },
        "Adoption Realism": {
            "status": volume_check["status"],
            "message": volume_check["message"]
        },
        "Growth Realism": {
            "status": growth_check["status"],
            "message": growth_check["message"]
        },
        "Operating Model Reality": {
            "status": opex_check["status"],
            "message": opex_check["message"]
        },
    }

    reds = sum(1 for item in checks.values() if item["status"] == "Red")
    yellows = sum(1 for item in checks.values() if item["status"] == "Yellow")

    if reds >= 2:
        overall = "Red"
        summary = "Reality Engine sees multiple real-world adoption or pricing issues."
    elif reds == 1 or yellows >= 2:
        overall = "Yellow"
        summary = "Reality Engine sees some real-world issues that may weaken investor credibility."
    else:
        overall = "Green"
        summary = "Reality Engine sees no major market realism issues at a high level."

    return {
        "checks": checks,
        "overall": overall,
        "summary": summary,
    }


def reality_status_icon(status: str) -> str:
    if status == "Green":
        return "🟢"
    if status == "Yellow":
        return "🟡"
    return "🔴"


# ==================================================
# RULE-BASED / SCORECARD LOGIC
# ==================================================
def run_rule_based_sanity_check(price, year1_units, growth_y2, growth_y3, cost_per_unit, opex_pct):
    warnings = []

    if price <= 0:
        warnings.append("🔴 Price per unit must be greater than zero.")
        return warnings

    gross_margin = (price - cost_per_unit) / price if price else 0

    if price < cost_per_unit:
        warnings.append("🔴 Price per unit is below cost per unit. This creates a structurally unprofitable business.")
    elif gross_margin < 0.20:
        warnings.append("🟠 Gross margin is below 20%. Investors usually expect stronger margins for scalable startups.")
    elif gross_margin < 0.40:
        warnings.append("🟡 Gross margin is moderate. Investors may ask how margins improve as the company scales.")

    if year1_units > 300000:
        warnings.append("🔴 Year 1 sales volume is extremely high for an early-stage startup and may be unrealistic.")
    elif year1_units > 150000:
        warnings.append("🟠 Year 1 sales are ambitious. Investors may ask for proof of early demand or strong distribution.")

    if growth_y2 > 1.5:
        warnings.append("🔴 Year 2 growth above 150% is very aggressive and may be difficult to sustain.")
    elif growth_y2 > 1.0:
        warnings.append("🟠 Year 2 growth above 100% will likely require strong evidence of market demand.")

    if growth_y3 > 1.2:
        warnings.append("🟠 Year 3 growth above 120% suggests very rapid scaling and may be questioned by investors.")

    if opex_pct > 0.70:
        warnings.append("🔴 Operating expenses exceed 70% of revenue, which may make profitability difficult.")
    elif opex_pct > 0.50:
        warnings.append("🟠 Operating expenses are high relative to revenue. Investors may ask how efficiency improves over time.")

    if gross_margin < 0.30 and opex_pct > 0.40:
        warnings.append("🔴 Low gross margin combined with high operating costs may make this model difficult to scale profitably.")

    if year1_units > 150000 and growth_y2 > 1.0:
        warnings.append("🟠 High initial sales combined with very fast growth may raise investor skepticism about demand assumptions.")

    if not warnings:
        warnings.append("🟢 No major structural issues detected in the assumptions. The model appears reasonable at a high level.")

    return warnings


def build_warning_summary(warnings):
    red_count = sum(1 for w in warnings if w.startswith("🔴"))
    amber_count = sum(1 for w in warnings if w.startswith("🟠") or w.startswith("🟡"))
    green_count = sum(1 for w in warnings if w.startswith("🟢"))

    if red_count >= 2:
        overall = "High Assumption Risk"
        icon = "🔴"
    elif red_count == 1 or amber_count >= 2:
        overall = "Moderate Assumption Risk"
        icon = "🟠"
    elif green_count > 0 and red_count == 0 and amber_count == 0:
        overall = "Low Assumption Risk"
        icon = "🟢"
    else:
        overall = "Watchlist"
        icon = "🟡"

    return {
        "icon": icon,
        "overall": overall,
        "red_count": red_count,
        "amber_count": amber_count,
    }


def build_scorecard(idea, industry, projection_df, price, cost_per_unit, year1_units, growth_y2, growth_y3, reality_engine_output):
    ending_cash_final = projection_df["Ending Cash"].iloc[-1]
    gross_margin = ((price - cost_per_unit) / price) if price else 0

    structural_pricing_status = "Green" if price >= cost_per_unit * 1.5 else "Yellow" if price >= cost_per_unit else "Red"
    market_pricing_status = reality_engine_output["checks"]["Pricing Market Fit"]["status"]

    if structural_pricing_status == "Red" or market_pricing_status == "Red":
        pricing_status = "Red"
    elif structural_pricing_status == "Yellow" or market_pricing_status == "Yellow":
        pricing_status = "Yellow"
    else:
        pricing_status = "Green"

    sales_status = reality_engine_output["checks"]["Adoption Realism"]["status"]

    growth_status = "Green" if growth_y2 <= 0.75 and growth_y3 <= 0.60 else "Yellow" if growth_y2 <= 1.0 and growth_y3 <= 0.80 else "Red"
    reality_growth = reality_engine_output["checks"]["Growth Realism"]["status"]
    if "Red" in [growth_status, reality_growth]:
        growth_status = "Red"
    elif "Yellow" in [growth_status, reality_growth]:
        growth_status = "Yellow"
    else:
        growth_status = "Green"

    margin_status = "Green" if gross_margin >= 0.50 else "Yellow" if gross_margin >= 0.25 else "Red"
    cash_status = "Green" if ending_cash_final > 0 else "Red"
    operating_model_status = reality_engine_output["checks"]["Operating Model Reality"]["status"]

    statuses = [pricing_status, sales_status, growth_status, margin_status, cash_status, operating_model_status]

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
        "Operating Model Reality": operating_model_status,
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


def build_benchmark_feedback(industry, price, cost_per_unit, year1_units, growth_y2, growth_y3, opex_pct):
    benchmark = INDUSTRY_BENCHMARKS.get(industry)
    if not benchmark:
        return ["No benchmark data available for this industry."]

    feedback = []

    gross_margin = ((price - cost_per_unit) / price) if price > 0 else 0

    gm_low, gm_high = benchmark["gross_margin"]
    y2_low, y2_high = benchmark["growth_y2"]
    y3_low, y3_high = benchmark["growth_y3"]
    opex_low, opex_high = benchmark["opex_pct"]
    unit_low, unit_high = benchmark["year1_units"]

    if gross_margin < gm_low:
        feedback.append(
            f"🔴 Gross margin is {gross_margin:.1%}, below the typical {industry} benchmark range of {gm_low:.0%}–{gm_high:.0%}. "
            f"Investors may question pricing power or scalability.\n{BENCHMARK_SOURCE_LABELS['gross_margin']}"
        )
    elif gross_margin > gm_high:
        feedback.append(
            f"🟢 Gross margin is {gross_margin:.1%}, above the typical {industry} benchmark range of {gm_low:.0%}–{gm_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['gross_margin']}"
        )
    else:
        feedback.append(
            f"🟢 Gross margin is {gross_margin:.1%}, within the typical {industry} benchmark range of {gm_low:.0%}–{gm_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['gross_margin']}"
        )

    if growth_y2 > y2_high:
        feedback.append(
            f"🟠 Year 2 growth of {growth_y2:.0%} is above the usual {industry} benchmark range of {y2_low:.0%}–{y2_high:.0%}. "
            f"Investors may expect stronger proof of traction.\n{BENCHMARK_SOURCE_LABELS['growth']}"
        )
    elif growth_y2 < y2_low:
        feedback.append(
            f"🟡 Year 2 growth of {growth_y2:.0%} is below the usual {industry} benchmark range of {y2_low:.0%}–{y2_high:.0%}. "
            f"This may look conservative, but could also limit investor excitement.\n{BENCHMARK_SOURCE_LABELS['growth']}"
        )
    else:
        feedback.append(
            f"🟢 Year 2 growth of {growth_y2:.0%} is within the typical {industry} benchmark range of {y2_low:.0%}–{y2_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['growth']}"
        )

    if growth_y3 > y3_high:
        feedback.append(
            f"🟠 Year 3 growth of {growth_y3:.0%} is above the usual {industry} benchmark range of {y3_low:.0%}–{y3_high:.0%}. "
            f"Investors may view this as aggressive unless supported by strong momentum.\n{BENCHMARK_SOURCE_LABELS['growth']}"
        )
    elif growth_y3 < y3_low:
        feedback.append(
            f"🟡 Year 3 growth of {growth_y3:.0%} is below the usual {industry} benchmark range of {y3_low:.0%}–{y3_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['growth']}"
        )
    else:
        feedback.append(
            f"🟢 Year 3 growth of {growth_y3:.0%} is within the typical {industry} benchmark range of {y3_low:.0%}–{y3_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['growth']}"
        )

    if opex_pct > opex_high:
        feedback.append(
            f"🔴 Operating expense ratio of {opex_pct:.0%} is above the typical {industry} benchmark range of {opex_low:.0%}–{opex_high:.0%}. "
            f"Investors may question efficiency and burn discipline.\n{BENCHMARK_SOURCE_LABELS['opex']}"
        )
    elif opex_pct < opex_low:
        feedback.append(
            f"🟡 Operating expense ratio of {opex_pct:.0%} is below the typical {industry} benchmark range of {opex_low:.0%}–{opex_high:.0%}. "
            f"This may look efficient, but investors may ask whether growth investment is too light.\n{BENCHMARK_SOURCE_LABELS['opex']}"
        )
    else:
        feedback.append(
            f"🟢 Operating expense ratio of {opex_pct:.0%} is within the typical {industry} benchmark range of {opex_low:.0%}–{opex_high:.0%}.\n"
            f"{BENCHMARK_SOURCE_LABELS['opex']}"
        )

    if year1_units > unit_high:
        feedback.append(
            f"🟠 Year 1 volume of {year1_units:,} is above the typical {industry} benchmark range of {unit_low:,}–{unit_high:,}. "
            f"Investors may ask for stronger evidence of demand and distribution capacity.\n{BENCHMARK_SOURCE_LABELS['units']}"
        )
    elif year1_units < unit_low:
        feedback.append(
            f"🟡 Year 1 volume of {year1_units:,} is below the typical {industry} benchmark range of {unit_low:,}–{unit_high:,}.\n"
            f"{BENCHMARK_SOURCE_LABELS['units']}"
        )
    else:
        feedback.append(
            f"🟢 Year 1 volume of {year1_units:,} is within the typical {industry} benchmark range of {unit_low:,}–{unit_high:,}.\n"
            f"{BENCHMARK_SOURCE_LABELS['units']}"
        )

    suggestions = []

    if gross_margin < gm_low:
        target_price = cost_per_unit / (1 - gm_low) if gm_low < 1 else price
        target_cogs = price * (1 - gm_low)
        suggestions.append(
            f"Suggested fix: To reach the low end of the benchmark gross margin ({gm_low:.0%}), price may need to increase to about ${target_price:,.2f}, "
            f"or COGS may need to decrease to about ${target_cogs:,.2f} at the current price."
        )

    if growth_y2 > y2_high:
        suggestions.append(
            f"Suggested fix: Consider lowering Year 2 growth toward the benchmark range, such as {y2_low:.0%}–{y2_high:.0%}."
        )

    if growth_y3 > y3_high:
        suggestions.append(
            f"Suggested fix: Consider lowering Year 3 growth toward the benchmark range, such as {y3_low:.0%}–{y3_high:.0%}."
        )

    if opex_pct > opex_high:
        suggestions.append(
            f"Suggested fix: Consider reducing operating expenses toward the benchmark range of {opex_low:.0%}–{opex_high:.0%} of revenue."
        )

    if year1_units > unit_high:
        suggestions.append(
            f"Suggested fix: Consider reducing Year 1 units toward a more supportable range, such as {unit_low:,}–{unit_high:,}."
        )

    if not suggestions:
        suggestions.append("Current assumptions are broadly aligned with this industry's benchmark ranges.")

    return feedback + ["", "Suggested Adjustments"] + suggestions


# ==================================================
# ASSUMPTION GENERATOR
# ==================================================
def generate_rule_based_assumptions(industry, idea_text=""):
    benchmark = INDUSTRY_BENCHMARKS.get(industry, INDUSTRY_BENCHMARKS["SaaS"])

    gm_low, gm_high = benchmark["gross_margin"]
    y2_low, y2_high = benchmark["growth_y2"]
    y3_low, y3_high = benchmark["growth_y3"]
    opex_low, opex_high = benchmark["opex_pct"]
    unit_low, unit_high = benchmark["year1_units"]

    target_gross_margin = round((gm_low + gm_high) / 2, 2)
    suggested_y2 = round((y2_low + y2_high) / 2, 2)
    suggested_y3 = round((y3_low + y3_high) / 2, 2)
    suggested_opex = round((opex_low + opex_high) / 2, 2)
    suggested_units = int((unit_low + unit_high) / 2)

    if industry == "SaaS":
        suggested_price = 99.0
        suggested_fixed_overhead = 250000.0
        suggested_starting_cash = 500000.0
    elif industry == "Marketplace":
        suggested_price = 35.0
        suggested_fixed_overhead = 300000.0
        suggested_starting_cash = 600000.0
    elif industry == "Consumer Product":
        suggested_price = 45.0
        suggested_fixed_overhead = 350000.0
        suggested_starting_cash = 500000.0
    elif industry == "Food / Delivery":
        suggested_price = 18.0
        suggested_fixed_overhead = 280000.0
        suggested_starting_cash = 400000.0
    elif industry == "AI Startup":
        suggested_price = 299.0
        suggested_fixed_overhead = 350000.0
        suggested_starting_cash = 700000.0
    else:
        suggested_price = 99.0
        suggested_fixed_overhead = 300000.0
        suggested_starting_cash = 500000.0

    idea_lower = (idea_text or "").lower()
    segment = detect_customer_segment(idea_text, industry)

    if any(word in idea_lower for word in ["enterprise", "b2b", "software", "platform", "analytics", "api", "workflow"]):
        suggested_price *= 1.35
        suggested_fixed_overhead *= 1.15

    if any(word in idea_lower for word in ["consumer", "app", "marketplace", "delivery", "subscription box"]):
        suggested_units = int(suggested_units * 1.20)

    if any(word in idea_lower for word in ["restaurant", "retail", "food", "cpg", "product"]):
        suggested_price *= 0.80

    if any(word in idea_lower for word in ["resume", "job", "career", "job seeker", "student", "consumer app"]):
        suggested_price = min(suggested_price, 39.0)
        suggested_units = int(suggested_units * 1.15)

    if segment == "B2C / Consumer" and industry in ["AI Startup", "SaaS"]:
        suggested_price = min(suggested_price, 49.0)

    if segment == "B2B / Enterprise" and suggested_price < 49:
        suggested_price = 49.0

    suggested_price = round(suggested_price, 2)
    suggested_cost = round(suggested_price * (1 - target_gross_margin), 2)
    suggested_fixed_overhead = round(suggested_fixed_overhead, 0)
    suggested_starting_cash = round(suggested_starting_cash, 0)

    return {
        "price_per_unit": suggested_price,
        "year1_units": suggested_units,
        "growth_y2": suggested_y2,
        "growth_y3": suggested_y3,
        "cost_per_unit": suggested_cost,
        "opex_pct": suggested_opex,
        "fixed_overhead": suggested_fixed_overhead,
        "starting_cash": suggested_starting_cash,
        "target_gross_margin": target_gross_margin,
        "detected_segment": segment,
    }


def build_reality_engine_explanation(idea, industry, suggested_values):
    reality = run_reality_engine(
        idea,
        industry,
        suggested_values["price_per_unit"],
        suggested_values["year1_units"],
        suggested_values["growth_y2"],
        suggested_values["growth_y3"],
        suggested_values["opex_pct"],
        suggested_values["fixed_overhead"],
    )

    lines = []
    lines.append("Reality Check Layer")
    lines.append(
        "TurboPitch also runs a Reality Engine after generating assumptions. This layer checks whether the assumptions look believable in the real world, not just whether they are mathematically consistent."
    )

    for label, item in reality["checks"].items():
        lines.append(f"{label}")
        lines.append(f"{reality_status_icon(item['status'])} {item['message']}")

    lines.append("Reality Engine Summary")
    lines.append(f"{reality_status_icon(reality['overall'])} {reality['summary']}")

    return "\n".join(lines)


def build_full_assumption_explanation(idea, industry, suggested_values):
    idea_text = idea if idea else "No startup idea was provided."

    price = suggested_values["price_per_unit"]
    units = suggested_values["year1_units"]
    g2 = suggested_values["growth_y2"]
    g3 = suggested_values["growth_y3"]
    cogs = suggested_values["cost_per_unit"]
    opex = suggested_values["opex_pct"]
    overhead = suggested_values["fixed_overhead"]
    cash = suggested_values["starting_cash"]
    target_margin = suggested_values["target_gross_margin"]
    segment = suggested_values.get("detected_segment", detect_customer_segment(idea, industry))

    benchmark = INDUSTRY_BENCHMARKS.get(industry, INDUSTRY_BENCHMARKS["SaaS"])
    gm_low, gm_high = benchmark["gross_margin"]
    y2_low, y2_high = benchmark["growth_y2"]
    y3_low, y3_high = benchmark["growth_y3"]
    opex_low, opex_high = benchmark["opex_pct"]
    unit_low, unit_high = benchmark["year1_units"]

    idea_lower = idea_text.lower()

    pricing_logic = []
    if any(word in idea_lower for word in ["enterprise", "b2b", "software", "platform", "analytics", "ai", "api"]):
        pricing_logic.append("The idea description suggests a B2B, software, platform, analytics, AI, or infrastructure-oriented concept, so the starting price was increased using a premium pricing heuristic.")
    if any(word in idea_lower for word in ["restaurant", "retail", "food", "cpg", "product"]):
        pricing_logic.append("The idea description suggests a lower-ticket or more price-sensitive product category, so the starting price was reduced using a consumer / product heuristic.")
    if any(word in idea_lower for word in ["resume", "job", "career", "job seeker", "student"]):
        pricing_logic.append("The idea description suggests an individual end-user and a price-sensitive audience, so the price was capped down using a consumer affordability heuristic.")
    if not pricing_logic:
        pricing_logic.append("The starting price was selected from the base pricing profile tied to the chosen industry, then checked against the target gross margin logic.")

    volume_logic = []
    if any(word in idea_lower for word in ["consumer", "app", "marketplace", "delivery", "subscription box"]):
        volume_logic.append("The idea description suggests a consumer or higher-volume model, so the Year 1 unit assumption was increased using a volume scaling heuristic.")
    else:
        volume_logic.append("The Year 1 unit assumption was anchored to the midpoint of the internal industry traction range so the model starts from a moderate rather than extreme adoption case.")

    reality_text = build_reality_engine_explanation(idea, industry, suggested_values)

    explanation = f"""
Suggested Assumption Rationale

Startup Idea Context
The current startup idea entered into TurboPitch is:
{idea_text}

Detected Customer Segment
TurboPitch interprets this concept as primarily:
{segment}

Why These Assumptions Make Sense
TurboPitch did not invent random numbers. The suggested assumptions were built in layers using a structured process.

First, the system looked at the selected industry, which is {industry}. That industry contains an internal set of benchmark ranges for gross margin, Year 2 growth, Year 3 growth, operating expense ratio, and Year 1 traction levels.

Second, TurboPitch created a baseline model from those benchmark ranges.
The system used the midpoint of the internal industry range for gross margin, which led to a target gross margin of about {target_margin:.0%}.
It also used the midpoint of the internal industry growth ranges to set Year 2 growth at about {g2:.0%} and Year 3 growth at about {g3:.0%}.
For Year 1 traction, it started from the midpoint of the internal range of {unit_low:,} to {unit_high:,}, which helped produce a starting Year 1 unit assumption of {units:,}.
For operating expenses, it used the midpoint of the internal opex range of {opex_low:.0%} to {opex_high:.0%}, which led to a starting opex ratio of {opex:.0%}.

Third, TurboPitch adjusted the base values using keyword and business model heuristics from the idea description.
{" ".join(pricing_logic)}
{" ".join(volume_logic)}

How The AI Came Up With Price
TurboPitch first selected a base starting price from the industry profile for {industry}.
It then adjusted that starting price using keyword-based business model rules tied to the idea description and customer segment.
After that, it recalculated cost per unit to keep the model aligned with the target gross margin.
The final suggested price is ${price:,.2f}.

How The AI Came Up With Cost Per Unit
TurboPitch did not guess cost per unit independently.
Instead, cost per unit was derived from the chosen price and the target gross margin.
The formula used is:
Cost per Unit = Price per Unit × (1 - Target Gross Margin)
With a suggested price of ${price:,.2f} and a target gross margin of about {target_margin:.0%}, the model produced a cost per unit of about ${cogs:,.2f}.
This makes cost structure consistent with the intended margin profile.

How The AI Came Up With Year 1 Units
TurboPitch used the internal industry Year 1 unit range for {industry}, which is {unit_low:,} to {unit_high:,}.
It started from the middle of that range to avoid making the first model too conservative or too aggressive.
If the idea description looked more consumer-oriented or volume-oriented, the model increased the Year 1 unit assumption using a simple scaling heuristic.
The resulting Year 1 unit assumption is {units:,}.

How The AI Came Up With Growth Rates
TurboPitch used the internal growth benchmark ranges for the selected industry.
For {industry}, the internal Year 2 growth range is {y2_low:.0%} to {y2_high:.0%}, and the internal Year 3 growth range is {y3_low:.0%} to {y3_high:.0%}.
The model used the middle of those ranges to create a balanced starting point.
That produced a Year 2 growth rate of {g2:.0%} and a Year 3 growth rate of {g3:.0%}.
These are intended to represent a believable early-stage scaling path rather than a maximum upside case.

How The AI Came Up With Operating Expense %
TurboPitch used the internal operating expense range for the selected industry.
For {industry}, the internal opex range is {opex_low:.0%} to {opex_high:.0%}.
The model used the midpoint of that range to generate a starting operating expense assumption of {opex:.0%}.
This is meant to reflect a practical early-stage operating burden before the founder fine-tunes hiring, marketing spend, support cost, and overhead structure.

How The AI Came Up With Fixed Overhead
Fixed overhead was not pulled from a live market database.
TurboPitch uses a baseline fixed overhead amount by industry type, then adjusts it upward when the idea appears more enterprise, software-heavy, or infrastructure-heavy.
That produced a starting fixed overhead value of ${overhead:,.0f}.
This number is meant to capture annual costs that do not directly scale with each unit sold, such as salaries, software, rent, and admin burden.

How The AI Came Up With Starting Cash
Starting cash is based on a baseline cash profile by industry type rather than on a live funding dataset.
TurboPitch uses higher starting cash for more capital-intensive or software-heavy concepts and lower starting cash for lighter models.
That produced a starting cash assumption of ${cash:,.0f}.
This is meant to represent a practical beginning runway assumption for an early-stage model, not a statement of what the founder must raise.

Reality Check Layer
{reality_text.replace("Reality Check Layer", "").strip()}

What The Founder Should Validate First
The founder should validate pricing first by checking what similar solutions, products, or services charge in the market.
Next, they should validate whether Year 1 volume is realistic based on access to customers, channels, partnerships, and demand generation capability.
Then they should validate costs, especially if the business depends on fulfillment, labor, onboarding, paid marketing, hosting, or AI infrastructure.
Finally, they should validate whether the chosen growth rates are actually possible given the go-to-market plan.

What May Need To Be Adjusted
These assumptions may need to change if the business is more premium, more commoditized, more service-heavy, more local, or more enterprise than the first draft suggests.
The founder may need to raise price, reduce Year 1 volume, increase cost per unit, increase opex, or raise starting cash depending on real-world conditions.
TurboPitch is giving a structured starting point, not a final truth.

How Conservative Or Aggressive This Starting Model Is
This starting model is intended to be moderately balanced.
It is not a worst-case stress test, and it is not a hyper-aggressive venture fantasy model either.
It is meant to give the founder a credible first-pass model that can be refined after more research and feedback.

Important Limitation
TurboPitch uses internal benchmark ranges, rule-based logic, Reality Engine heuristics, and AI explanation to build starter assumptions.
It does not currently pull live proprietary pricing databases or live real-time market feeds into this specific assumption generator.
The value of the feature is that it gives founders a rational starting model, challenges obvious real-world mismatches, and explains how it got there, so they are not guessing from scratch.
""".strip()

    return explanation


def render_ai_methodology_note():
    st.markdown("### How this analysis was generated")
    st.markdown(
        """
This review was generated using a combination of:

- founder-provided assumptions
- structured financial modeling
- rule-based benchmark logic
- market reality heuristics
- AI interpretation of realism, consistency, and investor risk

TurboPitch is designed to act like an investor-readiness filter, not a generic chatbot.
It evaluates whether assumptions look credible enough to survive investor scrutiny.
        """
    )


def render_trust_center():
    st.markdown("## Why Trust TurboPitch?")
    st.markdown(
        """
TurboPitch is designed to make AI output more credible by combining structured financial modeling,
industry benchmark logic, market reality checks, and investor-style analysis.

The goal is not to pretend the AI knows everything.
The goal is to show founders how their assumptions may look through the eyes of an investor.
        """
    )

    trust_tab1, trust_tab2, trust_tab3, trust_tab4 = st.tabs([
        "How the AI Works",
        "Data Sources",
        "Transparency",
        "Limitations"
    ])

    with trust_tab1:
        st.markdown("### How the AI Works")
        st.markdown(
            """
1. Founder inputs assumptions  
   The founder enters pricing, sales volume, growth, costs, and cash assumptions.

2. TurboPitch builds a structured financial model  
   Revenue, COGS, gross profit, operating expenses, taxes, net income, and ending cash are calculated.

3. Rule-based checks are applied  
   The model checks for structural issues such as negative margins, aggressive growth, high operating cost burden, and scale risk.

4. Benchmark logic is applied  
   Key assumptions are compared against industry ranges by business type.

5. Reality Engine checks are applied  
   The model checks for customer-segment fit, pricing plausibility, adoption realism, and operating model credibility.

6. AI interprets the results  
   The AI explains where investors may push back, what looks credible, and what should be improved.
            """
        )
        st.info("TurboPitch is designed to interpret structured financial logic and challenge obvious real-world mismatches.")

    with trust_tab2:
        st.markdown("### Data Sources Used for Benchmarking and Assumption Generation")
        for item in TRUST_DATA_SOURCES:
            st.markdown(f"**{item['category']}**")
            st.markdown(f"- Examples: {item['examples']}")
            st.markdown(f"- How it is used: {item['use_case']}")
        st.success("The purpose of this section is to show users that the analysis is grounded in recognizable business logic, benchmark structures, and explainable modeling steps.")

    with trust_tab3:
        st.markdown("### Trust and Transparency Principles")
        st.markdown(
            """
TurboPitch is built around a few simple trust principles:

- show the assumptions
- show the math
- show the benchmark ranges
- show the market-reality checks
- explain the reasoning
- acknowledge limitations

That means the platform should feel more like a structured venture analysis workflow and less like a black-box AI response.
            """
        )

        st.markdown("### Built on FP&A-Style and Investor-Style Thinking")
        st.markdown(
            """
TurboPitch uses logic similar to FP&A and business performance review frameworks:
pricing discipline, margin quality, growth realism, operating efficiency, cash viability, and customer willingness to pay.

This helps make the output more investor-relevant and more explainable.
            """
        )

    with trust_tab4:
        st.markdown("### Important Limitations")
        for item in TRUST_LIMITATIONS:
            st.warning(item)


# ==================================================
# AI FUNCTIONS
# ==================================================
def run_ai_assumption_helper(idea, industry, suggested_values):
    base_explanation = build_full_assumption_explanation(idea, industry, suggested_values)

    prompt = f"""
You are an FP&A analyst helping a first-time founder understand suggested startup assumptions.

You are given a structured internal explanation of how TurboPitch generated the starter assumptions.
Your job is to rewrite it so it sounds polished, clear, and founder-friendly while preserving the logic.

Startup idea:
{idea if idea else "No startup idea provided."}

Industry:
{industry}

Structured explanation:
{base_explanation}

Instructions:
- Keep the section title as Suggested Assumption Rationale
- Preserve the logic for how price, units, growth, cost per unit, opex, fixed overhead, and starting cash were derived
- Preserve the Reality Check Layer and explain it clearly
- Make it detailed and easy to understand
- Do not use markdown symbols like ##, ###, **, or bullet asterisks
- Do not invent live external data sources
- Make clear that TurboPitch uses internal benchmark ranges, heuristics, Reality Engine logic, and financial model logic
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a practical FP&A startup advisor helping founders understand starter business assumptions in a transparent way."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.4,
        )

        content = response.choices[0].message.content if response.choices and response.choices[0].message else ""
        if content and content.strip():
            return content.strip()

        return base_explanation

    except Exception:
        return base_explanation


def run_ai_sanity_check(
    idea,
    industry,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    projection_df,
    reality_engine_output
):
    fin_summary = financial_summary_text(projection_df)
    reality_summary = "\n".join(
        [f"{k}: {v['status']} - {v['message']}" for k, v in reality_engine_output["checks"].items()]
    )

    prompt = f"""
You are a skeptical but constructive venture capital investor reviewing an early-stage startup.

Your job is to evaluate whether this startup looks investable based on the idea, business model, financial assumptions, and market realism checks.

Write a report with these exact section headings:

Overall Impression
What Investors Will Like
What Investors Will Challenge
Biggest Risks
Final Investor Verdict
Model Adjustments
Suggestions to Improve Investor Appeal
How to Reframe This Pitch for Investors

Instructions:
- Be practical, sharp, and investor-minded.
- Think like a real early-stage investor, not a cheerleader.
- Be honest if the model looks weak or unrealistic.
- Use both the financial assumptions and the Reality Engine checks.
- If pricing looks mathematically strong but unrealistic for the customer, say so clearly.
- Do not just critique the model — explain what should change.
- In Final Investor Verdict, answer clearly with one of these:
  Yes
  Maybe
  No
- In Final Investor Verdict, explain the reasoning in 2 to 4 sentences.
- In Model Adjustments, give 4 to 6 specific recommended changes using concrete numbers whenever possible.
- Consider the startup's industry when evaluating realism.
- Do NOT write a full business plan.
- Do NOT write a pitch deck.
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks.
- Keep formatting clean and readable.

Startup Idea:
{idea if idea else "No startup idea provided."}

Industry:
{industry}

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

Reality Engine Summary:
{reality_summary}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a disciplined early-stage VC investor who evaluates startup ideas for investability, realism, and investor appeal."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.55,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI sanity check unavailable.\n\nError: {e}"


def run_ai_investor_interrogation(
    idea,
    industry,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    starting_cash,
    projection_df,
    reality_engine_output
):
    fin_summary = financial_summary_text(projection_df)
    reality_summary = "\n".join(
        [f"{k}: {v['status']} - {v['message']}" for k, v in reality_engine_output["checks"].items()]
    )

    prompt = f"""
You are a skeptical venture capitalist in a pitch meeting.

Your job is to challenge the founder with hard but realistic investor questions.

Write a section called:

Investor Questions

Then list 10 to 12 specific, tough questions an investor would ask this founder.

Instructions:
- Be sharp, skeptical, and practical.
- Consider the startup's industry when generating investor questions.
- Use the financial assumptions and the Reality Engine checks.
- Focus on weak assumptions, growth realism, market adoption, customer acquisition, competition, margins, operating costs, pricing realism, scaling risk, and funding use.
- Questions should sound like real VC pushback in a live meeting.
- Do NOT answer the questions.
- Do NOT write a business plan.
- Do NOT write in essay form.
- Keep each question concise and direct.
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks.

Startup Idea:
{idea if idea else "No startup idea provided."}

Industry:
{industry}

Assumptions:
Price per unit: ${price_per_unit:,.2f}
Year 1 units sold: {year1_units:,}
Year 2 growth rate: {growth_y2:.2%}
Year 3 growth rate: {growth_y3:.2%}
Cost per unit: ${cost_per_unit:,.2f}
Operating expense percent of revenue: {opex_pct:.2%}
Fixed annual overhead: ${fixed_overhead:,.0f}
Starting cash: ${starting_cash:,.0f}

Projection Summary:
{fin_summary}

Reality Engine Summary:
{reality_summary}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a skeptical venture capitalist who challenges startup founders with tough investor questions."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Investor interrogation unavailable.\n\nError: {e}"


def run_ai_founder_answer_builder(
    idea,
    industry,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    starting_cash,
    projection_df,
    reality_engine_output
):
    fin_summary = financial_summary_text(projection_df)
    reality_summary = "\n".join(
        [f"{k}: {v['status']} - {v['message']}" for k, v in reality_engine_output["checks"].items()]
    )

    prompt = f"""
You are a startup pitch coach helping a founder prepare for tough investor meetings.

Create a section called:

Founder Answer Prep

Generate 8 investor-style questions the founder is likely to face.

For each question, include these exact sub-sections:

Investor Question
What Investors Are Really Asking
Strong Answer Framework
Weak Answer Example
How To Improve The Answer

Instructions:
- Be practical, realistic, and investor-focused.
- Consider the startup's industry when creating the questions and answer prep.
- Use the startup idea, assumptions, and Reality Engine summary to create the questions.
- In Strong Answer Framework, give concise talking points, not a long essay.
- In Weak Answer Example, show the kind of vague answer founders often give.
- In How To Improve The Answer, explain how to make the response more credible.
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks.
- Keep formatting clean and readable.

Startup Idea:
{idea if idea else "No startup idea provided."}

Industry:
{industry}

Assumptions:
Price per unit: ${price_per_unit:,.2f}
Year 1 units sold: {year1_units:,}
Year 2 growth rate: {growth_y2:.2%}
Year 3 growth rate: {growth_y3:.2%}
Cost per unit: ${cost_per_unit:,.2f}
Operating expense percent of revenue: {opex_pct:.2%}
Fixed annual overhead: ${fixed_overhead:,.0f}
Starting cash: ${starting_cash:,.0f}

Projection Summary:
{fin_summary}

Reality Engine Summary:
{reality_summary}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a practical startup pitch coach who helps founders answer investor questions more effectively."
                },
                {"role": "user", "content": prompt},
            ],
            temperature=0.7,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Founder answer builder unavailable.\n\nError: {e}"


def generate_business_plan_and_deck(
    idea,
    industry,
    price_per_unit,
    year1_units,
    growth_y2,
    growth_y3,
    cost_per_unit,
    opex_pct,
    fixed_overhead,
    projection_df,
    reality_engine_output
):
    fin_summary = financial_summary_text(projection_df)
    reality_summary = "\n".join(
        [f"{k}: {v['status']} - {v['message']}" for k, v in reality_engine_output["checks"].items()]
    )

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
- Incorporate the Reality Engine concerns where appropriate

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
- Consider the startup's industry in the business plan and deck content.
- Use the Reality Engine to avoid blindly endorsing unrealistic assumptions.
- Do NOT use markdown symbols like ##, ###, **, or bullet asterisks
- Keep the business plan polished and readable
- Make the deck bullets concise and presentation-ready

Startup Idea:
{idea if idea else "No startup idea provided."}

Industry:
{industry}

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

Reality Engine Summary:
{reality_summary}
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
    "Describe Your Startup",
    value=st.session_state.idea,
    height=140,
    placeholder="Explain what the business does, who it serves, and why it matters..."
)


# ==================================================
# SIDEBAR INPUTS
# ==================================================
with st.sidebar:
    st.header("Financial Assumptions")

    st.session_state.industry = st.selectbox(
        "Startup industry",
        options=list(INDUSTRY_BENCHMARKS.keys()),
        index=list(INDUSTRY_BENCHMARKS.keys()).index(st.session_state.industry),
    )

    st.markdown("### Assumption Setup")

    st.session_state.assumption_mode = st.radio(
        "Choose assumption mode",
        ["Manual", "Help Me Generate Them"],
        index=0 if st.session_state.assumption_mode == "Manual" else 1
    )

    if st.session_state.assumption_mode == "Help Me Generate Them":
        st.caption("Not sure what numbers to use? TurboPitch can generate a starting model based on your industry, idea, and reality-engine heuristics.")

        if st.button("Generate Suggested Assumptions"):
            suggested = generate_rule_based_assumptions(
                st.session_state.industry,
                st.session_state.idea
            )

            st.session_state.price_per_unit = suggested["price_per_unit"]
            st.session_state.year1_units = suggested["year1_units"]
            st.session_state.growth_y2 = suggested["growth_y2"]
            st.session_state.growth_y3 = suggested["growth_y3"]
            st.session_state.cost_per_unit = suggested["cost_per_unit"]
            st.session_state.opex_pct = suggested["opex_pct"]
            st.session_state.fixed_overhead = suggested["fixed_overhead"]
            st.session_state.starting_cash = suggested["starting_cash"]

            with st.spinner("Explaining suggested assumptions..."):
                helper_text = run_ai_assumption_helper(
                    st.session_state.idea,
                    st.session_state.industry,
                    suggested
                )

            st.session_state.assumption_helper_output = helper_text
            st.success("Suggested starter assumptions generated with explanation.")

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
    st.markdown("### Investor Pushback Scenario")

    st.session_state.pushback_pct = st.number_input(
        "Reduce revenue assumptions by %",
        min_value=0,
        max_value=100,
        value=int(st.session_state.pushback_pct),
        step=5,
    )

    st.session_state.investor_feedback = st.text_area(
        "Investor feedback notes",
        value=st.session_state.investor_feedback,
        height=100,
    )

    if st.button("Apply Pushback Scenario"):
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
display_pnl = build_display_pnl(pnl_df)

reality_engine_output = run_reality_engine(
    st.session_state.idea,
    st.session_state.industry,
    st.session_state.price_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    st.session_state.opex_pct,
    st.session_state.fixed_overhead,
)

scorecard = build_scorecard(
    st.session_state.idea,
    st.session_state.industry,
    projection_df,
    st.session_state.price_per_unit,
    st.session_state.cost_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    reality_engine_output,
)

warnings = run_rule_based_sanity_check(
    st.session_state.price_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    st.session_state.cost_per_unit,
    st.session_state.opex_pct,
)

warning_summary = build_warning_summary(warnings)

benchmark_feedback = build_benchmark_feedback(
    st.session_state.industry,
    st.session_state.price_per_unit,
    st.session_state.cost_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3,
    st.session_state.opex_pct,
)


# ==================================================
# TOP ACTION BUTTONS
# ==================================================
col_top1, col_top2, col_top3, col_top4, _ = st.columns([1.2, 1.6, 1.4, 1.4, 2.8])

with col_top1:
    if st.button("Get Investor Verdict", key="run_vc_sanity_top"):
        with st.spinner("Analyzing investability..."):
            sanity_text = run_ai_sanity_check(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
                reality_engine_output,
            )
            st.session_state.sanity_output = sanity_text
        st.success("Investor verdict generated.")

with col_top2:
    if st.button("Generate Investor Materials", key="generate_full_plan_top"):
        with st.spinner("Generating investor materials..."):
            plan_text = generate_business_plan_and_deck(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
                reality_engine_output,
            )
            st.session_state.business_plan_output = plan_text
        st.success("Business plan and pitch deck content generated.")

with col_top3:
    if st.button("Generate VC Questions", key="generate_vc_questions_top"):
        with st.spinner("Generating investor questions..."):
            interrogation_text = run_ai_investor_interrogation(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                st.session_state.starting_cash,
                projection_df,
                reality_engine_output,
            )
            st.session_state.interrogation_output = interrogation_text
        st.success("Investor questions generated.")

with col_top4:
    if st.button("Build Founder Answers", key="build_founder_answers_top"):
        with st.spinner("Building founder answer prep..."):
            answer_text = run_ai_founder_answer_builder(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                st.session_state.starting_cash,
                projection_df,
                reality_engine_output,
            )
            st.session_state.answer_builder_output = answer_text
        st.success("Founder answer prep generated.")


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
    st.markdown("#### Profit & Loss Summary")
    st.dataframe(display_pnl, use_container_width=True, hide_index=True)

with summary_col2:
    st.markdown("#### Investor Snapshot")
    st.markdown(f"""
    <div class="snapshot-card">
        <p><strong>Industry:</strong> {st.session_state.industry}</p>
        <p><strong>Customer Segment:</strong> {detect_customer_segment(st.session_state.idea, st.session_state.industry)}</p>
        <p><strong>Overall Readiness:</strong> {scorecard['Overall Investor Readiness']}</p>
        <p><strong>Pricing Realism:</strong> {scorecard['Pricing Realism']}</p>
        <p><strong>Sales Volume:</strong> {scorecard['Sales Volume']}</p>
        <p><strong>Growth Assumptions:</strong> {scorecard['Growth Assumptions']}</p>
        <p><strong>Margin Quality:</strong> {scorecard['Margin Quality']}</p>
        <p><strong>Cash Viability:</strong> {scorecard['Cash Viability']}</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("#### Trust Summary")
st.info(
    "TurboPitch combines founder inputs, structured financial modeling, benchmark logic, market reality checks, and AI interpretation. "
    "It is designed as a decision-support tool to improve investor readiness, not as a guarantee of funding or business success."
)

if st.session_state.assumption_mode == "Help Me Generate Them":
    st.markdown("#### Starter Assumption Mode")
    st.info(
        "TurboPitch generated a suggested starting model based on internal industry benchmark ranges, business model heuristics, "
        "Reality Engine logic, and financial modeling logic. These values are intended as a starting point for refinement, not final answers."
    )

    if st.session_state.get("assumption_helper_output"):
        st.markdown("#### Suggested Assumption Explanation")
        st.markdown(clean_ai_text(st.session_state.assumption_helper_output))


# ==================================================
# TABS
# ==================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "Investor Review",
    "Benchmarking",
    "Investor Q&A",
    "Answer Builder",
    "Business Plan & Deck",
    "Financial Model",
    "Charts",
    "Downloads",
    "How TurboPitch Works",
    "Assumption Builder",
])

with tab1:
    st.markdown("## Investor Readiness Scorecard")
    for label, status in scorecard.items():
        st.write(f"{score_metric(status)} **{label}:** {status}")

    st.markdown("---")
    st.markdown("### Assumption Risk Summary")
    st.info(
        f"{warning_summary['icon']} {warning_summary['overall']}  |  "
        f"Red Flags: {warning_summary['red_count']}  |  "
        f"Watch Items: {warning_summary['amber_count']}"
    )

    st.markdown("---")
    st.markdown("### Reality Engine")
    st.write(f"{score_metric(reality_engine_output['overall'])} **Overall Reality Check:** {reality_engine_output['overall']}")
    st.write(reality_engine_output["summary"])

    for label, item in reality_engine_output["checks"].items():
        st.write(f"{reality_status_icon(item['status'])} **{label}:** {item['message']}")

    st.markdown("---")
    st.markdown("### Rule-Based Assumption Review")
    for warning in warnings:
        st.write(warning)

    st.markdown("---")
    st.markdown("### AI Investor Verdict, Risks & Recommendations")

    if st.button("Run VC Analysis", key="run_ai_sanity_tab"):
        with st.spinner("Reviewing startup assumptions..."):
            sanity_text = run_ai_sanity_check(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
                reality_engine_output,
            )
            st.session_state.sanity_output = sanity_text

    if st.session_state.get("sanity_output"):
        st.markdown(clean_ai_text(st.session_state.sanity_output))
        st.markdown("---")
        render_ai_methodology_note()
    else:
        st.info("No investor review available yet. Click 'Get Investor Verdict' or 'Run VC Analysis'.")

with tab2:
    st.markdown("## Industry Benchmark Feedback")
    st.write(f"Selected industry: **{st.session_state.industry}**")

    for item in benchmark_feedback:
        if item == "":
            st.markdown("---")
        elif item == "Suggested Adjustments":
            st.markdown("### Suggested Adjustments")
        else:
            st.write(item)

    st.markdown("---")
    st.markdown("### Benchmark Methodology")
    st.markdown(
        """
Benchmark comparisons in TurboPitch are based on internal benchmark ranges by business model.
These are intended to reflect directional industry norms, investor expectations, and common startup operating patterns.

They are best used as a credibility check, not as a substitute for full market diligence.
        """
    )

with tab3:
    st.markdown("## Investor Questions")
    st.write("Generate tough investor-style questions to pressure-test the pitch before a real meeting.")

    if st.button("Generate VC Questions", key="generate_vc_questions_tab"):
        with st.spinner("Generating investor questions..."):
            interrogation_text = run_ai_investor_interrogation(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                st.session_state.starting_cash,
                projection_df,
                reality_engine_output,
            )
            st.session_state.interrogation_output = interrogation_text

    if st.session_state.get("interrogation_output"):
        st.markdown(clean_ai_text(st.session_state.interrogation_output))
        st.markdown("---")
        render_ai_methodology_note()
    else:
        st.info("No investor questions yet. Click 'Generate VC Questions'.")

with tab4:
    st.markdown("## Founder Answer Prep")
    st.write("Prepare stronger answers to likely investor objections before your pitch.")

    if st.button("Build Founder Answers", key="build_founder_answers_tab"):
        with st.spinner("Building founder answer prep..."):
            answer_text = run_ai_founder_answer_builder(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                st.session_state.starting_cash,
                projection_df,
                reality_engine_output,
            )
            st.session_state.answer_builder_output = answer_text

    if st.session_state.get("answer_builder_output"):
        st.markdown(clean_ai_text(st.session_state.answer_builder_output))
        st.markdown("---")
        render_ai_methodology_note()
    else:
        st.info("No founder answer prep yet. Click 'Build Founder Answers'.")

with tab5:
    st.markdown("### Full Business Plan & Pitch Deck Content")

    if st.button("Build Business Plan + Deck", key="generate_plan_tab"):
        with st.spinner("Building investor materials..."):
            plan_text = generate_business_plan_and_deck(
                st.session_state.idea,
                st.session_state.industry,
                st.session_state.price_per_unit,
                st.session_state.year1_units,
                st.session_state.growth_y2,
                st.session_state.growth_y3,
                st.session_state.cost_per_unit,
                st.session_state.opex_pct,
                st.session_state.fixed_overhead,
                projection_df,
                reality_engine_output,
            )
            st.session_state.business_plan_output = plan_text

    if st.session_state.get("business_plan_output"):
        st.markdown(clean_ai_text(st.session_state.business_plan_output))
        st.markdown("---")
        render_ai_methodology_note()
    else:
        st.info("No business plan available yet. Click 'Generate Investor Materials' or 'Build Business Plan + Deck'.")

with tab6:
    st.dataframe(display_pnl, use_container_width=True, hide_index=True)

with tab7:
    chart_df = projection_df.set_index("Year")[["Revenue", "Net Income", "Ending Cash"]]
    st.line_chart(chart_df, use_container_width=True)

with tab8:
    st.markdown("### Export Investor Materials")
    st.write("Download a full business plan, financial package, or presentation-ready pitch deck.")

with tab9:
    render_trust_center()

with tab10:
    st.markdown("## Assumption Builder")
    st.write("Use this section if you are unsure how to price your product or what assumptions to start with.")

    if st.session_state.assumption_mode == "Help Me Generate Them":
        st.markdown("### Current Suggested Starter Assumptions")
        st.write(f"**Price per unit:** ${st.session_state.price_per_unit:,.2f}")
        st.write(f"**Year 1 units sold:** {st.session_state.year1_units:,}")
        st.write(f"**Year 2 growth:** {st.session_state.growth_y2:.0%}")
        st.write(f"**Year 3 growth:** {st.session_state.growth_y3:.0%}")
        st.write(f"**Cost per unit:** ${st.session_state.cost_per_unit:,.2f}")
        st.write(f"**Operating expense %:** {st.session_state.opex_pct:.0%}")
        st.write(f"**Fixed overhead:** ${st.session_state.fixed_overhead:,.0f}")
        st.write(f"**Starting cash:** ${st.session_state.starting_cash:,.0f}")
        st.write(f"**Detected customer segment:** {detect_customer_segment(st.session_state.idea, st.session_state.industry)}")
        st.markdown("---")

    st.markdown("### Reality Engine Preview")
    for label, item in reality_engine_output["checks"].items():
        st.write(f"{reality_status_icon(item['status'])} **{label}:** {item['message']}")

    st.markdown("---")

    if st.session_state.get("assumption_helper_output"):
        st.markdown(clean_ai_text(st.session_state.assumption_helper_output))
        st.markdown("---")
        render_ai_methodology_note()
    else:
        st.info("No suggested assumptions yet. In the sidebar, choose 'Help Me Generate Them' and click 'Generate Suggested Assumptions'.")


# ==================================================
# DOWNLOAD MATERIALS
# ==================================================
plan_raw = st.session_state.get("business_plan_output", "")
business_plan_text = extract_business_plan_section(plan_raw)
deck_slides = extract_pitch_deck_section(plan_raw)

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

doc.add_heading("Methodology Note", level=1)
doc.add_paragraph(
    "TurboPitch combines founder inputs, internal benchmark ranges, rule-based business heuristics, "
    "Reality Engine logic, and structured financial modeling logic to create investor-readiness analysis and starter assumptions."
)

doc.add_heading("Reality Engine Summary", level=1)
for label, item in reality_engine_output["checks"].items():
    doc.add_paragraph(f"{label}: {item['status']} - {item['message']}")

if st.session_state.get("assumption_helper_output"):
    doc.add_heading("Suggested Assumption Rationale", level=1)
    for paragraph in clean_ai_text(st.session_state.assumption_helper_output).split("\n\n"):
        paragraph = paragraph.strip()
        if paragraph:
            doc.add_paragraph(paragraph)

doc_buffer = io.BytesIO()
doc.save(doc_buffer)
doc_buffer.seek(0)

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

# ---------------- Dashboard Sheet ----------------
ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_dash.sheet_view.showGridLines = False

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
        ws_dash[f"B{i}"].number_format = "$#,##0"

ws_dash.merge_cells("D4:E4")
ws_dash["D4"] = "Investor Snapshot"
ws_dash["D4"].font = section_font
ws_dash["D4"].fill = section_fill

snapshot_items = [
    ("Industry", st.session_state.industry),
    ("Customer Segment", detect_customer_segment(st.session_state.idea, st.session_state.industry)),
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

ws_dash.merge_cells("A14:D14")
ws_dash["A14"] = "P&L Summary"
ws_dash["A14"].font = section_font
ws_dash["A14"].fill = section_fill

pnl_headers = list(pnl_df.columns)
for col_idx, col_name in enumerate(pnl_headers, start=1):
    cell = ws_dash.cell(row=15, column=col_idx)
    cell.value = col_name
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row_idx, (_, row) in enumerate(pnl_df.iterrows(), start=16):
    for col_idx, value in enumerate(row, start=1):
        cell = ws_dash.cell(row=row_idx, column=col_idx)
        cell.value = value
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if col_idx > 1:
            line_item = row["Line Item"]
            if line_item == "Gross Margin %":
                cell.number_format = "0.0%"
            elif line_item == "Units":
                cell.number_format = "#,##0"
            else:
                cell.number_format = "$#,##0"

for col in range(1, 6):
    col_letter = get_column_letter(col)
    max_length = 0
    for row in range(1, ws_dash.max_row + 1):
        value = ws_dash[f"{col_letter}{row}"].value
        if value is not None:
            max_length = max(max_length, len(str(value)))
    ws_dash.column_dimensions[col_letter].width = max_length + 3

ws_dash.merge_cells("G4:L4")
ws_dash["G4"] = "Projection Charts"
ws_dash["G4"].font = section_font
ws_dash["G4"].fill = section_fill

projection_chart_stream = create_excel_projection_chart_image(projection_df)
projection_chart_file = io.BytesIO(projection_chart_stream.getvalue())
projection_chart_img = XLImage(projection_chart_file)
projection_chart_img.width = 760
projection_chart_img.height = 360
ws_dash.add_image(projection_chart_img, "G5")

compare_chart_stream = create_excel_compare_chart_image(projection_df)
compare_chart_file = io.BytesIO(compare_chart_stream.getvalue())
compare_chart_img = XLImage(compare_chart_file)
compare_chart_img.width = 760
compare_chart_img.height = 360
ws_dash.add_image(compare_chart_img, "G24")

for row_num in range(5, 24):
    ws_dash.row_dimensions[row_num].height = 22

for row_num in range(24, 43):
    ws_dash.row_dimensions[row_num].height = 22

for col_letter in ["G", "H", "I", "J", "K", "L"]:
    ws_dash.column_dimensions[col_letter].width = 16

# ---------------- Financial Model Sheet ----------------
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
            cell.number_format = "0.0%"
        elif line_item == "Units":
            cell.number_format = "#,##0"
        else:
            cell.number_format = "$#,##0"

for col_idx, col_name in enumerate(headers, start=1):
    col_letter = get_column_letter(col_idx)
    max_length = len(str(col_name))
    for row_idx in range(2, ws_model.max_row + 1):
        value = ws_model[f"{col_letter}{row_idx}"].value
        if value is not None:
            max_length = max(max_length, len(str(value)))
    ws_model.column_dimensions[col_letter].width = max_length + 2

ws_model.freeze_panes = "A2"

# ---------------- Assumptions Sheet ----------------
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
    ("Industry", st.session_state.industry),
    ("Detected Customer Segment", detect_customer_segment(st.session_state.idea, st.session_state.industry)),
    ("Assumption Mode", st.session_state.assumption_mode),
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
        ws_assump[f"B{idx}"].number_format = "$#,##0.00"
    elif label in {"Year 2 Growth Rate", "Year 3 Growth Rate", "Operating Expense %"}:
        ws_assump[f"B{idx}"].number_format = "0.0%"
    elif label in {"Year 1 Units Sold"}:
        ws_assump[f"B{idx}"].number_format = "#,##0"

ws_assump.column_dimensions["A"].width = 30
ws_assump.column_dimensions["B"].width = 70
ws_assump.freeze_panes = "A4"

# ---------------- Methodology Sheet ----------------
ws_method = wb.create_sheet("Methodology")
ws_method.merge_cells("A1:B1")
ws_method["A1"] = "TurboPitch Methodology"
ws_method["A1"].font = title_font
ws_method["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_method["A3"] = "Section"
ws_method["B3"] = "Description"

for cell in ws_method[3]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

# ---------------- Methodology Sheet ----------------
ws_method = wb.create_sheet("Methodology")
ws_method.merge_cells("A1:B1")
ws_method["A1"] = "TurboPitch Methodology"
ws_method["A1"].font = title_font
ws_method["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_method["A3"] = "Section"
ws_method["B3"] = "Description"

for cell in ws_method[3]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

methodology_rows = [
    (
        "Core Purpose",
        "TurboPitch is a decision-support tool that helps founders pressure-test startup assumptions through financial modeling, benchmark logic, market reality checks, and AI interpretation."
    ),
    (
        "Founder Inputs",
        "The model starts with founder-provided inputs such as startup idea, industry, pricing, unit volume, growth, cost structure, operating expenses, fixed overhead, and starting cash."
    ),
    (
        "Financial Modeling Logic",
        "TurboPitch calculates revenue, COGS, gross profit, gross margin, operating expenses, operating income, taxes, net income, and ending cash across a 3-year projection."
    ),
    (
        "Benchmark Logic",
        "Key assumptions are compared against internal benchmark ranges by industry, including gross margin, growth, operating expense ratios, and Year 1 volume ranges."
    ),
    (
        "Reality Engine",
        "TurboPitch checks whether pricing, customer segment fit, adoption expectations, growth, and operating model assumptions look believable in the real world."
    ),
    (
        "AI Interpretation",
        "AI is used to explain investor risk, likely pushback, suggested changes, founder answer prep, investor questions, business plan language, and pitch deck content."
    ),
    (
        "Transparency Principle",
        "The platform is designed to show assumptions, show the math, explain the logic, and identify weak spots rather than act like a black-box oracle."
    ),
    (
        "Important Limitation",
        "TurboPitch uses internal benchmark logic, heuristics, and AI reasoning. It is not a substitute for primary market research, customer validation, or legal/accounting advice."
    ),
]

for idx, (section, description) in enumerate(methodology_rows, start=4):
    ws_method[f"A{idx}"] = section
    ws_method[f"B{idx}"] = description

    ws_method[f"A{idx}"].font = Font(bold=True)
    ws_method[f"A{idx}"].alignment = Alignment(vertical="top", wrap_text=True)
    ws_method[f"B{idx}"].alignment = Alignment(vertical="top", wrap_text=True)

    ws_method[f"A{idx}"].border = thin_border
    ws_method[f"B{idx}"].border = thin_border

ws_method.column_dimensions["A"].width = 28
ws_method.column_dimensions["B"].width = 95
ws_method.freeze_panes = "A4"

# ---------------- Benchmark Feedback Sheet ----------------
ws_bench = wb.create_sheet("Benchmark Feedback")
ws_bench.merge_cells("A1:B1")
ws_bench["A1"] = "Industry Benchmark Feedback"
ws_bench["A1"].font = title_font
ws_bench["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_bench["A3"] = "Category"
ws_bench["B3"] = "Feedback"

for cell in ws_bench[3]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

bench_row = 4
current_category = "Feedback"

for item in benchmark_feedback:
    if item == "":
        continue
    elif item == "Suggested Adjustments":
        current_category = "Suggested Adjustments"
        continue
    else:
        ws_bench[f"A{bench_row}"] = current_category
        ws_bench[f"B{bench_row}"] = item

        ws_bench[f"A{bench_row}"].font = Font(bold=True)
        ws_bench[f"A{bench_row}"].alignment = Alignment(vertical="top", wrap_text=True)
        ws_bench[f"B{bench_row}"].alignment = Alignment(vertical="top", wrap_text=True)

        ws_bench[f"A{bench_row}"].border = thin_border
        ws_bench[f"B{bench_row}"].border = thin_border
        bench_row += 1

ws_bench.column_dimensions["A"].width = 24
ws_bench.column_dimensions["B"].width = 110
ws_bench.freeze_panes = "A4"

# ---------------- Reality Engine Sheet ----------------
ws_reality = wb.create_sheet("Reality Engine")
ws_reality.merge_cells("A1:B1")
ws_reality["A1"] = "Reality Engine Review"
ws_reality["A1"].font = title_font
ws_reality["A1"].alignment = Alignment(horizontal="center", vertical="center")

ws_reality["A3"] = "Check"
ws_reality["B3"] = "Assessment"

for cell in ws_reality[3]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = thin_border

reality_row = 4
for label, item in reality_engine_output["checks"].items():
    ws_reality[f"A{reality_row}"] = f"{label} ({item['status']})"
    ws_reality[f"B{reality_row}"] = item["message"]

    ws_reality[f"A{reality_row}"].font = Font(bold=True)
    ws_reality[f"A{reality_row}"].alignment = Alignment(vertical="top", wrap_text=True)
    ws_reality[f"B{reality_row}"].alignment = Alignment(vertical="top", wrap_text=True)

    ws_reality[f"A{reality_row}"].border = thin_border
    ws_reality[f"B{reality_row}"].border = thin_border
    reality_row += 1

ws_reality[f"A{reality_row}"] = "Overall"
ws_reality[f"B{reality_row}"] = f"{reality_engine_output['overall']} - {reality_engine_output['summary']}"
ws_reality[f"A{reality_row}"].font = Font(bold=True)
ws_reality[f"A{reality_row}"].border = thin_border
ws_reality[f"B{reality_row}"].border = thin_border
ws_reality[f"A{reality_row}"].alignment = Alignment(vertical="top", wrap_text=True)
ws_reality[f"B{reality_row}"].alignment = Alignment(vertical="top", wrap_text=True)

ws_reality.column_dimensions["A"].width = 35
ws_reality.column_dimensions["B"].width = 100
ws_reality.freeze_panes = "A4"

# ---------------- Save Workbook ----------------
excel_buffer = io.BytesIO()
wb.save(excel_buffer)
excel_buffer.seek(0)

# ---------------- PowerPoint ----------------
prs = Presentation()

# Title Slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
slide.shapes.title.text = "TurboPitch Investor Deck"
slide.placeholders[1].text = "AI-generated investor materials with financial projections and reality checks"

# Financial Overview Slide
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "Financial Overview"

tx_box = slide.shapes.add_textbox(PPTInches(0.6), PPTInches(1.3), PPTInches(4.2), PPTInches(2.2))
tf = tx_box.text_frame
tf.word_wrap = True

financial_points = [
    f"Year 1 Revenue: ${projection_df.loc[0, 'Revenue']:,.0f}",
    f"Year 3 Revenue: ${projection_df.loc[2, 'Revenue']:,.0f}",
    f"Year 3 Net Income: ${projection_df.loc[2, 'Net Income']:,.0f}",
    f"Year 3 Ending Cash: ${projection_df.loc[2, 'Ending Cash']:,.0f}",
    f"Gross Margin: {projection_df.loc[2, 'Gross Margin %']:.1%}",
]

for i, point in enumerate(financial_points):
    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
    p.text = point
    p.font.size = Pt(18)

ppt_chart = create_ppt_financial_chart_image(projection_df)
slide.shapes.add_picture(ppt_chart, PPTInches(5.0), PPTInches(1.2), width=PPTInches(4.3))

# Reality Check Slide
slide = prs.slides.add_slide(prs.slide_layouts[5])
slide.shapes.title.text = "Reality Engine Summary"

box = slide.shapes.add_textbox(PPTInches(0.7), PPTInches(1.4), PPTInches(8.5), PPTInches(4.8))
tf = box.text_frame
tf.word_wrap = True

overall_p = tf.paragraphs[0]
overall_p.text = f"Overall: {reality_engine_output['overall']} - {reality_engine_output['summary']}"
overall_p.font.size = Pt(20)
overall_p.font.bold = True

for label, item in reality_engine_output["checks"].items():
    p = tf.add_paragraph()
    p.text = f"{label}: {item['status']} - {item['message']}"
    p.font.size = Pt(15)

# Pitch Deck Content Slides
for title, bullets in deck_slides:
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = title

    content_box = slide.shapes.add_textbox(PPTInches(0.8), PPTInches(1.5), PPTInches(8.4), PPTInches(4.8))
    tf = content_box.text_frame
    tf.word_wrap = True

    if bullets:
        for i, bullet in enumerate(bullets[:6]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = bullet
            p.font.size = Pt(20)
            p.level = 0
    else:
        tf.paragraphs[0].text = "No content available."

ppt_buffer = io.BytesIO()
prs.save(ppt_buffer)
ppt_buffer.seek(0)

# ---------------- Downloads UI ----------------
with tab8:
    st.download_button(
        label="Download Business Plan (Word)",
        data=doc_buffer,
        file_name="turbopitch_business_plan.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    st.download_button(
        label="Download Financial Model (Excel)",
        data=excel_buffer,
        file_name="turbopitch_financial_model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="Download Pitch Deck (PowerPoint)",
        data=ppt_buffer,
        file_name="turbopitch_pitch_deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

