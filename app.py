import os
import streamlit as st
import pandas as pd
from openai import OpenAI

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(page_title="TurboPitch AI", layout="wide")

st.title("TurboPitch AI")
st.subheader("Turn an idea into investor-ready materials — then pressure-test the assumptions.")

# -----------------------------
# Session state defaults
# -----------------------------
if "idea" not in st.session_state:
    st.session_state.idea = ""
if "price_per_unit" not in st.session_state:
    st.session_state.price_per_unit = 3.0
if "year1_units" not in st.session_state:
    st.session_state.year1_units = 1000000
if "growth_y2" not in st.session_state:
    st.session_state.growth_y2 = 0.50
if "growth_y3" not in st.session_state:
    st.session_state.growth_y3 = 0.33
if "last_output" not in st.session_state:
    st.session_state.last_output = ""
if "sanity_output" not in st.session_state:
    st.session_state.sanity_output = ""


# -----------------------------
# Helper functions
# -----------------------------
def build_projection(price, year1_units, growth_y2, growth_y3):
    year2_units = int(year1_units * (1 + growth_y2))
    year3_units = int(year2_units * (1 + growth_y3))

    year1_revenue = year1_units * price
    year2_revenue = year2_units * price
    year3_revenue = year3_units * price

    df = pd.DataFrame({
        "Year": ["Year 1", "Year 2", "Year 3"],
        "Units Sold": [year1_units, year2_units, year3_units],
        "Revenue": [year1_revenue, year2_revenue, year3_revenue]
    })
    return df


def classify_business_type(idea_text: str) -> str:
    text = idea_text.lower()

    hardware_keywords = [
        "device", "robot", "robotic", "physical", "mower", "car", "drone",
        "consumer product", "appliance", "equipment", "machine", "hardware"
    ]
    software_keywords = [
        "software", "saas", "platform", "app", "ai tool", "dashboard",
        "subscription", "cloud", "crm", "workflow"
    ]
    services_keywords = [
        "consulting", "agency", "service", "bookkeeping", "marketing service",
        "advisory", "outsourcing", "staffing"
    ]

    if any(word in text for word in hardware_keywords):
        return "hardware"
    if any(word in text for word in software_keywords):
        return "software"
    if any(word in text for word in services_keywords):
        return "services"
    return "general"


def run_rule_based_sanity_checks(idea_text, price, year1_units, growth_y2, growth_y3):
    business_type = classify_business_type(idea_text)
    checks = []

    # General volume check
    if year1_units >= 1000000:
        checks.append(("red", "Year 1 unit volume is extremely high for a brand-new business and likely unrealistic without existing distribution, capital, and brand awareness."))
    elif year1_units >= 100000:
        checks.append(("yellow", "Year 1 unit volume is aggressive and may require strong evidence such as channel partnerships, preorders, or paid distribution capacity."))
    else:
        checks.append(("green", "Year 1 unit volume is within a more plausible early-stage range."))

    # Growth checks
    if growth_y2 > 1.0 or growth_y3 > 1.0:
        checks.append(("red", "A growth rate above 100% year-over-year is very aggressive and would need strong traction proof."))
    elif growth_y2 > 0.7 or growth_y3 > 0.7:
        checks.append(("yellow", "Growth above 70% is possible, but investors will likely ask for proof of channel, demand, or retention strength."))
    else:
        checks.append(("green", "Growth assumptions are in a more defensible range."))

    # Business-type-specific checks
    if business_type == "hardware":
        if price < 50:
            checks.append(("yellow", "For a hardware product, this price may be too low unless manufacturing and logistics costs are extremely cheap."))
        if year1_units > 50000:
            checks.append(("red", "For a new hardware business, shipping more than 50,000 units in Year 1 is usually difficult without major retail access, capital, and operations."))
        else:
            checks.append(("green", "The hardware distribution assumption is not obviously impossible, though investors will still expect supply chain proof."))

    elif business_type == "software":
        if price < 5:
            checks.append(("yellow", "Software pricing below $5 may make customer acquisition economics difficult unless the model is high-volume or ad-supported."))
        elif price > 1000:
            checks.append(("yellow", "High software pricing can work for enterprise sales, but investors will expect proof of value, sales cycle, and target buyer."))
        else:
            checks.append(("green", "Software pricing is in a plausible range depending on the customer segment."))

    elif business_type == "services":
        if year1_units > 10000:
            checks.append(("red", "For a services business, very high unit volume is likely unrealistic unless 'units' means something automated or low-touch."))
        else:
            checks.append(("green", "Service volume is not obviously unrealistic, depending on the delivery model."))

    # Revenue magnitude
    year1_revenue = price * year1_units
    if year1_revenue > 100000000:
        checks.append(("red", "Year 1 revenue above $100M is highly unlikely for a newly launched company without an existing base, major contracts, or a proven distribution engine."))
    elif year1_revenue > 10000000:
        checks.append(("yellow", "Year 1 revenue above $10M is ambitious and should be justified with credible go-to-market assumptions."))

    return checks, business_type


def render_check(level, message):
    if level == "red":
        st.error(f"🔴 {message}")
    elif level == "yellow":
        st.warning(f"🟡 {message}")
    else:
        st.success(f"🟢 {message}")


# -----------------------------
# Sidebar inputs
# -----------------------------
st.sidebar.header("Core Assumptions")

idea = st.text_area("Describe your business idea", value=st.session_state.idea, height=140)

price_per_unit = st.sidebar.number_input(
    "Price per unit ($)",
    min_value=0.01,
    value=float(st.session_state.price_per_unit),
    step=0.5
)

year1_units = st.sidebar.number_input(
    "Year 1 units sold",
    min_value=1,
    value=int(st.session_state.year1_units),
    step=1000
)

growth_y2 = st.sidebar.slider(
    "Year 2 growth rate",
    min_value=0.0,
    max_value=2.0,
    value=float(st.session_state.growth_y2),
    step=0.05
)

growth_y3 = st.sidebar.slider(
    "Year 3 growth rate",
    min_value=0.0,
    max_value=2.0,
    value=float(st.session_state.growth_y3),
    step=0.05
)

st.session_state.idea = idea
st.session_state.price_per_unit = price_per_unit
st.session_state.year1_units = year1_units
st.session_state.growth_y2 = growth_y2
st.session_state.growth_y3 = growth_y3

projection_df = build_projection(price_per_unit, year1_units, growth_y2, growth_y3)
rule_checks, business_type = run_rule_based_sanity_checks(
    idea, price_per_unit, year1_units, growth_y2, growth_y3
)

# -----------------------------
# Top action buttons
# -----------------------------
col1, col2, col3 = st.columns([1, 1, 1])

with col1:
    if st.button("Run Sanity Check"):
        prompt = f"""
You are a skeptical startup investor and financial reviewer.

Business idea:
{idea}

Business type guess: {business_type}

Assumptions:
- Price per unit: ${price_per_unit}
- Year 1 units sold: {year1_units:,}
- Year 2 growth rate: {growth_y2:.0%}
- Year 3 growth rate: {growth_y3:.0%}

Projected revenues:
- Year 1: ${projection_df.iloc[0]['Revenue']:,.0f}
- Year 2: ${projection_df.iloc[1]['Revenue']:,.0f}
- Year 3: ${projection_df.iloc[2]['Revenue']:,.0f}

Give me:
1. Investor Reality Check
2. Most Likely Unrealistic Assumptions
3. What Needs Justification
4. Suggested More Credible Ranges

Be direct, skeptical, concise, and practical.
"""
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        st.session_state.sanity_output = response.choices[0].message.content

with col2:
    if st.button("Generate Investor Materials"):
        prompt = f"""
You are TurboPitch AI, an investor-readiness assistant.

A founder has this business idea:
{idea}

Business type guess: {business_type}

Current assumptions:
- Price per unit: ${price_per_unit}
- Year 1 units sold: {year1_units:,}
- Year 2 growth rate: {growth_y2:.0%}
- Year 3 growth rate: {growth_y3:.0%}

Projected revenues:
- Year 1: ${projection_df.iloc[0]['Revenue']:,.0f}
- Year 2: ${projection_df.iloc[1]['Revenue']:,.0f}
- Year 3: ${projection_df.iloc[2]['Revenue']:,.0f}

Generate these sections with headings:
1. Executive Summary
2. Key Assumptions
3. 3-Year Revenue Projection Summary
4. Investor Risks
5. Pitch Deck Outline

Be analytical and credible, not hypey.
"""
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        st.session_state.last_output = response.choices[0].message.content

with col3:
    adjustment_pct = st.number_input(
        "Reduce revenue assumptions by %",
        min_value=0,
        max_value=90,
        value=25,
        step=5
    )
    investor_feedback = st.text_area(
        "Investor feedback",
        value="Investor says revenue assumptions are too aggressive."
    )

    if st.button("Apply Investor Pushback"):
        new_year1_units = int(year1_units * (1 - adjustment_pct / 100))
        updated_projection_df = build_projection(price_per_unit, new_year1_units, growth_y2, growth_y3)

        prompt = f"""
You are TurboPitch AI, an investor-readiness assistant.

Original business idea:
{idea}

Investor feedback:
{investor_feedback}

Updated assumptions after pushback:
- Price per unit: ${price_per_unit}
- Revised Year 1 units sold: {new_year1_units:,}
- Year 2 growth rate: {growth_y2:.0%}
- Year 3 growth rate: {growth_y3:.0%}
- Revenue assumption reduced by: {adjustment_pct}%

Updated projected revenues:
- Year 1: ${updated_projection_df.iloc[0]['Revenue']:,.0f}
- Year 2: ${updated_projection_df.iloc[1]['Revenue']:,.0f}
- Year 3: ${updated_projection_df.iloc[2]['Revenue']:,.0f}

Generate these sections with headings:
1. Impact of Investor Pushback
2. Revised Executive Summary
3. Revised Key Assumptions
4. Revised Investor Risks
5. What Must Change in the Pitch Deck

Be skeptical, commercially realistic, and concise.
"""
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )

        st.session_state.last_output = response.choices[0].message.content
        st.session_state.year1_units = new_year1_units
        projection_df = updated_projection_df

st.divider()

# -----------------------------
# Tabs
# -----------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "Sanity Check",
    "AI Output",
    "Financial Table",
    "Revenue Chart"
])

with tab1:
    st.markdown("### Rule-Based Assumption Check")
    st.caption(f"Detected business type: **{business_type.title()}**")

    for level, message in rule_checks:
        render_check(level, message)

    st.markdown("### AI Investor Reality Check")
    if st.session_state.sanity_output:
        st.markdown(st.session_state.sanity_output)
    else:
        st.info("Click 'Run Sanity Check' to get AI feedback on whether the assumptions look too aggressive.")

with tab2:
    if st.session_state.last_output:
        st.markdown(st.session_state.last_output)
    else:
        st.info("Generate investor materials to see the output here.")

with tab3:
    st.dataframe(projection_df, use_container_width=True)

with tab4:
    chart_df = projection_df.set_index("Year")[["Revenue"]]
    st.line_chart(chart_df, use_container_width=True)
