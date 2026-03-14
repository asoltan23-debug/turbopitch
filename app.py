import streamlit as st
import pandas as pd
from openai import OpenAI

import os
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

st.set_page_config(page_title="TurboPitch AI", layout="wide")

st.title("TurboPitch AI")
st.subheader("Turn an idea into investor-ready materials — then update everything when assumptions change.")

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

st.sidebar.header("Core Assumptions")
idea = st.text_area("Describe your business idea", value=st.session_state.idea, height=120)
price_per_unit = st.sidebar.number_input("Price per unit ($)", min_value=0.01, value=float(st.session_state.price_per_unit), step=0.5)
year1_units = st.sidebar.number_input("Year 1 units sold", min_value=1, value=int(st.session_state.year1_units), step=1000)
growth_y2 = st.sidebar.slider("Year 2 growth rate", min_value=0.0, max_value=2.0, value=float(st.session_state.growth_y2), step=0.05)
growth_y3 = st.sidebar.slider("Year 3 growth rate", min_value=0.0, max_value=2.0, value=float(st.session_state.growth_y3), step=0.05)

st.session_state.idea = idea
st.session_state.price_per_unit = price_per_unit
st.session_state.year1_units = year1_units
st.session_state.growth_y2 = growth_y2
st.session_state.growth_y3 = growth_y3

col1, col2 = st.columns([1, 1])

with col1:
    if st.button("Generate Investor Materials"):
        projection_df = build_projection(price_per_unit, year1_units, growth_y2, growth_y3)

        prompt = f"""
You are TurboPitch AI, an investor-readiness assistant.

A founder has this business idea:
{idea}

Current assumptions:
- Price per unit: ${price_per_unit}
- Year 1 units sold: {year1_units}
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

with col2:
    adjustment_pct = st.number_input("Reduce revenue assumptions by %", min_value=0, max_value=90, value=25, step=5)
    investor_feedback = st.text_area(
        "Investor feedback",
        value="Investor says revenue assumptions are too aggressive."
    )

    if st.button("Apply Investor Pushback"):
        new_year1_units = int(year1_units * (1 - adjustment_pct / 100))
        projection_df = build_projection(price_per_unit, new_year1_units, growth_y2, growth_y3)

        prompt = f"""
You are TurboPitch AI, an investor-readiness assistant.

Original business idea:
{idea}

Investor feedback:
{investor_feedback}

Updated assumptions after pushback:
- Price per unit: ${price_per_unit}
- Revised Year 1 units sold: {new_year1_units}
- Year 2 growth rate: {growth_y2:.0%}
- Year 3 growth rate: {growth_y3:.0%}
- Revenue assumption reduced by: {adjustment_pct}%

Updated projected revenues:
- Year 1: ${projection_df.iloc[0]['Revenue']:,.0f}
- Year 2: ${projection_df.iloc[1]['Revenue']:,.0f}
- Year 3: ${projection_df.iloc[2]['Revenue']:,.0f}

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

st.divider()

projection_df = build_projection(
    st.session_state.price_per_unit,
    st.session_state.year1_units,
    st.session_state.growth_y2,
    st.session_state.growth_y3
)

tab1, tab2, tab3 = st.tabs(["AI Output", "Financial Table", "Revenue Chart"])

with tab1:
    if st.session_state.last_output:
        st.markdown(st.session_state.last_output)
    else:
        st.info("Generate investor materials to see the output here.")

with tab2:
    st.dataframe(projection_df, use_container_width=True)

with tab3:
    chart_df = projection_df.set_index("Year")[["Revenue"]]
    st.line_chart(chart_df)