import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Carbon Emissions Story", layout="wide")

# --- Load data ---
@st.cache_data
def load_data(file="emissions_sample.csv"):
    df = pd.read_csv(file, parse_dates=["date"])
    if "emissions_kgco2e" not in df.columns:
        df["emissions_kgco2e"] = df["activity_amount"] * df["emission_factor_kgco2e_per_unit"]
    df["month"] = df["date"].dt.to_period("M").astype(str)
    df["scope"] = df["scope"].astype(str)
    return df

df = load_data()

# Pre-aggregated
by_month = df.groupby("month", as_index=False)["emissions_kgco2e"].sum()
by_month["tCO2e"] = by_month["emissions_kgco2e"] / 1000

by_scope = df.groupby("scope", as_index=False)["emissions_kgco2e"].sum()
by_scope["tCO2e"] = by_scope["emissions_kgco2e"] / 1000

# --- STORY START ---
st.title("🌍 A Year of Carbon Emissions")
st.markdown("""
Welcome to our interactive **scrollytelling dashboard**.  
Scroll down to uncover how our emissions evolved during 2024, and what drives them.
""")

# Section 1
st.header("1️⃣ Monthly Emissions Trend")
st.markdown("Emissions fluctuate across the year. The chart below shows the **total emissions (tCO₂e) per month**.")

fig_trend = px.line(
    by_month,
    x="month",
    y="tCO2e",
    markers=True,
    title="Monthly Total Emissions (tCO₂e)"
)
fig_trend.update_traces(line_color="green", line_width=3)
fig_trend.update_layout(yaxis_title="tCO₂e", xaxis_title="Month")
st.plotly_chart(fig_trend, use_container_width=True)

# Section 2
st.header("2️⃣ Emissions by Scope")
st.markdown("Breaking it down by **Scope 1, 2, and 3**, we see which categories dominate.")

fig_scope = px.bar(
    by_scope,
    x="scope",
    y="tCO2e",
    text_auto=".2f",
    color="scope",
    title="Total Emissions by Scope"
)
fig_scope.update_layout(yaxis_title="tCO₂e", xaxis_title="Scope")
st.plotly_chart(fig_scope, use_container_width=True)

# Section 3
st.header("3️⃣ Category Breakdown")
st.markdown("Within each scope, different **categories** drive the total. Select a scope below.")

scope_choice = st.selectbox("Choose a scope:", sorted(df["scope"].unique()))
by_cat = (
    df[df["scope"] == scope_choice]
    .groupby("category", as_index=False)["emissions_kgco2e"].sum()
)
by_cat["tCO2e"] = by_cat["emissions_kgco2e"] / 1000

fig_cat = px.bar(
    by_cat.sort_values("tCO2e", ascending=False),
    x="category",
    y="tCO2e",
    text_auto=".2f",
    title=f"Emissions by Category (Scope {scope_choice})"
)
fig_cat.update_layout(yaxis_title="tCO₂e", xaxis_title="Category")
st.plotly_chart(fig_cat, use_container_width=True)

# Section 4 - Treemap
st.header("4️⃣ Treemap of Emissions (Scope → Category)")
st.markdown("This treemap shows emissions by **scope and category** hierarchically.")

fig_treemap = px.treemap(
    df,
    path=["scope", "category"],
    values="emissions_kgco2e",
    color="scope",
    title="Treemap: Scope → Category"
)
st.plotly_chart(fig_treemap, use_container_width=True)

# Section 5 - Sunburst
st.header("5️⃣ Sunburst of Emissions (Scope → Category)")
st.markdown("The sunburst chart is another way to explore the **hierarchy of emissions**.")

fig_sunburst = px.sunburst(
    df,
    path=["scope", "category"],
    values="emissions_kgco2e",
    color="scope",
    title="Sunburst: Scope → Category"
)
st.plotly_chart(fig_sunburst, use_container_width=True)

# Section 6 - Insights
st.header("6️⃣ Key Insights")
st.markdown("""
✨ **Highlights from the data:**
- **Scope 3** (e.g. flights, commuting) often dominates total emissions.  
- **Scope 2** (electricity) depends heavily on usage patterns and grid intensity.  
- Treemaps and sunbursts make it easy to spot the **largest categories at a glance**.  
- Reduction opportunities: electrification of fleet, travel demand management, energy efficiency.  
""")

