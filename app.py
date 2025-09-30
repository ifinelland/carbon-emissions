# app.py
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Carbon Emissions Dashboard", layout="wide")

@st.cache_data
def load_data(file):
    df = pd.read_csv(file, parse_dates=["date"])
    # Expect columns: date, month, scope, category, activity_amount, unit, emission_factor_kgco2e_per_unit
    if "emissions_kgco2e" not in df.columns:
        df["emissions_kgco2e"] = df["activity_amount"] * df["emission_factor_kgco2e_per_unit"]
    # normalize scope labels
    df["scope"] = df["scope"].astype(str)
    # month key
    if "month" not in df.columns:
        df["month"] = df["date"].dt.strftime("%Y-%m")
    return df

st.title("🌍 Total Carbon Emissions Dashboard")
st.caption("Upload your activity data and view total emissions by month, scope, and category.")

with st.sidebar:
    st.header("📤 Data")
    up = st.file_uploader("Upload CSV", type=["csv"])
    st.markdown("Or try the sample file above if you don’t have one yet.")
    st.markdown("[Download sample CSV](sandbox:/mnt/data/emissions_sample.csv)")

if up is not None:
    df = load_data(up)
else:
    # fallback to sample (only for quick demo)
    df = load_data("emissions_sample.csv")

# Filters
with st.sidebar:
    st.header("🔎 Filters")
    months = sorted(df["month"].unique())
    scopes = sorted(df["scope"].unique())
    categories = sorted(df["category"].unique())

    sel_months = st.multiselect("Months", months, default=months)
    sel_scopes = st.multiselect("Scopes", scopes, default=scopes)
    sel_cats = st.multiselect("Categories", categories, default=categories)

f = df[df["month"].isin(sel_months) & df["scope"].isin(sel_scopes) & df["category"].isin(sel_cats)]

# Aggregations
total_tCO2e = f["emissions_kgco2e"].sum() / 1000
by_scope = f.groupby("scope", as_index=False)["emissions_kgco2e"].sum()
by_scope["tCO2e"] = by_scope["emissions_kgco2e"] / 1000

by_month = f.groupby("month", as_index=False)["emissions_kgco2e"].sum()
by_month["tCO2e"] = by_month["emissions_kgco2e"] / 1000

by_cat = f.groupby(["scope","category"], as_index=False)["emissions_kgco2e"].sum()
by_cat["tCO2e"] = by_cat["emissions_kgco2e"] / 1000

# KPIs
c1, c2, c3 = st.columns(3)
c1.metric("Total Emissions (tCO₂e)", f"{total_tCO2e:,.2f}")
c2.metric("Scopes Covered", f"{f['scope'].nunique()}")
c3.metric("Months Covered", f"{f['month'].nunique()}")

# Charts
st.subheader("📈 Emissions Trend (Total)")
st.line_chart(by_month.set_index("month")["tCO2e"])

st.subheader("📊 Emissions by Scope")
st.bar_chart(by_scope.set_index("scope")["tCO2e"])

st.subheader("🏷️ Emissions by Category (within scope)")
sel_scope_for_cats = st.selectbox("Select scope", options=sorted(by_cat["scope"].unique()))
st.bar_chart(by_cat[by_cat["scope"] == sel_scope_for_cats].set_index("category")["tCO2e"])

# Tables + Export
st.subheader("🧾 Downloadable Tables")
tab1, tab2, tab3 = st.tabs(["By Month", "By Scope", "By Category"])

with tab1:
    st.dataframe(by_month[["month","tCO2e"]].rename(columns={"tCO2e":"tCO2e_total"}))
    st.download_button("Download by-month CSV",
                       data=by_month.to_csv(index=False),
                       file_name="emissions_by_month.csv",
                       mime="text/csv")

with tab2:
    st.dataframe(by_scope[["scope","tCO2e"]].rename(columns={"tCO2e":"tCO2e_total"}))
    st.download_button("Download by-scope CSV",
                       data=by_scope.to_csv(index=False),
                       file_name="emissions_by_scope.csv",
                       mime="text/csv")

with tab3:
    out = by_cat[["scope","category","tCO2e"]].rename(columns={"tCO2e":"tCO2e_total"})
    st.dataframe(out)
    st.download_button("Download by-category CSV",
                       data=out.to_csv(index=False),
                       file_name="emissions_by_category.csv",
                       mime="text/csv")

st.caption("Tip: Replace the sample emission factors with your country/market-specific, source-of-truth values (e.g., grid EF, fuel EF).")
