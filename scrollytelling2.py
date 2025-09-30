# scrollytelling2.py
# Requirements: streamlit, pandas, sqlalchemy, pymysql, plotly
# Install (if needed): pip3 install streamlit pandas sqlalchemy pymysql plotly

import streamlit as st
import pandas as pd
import plotly.express as px
from sqlalchemy import create_engine
from urllib.parse import quote_plus

st.set_page_config(page_title="Carbon Emissions (DB)", layout="wide")

# ---------------------------
# Database connection (uses st.secrets["mysql"])
def get_engine():
    try:
        db = st.secrets["mysql"]
    except Exception:
        st.error("Database credentials not found in st.secrets['mysql']. Create .streamlit/secrets.toml.")
        raise

    user = db.get("user") or db.get("username") or db.get("uname")
    raw_pwd = db.get("password") or db.get("pw") or ""
    host = db.get("host", "localhost")
    port = db.get("port", 3306)
    database = db.get("database") or db.get("db") or db.get("dbname")

    pwd = quote_plus(raw_pwd)  # URL-encode
    conn_str = f"mysql+pymysql://{user}:{pwd}@{host}:{port}/{database}"
    return create_engine(conn_str, pool_pre_ping=True)

# --- Initialize engine ---
engine = get_engine()

# --- Load data with time_periods.label for month/year ---
query = """
SELECT 
    f.*, 
    tp.label AS month_year
FROM v_fuel_buildings_emissions f
JOIN time_periods tp 
    ON f.time_period_id = tp.period_id
"""
df = pd.read_sql(query, engine)

# ---------------------------
# Filters (Dropdowns but multi-select)
# ---------------------------
st.sidebar.header("Filters")

scope_filter = st.sidebar.multiselect(
    "Select Scope(s):",
    options=df["scope"].unique(),
    default=df["scope"].unique()
)

month_filter = st.sidebar.multiselect(
    "Select Month/Year(s):",
    options=df["month_year"].unique(),
    default=df["month_year"].unique()
)

# Apply filters
df_filtered = df[
    (df["scope"].isin(scope_filter)) &
    (df["month_year"].isin(month_filter))
]

# --- KPIs ---
c1, c2, c3 = st.columns(3)
c1.metric("Total Emissions (tCO₂e)", f"{df_filtered['co2_tonnes'].sum():,.2f}")
c2.metric("Scopes Covered", f"{df_filtered['scope'].nunique()}")
c3.metric("Months Covered", f"{df_filtered['month_year'].nunique()}")

# --- Aggregation by month for charts ---
by_month = (
    df_filtered.groupby("month_year", as_index=False)
               .agg({"co2_tonnes": "sum"})
               .sort_values("month_year")
)

st.line_chart(by_month.set_index("month_year")["co2_tonnes"])
