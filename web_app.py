import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ACJ Company Dashboard", layout="wide")

# Hide scrollbars
st.markdown("""
<style>
::-webkit-scrollbar {
    display: none;
}
html {
    scrollbar-width: none; /* Firefox */
    -ms-overflow-style: none; /* IE and Edge */
    overflow: hidden;
}
body {
    overflow: hidden;
    overflow-x: hidden;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Load Excel outputs
# -----------------------------
df = pd.read_excel("HR_Analysis_Output.xlsx", sheet_name=None)
df_raw = pd.read_excel("HR Cleaned Data 01.09.26.xlsx", sheet_name="Data")

st.title("ACJ Company Dashboard")

# -----------------------------
# Year selector
# -----------------------------
years = [2020, 2021, 2022, 2023, 2024, 2025]
with st.spinner("Loading data..."):
    selected_year = st.radio("", years, horizontal=True)

# -----------------------------
# Tabs
# -----------------------------
tabs = st.tabs([
    "Workforce Profile & Demographics",
    "Attrition & Retention",
    "Career Progression",
    "Survey & Feedback Analytics",
    "Predictive & Diagnostic"
])

# -----------------------------
# Workforce Profile & Demographics
# -----------------------------
with tabs[0]:
    st.subheader("Headcount Overview")

    # Sheets
    tenure = df["Tenure Analysis"]
    resign = df["Resignation Trends"]
    hc = df["Headcount Per Year"]

    # Filter by selected year
    tenure_year = tenure[tenure["Year"] == selected_year]
    resign_year = resign[resign["Year"] == selected_year]
    hc_year = hc[hc["Year"] == selected_year]

    # Active employees = sum of Count from Tenure Analysis
    active_count = int(tenure_year["Count"].sum()) if not tenure_year.empty else 0

    # Leavers = LeaverCount from Resignation Trends
    leaver_count = int(resign_year["LeaverCount"].sum()) if not resign_year.empty else 0

    # Total headcount = Active + Leavers
    total_headcount = active_count + leaver_count

    # 📦 Summary metrics first
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Headcount", total_headcount)
    col2.metric("Active Employees", active_count)
    col3.metric("Leavers", leaver_count)

    # 📊 Headcount Trend chart below metrics
    st.markdown("### Headcount Trend")
    fig = px.line(hc, x="Year", y="Headcount", markers=True)
    st.plotly_chart(fig, use_container_width=True, height=120)

    # -----------------------------
    # Second row: Age, Gender, Tenure side by side
    # -----------------------------
    colA, colB, colC = st.columns(3)

    # Age Distribution
    with colA:
        st.markdown("### Age Distribution")
        age = df["Age Distribution"]
        age_year = age[age["Year"] == selected_year]
        avg_age = round(age_year["Age"].mean(), 1) if not age_year.empty else 0
        median_age = float(age_year["Age"].median()) if not age_year.empty else 0

        a1, a2 = st.columns(2)
        a1.metric("Average Age", avg_age)
        a2.metric("Median Age", median_age)

        if "Generation" in age_year.columns:
            fig = px.histogram(age_year, x="Age", y="Count", color="Generation", barmode="group",
                               title=f"Age Distribution ({selected_year})")
        else:
            fig = px.histogram(age_year, x="Age", y="Count", title=f"Age Distribution ({selected_year})")
        st.plotly_chart(fig, use_container_width=True, height=120)

    # Gender Diversity
    with colB:
        st.markdown("### Gender Diversity")
        gender = df["Gender Diversity"]
        gender_year = gender[gender["Year"] == selected_year]
        gender_counts = gender_year.groupby("Gender")["Count"].sum()

        gcols = st.columns(len(gender_counts))
        for i, (g, c) in enumerate(gender_counts.items()):
            gcols[i].metric(f"{g} Employees", int(c))

        fig = px.bar(gender_year, x="Position/Level", y="Count", color="Gender", barmode="stack",
                     title=f"Gender Diversity ({selected_year})")
        st.plotly_chart(fig, use_container_width=True, height=120)

    # Tenure Analysis
    with colC:
        st.markdown("### Tenure Analysis")
        avg_tenure = round(tenure_year["Tenure"].mean(), 1) if not tenure_year.empty else 0
        median_tenure = float(tenure_year["Tenure"].median()) if not tenure_year.empty else 0
        max_tenure = float(tenure_year["Tenure"].max()) if not tenure_year.empty else 0

        t1, t2, t3 = st.columns(3)
        t1.metric("Average Tenure", f"{avg_tenure} yrs")
        t2.metric("Median Tenure", f"{median_tenure} yrs")
        t3.metric("Longest Tenure", f"{max_tenure} yrs")

        fig = px.scatter(tenure, x="Tenure", y="Count", color="YearJoined", size="Count",
                         title="Tenure Analysis Across Years")
        st.plotly_chart(fig, use_container_width=True, height=120)

# -----------------------------
# Attrition & Retention
# -----------------------------
with tabs[1]:
    st.subheader("Attrition & Retention Overview")

    # --- Ensure Calendar Year is normalized ---
    df_raw["Calendar Year"] = pd.to_datetime(df_raw["Calendar Year"], errors="coerce").dt.year
    df_raw = df_raw.dropna(subset=["Calendar Year"])
    df_raw["Calendar Year"] = df_raw["Calendar Year"].astype(int)

    # --- Create ResignedFlag and Retention ---
    def to_resigned_flag(x):
        s = str(x).strip().upper()
        if s in {"LEAVER", "YES", "TRUE", "1"}:
            return 1
        try:
            return 1 if float(s) == 1.0 else 0
        except:
            return 0

    if "ResignedFlag" not in df_raw.columns:
        df_raw["ResignedFlag"] = df_raw["Resignee Checking"].apply(to_resigned_flag)
        df_raw["Retention"] = 1 - df_raw["ResignedFlag"]

    # --- Filter by selected year ---
    retention_year = df_raw[df_raw["Calendar Year"] == int(selected_year)]

    if not retention_year.empty:
        total_employees = len(retention_year)
        total_resigned = int(retention_year["ResignedFlag"].sum())
        retention_rate = retention_year["Retention"].mean() * 100
    else:
        total_employees = 0
        total_resigned = 0
        retention_rate = 0

    # --- Headline metrics ---
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Employees", total_employees)
    col2.metric("Resigned", total_resigned)
    col3.metric("Retention Rate", f"{retention_rate:.1f}%")

    # --- Attrition Trends ---
    st.markdown("## Attrition Trends")
    attrition_summary = df_raw.groupby("Calendar Year")["ResignedFlag"].sum().reset_index()
    fig = px.line(attrition_summary, x="Calendar Year", y="ResignedFlag",
                  markers=True, title="Resignations per Year")
    st.plotly_chart(fig, width="stretch", height=300)

    # --- Retention by Generation ---
    st.markdown("## Retention by Generation")
    gen_retention = df_raw.groupby("Generation")["Retention"].mean().reset_index()
    fig = px.bar(gen_retention, x="Generation", y="Retention",
                 title="Retention Rate by Generation")
    st.plotly_chart(fig, width="stretch", height=300)


# ----------------------------- # 
# Career Progression (employee data from df_raw) # 
# -----------------------------
with tabs[2]: 
    st.subheader("Career Progression Overview") 
    # Ensure numeric conversion for Promotion & Transfer 
    # Handles '1', '0', 'Yes', 'No', True/False 
    def to_num(x): 
        s = str(x).strip().upper() 
        if s in {"1", "YES", "TRUE"}: 
            return 1 
        if s in {"0", "NO", "FALSE"}: 
            return 0 
        try: 
            return float(s) 
        except: 
            return pd.NA 
    df_raw["Promotion & Transfer"] = df_raw["Promotion & Transfer"].apply(to_num) 
    career_year = df_raw[df_raw["Calendar Year"] == selected_year] 
    if not career_year.empty: 
        total_promotions_transfers = int(pd.Series(career_year["Promotion & Transfer"]).fillna(0).sum()) 
        avg_tenure = career_year["Tenure"].mean() 
        career_satisfaction = career_year["Career"].mean() 
    else: 
        total_promotions_transfers = 0 
        avg_tenure = 0 
        career_satisfaction = 0 
    col1, col2, col3 = st.columns(3) 
    col1.metric("Promotions & Transfers", total_promotions_transfers) 
    col2.metric("Average Tenure", f"{avg_tenure:.1f} yrs") 
    col3.metric("Career Satisfaction", f"{career_satisfaction:.1f}") 
    st.markdown("## Promotion & Transfer Tracking") 
    promo_summary = ( 
        df_raw.groupby("Calendar Year", as_index=False)["Promotion & Transfer"] 
        .sum() 
    ) 
    fig = px.bar( 
        promo_summary, x="Calendar Year", y="Promotion & Transfer", title="Promotions & Transfers per Year" 
    ) 
    st.plotly_chart(fig, width="stretch", height=250) 
    fig = px.bar( 
        df_raw.groupby(["Calendar Year", "Position/Level"], as_index=False)["Promotion & Transfer"].sum(), 
        x="Calendar Year", y="Promotion & Transfer", color="Position/Level", title="Promotions & Transfers by Position/Level" 
    ) 
    st.plotly_chart(fig, width="stretch", height=250) 
    st.markdown("## Promotion Predictors") 
    fig = px.scatter( 
        df_raw, x="Tenure", y="Promotion & Transfer", color="Generation", title="Promotion & Transfer Likelihood by Tenure & Generation", opacity=0.6 
    ) 
    st.plotly_chart(fig, width="stretch", height=250) 
    fig = px.scatter( 
        df_raw, x="Career", y="Promotion & Transfer", color="Position/Level", title="Career Satisfaction vs Promotion & Transfer", opacity=0.6 
    ) 
    st.plotly_chart(fig, width="stretch", height=250)

# -----------------------------
# Survey & Feedback Analytics
# -----------------------------
with tabs[3]:
    st.subheader("Satisfaction Heatmap")
    heat = df["Satisfaction Prct"]
    heat_year = heat[heat["Year"] == selected_year].reset_index(drop=True)
    heat_year.index = [""] * len(heat_year)
    st.dataframe(heat_year, use_container_width=True)

    fig = px.imshow(
        heat.set_index("Year").T,
        aspect="auto",
        color_continuous_scale="Blues",
        title="Satisfaction Heatmap"
    )
    st.plotly_chart(fig, width="stretch")

    st.subheader("Engagement Index")
    ei = df["Engagement Index"]
    ei_year = ei[ei["Year"] == selected_year].reset_index(drop=True)
    ei_year.index = [""] * len(ei_year)
    st.dataframe(ei_year, use_container_width=True)

    fig = px.line(
        ei, x="Year", y="Engagement Score",
        markers=True, title="Engagement Index Trend"
    )
    st.plotly_chart(fig, width="stretch")

    st.subheader("Driver Analysis (Resignation)")
    driver_res = df["Driver-Resignation"]
    driver_res_year = driver_res[driver_res["Year"] == selected_year].reset_index(drop=True)
    driver_res_year.index = [""] * len(driver_res_year)
    st.dataframe(driver_res_year, use_container_width=True)

    fig = px.bar(
        driver_res_year,
        x="Category", y="Correlation with Resignation",
        color="Year",
        title=f"Driver Analysis - Resignation ({selected_year})"
    )
    st.plotly_chart(fig, width="stretch")

    st.subheader("Driver Analysis (Promotion)")
    driver_prom = df["Driver-Promotion"]
    driver_prom_year = driver_prom[driver_prom["Year"] == selected_year].reset_index(drop=True)
    driver_prom_year.index = [""] * len(driver_prom_year)
    st.dataframe(driver_prom_year, use_container_width=True)

    fig = px.bar(
        driver_prom_year,
        x="Category", y="Correlation with Promotion",
        color="Year",
        title=f"Driver Analysis - Promotion ({selected_year})"
    )
    st.plotly_chart(fig, width="stretch")

# ----------------------------- #
# Predictive & Diagnostic #
# ----------------------------- #
with tabs[4]:
    st.subheader("Attrition Risk Modeling")
    # If you saved risk scores in Attrition_Risk_Output.xlsx, load and merge here
    # risk = pd.read_excel("Attrition_Risk_Output.xlsx", sheet_name="Attrition Risk Scores")
    # risk_year = risk[risk["Year"] == selected_year].reset_index(drop=True)
    # risk_year.index = [""] * len(risk_year)
    # st.dataframe(risk_year, use_container_width=True)

    st.subheader("Engagement vs Retention")
    evr = df["Engagement vs Retention"]
    evr_year = evr[evr["Year"] == selected_year].reset_index(drop=True)

    # Hide index numbers
    evr_year.index = [""] * len(evr_year)

    # Display clean table
    st.dataframe(evr_year, use_container_width=True)

    # Plot Engagement vs Retention
    fig = px.bar(
        evr_year,
        x="Resignee Checking ",
        y="Avg Engagement Score",
        color="Resignee Checking ",
        title=f"Engagement vs Retention ({selected_year})"
    )
    st.plotly_chart(fig, width="stretch")
