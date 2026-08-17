import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import re
import html


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="CareerVista",
    page_icon="📍",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown("""
<style>

    /* Main application */
    .stApp {
        background: #f7f9fc;
    }

    /* Hide the Streamlit deploy toolbar so page content starts at the top. */
    header[data-testid="stHeader"] {
        display: none;
    }

    [data-testid="stAppViewContainer"] {
        background: #f7f9fc;
    }

    .block-container {
        padding-top: 2.5rem;
        padding-bottom: 3rem;
        max-width: 1400px;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e5eaf2;
    }

    section[data-testid="stSidebar"] * {
        color: #172033;
    }

    section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {
        color: #667085;
    }

    section[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background: #eef5ff;
        border-radius: 8px;
    }

    /* Main headings */
    .main-title {
        font-size: 46px;
        font-weight: 750;
        color: #172033;
        line-height: 1.1;
        margin-bottom: 10px;
    }

    .subtitle {
        font-size: 18px;
        color: #667085;
        line-height: 1.6;
        margin-bottom: 25px;
    }

    /* Hero section */
    .hero {
        background: linear-gradient(
            120deg,
            #172033 0%,
            #253b63 55%,
            #326b78 100%
        );
        border-radius: 24px;
        padding: 45px 50px;
        color: white;
        margin-bottom: 30px;
    }

    .hero h1 {
        font-size: 44px;
        margin: 0;
        font-weight: 750;
    }

    .hero p {
        font-size: 18px;
        color: #dce5f2;
        max-width: 720px;
        line-height: 1.6;
    }

    /* Section title */
    .section-title {
        font-size: 28px;
        font-weight: 700;
        color: #172033;
        margin-top: 30px;
        margin-bottom: 15px;
    }

    /* Metric cards */
    .metric-card {
        background: white;
        padding: 22px;
        border-radius: 16px;
        border: 1px solid #e5e9f0;
        min-height: 125px;
        box-shadow: 0 6px 18px rgba(30, 40, 60, 0.05);
    }

    .metric-label {
        color: #667085;
        font-size: 14px;
        margin-bottom: 8px;
    }

    .metric-value {
        color: #172033;
        font-size: 28px;
        font-weight: 750;
    }

    /* Opportunity cards */
    .job-card {
        background: white;
        padding: 20px 22px;
        border-radius: 15px;
        border: 1px solid #e4e8ef;
        margin-bottom: 12px;
        transition: 0.2s;
    }

    .job-card:hover {
        border-color: #4f8f9d;
        box-shadow: 0 5px 18px rgba(30, 40, 60, 0.08);
    }

    .job-title {
        color: #172033;
        font-size: 18px;
        font-weight: 700;
    }

    .company {
        color: #4f8f9d;
        font-weight: 600;
        margin-top: 4px;
    }

    .job-meta {
        color: #667085;
        font-size: 14px;
        margin-top: 8px;
    }

    /* About page */
    .info-card {
        background: #ffffff;
        border: 1px solid #e5e9f0;
        border-radius: 16px;
        padding: 24px;
        min-height: 170px;
        box-shadow: 0 6px 18px rgba(30, 40, 60, 0.05);
    }

    .info-card h3 {
        color: #172033;
        font-size: 19px;
        margin: 0 0 10px;
    }

    .info-card p {
        color: #667085;
        font-size: 15px;
        line-height: 1.65;
        margin: 0;
    }

    .impact-banner {
        background: linear-gradient(120deg, #e9f6f2 0%, #eef5ff 100%);
        border: 1px solid #d8e8e3;
        border-radius: 18px;
        color: #172033;
        padding: 28px 32px;
        margin: 26px 0;
    }

    .impact-banner h2 {
        font-size: 25px;
        margin: 0 0 8px;
    }

    .impact-banner p {
        color: #526176;
        font-size: 16px;
        line-height: 1.65;
        margin: 0;
    }

    /* Location banner */
    .location-banner {
        background: #eaf5f4;
        border-left: 5px solid #4f8f9d;
        padding: 16px 20px;
        border-radius: 10px;
        margin: 20px 0;
        color: #172033;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #8a94a6;
        font-size: 13px;
        padding: 35px 0 10px 0;
    }

    /* Buttons */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
    }

    /* Selectbox */
    div[data-baseweb="select"] > div {
        border-radius: 10px;
    }

    /* Dataframe */
    .stDataFrame {
        border-radius: 12px;
    }

</style>
""", unsafe_allow_html=True)


# =========================================================
# PATHS
# =========================================================

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data" / "raw"


# =========================================================
# DATA LOADING
# =========================================================

def load_jobs():

    job_file = DATA_DIR / "indian_jobs.xlsx"

    if not job_file.exists():
        st.error(
            f"Job dataset not found.\n\n"
            f"Expected file:\n{job_file}"
        )
        st.stop()

    # Load only fields used by the dashboard. In particular, exclude the
    # large job-description field so location changes stay memory-efficient.
    df = pd.read_excel(
        job_file,
        usecols=[
            "title", "companyName", "tagsAndSkills", "experience", "jobUploaded",
            "salary", "location", "minimumSalary", "maximumSalary",
            "minimumExperience", "maximumExperience",
        ],
    )

    return df


def load_unemployment():

    # Look for CSV files in raw data folder
    csv_files = list(DATA_DIR.glob("*.csv"))

    if not csv_files:
        return pd.DataFrame()

    # Select the first CSV
    unemployment_file = csv_files[0]

    df = pd.read_csv(unemployment_file)

    return df


# =========================================================
# CLEAN JOB DATA
# =========================================================

def prepare_jobs(df):

    # Standardise column names
    df.columns = [str(c).strip() for c in df.columns]

    # Remove duplicate rows
    df.drop_duplicates(inplace=True)

    # Salary conversion
    for col in ["minimumSalary", "maximumSalary"]:

        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )

    # Average salary
    if "minimumSalary" in df.columns and "maximumSalary" in df.columns:

        df["average_salary"] = (
            df["minimumSalary"] +
            df["maximumSalary"]
        ) / 2

    # Experience conversion
    for col in [
        "minimumExperience",
        "maximumExperience"
    ]:

        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col],
                errors="coerce"
            )

    if (
        "minimumExperience" in df.columns
        and "maximumExperience" in df.columns
    ):

        df["average_experience"] = (
            df["minimumExperience"] +
            df["maximumExperience"]
        ) / 2

    # Fresher flag
    if "minimumExperience" in df.columns:

        df["is_fresher_job"] = (
            df["minimumExperience"].fillna(99) <= 0
        )

    return df


# =========================================================
# CLEAN UNEMPLOYMENT DATA
# =========================================================

def prepare_unemployment(df):

    if df.empty:
        return df

    df = df.copy()

    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # Rename common dataset columns
    rename_map = {}

    for col in df.columns:

        clean = col.lower().strip()

        if "region" in clean:
            rename_map[col] = "region"

        elif "date" == clean:
            rename_map[col] = "date"

        elif "unemployment rate" in clean:
            rename_map[col] = "unemployment_rate"

        elif "labour participation" in clean:
            rename_map[col] = "labour_participation_rate"

        elif clean == "area":
            rename_map[col] = "area"

    df = df.rename(columns=rename_map)

    if "unemployment_rate" in df.columns:

        df["unemployment_rate"] = pd.to_numeric(
            df["unemployment_rate"],
            errors="coerce"
        )

    return df


@st.cache_resource(show_spinner="Preparing employment data...")
def get_prepared_jobs():
    """Load and prepare the jobs dataset once per server process."""
    return prepare_jobs(load_jobs())


@st.cache_resource(show_spinner="Preparing unemployment data...")
def get_prepared_unemployment():
    """Load and prepare the unemployment dataset once per server process."""
    return prepare_unemployment(load_unemployment())


jobs = get_prepared_jobs()
unemployment = get_prepared_unemployment()


@st.cache_resource(show_spinner="Analysing skill demand...")
def get_skill_demand():
    """Calculate reusable, market-wide skill demand counts."""
    if "tagsAndSkills" not in jobs.columns:
        return pd.DataFrame(columns=["Skill", "Demand"])

    skills = (
        jobs["tagsAndSkills"].dropna().astype(str).str.lower()
        .str.split(",").explode().str.strip()
    )
    skills = skills[skills.str.len() > 2]
    demand = skills.value_counts().reset_index()
    demand.columns = ["Skill", "Demand"]
    return demand


# =========================================================
# SESSION STATE
# =========================================================

if "selected_location" not in st.session_state:
    st.session_state.selected_location = ""


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            font-size:26px;
            font-weight:750;
            margin-bottom:5px;
        ">
        💼 CareerVista
        </div>
        """,
        unsafe_allow_html=True
    )

    st.caption("Employment intelligence for India")

    st.divider()

    page = st.radio(
        "Navigate",
        [
            "Home",
            "Career Explorer",
            "Unemployment Insights",
            "Market Analytics",
            "Skill Gap Checker",
            "Regional Comparison",
            "About Us"
        ]
    )

    st.divider()

    st.markdown("### About")

    st.write(
        "This platform analyses employment data "
        "to identify job opportunities, skill demand "
        "and unemployment patterns across locations."
    )

    st.caption("Career data platform")


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def format_salary(value):

    if pd.isna(value) or value <= 0:
        return "Not disclosed"

    if value >= 100000:

        return f"₹{value/100000:.1f} LPA"

    return f"₹{value:,.0f}"


def render_metric_card(label, value):
    """Render a compact metric card without multiline HTML parsing."""
    st.markdown(
        '<div class="metric-card"><div class="metric-label">'
        f'{html.escape(str(label))}</div><div class="metric-value">'
        f'{html.escape(str(value))}</div></div>',
        unsafe_allow_html=True,
    )


def render_job_card(job):
    """Render an opportunity without using Streamlit's DataFrame component."""
    title = html.escape(str(job.get("title", "Opportunity")))
    company = html.escape(str(job.get("companyName", "Company not listed")))
    details = [
        str(job.get(column)).strip()
        for column in ("experience", "salary", "tagsAndSkills")
        if pd.notna(job.get(column)) and str(job.get(column)).strip()
    ]
    meta = " · ".join(html.escape(detail) for detail in details)
    st.markdown(
        f'<div class="job-card"><div class="job-title">{title}</div>'
        f'<div class="company">{company}</div>'
        f'<div class="job-meta">{meta}</div></div>',
        unsafe_allow_html=True,
    )


def get_location_list():

    if "location" not in jobs.columns:
        return []

    locations = (
        jobs["location"]
        .dropna()
        .astype(str)
        .str.strip()
    )

    locations = locations[
        locations.str.len() > 1
    ]

    return sorted(
        locations.unique()
    )


locations = get_location_list()


# =========================================================
# HOME PAGE
# =========================================================

if page == "Home":

    st.markdown(
        '<div class="hero"><h1>Find opportunities<br>where you live.</h1>'
        '<p>Explore employment opportunities, skills in demand, salary trends '
        'and unemployment patterns across different locations in India.</p></div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-title">Start exploring</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns([3, 1])

    with col1:

        location_input = st.selectbox(
            "Choose your location",
            ["Select a location"] + locations,
            index=0
        )

    with col2:

        st.write("")
        st.write("")

        if st.button(
            "Explore opportunities →",
            use_container_width=True
        ):

            if location_input != "Select a location":

                st.session_state.selected_location = location_input
                st.success(
                    f"Showing opportunities for {location_input}"
                )

            else:

                st.warning(
                    "Please select a location first."
                )


    # -----------------------------------------------------
    # OVERVIEW METRICS
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">What the platform provides</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        render_metric_card("Job Listings", f"{len(jobs):,}")

    with c2:

        unique_locations = (
            jobs["location"].nunique()
            if "location" in jobs.columns
            else 0
        )

        render_metric_card("Locations", f"{unique_locations:,}")

    with c3:

        unique_companies = (
            jobs["companyName"].nunique()
            if "companyName" in jobs.columns
            else 0
        )

        render_metric_card("Companies", f"{unique_companies:,}")

    with c4:

        if "is_fresher_job" in jobs.columns:

            fresher_percent = (
                jobs["is_fresher_job"].mean() * 100
            )

        else:

            fresher_percent = 0

        render_metric_card("Fresher Friendly", f"{fresher_percent:.1f}%")


    # -----------------------------------------------------
    # HOW IT WORKS
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">How it works</div>',
        unsafe_allow_html=True
    )

    a, b, c = st.columns(3)

    with a:

        st.markdown(
            """
            ### 01 · Choose a location

            Select the city or location you're interested in
            and discover the opportunities available there.
            """
        )

    with b:

        st.markdown(
            """
            ### 02 · Explore the market

            Compare job availability, salary information,
            experience requirements and skills.
            """
        )

    with c:

        st.markdown(
            """
            ### 03 · Identify opportunities

            Find fresher-friendly jobs and understand which
            skills are most demanded in your area.
            """
        )


# =========================================================
# LOCAL OPPORTUNITIES
# =========================================================

elif page == "Career Explorer":

    st.markdown(
        '<div class="main-title">Career Explorer</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">'
        'Discover jobs and skills in demand near your chosen location.'
        '</div>',
        unsafe_allow_html=True
    )

    location = st.selectbox(
        "📍 Select location",
        locations,
        index=(
            locations.index(
                st.session_state.selected_location
            )
            if st.session_state.selected_location in locations
            else 0
        )
    )

    st.session_state.selected_location = location

    # Filter jobs
    location_key = location.strip().casefold()
    local_jobs = jobs[
        jobs["location"]
        .astype("string")
        .str.strip()
        .str.casefold()
        .eq(location_key)
    ].copy()


    st.markdown(
        f"""
        <div class="location-banner">
            <strong>{location}</strong> ·
            {len(local_jobs):,} job listings found
        </div>
        """,
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # LOCAL METRICS
    # -----------------------------------------------------

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        render_metric_card("Local Jobs", f"{len(local_jobs):,}")

    with c2:

        if "companyName" in local_jobs.columns:

            companies = local_jobs["companyName"].nunique()

        else:

            companies = 0

        render_metric_card("Hiring Companies", f"{companies:,}")

    with c3:

        fresher_count = (
            local_jobs["is_fresher_job"].sum()
            if "is_fresher_job" in local_jobs.columns
            else 0
        )

        render_metric_card("Fresher Jobs", f"{fresher_count:,}")

    with c4:

        if "average_salary" in local_jobs.columns:

            salary = local_jobs[
                local_jobs["average_salary"] > 0
            ]["average_salary"].median()

        else:

            salary = 0

        render_metric_card("Median Advertised Salary", format_salary(salary))


    # -----------------------------------------------------
    # CHARTS
    # -----------------------------------------------------

    chart1, chart2 = st.columns(2)

    with chart1:

        if "companyName" in local_jobs.columns:

            company_counts = (
                local_jobs["companyName"]
                .value_counts()
                .head(10)
                .reset_index()
            )

            company_counts.columns = [
                "Company",
                "Jobs"
            ]

            fig = px.bar(
                company_counts,
                x="Jobs",
                y="Company",
                orientation="h",
                title="Companies with Most Listings"
            )

            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#f5f7fb",
                plot_bgcolor="#f5f7fb",
                margin=dict(l=10, r=10, t=50, b=20),
                height=400
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


    with chart2:

        if "is_fresher_job" in local_jobs.columns:

            fresher_data = pd.DataFrame({
                "Category": [
                    "Fresher Friendly",
                    "Experienced / Other"
                ],
                "Jobs": [
                    int(local_jobs["is_fresher_job"].sum()),
                    int(
                        len(local_jobs)
                        -
                        local_jobs["is_fresher_job"].sum()
                    )
                ]
            })

            fig = px.pie(
                fresher_data,
                names="Category",
                values="Jobs",
                hole=0.55,
                title="Experience Accessibility"
            )

            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#f5f7fb",
                plot_bgcolor="#f5f7fb",
                margin=dict(l=10, r=10, t=50, b=20),
                height=400
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


    # -----------------------------------------------------
    # JOB LISTINGS
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">Latest opportunities</div>',
        unsafe_allow_html=True
    )

    display_columns = [
        "title",
        "companyName",
        "experience",
        "salary",
        "tagsAndSkills"
    ]

    display_columns = [
        col
        for col in display_columns
        if col in local_jobs.columns
    ]

    latest_jobs = local_jobs[display_columns].head(20)

    if latest_jobs.empty:
        st.info("No opportunities are available for this location yet.")
    else:
        for _, job in latest_jobs.iterrows():
            render_job_card(job)


    # -----------------------------------------------------
    # SKILLS
    # -----------------------------------------------------

    st.markdown(
        '<div class="section-title">Skills employers are looking for</div>',
        unsafe_allow_html=True
    )

    if "tagsAndSkills" in local_jobs.columns:

        skills = (
            local_jobs["tagsAndSkills"]
            .dropna()
            .astype(str)
            .str.lower()
            .str.split(",")
            .explode()
            .str.strip()
        )

        skills = skills[
            skills.str.len() > 2
        ]

        top_local_skills = (
            skills
            .value_counts()
            .head(15)
            .reset_index()
        )

        top_local_skills.columns = [
            "Skill",
            "Demand"
        ]

        fig = px.bar(
            top_local_skills,
            x="Demand",
            y="Skill",
            orientation="h",
            title=f"Top Skills in {location}"
        )

        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="#f5f7fb",
            plot_bgcolor="#f5f7fb",
            height=500
        )

        st.plotly_chart(
            fig,
            use_container_width=True
        )


# =========================================================
# UNEMPLOYMENT INSIGHTS
# =========================================================

elif page == "Unemployment Insights":

    st.markdown(
        '<div class="main-title">Unemployment Insights</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">'
        'Understand unemployment patterns and labour participation '
        'across regions.'
        '</div>',
        unsafe_allow_html=True
    )

    if unemployment.empty:

        st.warning(
            "Unemployment dataset could not be found."
        )

    else:

        # -------------------------------------------------
        # LOCATION
        # -------------------------------------------------

        selected_region = st.selectbox(
            "Select region",
            sorted(
                unemployment["region"]
                .dropna()
                .unique()
            )
            if "region" in unemployment.columns
            else []
        )


        region_data = unemployment[
            unemployment["region"] == selected_region
        ].copy()


        # -------------------------------------------------
        # METRICS
        # -------------------------------------------------

        c1, c2, c3 = st.columns(3)

        avg_unemployment = (
            region_data["unemployment_rate"].mean()
            if "unemployment_rate" in region_data.columns
            else 0
        )

        avg_labour = (
            region_data["labour_participation_rate"].mean()
            if "labour_participation_rate" in region_data.columns
            else 0
        )

        records = len(region_data)


        with c1:
            render_metric_card("Average Unemployment", f"{avg_unemployment:.2f}%")


        with c2:
            render_metric_card("Labour Participation", f"{avg_labour:.2f}%")


        with c3:
            render_metric_card("Data Records", f"{records:,}")


        # -------------------------------------------------
        # UNEMPLOYMENT TREND
        # -------------------------------------------------

        if (
            "date" in region_data.columns
            and "unemployment_rate" in region_data.columns
        ):

            region_data["date"] = pd.to_datetime(
                region_data["date"],
                errors="coerce"
            )

            region_data = region_data.sort_values("date")

            fig = px.line(
                region_data,
                x="date",
                y="unemployment_rate",
                markers=True,
                title=f"Unemployment Trend — {selected_region}"
            )

            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#f5f7fb",
                plot_bgcolor="#f5f7fb",
                height=420
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


        # -------------------------------------------------
        # AREA COMPARISON
        # -------------------------------------------------

        if (
            "area" in unemployment.columns
            and "unemployment_rate" in unemployment.columns
        ):

            area_data = (
                unemployment
                .groupby("area")["unemployment_rate"]
                .mean()
                .reset_index()
            )

            fig = px.pie(
                area_data,
                names="area",
                values="unemployment_rate",
                hole=0.5,
                title="Average Unemployment by Area"
            )

            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#f5f7fb",
                plot_bgcolor="#f5f7fb",
                height=400
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


# =========================================================
# MARKET ANALYTICS
# =========================================================

elif page == "Market Analytics":

    st.markdown(
        '<div class="main-title">Market Analytics</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subtitle">Explore broader hiring patterns across locations, salaries, experience levels and skills.</div>',
        unsafe_allow_html=True,
    )

    chart_left, chart_right = st.columns(2)

    with chart_left:
        if "average_salary" in jobs.columns:
            salary_data = jobs.loc[
                jobs["average_salary"].gt(0), "average_salary"
            ].dropna()

            if not salary_data.empty:
                fig = px.histogram(
                    salary_data,
                    x="average_salary",
                    nbins=30,
                    title="Distribution of Advertised Salaries",
                    labels={"average_salary": "Average salary", "count": "Job listings"},
                    color_discrete_sequence=["#326b78"],
                )
                fig.update_layout(
                    template="plotly_white",
                    paper_bgcolor="#f7f9fc",
                    plot_bgcolor="#f7f9fc",
                    margin=dict(l=10, r=10, t=50, b=20),
                    height=400,
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Salary information is not available in the current dataset.")

    with chart_right:
        if "average_experience" in jobs.columns:
            experience_data = jobs.loc[
                jobs["average_experience"].ge(0), "average_experience"
            ].dropna()

            if not experience_data.empty:
                fig = px.histogram(
                    experience_data,
                    x="average_experience",
                    nbins=15,
                    title="Experience Requirements Across Listings",
                    labels={"average_experience": "Average years of experience", "count": "Job listings"},
                    color_discrete_sequence=["#4f8f9d"],
                )
                fig.update_layout(
                    template="plotly_white",
                    paper_bgcolor="#f7f9fc",
                    plot_bgcolor="#f7f9fc",
                    margin=dict(l=10, r=10, t=50, b=20),
                    height=400,
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Experience information is not available in the current dataset.")

    chart_left, chart_right = st.columns(2)

    with chart_left:
        if {"average_salary", "average_experience"}.issubset(jobs.columns):
            comparison_data = jobs.loc[
                jobs["average_salary"].gt(0)
                & jobs["average_experience"].ge(0),
                ["average_salary", "average_experience", "title"],
            ].dropna().head(1500)

            if not comparison_data.empty:
                fig = px.scatter(
                    comparison_data,
                    x="average_experience",
                    y="average_salary",
                    hover_name="title",
                    opacity=0.55,
                    title="Salary and Experience Relationship",
                    labels={
                        "average_experience": "Average years of experience",
                        "average_salary": "Average salary",
                    },
                    color_discrete_sequence=["#253b63"],
                )
                fig.update_layout(
                    template="plotly_white",
                    paper_bgcolor="#f7f9fc",
                    plot_bgcolor="#f7f9fc",
                    margin=dict(l=10, r=10, t=50, b=20),
                    height=420,
                )
                st.plotly_chart(fig, use_container_width=True)

    with chart_right:
        if "tagsAndSkills" in jobs.columns:
            market_skills = (
                jobs["tagsAndSkills"].dropna().astype(str).str.lower()
                .str.split(",").explode().str.strip()
            )
            market_skills = market_skills[market_skills.str.len() > 2]
            top_market_skills = market_skills.value_counts().head(18).reset_index()
            top_market_skills.columns = ["Skill", "Demand"]

            if not top_market_skills.empty:
                fig = px.treemap(
                    top_market_skills,
                    path=["Skill"],
                    values="Demand",
                    color="Demand",
                    color_continuous_scale="Teal",
                    title="Most In-Demand Skills",
                )
                fig.update_layout(
                    paper_bgcolor="#f7f9fc",
                    margin=dict(l=10, r=10, t=50, b=20),
                    height=420,
                )
                st.plotly_chart(fig, use_container_width=True)

    if "jobUploaded" in jobs.columns:
        posting_dates = pd.to_datetime(jobs["jobUploaded"], errors="coerce")
        monthly_jobs = (
            posting_dates.dropna().to_frame(name="date")
            .set_index("date").resample("ME").size()
            .reset_index(name="Listings")
        )

        if not monthly_jobs.empty:
            fig = px.area(
                monthly_jobs,
                x="date",
                y="Listings",
                title="Job Posting Activity Over Time",
                labels={"date": "Posting month", "Listings": "Job listings"},
                color_discrete_sequence=["#326b78"],
            )
            fig.update_layout(
                template="plotly_white",
                paper_bgcolor="#f7f9fc",
                plot_bgcolor="#f7f9fc",
                height=420,
            )
            st.plotly_chart(fig, use_container_width=True)


# =========================================================
# SKILL GAP CHECKER
# =========================================================

elif page == "Skill Gap Checker":

    st.markdown(
        '<div class="main-title">Skill Gap Checker</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subtitle">Select the skills you already have to see how they align with current job-market demand.</div>',
        unsafe_allow_html=True,
    )

    skill_demand = get_skill_demand()
    recommended_skills = skill_demand.head(100)["Skill"].tolist()
    selected_skills = st.multiselect(
        "Your current skills",
        options=recommended_skills,
        placeholder="Choose one or more skills",
    )

    if not selected_skills:
        st.info("Select your skills to receive a personalised market-readiness summary.")
    else:
        selected_set = {skill.casefold() for skill in selected_skills}
        priority_skills = skill_demand.head(15).copy()
        priority_skills["Status"] = np.where(
            priority_skills["Skill"].isin(selected_set),
            "You have this skill",
            "Recommended next skill",
        )
        matched = int((priority_skills["Status"] == "You have this skill").sum())
        missing_skills = priority_skills.loc[
            priority_skills["Status"] == "Recommended next skill", "Skill"
        ].tolist()

        metric_one, metric_two, metric_three = st.columns(3)
        with metric_one:
            st.metric("Skills selected", len(selected_skills))
        with metric_two:
            st.metric("Top-demand skills matched", matched)
        with metric_three:
            st.metric("Recommended next skills", len(missing_skills))

        st.markdown(
            '<div class="section-title">Your market skill alignment</div>',
            unsafe_allow_html=True,
        )
        fig = px.bar(
            priority_skills.sort_values("Demand"),
            x="Demand",
            y="Skill",
            color="Status",
            orientation="h",
            title="Top skills to strengthen for better job alignment",
            color_discrete_map={
                "You have this skill": "#326b78",
                "Recommended next skill": "#d99a3a",
            },
        )
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="#f7f9fc",
            plot_bgcolor="#f7f9fc",
            margin=dict(l=10, r=10, t=50, b=20),
            height=500,
            legend_title_text="",
        )
        st.plotly_chart(fig, use_container_width=True)

        selected_pattern = "|".join(
            re.escape(skill) for skill in selected_skills
        )
        matching_jobs = jobs[
            jobs["tagsAndSkills"]
            .astype("string")
            .str.contains(selected_pattern, case=False, na=False, regex=True)
        ].copy()

        st.markdown(
            '<div class="section-title">Companies hiring for your skills</div>',
            unsafe_allow_html=True,
        )

        if matching_jobs.empty or "companyName" not in matching_jobs.columns:
            st.info("No hiring companies were found for the selected skills.")
        else:
            company_jobs = matching_jobs.dropna(subset=["companyName"]).copy()
            company_jobs["companyName"] = company_jobs["companyName"].astype(str).str.strip()
            company_jobs = company_jobs[company_jobs["companyName"].ne("")]

            companies = (
                company_jobs.groupby("companyName")
                .agg(
                    Listings=("companyName", "size"),
                    Top_role=(
                        "title",
                        lambda titles: titles.dropna().astype(str).value_counts().index[0]
                        if not titles.dropna().empty
                        else "Role not listed",
                    ),
                )
                .reset_index()
                .sort_values("Listings", ascending=False)
                .head(12)
            )

            for _, company in companies.iterrows():
                company_name = html.escape(str(company["companyName"]))
                listings = int(company["Listings"])
                role = html.escape(str(company["Top_role"]))
                st.markdown(
                    f'<div class="job-card"><div class="job-title">{company_name}</div>'
                    f'<div class="company">{listings:,} matching listings</div>'
                    f'<div class="job-meta">Common matching role: {role}</div></div>',
                    unsafe_allow_html=True,
                )

        if missing_skills:
            st.markdown(
                '<div class="section-title">Suggested learning focus</div>',
                unsafe_allow_html=True,
            )
            st.write(
                "Prioritise: " + ", ".join(skill.title() for skill in missing_skills[:6])
            )


# =========================================================
# REGIONAL COMPARISON
# =========================================================

elif page == "Regional Comparison":

    st.markdown(
        '<div class="main-title">Regional Comparison</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subtitle">Compare unemployment and labour participation across two regions.</div>',
        unsafe_allow_html=True,
    )

    required_columns = {"region", "unemployment_rate", "labour_participation_rate"}
    regions = (
        sorted(unemployment["region"].dropna().unique())
        if "region" in unemployment.columns
        else []
    )

    if unemployment.empty or not required_columns.issubset(unemployment.columns):
        st.warning("Regional comparison data is not available in the current dataset.")
    elif len(regions) < 2:
        st.warning("At least two regions are required for a comparison.")
    else:
        selector_one, selector_two = st.columns(2)
        with selector_one:
            region_one = st.selectbox("First region", regions, key="region_one")
        with selector_two:
            remaining_regions = [region for region in regions if region != region_one]
            region_two = st.selectbox("Second region", remaining_regions, key="region_two")

        comparison_rows = []
        for region in [region_one, region_two]:
            data = unemployment[unemployment["region"] == region]
            comparison_rows.append({
                "Region": region,
                "Average unemployment": data["unemployment_rate"].mean(),
                "Labour participation": data["labour_participation_rate"].mean(),
                "Records": len(data),
            })
        comparison = pd.DataFrame(comparison_rows)

        first_metric, second_metric, third_metric = st.columns(3)
        with first_metric:
            st.metric(
                "Lower unemployment region",
                comparison.loc[comparison["Average unemployment"].idxmin(), "Region"],
            )
        with second_metric:
            unemployment_gap = abs(
                comparison.loc[0, "Average unemployment"]
                - comparison.loc[1, "Average unemployment"]
            )
            st.metric("Unemployment gap", f"{unemployment_gap:.2f}%")
        with third_metric:
            participation_gap = abs(
                comparison.loc[0, "Labour participation"]
                - comparison.loc[1, "Labour participation"]
            )
            st.metric("Participation gap", f"{participation_gap:.2f}%")

        comparison_chart = comparison.melt(
            id_vars="Region",
            value_vars=["Average unemployment", "Labour participation"],
            var_name="Measure",
            value_name="Percentage",
        )
        fig = px.bar(
            comparison_chart,
            x="Region",
            y="Percentage",
            color="Measure",
            barmode="group",
            title="Unemployment and Labour Participation Comparison",
            color_discrete_map={
                "Average unemployment": "#d99a3a",
                "Labour participation": "#326b78",
            },
        )
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="#f7f9fc",
            plot_bgcolor="#f7f9fc",
            height=440,
            yaxis_title="Percentage (%)",
            legend_title_text="",
        )
        st.plotly_chart(fig, use_container_width=True)

        if "date" in unemployment.columns:
            trend_data = unemployment[
                unemployment["region"].isin([region_one, region_two])
            ][["region", "date", "unemployment_rate"]].copy()
            trend_data["date"] = pd.to_datetime(trend_data["date"], errors="coerce")
            trend_data = trend_data.dropna(subset=["date", "unemployment_rate"])
            trend_data = (
                trend_data.groupby(["region", "date"], as_index=False)["unemployment_rate"]
                .mean().sort_values("date")
            )

            if not trend_data.empty:
                fig = px.line(
                    trend_data,
                    x="date",
                    y="unemployment_rate",
                    color="region",
                    markers=True,
                    title="Unemployment Trends Over Time",
                    labels={
                        "date": "Date",
                        "unemployment_rate": "Unemployment rate (%)",
                        "region": "Region",
                    },
                    color_discrete_sequence=["#326b78", "#d99a3a"],
                )
                fig.update_layout(
                    template="plotly_white",
                    paper_bgcolor="#f7f9fc",
                    plot_bgcolor="#f7f9fc",
                    height=420,
                )
                st.plotly_chart(fig, use_container_width=True)


# =========================================================
# ABOUT US
# =========================================================

elif page == "About Us":

    st.markdown(
        '<div class="main-title">About CareerVista</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="subtitle">Career intelligence that helps people discover clearer paths to meaningful work.</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="impact-banner"><h2>Our mission</h2><p>CareerVista brings local job-market information, in-demand skills and unemployment trends into one accessible platform. It is designed to turn complex data into practical career decisions.</p></div>',
        unsafe_allow_html=True,
    )

    objective, method, outcome = st.columns(3)

    with objective:
        st.markdown(
            '<div class="info-card"><h3>Our objective</h3><p>Make employment information easy to understand, so job seekers can identify realistic opportunities in their location and prepare for the skills employers value.</p></div>',
            unsafe_allow_html=True,
        )

    with method:
        st.markdown(
            '<div class="info-card"><h3>How we help</h3><p>We combine job listings with salary, experience and skill-demand data, alongside regional unemployment and labour participation insights, to reveal where support and opportunity are needed.</p></div>',
            unsafe_allow_html=True,
        )

    with outcome:
        st.markdown(
            '<div class="info-card"><h3>The result</h3><p>People can focus their job search, build relevant skills and pursue fresher-friendly roles with more confidence. The platform supports better connections between local talent and hiring demand.</p></div>',
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="section-title">How CareerVista addresses unemployment</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="info-card"><h3>1. Makes local demand visible</h3><p>Job seekers can explore available roles, employers and opportunities near them instead of relying on scattered information.</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="info-card"><h3>2. Connects skills to opportunities</h3><p>Skill-demand trends show users what to learn or strengthen, helping reduce the gap between candidate capabilities and employer needs.</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="info-card"><h3>3. Supports informed action</h3><p>Unemployment insights reveal regional patterns, while job data helps users make targeted choices about where and how to apply.</p></div>',
        unsafe_allow_html=True,
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">
        CareerVista · Employment & Unemployment Analytics
        <br>
        Helping people make informed career decisions
    </div>
    """,
    unsafe_allow_html=True
)
