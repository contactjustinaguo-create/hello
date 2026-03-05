import streamlit as st
import pandas as pd
import altair as alt




# Creating 1 joined sheet from the Excel file


# Separating Excel file into 2 sheets 
st.title("BFSI Walkthrough Impact Dashboard")

file_path = "Ops SWAT Take Home March 2026.xlsx"

hubspot = pd.read_excel(file_path, sheet_name="HubSpot_Mock")
projects = pd.read_excel(file_path, sheet_name="Monday_Projects")


# Fill in company id for fuzzy company names in projects to match the company ids in hubspot
hubspot_companies = hubspot[["company_id", "company_name"]].drop_duplicates()

# Compare project_name against all hubspot company_name values character by character.
# Eliminate candidates that do not match at each position.
# Return the single remaining row's company_id if exactly one remains, else None.
def infer_company_id_from_name(project_name: str, hubspot_df: pd.DataFrame):

    candidates = hubspot_df.copy()

    # Compare character by character
    for i, ch in enumerate(project_name):
        # Keep only names that are long enough and match at this position
        candidates = candidates[
            candidates["company_name"].apply(lambda x: len(x) > i and x[i] == ch)
        ]

        # Stop early if exactly one candidate remains
        if len(candidates) == 1:
            return candidates.iloc[0]["company_id"]

        # Stop if none remain
        if len(candidates) == 0:
            return None

    # If loop ends and exactly one remains, use it
    if len(candidates) == 1:
        return candidates.iloc[0]["company_id"]

    # Otherwise ambiguous
    return None

# Fill missing project company_id values
missing_mask = projects["company_id"].isna()

for idx in projects[missing_mask].index:
    proj_name = projects.at[idx, "company_name"]
    matched_company_id = infer_company_id_from_name(proj_name, hubspot_companies)

    if matched_company_id is not None:
        projects.at[idx, "company_id"] = matched_company_id

# Join HubSpot and Projects using company_id
joined = hubspot.merge(
    projects,
    on="company_id",
    how="left",
    suffixes=("_hubspot", "_project")
)




# Helper functions for charts

# Normalize yes/no values
def to_bool(val):
    if pd.isna(val):
        return False
    val = str(val).strip().lower()
    return val in ["yes", "true", "1"]

# Lifecycle stage ranking
lifecycle_sort_order = ["Lead", "MQL", "SQL", "Opportunity"]
lifecycle_order = {
    "Lead": 1,
    "MQL": 2,
    "SQL": 3,
    "Opportunity": 4
}
joined["lifecycle_stage_rank"] = joined["lifecycle_stage"].map(lifecycle_order)

# Deal stage ranking
stage_order = {
    "Discovery": 1,
    "Evaluation": 2,
    "Security Review": 3,
    "Proposal": 4,
    "Negotiation": 5
}
security_review_rank = stage_order["Security Review"]
joined["deal_stage_rank"] = joined["deal_stage"].map(stage_order)


# Difference label
def make_difference_label(diff):
    if diff > 0:
        return "Higher for engaged"
    elif diff < 0:
        return "Lower for engaged"
    else:
        return "No difference"
    
def comparison_table(metric_name, engaged_rate, non_engaged_rate):
    diff = engaged_rate - non_engaged_rate
    return pd.DataFrame({
        "Engaged": [f"{engaged_rate:.1%}"],
        "Non-engaged": [f"{non_engaged_rate:.1%}"],
        "Engaged v. Non-engaged": [f"{diff:+.1%}"],
        "Interpretation": [make_difference_label(diff)]
    })

# Charts by lifecycle
def stage_bar_chart(df, x_col, y_col, title, color="#2F6FDE"):
    bars = (
        alt.Chart(df)
        .mark_bar(color=color)
        .encode(
            x=alt.X(
                x_col,
                title="Lifecycle Stage",
                sort=lifecycle_sort_order
            ),
            y=alt.Y(
                y_col,
                title="Rate",
                axis=alt.Axis(format="%"),
                scale=alt.Scale(domain=[0, 1])
            ),
            tooltip=[
                alt.Tooltip(x_col, title="Lifecycle Stage"),
                alt.Tooltip(y_col, title="Rate", format=".1%")
            ]
        )
        .properties(
            title={
                "text": title,
                "anchor": "middle"
            }, 
            height=320)
    )

    labels = (
        alt.Chart(df)
        .mark_text(dy=-10, fontSize=12)
        .encode(
            x=alt.X(x_col, sort=lifecycle_sort_order),
            y=alt.Y(y_col, scale=alt.Scale(domain=[0, 1])),
            text=alt.Text(y_col, format=".1%")
        )
    )

    return bars + labels

# Charts by engaged v. non-engaged
def cohort_bar_chart(df, title):
    color_scale = alt.Scale(
        domain=["Engaged", "Non-engaged"],
        range=["#2F6FDE", "#FF7857"]
    )

    bars = (
        alt.Chart(df)
        .mark_bar()
        .encode(
            x=alt.X("cohort:N", title="Cohort"),
            y=alt.Y(
                "rate:Q",
                title="Rate",
                axis=alt.Axis(format="%"),
                scale=alt.Scale(domain=[0, 1])
            ),
            color=alt.Color("cohort:N", scale=color_scale, legend=None),
            tooltip=[
                alt.Tooltip("cohort:N", title="Cohort"),
                alt.Tooltip("rate:Q", title="Rate", format=".1%")
            ]
        )
        .properties(
            title={
                "text": title,
                "anchor": "middle"
            }, 
            height=320)
    )

    labels = (
        alt.Chart(df)
        .mark_text(dy=-10, fontSize=12, color="black")
        .encode(
            x="cohort:N",
            y=alt.Y("rate:Q", scale=alt.Scale(domain=[0, 1])),
            text=alt.Text("rate:Q", format=".1%")
        )
    )

    return bars + labels

# Create abbreviate company data
company_summary = (
    joined.groupby("company_id")
    .agg(
        company_name=("company_name_hubspot", "first") if "company_name_hubspot" in joined.columns else ("company_name", "first"),
        max_lifecycle_rank=("lifecycle_stage_rank", "max"),
        walkthrough_started=("walkthrough_started", lambda x: any(to_bool(v) for v in x)),
        walkthrough_step_completed=("walkthrough_step_completed", "max"),
        justification_pack_downloaded=("justification_pack_downloaded", lambda x: any(to_bool(v) for v in x)),
        sales_meeting_booked=("sales_meeting_booked", lambda x: any(to_bool(v) for v in x)),
        max_stage_rank=("deal_stage_rank", "max"),
        deal_stage=("deal_stage", "first"),
        project_count=("project_id", lambda x: x.notna().sum())
    )
    .reset_index()
)

rank_to_lifecycle = {v: k for k, v in lifecycle_order.items()}
company_summary["lifecycle_stage"] = company_summary["max_lifecycle_rank"].map(rank_to_lifecycle)

company_summary["security_or_beyond"] = company_summary["max_stage_rank"] >= security_review_rank
company_summary["has_project"] = company_summary["project_count"] > 0

walkthrough_max_step = company_summary["walkthrough_step_completed"].fillna(0).max()
walkthrough_completion_threshold = int(walkthrough_max_step / 2)

company_summary["substantive_completion"] = (
    company_summary["walkthrough_step_completed"].fillna(0) > walkthrough_completion_threshold
)

company_summary["cohort"] = company_summary["walkthrough_started"].map(
    {True: "Engaged", False: "Non-engaged"}
)

engaged = company_summary[company_summary["walkthrough_started"]]
non_engaged = company_summary[~company_summary["walkthrough_started"]]


# Calculate 5 metrics


# Metric 1: Walkthrough engagement
engagement_rate = company_summary["walkthrough_started"].mean()

engagement_by_lifecycle = (
    company_summary.groupby("lifecycle_stage")["walkthrough_started"]
    .mean()
    .reset_index(name="rate")
)

# Metric 2: Substantive walkthrough completion
completion_rate = (
    engaged["substantive_completion"].mean() if len(engaged) > 0 else 0
)

completion_by_lifecycle = (
    company_summary[company_summary["walkthrough_started"]]
    .groupby("lifecycle_stage")["substantive_completion"]
    .mean()
    .reset_index(name="rate")
)

# Metric 3: Meeting booked
meeting_rate_engaged = engaged["sales_meeting_booked"].mean() if len(engaged) > 0 else 0
meeting_rate_non_engaged = non_engaged["sales_meeting_booked"].mean() if len(non_engaged) > 0 else 0

meeting_chart_df = pd.DataFrame({
    "cohort": ["Engaged", "Non-engaged"],
    "rate": [meeting_rate_engaged, meeting_rate_non_engaged]
})

# Metric 4: Deal stage maturity
security_rate_engaged = engaged["security_or_beyond"].mean() if len(engaged) > 0 else 0
security_rate_non_engaged = non_engaged["security_or_beyond"].mean() if len(non_engaged) > 0 else 0

security_chart_df = pd.DataFrame({
    "cohort": ["Engaged", "Non-engaged"],
    "rate": [security_rate_engaged, security_rate_non_engaged]
})

# Metric 5: Project activation
project_rate_engaged = engaged["has_project"].mean() if len(engaged) > 0 else 0
project_rate_non_engaged = non_engaged["has_project"].mean() if len(non_engaged) > 0 else 0

project_chart_df = pd.DataFrame({
    "cohort": ["Engaged", "Non-engaged"],
    "rate": [project_rate_engaged, project_rate_non_engaged]
})





# Display dashboard

# Metric 1: Walkthrough engagement
st.markdown(f"### Walkthrough Engagement Rate: **{engagement_rate:.1%}**")
st.altair_chart(
    stage_bar_chart(
        engagement_by_lifecycle,
        "lifecycle_stage:N",
        "rate:Q",
        "Engagement Rate by Lifecycle Stage",
        color="#2F6FDE"
    ),
    use_container_width=True
)

# Metric 2: Substantive walkthrough completion
st.markdown(f"### Substantive Completion Rate Among Engaged Companies: **{completion_rate:.1%}**")
st.altair_chart(
    stage_bar_chart(
        completion_by_lifecycle,
        "lifecycle_stage:N",
        "rate:Q",
        "Substantive Walkthrough Completion by Lifecycle Stage",
        color="#2F6FDE"
    ),
    use_container_width=True
)

# Metric 3: Meeting booked
st.markdown("### Meeting Booked")
meeting_table = comparison_table(
    "Meetings booked",
    meeting_rate_engaged,
    meeting_rate_non_engaged
)
st.dataframe(meeting_table, use_container_width=True)
st.altair_chart(
    cohort_bar_chart(meeting_chart_df, "Meeting Booking Rate by Cohort"),
    use_container_width=True
)

# Metric 4: Deal stage maturity
st.markdown("### Deal Stage Maturity")
security_table = comparison_table(
    "Reached Security Review stage or later",
    security_rate_engaged,
    security_rate_non_engaged
)
st.dataframe(security_table, use_container_width=True)
st.altair_chart(
    cohort_bar_chart(security_chart_df, "Security Review or Later by Cohort"),
    use_container_width=True
)

# Metric 5: Project activation
st.markdown("### Project Activation")
project_table = comparison_table(
    "Has at least one implementation project",
    project_rate_engaged,
    project_rate_non_engaged
)
st.dataframe(project_table, use_container_width=True)
st.altair_chart(
    cohort_bar_chart(project_chart_df, "Project Activation by Cohort"),
    use_container_width=True
)