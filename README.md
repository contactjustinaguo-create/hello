# [BFSI PQC Pipeline Learning Dashboard](https://walkthroughdashboard.streamlit.app/#walkthrough-engagement-rate-80-0)
A Streamlit analytics dashboard paired with a separately built interactive Post-Quantum Cryptography (PQC) walkthrough to answer whether the walkthrough drives pipeline progression for Banking, Financial Services & Insurance (BFSI) prospects.
Built as a take-home assignment for QuSecure in a 3–4 hour timebox.

## What This Does
The project has two parts:
1. Interactive Walkthrough Prototype (Built elsewhere): A branching educational experience designed for BFSI CISOs. It explains why crypto-agility matters for PQC migration and drives toward a single CTA: downloading an internal justification pack. The walkthrough includes two purposeful interaction moments, one of which branches the path based on the CISO's infrastructure context (ATM vs. Core Banking).
2. Pipeline Learning Dashboard: A single-page Streamlit dashboard that joins HubSpot CRM data with project tracking data to measure whether walkthrough engagement correlates with pipeline progression across 5 metrics.

## Stack

- Python: Data processing and app logic
- Pandas: Data manipulation and dataset joining
- Altair: Interactive charts
- Streamlit: dashboard rendering and sharing


## The Messy Edge Case
The project CSV had missing company_id values with inconsistent company name formatting that didn't match the HubSpot CSV. Rather than using a fuzzy matching library, I wrote a character-by-character matching algorithm that eliminates candidates incrementally until exactly one match remains. This resolved all missing IDs without false positives.

## The 5 Metrics
1. Walkthrough Engagement Rate: % of companies that started the walkthrough out of all companies in the dataset
2. Substantive Completion Rate: % of walkthrough-engaged companies that completed more than half of the walkthrough steps
3. Meeting Booked: % of companies that booked a sales meeting, compared between engaged and non-engaged cohorts
4. Deal Stage Maturity: % of companies whose furthest deal stage reached Security Review or later, compared between cohorts
5. Project Activation: % of companies with at least one implementation project, compared between cohorts

"Progression" is defined as reaching Security Review or later since it is the stage most directly relevant to PQC adoption point where a buyer moves beyond early curiosity into active technical evaluation.

## Key Findings

- Engagement has a mixed effect on pipeline. Walkthrough-engaged companies book sales meetings at a higher rate, but show lower deal stage advancement and project activation than non-engaged companies.
- The strongest signal the walkthrough is working would be a higher rate of engaged companies reaching Security Review or later. Since the walkthrough is designed to accelerate technical conviction, its impact should appear at the point of serious evaluation.
- Causality is unconfirmed. The mixed results may reflect that engaged companies are earlier in the sales cycle. Timestamp data would allow comparison of walkthrough engagement dates against deal activity dates to distinguish cause from correlation.


## What I'd Change Next Iteration

1. Capture interaction timestamps: To determine whether walkthrough engagement causes progression or merely correlates with it
2. Add deal size as a variable; to see if contract value affects how the walkthrough influences decisions
3. Track project count and status: To understand whether more active implementation work correlates with deal progress
4. Control for owner: To test whether the sales relationship is a confounding factor in walkthrough engagement


## Assumptions

- Lifecycle stages progress linearly (Lead → MQL → SQL → Opportunity) with no skipping to simplify cohort analysis
- Lifecycle stage (not deal stage) used for the first two metrics because it has more granular values, making the analysis more statistically sound
- Security Review chosen as the deal maturity threshold because it represents meaningful technical commitment beyond early-stage curiosity

## How I Used AI

Coded with AI assistance; read through all generated code to ensure understanding, replaced hard-coded values with variables, and added documentation; AI suggested Streamlit as a sharing-friendly framework, which I validated on developer forums before adopting


## Running Locally
```
bashpip install streamlit pandas altair openpyxl
streamlit run dashboard.py
```

Note: The dataset included is synthetic. The original assignment used real company data provided by QuSecure.
