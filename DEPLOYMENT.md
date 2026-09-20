# NCVET submission validation and cleaning

This release includes downloads, validation, consolidation, cleaning, comparisons,
review queues and three workbook downloads. Scoring and grading are not included.
It starts empty and requires a new link workbook and reporting period.

## Free hosting: Streamlit Community Cloud

1. Extract the ZIP.
2. Sign up at https://share.streamlit.io/ and connect your GitHub account.
3. Create a **private GitHub repository** called `ncvet-monitoring`.
4. Upload the contents of the extracted `NCVET_Cleaning_Tool` folder into the
   repository root. Include the `.streamlit` directory and four reference Excel
   workbooks. Do not upload old Runs, submissions, secrets, or videos.
5. In Streamlit Community Cloud, create an app from that repository. Select its
   branch and use `streamlit_app.py` as the main file. Choose Python **3.12**
   in advanced settings, then deploy.
6. Use the generated `streamlit.app` address. Set the app's sharing to private
   and invite the intended viewers using the hosting dashboard.
7. Upload a fresh link workbook, choose the month/year, and select
   **Run validation & cleaning**. Download the outputs before closing the session.

Community Cloud hosting is free, has resource limits and may sleep when inactive.
This application uses temporary server storage; outputs are not a durable archive.
Reference: https://docs.streamlit.io/deploy/streamlit-community-cloud

For private Drive uploads, a current OAuth access token can be supplied as the
top-level `GOOGLE_DRIVE_ACCESS_TOKEN` secret in the hosting dashboard. Do not put
it in GitHub. Link-accessible files need no token. Tokens expire; refresh them
when necessary. The web app accepts Google Drive/Sheets HTTPS links only.

## Local Windows use

Install Python 3.12, extract the ZIP completely, then double-click `START_TOOL.cmd`.
The launcher creates a local environment, installs dependencies and opens the app.
Internet is required for setup and downloading agency submissions.

## Outputs

- `AA_Status_Report.xlsx`: submission coverage and format validation.
- `AA_Consolidated_Master.xlsx`: original consolidated records.
- `AA_Cleaned_Master.xlsx`: cleaned records, original values and the audit trail.

Unresolved values become `Missing`. FFCPL has the explicitly requested cleaning
alias `Fashion Future First Assessment Agency`; the reference master is unchanged.
The app reports cleaning activity, not verified accuracy or agency performance.
