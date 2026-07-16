import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO
from datetime import date

st.title("AIP Master Preparation")

# Upload only 1 file
t1 = st.file_uploader("Upload AIP Tracker_Backend Report (CSV or Excel)", type=["csv", "xlsx"])

if t1:
    # Helper function to load CSV/Excel
    def load_file(f):
        if f.name.endswith("xlsx"):
            return pd.read_excel(f, sheet_name="AIP Master")
        else:
            return pd.read_csv(f)

    # Load the file
    df = load_file(t1)

    # ---------------------------------------------------------------
    # Date filter selected on the frontend
    # ---------------------------------------------------------------
    cutoff_date = st.date_input(
        "Show records where 'AIP Follow-up Latest Date' is after:",
        value=date(2026, 7, 1)  # default = 1-Jul-26
    )

    # Parse the follow-up date column into a real datetime so the
    # comparison is a true date comparison (not a string comparison).
    # Handles values like "1-Jul-26", "01/07/2026", Excel datetimes, etc.
    df["_followup_dt"] = pd.to_datetime(
        df["AIP Follow-up Latest Date"], errors="coerce", dayfirst=True
    )

    # Store an ISO-formatted helper column (YYYY-MM-DD) so SQLite can
    # compare dates correctly.
    df["_followup_iso"] = df["_followup_dt"].dt.strftime("%Y-%m-%d")

    # Create SQLite in-memory DB
    conn = sqlite3.connect(":memory:")
    df.to_sql("aip_master", conn, index=False, if_exists="replace")

    # SQL Query (date cutoff passed as a parameter from the date picker)
    query = """

WITH base1 AS (SELECT "ADEK Applicant ID",
       "Student Name",
       "Current Mentor",
       "Regional Manager",
       "ADEK Advisor",
       "Grades Report-  Pathway Distribution ",
       "College at Time of Transcript Issuance",
       Country,
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Transcript Term" END AS "Transcript Term",
       "GPA Type",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Latest CGPA" END AS "Latest CGPA",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Latest GPA" END AS "Latest GPA",
       "AIP Submitted as on",
       "AIP Reason",
       CASE WHEN "Is the student's attainment improving or declining from last transcript?" = 0 THEN 'AIP cannot be submitted' ELSE "Is the student's attainment improving or declining from last transcript?" END AS "Is the student's attainment improving or declining from last transcript?",
       CASE WHEN "Courses failed during pathway status" = 0 THEN 'AIP cannot be submitted' ELSE "Courses failed during pathway status" END AS "Courses failed during pathway status",
       CASE WHEN "Summary of academic challenges" = 0 THEN 'AIP cannot be submitted' ELSE "Summary of academic challenges" END AS "Summary of academic challenges",
       CASE WHEN "Summary of student's attitude and attendance" = 0 THEN 'AIP cannot be submitted' ELSE "Summary of student's attitude and attendance" END AS "Summary of student's attitude and attendance",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Current cumulative credit hours / units completed" END AS "Current cumulative credit hours / units completed",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Credit hours / units remaining until pathway competition" END AS "Credit hours / units remaining until pathway competition",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Average CGPA / WAM required in remaining credit hours / units to achieve scholarship attainment benchmarks" END AS "Average CGPA / WAM required in remaining credit hours / units to achieve scholarship attainment benchmarks",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Semesters / terms available to achieve scholarship threshold of 2.5 CGPA / 55% WAM" END AS "Semesters / terms available to achieve scholarship threshold of 2.5 CGPA / 55% WAM",
       CASE WHEN "Intervention Planning - Semester / Term" = 0 THEN 'AIP cannot be submitted' ELSE "Intervention Planning - Semester / Term" END AS "Intervention Planning - Semester / Term",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Desired target CGPA / WAM for end of semester" END AS "Desired target CGPA / WAM for end of semester",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "No. of credits / units to be taken" END AS "No. of credits / units to be taken",
       CASE WHEN "Transcript Term" = 0 THEN 'AIP cannot be submitted' ELSE "Target GPA / WAM required in these credit courses / units to achieve desired CGPA" END AS "Target GPA / WAM required in these credit courses / units to achieve desired CGPA",
       CASE WHEN "First strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "First strategic intervention" END AS "First strategic intervention",
       CASE WHEN "Reason for first strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Reason for first strategic intervention" END AS "Reason for first strategic intervention",
       CASE WHEN "Monitoring of first strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Monitoring of first strategic intervention" END AS "Monitoring of first strategic intervention",
       CASE WHEN "Follow up (FSI) : 15th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (FSI) : 15th of the month" END AS "Follow up (FSI) : 15th of the month",
       CASE WHEN "Follow up (FSI) : 30th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (FSI) : 30th of the month" END AS "Follow up (FSI) : 30th of the month",
       CASE WHEN "Second strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Second strategic intervention" END AS "Second strategic intervention",
       CASE WHEN "Reason for second strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Reason for second strategic intervention" END AS "Reason for second strategic intervention",
       CASE WHEN "Monitoring of second strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Monitoring of second strategic intervention" END AS "Monitoring of second strategic intervention",
       CASE WHEN "Follow up (SSI) : 15th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (SSI) : 15th of the month" END AS "Follow up (SSI) : 15th of the month",
       CASE WHEN "Follow up (SSI) : 30th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (SSI) : 30th of the month" END AS "Follow up (SSI) : 30th of the month",
       CASE WHEN "Third strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Third strategic intervention" END AS "Third strategic intervention",
       CASE WHEN "Reason for third strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Reason for third strategic intervention" END AS "Reason for third strategic intervention",
       CASE WHEN "Monitoring of third strategic intervention" = 0 THEN 'AIP cannot be submitted' ELSE "Monitoring of third strategic intervention" END AS "Monitoring of third strategic intervention",
       CASE WHEN "Follow up (TSI) : 15th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (TSI) : 15th of the month" END AS "Follow up (TSI) : 15th of the month",
       CASE WHEN "Follow up (TSI) : 30th of the month" = 0 THEN 'Further AIP Follow-up not required - Refer to column AR for remark' ELSE "Follow up (TSI) : 30th of the month" END AS "Follow up (TSI) : 30th of the month",
       CASE WHEN "Further AIP Required" IN ("No","#N/A") OR "Further AIP Required" IS NULL THEN 'No' ELSE "Further AIP Required" END AS "Further AIP Required",
       CASE WHEN "Reason for not submitting an AIP" = 0 THEN '-' ELSE "Reason for not submitting an AIP" END AS "Reason for not submitting an AIP",
       CASE WHEN "Further AIP Required" = "Yes" THEN '-'
            WHEN ("Reason for not submitting follow up on AIP" = "#N/A" OR "Reason for not submitting follow up on AIP" IS NULL) AND "Third strategic intervention" = 0 THEN 'AIP cannot be submitted'
            ELSE "Reason for not submitting follow up on AIP" END AS "Reason for not submitting follow up on AIP",
       CASE WHEN "Reason for not submitting follow up on AIP" IN ("Transcript received: WAM above 55%", "Transcript received: CGPA above 2.5") THEN "Latest Transcript Grades (for students who don't require follow-up AIP)"
            WHEN "Reason for not submitting follow up on AIP" IN ("Student transferring to new institution", "Student terminated", "Student on National Service", "Medical Emergency") THEN 'Not Required'
            ELSE '-' END AS "Latest Transcript Grades (for students who don't require follow-up AIP)",
       "AIP Follow-up Latest Date",
       "_followup_iso"
FROM aip_master)

SELECT * FROM base1
WHERE "AIP Reason" <> "AIP cannot be submitted"
  AND "_followup_iso" IS NOT NULL
  AND "_followup_iso" > ?;

    """

    # Run query with the selected date as a parameter (ISO format)
    result_df = pd.read_sql_query(query, conn, params=[cutoff_date.strftime("%Y-%m-%d")])

    # Drop the internal helper column before displaying/exporting
    result_df = result_df.drop(columns=["_followup_iso"])

    st.subheader(f"Filtered Result — records after {cutoff_date.strftime('%d-%b-%y')}")
    st.write(f"Total records: {len(result_df)}")
    st.dataframe(result_df)

    # Create Excel file in memory
    output = BytesIO()

    # Read all sheets from original workbook
    all_sheets = pd.read_excel(t1, sheet_name=None)

    # Replace only AIP Master sheet with SQL result
    all_sheets["AIP Master"] = result_df

    # Write all sheets back
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, sheet_df in all_sheets.items():
            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)

    excel_data = output.getvalue()

    # Download button for Excel
    st.download_button(
        label="📥 Download Result as Excel",
        data=excel_data,
        file_name="AIP Tracker_Backend Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
