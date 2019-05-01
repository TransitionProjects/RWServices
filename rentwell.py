"""
A report for showing all Rent Well Services during a time period
"""

__version__ = "1.0"
__author__ = "David Marienburg"
__maintainer__ = "David Marienburg"

import pandas as pd
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

from datetime import datetime

# since the reports are always run after the first of the month this report
# and since the base report is always run for the year to date this report
# needs to slice based on last month
report_month = datetime.today().month - 1

# open the service consistency report
file = askopenfilename(
    title="Service Consistency Report",
    initialdir="//tproserver/Reports/Monthly Reports"
)
df = pd.read_excel(file)

# defining data frame for provider specific code
spsc = df[
    df["Provider Specific Code"].notna() &
    (
        df["Provider Specific Code"].str.contains("Rent Well") |
        df["Provider Specific Code"].str.contains("RentWell")
    ) &
    (df["Service Date"].dt.month == report_month)
].copy()

spsc["Service Date"]=spsc["Service Date"].dt.date

# slicing show services with incorrectly created Service Provider Specific Codes
spsc_is_blank = df[
    df["Provider Specific Code"].isna() &
    df["Service Type"].str.contains("Tenant Readiness")

]

# tracking those who attended a RentWell Class
attendee = spsc[
    spsc["Provider Specific Code"].str.contains("Attendance")
][["CTID","Staff Providing The Service", "Service Date"]]

# tracking those who graduated a RentWell Class
graduate = spsc[
    spsc["Provider Specific Code"].str.contains("Graduation")
][["CTID","Staff Providing The Service", "Service Date"]]

# tracking those who were referred to a RentWell Class
referred = spsc[
    spsc["Provider Specific Code"].str.contains("Referral")
][["CTID","Staff Providing The Service", "Service Date"]]

# create a summary
summary = spsc.pivot_table(
    index=["Provider Specific Code"],
    values=["CTID"],
    aggfunc=len
)

writer = pd.ExcelWriter(
    asksaveasfilename(
        title="RentWell Monthly Data",
        initialdir="//tproserver/Reports/Monthly Report",
        initialfile="RentWell Monthly Data.xlsx",
        defaultextension=".xlsx"
    ),
    engine="xlsxwriter"
)
summary.to_excel(writer, sheet_name="Summary")
graduate.to_excel(writer, sheet_name="RentWell Graduates", index=False)
attendee.to_excel(writer, sheet_name="RentWell Attendees", index=False)
referred.to_excel(writer, sheet_name="RentWell Referrals", index=False)
spsc_is_blank.to_excel(writer, sheet_name="Errors", index=False)
writer.save()
