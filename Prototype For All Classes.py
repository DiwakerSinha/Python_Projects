import pandas as pd
import numpy as np

# Importing the Raw Data
df_imported = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx", sheet_name="Interviews Raw Data")

# Splitting the Data into different Process Groupings
df_imported = df_imported[df_imported["Business Unit (Mapped)"].notna()]  # Removing NaN values
df_BE = df_imported[df_imported["Business Unit (Mapped)"].str.contains("Business Enablement", na=False)]
# Removing PSD Team since this is their job
df_BE = df_BE[~df_BE["Interviewer Vertical"].str.contains("PEOPLE STRATEGY & DEVELOPMENT", na=False)]

df_M_G = df_imported[df_imported["Business Unit (Mapped)"].str.contains("Marketing & Growth", na=False)]
df_Delivery = df_imported[df_imported["Business Unit (Mapped)"].str.contains("Delivery", na=False)]
df_Product = df_imported[df_imported["Business Unit (Mapped)"].str.contains("Product", na=False)]
# All the Process Groupings are splitting as expected
pd.set_option('mode.chained_assignment', None)

# Filtering By Month
df_BE["Interview Scheduled Date"] = np.datetime_as_string(df_BE["Interview Scheduled Date"], unit='D')
df_BE = df_BE.loc[df_BE["Interview Scheduled Date"].astype(str).str.contains("2021-06-")]


# Creating a df of the PSD people
df_BE_PSD = df_BE.loc[df_BE["Interviewer Vertical"] == "PEOPLE STRATEGY & DEVELOPMENT"]


"""
Tackling the problem of Conducted, Cancelled, and Rescheduled.
"""
# Sorting Rescheduled and Cancelled for BE
df_BE_conducted = df_BE[df_BE["Interview Status (Updated)"] == "Conducted"]
df_BE_conducted = df_BE_conducted.loc[:, 'Employee Name':'Adjusted Rating Date']
df_BE_cancelled = df_BE[df_BE["Interview Status (Updated)"] == "Cancelled"]
# Note: df_BE_cancelled includes both cancelled and rescheduled interviews under "Cancelled"

df_BE_rescheduled = pd.merge(df_BE_cancelled, df_BE_conducted, on=["Employee Name", "Interview Candidate ID",
                                                                   "Job Code"], how="inner")
df_BE_rescheduled["Interview Status (Updated)_x"] = "Rescheduled"
# Rescheduled Working Fine

df_BE_cancelled = pd.merge(df_BE_cancelled, df_BE_rescheduled,
                           on=["Employee Name", "Interview Candidate ID",
                               "Job Code"], how='outer', indicator=True) \
    .query("_merge != 'both'") \
    .drop('_merge', axis=1) \
    .reset_index(drop=True)
# Cancelled Working Fine


df_BE_rescheduled = df_BE_rescheduled.loc[:, 'Employee Name':"Adjusted Rating Date_x"]
df_BE_rescheduled.columns = [col.replace('_x', '') for col in df_BE_rescheduled.columns]

df_BE_cancelled = df_BE_cancelled.loc[:, 'Employee Name':'Adjusted Rating Date_x']
df_BE_cancelled.columns = [col.replace('_x', '') for col in df_BE_cancelled.columns]
df_BE_cancelled = df_BE_cancelled.iloc[:, 0:30]

df_BE = df_BE_conducted.merge(df_BE_rescheduled, how="outer")
df_BE = df_BE.merge(df_BE_cancelled, how="outer")

# Allotting Points for Interview Status -BE
df_BE["Points"] = 0
df_BE["Temp"] = 0
df_BE.loc[df_BE["Interview Status (Updated)"] == "Conducted", "Temp"] = 10
df_BE["Points"] = df_BE["Points"] + df_BE["Temp"]
df_BE["Temp"] = 0
df_BE.loc[df_BE["Interview Status (Updated)"] == "Cancelled", "Temp"] = -10
df_BE["Points"] = df_BE["Points"] + df_BE["Temp"]
df_BE["Temp"] = 0
df_BE.loc[df_BE["Interview Status (Updated)"] == "Rescheduled", "Temp"] = -5
df_BE["Points"] = df_BE["Points"] + df_BE["Temp"]
df_BE["Temp"] = 0

# Taking Feedback into Account -BE
df_BE.loc[df_BE["Feedback Interval"] == "Feedback within 1 working day", "Temp"] = 10
df_BE["Points"] = df_BE["Points"] + df_BE["Temp"]
df_BE["Temp"] = 0
df_BE.loc[df_BE["Feedback Interval"] == "More than 1 working day for Feedback", "Temp"] = -10
df_BE["Points"] = df_BE["Points"] + df_BE["Temp"]
df_BE["Temp"] = 0



# ************************** All 5 aspects working fine **************************





# Merging names with Points to Simplify 1st Five Attributes
df_BE_summary_one = df_BE.loc[:, ["Employee Name", "Points"]]
df_BE_summary_one = df_BE_summary_one.groupby(["Employee Name"]).sum()
df_BE_summary_one.to_excel("Merged.xlsx")
# Working









# ************************  Offered and Rejected In Later Interviews Part  ************************


# Importing the Data and Sorting by BE
df_offered_and_rejected = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx",
                                        sheet_name="All Candidates")
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Business Unit (Mapped)"]
                                                      == "Business Enablement"]
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Status"] != "Process Ongoing"]


# Taking Month Into Account
"""
Change when handling new data --> YYYY-MM-DD
Note: Basing the interview date on the Stage 1 Interview Date. Including the past month too in case the later stage of 
interviews run into the current month
"""
df_offered_and_rejected_part1 = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Interview Date"]
                                                            .astype(str).str.contains("2021-05-")]
df_offered_and_rejected_part2 = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Interview Date"]
                                                            .astype(str).str.contains("2021-06-")]
df_offered_and_rejected = df_offered_and_rejected_part1.merge(df_offered_and_rejected_part2, how="outer")
# Makes sure that the first stage interview happens either last month or this month









# **************** Refining Offered/Rejected df ****************



# Removes the interviews that never took place
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Status"]
                                                      != "Interview did not happen"]

# Removes the interviews which didn't make it to second round
df_offered_and_rejected.dropna(subset=["Stage 2 Interview Date"], inplace=True)
# Working Fine So far

# Splitting the imported df into 2 dfs based on whether the candidate was offered a job or not
df_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Offer Accepted"] == "Not offered"]
df_offered = df_offered_and_rejected.loc[(df_offered_and_rejected["Offer Accepted"] == "Yes") |
                                         (df_offered_and_rejected["Offer Accepted"] == "No") |
                                         (df_offered_and_rejected["Offer Accepted"] == "No response")]








# **************** Offered In Later Rounds ****************




# Stores the interviewers that passed the first stage and were ultimately accepted
df_offered_stage1 = df_offered.dropna(subset=["Stage 1 Interviewer(s)"])
df_offered_stage1 = df_offered_stage1.loc[df_offered_stage1["Stage 1 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage1_names = df_offered_stage1["Stage 1 Interviewer(s)"]
"""
ASK TANMAY: Figure out how to split the multiple interviewers
"""


# Removes the interviews which didn't make it to third stage
df_offered.dropna(subset=["Stage 3 Interview Date"])
# Stores the interviewers that passed the second stage and were ultimately accepted
df_offered_stage2 = df_offered.dropna(subset=["Stage 2 Interviewer(s)"])
df_offered_stage2 = df_offered_stage2.loc[df_offered_stage2["Stage 2 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage2_names = df_offered_stage2["Stage 2 Interviewer(s)"]

# Removes the interviews which didn't make it to fourth stage
df_offered.dropna(subset=["Stage 4 Interview Date"])
# Stores the interviewers that passed the third stage and were ultimately accepted
df_offered_stage3 = df_offered.dropna(subset=["Stage 3 Interviewer(s)"])
df_offered_stage3 = df_offered_stage3.loc[df_offered_stage3["Stage 3 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage3_names = df_offered_stage3["Stage 3 Interviewer(s)"]

# Removes the interviews which didn't make it to fifth stage
df_offered.dropna(subset=["Stage 5 Interview Date"])
# Stores the interviewers that passed the fourth stage and were ultimately accepted
df_offered_stage4 = df_offered.dropna(subset=["Stage 4 Interviewer(s)"])
df_offered_stage4 = df_offered_stage4.loc[df_offered_stage4["Stage 4 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage4_names = df_offered_stage4["Stage 4 Interviewer(s)"]

# Removes the interviews which didn't make it to sixth stage
df_offered.dropna(subset=["Stage 6 Interview Date"])
# Stores the interviewers that passed the fifth stage and were ultimately accepted
df_offered_stage5 = df_offered.dropna(subset=["Stage 5 Interviewer(s)"])
df_offered_stage5 = df_offered_stage5.loc[df_offered_stage5["Stage 5 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage5_names = df_offered_stage5["Stage 5 Interviewer(s)"]

# Removes the interviews which didn't make it to seventh stage
df_offered.dropna(subset=["Stage 7 Interview Date"])
# Stores the interviewers that passed the sixth stage and were ultimately accepted
df_offered_stage6 = df_offered.dropna(subset=["Stage 6 Interviewer(s)"])
df_offered_stage6 = df_offered_stage6.loc[df_offered_stage6["Stage 6 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage6_names = df_offered_stage6["Stage 6 Interviewer(s)"]

# Removes the interviews which didn't make it to eighth stage
df_offered.dropna(subset=["Stage 8 Interview Date"])
# Stores the interviewers that passed the sixth stage and were ultimately accepted
df_offered_stage7 = df_offered.dropna(subset=["Stage 7 Interviewer(s)"])
df_offered_stage7 = df_offered_stage7.loc[df_offered_stage7["Stage 7 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage7_names = df_offered_stage7["Stage 7 Interviewer(s)"]

# Merging all employees who were accepted in later rounds
df_offered_final_series = pd.concat([df_offered_stage1_names, df_offered_stage2_names,
                                     df_offered_stage3_names, df_offered_stage4_names, df_offered_stage5_names,
                                     df_offered_stage6_names, df_offered_stage7_names], ignore_index=True)


df_offered_final = df_offered_final_series.to_frame(name="Names")
df_offered_final = df_offered_final["Names"].str.split(" ")
pd.DataFrame(df_offered_final.str.split().values.tolist())
df_offered_final = pd.DataFrame(df_offered_final.str.get(0) + " " + df_offered_final.str.get(1))
df_offered_final = df_offered_final["Names"].value_counts()
df_offered_final = df_offered_final.to_frame(name="Names")
df_offered_final["Names"] = df_offered_final["Names"].astype(int) * 10
df_offered_final.rename(columns={"Names": "Points"}, inplace=True)

# To make the index (names) into a column and renaming it appropriately
df_offered_final.reset_index(level=0, inplace=True)
df_offered_final.columns = ["Interviewer", "Points"]

# Combining the BE interviewer names so that there is only a single name
df_BE = df_BE.groupby("Interviewer").count()
df_BE = (df_BE.iloc[:, 0:1]).reset_index()
df_offered_final["Interviewer"] = df_offered_final["Interviewer"].str.replace("(", " (", regex=False)

# To remove PSD people
df_offered_final = pd.merge(df_offered_final, df_BE, on="Interviewer", how="inner")
df_offered_final = df_offered_final.iloc[:, 0:2]
df_offered_final.columns = ["Employee Name", "Points"]

df_offered_final.to_excel("Offered in Later Stages.xlsx")
# Offered In Later Stages Working




# **************** Rejected In Later Rounds ****************



# Creating a new df for the Business Unit because the non PSD members have been filtered out
df_BE_new = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx", sheet_name="Interviews Raw Data")
df_BE_new = df_BE_new.loc[df_BE_new["Business Unit (Mapped)"] == "Business Enablement"]


# Making a df of the people that were rejected after the first round
df_BE_rejected = df_BE_new.loc[df_BE_new["Interview Result/Status"] == "Rejected"]
df_BE_rejected = df_BE_rejected.loc[df_BE_rejected["Interview Stage Name"] != "First Round"]
df_BE_rejected.to_excel("Rejected Check One.xlsx")


"""
Logic: Make a df of just the candidate name and job code. Then merge with the interviews raw data df. Merge on the 
Candidate Name and Job Code. Figure out which merge. After, only the interviews with the candidate name from BE_rej 
should remain. From there, remove the last interviews by loc "Interview Result/Status" != "Rejected".
Group them and count. Remove the PSD members using the same method as the one used in offered.

"""

# Obtaining all the interviews for the people that were ultimately rejected
df_BE_rejected_later = pd.merge(df_BE_new, df_BE_rejected, on=["Interview Candidate Name", "Interview Candidate ID",
                                                               "Job Code"], how="inner")

# Fixing the column names and making the df normal sized (columns)
df_BE_rejected_later = df_BE_rejected_later.iloc[:, 0:30]
df_BE_rejected_later.columns = [col.replace('_x', '') for col in df_BE_rejected_later.columns]

# Removing the rejected interviews since they are at the last stage and only sorting for interviews that were conducted
df_BE_rejected_later = df_BE_rejected_later.loc[(df_BE_rejected_later["Interview Result/Status"] != "Rejected") &
                                                (df_BE_rejected_later["Interview Status (Updated)"] == "Conducted")]

# Removing Interviewers that are in PSD
df_BE_rejected_later = df_BE_rejected_later.loc[(df_BE_rejected_later["Interviewer Vertical"]
                                                 != "PEOPLE STRATEGY & DEVELOPMENT")]

# Only keeping the interviewer name
df_BE_rejected_later = df_BE_rejected_later["Interviewer"].to_frame().reset_index()

# Adding points to each row
df_BE_rejected_later["Points"] = -10
df_BE_rejected_later = df_BE_rejected_later.iloc[:, 1:3]

# Combining common names and adding the points
df_BE_rejected_later.groupby(["Interviewer"]).sum()
df_BE_rejected_later.columns = ["Employee Name", "Points"]
# Rejected in Later Stages Working Fine



# **************** Combining Offered and Rejected in Later stages ****************

df_merged_two = pd.concat([df_offered_final, df_BE_rejected_later])
df_merged_two.groupby(["Employee Name"]).sum()
df_merged_two.reset_index(inplace=True)
df_merged_two.drop("index", inplace=True, axis=1)
df_merged_two.columns = ["Employee Name", "Points"]
# Concat working fine



# **************** Combining all parameters together ****************
df_BE_summary_one.reset_index(inplace=True)
df_BE_final = pd.concat([df_BE_summary_one, df_merged_two])
df_BE_final.groupby(["Employee Name"]).sum()
df_BE_final.reset_index(inplace=True)
df_BE_final.drop("index", inplace=True, axis=1)
df_BE_final = df_BE_final.sort_values('Employee Name')
df_BE_final.reset_index(inplace=True)
df_BE_final.drop("index", inplace=True, axis=1)
df_BE_final.to_excel("BE.xlsx")
# Works Fine
