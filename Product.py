import pandas as pd
import numpy as np

"""
I will use this to mark where the final df will be created. This starts the compilation of data in an Excel worksheet
as that of Tanmay.
"""

# Importing the Raw Data
df_imported = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx", sheet_name="Interviews Raw Data")

# Obtaining employees from Product and Removing PSD people
df_Product = df_imported.loc[df_imported["Business Unit (Mapped)"] == "Product"]
df_Product = df_Product.loc[df_Product["Interviewer Vertical"] != "PEOPLE STRATEGY & DEVELOPMENT"]

# To prevent runtime suggestions
pd.set_option('mode.chained_assignment', None)

# Filtering Interviews Conducted By Month
df_Product["Interview Scheduled Date"] = np.datetime_as_string(df_Product["Interview Scheduled Date"], unit='D')
df_Product = df_Product.loc[df_Product["Interview Scheduled Date"].astype(str).str.contains("2021-06-")]

# Creating df of interviews that were Conducted
df_Product_conducted = df_Product.loc[df_Product["Interview Status (Updated)"] == "Conducted"]

# Creating a df of interviews that were cancelled/rescheduled
df_Product_cancelled_and_rescheduled = df_Product[df_Product["Interview Status (Updated)"] == "Cancelled"]

# Creating a df of rescheduled interviews and handling the columns
df_Product_rescheduled = pd.merge(df_Product_cancelled_and_rescheduled, df_Product_conducted, on=["Employee Name",
                                                                                                     "Interview Candidate ID",
                                                                                                     "Job Code"], how="inner")
df_Product_rescheduled["Interview Status (Updated)_x"] = "Rescheduled"
df_Product_rescheduled = df_Product_rescheduled.loc[:, 'Employee Name':"Adjusted Rating Date_x"]
df_Product_rescheduled.columns = [col.replace('_x', '') for col in df_Product_rescheduled.columns]


# Creating a df of cancelled interviews and handling the columns
df_Product_cancelled = pd.merge(df_Product_cancelled_and_rescheduled, df_Product_rescheduled,
                                 on=["Employee Name", "Interview Candidate ID",
                                     "Job Code"], how='outer', indicator=True) \
    .query("_merge != 'both'") \
    .drop('_merge', axis=1) \
    .reset_index(drop=True)
df_Product_cancelled = df_Product_cancelled.loc[:, 'Employee Name':'Adjusted Rating Date_x']
df_Product_cancelled.columns = [col.replace('_x', '') for col in df_Product_cancelled.columns]
df_Product_cancelled = df_Product_cancelled.iloc[:, 0:30]


"""
Creating the first instance of the final df and instantiating it with the conducted interviews
"""
df_final = df_Product_conducted[["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]].reset_index()
df_final = df_final.groupby(["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]).count().reset_index()
df_final.rename(columns={"index": "Conducted"}, inplace=True)


# Adding the Rescheduled Column to the final df
df_rescheduled_temp = df_Product_rescheduled[["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]].reset_index()
df_rescheduled_temp = df_rescheduled_temp.groupby(["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]).\
    count().reset_index()
df_rescheduled_temp.rename(columns={"index": "Rescheduled"}, inplace=True)
df_final = df_final.merge(df_rescheduled_temp, on=["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"],
                          how="outer")


# Adding the Cancelled Column to the final df
df_cancelled_temp = df_Product_cancelled[["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]].reset_index()
df_cancelled_temp = df_cancelled_temp.groupby(["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]).\
    count().reset_index()
df_cancelled_temp.rename(columns={"index": "Cancelled"}, inplace=True)
df_final = df_final.merge(df_cancelled_temp, on=["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"],
                          how="outer")

# ************************** Conducted, Rescheduled, and Cancelled Added to final df **************************


# Creating df for Feedback Within 1 Day
df_feedback_one = df_Product.loc[df_Product["Feedback Interval"] == "Feedback within 1 working day"]
df_feedback_one = df_feedback_one.loc[df_feedback_one["Interview Status (Updated)"] == "Conducted"]
# Making a df to count instances of ^^^ and store in temp df to be added to final df
df_feedback_one_temp = df_feedback_one[["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]].reset_index()
df_feedback_one_temp = df_feedback_one_temp.groupby(["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"]).\
    count().reset_index()
df_feedback_one_temp.rename(columns={"index": "Feedback Within 1 Working Day"}, inplace=True)
df_final = df_final.merge(df_feedback_one_temp, on=["Interviewer", "Interviewer Vertical", "Business Unit (Mapped)"],
                          how="outer")


# Creating df for Feedback >1 day
df_feedback_more_than_one = df_Product.loc[df_Product["Feedback Interval"] == "More than 1 working day for Feedback"]
df_feedback_more_than_one = df_feedback_more_than_one[df_feedback_more_than_one["Interview Status (Updated)"] == "Conducted"]
# Making a df to count instances of ^^^ and store in temp df to be added to final df
df_feedback_more_than_one_temp = df_feedback_more_than_one[["Interviewer", "Interviewer Vertical",
                                                            "Business Unit (Mapped)"]].reset_index()
df_feedback_more_than_one_temp = df_feedback_more_than_one_temp.groupby(["Interviewer",
                                                                         "Interviewer Vertical",
                                                                         "Business Unit (Mapped)"])\
    .count().reset_index()
df_feedback_more_than_one_temp.rename(columns={"index": "More Than 1 Working Day For Feedback"}, inplace=True)
df_final = df_final.merge(df_feedback_more_than_one_temp, on=["Interviewer", "Interviewer Vertical",
                                                              "Business Unit (Mapped)"], how="outer")

# ************************** Feedback Categories Added to final df **************************



# Importing All Candidates Data For Product
df_offered_and_rejected = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx",
                                        sheet_name="All Candidates")
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Business Unit (Mapped)"]
                                                      == "Product"]
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Status"] != "Process Ongoing"]


# Taking Month Into Account
"""
Change when handling new data --> YYYY-MM-DD
Note: Basing the interview date on the Stage 1 Interview Date. Including the past month too in case the later stage of 
interviews run into the current month
"""
# Makes sure that the first stage interview happens either last month or this month
df_offered_and_rejected_part1 = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Interview Date"]
                                                            .astype(str).str.contains("2021-05-")]
df_offered_and_rejected_part2 = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Interview Date"]
                                                            .astype(str).str.contains("2021-06-")]
df_offered_and_rejected = df_offered_and_rejected_part1.merge(df_offered_and_rejected_part2, how="outer")


# Removes the interviews that never took place
df_offered_and_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Stage 1 Status"]
                                                      != "Interview did not happen"]

# Removes the interviews which didn't make it to second round
df_offered_and_rejected.dropna(subset=["Stage 2 Interview Date"], inplace=True)
# Working Fine So far

# Splitting the refined all candidates df into further dfs based on whether the candidate were offered a job or not
df_offered = df_offered_and_rejected.loc[(df_offered_and_rejected["Offer Accepted"] == "Yes") |
                                         (df_offered_and_rejected["Offer Accepted"] == "No") |
                                         (df_offered_and_rejected["Offer Accepted"] == "No response")]

df_rejected = df_offered_and_rejected.loc[df_offered_and_rejected["Offer Accepted"] == "Not offered"]



# **************** Offered In Later Rounds ****************


# Removes the interviews which didn't make it to second stage
df_offered.dropna(subset=["Stage 2 Interview Date"])
# Stores the interviewers that passed the first stage and were ultimately accepted
df_offered_stage1 = df_offered.dropna(subset=["Stage 1 Interviewer(s)"])
df_offered_stage1 = df_offered_stage1.loc[df_offered_stage1["Stage 1 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage1_names = df_offered_stage1["Stage 1 Interviewer(s)"] + "?" + \
                          df_offered_stage1["Candidate Name"] + "!" + df_offered_stage1["Job Code"]


# Removes the interviews which didn't make it to third stage
df_offered.dropna(subset=["Stage 3 Interview Date"])
# Stores the interviewers that passed the second stage and were ultimately accepted
df_offered_stage2 = df_offered.dropna(subset=["Stage 2 Interviewer(s)"])
df_offered_stage2 = df_offered_stage2.loc[df_offered_stage2["Stage 2 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage2_names = df_offered_stage2["Stage 2 Interviewer(s)"] + "?" + \
                          df_offered_stage2["Candidate Name"] + "!" + df_offered_stage2["Job Code"]


# Removes the interviews which didn't make it to fourth stage
df_offered.dropna(subset=["Stage 4 Interview Date"])
# Stores the interviewers that passed the third stage and were ultimately accepted
df_offered_stage3 = df_offered.dropna(subset=["Stage 3 Interviewer(s)"])
df_offered_stage3 = df_offered_stage3.loc[df_offered_stage3["Stage 3 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage3_names = df_offered_stage3["Stage 3 Interviewer(s)"] + "?" + \
                          df_offered_stage3["Candidate Name"] + "!" + df_offered_stage3["Job Code"]

# Removes the interviews which didn't make it to fifth stage
df_offered.dropna(subset=["Stage 5 Interview Date"])
# Stores the interviewers that passed the fourth stage and were ultimately accepted
df_offered_stage4 = df_offered.dropna(subset=["Stage 4 Interviewer(s)"])
df_offered_stage4 = df_offered_stage4.loc[df_offered_stage4["Stage 4 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage4_names = df_offered_stage4["Stage 4 Interviewer(s)"] + "?" + \
                          df_offered_stage4["Candidate Name"] + "!" + df_offered_stage4["Job Code"]

# Removes the interviews which didn't make it to sixth stage
df_offered.dropna(subset=["Stage 6 Interview Date"])
# Stores the interviewers that passed the fifth stage and were ultimately accepted
df_offered_stage5 = df_offered.dropna(subset=["Stage 5 Interviewer(s)"])
df_offered_stage5 = df_offered_stage5.loc[df_offered_stage5["Stage 5 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage5_names = df_offered_stage5["Stage 5 Interviewer(s)"] + "?" + \
                          df_offered_stage5["Candidate Name"] + "!" + df_offered_stage5["Job Code"]

# Removes the interviews which didn't make it to seventh stage
df_offered.dropna(subset=["Stage 7 Interview Date"])
# Stores the interviewers that passed the sixth stage and were ultimately accepted
df_offered_stage6 = df_offered.dropna(subset=["Stage 6 Interviewer(s)"])
df_offered_stage6 = df_offered_stage6.loc[df_offered_stage6["Stage 6 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage6_names = df_offered_stage6["Stage 6 Interviewer(s)"] + "?" + \
                          df_offered_stage6["Candidate Name"] + "!" + df_offered_stage6["Job Code"]

# Removes the interviews which didn't make it to eighth stage
df_offered.dropna(subset=["Stage 8 Interview Date"])
# Stores the interviewers that passed the seventh stage and were ultimately accepted
df_offered_stage7 = df_offered.dropna(subset=["Stage 7 Interviewer(s)"])
df_offered_stage7 = df_offered_stage7.loc[df_offered_stage7["Stage 7 Interview Date"].astype(str).str
                                                                                                 .contains("2021-06-")]
df_offered_stage7_names = df_offered_stage7["Stage 7 Interviewer(s)"] + "?" + \
                          df_offered_stage7["Candidate Name"] + "!" + df_offered_stage7["Job Code"]

# Vertically Stacking all employees who chose candidates that were accepted in later rounds in the form of a Series
df_offered_final_series = pd.concat([df_offered_stage1_names, df_offered_stage2_names,
                                     df_offered_stage3_names, df_offered_stage4_names, df_offered_stage5_names,
                                     df_offered_stage6_names, df_offered_stage7_names], ignore_index=True)


# Converting Series into a dataframe
df_offered_final = df_offered_final_series.to_frame(name="Interviewer?Candidate!Job Code")

# Splitting the elements from column 'Interviewer?Candidate!Job Code' using ? as that separates the interviewers from
# the people the candidate information. This is then stored as list of interviewers, Candidate Info
df_offered_final = df_offered_final["Interviewer?Candidate!Job Code"].str.split("?")
pd.DataFrame(df_offered_final.str.split().values.tolist())

# First part contains all the interviewers including the secondary interviewers
# Main goal is to only keep the first interviewer as they are the primary interviewer
df_offered_first_part = pd.DataFrame(df_offered_final.str.get(0))
df_offered_first_part.columns = ["Interviewer"]
df_offered_first_part = df_offered_first_part["Interviewer"].str.split(" ")
pd.DataFrame(df_offered_first_part.str.split().values.tolist())
df_offered_first_part = pd.DataFrame(df_offered_first_part.str.get(0) + " " + df_offered_first_part.str.get(1))

# The second part contains the Candidate Info
df_offered_second_part = pd.DataFrame(df_offered_final.str.get(1))
df_offered_second_part.columns = ["Candidate Info"]

# Creating offered temp df to add to the final df
df_offered_temp = df_offered_first_part[["Interviewer"]]
df_offered_temp["Candidate Info"] = df_offered_second_part["Candidate Info"]
df_offered_temp.reset_index()

# Removing any interviewer that conducted interviews for the same candidate
df_offered_temp["Temp"] = 0
df_offered_temp = df_offered_temp.groupby(["Interviewer", "Candidate Info"]).count().reset_index()
df_offered_temp.drop(columns=["Candidate Info", "Temp"], inplace=True)

# Counting how many instances of Offered in Later Rounds are present for a given interviewer
df_offered_temp["Offered In Later Rounds"] = 0
df_offered_temp = df_offered_temp.groupby(["Interviewer"]).count().reset_index()

# Creating a df of Product members such that each employee's name occurs only once
df_Product_non_PSD = df_Product[["Interviewer"]].reset_index()
df_Product_non_PSD = df_Product_non_PSD.groupby("Interviewer").count().reset_index()

# Changing the formatting of the name so that merge works properly
df_offered_temp["Interviewer"] = df_offered_temp["Interviewer"].str.replace("(", " (", regex=False)

# Merging the df so that only non-PSD members are left
df_offered_temp = pd.merge(df_offered_temp, df_Product_non_PSD, on="Interviewer", how="inner")
df_offered_temp.drop(columns=["index"], inplace=True)

df_final = df_final.merge(df_offered_temp, on="Interviewer", how="outer")



# **************** Rejected In Later Rounds ****************


# Creating a new df for the Business Unit because the non PSD members have been filtered out
df_Product_new = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\Interviewer R&R.xlsx", sheet_name="Interviews Raw Data")
df_Product_new = df_Product_new.loc[df_Product_new["Business Unit (Mapped)"] == "Product"]


# Making a df of the people that were rejected after the first round
df_Product_rejected = df_Product_new.loc[(df_Product_new["Interview Result/Status"] == "Rejected") |
                                           (df_Product_new["Interview Result/Status"] == "Pending for Decision - Rejected")]
df_Product_rejected = df_Product_rejected.loc[df_Product_rejected["Interview Stage Name"] != "First Round"]
df_Product_rejected.reset_index(inplace=True)
df_Product_rejected.drop(columns=["index"], inplace=True)
df_Product_rejected = df_Product_rejected.iloc[:, 0:30]

# Obtaining all the interviews for the people that were eventually rejected
df_Product_rejected_later = pd.merge(df_Product_new, df_Product_rejected, on=["Interview Candidate Name", "Interview Candidate ID",
                                                                                 "Job Code"], how="inner")

# Fixing the column names and making the df normal sized (columns)
df_Product_rejected_later = df_Product_rejected_later.iloc[:, 0:30]
df_Product_rejected_later.columns = [col.replace('_x', '') for col in df_Product_rejected_later.columns]

# Removing the rejected interviews since they are at the last stage and only sorting for interviews that were conducted
df_Product_rejected_later = df_Product_rejected_later.loc[df_Product_rejected_later["Interview Status (Updated)"] == "Conducted"]
df_Product_rejected_later = df_Product_rejected_later.loc[df_Product_rejected_later["Interview Result/Status"] != "Rejected"]
df_Product_rejected_later = df_Product_rejected_later.loc[df_Product_rejected_later["Interview Result/Status"] != "Pending for Decision - Rejected"]

# Removing Interviewers that are in PSD
df_Product_rejected_later = df_Product_rejected_later.loc[(df_Product_rejected_later["Interviewer Vertical"] != "PEOPLE STRATEGY & DEVELOPMENT")]
df_Product_rejected_later.reset_index(inplace=True)
df_Product_rejected_later.drop(columns=["index"], inplace=True)
df_Product_rejected_later = df_Product_rejected_later.iloc[:, 0:30]

# Only keeping the interviewer name
df_Product_rejected_later = df_Product_rejected_later["Interviewer"].to_frame().reset_index()
df_Product_rejected_later.drop(columns=["index"], inplace=True)
df_Product_rejected_later = df_Product_rejected_later.iloc[:, 0:30]

# Counting how many instances of Rejected in Later Rounds are present for a given interviewer
df_Product_rejected_later["Rejected In Later Rounds"] = 0
df_Product_rejected_later = df_Product_rejected_later.groupby(["Interviewer"]).count().reset_index()

# Adding to final df
df_final = df_final.merge(df_Product_rejected_later, on="Interviewer", how="outer")

# To fill any blank cells in Excel with 0 where points are necessary
df_final = df_final.apply(lambda x: x.fillna(0) if x.dtype.kind in 'biufc' else x.fillna('.'))

# Adding Score Column to final df
df_final["Score"] = 0
df_final["Score"] = (df_final["Conducted"] * 10) + (df_final["Rescheduled"] * -5) + (df_final["Cancelled"] * -10) + \
                    (df_final["Feedback Within 1 Working Day"] * 10) + (df_final["More Than 1 Working Day For Feedback"] * -10) \
                    + (df_final["Offered In Later Rounds"] * 15) + (df_final["Rejected In Later Rounds"] * -10)
df_final.sort_values("Score", ascending=False, inplace=True)
df_final.reset_index(inplace=True)
df_final.drop(columns=["index"], inplace=True)

df_final.to_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Product.xlsx")
