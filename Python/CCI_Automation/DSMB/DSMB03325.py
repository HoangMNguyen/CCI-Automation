#!/usr/bin/env python3
from itertools import count
import pandas as pd
import numpy as np
from DSMB.DSMB_util import (
    add_rename_column_corelisting,
    get_stats_df,
    get_stats_percentage,
    get_stats_percentage2,
    get_stats_perc_df,
    convert_float_2_sci_notation,
)
from util import (
    get_data_from_dict,
    get_data_from_dict_first,
    age_calculation,
    add_rename_column_df,
    convert_integers_to_strings,
)
from datetime import datetime
from typing import Optional
import re

# Opt-in to the future behavior
pd.set_option("future.no_silent_downcasting", True)


class DSMB03325:
    def __init__(
        self,
        data,
        output_dir,
        output_file_name,
    ):
        self.data = data
        self.output_dir = output_dir
        self.output_file_name = output_file_name
        self.data_loaded = False

    def run(self):
        """
        Process the given data and return the processed result.

        Args:
            None

        Returns:
            None
        """
        # process the listing only when Demographics is not empty
        data = self.data
        if not data["DM"].empty:
            self.enrollment_listing_df_output, self.enrollment_listing_df = self.enrollment_listing()
            self.enrollment_stat_table(self.enrollment_listing_df)
            # self.EGFR_listing_df_output, self.EGFR_listing_df = self.EGFR_listing()
            # self.infusion_df, self.infusionR_df = self.infusion_listing()
            # self.infusion_stats(self.infusion_df, self.infusionR_df)

            # self.AE_df, self.status_df, self.safetyCH1_total_df, self.safetyCH2_total_df, self.safetyCHN1_total_df = (
            #     self.status_listing(self.enrollment_listing_df)
            # )
            self.export(self.output_dir, self.output_file_name)

    def enrollment_listing(self):
        data = self.data
        # *: PREPARE DATA FOR ENROLLMENT LISTING

        # create dictionary for enrollment listing
        input_dict = {
            "DM": {
                "Legal Sex (ig_DM1.SEX)": "Legal Sex",
                "Sex Assigned at Birth (ig_DM1.BRTHSEX)": "Sex Assigned at Birth",
                "Gender Identity (ig_DM1.GENDERID)": "Gender Identity",
                "Specify Other Gender Identity (ig_DM1.GENDERIDOTH)": "Other Gender",
                "Date of Birth (ig_DM1.BRTHDAT)": "Date of Birth",
                "Apheresis Consent Date (ig_DM1.RFICDAT)": "Consent Date",
                "Race (ig_DM1.RACE)": "Race",
                "Specify Other or Multiple Races (ig_DM1.RACEOTH)": "Other Race",
                "Ethnicity (ig_DM1.ETHNIC)": "Ethnicity",
            },
            "DSCA": {"Cohort Assignment (ig_DSCA1.CACHASCOD)": "Cohort"},
            "IE": {
                "Main Consent Date (ig_IE1.MAINCDAT)": "Main Consent Date",
                "Subject Meets All Study Eligibility (ig_IE3.IEYN)": "Subject meets all study eligibility?",
                "Other Screen Fail Reason (ig_IE4.OTHRSFREAS)": "SF3",
                "Screen Failure Reason (ig_IE4.IECAT)": "Reason for Screen Failure",
                "Select the Primary Inclusion Criterion Excluding This Subject  (ig_IE4.ITESTCD)": "SF1",
                "Select the Primary Exclusion Criterion Excluding This Subject (ig_IE4.ETESTCD)": "SF2",
            },
            "EXINF": {
                "Event Group Label": "Event Group Label",
                "Was study treatment administered? (ig_EXINF1.INFOCCUR)": "Treated",
            },
            "DSEOS": {
                "End of Study Date (ig_DSEOS1.EOSDAT)": "End of Study Date",
                "Reason for End of Study? (ig_DSEOS2.EOSCOD1)": "End of Study Reason",
                "Provide Supportive Information (ig_DSEOS2.EOSTERM)": "Supportive Information",
            },
            "MHDIAG": {
                "Primary Diagnosis (IG_NS_NA_MHDIAG1.CL_NS_NH_PRMDIAG_cl_NS_MHDIAGGB1)": "Disease",
                "Specify Other Diagnosis (IG_NS_NA_MHDIAG1.TX_NS_NH_MHDIAGOTH)": "Disease2",
            },
        }

        enrollment_df = get_data_from_dict(data, input_dict)
        # Replace missing values in the "Cohort" column with "Pending"
        enrollment_df["Cohort"] = enrollment_df["Cohort"].fillna("Pending")

        enrollment_df = age_calculation(
            enrollment_df,
            "Age at Consent",
            "Date of Birth",
            "Consent Date",
            "Main Consent Date",
        )
        # fill NaN with empty string
        enrollment_df = enrollment_df.fillna("")
        enrollment_df.loc[
            enrollment_df["Disease"] == "Other",
            "Disease",
        ] = ""
        enrollment_df["Disease Type"] = enrollment_df["Disease"].fillna("") + " " + enrollment_df["Disease2"].fillna("")
        enrollment_df["Disease Type"].fillna(enrollment_df["Disease Type"], inplace=True)

        # Convert the entire column to string to avoid data type issues
        enrollment_df["Reason for Screen Failure"] = enrollment_df["Reason for Screen Failure"].astype(str)

        # if "Reason for Screen Failure" in enrollment_df.columns equal "Other", replace the value with SF3
        mask = enrollment_df["Reason for Screen Failure"] == "Other"
        # Replace "Other" with the corresponding values from the "SF3" column
        enrollment_df.loc[mask, "Reason for Screen Failure"] = enrollment_df.loc[mask, "SF3"]

        # if "Reason for Screen Failure" in enrollment_df.columns equal "Inclusion Criteria", concat the value of the column with "SF1"
        # Create a mask for rows where "Reason for Screen Failure" is "Inclusion Criteria"
        mask = (enrollment_df["Reason for Screen Failure"] == "Inclusion Criteria") & (enrollment_df["SF1"] != "")
        # Use the mask to update the "Reason for Screen Failure" column
        enrollment_df.loc[mask, "Reason for Screen Failure"] = (
            enrollment_df.loc[mask, "Reason for Screen Failure"] + " " + enrollment_df.loc[mask, "SF1"].astype(str)
        )

        # Create a mask for rows where "Reason for Screen Failure" is "Exclusion Criteria" and "SF2" is not empty
        mask = (enrollment_df["Reason for Screen Failure"] == "Exclusion Criteria") & (enrollment_df["SF2"] != "")

        # Use the mask to update the "Reason for Screen Failure" column
        enrollment_df.loc[mask, "Reason for Screen Failure"] = (
            enrollment_df.loc[mask, "Reason for Screen Failure"] + " " + enrollment_df.loc[mask, "SF2"].astype(str)
        )

        # if subject does not have IE data, check if the subject has DSEOS "End of Study Date" data. If yes, then the subject is "No" for "Subject meets all study eligibility?"
        enrollment_df.loc[
            (enrollment_df["Subject meets all study eligibility?"] != "Yes")
            & (enrollment_df["End of Study Date"].isnull()),
            "Subject meets all study eligibility?",
        ] = "Pending"
        enrollment_df.loc[
            (enrollment_df["Subject meets all study eligibility?"] != "Yes")
            & (enrollment_df["End of Study Date"].notna()),
            "Subject meets all study eligibility?",
        ] = "No"
        enrollment_df.loc[
            (enrollment_df["Subject meets all study eligibility?"] != "Yes")
            & (enrollment_df["End of Study Date"].notna()),
            "Reason for Screen Failure",
        ] = enrollment_df["Supportive Information"]
        # Add Screen Fail column
        enrollment_df["Screen Fail"] = None

        # Define conditions and corresponding values
        conditions = [
            enrollment_df["Subject meets all study eligibility?"].fillna("Unknown") == "No",
            enrollment_df["Subject meets all study eligibility?"].fillna("Unknown") == "Yes",
            enrollment_df["Subject meets all study eligibility?"].fillna("Unknown") == "Pending",
        ]

        values = [
            "Yes",
            "No",
            "Pending",
        ]

        # Use np.select to assign values based on conditions
        enrollment_df["Screen Fail"] = np.select(conditions, values, default="Unknown")
        # drop the columns that are not needed
        enrollment_df = enrollment_df.drop(
            columns=[
                "Disease",
                "Disease2",
                "SF1",
                "SF2",
                "SF3",
                "Supportive Information",
                "End of Study Reason",
            ]
        )

        # if add this filter, subjects with retx will be removed from enrollment_df because get_data_from_dict get the latest data which excluded Day 0 data and included Day 0-R1 data
        # enrollment_df = enrollment_df[enrollment_df["Event Group Label"] != "Day 0-R1"]
        enrollment_df = enrollment_df.drop(columns=["Event Group Label"])
        # Update 'Treated' column based on the conditions:
        enrollment_df.loc[
            (enrollment_df["Treated"] != "Yes") & (enrollment_df["End of Study Date"].isnull()),
            "Treated",
        ] = "Pending"
        enrollment_df.loc[
            (enrollment_df["Treated"] != "Yes") & (~enrollment_df["End of Study Date"].isnull()),
            "Treated",
        ] = "No"
        # enrollment_df = enrollment_df.drop(columns=["End of Study Date"])
        # Sort
        enrollment_df = enrollment_df.sort_values(["Subject"])

        # prepare the output dataframe
        enrollment_output_df = enrollment_df.copy()
        for col in enrollment_output_df.columns:
            if "Date" in col:
                enrollment_output_df[col] = enrollment_output_df[col].dt.strftime("%m/%d/%Y")
        # If Gender is "Other", replace the value with "Other Gender"
        enrollment_output_df.loc[
            enrollment_output_df["Gender Identity"] == "Other",
            "Gender Identity",
        ] = enrollment_output_df["Other Gender"]
        # Use `.loc` to filter where Race is "Other" and update it with the value in "Other Race"
        enrollment_output_df.loc[enrollment_output_df["Race"] == "Other", "Race"] = (
            "Other- " + enrollment_output_df.loc[enrollment_output_df["Race"] == "Other", "Other Race"].astype(str)
        )
        # *Re-order the columns and remove the columns that are not needed
        enrollment_output_df = enrollment_output_df[
            [
                "Subject",
                "Cohort",
                "Disease Type",
                "Legal Sex",
                "Gender Identity",
                "Sex Assigned at Birth",
                "Age at Consent",
                "Race",
                "Ethnicity",
                "Screen Fail",
                "Reason for Screen Failure",
                "Treated",
            ]
        ]
        return enrollment_output_df, enrollment_df

    def enrollment_stat_table(self, enrollment_df):
        ### TODO: Demo Stats Table
        # !Update this filter options to each cohort
        filter_options = [
            enrollment_df["Consent Date"].notna() | enrollment_df["Main Consent Date"].notna(),
            enrollment_df["Cohort"] == "Cohort 1",
            enrollment_df["Cohort"] == "Cohort 2",
            enrollment_df["Cohort"] == "Cohort -1",
        ]
        self.status_list = []
        self.LegalSex_list = []
        self.Age_at_Consent_list = []
        self.Race_list = []
        self.Ethnicity_list = []

        for filter_index, filter_option in enumerate(filter_options):
            # Apply the filter to the dataframe
            filtered_df = enrollment_df[filter_option].copy()
            filtered_df = filtered_df[
                (filtered_df["Consent Date"].notna()) | (filtered_df["Main Consent Date"].notna())
            ]
            # Calculate the stats
            ## Total Consented
            TT_df = filtered_df.copy()
            TT = filtered_df["Subject"].count()
            ## Screen Failed
            SF_df = filtered_df[filtered_df["Subject meets all study eligibility?"] == "No"].copy()
            SF = SF_df["Subject"].count()
            ## Eligible
            EL_df = filtered_df[filtered_df["Subject meets all study eligibility?"] == "Yes"].copy()
            EL = EL_df["Subject"].count()
            ## Treated
            INFR_df = filtered_df[filtered_df["Treated"] == "Yes"].copy()
            INF = INFR_df["Subject"].count()

            # Define a dictionary containing the status of each variable
            self.status_list.append(
                {
                    "Total Consented": TT,
                    "Screen Failed": SF,
                    "Eligible": EL,
                    "Treated": INF,
                }
            )

            # Calculate the stats for the filtered dataframe
            Legal_Sex_Codelist = ["Male", "Female", "X (Nonbinary)", "Not Reported"]
            self.LegalSex_list.append(
                get_stats_percentage2("Legal Sex", Legal_Sex_Codelist, TT_df, SF_df, EL_df, INFR_df)
            )
            self.Age_at_Consent_list.append(get_stats_df("Age at Consent", TT_df, SF_df, EL_df, INFR_df))
            self.Race_list.append(get_stats_percentage("Race", TT_df, SF_df, EL_df, INFR_df))
            self.Ethnicity_list.append(get_stats_percentage("Ethnicity", TT_df, SF_df, EL_df, INFR_df))

    # def EGFR_listing(self):
    #     data = self.data
    #     # Find eligible subjects
    #     IE_df = data["IE"][["Subject", "Subject Meets All Study Eligibility (ig_IE3.IEYN)"]].copy()
    #     IE_new_col_name = {
    #         "Subject Meets All Study Eligibility (ig_IE3.IEYN)": "Subject meets all study eligibility?",
    #     }
    #     IE_df = IE_df.rename(columns=IE_new_col_name)
    #     Eligible_df = IE_df[IE_df["Subject meets all study eligibility?"] == "Yes"].copy()

    #     # *: Get latest date from Sugery in Initial Study Enrollment/Apheresis event
    #     MHSG_subject_df = Eligible_df["Subject"].copy()
    #     MHSG1_df = MHSG_subject_df
    #     MHSG1_df = add_rename_column_corelisting(MHSG1_df, data, "MHSG", "Event Group Label", "Event Group Label")

    #     MHSG1_df = MHSG1_df[(MHSG1_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")].copy()

    #     MHSG_collectionDT_df = data["MHSG"][
    #         [
    #             "Subject",
    #             "Event Group Label",
    #             "Was EGFR amplification testing performed on this sample? (IG_NS_NA_MHSG4.CL_NS_NH_EGFRPERF_cl_YS_YN1)",
    #             "Date of Surgery (IG_NS_NA_MHSG2.DT_NS_NH_SGDAT)",
    #         ]
    #     ].copy()

    #     MHSG_new_col_name = {
    #         "Was EGFR amplification testing performed on this sample? (IG_NS_NA_MHSG4.CL_NS_NH_EGFRPERF_cl_YS_YN1)": "Was EGFR amplification testing performed on this sample?",
    #         "Date of Surgery (IG_NS_NA_MHSG2.DT_NS_NH_SGDAT)": "Date of Surgery",
    #     }

    #     MHSG_collectionDT_df = MHSG_collectionDT_df.rename(columns=MHSG_new_col_name)

    #     # Filter the DataFrame based on the conditions
    #     filtered_MHSGINIT_df = MHSG_collectionDT_df[
    #         (MHSG_collectionDT_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")
    #         & (MHSG_collectionDT_df["Was EGFR amplification testing performed on this sample?"] == "Yes")
    #         & (MHSG_collectionDT_df["Date of Surgery"].notna())
    #     ].copy()

    #     # Sort and get the last row for each subject, need test result from last collection date
    #     filtered_MHSGINIT_df = filtered_MHSGINIT_df.sort_values(["Date of Surgery"])
    #     filtered_MHSGINIT_df = filtered_MHSGINIT_df.groupby("Subject").tail(1)  # Get the last row per subject
    #     filtered_MHSGINIT_df = filtered_MHSGINIT_df.sort_values(["Subject"])
    #     # print(filtered_MHSGINIT_df)
    #     # drop "Event Group Label" and "Was EGFR amplification testing performed on this sample?" column
    #     filtered_MHSGINIT_df = filtered_MHSGINIT_df.drop(
    #         columns=["Event Group Label", "Was EGFR amplification testing performed on this sample?"]
    #     )
    #     # Renmae filtered_MHSGINIT_df["Date of Surgery"] to "Collection Date"
    #     filtered_MHSGINIT_df = filtered_MHSGINIT_df.rename(columns={"Date of Surgery": "Collection Date"})

    #     # *: PREPARE DATA FOR EGFR LISTING
    #     # EGFR_subject_df = Eligible_df["Subject"].copy()
    #     # EGFR1_df = EGFR_subject_df
    #     # EGFR1_df = add_rename_column_corelisting(EGFR1_df, data, "LBEGFR", "Event Group Label", "Event Group Label")

    #     # EGFR1_df = EGFR1_df[(EGFR1_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")].copy()

    #     EGFR_df = data["LBEGFR"][
    #         [
    #             "Subject",
    #             # "Event Group Label",
    #             # "Was EGFR Amplification testing performed? (IG_NS_NA_LBEGFR1.CL_NS_NH_EGFRAMPERF_cl_YS_YN1)",
    #             "Collection Date (IG_NS_NA_LBEGFR1.DT_YS_NH_LBDAT)",
    #             "Amplification of EGFR (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR_cl_YS_DTNDT1)",
    #             "EGFRvIII Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR8_cl_YS_DTNDT1)",
    #             "EGFR Extracellular Domain Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_EGFRMUT_cl_YS_DTNDT1)",
    #         ]
    #     ].copy()

    #     EGFR_new_col_name = {
    #         # "Was EGFR Amplification testing performed? (IG_NS_NA_LBEGFR1.CL_NS_NH_EGFRAMPERF_cl_YS_YN1)": "Was EGFR Amplification testing performed?",
    #         "Collection Date (IG_NS_NA_LBEGFR1.DT_YS_NH_LBDAT)": "Collection Date",
    #         "Amplification of EGFR (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR_cl_YS_DTNDT1)": "Amplification of EGFR",
    #         "EGFRvIII Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR8_cl_YS_DTNDT1)": "EGFRvIII Mutation",
    #         "EGFR Extracellular Domain Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_EGFRMUT_cl_YS_DTNDT1)": "EGFR Extracellular Domain Mutation",
    #     }

    #     EGFR_df = EGFR_df.rename(columns=EGFR_new_col_name)

    #     # Merge the filtered_MHSGINIT_df with the EGFR_collectionDT_df on "Subject" and "Collection Date"
    #     EGFR_df = pd.merge(
    #         filtered_MHSGINIT_df[["Subject", "Collection Date"]].drop_duplicates(),
    #         EGFR_df,
    #         on=["Collection Date", "Subject"],
    #         how="left",
    #     )

    #     # Filter the DataFrame based on the conditions
    #     # filtered_INIT_df = EGFR_collectionDT_df[
    #     #     (EGFR_collectionDT_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")
    #     #     & (EGFR_collectionDT_df["Was EGFR Amplification testing performed?"] == "Yes")
    #     #     & (EGFR_collectionDT_df["Collection Date"].notna())
    #     # ].copy()

    #     # # Sort and get the last row for each subject, need test result from last collection date
    #     # filtered_INIT_df = filtered_INIT_df.sort_values(["Collection Date"])
    #     # filtered_INIT_df = filtered_INIT_df.groupby("Subject").tail(2)  # Get the last 2 rows per subject
    #     # filtered_INIT_df = filtered_INIT_df.sort_values(["Subject"])

    #     replacement_date = pd.Timestamp("1900-01-01")
    #     # filtered_INIT_df["Collection Date"] = filtered_INIT_df["Collection Date"].fillna(replacement_date)

    #     # Define the columns to concatenate
    #     columns_to_concatenate_init = [
    #         "Amplification of EGFR",
    #         "EGFRvIII Mutation",
    #         "EGFR Extracellular Domain Mutation",
    #     ]

    #     # Perform aggregation and retain only the first non-empty value for each column, handling NaN properly
    #     agg_init_df = (
    #         EGFR_df.groupby(["Subject", "Collection Date"])[columns_to_concatenate_init]
    #         .apply(
    #             lambda group: group.apply(lambda col: col.dropna().iloc[0] if not col.dropna().empty else "")
    #         )  # Retain first non-empty value
    #         .reset_index()
    #     )

    #     # Fill any remaining NaNs in the "Collection Date" column with the placeholder date
    #     agg_init_df["Collection Date"] = agg_init_df["Collection Date"].fillna(replacement_date)

    #     EGFR_df["Collection Date"] = EGFR_df["Collection Date"].fillna(replacement_date)
    #     EGFR_df["Collection Date"] = pd.to_datetime(EGFR_df["Collection Date"])

    #     EGFR_final_df = agg_init_df.sort_values(["Subject"])
    #     # print(EGFR_final_df)

    #     MHDIAG_df = data["MHDIAG"][["Subject", "MGMT Result (IG_NS_NA_MHDIAG2.CL_NS_NH_MGMTRES_cl_NS_MGMTRES1)"]].copy()
    #     MHDIAG_new_col_name = {
    #         "MGMT Result (IG_NS_NA_MHDIAG2.CL_NS_NH_MGMTRES_cl_NS_MGMTRES1)": "MGMT Result",
    #     }
    #     MHDIAG_df = MHDIAG_df.rename(columns=MHDIAG_new_col_name)
    #     MHDIAG_df = pd.merge(
    #         Eligible_df,
    #         MHDIAG_df,
    #         on=["Subject"],
    #         how="left",
    #     )

    #     EGFR_final2_df = pd.merge(EGFR_final_df, MHDIAG_df, on="Subject", how="right")
    #     EGFR_final2_df["Collection Date"] = EGFR_final2_df["Collection Date"].fillna(replacement_date)
    #     EGFR_final2_df = EGFR_final2_df.sort_values(["Subject"])
    #     EGFR_final2_df = EGFR_final2_df.fillna("")

    #     # replace "" with "Not Done" for all columns except "Subject" and "Collection Date"
    #     EGFR_final2_df.loc[EGFR_final2_df["Amplification of EGFR"] == "", "Amplification of EGFR"] = (
    #         "Not Done"  # Replace empty strings with "Not Done"
    #     )
    #     EGFR_final2_df.loc[EGFR_final2_df["EGFRvIII Mutation"] == "", "EGFRvIII Mutation"] = "Not Done"
    #     EGFR_final2_df.loc[
    #         EGFR_final2_df["EGFR Extracellular Domain Mutation"] == "", "EGFR Extracellular Domain Mutation"
    #     ] = "Not Done"

    #     EGFR_output_df = EGFR_final2_df.copy()
    #     EGFR_output_df = EGFR_output_df[
    #         [
    #             "Subject",
    #             "Amplification of EGFR",
    #             "EGFRvIII Mutation",
    #             "EGFR Extracellular Domain Mutation",
    #             "MGMT Result",
    #         ]
    #     ]
    #     return EGFR_output_df, EGFR_final2_df

    # def infusion_listing(self):
    #     data = self.data
    #     ### TODO: INFUSION LISTING
    #     # adding Target Cell Dose dictionary
    #     # !: Update this dictionary to the new study
    #     TCD_dict = {
    #         "Cohort -1": 5000000,
    #         "Cohort 1": 10000000,
    #         "Cohort 2": 25000000,
    #         "Cohort 3": 50000000,
    #         "Not Assigned": "Not Assigned",
    #     }

    #     # *: PREPARE DATA FOR INFUSION LISTING

    #     # create dictionary for enrollment listing
    #     input_dict1 = {
    #         "EXINF": {
    #             "Event Group Label": "Event Group Label",
    #             "Study Treatment Date (ig_EXINF1.INFDAT)": "Study Treatment Date",
    #             "Volume CSF Removed for Cell Product Administration (mL) (ig_EXINF1.CSFVOL)": "Volume CSF Withdrawn",
    #             "Total Volume Administered (mL) (ig_EXINF1.INFTOTVOL)": "Volume Dose Administered",
    #             "CAR T Cell Dose Administered (ig_EXINF1.INFDOS)": "CART-EGFR-IL13Rα2 Cell Dose",
    #             "x 10 to the power of (ig_EXINF1.INFDOSXP)": "x 10 to the power of (ig_EXINF1.INFDOSXP)",
    #             "Total Cell Dose Administered (ig_EXINF1.INFDOSTOT)": "Total Cell Dose",
    #             "x 10 to the power of (ig_EXINF1.INFDOSTOTXP)": "x 10 to the power of (ig_EXINF1.INFDOSTOTXP)",
    #             "EGFR Transduction Efficiency (%) (ig_EXINF1.EGFRINFTEFFP)": "%scFV (EGFR)",
    #             "IL13Ra2 Transduction Efficiency (%) (ig_EXINF1.IL13INFTEFFP)": "%scFV (Il13Rα2)",
    #             # "Event Date": "Event Date INF",
    #         },
    #     }
    #     input_dict2 = {
    #         "DSCA": {"Cohort Assignment (ig_DSCA1.CACHASCOD)": "Cohort"},
    #     }

    #     # For infusion CRF, get the earliest data (Day 0)
    #     raw_infusion_df1 = get_data_from_dict_first(data, input_dict1)
    #     raw_infusion_df2 = get_data_from_dict(data, input_dict2)
    #     raw_infusion_df = pd.merge(raw_infusion_df1, raw_infusion_df2, on="Subject", how="left")
    #     # convert the date to datetime object and format it to MM-DD-YYYY
    #     raw_infusion_df["Study Treatment Date"] = raw_infusion_df["Study Treatment Date"].apply(
    #         lambda x: datetime.strptime(x.strftime("%Y-%m-%d"), "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
    #     )

    #     # print(raw_infusion_df)
    #     # TODO: INFUSION LISTING Day 0

    #     infusion_df = raw_infusion_df[raw_infusion_df["Event Group Label"] == "Day 0"]
    #     # print(infusion_df)

    #     # adding Target Dose using TCD_dict
    #     infusion_df["Target Dose"] = infusion_df["Cohort"].map(TCD_dict)

    #     # combine CART-EGFR-IL13Rα2 Cell Dose and x 10 to the power of (ig_EXINF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the CART-EGFR-IL13Rα2 Cell Dose column to string
    #     infusion_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusion_df["CART-EGFR-IL13Rα2 Cell Dose"].multiply(
    #         10 ** infusion_df["x 10 to the power of (ig_EXINF1.INFDOSXP)"]
    #     )
    #     infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSXP)"])
    #     infusion_df["Total Cell Dose"] = infusion_df["Total Cell Dose"].multiply(
    #         10 ** infusion_df["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"]
    #     )
    #     infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"])

    #     # Adding Met Target Dose column based on the condition of Total Cell Dose and CART-EGFR-IL13Rα2 Cell Dose if 'Target Cell Dose' is integer
    #     infusion_df["Met Target Dose (Y/N)"] = infusion_df.apply(
    #         lambda row: "Y"
    #         if isinstance(row["Target Dose"], int) and row["CART-EGFR-IL13Rα2 Cell Dose"] >= row["Target Dose"]
    #         else "",
    #         axis=1,
    #     )
    #     infusion_df["Met Target Dose (Y/N)"] = infusion_df.apply(
    #         lambda row: "N"
    #         if isinstance(row["Target Dose"], int) and row["CART-EGFR-IL13Rα2 Cell Dose"] < row["Target Dose"]
    #         else row["Met Target Dose (Y/N)"],
    #         axis=1,
    #     )

    #     # adding Met Target %scFv and fillter out the rows that have NaN in Met Target %scFv
    #     infusion_df["Met Target % scFV Flow (Y/N) (≥2%)"] = infusion_df[infusion_df["%scFV (EGFR)"].notna()][
    #         "%scFV (EGFR)"
    #     ].apply(lambda x: "Y" if x >= 2 else "N")
    #     # fill NaN with empty string
    #     infusion_df = infusion_df.fillna(
    #         "",
    #     ).infer_objects(copy=False)

    #     # Only keep the rows that have Event Group Label
    #     infusion_df = infusion_df[infusion_df["Event Group Label"] != ""]

    #     # *Re-order the columns and remove the columns that are not needed
    #     infusion_df = infusion_df[
    #         [
    #             "Subject",
    #             "Event Group Label",
    #             "Cohort",
    #             "Target Dose",
    #             "Study Treatment Date",
    #             "Volume CSF Withdrawn",
    #             "Volume Dose Administered",
    #             "Total Cell Dose",
    #             "CART-EGFR-IL13Rα2 Cell Dose",
    #             "Met Target Dose (Y/N)",
    #             "%scFV (EGFR)",
    #             "Met Target % scFV Flow (Y/N) (≥2%)",
    #             "%scFV (Il13Rα2)",
    #         ]
    #     ]

    #     # TODO: Infusion Listing Day 0-R1, Day 0-R2
    #     # Get the latest data to include Day 0-R1 and Day 0-R2 data
    #     raw_infusionR_df1 = get_data_from_dict(data, input_dict1)
    #     raw_infusionR_df2 = get_data_from_dict(data, input_dict2)
    #     raw_infusionR_df = pd.merge(raw_infusionR_df1, raw_infusionR_df2, on="Subject", how="left")
    #     # convert the date to datetime object and format it to MM-DD-YYYY
    #     raw_infusionR_df["Study Treatment Date"] = raw_infusionR_df["Study Treatment Date"].apply(
    #         lambda x: datetime.strptime(x.strftime("%Y-%m-%d"), "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
    #     )
    #     infusionR_df = raw_infusionR_df[
    #         (raw_infusionR_df["Event Group Label"] == "Day 0-R1")
    #         | (raw_infusionR_df["Event Group Label"] == "Day 0-R2")
    #     ]

    #     # combine CART-EGFR-IL13Rα2 Cell Dose and x 10 to the power of (ig_EXINF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the CART-EGFR-IL13Rα2 Cell Dose column to string
    #     infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"].multiply(
    #         10 ** infusionR_df["x 10 to the power of (ig_EXINF1.INFDOSXP)"]
    #     )
    #     infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSXP)"])

    #     infusionR_df["Total Cell Dose"] = infusionR_df["Total Cell Dose"].multiply(
    #         10 ** infusionR_df["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"]
    #     )
    #     infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"])

    #     # adding Met Target %scFv
    #     infusionR_df = add_rename_column_corelisting(
    #         infusionR_df,
    #         data,
    #         "EXINF",
    #         "IL13Ra2 Transduction Efficiency (%) (ig_EXINF1.IL13INFTEFFP)",
    #         "Met Target % scFV Flow (Y/N) (≥2%)",
    #         "Subject",
    #         "Event Group Label",
    #     )
    #     # adding Met Target %scFv and fillter out the rows that have NaN in Met Target %scFv
    #     infusionR_df["Met Target % scFV Flow (Y/N) (≥2%)"] = infusionR_df[
    #         infusionR_df["Met Target % scFV Flow (Y/N) (≥2%)"].notna()
    #     ]["Met Target % scFV Flow (Y/N) (≥2%)"].apply(lambda x: "Y" if x >= 2 else "N")
    #     # fill NaN with empty string
    #     infusionR_df = infusionR_df.fillna("")

    #     # Only keep the rows that have Event Group Label
    #     infusionR_df = infusionR_df[infusionR_df["Event Group Label"] != ""]

    #     # *Re-order the columns and remove the columns that are not needed
    #     infusionR_df = infusionR_df[
    #         [
    #             "Subject",
    #             "Event Group Label",
    #             "Cohort",
    #             "Study Treatment Date",
    #             "Volume CSF Withdrawn",
    #             "Volume Dose Administered",
    #             "Total Cell Dose",
    #             "CART-EGFR-IL13Rα2 Cell Dose",
    #             "%scFV (EGFR)",
    #             "Met Target % scFV Flow (Y/N) (≥2%)",
    #             "%scFV (Il13Rα2)",
    #         ]
    #     ]

    #     return infusion_df, infusionR_df

    # def infusion_stats(self, infusion_df, infusionR_df):
    #     # TODO: INFUSION STATISTICS
    #     infusion_count = []
    #     # * Cohort 1
    #     # Create a new dataframe for Cohort 1 with infusion_df
    #     infusion1_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort 1"]
    #     # Create a new dataframe for Total Cell Dose table with infusion_df
    #     infusion_statA1 = get_stats_df("Total Cell Dose", infusion1_df)

    #     infusion_statA2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion1_df)

    #     # Count the number of subjects that met the target dose
    #     met_targetDose_count = infusion1_df[infusion1_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
    #     # Count the number of subjects
    #     total_subject_count = infusion1_df["Subject"].nunique()
    #     # Add column "Met Target Dose (Y/N)" to infusion_statA2 data frame, since data frames are concated, the order of the data frame should match the output order
    #     infusion_statA2["Met Target Dose (Y/N)"] = (
    #         str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFv Flow table with infusion_df
    #     infusion_statA3 = get_stats_perc_df("%scFV (EGFR)", infusion1_df)
    #     # Count the number of subjects that met the target %scFv
    #     met_targetFlow_count = infusion1_df[infusion1_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
    #         "Subject"
    #     ]
    #     infusion_statA3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
    #         str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
    #     infusion_statA4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion1_df)

    #     # Combine the three dataframes
    #     infusion_statA = pd.concat(
    #         [infusion_statA1, infusion_statA2, infusion_statA3, infusion_statA4],
    #         axis=1,
    #     )
    #     infusion_statA = infusion_statA.replace([np.inf, -np.inf], "")
    #     infusion_statA = infusion_statA.fillna("")
    #     self.infusion_statA = infusion_statA
    #     infusion_count.append(total_subject_count)
    #     self.infusion_count = infusion_count

    #     # * Cohort 2
    #     # Create a new dataframe for Cohort 2 with infusion_df
    #     infusion2_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort 2"]
    #     infusion_statB1 = get_stats_df("Total Cell Dose", infusion2_df)

    #     infusion_statB2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion2_df)

    #     # Count the number of subjects that met the target dose
    #     # infusion_statB3 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion2_df)
    #     met_targetDose_count = infusion2_df[infusion2_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
    #     # Count the number of subjects
    #     total_subject_count = infusion2_df["Subject"].nunique()
    #     infusion_statB2["Met Target Dose (Y/N)"] = (
    #         str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFv Flow table with infusion_df
    #     infusion_statB3 = get_stats_perc_df("%scFV (EGFR)", infusion2_df)
    #     # Count the number of subjects that met the target %scFv
    #     met_targetFlow_count = infusion2_df[infusion2_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
    #         "Subject"
    #     ]
    #     infusion_statB3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
    #         str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
    #     infusion_statB4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion2_df)

    #     # Combine the three dataframes
    #     infusion_statB = pd.concat([infusion_statB1, infusion_statB2, infusion_statB3, infusion_statB4], axis=1)
    #     infusion_statB = infusion_statB.replace([np.inf, -np.inf], "")
    #     infusion_statB = infusion_statB.fillna("")
    #     self.infusion_statB = infusion_statB
    #     infusion_count.append(total_subject_count)
    #     self.infusion_count = infusion_count

    #     # * Cohort -1
    #     # Create a new dataframe for Cohort 2 with infusion_df
    #     infusion3_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort -1"]
    #     infusion_statC1 = get_stats_df("Total Cell Dose", infusion3_df)

    #     infusion_statC2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion3_df)

    #     # Count the number of subjects that met the target dose
    #     # infusion_statC3 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion3_df)
    #     met_targetDose_count = infusion3_df[infusion3_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
    #     # Count the number of subjects
    #     total_subject_count = infusion3_df["Subject"].nunique()
    #     infusion_statC2["Met Target Dose (Y/N)"] = (
    #         str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFv Flow table with infusion_df
    #     infusion_statC3 = get_stats_perc_df("%scFV (EGFR)", infusion3_df)
    #     # Count the number of subjects that met the target %scFv
    #     met_targetFlow_count = infusion3_df[infusion3_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
    #         "Subject"
    #     ]
    #     infusion_statC3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
    #         str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
    #     )

    #     # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
    #     infusion_statC4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion3_df)
    #     # Combine the three dataframes
    #     infusion_statC = pd.concat([infusion_statC1, infusion_statC2, infusion_statC3, infusion_statC4], axis=1)
    #     infusion_statC = infusion_statC.replace([np.inf, -np.inf], "")
    #     infusion_statC = infusion_statC.fillna("")
    #     self.infusion_statC = infusion_statC
    #     infusion_count.append(total_subject_count)
    #     self.infusion_count = infusion_count

    #     ## TODO: FORMATTING THE DATAFRAME
    #     # TODO: Day 0
    #     # Convert the columns to scientific notation if the value is not NaN
    #     infusion_df["Target Dose"] = infusion_df["Target Dose"].apply(
    #         lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
    #     )
    #     infusion_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusion_df["CART-EGFR-IL13Rα2 Cell Dose"].apply(
    #         lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
    #     )
    #     infusion_df["Total Cell Dose"] = infusion_df["Total Cell Dose"].apply(
    #         lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
    #     )
    #     # adding '%' sign to %scFv Flow
    #     infusion_df["%scFV (EGFR)"] = infusion_df.apply(
    #         lambda row: str(x) + "%" if pd.notna(x := row["%scFV (EGFR)"]) else x, axis=1
    #     )
    #     infusion_df["%scFV (Il13Rα2)"] = infusion_df.apply(
    #         lambda row: str(x) + "%" if pd.notna(x := row["%scFV (Il13Rα2)"]) else x, axis=1
    #     )

    #     # TODO: Day 0-R1
    #     # Convert the columns to scientific notation if the value is not NaN
    #     infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"].apply(
    #         lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    #     )
    #     infusionR_df["Total Cell Dose"] = infusionR_df["Total Cell Dose"].apply(
    #         lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    #     )
    #     # adding '%' sign to %scFv Flow
    #     infusionR_df["%scFV (EGFR)"] = infusionR_df.apply(
    #         lambda row: str(x) + "%" if pd.notna(x := row["%scFV (EGFR)"]) else x, axis=1
    #     )
    #     infusionR_df["%scFV (Il13Rα2)"] = infusionR_df.apply(
    #         lambda row: str(x) + "%" if pd.notna(x := row["%scFV (Il13Rα2)"]) else x, axis=1
    #     )

    #     data = self.data

    # def status_listing(self, enrollment_listing_df):
    #     data = self.data
    #     # AE and SAE data
    #     #    if not data["AE"].empty:
    #     AE_df = data["AE"][
    #         [
    #             "Subject",
    #             "AE or SAE? (ig_AE2.AESEV)",
    #         ]
    #     ].copy()
    #     AE_new_col_name = {
    #         "AE or SAE? (ig_AE2.AESEV)": "AE or SAE?",
    #     }
    #     AE_df = AE_df.rename(columns=AE_new_col_name)

    #     # # replaces all occurrences of NaN, positive infinity, and negative infinity in the infusion_df dataframe with empty strings.
    #     # infusion_df = infusion_df.replace([np.nan, np.inf, -np.inf], "N/A")
    #     status_df = self.enrollment_listing_df[self.enrollment_listing_df["Screen Fail"].str.strip() == "No"][
    #         ["Subject", "Cohort"]
    #     ]

    #     status_df["AE"] = status_df["Subject"].apply(lambda x: "Y" if x in AE_df["Subject"].values else "N")

    #     status_df["SAE"] = status_df["Subject"].apply(
    #         lambda x: "Y" if x in AE_df[AE_df["AE or SAE?"] == "SAE"]["Subject"].values else "N"
    #     )
    #     # replaces all occurrences of NaN, positive infinity, and negative infinity in the infusion_df dataframe with empty strings.
    #     status_df = status_df.replace([np.nan, np.inf, -np.inf], "")

    #     # Event Label Update dictionary
    #     event_1_dict = {
    #         "Pre-Treatment Safety Visit": "Pre-Treatment",
    #         "Primary Treatment and Follow-Up": "Primary Follow-up",
    #         "Long-Term Follow-Up Months 3-60": "LTFU",
    #         "Pre-Retreatment Safety 1": "Pre-Retreatment",
    #         "Primary Retreatment and Follow-Up 1": "Primary Retreatment Follow-up",
    #         "Retreatment Long-Term Follow-Up Months 3-60 (1)": "Retreatment LTFU",
    #     }

    #     # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
    #     status_SV_df = data["DSSV"][["Subject", "Event Label", "Event Date"]]
    #     # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
    #     status_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]

    #     # status_DSSVLTFU_df["Event Group Label"] = status_DSSVLTFU_df["Event Label"].apply(map_event)

    #     # Combine DSSVLTFU with SV dataframe vertically
    #     status_SV_df = pd.concat([status_SV_df, status_DSSVLTFU_df])
    #     # Sort the dataframe by Subject and Event Date
    #     status_SV_df = status_SV_df.sort_values(by=["Subject", "Event Date"])

    #     # For each unique subject, get the last row of the dataframe
    #     status_SV_df = status_SV_df.groupby("Subject").tail(1)

    #     # Merge left with the current response dataframe
    #     status_df = pd.merge(
    #         status_df,
    #         status_SV_df[["Subject", "Event Label"]],
    #         on="Subject",
    #         how="left",
    #     )

    #     # Rename the column Event Label to Event Label (Study Status)
    #     status_df["Event Label"] = status_df["Event Label"].map(event_1_dict)

    #     status_df["Event Label3"] = status_df["Subject"].apply(
    #         lambda x: "Pre-Treatment"
    #         if (
    #             self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Treated"]
    #             .fillna("")
    #             .str.strip()
    #             .values[0]
    #             == "Pending"
    #         )
    #         else ""
    #     )
    #     # print(self.enrollment_listing_df)
    #     status_df["Event Label4"] = status_df["Subject"].apply(
    #         lambda x: "Withdrawn Prior to Study Treatment"
    #         if (
    #             self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Treated"]
    #             .fillna("")
    #             .str.strip()
    #             .values[0]
    #             == "No"
    #         )
    #         & (
    #             self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Screen Fail"]
    #             .fillna("")
    #             .str.strip()
    #             .values[0]
    #             == "No"
    #         )
    #         else ""
    #     )
    #     # Merge all event group label into study status
    #     status_df["Event Label"] = (
    #         status_df["Event Label"].fillna("")
    #         + status_df["Event Label3"].fillna("")
    #         + status_df["Event Label4"].fillna("")
    #     )
    #     status_df["Event Label"].fillna(status_df["Event Label"], inplace=True)
    #     status_df = status_df.drop(
    #         columns=[
    #             "Event Label3",
    #             "Event Label4",
    #         ]
    #     )
    #     # print(status_df)

    #     # filter the data frame to only include subjects whose end of study date is later than or equal to main consent date
    #     filteredemrollment_df = self.enrollment_listing_df[
    #         self.enrollment_listing_df["End of Study Date"] >= self.enrollment_listing_df["Main Consent Date"]
    #     ]
    #     DSEOS_df = data["DSEOS"][
    #         [
    #             "Subject",
    #             "Reason for End of Study? (ig_DSEOS2.EOSCOD1)",
    #             "Provide Supportive Information (ig_DSEOS2.EOSTERM)",
    #             "Principal Cause of Death (ig_DSEOS2.PRCDTH)",
    #             "Specify Principal Cause of Death (ig_DSEOS2.PRCDTHOS)",
    #             "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)",
    #             "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)",
    #             "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)",
    #         ]
    #     ].copy()
    #     DSEOS_new_col_name = {
    #         "Reason for End of Study? (ig_DSEOS2.EOSCOD1)": "Off-Study Reason",
    #         "Provide Supportive Information (ig_DSEOS2.EOSTERM)": "Off-Study Reason sp1",
    #         "Principal Cause of Death (ig_DSEOS2.PRCDTH)": "Off-Study Reason sp2",
    #         "Specify Principal Cause of Death (ig_DSEOS2.PRCDTHOS)": "Off-Study Reason sp3",
    #         "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)": "Last Study Phase Completed",
    #         "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)": "Last Primary FUP",
    #         "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)": "Last Primary Retreatment",
    #     }
    #     DSEOS_df = DSEOS_df.rename(columns=DSEOS_new_col_name)

    #     # Merge last study visit
    #     DSEOS_df["Last Study Visit"] = DSEOS_df["Last Primary FUP"].fillna("") + DSEOS_df[
    #         "Last Primary Retreatment"
    #     ].fillna("")

    #     # Merge off-study reason
    #     DSEOS_df["Off-Study Reason"] = (
    #         DSEOS_df["Off-Study Reason"].fillna("")
    #         + " "
    #         + DSEOS_df["Off-Study Reason sp1"].fillna("")
    #         + DSEOS_df["Off-Study Reason sp2"].fillna("")
    #         + " "
    #         + DSEOS_df["Off-Study Reason sp3"].fillna("")
    #     )

    #     filteredDSEOS_df = DSEOS_df[(DSEOS_df["Subject"].isin(filteredemrollment_df["Subject"].values))].copy()
    #     # on perform the replace when status_df is not empty
    #     if not status_df.empty:
    #         status_df["Event Label"] = status_df.apply(
    #             lambda row: "Off Study"
    #             if (row["Subject"] in filteredDSEOS_df["Subject"].values)
    #             & ("Withdrawn Prior to Study Treatment" not in row["Event Label"])
    #             else "On Study/" + row["Event Label"],
    #             axis=1,
    #         )
    #         status_df = status_df.replace(
    #             "On Study/Withdrawn Prior to Study Treatment", "Withdrawn Prior to Study Treatment"
    #         )
    #         status_df = status_df.replace(
    #             "On Study/Pre-TreatmentWithdrawn Prior to Study Treatment", "Withdrawn Prior to Study Treatment"
    #         )
    #         status_df = pd.merge(
    #             status_df,
    #             filteredDSEOS_df[["Subject", "Off-Study Reason", "Last Study Visit"]],
    #             on="Subject",
    #             how="left",
    #         )
    #         status_df = status_df.replace("On Study/Pre-TreatmentPre-Treatment", "On Study/Pre-Treatment")

    #     # replaces all occurrences of NaN, positive infinity, and negative infinity with empty strings.
    #     status_df = status_df.replace([np.nan, np.inf, -np.inf], "N/A")

    #     # Gather all stats of each cohort
    #     total_status_df = status_df.copy()

    #     totalCH1_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort 1"])].copy()
    #     # Total number of subjects for Cohort 1
    #     AECH1_total_count = get_stats_percentage("AE", totalCH1_status_df).T
    #     SAECH1_total_count = get_stats_percentage("SAE", totalCH1_status_df).T
    #     # merge AE and SAE dataframes
    #     safetyCH1_total_df = pd.concat([AECH1_total_count, SAECH1_total_count], axis=1)

    #     totalCH2_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort 2"])].copy()
    #     # Total number of subjects for Cohort 2
    #     AECH2_total_count = get_stats_percentage("AE", totalCH2_status_df).T
    #     SAECH2_total_count = get_stats_percentage("SAE", totalCH2_status_df).T
    #     # merge AE and SAE dataframes
    #     safetyCH2_total_df = pd.concat([AECH2_total_count, SAECH2_total_count], axis=1)

    #     totalCHN1_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort -1"])].copy()
    #     # Total number of subjects for Cohort -1
    #     AECHN1_total_count = get_stats_percentage("AE", totalCHN1_status_df).T
    #     SAECHN1_total_count = get_stats_percentage("SAE", totalCHN1_status_df).T
    #     # merge AE and SAE dataframes
    #     safetyCHN1_total_df = pd.concat([AECHN1_total_count, SAECHN1_total_count], axis=1)

    #     return AE_df, status_df, safetyCH1_total_df, safetyCH2_total_df, safetyCHN1_total_df

    def export(self, output_dir, output_file_name):
        data = self.data
        with pd.ExcelWriter(output_dir + "/" + output_file_name + ".xlsx", engine="xlsxwriter") as writer:
            # TODO: - Add formatting and coloring
            # TODO: - for each tab: write data, format data, write header, format header

            ## * FORMATING AND COLORING
            bold_11_format = writer.book.add_format(
                {
                    "bg_color": "#FFFFFF",
                    "text_wrap": True,
                    "valign": "vcenter",
                    "align": "center",
                    "bold": True,
                    "font_name": "Calibri",
                    "font_size": 11,
                    "border": 1,
                }
            )
            bold_12_format = writer.book.add_format(
                {
                    "bg_color": "#FFFFFF",
                    "text_wrap": False,
                    "valign": "vcenter",
                    "align": "center",
                    "bold": True,
                    "font_name": "Calibri",
                    "font_size": 12,
                    "border": 1,
                }
            )
            bold_12_wrap_format = writer.book.add_format(
                {
                    "bg_color": "#FFFFFF",
                    "text_wrap": True,
                    "valign": "vcenter",
                    "align": "center",
                    "bold": True,
                    "font_name": "Calibri",
                    "font_size": 12,
                    "border": 1,
                }
            )
            bold_11_wrap_format = writer.book.add_format(
                {
                    "bg_color": "#FFFFFF",
                    "text_wrap": True,
                    "valign": "vcenter",
                    "align": "center",
                    "bold": True,
                    "font_name": "Calibri",
                    "font_size": 11,
                    "border": 1,
                }
            )
            normal_data_format = writer.book.add_format(
                {
                    "bg_color": "#FFFFFF",
                    "text_wrap": False,
                    "valign": "vcenter",
                    "align": "center",
                    "bold": False,
                    "font_name": "Calibri",
                    "font_size": 11,
                    "border": 1,
                }
            )
            # Create a format for a black cell
            black_cell = writer.book.add_format({"bg_color": "black"})
            if data["DM"]["Subject"].count() > 0:
                ## TODO: DSMB-Demo Stats Table
                if self.enrollment_listing_df_output["Subject"].count() > 0:
                    # * WRITING DATA: LegalSex_list, Age_at_Consent_list, Race_list, Ethnicity_list
                    worksheet1 = writer.book.add_worksheet("Demographics Statistics")

                    # * FORMAT DATA
                    for i in range(0, len(self.status_list)):
                        for j in range(0, len(self.LegalSex_list[i])):
                            for k in range(0, len(self.LegalSex_list[i].columns)):
                                worksheet1.write(
                                    j + 3,
                                    k + 1 + i * 4,
                                    self.LegalSex_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(self.Age_at_Consent_list[i])):
                            for k in range(0, len(self.Age_at_Consent_list[i].columns)):
                                worksheet1.write(
                                    j + 8,
                                    k + 1 + i * 4,
                                    self.Age_at_Consent_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(self.Race_list[i])):
                            for k in range(0, len(self.Race_list[i].columns)):
                                worksheet1.write(
                                    j + 12,
                                    k + 1 + i * 4,
                                    self.Race_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(self.Ethnicity_list[i])):
                            for k in range(0, len(self.Ethnicity_list[i].columns)):
                                worksheet1.write(
                                    j + 22,
                                    k + 1 + i * 4,
                                    self.Ethnicity_list[i].iloc[j, k],
                                    normal_data_format,
                                )

                    # * WRITING HEADER AND FORMATTING
                    Sex_order = ["Male", "Female", "X (Nonbinary)", "Not Reported"]
                    Age_order = ["Mean SD", "Median", "Range"]
                    Race_order = [
                        "African American",
                        "Alaska Native",
                        "American Indian",
                        "Asian",
                        "Caucasian",
                        "Multiple Races",
                        "Pacific Islander",
                        "Other",
                        "Unknown",
                    ]
                    Ethnicity_order = ["Hispanic", "Non-Hispanic", "Unknown"]

                    for i in range(0, len(Sex_order)):
                        worksheet1.write(i + 3, 0, Sex_order[i], bold_11_format)
                    for i in range(0, len(Age_order)):
                        worksheet1.write(i + 8, 0, Age_order[i], bold_11_format)
                    for i in range(0, len(Race_order)):
                        worksheet1.write(i + 12, 0, Race_order[i], bold_11_format)
                    for i in range(0, len(Ethnicity_order)):
                        worksheet1.write(i + 22, 0, Ethnicity_order[i], bold_11_format)

                    worksheet1.merge_range("B1:E1", "Overall Study Enrollment", bold_12_format)
                    worksheet1.merge_range("F1:I1", "Cohort 1 Enrollment", bold_12_format)
                    worksheet1.merge_range("J1:M1", "Cohort 2 Enrollment", bold_12_format)
                    worksheet1.merge_range("N1:Q1", "Cohort -1 Enrollment", bold_12_format)
                    worksheet1.write(1, 0, "Status", bold_11_format)
                    for i in range(len(self.status_list)):
                        # skip "Consented" for Cohort 1, Cohort 2 and Cohort -1
                        worksheet1.write(
                            1,
                            1 + i * 4,
                            "Consented\nN=" + str(self.status_list[i]["Total Consented"]),
                            bold_11_wrap_format,
                        )

                        worksheet1.write(
                            1,
                            2 + i * 4,
                            "Screen Fail\nN=" + str(self.status_list[i]["Screen Failed"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            3 + i * 4,
                            "Eligible\nN=" + str(self.status_list[i]["Eligible"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            4 + i * 4,
                            "Treated\nN=" + str(self.status_list[i]["Treated"]),
                            bold_11_wrap_format,
                        )

                    worksheet1.merge_range("A3:Q3", "Legal Sex", bold_11_format)
                    worksheet1.merge_range("A8:Q8", "Age at Consent", bold_11_format)
                    worksheet1.merge_range("A12:Q12", "Race", bold_11_format)
                    worksheet1.merge_range("A22:Q22", "Ethnicity", bold_11_format)
                    # # remove columns F, G, J, K, N, O
                    # worksheet1.delete_cols(15)
                    # worksheet1.delete_cols(14)
                    # worksheet1.delete_cols(10)
                    # worksheet1.delete_cols(9)
                    # worksheet1.delete_cols(7)
                    # worksheet1.delete_cols(6)
                    worksheet1.autofit()

                    ## TODO: Enrollment Listing
                    # * WRITING DATA: enrollment_listing_df_output
                    worksheet2 = writer.book.add_worksheet("Enrollment Listing")
                    # self.enrollment_df = self.enrollment_df.drop(columns=["End of Study Date"])

                    # * WRITING HEADER AND FORMATTING
                    # Assuming 'enrollment_listing_df_output' is your DataFrame
                    self.enrollment_listing_df_output = self.enrollment_listing_df_output.replace(
                        [np.inf, -np.inf], np.nan
                    )  # Replace INF with NaN

                    self.enrollment_listing_df_output = self.enrollment_listing_df_output.fillna(
                        ""
                    )  # Replace NaN with a placeholder
                    # Replace column header for A1 from "Subject" to "Subject ID", columns name starting from 1 instead of 0
                    worksheet2.write("A1", "Subject ID", bold_12_wrap_format)
                    for i in range(1, len(self.enrollment_listing_df_output.columns)):
                        worksheet2.write(0, i, self.enrollment_listing_df_output.columns[i], bold_11_format)
                    # * FORMAT DATA
                    for i in range(0, len(self.enrollment_listing_df_output)):
                        for j in range(0, len(self.enrollment_listing_df_output.columns)):
                            worksheet2.write(i + 1, j, self.enrollment_listing_df_output.iloc[i, j], normal_data_format)
                    # Autofit
                    worksheet2.autofit()

                    ## TODO: EGFR for Eligible Subjects
                    # worksheet3 = writer.book.add_worksheet("EGFR for Eligible Subjects")

                    # # * WRITING HEADER AND FORMATTING
                    # # Assuming 'EGFR_listing_df_output' is your DataFrame
                    # self.EGFR_listing_df_output = self.EGFR_listing_df_output.replace(
                    #     [np.inf, -np.inf], np.nan
                    # )  # Replace INF with NaN

                    # self.EGFR_listing_df_output = self.EGFR_listing_df_output.fillna(
                    #     ""
                    # )  # Replace NaN with a placeholder
                    # # Replace column header for A1 from "Subject" to "Subject ID", columns name starting from 1 instead of 0
                    # worksheet3.write("A1", "Subject ID", bold_12_wrap_format)
                    # for i in range(1, len(self.EGFR_listing_df_output.columns)):
                    #     worksheet3.write(0, i, self.EGFR_listing_df_output.columns[i], bold_11_format)
                    # # * FORMAT DATA
                    # for i in range(0, len(self.EGFR_listing_df_output)):
                    #     for j in range(0, len(self.EGFR_listing_df_output.columns)):
                    #         worksheet3.write(i + 1, j, self.EGFR_listing_df_output.iloc[i, j], normal_data_format)
                    # # Autofit
                    # worksheet3.autofit()

                    # ## TODO: Study Tx Statistics
                    # worksheet4 = writer.book.add_worksheet("Study Tx Statistics")

                    # # * FORMATING DATA
                    # # Cohort 1
                    # for i in range(0, len(self.infusion_statA)):
                    #     for j in range(0, len(self.infusion_statA.columns)):
                    #         worksheet4.write(
                    #             i + 4,
                    #             j + 1,
                    #             self.infusion_statA.iloc[i, j],
                    #             normal_data_format,
                    #         )
                    # # Cohort 2
                    # for i in range(0, len(self.infusion_statB)):
                    #     for j in range(0, len(self.infusion_statB.columns)):
                    #         worksheet4.write(
                    #             i + 8,
                    #             j + 1,
                    #             self.infusion_statB.iloc[i, j],
                    #             normal_data_format,
                    #         )
                    # # Cohort -1
                    # for i in range(0, len(self.infusion_statC)):
                    #     for j in range(0, len(self.infusion_statC.columns)):
                    #         worksheet4.write(
                    #             i + 12,
                    #             j + 1,
                    #             self.infusion_statC.iloc[i, j],
                    #             normal_data_format,
                    #         )

                    # # * WRITING HEADER AND FORMATTING
                    # stat_order = ["Mean SD", "Median", "Range"]
                    # worksheet4.merge_range(
                    #     "B1:G1",
                    #     "Study Treatment Statistics (N="
                    #     + str(self.infusion_count[0] + self.infusion_count[1] + self.infusion_count[2])
                    #     + ")",
                    #     bold_12_wrap_format,
                    # )

                    # worksheet4.merge_range("B2:D2", "Cells Infused", bold_12_wrap_format)
                    # worksheet4.merge_range("E2:G2", "Transduction Efficiency", bold_12_wrap_format)
                    # worksheet4.write("B3", "Total Cell Dose", bold_12_wrap_format)
                    # worksheet4.write("C3", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                    # worksheet4.write("D3", "Met Target Dose", bold_12_wrap_format)
                    # worksheet4.write("E3", "%scFV (EGFR)", bold_12_wrap_format)
                    # worksheet4.write("F3", "Met Target % scFV Flow(Y/N) (≥2%)", bold_12_wrap_format)
                    # worksheet4.write("G3", "%scFV (Il13Rα2)", bold_12_wrap_format)
                    # worksheet4.merge_range(
                    #     "A4:G4",
                    #     "Cohort 1 (N=" + str(self.infusion_count[0]) + ")",
                    #     bold_12_wrap_format,
                    # )
                    # worksheet4.merge_range(
                    #     "A8:G8",
                    #     "Cohort 2 (N=" + str(self.infusion_count[1]) + ")",
                    #     bold_12_wrap_format,
                    # )
                    # worksheet4.merge_range(
                    #     "A12:G12",
                    #     "Cohort -1 (N=" + str(self.infusion_count[2]) + ")",
                    #     bold_12_wrap_format,
                    # )

                    # # Merge and format data
                    # worksheet4.merge_range("D5:D7", self.infusion_statA.iloc[0, 2], normal_data_format)
                    # worksheet4.merge_range("F5:F7", self.infusion_statA.iloc[0, 4], normal_data_format)
                    # worksheet4.merge_range("D9:D11", self.infusion_statB.iloc[0, 2], normal_data_format)
                    # worksheet4.merge_range("F9:F11", self.infusion_statB.iloc[0, 4], normal_data_format)
                    # worksheet4.merge_range("D13:D15", self.infusion_statC.iloc[0, 2], normal_data_format)
                    # worksheet4.merge_range("F13:F15", self.infusion_statC.iloc[0, 4], normal_data_format)

                    # for i in range(0, len(stat_order)):
                    #     worksheet4.write(i + 4, 0, stat_order[i], bold_11_format)  # Cohort 1
                    # for i in range(0, len(stat_order)):
                    #     worksheet4.write(i + 8, 0, stat_order[i], bold_11_format)  # Cohort 2
                    # for i in range(0, len(stat_order)):
                    #     worksheet4.write(i + 12, 0, stat_order[i], bold_11_format)  # Cohort -1

                    # # * Autofit
                    # worksheet4.autofit()

                    # ## TODO: DSMB-Infusion Listing
                    # worksheet5 = writer.book.add_worksheet("Study Tx Listing")
                    # # * WRITING AND FORMATING DATA
                    # for i in range(0, len(self.infusion_df)):
                    #     for j in range(0, len(self.infusion_df.columns)):
                    #         worksheet5.write(i + 2, j, self.infusion_df.iloc[i, j], normal_data_format)
                    # # if there are subjects in infusionR_df
                    # if len(self.infusionR_df) > 0:
                    #     for i in range(0, len(self.infusionR_df)):
                    #         for j in range(0, len(self.infusionR_df.columns)):
                    #             worksheet5.write(
                    #                 i + 2,
                    #                 j + 15,
                    #                 self.infusionR_df.iloc[i, j],
                    #                 normal_data_format,
                    #             )
                    # # * WRITING HEADER AND FORMATTING
                    # worksheet5.merge_range("A1:A2", "Subject ID", bold_12_wrap_format)
                    # worksheet5.merge_range("B1:B2", "Study Day (Primary)", bold_12_wrap_format)
                    # worksheet5.merge_range("C1:C2", "Cohort", bold_12_wrap_format)
                    # worksheet5.merge_range("D1:D2", "Target Dose", bold_12_wrap_format)
                    # worksheet5.merge_range(
                    #     "E1:E2",
                    #     "Study Treatment Date",
                    #     bold_12_wrap_format,
                    # )
                    # worksheet5.merge_range("F1:G1", "Product Administration Volumes", bold_12_wrap_format)
                    # worksheet5.merge_range("H1:J1", "Cells Administered", bold_12_wrap_format)
                    # worksheet5.merge_range("K1:M1", "Transduction Efficiency", bold_12_wrap_format)
                    # worksheet5.write("F2", "Volume CSF Withdrawn", bold_12_wrap_format)
                    # worksheet5.write(
                    #     "G2",
                    #     "Volume Dose Administered",
                    #     bold_12_wrap_format,
                    # )
                    # worksheet5.write("H2", "Total Cell Dose", bold_12_wrap_format)
                    # worksheet5.write("I2", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                    # worksheet5.write("J2", "Met Target Dose (Y/N)", bold_12_wrap_format)
                    # worksheet5.write("K2", "%scFV (EGFR)", bold_12_wrap_format)
                    # worksheet5.write("L2", "Met Target % scFV Flow (Y/N) (≥2%)", bold_12_wrap_format)
                    # worksheet5.write("M2", "%scFV (Il13Rα2)", bold_12_wrap_format)
                    # if len(self.infusionR_df) > 0:
                    #     worksheet5.merge_range("P1:P2", "Subject ID", bold_12_wrap_format)
                    #     worksheet5.merge_range("Q1:Q2", "Study Day (Retreatment)", bold_12_wrap_format)
                    #     worksheet5.merge_range("R1:R2", "Cohort", bold_12_wrap_format)
                    #     worksheet5.merge_range(
                    #         "S1:S2",
                    #         "Study Treatment Date",
                    #         bold_12_wrap_format,
                    #     )

                    #     worksheet5.merge_range("T1:U1", "Product Administration Volumes", bold_12_wrap_format)
                    #     worksheet5.merge_range("V1:W1", "Cells Administered", bold_12_wrap_format)
                    #     worksheet5.merge_range("X1:Z1", "Transduction Efficiency", bold_12_wrap_format)
                    #     worksheet5.write(
                    #         "T2",
                    #         "Volume CSF Withdrawn",
                    #         bold_12_wrap_format,
                    #     )
                    #     worksheet5.write("U2", "Volume Dose Administered", bold_12_wrap_format)
                    #     worksheet5.write("V2", "Total Cell Dose", bold_12_wrap_format)
                    #     worksheet5.write("W2", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                    #     worksheet5.write("X2", "%scFV (EGFR)", bold_12_wrap_format)
                    #     worksheet5.write("Y2", "Met Target % scFV Flow (Y/N) (≥2%)", bold_12_wrap_format)
                    #     worksheet5.write("Z2", "%scFV (Il13Rα2)", bold_12_wrap_format)

                    # # Autofit
                    # worksheet5.autofit()

                    # worksheet6 = writer.book.add_worksheet("Status for Eligible Subjects")
                    # # * WRITING AND FORMATING DATA
                    # for i in range(0, len(self.status_df)):
                    #     for j in range(0, len(self.status_df.columns)):
                    #         worksheet6.write(i + 2, j, self.status_df.iloc[i, j], normal_data_format)

                    # # * WRITING HEADER AND FORMATTING
                    # worksheet6.merge_range("A1:A2", "Subject ID", bold_12_wrap_format)
                    # worksheet6.merge_range("B1:B2", "Cohort", bold_12_wrap_format)
                    # worksheet6.merge_range("C1:C2", "Adverse Events (Y/N)", bold_12_wrap_format)
                    # worksheet6.merge_range("D1:D2", "Serious Adverse Events (Y/N)", bold_12_wrap_format)
                    # worksheet6.merge_range("E1:E2", "Study Status", bold_12_wrap_format)
                    # worksheet6.merge_range("F1:F2", "Off-Study Reason", bold_12_wrap_format)
                    # worksheet6.merge_range(
                    #     "G1:G2",
                    #     "Last Study Visit Performed for Off-Study Subject",
                    #     bold_12_wrap_format,
                    # )

                    # # Safety Headers
                    # # number of subject of safety_total_df
                    # safety_total_df_subject_count = len(self.status_df["Subject"].unique())
                    # worksheet6.merge_range(
                    #     "K1:N1",
                    #     "Safety Statistics (N=" + str(safety_total_df_subject_count) + ")",
                    #     bold_12_wrap_format,
                    # )
                    # worksheet6.merge_range("K2:L2", "Adverse Events", bold_11_format)
                    # worksheet6.merge_range("M2:N2", "Serious Adverse Events ", bold_11_format)
                    # worksheet6.write("K3", "Yes", bold_11_format)
                    # worksheet6.write("L3", "No", bold_11_format)
                    # worksheet6.write("M3", "Yes", bold_11_format)
                    # worksheet6.write("N3", "No", bold_11_format)
                    # worksheet6.write("J4", "Cohort 1", bold_11_format)
                    # worksheet6.write("J5", "Cohort 2", bold_11_format)
                    # worksheet6.write("J6", "Cohort -1", bold_11_format)

                    # # Safety Data
                    # # Cohort 1
                    # for i in range(0, len(self.safetyCH1_total_df)):
                    #     for j in range(0, len(self.safetyCH1_total_df.columns)):
                    #         worksheet6.write(
                    #             i + 3,
                    #             j + 10,
                    #             self.safetyCH1_total_df.iloc[i, j],
                    #             normal_data_format,
                    #         )
                    # # Cohort 2
                    # for i in range(0, len(self.safetyCH2_total_df)):
                    #     for j in range(0, len(self.safetyCH2_total_df.columns)):
                    #         worksheet6.write(
                    #             i + 4,
                    #             j + 10,
                    #             self.safetyCH2_total_df.iloc[i, j],
                    #             normal_data_format,
                    #         )

                    # # Cohort -1
                    # for i in range(0, len(self.safetyCHN1_total_df)):
                    #     for j in range(0, len(self.safetyCHN1_total_df.columns)):
                    #         worksheet6.write(
                    #             i + 5,
                    #             j + 10,
                    #             self.safetyCHN1_total_df.iloc[i, j],
                    #             normal_data_format,
                    #         )

                    # # Autofit
                    # worksheet6.autofit()
