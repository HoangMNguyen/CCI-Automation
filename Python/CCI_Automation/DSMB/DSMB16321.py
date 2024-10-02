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

# Opt-in to the future behavior
pd.set_option("future.no_silent_downcasting", True)


class DSMB16321:
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
        # process the enrollment listing
        self.enrollment_listing_df_output, self.enrollment_listing_df = self.enrollment_listing()
        self.enrollment_stat_table(self.enrollment_listing_df)
        self.EGFR_listing_df_output, self.EGFR_listing_df = self.EGFR_listing()
        self.infusion_df, self.infusionR_df = self.infusion_listing()
        self.infusion_stats(self.infusion_df, self.infusionR_df)
        self.response_df, self.responseR_df = self.response_listing()
        self.response_stats()
        self.AE_df, self.status_df, self.safetyCH1_total_df, self.safetyCH2_total_df, self.safetyCHN1_total_df = (
            self.status_listing(self.enrollment_listing_df)
        )
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
        enrollment_df["Disease Type"] = enrollment_df["Disease"].fillna("") + " " + enrollment_df["Disease"].fillna("")
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
        # If Race is "Other", replace the value with "Other Race"
        enrollment_output_df.loc[
            enrollment_output_df["Race"] == "Other",
            "Race",
        ] = enrollment_output_df["Other Race"]
        # *Re-order the columns and remove the columns that are not needed
        enrollment_output_df = enrollment_output_df[
            [
                "Subject",
                "Cohort",
                "Disease Type",
                "Legal Sex",
                "Sex Assigned at Birth",
                "Gender Identity",
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

    def EGFR_listing(self):
        data = self.data
        # Find eligible subjects
        IE_df = data["IE"][["Subject", "Subject Meets All Study Eligibility (ig_IE3.IEYN)"]].copy()
        IE_new_col_name = {
            "Subject Meets All Study Eligibility (ig_IE3.IEYN)": "Subject meets all study eligibility?",
        }
        IE_df = IE_df.rename(columns=IE_new_col_name)
        Eligible_df = IE_df[IE_df["Subject meets all study eligibility?"] == "Yes"].copy()
        # *: PREPARE DATA FOR EGFR LISTING
        EGFR_subject_df = Eligible_df["Subject"].copy()
        EGFR1_df = EGFR_subject_df
        EGFR1_df = add_rename_column_corelisting(EGFR1_df, data, "LBEGFR", "Event Group Label", "Event Group Label")
        EGFR1_df = EGFR1_df[(EGFR1_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")].copy()

        EGFR_collectionDT_df = data["LBEGFR"][
            [
                "Subject",
                "Event Group Label",
                "Was EGFR Amplification testing performed? (IG_NS_NA_LBEGFR1.CL_NS_NH_EGFRAMPERF_cl_YS_YN1)",
                "Collection Date (IG_NS_NA_LBEGFR1.DT_YS_NH_LBDAT)",
                "Amplification of EGFR (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR_cl_YS_DTNDT1)",
                "EGFRvIII Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR8_cl_YS_DTNDT1)",
                "EGFR Extracellular Domain Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_EGFRMUT_cl_YS_DTNDT1)",
            ]
        ].copy()
        EGFR_new_col_name = {
            "Was EGFR Amplification testing performed? (IG_NS_NA_LBEGFR1.CL_NS_NH_EGFRAMPERF_cl_YS_YN1)": "Was EGFR Amplification testing performed?",
            "Collection Date (IG_NS_NA_LBEGFR1.DT_YS_NH_LBDAT)": "Collection Date",
            "Amplification of EGFR (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR_cl_YS_DTNDT1)": "Amplification of EGFR",
            "EGFRvIII Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_AMPEGFR8_cl_YS_DTNDT1)": "EGFRvIII Mutation",
            "EGFR Extracellular Domain Mutation (IG_NS_NA_LBEGFR2.CL_NS_NH_EGFRMUT_cl_YS_DTNDT1)": "EGFR Extracellular Domain Mutation",
        }
        EGFR_collectionDT_df = EGFR_collectionDT_df.rename(columns=EGFR_new_col_name)

        # Filter the DataFrame based on the conditions
        filtered_df = EGFR_collectionDT_df[
            (EGFR_collectionDT_df["Event Group Label"] == "Initial Study Enrollment/Apheresis")
            & (EGFR_collectionDT_df["Was EGFR Amplification testing performed?"] == "Yes")
            & (EGFR_collectionDT_df["Collection Date"].notna())
        ].copy()
        EGFR_collectionDT2_df = filtered_df
        EGFR_collectionDT2_df = EGFR_collectionDT2_df.drop(
            columns=[
                "Event Group Label",
                "Was EGFR Amplification testing performed?",
            ]
        )
        EGFR2_df = pd.merge(
            EGFR1_df,
            EGFR_collectionDT2_df,
            on=["Subject"],
            how="left",
        ).drop_duplicates()

        # Sort and get the last row for each subject
        EGFR2_df = EGFR2_df.sort_values(["Collection Date"])
        EGFR_df = EGFR2_df.groupby("Subject").tail(2)  # same collection date will have 2 rows with different labs
        EGFR_df = EGFR_df.sort_values(["Subject"])
        # print(EGFR_df)

        replacement_date = pd.Timestamp("1900-01-01")

        # Replace missing dates with the specified value
        EGFR_df["Collection Date"] = EGFR_df["Collection Date"].fillna(replacement_date)
        EGFR_df["Collection Date"] = pd.to_datetime(EGFR_df["Collection Date"])

        # List of columns to concatenate
        columns_to_concatenate = [
            "Amplification of EGFR",
            "EGFRvIII Mutation",
            "EGFR Extracellular Domain Mutation",
        ]

        # Group by "Collection Date" and concatenate strings within each group, when some tests are performed twice by different labs, there will be duplicate results in one cell
        # Perform the groupby operation and aggregation
        agg_df = (
            EGFR_df.groupby(["Collection Date"])[columns_to_concatenate]
            .apply(lambda group: group.fillna("").astype(str).agg(" ".join))
            .reset_index()
            .drop_duplicates()
        )

        # Rename the aggregated columns if needed
        agg_df.columns = ["Collection Date"] + columns_to_concatenate

        # # Replace missing dates with the specified value
        agg_df["Collection Date"] = agg_df["Collection Date"].fillna(replacement_date)

        # Merge aggregated data back to the original DataFrame
        EGFR_df = pd.merge(
            EGFR_df[["Subject", "Collection Date"]].drop_duplicates(),
            agg_df,
            on="Collection Date",
            how="left",
        )

        EGFR_final_df = EGFR_df.sort_values(["Subject"])

        # get "MGMT Result" data from MHDIAG
        MHDIAG_df = data["MHDIAG"][["Subject", "MGMT Result (IG_NS_NA_MHDIAG2.CL_NS_NH_MGMTRES_cl_NS_MGMTRES1)"]].copy()
        MHDIAG_new_col_name = {
            "MGMT Result (IG_NS_NA_MHDIAG2.CL_NS_NH_MGMTRES_cl_NS_MGMTRES1)": "MGMT Result",
        }
        MHDIAG_df = MHDIAG_df.rename(columns=MHDIAG_new_col_name)
        MHDIAG_df = pd.merge(
            Eligible_df,
            MHDIAG_df,
            on=["Subject"],
            how="left",
        ).drop_duplicates()
        # Merge EGFR data with MHDIAG data, keep eligible subjects even EGFR was not tested
        EGFR_final2_df = pd.merge(EGFR_final_df, MHDIAG_df, on="Subject", how="right")
        EGFR_final2_df["Collection Date"] = EGFR_final2_df["Collection Date"].fillna(replacement_date)
        EGFR_final2_df = EGFR_final2_df.sort_values(["Subject"])
        EGFR_final2_df = EGFR_final2_df.fillna("")

        # prepare the output dataframe
        EGFR_output_df = EGFR_final2_df.copy()
        # *Re-order the columns and remove the columns that are not needed
        EGFR_output_df = EGFR_output_df[
            [
                "Subject",
                #   "Collection Date",
                #  "Laboratory Name",
                "Amplification of EGFR",
                "EGFRvIII Mutation",
                "EGFR Extracellular Domain Mutation",
                "MGMT Result",
            ]
        ]
        return EGFR_output_df, EGFR_final2_df

    def infusion_listing(self):
        data = self.data
        ### TODO: INFUSION LISTING
        # adding Target Cell Dose dictionary
        # !: Update this dictionary to the new study
        TCD_dict = {
            "Cohort -1": 5000000,
            "Cohort 1": 10000000,
            "Cohort 2": 25000000,
            "Cohort 3": 50000000,
            "Not Assigned": "Not Assigned",
        }

        # *: PREPARE DATA FOR INFUSION LISTING

        # create dictionary for enrollment listing
        input_dict1 = {
            "EXINF": {
                "Event Group Label": "Event Group Label",
                "Study Treatment Date (ig_EXINF1.INFDAT)": "Study Treatment Date",
                "Volume CSF Removed for Cell Product Administration (mL) (ig_EXINF1.CSFVOL)": "Volume CSF Withdrawn",
                "Total Volume Administered (mL) (ig_EXINF1.INFTOTVOL)": "Volume Dose Administered",
                "CAR T Cell Dose Administered (ig_EXINF1.INFDOS)": "CART-EGFR-IL13Rα2 Cell Dose",
                "x 10 to the power of (ig_EXINF1.INFDOSXP)": "x 10 to the power of (ig_EXINF1.INFDOSXP)",
                "Total Cell Dose Administered (ig_EXINF1.INFDOSTOT)": "Total Cell Dose",
                "x 10 to the power of (ig_EXINF1.INFDOSTOTXP)": "x 10 to the power of (ig_EXINF1.INFDOSTOTXP)",
                "EGFR Transduction Efficiency (%) (ig_EXINF1.EGFRINFTEFFP)": "%scFV (EGFR)",
                "IL13Ra2 Transduction Efficiency (%) (ig_EXINF1.IL13INFTEFFP)": "%scFV (Il13Rα2)",
                # "Event Date": "Event Date INF",
            },
        }
        input_dict2 = {
            "DSCA": {"Cohort Assignment (ig_DSCA1.CACHASCOD)": "Cohort"},
        }

        # For infusion CRF, get the earliest data (Day 0)
        raw_infusion_df1 = get_data_from_dict_first(data, input_dict1)
        raw_infusion_df2 = get_data_from_dict(data, input_dict2)
        raw_infusion_df = pd.merge(raw_infusion_df1, raw_infusion_df2, on="Subject", how="left")
        # convert the date to datetime object and format it to MM-DD-YYYY
        raw_infusion_df["Study Treatment Date"] = raw_infusion_df["Study Treatment Date"].apply(
            lambda x: datetime.strptime(x.strftime("%Y-%m-%d"), "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
        )

        # print(raw_infusion_df)
        # TODO: INFUSION LISTING Day 0

        infusion_df = raw_infusion_df[raw_infusion_df["Event Group Label"] == "Day 0"]
        # print(infusion_df)

        # adding Target Dose using TCD_dict
        infusion_df["Target Dose"] = infusion_df["Cohort"].map(TCD_dict)

        # combine CART-EGFR-IL13Rα2 Cell Dose and x 10 to the power of (ig_EXINF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the CART-EGFR-IL13Rα2 Cell Dose column to string
        infusion_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusion_df["CART-EGFR-IL13Rα2 Cell Dose"].multiply(
            10 ** infusion_df["x 10 to the power of (ig_EXINF1.INFDOSXP)"]
        )
        infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSXP)"])
        infusion_df["Total Cell Dose"] = infusion_df["Total Cell Dose"].multiply(
            10 ** infusion_df["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"]
        )
        infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"])

        # Adding Met Target Dose column based on the condition of Total Cell Dose and CART-EGFR-IL13Rα2 Cell Dose if 'Target Cell Dose' is integer
        infusion_df["Met Target Dose (Y/N)"] = infusion_df.apply(
            lambda row: "Y"
            if isinstance(row["Target Dose"], int) and row["CART-EGFR-IL13Rα2 Cell Dose"] >= row["Target Dose"]
            else "",
            axis=1,
        )
        infusion_df["Met Target Dose (Y/N)"] = infusion_df.apply(
            lambda row: "N"
            if isinstance(row["Target Dose"], int) and row["CART-EGFR-IL13Rα2 Cell Dose"] < row["Target Dose"]
            else row["Met Target Dose (Y/N)"],
            axis=1,
        )

        # adding Met Target %scFv and fillter out the rows that have NaN in Met Target %scFv
        infusion_df["Met Target % scFV Flow (Y/N) (≥2%)"] = infusion_df[infusion_df["%scFV (EGFR)"].notna()][
            "%scFV (EGFR)"
        ].apply(lambda x: "Y" if x >= 2 else "N")
        # fill NaN with empty string
        infusion_df = infusion_df.fillna(
            "",
        ).infer_objects(copy=False)

        # Only keep the rows that have Event Group Label
        infusion_df = infusion_df[infusion_df["Event Group Label"] != ""]

        # *Re-order the columns and remove the columns that are not needed
        infusion_df = infusion_df[
            [
                "Subject",
                "Event Group Label",
                "Cohort",
                "Target Dose",
                "Study Treatment Date",
                "Volume CSF Withdrawn",
                "Volume Dose Administered",
                "Total Cell Dose",
                "CART-EGFR-IL13Rα2 Cell Dose",
                "Met Target Dose (Y/N)",
                "%scFV (EGFR)",
                "Met Target % scFV Flow (Y/N) (≥2%)",
                "%scFV (Il13Rα2)",
            ]
        ]

        # TODO: Infusion Listing Day 0-R1, Day 0-R2
        # Get the latest data to include Day 0-R1 and Day 0-R2 data
        raw_infusionR_df1 = get_data_from_dict(data, input_dict1)
        raw_infusionR_df2 = get_data_from_dict(data, input_dict2)
        raw_infusionR_df = pd.merge(raw_infusionR_df1, raw_infusionR_df2, on="Subject", how="left")
        # convert the date to datetime object and format it to MM-DD-YYYY
        raw_infusionR_df["Study Treatment Date"] = raw_infusionR_df["Study Treatment Date"].apply(
            lambda x: datetime.strptime(x.strftime("%Y-%m-%d"), "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
        )
        infusionR_df = raw_infusionR_df[
            (raw_infusionR_df["Event Group Label"] == "Day 0-R1")
            | (raw_infusionR_df["Event Group Label"] == "Day 0-R2")
        ]

        # combine CART-EGFR-IL13Rα2 Cell Dose and x 10 to the power of (ig_EXINF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the CART-EGFR-IL13Rα2 Cell Dose column to string
        infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"].multiply(
            10 ** infusionR_df["x 10 to the power of (ig_EXINF1.INFDOSXP)"]
        )
        infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSXP)"])

        infusionR_df["Total Cell Dose"] = infusionR_df["Total Cell Dose"].multiply(
            10 ** infusionR_df["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"]
        )
        infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_EXINF1.INFDOSTOTXP)"])

        # adding Met Target %scFv
        infusionR_df = add_rename_column_corelisting(
            infusionR_df,
            data,
            "EXINF",
            "IL13Ra2 Transduction Efficiency (%) (ig_EXINF1.IL13INFTEFFP)",
            "Met Target % scFV Flow (Y/N) (≥2%)",
            "Subject",
            "Event Group Label",
        )
        # adding Met Target %scFv and fillter out the rows that have NaN in Met Target %scFv
        infusionR_df["Met Target % scFV Flow (Y/N) (≥2%)"] = infusionR_df[
            infusionR_df["Met Target % scFV Flow (Y/N) (≥2%)"].notna()
        ]["Met Target % scFV Flow (Y/N) (≥2%)"].apply(lambda x: "Y" if x >= 2 else "N")
        # fill NaN with empty string
        infusionR_df = infusionR_df.fillna("")

        # Only keep the rows that have Event Group Label
        infusionR_df = infusionR_df[infusionR_df["Event Group Label"] != ""]

        # *Re-order the columns and remove the columns that are not needed
        infusionR_df = infusionR_df[
            [
                "Subject",
                "Event Group Label",
                "Cohort",
                "Study Treatment Date",
                "Volume CSF Withdrawn",
                "Volume Dose Administered",
                "Total Cell Dose",
                "CART-EGFR-IL13Rα2 Cell Dose",
                "%scFV (EGFR)",
                "Met Target % scFV Flow (Y/N) (≥2%)",
                "%scFV (Il13Rα2)",
            ]
        ]

        return infusion_df, infusionR_df

    def infusion_stats(self, infusion_df, infusionR_df):
        # TODO: INFUSION STATISTICS
        infusion_count = []
        # * Cohort 1
        # Create a new dataframe for Cohort 1 with infusion_df
        infusion1_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort 1"]
        # Create a new dataframe for Total Cell Dose table with infusion_df
        infusion_statA1 = get_stats_df("Total Cell Dose", infusion1_df)

        infusion_statA2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion1_df)

        # Count the number of subjects that met the target dose
        met_targetDose_count = infusion1_df[infusion1_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
        # Count the number of subjects
        total_subject_count = infusion1_df["Subject"].nunique()
        # Add column "Met Target Dose (Y/N)" to infusion_statA2 data frame, since data frames are concated, the order of the data frame should match the output order
        infusion_statA2["Met Target Dose (Y/N)"] = (
            str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFv Flow table with infusion_df
        infusion_statA3 = get_stats_perc_df("%scFV (EGFR)", infusion1_df)
        # Count the number of subjects that met the target %scFv
        met_targetFlow_count = infusion1_df[infusion1_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
            "Subject"
        ]
        infusion_statA3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
            str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
        infusion_statA4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion1_df)

        # Combine the three dataframes
        infusion_statA = pd.concat(
            [infusion_statA1, infusion_statA2, infusion_statA3, infusion_statA4],
            axis=1,
        )
        infusion_statA = infusion_statA.replace([np.inf, -np.inf], "")
        infusion_statA = infusion_statA.fillna("")
        self.infusion_statA = infusion_statA
        infusion_count.append(total_subject_count)
        self.infusion_count = infusion_count

        # * Cohort 2
        # Create a new dataframe for Cohort 2 with infusion_df
        infusion2_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort 2"]
        infusion_statB1 = get_stats_df("Total Cell Dose", infusion2_df)

        infusion_statB2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion2_df)

        # Count the number of subjects that met the target dose
        # infusion_statB3 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion2_df)
        met_targetDose_count = infusion2_df[infusion2_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
        # Count the number of subjects
        total_subject_count = infusion2_df["Subject"].nunique()
        infusion_statB2["Met Target Dose (Y/N)"] = (
            str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFv Flow table with infusion_df
        infusion_statB3 = get_stats_perc_df("%scFV (EGFR)", infusion2_df)
        # Count the number of subjects that met the target %scFv
        met_targetFlow_count = infusion2_df[infusion2_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
            "Subject"
        ]
        infusion_statB3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
            str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
        infusion_statB4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion2_df)

        # Combine the three dataframes
        infusion_statB = pd.concat([infusion_statB1, infusion_statB2, infusion_statB3, infusion_statB4], axis=1)
        infusion_statB = infusion_statB.replace([np.inf, -np.inf], "")
        infusion_statB = infusion_statB.fillna("")
        self.infusion_statB = infusion_statB
        infusion_count.append(total_subject_count)
        self.infusion_count = infusion_count

        # * Cohort -1
        # Create a new dataframe for Cohort 2 with infusion_df
        infusion3_df = self.infusion_df[self.infusion_df["Cohort"] == "Cohort -1"]
        infusion_statC1 = get_stats_df("Total Cell Dose", infusion3_df)

        infusion_statC2 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion3_df)

        # Count the number of subjects that met the target dose
        # infusion_statC3 = get_stats_df("CART-EGFR-IL13Rα2 Cell Dose", infusion3_df)
        met_targetDose_count = infusion3_df[infusion3_df["Met Target Dose (Y/N)"] == "Y"].count()["Subject"]
        # Count the number of subjects
        total_subject_count = infusion3_df["Subject"].nunique()
        infusion_statC2["Met Target Dose (Y/N)"] = (
            str(met_targetDose_count) + " (" + str(round(met_targetDose_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFv Flow table with infusion_df
        infusion_statC3 = get_stats_perc_df("%scFV (EGFR)", infusion3_df)
        # Count the number of subjects that met the target %scFv
        met_targetFlow_count = infusion3_df[infusion3_df["Met Target % scFV Flow (Y/N) (≥2%)"] == "Y"].count()[
            "Subject"
        ]
        infusion_statC3["Met Target % scFV Flow (Y/N) (≥2%)"] = (
            str(met_targetFlow_count) + " (" + str(round(met_targetFlow_count / total_subject_count * 100, 2)) + "%)"
        )

        # Create a new dataframe for %scFV (Il13Rα2) with infusion_df
        infusion_statC4 = get_stats_perc_df("%scFV (Il13Rα2)", infusion3_df)
        # Combine the three dataframes
        infusion_statC = pd.concat([infusion_statC1, infusion_statC2, infusion_statC3, infusion_statC4], axis=1)
        infusion_statC = infusion_statC.replace([np.inf, -np.inf], "")
        infusion_statC = infusion_statC.fillna("")
        self.infusion_statC = infusion_statC
        infusion_count.append(total_subject_count)
        self.infusion_count = infusion_count

        ## TODO: FORMATTING THE DATAFRAME
        # TODO: Day 0
        # Convert the columns to scientific notation if the value is not NaN
        infusion_df["Target Dose"] = infusion_df["Target Dose"].apply(
            lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
        )
        infusion_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusion_df["CART-EGFR-IL13Rα2 Cell Dose"].apply(
            lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
        )
        infusion_df["Total Cell Dose"] = infusion_df["Total Cell Dose"].apply(
            lambda x: convert_float_2_sci_notation(x) if not isinstance(x, str) and pd.notna(x) else x
        )
        # adding '%' sign to %scFv Flow
        infusion_df["%scFV (EGFR)"] = infusion_df.apply(
            lambda row: str(x) + "%" if pd.notna(x := row["%scFV (EGFR)"]) else x, axis=1
        )
        infusion_df["%scFV (Il13Rα2)"] = infusion_df.apply(
            lambda row: str(x) + "%" if pd.notna(x := row["%scFV (Il13Rα2)"]) else x, axis=1
        )

        # TODO: Day 0-R1
        # Convert the columns to scientific notation if the value is not NaN
        infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"] = infusionR_df["CART-EGFR-IL13Rα2 Cell Dose"].apply(
            lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
        )
        infusionR_df["Total Cell Dose"] = infusionR_df["Total Cell Dose"].apply(
            lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
        )
        # adding '%' sign to %scFv Flow
        infusionR_df["%scFV (EGFR)"] = infusionR_df.apply(
            lambda row: str(x) + "%" if pd.notna(x := row["%scFV (EGFR)"]) else x, axis=1
        )
        infusionR_df["%scFV (Il13Rα2)"] = infusionR_df.apply(
            lambda row: str(x) + "%" if pd.notna(x := row["%scFV (Il13Rα2)"]) else x, axis=1
        )

    def response_listing(self):
        data = self.data

        # TODO: RESPONSE LISTING
        INF_df = data["EXINF"][
            ["Subject", "Event Group Label", "Was study treatment administered? (ig_EXINF1.INFOCCUR)"]
        ].copy()
        INF_new_col_name = {
            "Was study treatment administered? (ig_EXINF1.INFOCCUR)": "Treated",
        }
        INF_df = INF_df.rename(columns=INF_new_col_name)
        Treated_df = INF_df[(INF_df["Event Group Label"] == "Day 0") & (INF_df["Treated"] == "Yes")].copy()
        Treated_df = Treated_df.drop(columns="Event Group Label")

        Treated_subject_df = Treated_df["Subject"].copy()

        response_df = Treated_subject_df

        response_df = add_rename_column_corelisting(
            response_df,
            data,
            "RS",
            "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
            "Measurable vs. Non-Measurable Disease",
            "Subject",
            #  "Event Group Label",
        ).drop_duplicates()

        # Drop rows where 'Measurable vs. Non-Measurable Disease' is NaN or empty string
        response_df = response_df[
            response_df["Measurable vs. Non-Measurable Disease"].notna()
            & response_df["Measurable vs. Non-Measurable Disease"].astype(bool)
        ]

        # Corrected lambda function to replace values based on condition and handle missing values
        response_df["Measurable vs. Non-Measurable Disease"] = response_df.apply(
            lambda row: "Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "Yes"
            else "Non-Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "No"
            else "Unknown",
            axis=1,
        )

        responseP_RSBMRI_df = data["RSBMRI"][
            [
                "Subject",
                "Event Group Label",
                "Event Date",
                "Study Phase (IG_NS_NA_RSBMRI1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)",
                # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)",
                # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)",
                "Lesion # (IG_NS_NA_RSBMRI3.NM_YS_YH_LESNUM)",
                "Measurable Lesion Change Percentage from Baseline (%) (IG_NS_NA_RSBMRI4.NM_NS_YH_TLCP)",
            ]
        ].copy()
        RSBMRI_new_col_name = {
            "Study Phase (IG_NS_NA_RSBMRI1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)": "Study Phase",
            # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)": "Primary Time Point",
            # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)": "Unscheduled Primary Day#",
            "Lesion # (IG_NS_NA_RSBMRI3.NM_YS_YH_LESNUM)": "Lesion #",
            "Measurable Lesion Change Percentage from Baseline (%) (IG_NS_NA_RSBMRI4.NM_NS_YH_TLCP)": "% Change is SPD",
        }
        responseP_RSBMRI_df = responseP_RSBMRI_df.rename(columns=RSBMRI_new_col_name).drop_duplicates()
        #
        # Convert 'Lesion #' to numeric (if needed) and handle errors by coercing non-numeric values to NaN
        responseP_RSBMRI_df["Lesion #"] = pd.to_numeric(responseP_RSBMRI_df["Lesion #"], errors="coerce")

        # Filter the DataFrame where 'Lesion #' is 1
        responseP_RSBMRI_df_filtered = responseP_RSBMRI_df[responseP_RSBMRI_df["Lesion #"] == 1].copy()

        responseP_RSBMRI_df = responseP_RSBMRI_df_filtered.drop(columns=["Lesion #"])

        responseP1_df = data["RS"][
            [
                "Subject",
                "Event Group Label",
                "Event Date",
                "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)",
                "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)",
                "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)",
                #   "Retreatment Cycle Number (IG_NS_NA_RS1.RETXCYCLENUM)",
                #   "For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_RS1.TX_YS_YH_UNSDAYR)",
                # "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
                "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)",
            ]
        ].copy()
        RS_new_col_name = {
            "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)": "Study Phase",
            "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)": "Primary Time Point",
            "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)": "Unscheduled Primary Day#",
            "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)": "Overall Objective Status",
        }
        responseP1_df = responseP1_df.rename(columns=RS_new_col_name).drop_duplicates()
        responseP1_df = responseP1_df.sort_values(by=["Subject"])

        responseP_df = pd.merge(
            responseP1_df,
            responseP_RSBMRI_df,
            on=["Subject", "Event Group Label", "Event Date", "Study Phase"],
            how="left",
        )

        # Convert Event Date to datetime object
        responseP_df["Event Date"] = pd.to_datetime(responseP_df["Event Date"])

        if not data["DSINITRT"].empty:
            # Primary FUP Response
            initRetxLastVisit_df = data["DSINITRT"][
                [
                    "Subject",
                    "Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)",
                    "Last Visit Completed in Long-Term Follow-Up (ig_DSINITRT1.DSLVCLTFUR)",
                ]
            ].copy()
            DSINITRT_new_col_name = {
                "Last Visit Completed in Primary Follow-Up (ig_DSINITRT1.DSLVCPFUR)": "Last Visit Completed in Primary FUP",
                "Last Visit Completed in Long-Term Follow-Up (ig_DSINITRT1.DSLVCLTFUR)": "Last Visit Completed in LTFU",
            }
            initRetxLastVisit_df = initRetxLastVisit_df.rename(columns=DSINITRT_new_col_name)

            initRetxLastVisit_df["Last Visit Completed in Primary FUP"] = initRetxLastVisit_df[
                initRetxLastVisit_df["Last Visit Completed in Primary FUP"].notna()
            ]["Last Visit Completed in Primary FUP"].astype(str)

            initRetxLastVisit_df["Last Visit Completed in LTFU"] = initRetxLastVisit_df[
                initRetxLastVisit_df["Last Visit Completed in LTFU"].notna()
            ]["Last Visit Completed in LTFU"].astype(str)

            initRetxLastVisit_df["Last Visit"] = None
            initRetxLastVisit_df["Last Visit"] = initRetxLastVisit_df["Last Visit Completed in Primary FUP"].fillna(
                ""
            ) + initRetxLastVisit_df["Last Visit Completed in LTFU"].fillna("")

            initRetxLastVisit_df = initRetxLastVisit_df.drop(
                columns=[
                    "Last Visit Completed in Primary FUP",
                    "Last Visit Completed in LTFU",
                ]
            )

            pd.set_option("future.no_silent_downcasting", True)
            initRetxLastVisit_df = initRetxLastVisit_df.fillna("").infer_objects(copy=False)

            DSEOS_df = data["DSEOS"][
                [
                    "Subject",
                    #  "Reason for End of Study? (ig_DSEOS2.EOSCOD1)",
                    "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)",
                    "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)",
                    "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)",
                ]
            ].copy()
            DSEOS_new_col_name = {
                # "Reason for End of Study? (ig_DSEOS2.EOSCOD1)": "Off-Study Reason",
                "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)": "Last Study Phase Completed",
                "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)": "Last Primary FUP",
                "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)": "Last Primary Retreatment",
            }
            DSEOS_df = DSEOS_df.rename(columns=DSEOS_new_col_name)
            # End of study on primary treatment
            EOS_df = DSEOS_df[(DSEOS_df["Last Study Phase Completed"] == "Primary Treatment")].copy()

            EOS_df["Last Visit"] = None
            EOS_df["Last Visit"] = EOS_df["Last Primary FUP"].fillna("")

            EOS_df = EOS_df.drop(
                columns=[
                    "Last Primary FUP",
                    "Last Study Phase Completed",
                    "Last Primary Retreatment",
                    # "Off-Study Reason",
                ]
            )

            pd.set_option("future.no_silent_downcasting", True)
            EOS_df = EOS_df.fillna("").infer_objects(copy=False)

            # End of study on retreatment
            EOSR_df = DSEOS_df[(DSEOS_df["Last Study Phase Completed"] == "Retreatment")].copy()

            EOSR_df["Last Visit"] = None
            EOSR_df["Last Visit"] = EOSR_df["Last Primary Retreatment"].fillna("")

            EOSR_df = EOSR_df.drop(
                columns=[
                    "Last Primary Retreatment",
                    "Last Study Phase Completed",
                    "Last Primary FUP",
                    #   "Off-Study Reason",
                ]
            )

            pd.set_option("future.no_silent_downcasting", True)
            EOSR_df = EOSR_df.fillna("").infer_objects(copy=False)

        drop_RS_columns = [
            "Event Group Label",
            "Study Phase",
            "Primary Time Point",
            "Unscheduled Primary Day#",
            "Event Date",
        ]
        drop_LASTV_columns = [
            "Event Group Label",
            "Study Phase",
            "Primary Time Point",
            "Unscheduled Primary Day#",
            "Event Date",
            "Last Visit",
        ]

        responseP_df = responseP_df[responseP_df["Study Phase"] == "Primary Treatment"].copy()

        DAY1_df = responseP_df[responseP_df["Event Group Label"] == "Day 1"].copy()
        DAY1_df = DAY1_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, DAY1_df, on="Subject", how="left")

        DAY28_df = responseP_df[responseP_df["Event Group Label"] == "Day 28"].copy()
        DAY28_df = DAY28_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, DAY28_df, on="Subject", how="left", suffixes=("", "_D28"))

        M2_df = responseP_df[responseP_df["Event Group Label"] == "Month 2"].copy()
        M2_df = pd.merge(M2_df, initRetxLastVisit_df, on="Subject", how="outer")

        M2_df["Overall Objective Status"] = M2_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"]) and row["Last Visit"].strip() == "Day 28"
            else row["Overall Objective Status"],
            axis=1,
        )
        M2_df["% Change is SPD"] = M2_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"]) and row["Last Visit"].strip() == "Day 28"
            else row["% Change is SPD"],
            axis=1,
        )
        M2_df = M2_df.drop(columns=drop_LASTV_columns)
        M2_df = pd.merge(M2_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M2_df["Overall Objective Status"] = M2_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"]) and (row["Last Visit"].strip() == "Day 28")
            else row["Overall Objective Status"],
            axis=1,
        )
        M2_df["% Change is SPD"] = M2_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"]) and (row["Last Visit"].strip() == "Day 28")
            else row["% Change is SPD"],
            axis=1,
        )
        # print(M2_df)
        M2_df = M2_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M2_df = M2_df.drop(columns=drop_RS_columns)

        response_df = pd.merge(response_df, M2_df, on="Subject", how="left", suffixes=("", "_M2"))

        M4_df = responseP_df[responseP_df["Event Group Label"] == "Month 4"].copy()
        M4_df = pd.merge(M4_df, initRetxLastVisit_df, on="Subject", how="outer")

        M4_df["Overall Objective Status"] = M4_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M4_df["% Change is SPD"] = M4_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M4_df = M4_df.drop(columns=drop_LASTV_columns)
        M4_df = pd.merge(M4_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M4_df["Overall Objective Status"] = M4_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M4_df["% Change is SPD"] = M4_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        # print(M4_df)
        M4_df = M4_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M4_df = M4_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, M4_df, on="Subject", how="left", suffixes=("", "_M4"))

        M6_df = responseP_df[responseP_df["Event Group Label"] == "Month 6"].copy()
        M6_df = pd.merge(M6_df, initRetxLastVisit_df, on="Subject", how="outer")

        M6_df["Overall Objective Status"] = M6_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M6_df["% Change is SPD"] = M6_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M6_df = M6_df.drop(columns=drop_LASTV_columns)
        M6_df = pd.merge(M6_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M6_df["Overall Objective Status"] = M6_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M6_df["% Change is SPD"] = M6_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
            )
            else row["% Change is SPD"],
            axis=1,
        )

        M6_df = M6_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M6_df = M6_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, M6_df, on="Subject", how="left", suffixes=("", "_M6"))

        M8_df = responseP_df[responseP_df["Event Group Label"] == "Month 8"].copy()
        M8_df = pd.merge(M8_df, initRetxLastVisit_df, on="Subject", how="outer")

        M8_df["Overall Objective Status"] = M8_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M8_df["% Change is SPD"] = M8_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M8_df = M8_df.drop(columns=drop_LASTV_columns)
        M8_df = pd.merge(M8_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M8_df["Overall Objective Status"] = M8_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M8_df["% Change is SPD"] = M8_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
            )
            else row["% Change is SPD"],
            axis=1,
        )

        M8_df = M8_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M8_df = M8_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, M8_df, on="Subject", how="left", suffixes=("", "_M8"))

        M10_df = responseP_df[responseP_df["Event Group Label"] == "Month 10"].copy()
        M10_df = pd.merge(M10_df, initRetxLastVisit_df, on="Subject", how="outer")

        M10_df["Overall Objective Status"] = M10_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M10_df["% Change is SPD"] = M10_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M10_df = M10_df.drop(columns=drop_LASTV_columns)
        M10_df = pd.merge(M10_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M10_df["Overall Objective Status"] = M10_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M10_df["% Change is SPD"] = M10_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
            )
            else row["% Change is SPD"],
            axis=1,
        )

        M10_df = M10_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M10_df = M10_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, M10_df, on="Subject", how="left", suffixes=("", "_M10"))

        M12_df = responseP_df[responseP_df["Event Group Label"] == "Month 12"].copy()
        M12_df = pd.merge(M12_df, initRetxLastVisit_df, on="Subject", how="outer")

        M12_df["Overall Objective Status"] = M12_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
                | (row["Last Visit"].strip() == "Month 10")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M12_df["% Change is SPD"] = M12_df.apply(
            lambda row: "Transitioned to Retreatment"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
                | (row["Last Visit"].strip() == "Month 10")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M12_df = M12_df.drop(columns=drop_LASTV_columns)
        M12_df = pd.merge(M12_df, EOS_df, on="Subject", how="outer")

        # Replace "Overall Objective Status" to "Off-Study", handling possible trailing spaces and NaN values
        M12_df["Overall Objective Status"] = M12_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
                | (row["Last Visit"].strip() == "Month 10")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M12_df["% Change is SPD"] = M12_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28")
                | (row["Last Visit"].strip() == "Month 2")
                | (row["Last Visit"].strip() == "Month 3")
                | (row["Last Visit"].strip() == "Month 4")
                | (row["Last Visit"].strip() == "Month 5")
                | (row["Last Visit"].strip() == "Month 6")
                | (row["Last Visit"].strip() == "Month 8")
                | (row["Last Visit"].strip() == "Month 9")
                | (row["Last Visit"].strip() == "Month 10")
            )
            else row["% Change is SPD"],
            axis=1,
        )

        M12_df = M12_df.drop(
            columns=[
                "Last Visit",
            ]
        )
        # M12_df = M12_df.drop(columns=drop_RS_columns)
        response_df = pd.merge(response_df, M12_df, on="Subject", how="left", suffixes=("", "_M12"))

        UNS_df = responseP_df[(responseP_df["Event Group Label"] == "Unscheduled Disease Assessments")].copy()

        # Add Day in front of unscheduedled Day#
        # Define a function to handle conversion safely
        def safe_int_conversion(value):
            try:
                # Attempt to convert the value to an integer
                return "Day " + str(int(value))
            except (ValueError, TypeError):
                # Return an empty string if conversion fails
                return "Day " + str(value)

        UNS_df["Primary Time Point"] = UNS_df[UNS_df["Primary Time Point"].notna()]["Primary Time Point"].astype(str)
        UNS_df.loc[
            UNS_df["Primary Time Point"] == "Unscheduled",
            "Primary Time Point",
        ] = ""

        UNS_df["Unscheduled Primary Day#"] = UNS_df.apply(
            lambda row: safe_int_conversion(row["Unscheduled Primary Day#"])
            if pd.notna(row["Unscheduled Primary Day#"]) and row["Unscheduled Primary Day#"] != ""
            else "",
            axis=1,
        )

        UNS_df["Primary Time Point"] = UNS_df["Primary Time Point"].fillna("") + UNS_df[
            "Unscheduled Primary Day#"
        ].fillna("")

        # Alternatively, if you want to ensure unique concatenated values for the same time point only, you can use the groupby approach
        def concatenate_group(group):
            group["Overall Objective Status"] = (
                group["Primary Time Point"].fillna("").astype(str)
                + " "
                + group["Overall Objective Status"].fillna("").astype(str)
                + "/"
                + group["% Change is SPD"].fillna("").astype(str)
            )
            return group

        UNS_df = UNS_df.groupby(["Primary Time Point"]).apply(concatenate_group).reset_index(drop=True)

        # Removing duplicates if there are any within the same group
        UNS_df = UNS_df.drop_duplicates(subset=["Overall Objective Status"])

        # UNS_df = UNS_df_grouped_unique
        UNS_df = UNS_df.drop(columns=drop_RS_columns)

        UNS_df = UNS_df.drop(columns=["% Change is SPD"])
        response_df = pd.merge(response_df, UNS_df, on="Subject", how="left", suffixes=("", "_UNS"))

        pd.set_option("future.no_silent_downcasting", True)
        response_df = response_df.fillna("").infer_objects(copy=False)
        # replace "-" for blank unscheduled response
        response_df["Overall Objective Status_UNS"] = response_df["Overall Objective Status_UNS"].replace("", "-")

        # replacement_date = pd.Timestamp("1900-01-01")
        # # Replace missing dates with the specified value
        # response_df["Event Date"] = response_df["Event Date"].fillna(replacement_date)

        response_df = pd.merge(response_df, Treated_df, on="Subject", how="right")
        response_df = response_df.fillna("")
        response_df = response_df.sort_values(by=["Subject"])
        response_df = response_df.drop(columns=["Treated"])
        response_df = response_df.drop_duplicates()

        # print(response_df)

        # TODO: RESPONSE LISTING for Retreated Subjects
        INFR_df = data["EXINF"][
            ["Subject", "Event Group Label", "Was study treatment administered? (ig_EXINF1.INFOCCUR)"]
        ].copy()
        INFR_new_col_name = {
            "Was study treatment administered? (ig_EXINF1.INFOCCUR)": "Treated",
        }
        INFR_df = INFR_df.rename(columns=INFR_new_col_name)
        Retreated_df = INFR_df[(INFR_df["Event Group Label"] == "Day 0-R1") & (INFR_df["Treated"] == "Yes")].copy()
        Retreated_df = Retreated_df.drop(columns="Event Group Label")

        Retreated_subject_df = Retreated_df["Subject"].copy()

        responseR_df = Retreated_subject_df

        responseR_df = add_rename_column_corelisting(
            responseR_df,
            data,
            "RS",
            "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
            "Measurable vs. Non-Measurable Disease",
            "Subject",
            # "Event Group Label",
        ).drop_duplicates()

        # Drop rows where 'Measurable vs. Non-Measurable Disease' is NaN or empty string
        responseR_df = responseR_df[
            responseR_df["Measurable vs. Non-Measurable Disease"].notna()
            & responseR_df["Measurable vs. Non-Measurable Disease"].astype(bool)
        ]

        # Corrected lambda function to replace values based on condition and handle missing values
        responseR_df["Measurable vs. Non-Measurable Disease"] = responseR_df.apply(
            lambda row: "Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "Yes"
            else "Non-Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "No"
            else "Unknown",
            axis=1,
        )

        responsePR_RSBMRI_df = data["RSBMRI"][
            [
                "Subject",
                "Event Group Label",
                "Event Date",
                "Study Phase (IG_NS_NA_RSBMRI1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)",
                # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)",
                # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)",
                "Lesion # (IG_NS_NA_RSBMRI3.NM_YS_YH_LESNUM)",
                "Measurable Lesion Change Percentage from Baseline (%) (IG_NS_NA_RSBMRI4.NM_NS_YH_TLCP)",
            ]
        ].copy()
        RSBMRIR_new_col_name = {
            "Study Phase (IG_NS_NA_RSBMRI1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)": "Study Phase",
            # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)": "Primary Time Point",
            # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)": "Unscheduled Primary Day#",
            "Lesion # (IG_NS_NA_RSBMRI3.NM_YS_YH_LESNUM)": "Lesion #",
            "Measurable Lesion Change Percentage from Baseline (%) (IG_NS_NA_RSBMRI4.NM_NS_YH_TLCP)": "% Change is SPD",
        }
        responsePR_RSBMRI_df = responsePR_RSBMRI_df.rename(columns=RSBMRIR_new_col_name).drop_duplicates()
        #
        # Convert 'Lesion #' to numeric (if needed) and handle errors by coercing non-numeric values to NaN
        responsePR_RSBMRI_df["Lesion #"] = pd.to_numeric(responsePR_RSBMRI_df["Lesion #"], errors="coerce")

        # Filter the DataFrame where 'Lesion #' is 1
        responsePR_RSBMRI_df_filtered = responsePR_RSBMRI_df[
            (responsePR_RSBMRI_df["Study Phase"] == "Retreatment") & (responsePR_RSBMRI_df["Lesion #"] == 1)
        ].copy()

        responsePR_RSBMRI_df = responsePR_RSBMRI_df_filtered.drop(columns=["Lesion #"])

        responseP1R_df = data["RS"][
            [
                "Subject",
                "Event Group Label",
                "Event Date",
                "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)",
                "Retreatment Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPTR_cl_NS_NETPT2)",
                "For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_RS1.TX_YS_YH_UNSDAYR)",
                #   "Retreatment Cycle Number (IG_NS_NA_RS1.RETXCYCLENUM)",
                #   "For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_RS1.TX_YS_YH_UNSDAYR)",
                # "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
                "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)",
            ]
        ].copy()
        RSR_new_col_name = {
            "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)": "Study Phase",
            "Retreatment Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPTR_cl_NS_NETPT2)": "Retreatment Time Point",
            "For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_RS1.TX_YS_YH_UNSDAYR)": "Unscheduled Retreatment Day#",
            "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)": "Overall Objective Status",
        }
        responseP1R_df = responseP1R_df.rename(columns=RSR_new_col_name).drop_duplicates()
        responseP1R_df = responseP1R_df.sort_values(by=["Subject"])

        responsePR_df = pd.merge(
            responseP1R_df,
            responsePR_RSBMRI_df,
            on=["Subject", "Event Group Label", "Event Date", "Study Phase"],
            how="left",
        )

        # Convert Event Date to datetime object
        responsePR_df["Event Date"] = pd.to_datetime(responsePR_df["Event Date"])

        drop_RSR_columns = [
            "Event Group Label",
            "Study Phase",
            "Retreatment Time Point",
            "Unscheduled Retreatment Day#",
            "Event Date",
        ]
        drop_LASTVR_columns = [
            "Event Group Label",
            "Study Phase",
            "Retreatment Time Point",
            "Unscheduled Retreatment Day#",
            "Event Date",
            "Last Visit",
        ]
        responsePR_df = responsePR_df[responsePR_df["Study Phase"] == "Retreatment"].copy()

        DAY1R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Day 1-R1"].copy()
        DAY1R1_df = DAY1R1_df.drop(columns=drop_RSR_columns)
        responseR_df = pd.merge(responseR_df, DAY1R1_df, on="Subject", how="left")

        DAY28R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Day 28-R1"].copy()
        DAY28R1_df = DAY28R1_df.drop(columns=drop_RSR_columns)
        responseR_df = pd.merge(responseR_df, DAY28R1_df, on="Subject", how="left", suffixes=("", "_D28R1"))

        M2R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 2-R1"].copy()
        M2R1_df = pd.merge(M2R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M2R1_df["Overall Objective Status"] = M2R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"]) and (row["Last Visit"].strip() == "Day 28-R")
            else row["Overall Objective Status"],
            axis=1,
        )
        M2R1_df["% Change is SPD"] = M2R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"]) and (row["Last Visit"].strip() == "Day 28-R")
            else row["% Change is SPD"],
            axis=1,
        )
        M2R1_df = M2R1_df.drop(columns=drop_LASTVR_columns)

        # M2R1_df = M2R1_df.drop(columns=drop_RSR_columns)
        responseR_df = pd.merge(responseR_df, M2R1_df, on="Subject", how="left", suffixes=("", "_M2R1"))

        M4R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 4-R1"].copy()
        M4R1_df = pd.merge(M4R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M4R1_df["Overall Objective Status"] = M4R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M4R1_df["% Change is SPD"] = M4R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M4R1_df = M4R1_df.drop(columns=drop_LASTVR_columns)
        responseR_df = pd.merge(responseR_df, M4R1_df, on="Subject", how="left", suffixes=("", "_M4R1"))

        M6R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 6-R1"].copy()
        M6R1_df = pd.merge(M6R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M6R1_df["Overall Objective Status"] = M6R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M6R1_df["% Change is SPD"] = M6R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M6R1_df = M6R1_df.drop(columns=drop_LASTVR_columns)
        responseR_df = pd.merge(responseR_df, M6R1_df, on="Subject", how="left", suffixes=("", "_M6R1"))

        M8R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 8-R1"].copy()
        M8R1_df = pd.merge(M8R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M8R1_df["Overall Objective Status"] = M8R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M8R1_df["% Change is SPD"] = M8R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M8R1_df = M8R1_df.drop(columns=drop_LASTVR_columns)
        responseR_df = pd.merge(responseR_df, M8R1_df, on="Subject", how="left", suffixes=("", "_M8R1"))

        M10R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 10-R1"].copy()
        M10R1_df = pd.merge(M10R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M10R1_df["Overall Objective Status"] = M10R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
                | (row["Last Visit"].strip() == "Month 8-R")
                | (row["Last Visit"].strip() == "Month 9-R")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M10R1_df["% Change is SPD"] = M10R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
                | (row["Last Visit"].strip() == "Month 8-R")
                | (row["Last Visit"].strip() == "Month 9-R")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M10R1_df = M10R1_df.drop(columns=drop_LASTVR_columns)
        # M10R1_df = M10R1_df.drop(columns=drop_RSR_columns)
        responseR_df = pd.merge(responseR_df, M10R1_df, on="Subject", how="left", suffixes=("", "_M10R1"))

        M12R1_df = responsePR_df[responsePR_df["Event Group Label"] == "Month 12-R1"].copy()
        M12R1_df = pd.merge(M12R1_df, EOSR_df, on="Subject", how="outer")

        # Replace "Overall Tumor Response" to "Off-Study", handling possible trailing spaces and NaN values
        M12R1_df["Overall Objective Status"] = M12R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
                | (row["Last Visit"].strip() == "Month 8-R")
                | (row["Last Visit"].strip() == "Month 9-R")
                | (row["Last Visit"].strip() == "Month 10-R")
            )
            else row["Overall Objective Status"],
            axis=1,
        )
        M12R1_df["% Change is SPD"] = M12R1_df.apply(
            lambda row: "Off-Study"
            if pd.notna(row["Last Visit"])
            and (
                (row["Last Visit"].strip() == "Day 28-R")
                | (row["Last Visit"].strip() == "Month 2-R")
                | (row["Last Visit"].strip() == "Month 3-R")
                | (row["Last Visit"].strip() == "Month 4-R")
                | (row["Last Visit"].strip() == "Month 5-R")
                | (row["Last Visit"].strip() == "Month 6-R")
                | (row["Last Visit"].strip() == "Month 8-R")
                | (row["Last Visit"].strip() == "Month 9-R")
                | (row["Last Visit"].strip() == "Month 10-R")
            )
            else row["% Change is SPD"],
            axis=1,
        )
        M12R1_df = M12R1_df.drop(columns=drop_LASTVR_columns)

        responseR_df = pd.merge(responseR_df, M12R1_df, on="Subject", how="left", suffixes=("", "_M12R1"))

        UNSR_df = responsePR_df[(responsePR_df["Event Group Label"] == "Unscheduled Disease Assessments")].copy()

        # Add Day in front of unscheduedled Day#
        # Define a function to handle conversion safely
        def safe_int_conversion_R(value):
            try:
                # Attempt to convert the value to an integer
                return "Day " + str(int(value)) + "-R"
            except (ValueError, TypeError):
                # Return an empty string if conversion fails
                return "Day " + str(value) + "-R"

        UNSR_df["Retreatment Time Point"] = UNSR_df[UNSR_df["Retreatment Time Point"].notna()][
            "Retreatment Time Point"
        ].astype(str)
        UNSR_df.loc[
            UNSR_df["Retreatment Time Point"] == "Unscheduled",
            "Retreatment Time Point",
        ] = ""

        UNSR_df["Unscheduled Retreatment Day#"] = UNSR_df.apply(
            lambda row: safe_int_conversion_R(row["Unscheduled Retreatment Day#"])
            if pd.notna(row["Unscheduled Retreatment Day#"]) and row["Unscheduled Retreatment Day#"] != ""
            else "",
            axis=1,
        )

        UNSR_df["Retreatment Time Point"] = UNSR_df["Retreatment Time Point"].fillna("") + UNSR_df[
            "Unscheduled Retreatment Day#"
        ].fillna("")

        # Alternatively, if you want to ensure unique concatenated values for the same time point only, you can use the groupby approach
        def concatenate_group(group):
            group["Overall Objective Status"] = (
                group["Retreatment Time Point"].fillna("").astype(str)
                + " "
                + group["Overall Objective Status"].fillna("").astype(str)
                + "/"
                + group["% Change is SPD"].fillna("").astype(str)
            )
            return group

        UNSR_df = UNSR_df.groupby(["Retreatment Time Point"]).apply(concatenate_group).reset_index(drop=True)

        # Removing duplicates if there are any within the same group
        UNSR_df = UNSR_df.drop_duplicates(subset=["Overall Objective Status"])

        UNSR_df = UNSR_df.drop(columns=drop_RSR_columns)

        UNSR_df = UNSR_df.drop(columns=["% Change is SPD"])
        responseR_df = pd.merge(responseR_df, UNSR_df, on="Subject", how="left", suffixes=("", "_UNSR"))

        pd.set_option("future.no_silent_downcasting", True)
        responseR_df = responseR_df.fillna("").infer_objects(copy=False)
        # replace "-" for blank unscheduled response
        responseR_df["Overall Objective Status_UNSR"] = responseR_df["Overall Objective Status_UNSR"].replace("", "-")

        # replacement_date = pd.Timestamp("1900-01-01")
        # # Replace missing dates with the specified value
        # response_df["Event Date"] = response_df["Event Date"].fillna(replacement_date)

        responseR_df = pd.merge(responseR_df, Retreated_df, on="Subject", how="right")
        responseR_df = responseR_df.fillna("")
        responseR_df = responseR_df.sort_values(by=["Subject"])
        responseR_df = responseR_df.drop(columns=["Treated"])
        responseR_df = responseR_df.drop_duplicates()

        return response_df, responseR_df

    def response_stats(self):
        data = self.data
        ### TODO: REPONSE STATS

        INF_df = data["EXINF"][
            ["Subject", "Event Group Label", "Was study treatment administered? (ig_EXINF1.INFOCCUR)"]
        ].copy()
        INF_new_col_name = {
            "Was study treatment administered? (ig_EXINF1.INFOCCUR)": "Treated",
        }
        INF_df = INF_df.rename(columns=INF_new_col_name)
        Treated_df = INF_df[(INF_df["Event Group Label"] == "Day 0") & (INF_df["Treated"] == "Yes")].copy()
        Treated_df = Treated_df.drop(columns="Event Group Label")

        Treated_subject_df = Treated_df["Subject"].copy()

        response_stat_df = Treated_subject_df

        response_stat_df = add_rename_column_corelisting(
            response_stat_df,
            data,
            "RS",
            "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
            "Measurable vs. Non-Measurable Disease",
            "Subject",
        ).drop_duplicates()

        # Drop rows where 'Measurable vs. Non-Measurable Disease' is NaN or empty string
        response_stat_df = response_stat_df[
            response_stat_df["Measurable vs. Non-Measurable Disease"].notna()
            & response_stat_df["Measurable vs. Non-Measurable Disease"].astype(bool)
        ]

        # Corrected lambda function to replace values based on condition and handle missing values
        response_stat_df["Measurable vs. Non-Measurable Disease"] = response_stat_df.apply(
            lambda row: "Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "Yes"
            else "Non-Measurable"
            if row["Measurable vs. Non-Measurable Disease"] == "No"
            else "Unknown",
            axis=1,
        )

        responseP1_stat_df = data["RS"][
            [
                "Subject",
                # "Event Group Label",
                # "Event Date",
                # "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)",
                # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)",
                # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)",
                # #   "Retreatment Cycle Number (IG_NS_NA_RS1.RETXCYCLENUM)",
                # #   "For Unscheduled Retreatment Time Point, Specify Day # (IG_NS_NA_RS1.TX_YS_YH_UNSDAYR)",
                # # "Did the subject have measurable disease at study treatment baseline? (IG_NS_NA_RS1.CL_NS_YH_MDATBL_cl_YS_YN1)",
                "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)",
            ]
        ].copy()
        RS_new_col_name = {
            # "Study Phase (IG_NS_NA_RS1.CL_YS_NH_STUDYPHS_cl_YS_STUDYPHS)": "Study Phase",
            # "Primary Time Point (IG_NS_NA_RS1.CL_YS_NH_RSTPT_cl_NS_NETPT1)": "Primary Time Point",
            # "For Unscheduled Primary Time Point, Specify Day #  (IG_NS_NA_RS1.TX_YS_YH_UNSDAY)": "Unscheduled Primary Day#",
            "Overall Objective Status (IG_NS_NA_RS2.CL_NS_NH_OOS_cl_NS_OOSRESP1)": "Overall Objective Status",
        }
        responseP1_stat_df = responseP1_stat_df.rename(columns=RS_new_col_name).drop_duplicates()
        responseP1_stat_df = responseP1_stat_df.sort_values(by=["Subject"])

        response_stat_df = pd.merge(
            response_stat_df,
            responseP1_stat_df,
            on=["Subject"],
            how="left",
        )

        # # Convert Event Date to datetime object
        # response_stat_df["Event Date"] = pd.to_datetime(response_stat_df["Event Date"])

        # Disease Response Overall Objective Status dictionary
        DR_OOS_dict = {
            "Confirmed CR": 1,
            "Confirmed PR": 2,
            "SD": 3,
            "Confirmed PD": 4,
            "Not Evaludated": 5,  # Pending Confirmation of Response
            "Preliminary CR": 9,
            "Preliminary PR": 9,
            "Preliminary PD": 9,
            "Not Evaluable": 9,
            "": 9,
        }
        response_stat_df = response_stat_df.replace([np.nan, np.inf, -np.inf], "")
        # # Gather all stats of treated subjects
        # total_infused_df = self.infusion_df.copy()
        # Gather treated subjects with measurable disease at baseline
        total_infused_measurable_df = response_stat_df[
            response_stat_df["Measurable vs. Non-Measurable Disease"] == "Measurable"
        ][["Subject", "Overall Objective Status"]].copy()
        self.subject_infused_measurable_count = len(total_infused_measurable_df["Subject"].unique())

        # print(total_infused_measurable_df)

        OOS_Response_Codelist = [
            "Confirmed CR",
            "Confirmed PR",
            "SD",
            "Confirmed PD",
            "Pending Confirmation of Response",  # "Not Evaludated",
        ]
        # Convert Overall Objective Status to numeric values
        total_infused_measurable_df["OOS-Score"] = total_infused_measurable_df["Overall Objective Status"].map(
            DR_OOS_dict
        )
        # print(total_infused_measurable_df)
        # * BEST RESPONSE
        ## Best Disease Response
        # Get the indices of the rows with the minimum 'Overall Objective Status' for each 'Subject'
        response_best_OOS_idx = total_infused_measurable_df.groupby("Subject")["OOS-Score"].idxmin()

        # Select these rows for the best PET-based response
        response_best_OOS_df = total_infused_measurable_df.loc[response_best_OOS_idx].copy()
        # Select the columns subject and PET-Based NHL Disease Response from responseA_best_PET_df
        response_best_OOS_df = response_best_OOS_df[
            [
                "Subject",
                "Overall Objective Status",
            ]
        ]
        # TODO: RESPONSE STATS
        if self.subject_infused_measurable_count > 0:
            response_stat = response_best_OOS_df.copy()
            # replace 'Not Evaludated' with 'Pending Confirmation of Response' for all columns in response_stat
            response_stat = response_stat.replace("Not Evaludated", "Pending Confirmation of Response")
            self.response_stat_OOS = get_stats_percentage2(
                "Overall Objective Status", OOS_Response_Codelist, response_stat
            )
        # print(self.response_stat_OOS)

    def status_listing(self, enrollment_listing_df):
        data = self.data
        # AE and SAE data
        #    if not data["AE"].empty:
        AE_df = data["AE"][
            [
                "Subject",
                "AE or SAE? (ig_AE2.AESEV)",
            ]
        ].copy()
        AE_new_col_name = {
            "AE or SAE? (ig_AE2.AESEV)": "AE or SAE?",
        }
        AE_df = AE_df.rename(columns=AE_new_col_name)

        # # replaces all occurrences of NaN, positive infinity, and negative infinity in the infusion_df dataframe with empty strings.
        # infusion_df = infusion_df.replace([np.nan, np.inf, -np.inf], "N/A")
        status_df = self.enrollment_listing_df[self.enrollment_listing_df["Screen Fail"].str.strip() == "No"][
            ["Subject", "Cohort"]
        ]

        status_df["AE"] = status_df["Subject"].apply(lambda x: "Y" if x in AE_df["Subject"].values else "N")

        status_df["SAE"] = status_df["Subject"].apply(
            lambda x: "Y" if x in AE_df[AE_df["AE or SAE?"] == "SAE"]["Subject"].values else "N"
        )
        # replaces all occurrences of NaN, positive infinity, and negative infinity in the infusion_df dataframe with empty strings.
        status_df = status_df.replace([np.nan, np.inf, -np.inf], "")

        # Event Label Update dictionary
        event_1_dict = {
            "Pre-Treatment Safety Visit": "Pre-Treatment",
            "Primary Treatment and Follow-Up": "Primary Follow-up",
            "Long-Term Follow-Up Months 3-60": "LTFU",
            "Pre-Retreatment Safety 1": "Pre-Retreatment",
            "Primary Retreatment and Follow-Up 1": "Primary Retreatment Follow-up",
            "Retreatment Long-Term Follow-Up Months 3-60 (1)": "Retreatment LTFU",
        }

        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        status_SV_df = data["DSSV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        status_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]

        # status_DSSVLTFU_df["Event Group Label"] = status_DSSVLTFU_df["Event Label"].apply(map_event)

        # Combine DSSVLTFU with SV dataframe vertically
        status_SV_df = pd.concat([status_SV_df, status_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        status_SV_df = status_SV_df.sort_values(by=["Subject", "Event Date"])

        # For each unique subject, get the last row of the dataframe
        status_SV_df = status_SV_df.groupby("Subject").tail(1)

        # Merge left with the current response dataframe
        status_df = pd.merge(
            status_df,
            status_SV_df[["Subject", "Event Label"]],
            on="Subject",
            how="left",
        )

        # Rename the column Event Label to Event Label (Study Status)
        status_df["Event Label"] = status_df["Event Label"].map(event_1_dict)

        status_df["Event Label3"] = status_df["Subject"].apply(
            lambda x: "Pre-Treatment"
            if (
                self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Treated"]
                .fillna("")
                .str.strip()
                .values[0]
                == "Pending"
            )
            else ""
        )
        # print(self.enrollment_listing_df)
        status_df["Event Label4"] = status_df["Subject"].apply(
            lambda x: "Withdrawn Prior to Study Treatment"
            if (
                self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Treated"]
                .fillna("")
                .str.strip()
                .values[0]
                == "No"
            )
            & (
                self.enrollment_listing_df[self.enrollment_listing_df["Subject"] == x]["Screen Fail"]
                .fillna("")
                .str.strip()
                .values[0]
                == "No"
            )
            else ""
        )
        # Merge all event group label into study status
        status_df["Event Label"] = (
            status_df["Event Label"].fillna("")
            + status_df["Event Label3"].fillna("")
            + status_df["Event Label4"].fillna("")
        )
        status_df["Event Label"].fillna(status_df["Event Label"], inplace=True)
        status_df = status_df.drop(
            columns=[
                "Event Label3",
                "Event Label4",
            ]
        )
        # print(status_df)

        # filter the data frame to only include subjects whose end of study date is later than or equal to main consent date
        filteredemrollment_df = self.enrollment_listing_df[
            self.enrollment_listing_df["End of Study Date"] >= self.enrollment_listing_df["Main Consent Date"]
        ]
        DSEOS_df = data["DSEOS"][
            [
                "Subject",
                "Reason for End of Study? (ig_DSEOS2.EOSCOD1)",
                "Provide Supportive Information (ig_DSEOS2.EOSTERM)",
                "Principal Cause of Death (ig_DSEOS2.PRCDTH)",
                "Specify Principal Cause of Death (ig_DSEOS2.PRCDTHOS)",
                "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)",
                "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)",
                "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)",
            ]
        ].copy()
        DSEOS_new_col_name = {
            "Reason for End of Study? (ig_DSEOS2.EOSCOD1)": "Off-Study Reason",
            "Provide Supportive Information (ig_DSEOS2.EOSTERM)": "Off-Study Reason sp1",
            "Principal Cause of Death (ig_DSEOS2.PRCDTH)": "Off-Study Reason sp2",
            "Specify Principal Cause of Death (ig_DSEOS2.PRCDTHOS)": "Off-Study Reason sp3",
            "Last Study Phase (ig_DSEOS1.STUDYPHSEOS)": "Last Study Phase Completed",
            "Last Study Visit Completed in Primary Treatment (ig_DSEOS1.EOSLASTVISIT)": "Last Primary FUP",
            "Last Study Visit Completed in Retreatment (ig_DSEOS1.EOSLASTVISITR)": "Last Primary Retreatment",
        }
        DSEOS_df = DSEOS_df.rename(columns=DSEOS_new_col_name)

        # Merge last study visit
        DSEOS_df["Last Study Visit"] = DSEOS_df["Last Primary FUP"].fillna("") + DSEOS_df[
            "Last Primary Retreatment"
        ].fillna("")

        # Merge off-study reason
        DSEOS_df["Off-Study Reason"] = (
            DSEOS_df["Off-Study Reason"].fillna("")
            + " "
            + DSEOS_df["Off-Study Reason sp1"].fillna("")
            + DSEOS_df["Off-Study Reason sp2"].fillna("")
            + " "
            + DSEOS_df["Off-Study Reason sp3"].fillna("")
        )

        filteredDSEOS_df = DSEOS_df[(DSEOS_df["Subject"].isin(filteredemrollment_df["Subject"].values))].copy()

        status_df["Event Label"] = status_df.apply(
            lambda row: "Off Study"
            if (row["Subject"] in filteredDSEOS_df["Subject"].values)
            & ("Withdrawn Prior to Study Treatment" not in row["Event Label"])
            else "On Study/" + row["Event Label"],
            axis=1,
        )
        status_df = status_df.replace(
            "On Study/Withdrawn Prior to Study Treatment", "Withdrawn Prior to Study Treatment"
        )
        status_df = status_df.replace(
            "On Study/Pre-TreatmentWithdrawn Prior to Study Treatment", "Withdrawn Prior to Study Treatment"
        )
        status_df = pd.merge(
            status_df,
            filteredDSEOS_df[["Subject", "Off-Study Reason", "Last Study Visit"]],
            on="Subject",
            how="left",
        )
        status_df = status_df.replace("On Study/Pre-TreatmentPre-Treatment", "On Study/Pre-Treatment")

        # replaces all occurrences of NaN, positive infinity, and negative infinity with empty strings.
        status_df = status_df.replace([np.nan, np.inf, -np.inf], "N/A")

        # Gather all stats of each cohort
        total_status_df = status_df.copy()

        totalCH1_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort 1"])].copy()
        # Total number of subjects for Cohort 1
        AECH1_total_count = get_stats_percentage("AE", totalCH1_status_df).T
        SAECH1_total_count = get_stats_percentage("SAE", totalCH1_status_df).T
        # merge AE and SAE dataframes
        safetyCH1_total_df = pd.concat([AECH1_total_count, SAECH1_total_count], axis=1)

        totalCH2_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort 2"])].copy()
        # Total number of subjects for Cohort 2
        AECH2_total_count = get_stats_percentage("AE", totalCH2_status_df).T
        SAECH2_total_count = get_stats_percentage("SAE", totalCH2_status_df).T
        # merge AE and SAE dataframes
        safetyCH2_total_df = pd.concat([AECH2_total_count, SAECH2_total_count], axis=1)

        totalCHN1_status_df = total_status_df[total_status_df["Cohort"].isin(["Cohort -1"])].copy()
        # Total number of subjects for Cohort -1
        AECHN1_total_count = get_stats_percentage("AE", totalCHN1_status_df).T
        SAECHN1_total_count = get_stats_percentage("SAE", totalCHN1_status_df).T
        # merge AE and SAE dataframes
        safetyCHN1_total_df = pd.concat([AECHN1_total_count, SAECHN1_total_count], axis=1)

        return AE_df, status_df, safetyCH1_total_df, safetyCH2_total_df, safetyCHN1_total_df

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
                        worksheet1.write(
                            1,
                            1 + i * 4,
                            "Total Consented\nN=" + str(self.status_list[i]["Total Consented"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            2 + i * 4,
                            "Screen Failed\nN=" + str(self.status_list[i]["Screen Failed"]),
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

                    worksheet1.merge_range("A3:I3", "Legal Sex", bold_11_format)
                    worksheet1.merge_range("A8:I8", "Age at Consent", bold_11_format)
                    worksheet1.merge_range("A12:I12", "Race", bold_11_format)
                    worksheet1.merge_range("A22:I22", "Ethnicity", bold_11_format)
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
                    worksheet3 = writer.book.add_worksheet("EGFR for Eligible Subjects")

                    # * WRITING HEADER AND FORMATTING
                    # Assuming 'EGFR_listing_df_output' is your DataFrame
                    self.EGFR_listing_df_output = self.EGFR_listing_df_output.replace(
                        [np.inf, -np.inf], np.nan
                    )  # Replace INF with NaN

                    self.EGFR_listing_df_output = self.EGFR_listing_df_output.fillna(
                        ""
                    )  # Replace NaN with a placeholder
                    # Replace column header for A1 from "Subject" to "Subject ID", columns name starting from 1 instead of 0
                    worksheet3.write("A1", "Subject ID", bold_12_wrap_format)
                    for i in range(1, len(self.EGFR_listing_df_output.columns)):
                        worksheet3.write(0, i, self.EGFR_listing_df_output.columns[i], bold_11_format)
                    # * FORMAT DATA
                    for i in range(0, len(self.EGFR_listing_df_output)):
                        for j in range(0, len(self.EGFR_listing_df_output.columns)):
                            worksheet3.write(i + 1, j, self.EGFR_listing_df_output.iloc[i, j], normal_data_format)
                    # Autofit
                    worksheet3.autofit()

                    ## TODO: Study Tx Statistics
                    worksheet4 = writer.book.add_worksheet("Study Tx Statistics")

                    # * FORMATING DATA
                    # Cohort 1
                    for i in range(0, len(self.infusion_statA)):
                        for j in range(0, len(self.infusion_statA.columns)):
                            worksheet4.write(
                                i + 4,
                                j + 1,
                                self.infusion_statA.iloc[i, j],
                                normal_data_format,
                            )
                    # Cohort 2
                    for i in range(0, len(self.infusion_statB)):
                        for j in range(0, len(self.infusion_statB.columns)):
                            worksheet4.write(
                                i + 8,
                                j + 1,
                                self.infusion_statB.iloc[i, j],
                                normal_data_format,
                            )
                    # Cohort -1
                    for i in range(0, len(self.infusion_statC)):
                        for j in range(0, len(self.infusion_statC.columns)):
                            worksheet4.write(
                                i + 12,
                                j + 1,
                                self.infusion_statC.iloc[i, j],
                                normal_data_format,
                            )

                    # * WRITING HEADER AND FORMATTING
                    stat_order = ["Mean SD", "Median", "Range"]
                    worksheet4.merge_range(
                        "B1:G1",
                        "Study Treatment Statistics (N="
                        + str(self.infusion_count[0] + self.infusion_count[1] + self.infusion_count[2])
                        + ")",
                        bold_12_wrap_format,
                    )

                    worksheet4.merge_range("B2:D2", "Cells Infused", bold_12_wrap_format)
                    worksheet4.merge_range("E2:G2", "Transduction Efficiency", bold_12_wrap_format)
                    worksheet4.write("B3", "Total Cell Dose", bold_12_wrap_format)
                    worksheet4.write("C3", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                    worksheet4.write("D3", "Met Target Dose", bold_12_wrap_format)
                    worksheet4.write("E3", "%scFV (EGFR)", bold_12_wrap_format)
                    worksheet4.write("F3", "Met Target % scFV Flow(Y/N) (≥2%)", bold_12_wrap_format)
                    worksheet4.write("G3", "%scFV (Il13Rα2)", bold_12_wrap_format)
                    worksheet4.merge_range(
                        "A4:G4",
                        "Cohort 1 (N=" + str(self.infusion_count[0]) + ")",
                        bold_12_wrap_format,
                    )
                    worksheet4.merge_range(
                        "A8:G8",
                        "Cohort 2 (N=" + str(self.infusion_count[1]) + ")",
                        bold_12_wrap_format,
                    )
                    worksheet4.merge_range(
                        "A12:G12",
                        "Cohort -1 (N=" + str(self.infusion_count[2]) + ")",
                        bold_12_wrap_format,
                    )

                    # Merge and format data
                    worksheet4.merge_range("D5:D7", self.infusion_statA.iloc[0, 2], normal_data_format)
                    worksheet4.merge_range("F5:F7", self.infusion_statA.iloc[0, 4], normal_data_format)
                    worksheet4.merge_range("D9:D11", self.infusion_statB.iloc[0, 2], normal_data_format)
                    worksheet4.merge_range("F9:F11", self.infusion_statB.iloc[0, 4], normal_data_format)
                    worksheet4.merge_range("D13:D15", self.infusion_statC.iloc[0, 2], normal_data_format)
                    worksheet4.merge_range("F13:F15", self.infusion_statC.iloc[0, 4], normal_data_format)

                    for i in range(0, len(stat_order)):
                        worksheet4.write(i + 4, 0, stat_order[i], bold_11_format)  # Cohort 1
                    for i in range(0, len(stat_order)):
                        worksheet4.write(i + 8, 0, stat_order[i], bold_11_format)  # Cohort 2
                    for i in range(0, len(stat_order)):
                        worksheet4.write(i + 12, 0, stat_order[i], bold_11_format)  # Cohort -1

                    # * Autofit
                    worksheet4.autofit()

                    ## TODO: DSMB-Infusion Listing
                    worksheet5 = writer.book.add_worksheet("Study Tx Listing")
                    # * WRITING AND FORMATING DATA
                    for i in range(0, len(self.infusion_df)):
                        for j in range(0, len(self.infusion_df.columns)):
                            worksheet5.write(i + 2, j, self.infusion_df.iloc[i, j], normal_data_format)
                    # if there are subjects in infusionR_df
                    if len(self.infusionR_df) > 0:
                        for i in range(0, len(self.infusionR_df)):
                            for j in range(0, len(self.infusionR_df.columns)):
                                worksheet5.write(
                                    i + 2,
                                    j + 15,
                                    self.infusionR_df.iloc[i, j],
                                    normal_data_format,
                                )
                    # * WRITING HEADER AND FORMATTING
                    worksheet5.merge_range("A1:A2", "Subject ID", bold_12_wrap_format)
                    worksheet5.merge_range("B1:B2", "Study Day (Primary)", bold_12_wrap_format)
                    worksheet5.merge_range("C1:C2", "Cohort", bold_12_wrap_format)
                    worksheet5.merge_range("D1:D2", "Target Dose", bold_12_wrap_format)
                    worksheet5.merge_range(
                        "E1:E2",
                        "Study Treatment Date",
                        bold_12_wrap_format,
                    )
                    worksheet5.merge_range("F1:G1", "Product Administration Volumes", bold_12_wrap_format)
                    worksheet5.merge_range("H1:J1", "Cells Administered", bold_12_wrap_format)
                    worksheet5.merge_range("K1:M1", "Transduction Efficiency", bold_12_wrap_format)
                    worksheet5.write("F2", "Volume CSF Withdrawn", bold_12_wrap_format)
                    worksheet5.write(
                        "G2",
                        "Volume Dose Administered",
                        bold_12_wrap_format,
                    )
                    worksheet5.write("H2", "Total Cell Dose", bold_12_wrap_format)
                    worksheet5.write("I2", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                    worksheet5.write("J2", "Met Target Dose (Y/N)", bold_12_wrap_format)
                    worksheet5.write("K2", "%scFV (EGFR)", bold_12_wrap_format)
                    worksheet5.write("L2", "Met Target % scFV Flow (Y/N) (≥2%)", bold_12_wrap_format)
                    worksheet5.write("M2", "%scFV (Il13Rα2)", bold_12_wrap_format)
                    if len(self.infusionR_df) > 0:
                        worksheet5.merge_range("P1:P2", "Subject ID", bold_12_wrap_format)
                        worksheet5.merge_range("Q1:Q2", "Study Day (Retreatment)", bold_12_wrap_format)
                        worksheet5.merge_range("R1:R2", "Cohort", bold_12_wrap_format)
                        worksheet5.merge_range(
                            "S1:S2",
                            "Study Treatment Date",
                            bold_12_wrap_format,
                        )

                        worksheet5.merge_range("T1:U1", "Product Administration Volumes", bold_12_wrap_format)
                        worksheet5.merge_range("V1:W1", "Cells Administered", bold_12_wrap_format)
                        worksheet5.merge_range("X1:Z1", "Transduction Efficiency", bold_12_wrap_format)
                        worksheet5.write(
                            "T2",
                            "Volume CSF Withdrawn",
                            bold_12_wrap_format,
                        )
                        worksheet5.write("U2", "Volume Dose Administered", bold_12_wrap_format)
                        worksheet5.write("V2", "Total Cell Dose", bold_12_wrap_format)
                        worksheet5.write("W2", "CART-EGFR-IL13Rα2 Cell Dose", bold_12_wrap_format)
                        worksheet5.write("X2", "%scFV (EGFR)", bold_12_wrap_format)
                        worksheet5.write("Y2", "Met Target % scFV Flow (Y/N) (≥2%)", bold_12_wrap_format)
                        worksheet5.write("Z2", "%scFV (Il13Rα2)", bold_12_wrap_format)

                    # Autofit
                    worksheet5.autofit()

                    ## TODO: Response Stat for Treated
                    worksheet6 = writer.book.add_worksheet("Response Stat for Treated")
                    # Response Headers
                    worksheet6.merge_range(
                        "A1:B1",
                        "Best Response Reported (N=" + str(self.subject_infused_measurable_count) + ")",
                        bold_12_format,
                    )
                    # Create a format with text wrapping
                    normal_data_format_note = writer.book.add_format({"text_wrap": True, "border": 0})

                    # Merge the range and apply the wrapped format
                    worksheet6.merge_range(
                        "A8:B9",
                        "Note: Sample size includes all treated subjects with measurable disease at baseline and post-treatment response evaluations are available",
                        normal_data_format_note,
                    )

                    # Adjust row height if necessary to make the text visible
                    worksheet6.set_row(8, 9)  # Adjust this number as needed

                    # Listing Response Criteria
                    OOS_Response_Codelist = [
                        "Confirmed CR",
                        "Confirmed PR",
                        "SD",
                        "Confirmed PD",
                        "Pending Confirmation of Response (Not Evaluated)",
                    ]
                    for i in range(0, len(OOS_Response_Codelist)):
                        worksheet6.write(i + 1, 0, OOS_Response_Codelist[i], bold_11_format)

                    for i in range(0, len(self.response_stat_OOS)):
                        for j in range(0, len(self.response_stat_OOS.columns)):
                            worksheet6.write(
                                i + 1,
                                j + 1,
                                self.response_stat_OOS.iloc[i, j],
                                normal_data_format,
                            )

                    worksheet6.autofit()

                    # TODO: Response Listing for Treated

                    worksheet7 = writer.book.add_worksheet("Response Listing for Treated")

                    worksheet7.merge_range("A2:A3", "Subject ID", bold_11_format)
                    worksheet7.merge_range("B2:B3", "Measurable vs. Non-Measurable Disease", bold_11_format)
                    worksheet7.merge_range("C2:D2", "Day 1", bold_11_format)
                    worksheet7.merge_range("E2:F2", "Day 28", bold_11_format)
                    worksheet7.merge_range("G2:H2", "Month 2", bold_11_format)
                    worksheet7.merge_range("I2:J2", "Month 4", bold_11_format)
                    worksheet7.merge_range("K2:L2", "Month 6", bold_11_format)
                    worksheet7.merge_range("M2:N2", "Month 8", bold_11_format)
                    worksheet7.merge_range("O2:P2", "Month 10", bold_11_format)
                    worksheet7.merge_range("Q2:R2", "Month 12", bold_11_format)
                    worksheet7.write("S2", "Unscheduled", bold_11_format)

                    worksheet7.write("C3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("D3", "% Change is SPD", bold_11_format)
                    worksheet7.write("E3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("F3", "% Change is SPD", bold_11_format)
                    worksheet7.write("G3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("H3", "% Change is SPD", bold_11_format)
                    worksheet7.write("I3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("J3", "% Change is SPD", bold_11_format)
                    worksheet7.write("K3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("L3", "% Change is SPD", bold_11_format)
                    worksheet7.write("M3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("N3", "% Change is SPD", bold_11_format)
                    worksheet7.write("O3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("P3", "% Change is SPD", bold_11_format)
                    worksheet7.write("Q3", "Overall Objective Status", bold_11_format)
                    worksheet7.write("R3", "% Change is SPD", bold_11_format)
                    worksheet7.write("S3", "Timepoint: Overall Objective Status/% Change in SPD", bold_11_format)

                    # sort the response_df by Subject
                    self.response_df = self.response_df.sort_values(by=["Subject"])
                    # reset index
                    self.response_df = self.response_df.reset_index(drop=True)
                    for i in range(0, len(self.response_df)):
                        for j in range(0, len(self.response_df.columns)):
                            worksheet7.write(
                                i + 3,
                                j,
                                self.response_df.iloc[i, j],
                                normal_data_format,
                            )

                    self.response_df = self.response_df.replace([np.inf, -np.inf], np.nan)  # Replace INF with NaN
                    unique_subject_list = self.response_df["Subject"].unique()
                    self.subject_prim_count = len(unique_subject_list)
                    for unique_subject in unique_subject_list:
                        count_number = self.response_df["Subject"].value_counts().get(unique_subject, 0)
                        if count_number > 1:
                            # print(f"{unique_subject}: {count_number}")
                            # find the first row of unique_subject within self.response_df and get the index
                            first_row_index = self.response_df[self.response_df["Subject"] == unique_subject].index[0]
                            # for each column within the range of self.response_df -1 (minus the unscheduled column)
                            for column in range(0, len(self.response_df.columns) - 1):
                                # merger the rows starting from the index of the first row all the way to the row with index = first row index + count_number - 1
                                worksheet7.merge_range(
                                    first_row_index + 3,
                                    column,
                                    first_row_index + count_number + 2,
                                    column,
                                    self.response_df.iloc[first_row_index, column],
                                    normal_data_format,
                                )

                    # Autofit
                    worksheet7.autofit()

                    # TODO: Response Listing for Treated
                    worksheet8 = writer.book.add_worksheet("Response Listing for Retreated")

                    self.responseR_df = self.responseR_df.replace([np.inf, -np.inf], np.nan)  # Replace INF with NaN
                    unique_subjectR_list = self.responseR_df["Subject"].unique()
                    self.subject_retx_count = len(unique_subjectR_list)

                    worksheet8.merge_range(
                        "A1:S1",
                        "Disease Response for Retreated Subjects \nN=" + str(self.subject_retx_count),
                        bold_12_format,
                    )

                    worksheet8.merge_range("A2:A3", "Subject ID", bold_11_format)
                    worksheet8.merge_range("B2:B3", "Measurable vs. Non-Measurable Disease", bold_11_format)
                    worksheet8.merge_range("C2:D2", "Day 1-R1", bold_11_format)
                    worksheet8.merge_range("E2:F2", "Day 28-R1", bold_11_format)
                    worksheet8.merge_range("G2:H2", "Month 2-R1", bold_11_format)
                    worksheet8.merge_range("I2:J2", "Month 4-R1", bold_11_format)
                    worksheet8.merge_range("K2:L2", "Month 6-R1", bold_11_format)
                    worksheet8.merge_range("M2:N2", "Month 8-R1", bold_11_format)
                    worksheet8.merge_range("O2:P2", "Month 10-R1", bold_11_format)
                    worksheet8.merge_range("Q2:R2", "Month 12-R1", bold_11_format)
                    worksheet8.write("S2", "Unscheduled-R1", bold_11_format)

                    worksheet8.write("C3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("D3", "% Change is SPD", bold_11_format)
                    worksheet8.write("E3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("F3", "% Change is SPD", bold_11_format)
                    worksheet8.write("G3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("H3", "% Change is SPD", bold_11_format)
                    worksheet8.write("I3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("J3", "% Change is SPD", bold_11_format)
                    worksheet8.write("K3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("L3", "% Change is SPD", bold_11_format)
                    worksheet8.write("M3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("N3", "% Change is SPD", bold_11_format)
                    worksheet8.write("O3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("P3", "% Change is SPD", bold_11_format)
                    worksheet8.write("Q3", "Overall Objective Status", bold_11_format)
                    worksheet8.write("R3", "% Change is SPD", bold_11_format)
                    worksheet8.write("S3", "Timepoint: Overall Objective Status/% Change in SPD", bold_11_format)

                    for i in range(0, len(self.responseR_df)):
                        for j in range(0, len(self.responseR_df.columns)):
                            worksheet8.write(
                                i + 3,
                                j,
                                self.responseR_df.iloc[i, j],
                                normal_data_format,
                            )

                    # Autofit
                    worksheet8.autofit()

                    worksheet9 = writer.book.add_worksheet("Status for Eligible Subjects")
                    # * WRITING AND FORMATING DATA
                    for i in range(0, len(self.status_df)):
                        for j in range(0, len(self.status_df.columns)):
                            worksheet9.write(i + 2, j, self.status_df.iloc[i, j], normal_data_format)

                    # * WRITING HEADER AND FORMATTING
                    worksheet9.merge_range("A1:A2", "Subject ID", bold_12_wrap_format)
                    worksheet9.merge_range("B1:B2", "Cohort", bold_12_wrap_format)
                    worksheet9.merge_range("C1:C2", "Adverse Events (Y/N)", bold_12_wrap_format)
                    worksheet9.merge_range("D1:D2", "Serious Adverse Events (Y/N)", bold_12_wrap_format)
                    worksheet9.merge_range("E1:E2", "Study Status", bold_12_wrap_format)
                    worksheet9.merge_range("F1:F2", "Off-Study Reason", bold_12_wrap_format)
                    worksheet9.merge_range(
                        "G1:G2",
                        "Last Study Visit Performed for Off-Study Subject",
                        bold_12_wrap_format,
                    )

                    # Safety Headers
                    # number of subject of safety_total_df
                    safety_total_df_subject_count = len(self.status_df["Subject"].unique())
                    worksheet9.merge_range(
                        "K1:N1",
                        "Safety Statistics (N=" + str(safety_total_df_subject_count) + ")",
                        bold_12_wrap_format,
                    )
                    worksheet9.merge_range("K2:L2", "Adverse Events", bold_11_format)
                    worksheet9.merge_range("M2:N2", "Serious Adverse Events ", bold_11_format)
                    worksheet9.write("K3", "Yes", bold_11_format)
                    worksheet9.write("L3", "No", bold_11_format)
                    worksheet9.write("M3", "Yes", bold_11_format)
                    worksheet9.write("N3", "No", bold_11_format)
                    worksheet9.write("J4", "Cohort 1", bold_11_format)
                    worksheet9.write("J5", "Cohort 2", bold_11_format)
                    worksheet9.write("J6", "Cohort -1", bold_11_format)

                    # Safety Data
                    # Cohort 1
                    for i in range(0, len(self.safetyCH1_total_df)):
                        for j in range(0, len(self.safetyCH1_total_df.columns)):
                            worksheet9.write(
                                i + 3,
                                j + 10,
                                self.safetyCH1_total_df.iloc[i, j],
                                normal_data_format,
                            )
                    # Cohort 2
                    for i in range(0, len(self.safetyCH2_total_df)):
                        for j in range(0, len(self.safetyCH2_total_df.columns)):
                            worksheet9.write(
                                i + 4,
                                j + 10,
                                self.safetyCH2_total_df.iloc[i, j],
                                normal_data_format,
                            )

                    # Cohort -1
                    for i in range(0, len(self.safetyCHN1_total_df)):
                        for j in range(0, len(self.safetyCHN1_total_df.columns)):
                            worksheet9.write(
                                i + 5,
                                j + 10,
                                self.safetyCHN1_total_df.iloc[i, j],
                                normal_data_format,
                            )

                    # Autofit
                    worksheet9.autofit()
