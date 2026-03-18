#!/usr/bin/env python3
import pandas as pd
from dateutil.relativedelta import *
from util import *


def EnrollmentLog03325(final_data):
    if not final_data["DM"].empty:
        # DM
        DM_df = final_data["DM"][
            [
                "Subject",
                "Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)",
                "Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)",
                "Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)",
                "Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)",
                "Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)",
                "Main Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)",
            ]
        ].copy()
        DM_new_col_name = {
            "Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)": "Race",
            "Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)": "Ethnicity",
            "Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)": "Sex Assigned at Birth",
            "Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)": "Legal Sex",
            "Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)": "Date of Birth",
            "Main Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)": "Main Consent Date",
        }
        DM_df = DM_df.rename(columns=DM_new_col_name)
        DM_df["Main Consent Date"] = pd.to_datetime(DM_df["Main Consent Date"])
        sorted_DM_df = DM_df.sort_values(["Subject"])
        merged_df = sorted_DM_df

        # DSDLA
        DSDLA_df = final_data["DSDLA"][
            ["Subject", "Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)"]
        ].copy()
        DSDLA_new_col_name = {
            "Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)": "Assigned Dose Level"
        }
        DSDLA_df = DSDLA_df.rename(columns=DSDLA_new_col_name)
        merged_df = pd.merge(sorted_DM_df, DSDLA_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Race")
        merged_df.insert(index_reference, "Assigned Dose Level", merged_df.pop("Assigned Dose Level"))

        # DSCA
        DSCA_df = final_data["DSCA"][
            ["Subject", "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)"]
        ].copy()
        DSCA_new_col_name = {"Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)": "Assigned Cohort"}
        DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
        merged_df = pd.merge(merged_df, DSCA_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Assigned Dose Level")
        merged_df.insert(index_reference, "Assigned Cohort", merged_df.pop("Assigned Cohort"))
        # print(merged_df)

        # IE Date of Confirmation of Step #1 Eligibility by PI
        IE_df = final_data["IE"][
            [
                "Subject",
                "Event Group Label",
                "Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)",
            ]
        ].copy()
        IE_df = IE_df[IE_df["Event Group Label"] == "Step #1 Screening/Eligibility"]
        IE_new_col_name = {
            "Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)": "Date of Confirmation of Step #1 Eligibility by PI"
        }
        IE_df = IE_df.rename(columns=IE_new_col_name)
        IE_df = IE_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, IE_df, on="Subject", how="left")
        merged_df["Date of Confirmation of Step #1 Eligibility by PI"] = pd.to_datetime(
            merged_df["Date of Confirmation of Step #1 Eligibility by PI"]
        ).dt.strftime("%m/%d/%Y")

        # convert date columns to datetime type
        merged_df["Date of Birth"] = pd.to_datetime(merged_df["Date of Birth"])

        # Calculate time difference in days and convert to years
        # create a mask for non-NaT values in the two columns
        mask = ~merged_df[["Main Consent Date", "Date of Birth"]].isnull().any(axis=1)

        # apply relativedelta only to rows with non-NaT values in both columns
        merged_df.loc[mask, "Age at Consent"] = merged_df[mask].apply(
            lambda x: relativedelta(x["Main Consent Date"], x["Date of Birth"]).years, axis=1
        )
        merged_df["Main Consent Date"] = pd.to_datetime(merged_df["Main Consent Date"]).dt.strftime("%m/%d/%Y")
        merged_df = merged_df.drop("Date of Birth", axis=1)
        # move the column
        index_reference = merged_df.columns.get_loc("Main Consent Date")
        merged_df.insert(index_reference, "Age at Consent", merged_df.pop("Age at Consent"))

        # IE Date of Confirmation of Step #2 Eligibility by PI
        IE_df = final_data["IE"][
            [
                "Subject",
                "Event Group Label",
                "Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)",
            ]
        ].copy()
        IE_df = IE_df[IE_df["Event Group Label"] == "Step #2 Screening/Eligibility"]
        IE_new_col_name = {
            "Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)": "Date of Confirmation of Step #2 Eligibility by PI"
        }
        IE_df = IE_df.rename(columns=IE_new_col_name)
        IE_df = IE_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, IE_df, on="Subject", how="left")
        merged_df["Date of Confirmation of Step #2 Eligibility by PI"] = pd.to_datetime(
            merged_df["Date of Confirmation of Step #2 Eligibility by PI"]
        ).dt.strftime("%m/%d/%Y")
        # print(merged_df)

        # IE Date of Confirmation of Step #2 Eligibility by Monitoring
        IE_df = final_data["IE"][
            [
                "Subject",
                "Event Group Label",
                "Date of Completion of Monitoring Visit for Eligibility (IG_NS_NA_IE5.DT_NS_YH_ELIGMONDAT)",
            ]
        ].copy()
        IE_df = IE_df[IE_df["Event Group Label"] == "Step #2 Screening/Eligibility"]
        IE_new_col_name = {
            "Date of Completion of Monitoring Visit for Eligibility (IG_NS_NA_IE5.DT_NS_YH_ELIGMONDAT)": "Date of Confirmation of Step #2 Eligibility by Monitoring"
        }
        IE_df = IE_df.rename(columns=IE_new_col_name)
        IE_df = IE_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, IE_df, on="Subject", how="left")
        merged_df["Date of Confirmation of Step #2 Eligibility by Monitoring"] = pd.to_datetime(
            merged_df["Date of Confirmation of Step #2 Eligibility by Monitoring"]
        ).dt.strftime("%m/%d/%Y")
        # print(merged_df)

        # PRAPH
        APH_df = final_data["PRAPH"][
            [
                "Subject",
                "Apheresis Type (IG_NS_NA_PRAPH1.CL_NS_YH_APHTP_cl_NS_APHTP1)",
                "Apheresis Date (IG_NS_NA_PRAPH1.DT_NS_NH_APHDAT)",
            ]
        ].copy()
        APH_new_col_name = {
            "Apheresis Type (IG_NS_NA_PRAPH1.CL_NS_YH_APHTP_cl_NS_APHTP1)": "Apheresis Type (Fresh or Historical)",
            "Apheresis Date (IG_NS_NA_PRAPH1.DT_NS_NH_APHDAT)": "Date of Apheresis Collection",
        }
        APH_df = APH_df.rename(columns=APH_new_col_name)
        merged_df = pd.merge(merged_df, APH_df, on="Subject", how="left")
        merged_df["Date of Apheresis Collection"] = pd.to_datetime(
            merged_df["Date of Apheresis Collection"]
        ).dt.strftime("%m/%d/%Y")
        # print(merged_df)

        # EXINF
        EXINF_df = final_data["EXINF"][
            ["Subject", "Event Group Label", "Study Treatment Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)"]
        ].copy()
        EXINF_df = EXINF_df[EXINF_df["Event Group Label"] == "Study Treatment"]
        EXINF_new_col_name = {
            "Study Treatment Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)": "CAR T cell Administration Date (Day 0)"
        }
        EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
        EXINF_df = EXINF_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXINF_df, on="Subject", how="left")
        merged_df["CAR T cell Administration Date (Day 0)"] = pd.to_datetime(
            merged_df["CAR T cell Administration Date (Day 0)"]
        ).dt.strftime("%m/%d/%Y")
        # print(merged_df)

        # DSINITLF
        INITLF_df = final_data["DSINITLF"][
            [
                "Subject",
                "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)",
                "Last Study Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCPFU_cl_YS_LVCPFU1)",
                "End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)",
            ]
        ].copy()
        INITLF_df = INITLF_df[
            INITLF_df[
                "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)"
            ]
            != "Retreatment"
        ]
        INITLF_new_col_name = {
            "Last Study Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCPFU_cl_YS_LVCPFU1)": "Last Study Visit Completed in Primary Follow-Up",
            "End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)": "Initiation of LTFU Date",
        }
        INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
        INITLF_df = INITLF_df.drop(
            "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)",
            axis=1,
        )
        merged_df = pd.merge(merged_df, INITLF_df, on="Subject", how="left")
        merged_df["Initiation of LTFU Date"] = pd.to_datetime(merged_df["Initiation of LTFU Date"]).dt.strftime(
            "%m/%d/%Y"
        )
        # print(merged_df)

        # DSINITRT
        DSINITRT_df = final_data["DSINITRT"][
            [
                "Subject",
                "Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_YS_LVCPFU1)",
                "Will the Subject receive Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_RTYN_cl_YS_YN1)",
                "From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)",
                "End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)",
                "End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)",
            ]
        ].copy()
        DSINITRT_new_col_name = {
            "Will the Subject receive Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_RTYN_cl_YS_YN1)": "Retreatment?",
            "From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)": "Phase",
            "End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)": "Initiation of Retx Date",
        }
        DSINITRT_df = DSINITRT_df.rename(columns=DSINITRT_new_col_name)
        DSINITRT_df["Initiation of Retx Date"] = DSINITRT_df["Initiation of Retx Date"].fillna(
            DSINITRT_df["End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)"]
        )
        DSINITRT_df = DSINITRT_df.drop(
            "End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)", axis=1
        )
        merged_df = pd.merge(merged_df, DSINITRT_df, on="Subject", how="left")
        # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
        merged_df.loc[merged_df["Phase"] == "Primary Follow-Up", "Initiation of LTFU Date"] = "N/A"
        merged_df["Last Study Visit Completed in Primary Follow-Up"] = merged_df[
            "Last Study Visit Completed in Primary Follow-Up"
        ].fillna(
            merged_df["Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_YS_LVCPFU1)"]
        )
        merged_df = merged_df.drop(
            ["Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_YS_LVCPFU1)", "Phase"],
            axis=1,
        )
        merged_df["Initiation of Retx Date"] = pd.to_datetime(merged_df["Initiation of Retx Date"]).dt.strftime(
            "%m/%d/%Y"
        )

        # EXINF Retx
        EXINF_df = final_data["EXINF"][
            ["Subject", "Event Group Label", "Study Treatment Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)"]
        ].copy()
        EXINF_df = EXINF_df[EXINF_df["Event Group Label"] == "Study Retreatment"]
        EXINF_new_col_name = {
            "Study Treatment Date (IG_NS_NA_EXINF1.DT_NS_NH_INFDAT)": "CAR T Cell Retreatment Date (Day 0-R)"
        }
        EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
        EXINF_df = EXINF_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXINF_df, on="Subject", how="left")
        merged_df["CAR T Cell Retreatment Date (Day 0-R)"] = pd.to_datetime(
            merged_df["CAR T Cell Retreatment Date (Day 0-R)"]
        ).dt.strftime("%m/%d/%Y")
        # print(merged_df)

        # INITLF Retx
        INITLF_df = final_data["DSINITLF"][
            [
                "Subject",
                "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)",
                "Last Study Visit Completed in Retreatment (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCRETX_cl_NS_LVCPFUR1)",
                "End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)",
            ]
        ].copy()
        INITLF_df = INITLF_df[
            INITLF_df[
                "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)"
            ]
            == "Retreatment"
        ]
        INITLF_new_col_name = {
            "Last Study Visit Completed in Retreatment (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCRETX_cl_NS_LVCPFUR1)": "Last Study Visit Completed in Retreatment (X-R)",
            "End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)": "Initiation of Retreatment LTFU Date",
        }
        INITLF_df = INITLF_df.rename(columns=INITLF_new_col_name)
        INITLF_df = INITLF_df.drop(
            "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)",
            axis=1,
        )
        merged_df = pd.merge(merged_df, INITLF_df, on="Subject", how="left")
        merged_df["Initiation of Retreatment LTFU Date"] = pd.to_datetime(
            merged_df["Initiation of Retreatment LTFU Date"]
        ).dt.strftime("%m/%d/%Y")

        # EOS
        EOS_df = final_data["DSEOS"][["Subject", "End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)"]].copy()
        EOS_new_col_name = {"End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)": "End of Study Date"}
        EOS_df = EOS_df.rename(columns=EOS_new_col_name)
        merged_df = pd.merge(merged_df, EOS_df, on="Subject", how="left")
        merged_df["End of Study Date"] = pd.to_datetime(merged_df["End of Study Date"]).dt.strftime("%m/%d/%Y")
        # update headers and fill N/A
        merged_df = merged_df.rename(columns={"Subject": "Subject ID#"})
        # merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()] = merged_df.loc[merged_df['End of Study Date'].notna(), merged_df.columns.tolist()].fillna('N/A')
        # Remove duplicate rows based on both columns
        merged_df.drop_duplicates(inplace=True)
        return merged_df

    else:
        # If no DM data found, return an empty DataFrame with expected columns
        # select and reorder columns
        column_list = [
            "Subject ID#",
            "Assigned Cohort",
            "Assigned Dose Level",
            "Race",
            "Ethnicity",
            "Sex Assigned at Birth",
            "Legal Sex",
            "Age at Consent",
            "Main Consent Date",
            "Date of Confirmation of Step #1 Eligibility by PI",
            "Date of Confirmation of Step #2 Eligibility by PI",
            "Date of Confirmation of Step #2 Eligibility by Monitoring",
            "Apheresis Type (Fresh or Historical)",
            "Date of Apheresis Collection",
            "CAR T cell Administration Date (Day 0)",
            "Last Study Visit Completed in Primary Follow-Up",
            "Initiation of LTFU Date",
            "Retreatment?",
            "Initiation of Retx Date",
            "CAR T Cell Retreatment Date (Day 0-R)",
            "Last Study Visit Completed in Retreatment (X-R)",
            "Initiation of Retreatment LTFU Date",
            "End of Study Date",
        ]
        merged_df = pd.DataFrame()
        # if merged_df does not have all the columns in column_list, add the missing columns
        for col in column_list:
            if col not in merged_df.columns:
                merged_df[col] = np.nan
        merged_df = merged_df[column_list]
        return merged_df
