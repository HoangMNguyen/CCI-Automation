#!/usr/bin/env python3
import pandas as pd
from dateutil.relativedelta import *
from util import *


def EnrollmentLog50425(final_data):
    # DM
    if "DM" in final_data:
        DM_df = final_data["DM"][
            [
                "Subject",
                "Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)",
                "Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)",
                "Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)",
                "Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)",
                "Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)",
            ]
        ].copy()
        DM_new_col_name = {
            "Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)": "Race",
            "Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)": "Ethnicity",
            "Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)": "Sex Assigned at Birth",
            "Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)": "Legal Sex",
            "Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)": "Date of Birth",
        }
        DM_df = DM_df.rename(columns=DM_new_col_name)
        sorted_DM_df = DM_df.sort_values(["Subject"])

    # DSCA
    if "DSCA" in final_data:
        DSCA_df = final_data["DSCA"][
            ["Subject", "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)"]
        ].copy()
        DSCA_new_col_name = {"Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)": "Cohort"}
        DSCA_df = DSCA_df.rename(columns=DSCA_new_col_name)
        merged_df = pd.merge(sorted_DM_df, DSCA_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Race")
        merged_df.insert(index_reference, "Cohort", merged_df.pop("Cohort"))
    else:
        merged_df = sorted_DM_df

    # MHDIAGA
    if "MHDIAGA" in final_data:
        MHDIAGA_df = final_data["MHDIAGA"][
            ["Subject", "Disease Type (IG_NS_NA_MHDIAGA1.CL_NS_YH_RSTYP_cl_NS_RSCAT1)"]
        ].copy()
        MHDIAGA_new_col_name = {
            "Disease Type (IG_NS_NA_MHDIAGA1.CL_NS_YH_RSTYP_cl_NS_RSCAT1)": "Disease Type at Study Enrollment Cohort A"
        }
        MHDIAGA_df = MHDIAGA_df.rename(columns=MHDIAGA_new_col_name)
        merged_df = pd.merge(merged_df, MHDIAGA_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Race")
        merged_df.insert(index_reference, "Disease Type at Study Enrollment Cohort A", merged_df.pop("Disease Type at Study Enrollment Cohort A"))

    # MHDIAGB
    if "MHDIAGB" in final_data:
        MHDIAGB_df = final_data["MHDIAGB"][
            ["Subject", "Disease Type (IG_NS_NA_MHDIAGB1.CL_NS_YH_RSTYP_cl_NS_RSCAT2)"]
        ].copy()
        MHDIAGB_new_col_name = {
            "Disease Type (IG_NS_NA_MHDIAGB1.CL_NS_YH_RSTYP_cl_NS_RSCAT2)": "Disease Type at Study Enrollment Cohort B"
        }
        MHDIAGB_df = MHDIAGB_df.rename(columns=MHDIAGB_new_col_name)
        merged_df = pd.merge(merged_df, MHDIAGB_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Race")
        merged_df.insert(index_reference, "Disease Type at Study Enrollment Cohort B", merged_df.pop("Disease Type at Study Enrollment Cohort B"))

    if "MHDIAGA" in final_data and "MHDIAGB" in final_data:    
        # Create Disease Type at Study Enrollment based on Cohort
        merged_df["Disease Type at Study Enrollment"] = np.select(
            [
                merged_df["Cohort"].str.contains("Cohort A", na=False),
                merged_df["Cohort"].str.contains("Cohort B", na=False)
            ],
            [
                merged_df["Disease Type at Study Enrollment Cohort A"],
                merged_df["Disease Type at Study Enrollment Cohort B"]
            ],
            default=pd.NA
        )

        # Move the new column before Race
        if "Disease Type at Study Enrollment" in merged_df.columns:
            index_reference = merged_df.columns.get_loc("Race")
            merged_df.insert(
                index_reference,
                "Disease Type at Study Enrollment",
                merged_df.pop("Disease Type at Study Enrollment")
            )

        # remove temporary columns
        merged_df = merged_df.drop(
            columns=[
                "Disease Type at Study Enrollment Cohort A",
                "Disease Type at Study Enrollment Cohort B"
            ],
            errors="ignore"
        )

    # DSDLA
    if "DSDLA" in final_data:
        DSDLA_df = final_data["DSDLA"][
            ["Subject", "CART-45 Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)"]
        ].copy()
        DSDLA_new_col_name = {
            "CART-45 Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)": "Assigned Dose Level"
        }
        DSDLA_df = DSDLA_df.rename(columns=DSDLA_new_col_name)
        merged_df = pd.merge(merged_df, DSDLA_df, on="Subject", how="left")
        index_reference = merged_df.columns.get_loc("Race")
        merged_df.insert(index_reference, "Assigned Dose Level", merged_df.pop("Assigned Dose Level"))

    # IE
    if "IE" in final_data:
        # Subset and rename columns simultaneously
        IE_cols_to_select = {
            "Subject": "Subject",
            "Consent Date (IG_NS_NA_IE1.DT_NS_NH_MAINCDAT)": "Consent Date",
            "Subject Meets All Study Eligibility (IG_NS_NA_IE3.CL_NS_YH_ELIGYN_cl_YS_YN1)": "Subject Meets All Study Eligibility",
            "Date of Eligibility Confirmation by Physician-Investigator (IG_NS_NA_IE5.DT_NS_YH_ELIGPIDAT)": "Date PI Confirmed Eligibility",
            "Date of Completion of Monitoring Visit for Eligibility (IG_NS_NA_IE5.DT_NS_YH_ELIGMONDAT)": "Date of Confirmation of Eligibility by Monitoring",
        }
        IE_df = final_data["IE"][list(IE_cols_to_select.keys())].rename(columns=IE_cols_to_select)
        # Filter Subject Meets All Study Eligibility == Yes
        IE_df = IE_df[IE_df["Subject Meets All Study Eligibility"] == "Yes"].drop(
            "Subject Meets All Study Eligibility", axis=1
        )
        # convert date columsn to datetime type
        IE_df["Date PI Confirmed Eligibility"] = pd.to_datetime(
            IE_df["Date PI Confirmed Eligibility"]
        ).dt.strftime("%m/%d/%Y")
        IE_df["Date of Confirmation of Eligibility by Monitoring"] = pd.to_datetime(
            IE_df["Date of Confirmation of Eligibility by Monitoring"]
        ).dt.strftime("%m/%d/%Y")
        merged_df = pd.merge(merged_df, IE_df, on="Subject", how="left")
        merged_df["Consent Date"] = pd.to_datetime(merged_df["Consent Date"])
        merged_df["Date of Birth"] = pd.to_datetime(merged_df["Date of Birth"])
        # Calculate time difference in days and convert to years
        mask = ~merged_df[["Consent Date", "Date of Birth"]].isnull().any(axis=1)
        # apply relativedelta only to rows with non-NaT values in both columns
        merged_df.loc[mask, "Age at Consent"] = merged_df[mask].apply(
            lambda x: relativedelta(x["Consent Date"], x["Date of Birth"]).years, axis=1
        )
        
        merged_df["Consent Date"] = merged_df["Consent Date"].dt.strftime("%m/%d/%Y")
        merged_df = merged_df.drop("Date of Birth", axis=1)
        # move the column
        index_reference = merged_df.columns.get_loc("Consent Date")
        merged_df.insert(index_reference, "Age at Consent", merged_df.pop("Age at Consent"))

    # APH1
    if "PRAPH1" in final_data:
        PRAPH1_cols_to_select = {
            "Subject": "Subject",
            "Event Group Label": "Event Group Label",
            "Apheresis Date (IG_NS_NA_PRAPH11.DT_YS_NH_APHDAT)": "Date of Apheresis #1 (CART-45 Mfg)",
        }
        PRAPH1_df = final_data["PRAPH1"][list(PRAPH1_cols_to_select.keys())].rename(columns=PRAPH1_cols_to_select)

        # Filter rows and drop unnecessary column
        PRAPH1_df = PRAPH1_df[PRAPH1_df["Event Group Label"] == "Apheresis 1"].drop(
            "Event Group Label", axis=1
        )
        merged_df = pd.merge(merged_df, PRAPH1_df, on="Subject", how="left")
        merged_df["Date of Apheresis #1 (CART-45 Mfg)"] = pd.to_datetime(
            merged_df["Date of Apheresis #1 (CART-45 Mfg)"]
        ).dt.strftime("%m/%d/%Y")

    # APH2
    if "PRAPH2" in final_data:
        PRAPH2_cols_to_select = {
            "Subject": "Subject",
            "Event Group Label": "Event Group Label",
            "Collection Day (IG_NS_NA_PRAPH21.CL_NS_NH_COLDAY_cl_NS_COLDAY1)": "Collection Day",
            "Apheresis Date (IG_NS_NA_PRAPH21.DT_YS_NH_APHDAT)": "Date of Apheresis #2 (CD45BE-HSPC Mfg)",
        }
        PRAPH2_df = final_data["PRAPH2"][list(PRAPH2_cols_to_select.keys())].rename(columns=PRAPH2_cols_to_select)

        # Filter rows and drop unnecessary column
        PRAPH2_df = PRAPH2_df[PRAPH2_df["Event Group Label"] == "Apheresis 2"].drop(
            "Event Group Label", axis=1
        )
        # Filter rows and drop unnecessary column
        PRAPH2_df = PRAPH2_df[PRAPH2_df["Collection Day"] == "Collection Day A"].drop(
            "Collection Day", axis=1
        )
        merged_df = pd.merge(merged_df, PRAPH2_df, on="Subject", how="left")
        merged_df["Date of Apheresis #2 (CD45BE-HSPC Mfg)"] = pd.to_datetime(
            merged_df["Date of Apheresis #2 (CD45BE-HSPC Mfg)"]
        ).dt.strftime("%m/%d/%Y")

    #EXPTRPLCD
    if "EXPTRPLCD" in final_data:
        EXPTRPLCD_df = final_data["EXPTRPLCD"][
            ["Subject", "Event Group Label", "Start Date (IG_NS_NA_EXPTRPLCD2.DT_YS_NH_EXSTDAT)"]
        ].copy()
        EXPTRPLCD_df = EXPTRPLCD_df[EXPTRPLCD_df["Start Date (IG_NS_NA_EXPTRPLCD2.DT_YS_NH_EXSTDAT)"] != "NaN"]
        EXPTRPLCD_df = EXPTRPLCD_df[EXPTRPLCD_df["Event Group Label"] == "CD45BE-HSPC Pre-Transplant Conditioning"]
        EXPTRPLCD_df = EXPTRPLCD_df.drop_duplicates(subset=["Subject"])
        EXPTRPLCD_new_col_name = {"Start Date (IG_NS_NA_EXPTRPLCD2.DT_YS_NH_EXSTDAT)": "Date of Initiation of CD45BE-HSPC Pre-Transplant Conditioning"}
        EXPTRPLCD_df = EXPTRPLCD_df.rename(columns=EXPTRPLCD_new_col_name)
        EXPTRPLCD_df = EXPTRPLCD_df.sort_values(["Subject"])
        EXPTRPLCD_df = EXPTRPLCD_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXPTRPLCD_df, on="Subject", how="left")
        merged_df["Date of Initiation of CD45BE-HSPC Pre-Transplant Conditioning"] = pd.to_datetime(
            merged_df["Date of Initiation of CD45BE-HSPC Pre-Transplant Conditioning"]
        ).dt.strftime("%m/%d/%Y")

    #EXTRPL
    if "EXTRPL" in final_data:
        EXTRPL_df = final_data["EXTRPL"][
            ["Subject", "Event Group Label", "Infusion Date (IG_NS_NA_EXTRPL1.DT_YS_NH_INFDAT)"]
        ].copy()
        EXTRPL_df = EXTRPL_df[EXTRPL_df["Infusion Date (IG_NS_NA_EXTRPL1.DT_YS_NH_INFDAT)"] != "NaN"]
        EXTRPL_df = EXTRPL_df[EXTRPL_df["Event Group Label"] == "CD45BE-HSPC Transplant"]
        EXTRPL_df = EXTRPL_df.drop_duplicates(subset=["Subject"])
        EXTRPL_new_col_name = {"Infusion Date (IG_NS_NA_EXTRPL1.DT_YS_NH_INFDAT)": "CD45BE-HSPC Transplant Date (Day 0)"}
        EXTRPL_df = EXTRPL_df.rename(columns=EXTRPL_new_col_name)
        EXTRPL_df = EXTRPL_df.sort_values(["Subject"])
        EXTRPL_df = EXTRPL_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXTRPL_df, on="Subject", how="left")
        merged_df["CD45BE-HSPC Transplant Date (Day 0)"] = pd.to_datetime(
            merged_df["CD45BE-HSPC Transplant Date (Day 0)"]
        ).dt.strftime("%m/%d/%Y")

    #EXCHMO
    if "EXCHMO" in final_data:
        EXCHMO_df = final_data["EXCHMO"][
            ["Subject", "Event Group Label", "Start Date (IG_NS_NA_EXCHMO2.DT_YS_NH_EXSTDAT)"]
        ].copy()
        EXCHMO_df = EXCHMO_df[EXCHMO_df["Start Date (IG_NS_NA_EXCHMO2.DT_YS_NH_EXSTDAT)"] != "NaN"]
        EXCHMO_df = EXCHMO_df[EXCHMO_df["Event Group Label"] == "Lymphodepleting Chemotherapy"]
        EXCHMO_df = EXCHMO_df.drop_duplicates(subset=["Subject"])
        EXCHMO_new_col_name = {"Start Date (IG_NS_NA_EXCHMO2.DT_YS_NH_EXSTDAT)": "Date of Initiation of LD Chemo"}
        EXCHMO_df = EXCHMO_df.rename(columns=EXCHMO_new_col_name)
        EXCHMO_df = EXCHMO_df.sort_values(["Subject"])
        EXCHMO_df = EXCHMO_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXCHMO_df, on="Subject", how="left")
        merged_df["Date of Initiation of LD Chemo"] = pd.to_datetime(
            merged_df["Date of Initiation of LD Chemo"]
        ).dt.strftime("%m/%d/%Y")

    #INF
    if "EXINF" in final_data:
        EXINF_df = final_data["EXINF"][
            ["Subject", "Event Group Label", "Infusion Date (IG_NS_NA_EXINF1.DT_YS_NH_INFDAT)"]
        ].copy()
        EXINF_df = EXINF_df[EXINF_df["Event Group Label"] == "CART-45 Cell Infusion"]
        EXINF_new_col_name = {"Infusion Date (IG_NS_NA_EXINF1.DT_YS_NH_INFDAT)": "CART-45 Infusion Date"}
        EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
        EXINF_df = EXINF_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXINF_df, on="Subject", how="left")
        merged_df["CART-45 Infusion Date"] = pd.to_datetime(
            merged_df["CART-45 Infusion Date"]
        ).dt.strftime("%m/%d/%Y")

    # DSINITLF
    if "DSINITLF" in final_data:
        DSINITLF_cols_to_select = {
            "Subject": "Subject",
            "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)": "Phase Entering LTFU",
            "Last Study Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCPFU_cl_NS_LVCPFU2)": "Last Study Visit Completed in Primary Follow-Up",
            "End of Primary Follow-Up Date (IG_NS_NA_DSINITLF1.DT_NS_YH_INITLFPFUENDDAT)": "Date Initiation of LTFU",
        }
        DSINITLF_df = final_data["DSINITLF"][list(DSINITLF_cols_to_select.keys())].rename(
            columns=DSINITLF_cols_to_select
        )
    # Filter rows and drop unnecessary column
        DSINITLF_df = DSINITLF_df[DSINITLF_df["Phase Entering LTFU"] == "Primary Follow-Up"].drop(
            "Phase Entering LTFU", axis=1
        )
        # Merge dataframes
        merged_df = pd.merge(merged_df, DSINITLF_df, on="Subject", how="left")
        # Format date
        merged_df["Date Initiation of LTFU"] = pd.to_datetime(merged_df["Date Initiation of LTFU"]).dt.strftime(
            "%m/%d/%Y"
        )

    # DSINITRT
    if "DSINITRT" in final_data:
        DSINITRT_cols_to_select = {
            "Subject": "Subject",
            "Last Visit Completed in Primary Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_NH_RELVCPFU_cl_NS_LVCPFU1)": "Last Visit Completed in PFU",
            "Last Visit Completed in Long-Term Follow-Up (IG_NS_NA_DSINITRT1.CL_NS_YH_LVCLTFU_cl_NS_LVLTFUTP1)": "Last Visit Completed in LTFU",
            "From which Phase is the Subject entering Retreatment? (IG_NS_NA_DSINITRT1.CL_NS_NH_PHASER_cl_NS_PHASE2)": "Phase",
            "End of Primary Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_PFUENDDAT)": "Initiation of Retx Date",
            "End of Long-Term Follow-Up Date (IG_NS_NA_DSINITRT1.DT_NS_YH_LTFUENDDAT)": "End of Long-Term Follow-Up Date",
        }
        DSINITRT_df = final_data["DSINITRT"][list(DSINITRT_cols_to_select.keys())].rename(
            columns=DSINITRT_cols_to_select
        )
        merged_df = pd.merge(merged_df, DSINITRT_df, on="Subject", how="left")
        # if 'Phase' = Primary Follow-Up, convert Initiation of LTFU Date to N/A
        merged_df.loc[merged_df["Phase"] == "Primary Follow-Up", "Date Initiation of LTFU"] = "N/A"
        
        merged_df["Last Study Visit Completed in Primary Follow-Up"] = merged_df[
            "Last Study Visit Completed in Primary Follow-Up"
        ].fillna(merged_df["Last Visit Completed in PFU"])
        merged_df = merged_df.drop("Last Visit Completed in PFU", axis=1)


        merged_df["Initiation of Retx Date"] = merged_df["Initiation of Retx Date"].fillna(
            merged_df["End of Long-Term Follow-Up Date"]
        )
        merged_df = merged_df.drop(["End of Long-Term Follow-Up Date", "Phase"], axis=1)
        # drop End of Long-Term Follow-Up Date column
        merged_df["Initiation of Retx Date"] = pd.to_datetime(merged_df["Initiation of Retx Date"]).dt.strftime(
            "%m/%d/%Y"
        )

    # EXCHMO Retreatment
    if "EXCHMO" in final_data:
        # Subset and rename columns simultaneously
        EXCHMO_cols_to_select = {
            "Subject": "Subject",
            "Event Group Label": "Event Group Label",
            "Start Date (IG_NS_NA_EXCHMO2.DT_YS_NH_EXSTDAT)": "Date of Initiation of Retx LD Chemo",
        }
        EXCHMO_df = final_data["EXCHMO"][list(EXCHMO_cols_to_select.keys())].rename(columns=EXCHMO_cols_to_select)

        # Filter rows
        EXCHMO_df = EXCHMO_df[
            (EXCHMO_df["Date of Initiation of Retx LD Chemo"] != "NaN")
            & (EXCHMO_df["Event Group Label"] == "Retreatment Lymphodepleting Chemotherapy")
        ]

        # Drop duplicates and unnecessary columns
        EXCHMO_df = EXCHMO_df.drop_duplicates(subset=["Subject"]).drop("Event Group Label", axis=1)
        # Sort by Subject
        EXCHMO_df = EXCHMO_df.sort_values(["Subject"])
        merged_df = pd.merge(merged_df, EXCHMO_df, on="Subject", how="left")
        merged_df["Date of Initiation of Retx LD Chemo"] = pd.to_datetime(
            merged_df["Date of Initiation of Retx LD Chemo"]
        ).dt.strftime("%m/%d/%Y")

    # EXINF Retx
    if "EXINF" in final_data:
        EXINF_df = final_data["EXINF"][
            ["Subject", "Event Group Label", "Infusion Date (IG_NS_NA_EXINF1.DT_YS_NH_INFDAT)"]
        ].copy()
        EXINF_df = EXINF_df[EXINF_df["Event Group Label"] == "CART-45 Cell Retreatment Infusion"]
        EXINF_new_col_name = {"Infusion Date (IG_NS_NA_EXINF1.DT_YS_NH_INFDAT)": "Date of CART-45 Retx (Day 0-R)"}
        EXINF_df = EXINF_df.rename(columns=EXINF_new_col_name)
        EXINF_df = EXINF_df.drop("Event Group Label", axis=1)
        merged_df = pd.merge(merged_df, EXINF_df, on="Subject", how="left")
        merged_df["Date of CART-45 Retx (Day 0-R)"] = pd.to_datetime(
            merged_df["Date of CART-45 Retx (Day 0-R)"]
        ).dt.strftime("%m/%d/%Y")

    # INITLF Retx
    if "DSINITLF" in final_data:
        # Subset and rename columns simultaneously
        DSINITLF_cols_to_select = {
            "Subject": "Subject",
            "From which Phase is the Subject entering Long-Term Follow-Up? (IG_NS_NA_DSINITLF1.CL_NS_NH_PHASELTFU_cl_NS_PHASE1)": "Phase Entering LTFU",
            "Last Study Visit Completed in Retreatment (IG_NS_NA_DSINITLF1.CL_NS_NH_LVCRETX_cl_NS_LVCPFUR1)": "Last Study Visit Completed in Retreatment",
            "End of Retreatment Date (IG_NS_NA_DSINITLF1.DT_NS_YH_RETXENDDAT)": "Initiation of Retreatment LTFU Date",
        }
        DSINITLF_df = final_data["DSINITLF"][list(DSINITLF_cols_to_select.keys())].rename(
            columns=DSINITLF_cols_to_select
        )

        # Filter rows and drop unnecessary column
        DSINITLF_df = DSINITLF_df[DSINITLF_df["Phase Entering LTFU"] == "Retreatment"].drop(
            "Phase Entering LTFU", axis=1
        )

        # Merge dataframes
        merged_df = pd.merge(merged_df, DSINITLF_df, on="Subject", how="left")

        # Format date
        merged_df["Initiation of Retreatment LTFU Date"] = pd.to_datetime(
            merged_df["Initiation of Retreatment LTFU Date"]
        ).dt.strftime("%m/%d/%Y")

    # EOS
    if "DSEOS" in final_data:
        # Subset and rename columns simultaneously
        DSEOS_cols_to_select = {
            "Subject": "Subject",
            "End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)": "End of Study Date",
            "Last Study Phase (IG_NS_NA_DSEOS1.CL_NS_YH_LSTUDYPS_cl_YS_STUDYPS1)": "Last Study Phase",
            "Last Study Visit Completed in Primary Treatment (IG_NS_NA_DSEOS1.CL_NS_YH_EOSLSVPR_cl_NS_EOSTP1)": "Last Study Visit Completed in Primary Treatment",
            "Last Study Visit Completed in Retreatment (IG_NS_NA_DSEOS1.CL_NS_YH_EOSLSVRE_cl_NS_EOSTP2)": "Last Study Visit Completed in Retreatment",
        }
        DSEOS_df = final_data["DSEOS"][list(DSEOS_cols_to_select.keys())].rename(
            columns=DSEOS_cols_to_select
        )

        # Create unified column
        DSEOS_df["Last Study Visit Completed before End of Study"] = (
        DSEOS_df["Last Study Visit Completed in Primary Treatment"]
        .fillna(DSEOS_df["Last Study Visit Completed in Retreatment"])
        )
        # remove temporary columns
        DSEOS_df = DSEOS_df.drop(
            columns=[
                "Last Study Phase",
                "Last Study Visit Completed in Primary Treatment",
                "Last Study Visit Completed in Retreatment",
            ],
            errors="ignore"
        )
        merged_df = pd.merge(merged_df, DSEOS_df, on="Subject", how="left")
        merged_df["End of Study Date"] = pd.to_datetime(merged_df["End of Study Date"]).dt.strftime("%m/%d/%Y")

        

    # update headers and fill N/A
    merged_df = merged_df.rename(columns={"Subject": "Subject ID#"})
    return merged_df