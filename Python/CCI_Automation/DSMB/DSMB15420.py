#!/usr/bin/env python3
import pandas as pd
import numpy as np
from util import *
from DSMB.DSMB_util import *
from dateutil.relativedelta import *
from datetime import datetime, date
from typing import Optional


def DSMB15420(
    data,
    export=False,
    output_dir="C:/Users/Hoang Nguyen/Dropbox/Current Work/Download",
    output_file_name=datetime.now().strftime("%Y%m%d%H%M%S") + "-15420-DSMB Report",
    debug=False,
):
    # if not data['DM'].empty:
    # Subject
    enrollment_df = data["DM"][["Subject"]].copy()
    enrollment_df = enrollment_df.sort_values(["Subject"])
    # Cohort Assignment
    enrollment_df = add_rename_column_corelisting(
        enrollment_df,
        data,
        "DSCA",
        "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)",
        "Cohort Assignment",
    )
    # Disease
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAG", "Primary Diagnosis of CLL (ig_PRDIAG2.PRDIAGCLL)", "Disease CLL"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAG", "Specify Other Diagnosis (ig_PRDIAG2.PRDIAGCLLOTH)", "Disease CLL2"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAG", "Primary Diagnosis of NHL (ig_PRDIAG3.PRDIAGNHL)", "Disease NHL"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAG", "Specify Other Diagnosis (ig_PRDIAG3.PRDIAGNHLOTH)", "Disease NHL2"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAGALL", "Primary Diagnosis of ALL (ig_PRDIAGALL1.PRDIAGALL)", "Disease ALL"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAGALL", "Specify Other Diagnosis (ig_PRDIAGALL1.PRDIAGALLOTH)", "Disease ALL2"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "PRDIAG", "Primary Diagnosis at Enrollment (ig_PRDIAG5.PRDIAGRICH)", "Disease RICH"
    )
    enrollment_df["Disease"] = None
    # List the columns in the order you want to use them for filling 'Disease'
    columns_to_fill_from = [
        "Disease CLL2",
        "Disease CLL",
        "Disease NHL2",
        "Disease NHL",
        "Disease ALL2",
        "Disease ALL",
        "Disease RICH",
    ]
    # Use fillna() in a loop to fill 'Disease' from the specified columns
    for col in columns_to_fill_from:
        enrollment_df["Disease"] = enrollment_df["Disease"].fillna(enrollment_df[col])
    enrollment_df = enrollment_df.drop(
        columns=[
            "Disease CLL",
            "Disease CLL2",
            "Disease NHL2",
            "Disease NHL",
            "Disease ALL2",
            "Disease ALL",
            "Disease RICH",
        ]
    )
    # Dose Level
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "DLA", "Dose Level Assignment (ig_DLA1.DLADOSELVL)", "Dose Level"
    )
    # Legal Sex
    enrollment_df = add_rename_column_corelisting(enrollment_df, data, "DM", "Legal Sex (ig_DM1.SEX)", "Legal Sex")
    # Sex Assigned at Birth
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "DM", "Sex Assigned at Birth (ig_DM1.BRTHSEX)", "Sex Assigned at Birth"
    )
    # Gender Identity
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "DM", "Gender Identity (ig_DM1.GENDERID2)", "Gender Identity"
    )
    # Age
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "DM", "Date of Birth (ig_DM1.BRTHDAT)", "Date of Birth"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "IE", "Consent Date (ig_IE1.MAINCDAT)", "Consent Date"
    )
    enrollment_df["Consent Date"] = pd.to_datetime(enrollment_df["Consent Date"])
    enrollment_df["Date of Birth"] = pd.to_datetime(enrollment_df["Date of Birth"])
    mask = ~enrollment_df[["Consent Date", "Date of Birth"]].isnull().any(axis=1)
    enrollment_df.loc[mask, "Age"] = enrollment_df[mask].apply(
        lambda x: relativedelta(x["Consent Date"], x["Date of Birth"]).years, axis=1
    )
    enrollment_df = enrollment_df.drop(columns=["Consent Date", "Date of Birth"])
    # Race
    enrollment_df = add_rename_column_corelisting(enrollment_df, data, "DM", "Race (ig_DM1.RACE)", "Race")
    # Subject meets all study eligibility?
    enrollment_df = add_rename_column_corelisting(
        enrollment_df,
        data,
        "IE",
        "Subject Meets All Study Eligibility (ig_IE3.IEYN)",
        "Subject meets all study eligibility?",
    )
    # Reason for Screen Failure
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "IE", "Other Screen Fail Reason (ig_IE4.OTHRSFREAS)", "Reason for Screen Failure"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "IE", "Screen Failure Reason (ig_IE4.IECAT)", "SF1"
    )
    enrollment_df = add_rename_column_corelisting(
        enrollment_df,
        data,
        "IE",
        "Select the primary inclusion criterion excluding this subject  (ig_IE4.ITESTCD)",
        "SF2",
    )
    enrollment_df["SF2"] = enrollment_df[enrollment_df["SF2"].notna()]["SF2"].astype(int)
    enrollment_df = add_rename_column_corelisting(
        enrollment_df,
        data,
        "IE",
        "Select the primary exclusion criterion excluding this subject (ig_IE4.ETESTCD)",
        "SF3",
    )
    enrollment_df["SF3"] = enrollment_df[enrollment_df["SF3"].notna()]["SF3"].astype(int)
    enrollment_df["SF4"] = (
        enrollment_df["SF1"].fillna("").astype(str)
        + " "
        + enrollment_df["SF2"].round().fillna("").astype(str)
        + enrollment_df["SF3"].round().fillna("").astype(str)
    )
    enrollment_df["Reason for Screen Failure"].fillna(enrollment_df["SF4"], inplace=True)
    enrollment_df = enrollment_df.drop(columns=["SF1", "SF2", "SF3", "SF4"])
    # Infused
    enrollment_df = add_rename_column_corelisting(enrollment_df, data, "INF", "Event Group Label", "Event Group Label")
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "INF", "Was Infusion Administered? (ig_INF1.INFOCCUR)", "Infused"
    )
    enrollment_df = enrollment_df[enrollment_df["Event Group Label"] != "Day 0-R"]
    enrollment_df = enrollment_df.drop(columns=["Event Group Label"])
    enrollment_df = add_rename_column_corelisting(
        enrollment_df, data, "EOS", "End of Study Date (ig_EOS1.EOSDAT)", "End of Study Date (ig_EOS1.EOSDAT)"
    )
    # Update 'Infused' column based on the conditions:
    enrollment_df.loc[
        (enrollment_df["Infused"] != "Yes") & (enrollment_df["End of Study Date (ig_EOS1.EOSDAT)"].isnull()), "Infused"
    ] = "Pending"
    enrollment_df.loc[
        (enrollment_df["Infused"] != "Yes") & (~enrollment_df["End of Study Date (ig_EOS1.EOSDAT)"].isnull()), "Infused"
    ] = "No"
    enrollment_df = enrollment_df.drop(columns=["End of Study Date (ig_EOS1.EOSDAT)"])
    enrollment_df = enrollment_df.drop_duplicates()
    final_enrollment_df = enrollment_df.copy()

    ### TODO: Demo Stats Table
    # Calculate Stats of enrollment table
    TT = enrollment_df["Subject"].count()
    TT_df = enrollment_df.copy()
    # Screen Failed
    SF = enrollment_df[enrollment_df["Subject meets all study eligibility?"] == "No"].count()["Subject"]
    SF_df = enrollment_df[enrollment_df["Subject meets all study eligibility?"] == "No"]
    # Cohort A Enrolled & Infused
    CAE = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort A"))
            & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
        ]
        .count()["Subject"]
    )
    CAI = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort A"))
            & (enrollment_df["Infused"] == "Yes")
        ]
        .count()["Subject"]
    )
    CAE_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort A"))
        & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
    ]
    CAI_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort A")) & (enrollment_df["Infused"] == "Yes")
    ]
    # Cohort B Enrolled & Infused
    CBE = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort B"))
            & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
        ]
        .count()["Subject"]
    )
    CBI = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort B"))
            & (enrollment_df["Infused"] == "Yes")
        ]
        .count()["Subject"]
    )
    CBE_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort B"))
        & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
    ]
    CBI_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort B")) & (enrollment_df["Infused"] == "Yes")
    ]
    # Cohort C Enrolled & Infused
    CCE = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort C"))
            & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
        ]
        .count()["Subject"]
    )
    CCI = (
        enrollment_df.fillna("")
        .loc[
            (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort C"))
            & (enrollment_df["Infused"] == "Yes")
        ]
        .count()["Subject"]
    )
    CCE_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort C"))
        & (enrollment_df["Subject meets all study eligibility?"] == "Yes")
    ]
    CCI_df = enrollment_df.fillna("").loc[
        (enrollment_df["Cohort Assignment"].fillna("").str.contains("Cohort C")) & (enrollment_df["Infused"] == "Yes")
    ]

    # Fill NaN with empty string for enrollment_df
    final_enrollment_df = final_enrollment_df.fillna("")

    # Define a dictionary containing the status of each variable
    final_status = {
        "Total Screened": TT,
        "Screen Failed": SF,
        "Cohort A Enrolled": CAE,
        "Cohort A Infused": CAI,
        "Cohort B Enrolled": CBE,
        "Cohort B Infused": CBI,
        "Cohort C Enrolled": CCE,
        "Cohort C Infused": CCI,
    }

    # Create a new dataframe for Legal Sex table with TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df
    final_LegalSex_df = get_stats_percentage("Legal Sex", TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df)

    # Create a new dataframe for Age table with TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df for Age at Consent
    final_Age_df = get_stats_df("Age", TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df)
    final_Age_df = final_Age_df.replace([np.inf, -np.inf], "")
    final_Age_df = final_Age_df.fillna("")

    # Create a new dataframe for Race table with TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df
    final_Race_df = get_stats_percentage("Race", TT_df, SF_df, CAE_df, CAI_df, CBE_df, CBI_df, CCE_df, CCI_df)

    ### TODO: INFUSION LISTING
    # if not data['INF'].empty:

    # adding Target Cell Dose dictionary
    TCD_dict = {
        "Dose Level 1a (DL1a)": 3000000,
        "Dose Level 1b (DL1b)": 3000000,
        "Dose Level 2 (DL2)": 7000000,
        "Dose Level 3 (DL3)": 30000000,
        "Dose Level 4 (DL4)": 70000000,
        "Dose Level 5 (DL5)": 300000000,
        "Not Assigned": "Not Assigned",
    }

    # TODO: PREPARE DATA FOR INFUSION LISTING
    EXCHMO_df = data["EXCHMO"].copy()
    # select unique subject and Event Group Label
    grouped_df = EXCHMO_df.groupby(["Subject", "Event Group Label"])["Medication (ig_EXCHMO2.EXCCAT)"].unique()
    # convert the unique list to string by joining the list with '+' if the list has more than 1 medication
    grouped_df = grouped_df.apply(
        lambda x: " + ".join(str(val) for val in x if pd.notna(val)) if len(x) > 1 else x[0]
    ).reset_index()
    # replace the Event Group Label with Day 0 and Day 0-R
    grouped_df.loc[
        (grouped_df["Event Group Label"] == "Lymphodepleting Chemotherapy")
        | (grouped_df["Event Group Label"] == "Lymphodepleting Chemotherapy - ALL"),
        "Event Group Label",
    ] = "Day 0"
    grouped_df.loc[
        (grouped_df["Event Group Label"] == "Retreatment Lymphodepleting Chemotherapy")
        | (grouped_df["Event Group Label"] == "Retreatment Lymphodepleting Chemotherapy - ALL"),
        "Event Group Label",
    ] = "Day 0-R"
    grouped_df.loc[grouped_df["Subject"] == "100-15420-01", "Event Group Label"] = "Day 0-R"
    # reassign the dataframe to EXCHMO_df with subject, Study Day, and Medication
    EXCHMO_df = grouped_df

    # TODO: INFUSION LISTING Day 0
    # Subject
    infusion_df = data["DM"][["Subject"]].copy()
    infusion_df = infusion_df.sort_values(["Subject"])
    infusion_df = add_rename_column_corelisting(infusion_df, data, "INF", "Event Group Label", "Event Group Label")
    infusion_df = infusion_df[infusion_df["Event Group Label"] != "Day 0-R"]

    # Cohort Assignment
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "DSCA",
        "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)",
        "Cohort Assignment",
    )
    # Dose Level
    infusion_df = add_rename_column_corelisting(
        infusion_df, data, "DLA", "Dose Level Assignment (ig_DLA1.DLADOSELVL)", "Dose Level Assignment"
    )
    # Lymphodepleting Chemotherapy Regimen
    infusion_df = add_rename_column_df(
        infusion_df,
        EXCHMO_df[EXCHMO_df["Event Group Label"] == "Day 0"],
        "EXCHMO",
        "Medication (ig_EXCHMO2.EXCCAT)",
        "Lymphodepleting Chemotherapy Regimen",
    )
    # Fill NaN with
    # Infusion Date
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "Infusion Date (ig_INF1.INFDAT)",
        "Date of huCART19-IL18 Infusion",
        "Subject",
        "Event Group Label",
    )
    # convert the date to datetime object and format it to MM-DD-YYYY
    infusion_df["Date of huCART19-IL18 Infusion"] = infusion_df["Date of huCART19-IL18 Infusion"].apply(
        lambda x: datetime.strptime(x, "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
    )

    # adding Target Cell Dose using TCD_dict
    infusion_df["Target Cell Dose"] = infusion_df["Dose Level Assignment"].map(TCD_dict)

    # Total huCart19-IL18 Cell Dose
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "CAR T Cell Dose Administered (ig_INF1.INFDOS)",
        "Total huCAR T Cell Dose Administered",
        "Subject",
        "Event Group Label",
    )
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "x 10 to the power of (ig_INF1.INFDOSXP)",
        "x 10 to the power of (ig_INF1.INFDOSXP)",
        "Subject",
        "Event Group Label",
    )
    # combine Total huCart19-IL18 Cell Dose and x 10 to the power of (ig_INF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the Total huCart19-IL18 Cell Dose column to string
    infusion_df["Total huCAR T Cell Dose Administered"] = infusion_df["Total huCAR T Cell Dose Administered"].multiply(
        10 ** infusion_df["x 10 to the power of (ig_INF1.INFDOSXP)"]
    )
    infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_INF1.INFDOSXP)"])

    # Total Cell Dose Administered column
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "Total Cell Dose Administered (ig_INF1.INFDOSTOT)",
        "Total Cell Dose Administered",
        "Subject",
        "Event Group Label",
    )
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "x 10 to the power of (ig_INF1.INFDOSTOTXP)",
        "x 10 to the power of (ig_INF1.INFDOSTOTXP)",
        "Subject",
        "Event Group Label",
    )
    infusion_df["Total Cell Dose Administered"] = infusion_df["Total Cell Dose Administered"].multiply(
        10 ** infusion_df["x 10 to the power of (ig_INF1.INFDOSTOTXP)"]
    )
    infusion_df = infusion_df.drop(columns=["x 10 to the power of (ig_INF1.INFDOSTOTXP)"])

    # Adding Met Target Dose column based on the condition of Total Cell Dose Administered and Total huCAR T Cell Dose Administered if 'Target Cell Dose' is integer
    infusion_df["Met Target Dose"] = infusion_df.apply(
        lambda row: "Y"
        if isinstance(row["Target Cell Dose"], int)
        and row["Total huCAR T Cell Dose Administered"] >= row["Target Cell Dose"]
        else "",
        axis=1,
    )
    infusion_df["Met Target Dose"] = infusion_df.apply(
        lambda row: "N"
        if isinstance(row["Target Cell Dose"], int)
        and row["Total huCAR T Cell Dose Administered"] < row["Target Cell Dose"]
        else row["Met Target Dose"],
        axis=1,
    )

    # %scFv Flow
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "Transduction Efficiency (%) (ig_INF1.INFTEFFP)",
        "%scFv Flow",
        "Subject",
        "Event Group Label",
    )

    # adding Met Target %scFv
    infusion_df = add_rename_column_corelisting(
        infusion_df,
        data,
        "INF",
        "Transduction Efficiency (%) (ig_INF1.INFTEFFP)",
        "Met Target %scFv",
        "Subject",
        "Event Group Label",
    )
    # fillter out the rows that have NaN in Met Target %scFv
    infusion_df["Met Target %scFv"] = infusion_df[infusion_df["Met Target %scFv"].notna()]["Met Target %scFv"].apply(
        lambda x: "Y" if x >= 2 else "N"
    )
    # fill NaN with empty string
    infusion_df = infusion_df.fillna("")

    # Only keep the rows that have Event Group Label
    infusion_df = infusion_df[infusion_df["Event Group Label"] != ""]

    # TODO: Infusion Listing Day 0-R
    # prepare EXCHO_df
    EXCHMO_df = data["EXCHMO"].copy()
    # select unique subject and Event Group Label
    grouped_df = EXCHMO_df.groupby(["Subject", "Event Group Label"])["Medication (ig_EXCHMO2.EXCCAT)"].unique()
    # convert the unique list to string
    grouped_df = grouped_df.apply(
        lambda x: "+".join(str(val) for val in x if pd.notna(val)) if len(x) > 1 else x[0]
    ).reset_index()
    # replace the Event Group Label with Day 0 and Day 0-R
    grouped_df.loc[grouped_df["Event Group Label"] == "Lymphodepleting Chemotherapy", "Event Group Label"] = "Day 0"
    grouped_df.loc[
        grouped_df["Event Group Label"] == "Retreatment Lymphodepleting Chemotherapy", "Event Group Label"
    ] = "Day 0-R"
    grouped_df.loc[grouped_df["Subject"] == "15420-01", "Event Group Label"] = "Day 0-R"
    # reassign the dataframe to EXCHMO_df with subject, Study Day, and Medication
    EXCHMO_df = grouped_df

    # Subject
    infusionR_df = data["DM"][["Subject"]].copy()
    infusionR_df = infusionR_df.sort_values(["Subject"])
    infusionR_df = add_rename_column_corelisting(infusionR_df, data, "INF", "Event Group Label", "Event Group Label")
    infusionR_df = infusionR_df[infusionR_df["Event Group Label"] == "Day 0-R"]

    # Cohort Assignment
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "DSCA",
        "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_NH_CACHASCOD_cl_NS_COHORT1)",
        "Cohort Assignment",
    )
    # Lymphodepleting Chemotherapy Regimen
    infusionR_df = add_rename_column_df(
        infusionR_df,
        EXCHMO_df[EXCHMO_df["Event Group Label"] == "Day 0-R"],
        "EXCHMO",
        "Medication (ig_EXCHMO2.EXCCAT)",
        "Lymphodepleting Chemotherapy Regimen",
    )
    # Infusion Date
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "Infusion Date (ig_INF1.INFDAT)",
        "Date of huCART19-IL18 Infusion",
        "Subject",
        "Event Group Label",
    )
    # convert the date to datetime object and format it to MM-DD-YYYY
    infusionR_df["Date of huCART19-IL18 Infusion"] = infusionR_df["Date of huCART19-IL18 Infusion"].apply(
        lambda x: datetime.strptime(x, "%Y-%m-%d").strftime("%m-%d-%Y") if pd.notna(x) else x
    )

    # Total huCart19-IL18 Cell Dose
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "CAR T Cell Dose Administered (ig_INF1.INFDOS)",
        "Total huCAR T Cell Dose Administered",
        "Subject",
        "Event Group Label",
    )
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "x 10 to the power of (ig_INF1.INFDOSXP)",
        "x 10 to the power of (ig_INF1.INFDOSXP)",
        "Subject",
        "Event Group Label",
    )
    # combine Total huCart19-IL18 Cell Dose and x 10 to the power of (ig_INF1.INFDOSXP) columns, compare the new value with 'Target Cell Dose', and convert the Total huCart19-IL18 Cell Dose column to string
    infusionR_df["Total huCAR T Cell Dose Administered"] = infusionR_df[
        "Total huCAR T Cell Dose Administered"
    ].multiply(10 ** infusionR_df["x 10 to the power of (ig_INF1.INFDOSXP)"])
    infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_INF1.INFDOSXP)"])

    # Total Cell Dose Administered column
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "Total Cell Dose Administered (ig_INF1.INFDOSTOT)",
        "Total Cell Dose Administered",
        "Subject",
        "Event Group Label",
    )
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "x 10 to the power of (ig_INF1.INFDOSTOTXP)",
        "x 10 to the power of (ig_INF1.INFDOSTOTXP)",
        "Subject",
        "Event Group Label",
    )
    infusionR_df["Total Cell Dose Administered"] = infusionR_df["Total Cell Dose Administered"].multiply(
        10 ** infusionR_df["x 10 to the power of (ig_INF1.INFDOSTOTXP)"]
    )
    infusionR_df = infusionR_df.drop(columns=["x 10 to the power of (ig_INF1.INFDOSTOTXP)"])

    # %scFv Flow
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "Transduction Efficiency (%) (ig_INF1.INFTEFFP)",
        "%scFv Flow",
        "Subject",
        "Event Group Label",
    )

    # adding Met Target %scFv
    infusionR_df = add_rename_column_corelisting(
        infusionR_df,
        data,
        "INF",
        "Transduction Efficiency (%) (ig_INF1.INFTEFFP)",
        "Met Target %scFv",
        "Subject",
        "Event Group Label",
    )
    # fillter out the rows that have NaN in Met Target %scFv
    infusionR_df["Met Target %scFv"] = infusionR_df[infusionR_df["Met Target %scFv"].notna()]["Met Target %scFv"].apply(
        lambda x: "Y" if x >= 2 else "N"
    )
    # fill NaN with empty string
    infusionR_df = infusionR_df.fillna("")

    # Only keep the rows that have Event Group Label
    infusionR_df = infusionR_df[infusionR_df["Event Group Label"] != ""]

    # TODO: INFUSION STATISTICS
    infusion_count = []
    # * Cohort A
    # Create a new dataframe for Total huCAR T Cell Dose Administered table with infusion_df
    infusionA_df = infusion_df[infusion_df["Cohort Assignment"] == "Cohort A: Non-Hodgkin Lymphoma (NHL)"]
    infusion_statA1 = get_stats_df("Total huCAR T Cell Dose Administered", infusionA_df)
    # Create a new dataframe for Total Cell Dose Administered table with infusion_df
    infusion_statA2 = get_stats_df("Total Cell Dose Administered", infusionA_df)
    # Count the number of subjects that met the target dose
    met_target_count = infusionA_df[infusionA_df["Met Target Dose"] == "Y"].count()["Subject"]
    # Count the number of subjects
    total_subject_count = infusionA_df["Subject"].nunique()
    infusion_statA2["Met Target Dose"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Create a new dataframe for %scFv Flow table with infusion_df
    infusion_statA3 = get_stats_perc_df("%scFv Flow", infusionA_df)
    # Count the number of subjects that met the target %scFv
    met_target_count = infusionA_df[infusionA_df["Met Target %scFv"] == "Y"].count()["Subject"]
    infusion_statA3["Met Target %scFv"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Combine the three dataframes
    infusion_statA = pd.concat([infusion_statA1, infusion_statA2, infusion_statA3], axis=1)
    infusion_statA = infusion_statA.replace([np.inf, -np.inf], "")
    infusion_statA = infusion_statA.fillna("")
    infusion_count.append(total_subject_count)

    # * Cohort B
    # Create a new dataframe for Total huCAR T Cell Dose Administered table with infusion_df
    infusionB_df = infusion_df[infusion_df["Cohort Assignment"] == "Cohort B: Chronic Lymphocytic Leukemia (CLL)"]
    infusion_statB1 = get_stats_df("Total huCAR T Cell Dose Administered", infusionB_df)
    # Create a new dataframe for Total Cell Dose Administered table with infusion_df
    infusion_statB2 = get_stats_df("Total Cell Dose Administered", infusionB_df)
    # Count the number of subjects that met the target dose
    met_target_count = infusionB_df[infusionB_df["Met Target Dose"] == "Y"].count()["Subject"]
    # Count the number of subjects
    total_subject_count = infusionB_df["Subject"].nunique()
    infusion_statB2["Met Target Dose"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Create a new dataframe for %scFv Flow table with infusion_df
    infusion_statB3 = get_stats_perc_df("%scFv Flow", infusionB_df)
    # Count the number of subjects that met the target %scFv
    met_target_count = infusionB_df[infusionB_df["Met Target %scFv"] == "Y"].count()["Subject"]
    infusion_statB3["Met Target %scFv"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Combine the three dataframes
    infusion_statB = pd.concat([infusion_statB1, infusion_statB2, infusion_statB3], axis=1)
    infusion_statB = infusion_statB.replace([np.inf, -np.inf], "")
    infusion_statB = infusion_statB.fillna("")
    infusion_count.append(total_subject_count)

    # * Cohort C
    # Create a new dataframe for Total huCAR T Cell Dose Administered table with infusion_df
    infusionC_df = infusion_df[infusion_df["Cohort Assignment"] == "Cohort C: Acute Lymphoblastic Leukemia (ALL)"]
    infusion_statC1 = get_stats_df("Total huCAR T Cell Dose Administered", infusionC_df)
    # Create a new dataframe for Total Cell Dose Administered table with infusion_df
    infusion_statC2 = get_stats_df("Total Cell Dose Administered", infusionC_df)
    # Count the number of subjects that met the target dose
    met_target_count = infusionC_df[infusionC_df["Met Target Dose"] == "Y"].count()["Subject"]
    # Count the number of subjects
    total_subject_count = infusionC_df["Subject"].nunique()
    infusion_statC2["Met Target Dose"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Create a new dataframe for %scFv Flow table with infusion_df
    infusion_statC3 = get_stats_perc_df("%scFv Flow", infusionC_df)
    # Count the number of subjects that met the target %scFv
    met_target_count = infusionC_df[infusionC_df["Met Target %scFv"] == "Y"].count()["Subject"]
    infusion_statC3["Met Target %scFv"] = (
        str(met_target_count) + " (" + str(round(met_target_count / total_subject_count * 100, 2)) + "%)"
    )
    # Combine the three dataframes
    infusion_statC = pd.concat([infusion_statC1, infusion_statC2, infusion_statC3], axis=1)
    infusion_statC = infusion_statC.replace([np.inf, -np.inf], "")
    infusion_statC = infusion_statC.fillna("")
    infusion_count.append(total_subject_count)

    ## TODO: FORMATTING THE DATAFRAME
    # TODO: Day 0
    # Convert the columns to scientific notation if the value is not NaN
    infusion_df["Target Cell Dose"] = infusion_df["Target Cell Dose"].apply(
        lambda x: convert_float_2_sci_notation(x) if isinstance(x, int) and pd.notna(x) else x
    )
    infusion_df["Total huCAR T Cell Dose Administered"] = infusion_df["Total huCAR T Cell Dose Administered"].apply(
        lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    )
    infusion_df["Total Cell Dose Administered"] = infusion_df["Total Cell Dose Administered"].apply(
        lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    )
    # adding '%' sign to %scFv Flow
    infusion_df["%scFv Flow"] = infusion_df.apply(
        lambda row: str(x) + "%" if pd.notna(x := row["%scFv Flow"]) else x, axis=1
    )

    # TODO: Day 0-R
    # Convert the columns to scientific notation if the value is not NaN
    infusionR_df["Total huCAR T Cell Dose Administered"] = infusionR_df["Total huCAR T Cell Dose Administered"].apply(
        lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    )
    infusionR_df["Total Cell Dose Administered"] = infusionR_df["Total Cell Dose Administered"].apply(
        lambda x: convert_float_2_sci_notation(x) if isinstance(x, float) and pd.notna(x) else x
    )
    # adding '%' sign to %scFv Flow
    infusionR_df["%scFv Flow"] = infusionR_df.apply(
        lambda row: str(x) + "%" if pd.notna(x := row["%scFv Flow"]) else x, axis=1
    )

    # TODO: ASSIGNING THE DATAFRAME TO THE CLASS VARIABLE
    final_infusion_df = infusion_df
    final_infusionR_df = infusionR_df
    final_infusion_statA = infusion_statA
    final_infusion_statB = infusion_statB
    final_infusion_statC = infusion_statC
    final_infusion_count = infusion_count

    # TODO: RESPONSE LISTING
    # if not data['RS'].empty:
    # TODO: PREPARE
    # Disease Response NHL PET based dictionary
    DR_NHL_PET_dict = {
        "Complete Metabolic Response (CMR)": 1,
        "Partial Metabolic Response (PMR)": 2,
        "No Metabolic Response (NMR)": 3,
        "Indeterminate Response (IR)": 4,
        "Progressive Metabolic Disease (PMD)": 5,
        "Not Assessed": 6,
        "Not Evaluable": 7,
        "Not Reported": 10,
    }
    # Disease Response NHL CT based dictionary
    DR_NHL_CT_dict = {
        "Complete Radiologic Response (CR)": 1,
        "Partial Response (PR)": 2,
        "Stable Disease (SD)": 3,
        "Indeterminate Response (IR)": 4,
        "Progressive Disease (PD)": 5,
        "Not Assessed": 6,
        "Not Evaluable": 7,
        "Not Reported": 10,
    }
    # Disease Response CLL Overall dictionary
    DR_CLL_OV_dict = {
        "Complete Remission (CR)": 1,
        "Complete Remission with Incomplete Marrow Recovery (CRi)": 2,
        "Partial Remission (PR)": 3,
        "Stable Disease (SD)": 4,
        "Progressive Disease (PD)": 5,
        "Not Assessed": 6,
        "Not Evaluable": 7,
        "Not Reported": 10,
    }
    # Disease Response CLL Bone Marrow dictionary
    DR_CLL_BM_dict = {
        "Complete Remission (CR)": 1,
        "Partial Remission (PR)": 2,
        "Progressive Disease (PD)": 3,
        "Stable Disease (SD)": 4,
        "Not Assessed": 5,
        "Not Evaluable": 6,
        "Not Reported": 10,
    }
    # Disease Response ALL Overall Resposne dictionary
    DR_ALL_OV_dict = {
        "Complete Remission (CR)": 1,
        "Complete Remission with Incomplete Blood Count Recovery (CRi)": 2,
        "Complete Remission with Residual Mediastinal Disease (CRu)": 3,
        "Treatment Failure (TF)": 4,
        "Relapsed Disease (RD)": 5,
        "Extramedullary Disease Without Bone Marrow Involvement": 6,
        "No Clinical Evidence of Relapse": 7,
        "Unknown/Not Assessed": 8,
        "Not Reported": 9,
        "Not Applicable": 10,
    }
    # Disease Response ALL Extramedullary Disease without Bone Marrow Involvement dictionary
    DR_ALL_ED_dict = {
        "Complete Remission (CR)": 1,
        "Partial Remission (PR)": 2,
        "Stable Disease (SD)": 3,
        "Indeterminate Response": 4,
        "Progressive Disease (PD)": 5,
        "Not Assessed": 6,
        "Not Reported": 9,
        "Not Applicable": 10,
    }
    # Event Label Update dictionary for cohort A and B
    event_AB_dict = {
        "Primary Treatment and Follow-Up": "Primary Treatment",
        "Primary Retreatment and Follow-up": "Primary Retreatment",
        "Pre-Retreatment Safety Visit": "Pre-Retreatment",
        "Long Term Follow-up Months 3-60": "Long Term Follow-up",
        "Retreatment Long Term Follow-up Months 3-60": "Retreatment Long Term Follow-up",
    }
    # Event Label Update dictionary for cohort C
    event_C_dict = {
        "Primary Treatment and Follow-up - ALL": "Primary Treatment",
        "Primary Retreatment and Follow-up - ALL": "Primary Retreatment",
        "Pre-Retreatment Safety Visit - ALL": "Pre-Retreatment",
        "Retreatment Lymphodepleting Chemotherapy - ALL": "Retreatment Lymphodepleting",
        "Long Term Follow-up Months 3-60 - ALL": "Long Term Follow-up",
        "Retreatment Long Term Follow-up Months 3-60 - ALL": "Retreatment Long Term Follow-up",
    }

    # Get data from Initiation of Long Term Follow up
    PD_df = data["INITLF"][
        [
            "Subject",
            "From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)",
            "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)",
            "Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE)",
        ]
    ].copy()
    # Filter the data to only subject with 'Disease progression' in Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE) column
    PD_df = PD_df[
        PD_df["Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE)"]
        == "Disease progression"
    ]
    # Filter the data to subject in Primary Follow up
    PD_df = PD_df[
        PD_df["From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)"]
        == "Primary Follow-Up"
    ]
    # Convert End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) to datetime object
    PD_df["End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"] = pd.to_datetime(
        PD_df["End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"]
    )

    # Get data from Initiation of Long Term Follow up
    PD_Retx_df = data["INITLF"][
        [
            "Subject",
            "From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)",
            "End of Retreatment Date (ig_INITLF1.DSENRETXDAT)",
            "Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE)",
        ]
    ].copy()
    # Filter the data to only subject with 'Disease progression' in Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE) column
    PD_Retx_df = PD_Retx_df[
        PD_Retx_df["Provide reason the Subject is entering into the Long-Term Follow-Up Phase (ig_INITLF2.DSLTFURE)"]
        == "Disease progression"
    ]
    # Filter the data to subject in Retreatment
    PD_Retx_df = PD_Retx_df[
        PD_Retx_df["From which Phase is the Subject entering Long-Term Follow-Up? (ig_INITLF1.DSPHASE)"]
        == "Retreatment"
    ]
    # Convert End of Retreatment Date (ig_INITLF1.DSENRETXDAT) to datetime object
    PD_Retx_df["End of Retreatment Date (ig_INITLF1.DSENRETXDAT)"] = pd.to_datetime(
        PD_Retx_df["End of Retreatment Date (ig_INITLF1.DSENRETXDAT)"]
    )

    # Get data from DSINITRT
    DSINITRT_df = data["DSINITRT"][
        [
            "Subject",
            "From which Phase is the Subject entering Retreatment? (ig_DSINITRT1.DSPHASER)",
            "End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)",
        ]
    ].copy()
    # Convert End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT) to datetime object
    DSINITRT_df["End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)"] = pd.to_datetime(
        DSINITRT_df["End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)"]
    )

    # TODO: RESPONSE LISTING NHL ONLY
    # Response data dataframe for NHL only
    response_df = data["RS"][
        [
            "Subject",
            "Event Group Label",
            "Event Date",
            "Disease Type (ig_RS1.RSCAT)",
            "Study Phase (ig_RS1.STUDYPHS2)",
            "Primary Treatment Time Point (ig_RS1.RSTPT)",
            "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)",
            "Retreatment Time Point (ig_RS1.RSTPTR)",
            "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)",
            "PET-Based Response (ig_RS3.RSIMAGTYP1)",
            "CT-Based Response (ig_RS3.RSIMAGTYP2)",
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
        ]
    ].copy()
    responseA_df = response_df[response_df["Disease Type (ig_RS1.RSCAT)"] == "Non-Hodgkin Lymphoma"]
    responseA_df = responseA_df.sort_values(by=["Subject", "Event Date"])
    # replace Not Assessed with Not Reported in all columns
    responseA_df = responseA_df.replace("Not Assessed", "Not Reported")
    # replace Not Evaluable with Not Reported in all columns
    responseA_df = responseA_df.replace("Not Evaluable", "Not Reported")
    # Filter to Richter's Transformation
    responseBRT_df = response_df[response_df["Disease Type (ig_RS1.RSCAT)"] == "Richter's Transformation"]
    responseBRT_df = responseBRT_df.sort_values(by=["Subject", "Event Date"])
    # replace Not Assessed with Not Reported in all columns
    responseBRT_df = responseBRT_df.replace("Not Assessed", "Not Reported")
    # replace Not Evaluable with Not Reported in all columns
    responseBRT_df = responseBRT_df.replace("Not Evaluable", "Not Reported")

    # TODO: Cohort A - NHL Primary
    # Filter to only Primary Treatment
    responseA_primary_df = responseA_df[responseA_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Primary Treatment"]
    # Replace value of "Unscheduled" in column Primary Treatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseA_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Unscheduled"
    responseA_primary_df = convert_integers_to_strings(
        responseA_primary_df, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    )
    responseA_primary_df.loc[temp_mask, "Primary Treatment Time Point (ig_RS1.RSTPT)"] = (
        "Day "
        + responseA_primary_df.loc[
            temp_mask, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseA_primary_df = responseA_primary_df[
        responseA_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] != "Pre-Treatment Safety Visit"
    ]
    # Convert Event Date to datetime object
    responseA_primary_df["Event Date"] = pd.to_datetime(responseA_primary_df["Event Date"])
    # Snapshot the responseA_primary_df
    responseA_primary_df_snapshot = responseA_primary_df.copy()
    # check the number of subject for cohort B - CLL Retreatment
    final_subject_A_prim_count = len(responseA_primary_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_A_prim_count > 0:
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseA_primary_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseA_primary_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna("Not Reported", inplace=True)
        # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
        responseA_primary_df["PET-Score"] = responseA_primary_df[
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
        ].map(DR_NHL_PET_dict)
        responseA_primary_df["CT-Score"] = responseA_primary_df[
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"
        ].map(DR_NHL_CT_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseA_primary_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseA_primary_current_df = responseA_primary_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseA_primary_current_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseA_primary_current_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best PET-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
        responseA_best_PET_idx = responseA_primary_df.groupby("Subject")["PET-Score"].idxmin()
        # Select these rows for the best PET-based response
        responseA_best_PET_df = responseA_primary_df.loc[responseA_best_PET_idx].copy()
        # Select the columns subject and PET-Based NHL Disease Response from responseA_best_PET_df
        responseA_best_PET_df = responseA_best_PET_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
            ]
        ]
        # Rename the column PET-Based NHL Disease Response to PET-Based Response
        responseA_best_PET_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best PET Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary current response dataframe
        final_responseA_primary_df = pd.merge(
            responseA_primary_current_df, responseA_best_PET_df, on="Subject", how="left"
        )

        ## Best CT-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseA_best_CT_idx = responseA_primary_df.groupby("Subject")["CT-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseA_best_CT_df = responseA_primary_df.loc[responseA_best_CT_idx]
        # Select the columns subject and CT-Based NHL Disease Response from responseA_best_CT_df
        responseA_best_CT_df = responseA_best_CT_df[
            [
                "Subject",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
            ]
        ]
        # Rename the column CT-Based NHL Disease Response to CT-Based Response
        responseA_best_CT_df.rename(
            columns={
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best CT Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary response dataframe
        final_responseA_primary_df = pd.merge(
            final_responseA_primary_df, responseA_best_CT_df, on="Subject", how="left"
        )

        ## Overall NHL Disease Response at Month 3 primary
        # Filter responseA_primary_df to only Month 3
        responseA_primary_M3_df = responseA_df[responseA_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Month 3"]
        # Selec the columns subject and PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) from responseA_primary_M3_df
        responseA_primary_M3_df = responseA_primary_M3_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Event Date",
            ]
        ]
        # Compare responseA_primary_M3_df with responseA_df, and add the subjects (do it once) that are not in responseA_primary_M3_df to responseA_primary_M3_df
        responseA_primary_M3_df = pd.concat(
            [
                responseA_primary_M3_df,
                responseA_df[~responseA_df["Subject"].isin(responseA_primary_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseA_primary_M3_df = responseA_primary_M3_df.drop_duplicates(subset=["Subject"])
        # Copy snapshot of responseA_primary_df to a temporary dataframe
        temp_df = responseA_primary_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        temp_df = temp_df[
            temp_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].notna()
            | temp_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].notna()
        ]
        # Create a for loop that will check the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of each subject in responseA_primary_M3_df
        for index, row in responseA_primary_M3_df.iterrows():
            # check if the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject is nan, and check if the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject is nan
            if pd.isna(row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]) and pd.isna(
                row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
            ):
                # if yes, check to see if the subject in in PD_df
                if row["Subject"] in PD_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in PD_df
                    end_date = PD_df[PD_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"
                    ].values[0]
                    # find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseA_primary_M3_df.loc[index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"] = (
                            temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        )
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseA_primary_M3_df.loc[index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"] = (
                            temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
                        )
                # check if the subject is in Initiation of REtx before month 3
                elif row["Subject"] in DSINITRT_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in DSINITRT_df
                    end_date = DSINITRT_df[DSINITRT_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)"
                    ].values[0]
                    # find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseA_primary_M3_df.loc[index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"] = (
                            temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        )
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseA_primary_M3_df.loc[index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"] = (
                            temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
                        )
        # Rename the column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) to PET-Based ORR
        responseA_primary_M3_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based ORR",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based ORR",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseA_primary_df = pd.merge(
            final_responseA_primary_df, responseA_primary_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Reported" in column PET-Based ORR
        final_responseA_primary_df["PET-Based ORR"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CT-Based ORR
        final_responseA_primary_df["CT-Based ORR"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for NHL primary
        # Getting AE and SAE dataframes
        responseA_primary_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseA_primary_AE_df if the subject of responseA_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_primary_df, else add 'N'
        final_responseA_primary_df["AE"] = final_responseA_primary_df["Subject"].apply(
            lambda x: "Y" if x in responseA_primary_AE_df["Subject"].values else "N"
        )
        # Check responseA_primary_AE_df if the subject of responseA_primary_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseA_primary_df, else add 'N'
        final_responseA_primary_df["SAE"] = final_responseA_primary_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseA_primary_AE_df[responseA_primary_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"]["Subject"].values
            else "N"
        )

        ## Checking Study Status for NHL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseA_primary_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseA_primary_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseA_primary_SV_df = pd.concat([responseA_primary_SV_df, responseA_primary_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseA_primary_SV_df = responseA_primary_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseA_primary_SV_df = responseA_primary_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseA_primary_df = pd.merge(
            final_responseA_primary_df, responseA_primary_SV_df[["Subject", "Event Label"]], on="Subject", how="left"
        )
        # Rename the column Event Label to Event Label (Study Status)
        final_responseA_primary_df["Event Label"] = final_responseA_primary_df["Event Label"].map(event_AB_dict)

        # Select the columns needed only
        final_responseA_primary_df = final_responseA_primary_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
                "PET-Based Response",
                "Best PET Time Point",
                "CT-Based Response",
                "Best CT Time Point",
                "PET-Based ORR",
                "CT-Based ORR",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseA_primary_df = final_responseA_primary_df.replace([np.nan, np.inf, -np.inf], "")
        # * Formatting the dataframe

    # TODO: Cohort A - NHL Retreatment
    # Filter to only Primary Treatment
    responseA_retreatment_df = responseA_df[responseA_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Retreatment"]
    # Replace value of "Unscheduled" in column Retreatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Retreatment  Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseA_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] == "Unscheduled"
    # Convert the column For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR) to string
    responseA_retreatment_df = convert_integers_to_strings(
        responseA_retreatment_df, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
    )
    responseA_retreatment_df.loc[temp_mask, "Retreatment Time Point (ig_RS1.RSTPTR)"] = (
        "Day "
        + responseA_retreatment_df.loc[
            temp_mask, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseA_retreatment_df = responseA_retreatment_df[
        responseA_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] != "Pre-Retreatment Safety Visit"
    ]
    # Convert Event Date to datetime object
    responseA_retreatment_df["Event Date"] = pd.to_datetime(responseA_retreatment_df["Event Date"])
    # Snapshot the responseA_retreatment_df
    responseA_retreatment_df_snapshot = responseA_retreatment_df.copy()
    # check the number of subject for cohort B - CLL Retreatment
    final_subject_A_retx_count = len(responseA_retreatment_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_A_retx_count > 0:
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseA_retreatment_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseA_retreatment_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
        responseA_retreatment_df["PET-Score"] = responseA_retreatment_df[
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
        ].map(DR_NHL_PET_dict)
        responseA_retreatment_df["CT-Score"] = responseA_retreatment_df[
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"
        ].map(DR_NHL_CT_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseA_retreatment_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseA_retreatment_current_df = responseA_retreatment_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseA_retreatment_current_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseA_retreatment_current_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best PET-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
        responseA_best_PET_idx = responseA_retreatment_df.groupby("Subject")["PET-Score"].idxmin()
        # Select these rows for the best PET-based response
        responseA_best_PET_df = responseA_retreatment_df.loc[responseA_best_PET_idx].copy()
        # Select the columns subject and PET-Based NHL Disease Response from responseA_best_PET_df
        responseA_best_PET_df = responseA_best_PET_df[
            ["Subject", "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)", "Retreatment Time Point (ig_RS1.RSTPTR)"]
        ]
        # Rename the column PET-Based NHL Disease Response to PET-Based Response
        responseA_best_PET_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based Response",
                "Retreatment Time Point (ig_RS1.RSTPTR)": "Best PET Time Point",
            },
            inplace=True,
        )
        # Fill NaN with "Not Reported" in column PET-Based Response
        responseA_best_PET_df["PET-Based Response"].fillna("Not Reported", inplace=True)
        # Merge left with the primary current response dataframe
        final_responseA_retreatment_df = pd.merge(
            responseA_retreatment_current_df, responseA_best_PET_df, on="Subject", how="left"
        )

        ## Best CT-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseA_best_CT_idx = responseA_retreatment_df.groupby("Subject")["CT-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseA_best_CT_df = responseA_retreatment_df.loc[responseA_best_CT_idx]
        # Select the columns subject and CT-Based NHL Disease Response from responseA_best_CT_df
        responseA_best_CT_df = responseA_best_CT_df[
            ["Subject", "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)", "Retreatment Time Point (ig_RS1.RSTPTR)"]
        ]
        # Rename the column CT-Based NHL Disease Response to CT-Based Response
        responseA_best_CT_df.rename(
            columns={
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based Response",
                "Retreatment Time Point (ig_RS1.RSTPTR)": "Best CT Time Point",
            },
            inplace=True,
        )
        # Fill NaN with "Not Reported" in column CT-Based Response
        responseA_best_CT_df["CT-Based Response"].fillna("Not Reported", inplace=True)
        # Merge left with the primary response dataframe
        final_responseA_retreatment_df = pd.merge(
            final_responseA_retreatment_df, responseA_best_CT_df, on="Subject", how="left"
        )

        ## * Overall NHL Disease Response at Month 3-R
        # Filter responseA_retreatment_df to only Month 3-R
        responseA_retreatment_M3_df = responseA_df[
            responseA_df["Retreatment Time Point (ig_RS1.RSTPTR)"] == "Month 3-R"
        ]
        # Selec the columns subject and PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) from responseA_retreatment_M3_df
        responseA_retreatment_M3_df = responseA_retreatment_M3_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Event Date",
            ]
        ]
        # Compare responseA_retreatment_M3_df with responseA_df, and add the subjects (do it once) that are not in responseA_retreatment_M3_df to responseA_retreatment_M3_df
        responseA_retreatment_M3_df = pd.concat(
            [
                responseA_retreatment_M3_df,
                responseA_df[~responseA_df["Subject"].isin(responseA_retreatment_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseA_retreatment_M3_df = responseA_retreatment_M3_df.drop_duplicates(subset=["Subject"])

        # Copy snapshot of responseA_retreatment_df to a temporary dataframe
        temp_df = responseA_retreatment_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        temp_df = temp_df[
            temp_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].notna()
            | temp_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].notna()
        ]
        # Create a for loop that will check the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of each subject in responseA_retreatment_M3_df
        for index, row in responseA_retreatment_M3_df.iterrows():
            # check if the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject is nan, and check if the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject is nan
            if pd.isna(row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]) and pd.isna(
                row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
            ):
                # if yes, check to see if the subject in in PD_Retx_df
                if row["Subject"] in PD_Retx_df["Subject"].values:
                    # get the End of Retreatment Date (ig_INITLF1.DSENRETXDAT) date of the subject in PD_Retx_df
                    end_date = PD_Retx_df[PD_Retx_df["Subject"] == row["Subject"]][
                        "End of Retreatment Date (ig_INITLF1.DSENRETXDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseA_retreatment_M3_df.loc[
                            index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
                        ] = temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseA_retreatment_M3_df.loc[index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"] = (
                            temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
                        )
        # Rename the column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) to PET-Based ORR
        responseA_retreatment_M3_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based ORR",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based ORR",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseA_retreatment_df = pd.merge(
            final_responseA_retreatment_df, responseA_retreatment_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Reported" in column PET-Based ORR
        final_responseA_retreatment_df["PET-Based ORR"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CT-Based ORR
        final_responseA_retreatment_df["CT-Based ORR"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for NHL primary
        # Getting AE and SAE dataframes
        responseA_retreatment_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseA_retreatment_AE_df if the subject of responseA_retreatment_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_retreatment_df, else add 'N'
        final_responseA_retreatment_df["AE"] = final_responseA_retreatment_df["Subject"].apply(
            lambda x: "Y" if x in responseA_retreatment_AE_df["Subject"].values else "N"
        )
        # Check responseA_retreatment_AE_df if the subject of responseA_retreatment_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseA_retreatment_df, else add 'N'
        final_responseA_retreatment_df["SAE"] = final_responseA_retreatment_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseA_retreatment_AE_df[responseA_retreatment_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"][
                "Subject"
            ].values
            else "N"
        )

        ## Checking Study Status for NHL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseA_retreatment_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseA_retreatment_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseA_retreatment_SV_df = pd.concat([responseA_retreatment_SV_df, responseA_retreatment_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseA_retreatment_SV_df = responseA_retreatment_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseA_retreatment_SV_df = responseA_retreatment_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseA_retreatment_df = pd.merge(
            final_responseA_retreatment_df,
            responseA_retreatment_SV_df[["Subject", "Event Label"]],
            on="Subject",
            how="left",
        )

        # * Formatting the dataframe
        # Rename the column Event Label to Event Label (Study Status)
        final_responseA_retreatment_df["Event Label"] = final_responseA_retreatment_df["Event Label"].map(event_AB_dict)

        # Select the columns needed only
        final_responseA_retreatment_df = final_responseA_retreatment_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Retreatment Time Point (ig_RS1.RSTPTR)",
                "PET-Based Response",
                "Best PET Time Point",
                "CT-Based Response",
                "Best CT Time Point",
                "PET-Based ORR",
                "CT-Based ORR",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseA_retreatment_df = final_responseA_retreatment_df.replace([np.nan, np.inf, -np.inf], "")
    # TODO: Cohort B - CLL for Primary
    # TODO: RESPONSE LISTING CLL ONLY
    # Response data dataframe for CLL only
    response_df = data["RS"][
        [
            "Subject",
            "Form Status",
            "Event Group Label",
            "Event Date",
            "Disease Type (ig_RS1.RSCAT)",
            "Study Phase (ig_RS1.STUDYPHS2)",
            "Primary Treatment Time Point (ig_RS1.RSTPT)",
            "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)",
            "Retreatment Time Point (ig_RS1.RSTPTR)",
            "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)",
            "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
            "Disease Response Date (ig_RS2.RSDAT)",
            "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
        ]
    ].copy()
    response_df = response_df[response_df["Form Status"] == "Submitted"]
    # Filter to Chronic Lymphocytic Leukemia
    responseB_df = response_df[response_df["Disease Type (ig_RS1.RSCAT)"] == "Chronic Lymphocytic Leukemia"]
    responseB_df = responseB_df.sort_values(by=["Subject", "Event Date"])
    # Convert Event Date to datetime object
    responseB_df["Event Date"] = pd.to_datetime(responseB_df["Event Date"])
    # Replace Not Assessed with Not Reported in all columns
    responseB_df = responseB_df.replace("Not Assessed", "Not Reported")
    # Replace Not Evaluable with Not Reported in all columns
    responseB_df = responseB_df.replace("Not Evaluable", "Not Reported")

    # TODO: Cohort B - CLL Primary
    # Filter to only Primary Treatment
    responseB_primary_df = responseB_df[responseB_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Primary Treatment"]
    # Remove rows with Pre-Treatment Safety Visit
    responseB_primary_df = responseB_primary_df[
        responseB_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] != "Pre-Treatment Safety Visit"
    ]
    # Snapshot the responseB_primary_df
    responseB_primary_df_snapshot = responseB_primary_df.copy()
    # check the number of subject for cohort B - CLL Retreatment
    final_subject_B_prim_count = len(responseB_primary_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_B_prim_count > 0:
        # Replace value of "Unscheduled" in column Primary Treatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
        temp_mask = responseB_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Unscheduled"
        # Convert the column For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY) to string
        responseB_primary_df = convert_integers_to_strings(
            responseB_primary_df, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
        )
        responseB_primary_df.loc[temp_mask, "Primary Treatment Time Point (ig_RS1.RSTPT)"] = (
            "Day "
            + responseB_primary_df.loc[
                temp_mask, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
            ].astype(str)
        )

        # Fill NaN with "Not Reported" in column Overall CLL Disease Response (ig_RS2.RSCLLCAT)
        responseB_primary_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        responseB_primary_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].fillna("Not Reported", inplace=True)
        # Convert OV-Based CLL Disease Response and CT-Based CLL Disease Response to numeric values
        responseB_primary_df["OV-Score"] = responseB_primary_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].map(
            DR_CLL_OV_dict
        )
        responseB_primary_df["BM-Score"] = responseB_primary_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].map(
            DR_CLL_BM_dict
        )

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseB_primary_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseB_primary_current_df = responseB_primary_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column Overall CLL Disease Response (ig_RS2.RSCLLCAT)
        responseB_primary_current_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        responseB_primary_current_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best OV-Based CLL Disease Response primary
        # Get the indices of the rows with the minimum 'OV-Best' for each 'Subject'
        responseB_best_OV_idx = responseB_primary_df.groupby("Subject")["OV-Score"].idxmin()
        # Select these rows for the best OV-based response
        responseB_best_OV_df = responseB_primary_df.loc[responseB_best_OV_idx].copy()
        # Select the columns subject and OV-Based CLL Disease Response from responseB_best_OV_df
        responseB_best_OV_df = responseB_best_OV_df[
            ["Subject", "Overall CLL Disease Response (ig_RS2.RSCLLCAT)", "Primary Treatment Time Point (ig_RS1.RSTPT)"]
        ]
        # Rename the column OV-Based CLL Disease Response to OV-Based Response
        responseB_best_OV_df.rename(
            columns={
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)": "OV-Best Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best OV Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary current response dataframe
        final_responseB_primary_df = pd.merge(
            responseB_primary_current_df, responseB_best_OV_df, on="Subject", how="left"
        )

        ## Best CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseB_best_CT_idx = responseB_primary_df.groupby("Subject")["BM-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseB_best_CT_df = responseB_primary_df.loc[responseB_best_CT_idx]
        # Select the columns subject and CT-Based CLL Disease Response from responseB_best_CT_df
        responseB_best_CT_df = responseB_best_CT_df[
            ["Subject", "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)", "Primary Treatment Time Point (ig_RS1.RSTPT)"]
        ]
        # Rename the column CT-Based CLL Disease Response to CT-Based Response
        responseB_best_CT_df.rename(
            columns={
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)": "BM-Best Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best BM Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary response dataframe
        final_responseB_primary_df = pd.merge(
            final_responseB_primary_df, responseB_best_CT_df, on="Subject", how="left"
        )

        ## * Overall CLL Disease Response at Month 3 primary
        # Filter responseB_primary_df to only Month 3
        responseB_primary_M3_df = responseB_df[responseB_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Month 3"]
        # Selec the columns subject and Overall CLL Disease Response (ig_RS2.RSCLLCAT) from responseB_primary_M3_df
        responseB_primary_M3_df = responseB_primary_M3_df[
            [
                "Subject",
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
            ]
        ]
        # Compare responseB_primary_M3_df with responseB_df, and add the subjects (do it once) that are not in responseB_primary_M3_df to responseB_primary_M3_df
        responseB_primary_M3_df = pd.concat(
            [
                responseB_primary_M3_df,
                responseB_df[~responseB_df["Subject"].isin(responseB_primary_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseB_primary_M3_df = responseB_primary_M3_df.drop_duplicates(subset=["Subject"])
        # Copy snapshot of responseB_primary_df to a temporary dataframe
        temp_df = responseB_primary_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in Overall CLL Disease Response (ig_RS2.RSCLLCAT) and CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        temp_df = temp_df[
            temp_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].notna()
            | temp_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].notna()
        ]
        # Create a for loop that will check the Overall CLL Disease Response (ig_RS2.RSCLLCAT) and CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of each subject in responseB_primary_M3_df
        for index, row in responseB_primary_M3_df.iterrows():
            # check if the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the subject is nan, and check if the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the subject is nan
            if pd.isna(row["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"]) and pd.isna(
                row["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"]
            ):
                # if yes, check to see if the subject in in PD_df
                if row["Subject"] in PD_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in PD_df
                    end_date = PD_df[PD_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the subject with the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the temp_row
                        responseB_primary_M3_df.loc[index, "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"] = temp_row[
                            "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"
                        ]
                        # Replace the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the subject with the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the temp_row
                        responseB_primary_M3_df.loc[index, "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"] = temp_row[
                            "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"
                        ]
                # check if the subject is in Initiation of REtx before month 3
                elif row["Subject"] in DSINITRT_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) of the subject in DSINITRT_df
                    end_date = DSINITRT_df[DSINITRT_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the subject with the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the temp_row
                        responseB_primary_M3_df.loc[index, "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"] = temp_row[
                            "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"
                        ]
                        # Replace the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the subject with the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the temp_row
                        responseB_primary_M3_df.loc[index, "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"] = temp_row[
                            "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"
                        ]
        # Rename the column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) to OV-Based ORR
        responseB_primary_M3_df.rename(
            columns={
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)": "Overall Response",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)": "Bone Marrow Response",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseB_primary_df = pd.merge(
            final_responseB_primary_df, responseB_primary_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Reported" in column Overall Response
        final_responseB_primary_df["Overall Response"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column Bone Marrow Response
        final_responseB_primary_df["Bone Marrow Response"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for CLL primary
        # Getting AE and SAE dataframes
        responseB_primary_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseB_primary_AE_df if the subject of responseB_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseB_primary_df, else add 'N'
        final_responseB_primary_df["AE"] = final_responseB_primary_df["Subject"].apply(
            lambda x: "Y" if x in responseB_primary_AE_df["Subject"].values else "N"
        )
        # Check responseB_primary_AE_df if the subject of responseB_primary_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseB_primary_df, else add 'N'
        final_responseB_primary_df["SAE"] = final_responseB_primary_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseB_primary_AE_df[responseB_primary_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"]["Subject"].values
            else "N"
        )

        ## Checking Study Status for CLL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseB_primary_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseB_primary_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseB_primary_SV_df = pd.concat([responseB_primary_SV_df, responseB_primary_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseB_primary_SV_df = responseB_primary_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseB_primary_SV_df = responseB_primary_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseB_primary_df = pd.merge(
            final_responseB_primary_df, responseB_primary_SV_df[["Subject", "Event Label"]], on="Subject", how="left"
        )
        # Select the columns needed only
        final_responseB_primary_df = final_responseB_primary_df[
            [
                "Subject",
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
                "OV-Best Response",
                "Best OV Time Point",
                "BM-Best Response",
                "Best BM Time Point",
                "Overall Response",
                "Bone Marrow Response",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseB_primary_df = final_responseB_primary_df.replace([np.nan, np.inf, -np.inf], "")

        # * Formatting the dataframe
        # Rename the column Event Label to Event Label (Study Status)
        final_responseB_primary_df["Event Label"] = final_responseB_primary_df["Event Label"].map(event_AB_dict)

    # TODO: Cohort B - CLL Retreatment
    # Filter to only Primary Treatment
    responseB_retreatment_df = responseB_df[responseB_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Retreatment"]
    # Replace value of "Unscheduled" in column Primary Treatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseB_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] == "Unscheduled"
    # Convert the column For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR) to string
    responseB_retreatment_df = convert_integers_to_strings(
        responseB_retreatment_df, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
    )
    responseB_retreatment_df.loc[temp_mask, "Retreatment Time Point (ig_RS1.RSTPTR)"] = (
        "Day "
        + responseB_retreatment_df.loc[
            temp_mask, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseB_retreatment_df = responseB_retreatment_df[
        responseB_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] != "Pre-Retreatment Safety Visit"
    ]
    # Snapshot the responseB_retreatment_df
    responseB_retreatment_df_snapshot = responseB_retreatment_df.copy()
    # check the number of subject for cohort B - CLL Retreatment
    final_subject_B_retx_count = len(responseB_retreatment_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_B_retx_count > 0:
        # Fill NaN with "Not Reported" in column Overall CLL Disease Response (ig_RS2.RSCLLCAT)
        responseB_retreatment_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        responseB_retreatment_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].fillna("Not Reported", inplace=True)
        # Convert OV-Based CLL Disease Response and CT-Based CLL Disease Response to numeric values
        responseB_retreatment_df["OV-Score"] = responseB_retreatment_df[
            "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"
        ].map(DR_CLL_OV_dict)
        responseB_retreatment_df["BM-Score"] = responseB_retreatment_df[
            "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"
        ].map(DR_CLL_BM_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseB_retreatment_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseB_retreatment_current_df = responseB_retreatment_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column Overall CLL Disease Response (ig_RS2.RSCLLCAT)
        responseB_retreatment_current_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        responseB_retreatment_current_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best OV-Based CLL Disease Response primary
        # Get the indices of the rows with the minimum 'OV-Best' for each 'Subject'
        responseB_best_OV_idx = responseB_retreatment_df.groupby("Subject")["OV-Score"].idxmin()
        # Select these rows for the best OV-based response
        responseB_best_OV_df = responseB_retreatment_df.loc[responseB_best_OV_idx].copy()
        # Select the columns subject and OV-Based CLL Disease Response from responseB_best_OV_df
        responseB_best_OV_df = responseB_best_OV_df[
            ["Subject", "Overall CLL Disease Response (ig_RS2.RSCLLCAT)", "Primary Treatment Time Point (ig_RS1.RSTPT)"]
        ]
        # Rename the column OV-Based CLL Disease Response to OV-Based Response
        responseB_best_OV_df.rename(
            columns={
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)": "OV-Best Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best OV Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary current response dataframe
        final_responseB_retreatment_df = pd.merge(
            responseB_retreatment_current_df, responseB_best_OV_df, on="Subject", how="left"
        )

        ## Best CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseB_best_CT_idx = responseB_retreatment_df.groupby("Subject")["BM-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseB_best_CT_df = responseB_retreatment_df.loc[responseB_best_CT_idx]
        # Select the columns subject and CT-Based CLL Disease Response from responseB_best_CT_df
        responseB_best_CT_df = responseB_best_CT_df[
            ["Subject", "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)", "Primary Treatment Time Point (ig_RS1.RSTPT)"]
        ]
        # Rename the column CT-Based CLL Disease Response to CT-Based Response
        responseB_best_CT_df.rename(
            columns={
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)": "BM-Best Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best BM Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary response dataframe
        final_responseB_retreatment_df = pd.merge(
            final_responseB_retreatment_df, responseB_best_CT_df, on="Subject", how="left"
        )

        ## * Overall CLL Disease Response at Month 3 primary
        # Filter responseB_retreatment_df to only Month 3
        responseB_retreatment_M3_df = responseB_df[
            responseB_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Month 3"
        ]
        # Selec the columns subject and Overall CLL Disease Response (ig_RS2.RSCLLCAT) from responseB_retreatment_M3_df
        responseB_retreatment_M3_df = responseB_retreatment_M3_df[
            [
                "Subject",
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
            ]
        ]
        # Compare responseB_retreatment_M3_df with responseB_df, and add the subjects (do it once) that are not in responseB_retreatment_M3_df to responseB_retreatment_M3_df
        responseB_retreatment_M3_df = pd.concat(
            [
                responseB_retreatment_M3_df,
                responseB_df[~responseB_df["Subject"].isin(responseB_retreatment_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseB_retreatment_M3_df = responseB_retreatment_M3_df.drop_duplicates(subset=["Subject"])
        # Copy snapshot of responseB_retreatment_df to a temporary dataframe
        temp_df = responseB_retreatment_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in Overall CLL Disease Response (ig_RS2.RSCLLCAT) and CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)
        temp_df = temp_df[
            temp_df["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"].notna()
            | temp_df["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"].notna()
        ]
        # Create a for loop that will check the Overall CLL Disease Response (ig_RS2.RSCLLCAT) and CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of each subject in responseB_retreatment_M3_df
        for index, row in responseB_retreatment_M3_df.iterrows():
            # check if the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the subject is nan, and check if the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the subject is nan
            if pd.isna(row["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"]) and pd.isna(
                row["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"]
            ):
                # if yes, check to see if the subject in in PD_Retx_df
                if row["Subject"] in PD_Retx_df["Subject"].values:
                    # get the End of Retreatment Date (ig_INITLF1.DSENRETXDAT) date of the subject in PD_Retx_df
                    end_date = PD_Retx_df[PD_Retx_df["Subject"] == row["Subject"]][
                        "End of Retreatment Date (ig_INITLF1.DSENRETXDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the subject with the Overall CLL Disease Response (ig_RS2.RSCLLCAT) of the temp_row
                        responseB_retreatment_M3_df.loc[index, "Overall CLL Disease Response (ig_RS2.RSCLLCAT)"] = (
                            temp_row["Overall CLL Disease Response (ig_RS2.RSCLLCAT)"]
                        )
                        # Replace the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the subject with the CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) of the temp_row
                        responseB_retreatment_M3_df.loc[index, "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"] = (
                            temp_row["CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)"]
                        )
        # Rename the column CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP) to OV-Based ORR
        responseB_retreatment_M3_df.rename(
            columns={
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)": "Overall Response",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)": "Bone Marrow Response",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseB_retreatment_df = pd.merge(
            final_responseB_retreatment_df, responseB_retreatment_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Reported" in column Overall Response and Bone Marrow Response
        final_responseB_retreatment_df["Overall Response"].fillna("Not Reported", inplace=True)
        final_responseB_retreatment_df["Bone Marrow Response"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for CLL primary
        # Getting AE and SAE dataframes
        responseB_retreatment_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseB_retreatment_AE_df if the subject of responseB_retreatment_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseB_retreatment_df, else add 'N'
        final_responseB_retreatment_df["AE"] = final_responseB_retreatment_df["Subject"].apply(
            lambda x: "Y" if x in responseB_retreatment_AE_df["Subject"].values else "N"
        )
        # Check responseB_retreatment_AE_df if the subject of responseB_retreatment_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseB_retreatment_df, else add 'N'
        final_responseB_retreatment_df["SAE"] = final_responseB_retreatment_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseB_retreatment_AE_df[responseB_retreatment_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"][
                "Subject"
            ].values
            else "N"
        )

        ## Checking Study Status for CLL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseB_retreatment_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseB_retreatment_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseB_retreatment_SV_df = pd.concat([responseB_retreatment_SV_df, responseB_retreatment_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseB_retreatment_SV_df = responseB_retreatment_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseB_retreatment_SV_df = responseB_retreatment_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseB_retreatment_df = pd.merge(
            final_responseB_retreatment_df,
            responseB_retreatment_SV_df[["Subject", "Event Label"]],
            on="Subject",
            how="left",
        )

        # * Formatting the dataframe
        # Rename the column Event Label to Event Label (Study Status)
        final_responseB_retreatment_df["Event Label"] = final_responseB_retreatment_df["Event Label"].map(event_AB_dict)

        # Select the columns needed only
        final_responseB_retreatment_df = final_responseB_retreatment_df[
            [
                "Subject",
                "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
                "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
                "OV-Best Response",
                "Best OV Time Point",
                "BM-Best Response",
                "Best BM Time Point",
                "Overall Response",
                "Bone Marrow Response",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseB_retreatment_df = final_responseB_retreatment_df.replace([np.nan, np.inf, -np.inf], "")

    # TODO: RESPONSE LISTING RICHTER'S TRANSFORMATION ONLY
    # TODO: Cohort B - RESPONSE LISTING RICHTER'S TRANSFORMATION
    # Filter to only Primary Treatment
    responseBRT_primary_df = responseBRT_df[responseBRT_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Primary Treatment"]
    # Replace value of "Unscheduled" in column Primary Treatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseBRT_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Unscheduled"
    # Convert the column For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY) to string
    responseBRT_primary_df = convert_integers_to_strings(
        responseBRT_primary_df, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    )
    responseBRT_primary_df.loc[temp_mask, "Primary Treatment Time Point (ig_RS1.RSTPT)"] = (
        "Day "
        + responseBRT_primary_df.loc[
            temp_mask, "For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseBRT_primary_df = responseBRT_primary_df[
        responseBRT_primary_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] != "Pre-Treatment Safety Visit"
    ]
    # Snapshot the responseBRT_primary_df
    responseBRT_primary_df_snapshot = responseBRT_primary_df.copy()
    # check the number of subject for cohort B - Richter's Transformation
    final_subject_BRT_prim_count = len(responseBRT_primary_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_BRT_prim_count > 0:
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseBRT_primary_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseBRT_primary_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna("Not Reported", inplace=True)
        # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
        responseBRT_primary_df["PET-Score"] = responseBRT_primary_df[
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
        ].map(DR_NHL_PET_dict)
        responseBRT_primary_df["CT-Score"] = responseBRT_primary_df[
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"
        ].map(DR_NHL_CT_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseBRT_primary_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseBRT_primary_current_df = responseBRT_primary_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseBRT_primary_current_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseBRT_primary_current_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best PET-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
        responseBRT_best_PET_idx = responseBRT_primary_df.groupby("Subject")["PET-Score"].idxmin()
        # Select these rows for the best PET-based response
        responseBRT_best_PET_df = responseBRT_primary_df.loc[responseBRT_best_PET_idx].copy()
        # Select the columns subject and PET-Based NHL Disease Response from responseBRT_best_PET_df
        responseBRT_best_PET_df = responseBRT_best_PET_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
            ]
        ]
        # Rename the column PET-Based NHL Disease Response to PET-Based Response
        responseBRT_best_PET_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best PET Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary current response dataframe
        final_responseBRT_primary_df = pd.merge(
            responseBRT_primary_current_df, responseBRT_best_PET_df, on="Subject", how="left"
        )

        ## Best CT-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseBRT_best_CT_idx = responseBRT_primary_df.groupby("Subject")["CT-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseBRT_best_CT_df = responseBRT_primary_df.loc[responseBRT_best_CT_idx]
        # Select the columns subject and CT-Based NHL Disease Response from responseBRT_best_CT_df
        responseBRT_best_CT_df = responseBRT_best_CT_df[
            [
                "Subject",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
            ]
        ]
        # Rename the column CT-Based NHL Disease Response to CT-Based Response
        responseBRT_best_CT_df.rename(
            columns={
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based Response",
                "Primary Treatment Time Point (ig_RS1.RSTPT)": "Best CT Time Point",
            },
            inplace=True,
        )
        # Merge left with the primary response dataframe
        final_responseBRT_primary_df = pd.merge(
            final_responseBRT_primary_df, responseBRT_best_CT_df, on="Subject", how="left"
        )

        ## Overall NHL Disease Response at Month 3 primary
        # Filter responseBRT_primary_df to only Month 3
        responseBRT_primary_M3_df = responseBRT_df[
            responseBRT_df["Primary Treatment Time Point (ig_RS1.RSTPT)"] == "Month 3"
        ]
        # Selec the columns subject and PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) from responseBRT_primary_M3_df
        responseBRT_primary_M3_df = responseBRT_primary_M3_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
            ]
        ]
        # Compare responseBRT_primary_M3_df with responseBRT_df, and add the subjects (do it once) that are not in responseBRT_primary_M3_df to responseBRT_primary_M3_df
        responseBRT_primary_M3_df = pd.concat(
            [
                responseBRT_primary_M3_df,
                responseBRT_df[~responseBRT_df["Subject"].isin(responseBRT_primary_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseBRT_primary_M3_df = responseBRT_primary_M3_df.drop_duplicates(subset=["Subject"])

        # Copy snapshot of responseBRT_primary_df to a temporary dataframe
        temp_df = responseBRT_primary_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        temp_df = temp_df[
            temp_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].notna()
            | temp_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].notna()
        ]
        # Create a for loop that will check the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of each subject in responseBRT_primary_M3_df
        for index, row in responseBRT_primary_M3_df.iterrows():
            # check if the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject is nan, and check if the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject is nan
            if pd.isna(row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]) and pd.isna(
                row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
            ):
                # if yes, check to see if the subject in in PD_df
                if row["Subject"] in PD_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in PD_df
                    end_date = PD_df[PD_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseBRT_primary_M3_df.loc[index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"] = (
                            temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        )
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseBRT_primary_M3_df.loc[index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"] = (
                            temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
                        )
                # check if the subject is in Initiation of REtx before month 3
                elif row["Subject"] in DSINITRT_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in DSINITRT_df
                    end_date = DSINITRT_df[DSINITRT_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseBRT_primary_M3_df.loc[index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"] = (
                            temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        )
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseBRT_primary_M3_df.loc[index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"] = (
                            temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
                        )
        # Rename the column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) to PET-Based ORR
        responseBRT_primary_M3_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based ORR",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based ORR",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseBRT_primary_df = pd.merge(
            final_responseBRT_primary_df, responseBRT_primary_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Reported" in column PET-Based ORR and CT-Based ORR
        final_responseBRT_primary_df["PET-Based ORR"].fillna("Not Reported", inplace=True)
        final_responseBRT_primary_df["CT-Based ORR"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for NHL primary
        # Getting AE and SAE dataframes
        responseBRT_primary_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseBRT_primary_AE_df if the subject of responseBRT_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseBRT_primary_df, else add 'N'
        final_responseBRT_primary_df["AE"] = final_responseBRT_primary_df["Subject"].apply(
            lambda x: "Y" if x in responseBRT_primary_AE_df["Subject"].values else "N"
        )
        # Check responseBRT_primary_AE_df if the subject of responseBRT_primary_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseBRT_primary_df, else add 'N'
        final_responseBRT_primary_df["SAE"] = final_responseBRT_primary_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseBRT_primary_AE_df[responseBRT_primary_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"][
                "Subject"
            ].values
            else "N"
        )

        ## Checking Study Status for NHL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseBRT_primary_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseBRT_primary_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseBRT_primary_SV_df = pd.concat([responseBRT_primary_SV_df, responseBRT_primary_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseBRT_primary_SV_df = responseBRT_primary_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseBRT_primary_SV_df = responseBRT_primary_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseBRT_primary_df = pd.merge(
            final_responseBRT_primary_df,
            responseBRT_primary_SV_df[["Subject", "Event Label"]],
            on="Subject",
            how="left",
        )
        # Rename the column Event Label to Event Label (Study Status)
        final_responseBRT_primary_df["Event Label"] = final_responseBRT_primary_df["Event Label"].map(event_AB_dict)

        # Select the columns needed only
        final_responseBRT_primary_df = final_responseBRT_primary_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Primary Treatment Time Point (ig_RS1.RSTPT)",
                "PET-Based Response",
                "Best PET Time Point",
                "CT-Based Response",
                "Best CT Time Point",
                "PET-Based ORR",
                "CT-Based ORR",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseBRT_primary_df = final_responseBRT_primary_df.replace([np.nan, np.inf, -np.inf], "")
        # * Formatting the dataframe

    # TODO: Cohort B - RT Retreatment
    # Filter to only Primary Treatment
    responseBRT_retreatment_df = responseBRT_df[responseBRT_df["Study Phase (ig_RS1.STUDYPHS2)"] == "Retreatment"]
    # Replace value of "Unscheduled" in column Retreatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Retreatment  Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseBRT_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] == "Unscheduled"
    # Convert the column For Unscheduled Retreatment Time Point, Specify Day #  (ig_RS1.UNSDAYR) to string
    responseBRT_retreatment_df = convert_integers_to_strings(
        responseBRT_retreatment_df, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
    )
    responseBRT_retreatment_df.loc[temp_mask, "Retreatment Time Point (ig_RS1.RSTPTR)"] = (
        "Day "
        + responseBRT_retreatment_df.loc[
            temp_mask, "For Unscheduled Retreatment Time Point, Specify Day # (ig_RS1.UNSDAYR)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseBRT_retreatment_df = responseBRT_retreatment_df[
        responseBRT_retreatment_df["Retreatment Time Point (ig_RS1.RSTPTR)"] != "Pre-Retreatment Safety Visit"
    ]
    # Snapshot the responseBRT_retreatment_df
    responseBRT_retreatment_df_snapshot = responseBRT_retreatment_df.copy()
    # check the number of subject for cohort B - CLL Retreatment
    final_subject_BRT_retx_count = len(responseBRT_retreatment_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_BRT_retx_count > 0:
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseBRT_retreatment_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseBRT_retreatment_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
        responseBRT_retreatment_df["PET-Score"] = responseBRT_retreatment_df[
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
        ].map(DR_NHL_PET_dict)
        responseBRT_retreatment_df["CT-Score"] = responseBRT_retreatment_df[
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"
        ].map(DR_NHL_CT_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseBRT_retreatment_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseBRT_retreatment_current_df = responseBRT_retreatment_df.loc[idx].copy()
        # Fill NaN with "Not Reported" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseBRT_retreatment_current_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].fillna(
            "Not Reported", inplace=True
        )
        # Fill NaN with "Not Reported" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseBRT_retreatment_current_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].fillna(
            "Not Reported", inplace=True
        )

        # * BEST RESPONSE
        ## Best PET-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'PET-Best' for each 'Subject'
        responseBRT_best_PET_idx = responseBRT_retreatment_df.groupby("Subject")["PET-Score"].idxmin()
        # Select these rows for the best PET-based response
        responseBRT_best_PET_df = responseBRT_retreatment_df.loc[responseBRT_best_PET_idx].copy()
        # Select the columns subject and PET-Based NHL Disease Response from responseBRT_best_PET_df
        responseBRT_best_PET_df = responseBRT_best_PET_df[
            ["Subject", "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)", "Retreatment Time Point (ig_RS1.RSTPTR)"]
        ]
        # Rename the column PET-Based NHL Disease Response to PET-Based Response
        responseBRT_best_PET_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based Response",
                "Retreatment Time Point (ig_RS1.RSTPTR)": "Best PET Time Point",
            },
            inplace=True,
        )
        # Fill NaN with "Not Reported" in column PET-Based Response
        responseBRT_best_PET_df["PET-Based Response"].fillna("Not Reported", inplace=True)
        # Merge left with the primary current response dataframe
        final_responseBRT_retreatment_df = pd.merge(
            responseBRT_retreatment_current_df, responseBRT_best_PET_df, on="Subject", how="left"
        )

        ## Best CT-Based NHL Disease Response primary
        # Get the indices of the rows with the minimum 'CT-Best' for each 'Subject'
        responseBRT_best_CT_idx = responseBRT_retreatment_df.groupby("Subject")["CT-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseBRT_best_CT_df = responseBRT_retreatment_df.loc[responseBRT_best_CT_idx]
        # Select the columns subject and CT-Based NHL Disease Response from responseBRT_best_CT_df
        responseBRT_best_CT_df = responseBRT_best_CT_df[
            ["Subject", "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)", "Retreatment Time Point (ig_RS1.RSTPTR)"]
        ]
        # Rename the column CT-Based NHL Disease Response to CT-Based Response
        responseBRT_best_CT_df.rename(
            columns={
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based Response",
                "Retreatment Time Point (ig_RS1.RSTPTR)": "Best CT Time Point",
            },
            inplace=True,
        )
        # Fill NaN with "Not Reported" in column CT-Based Response
        responseBRT_best_CT_df["CT-Based Response"].fillna("Not Reported", inplace=True)
        # Merge left with the primary response dataframe
        final_responseBRT_retreatment_df = pd.merge(
            final_responseBRT_retreatment_df, responseBRT_best_CT_df, on="Subject", how="left"
        )

        ## * Overall NHL Disease Response at Month 3 R
        # Filter responseBRT_retreatment_df to only Month 3
        responseBRT_retreatment_M3_df = responseBRT_df[
            responseBRT_df["Retreatment Time Point (ig_RS1.RSTPTR)"] == "Month 3"
        ]
        # Selec the columns subject and PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) from responseBRT_retreatment_M3_df
        responseBRT_retreatment_M3_df = responseBRT_retreatment_M3_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
            ]
        ]
        # Compare responseBRT_retreatment_M3_df with responseBRT_df, and add the subjects (do it once) that are not in responseBRT_retreatment_M3_df to responseBRT_retreatment_M3_df
        responseBRT_retreatment_M3_df = pd.concat(
            [
                responseBRT_retreatment_M3_df,
                responseBRT_df[~responseBRT_df["Subject"].isin(responseBRT_retreatment_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseBRT_retreatment_M3_df = responseBRT_retreatment_M3_df.drop_duplicates(subset=["Subject"])
        # Copy snapshot of responseBRT_retreatment_df to a temporary dataframe
        temp_df = responseBRT_retreatment_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        temp_df = temp_df[
            temp_df["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"].notna()
            | temp_df["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"].notna()
        ]
        # Create a for loop that will check the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) and CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of each subject in responseBRT_retreatment_M3_df
        for index, row in responseBRT_retreatment_M3_df.iterrows():
            # check if the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject is nan, and check if the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject is nan
            if pd.isna(row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]) and pd.isna(
                row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
            ):
                # if yes, check to see if the subject in in PD_Retx_df
                if row["Subject"] in PD_Retx_df["Subject"].values:
                    # get the End of Retreatment Date (ig_INITLF1.DSENRETXDAT) date of the subject in PD_Retx_df
                    end_date = PD_Retx_df[PD_Retx_df["Subject"] == row["Subject"]][
                        "End of Retreatment Date (ig_INITLF1.DSENRETXDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the subject with the PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) of the temp_row
                        responseBRT_retreatment_M3_df.loc[
                            index, "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"
                        ] = temp_row["PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)"]
                        # Replace the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the subject with the CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT) of the temp_row
                        responseBRT_retreatment_M3_df.loc[
                            index, "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"
                        ] = temp_row["CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)"]
        # Rename the column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT) to PET-Based ORR
        responseBRT_retreatment_M3_df.rename(
            columns={
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)": "PET-Based ORR",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)": "CT-Based ORR",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseBRT_retreatment_df = pd.merge(
            final_responseBRT_retreatment_df, responseBRT_retreatment_M3_df, on="Subject", how="left"
        )
        final_responseBRT_retreatment_df["PET-Based ORR"].fillna("Not Reported", inplace=True)
        # Fill NaN with "Not Reported" in column CT-Based ORR
        final_responseBRT_retreatment_df["CT-Based ORR"].fillna("Not Reported", inplace=True)

        ## Checking AE and SAE for NHL primary
        # Getting AE and SAE dataframes
        responseBRT_retreatment_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseBRT_retreatment_AE_df if the subject of responseBRT_retreatment_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseBRT_retreatment_df, else add 'N'
        final_responseBRT_retreatment_df["AE"] = final_responseBRT_retreatment_df["Subject"].apply(
            lambda x: "Y" if x in responseBRT_retreatment_AE_df["Subject"].values else "N"
        )
        # Check responseBRT_retreatment_AE_df if the subject of responseBRT_retreatment_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseBRT_retreatment_df, else add 'N'
        final_responseBRT_retreatment_df["SAE"] = final_responseBRT_retreatment_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseBRT_retreatment_AE_df[responseBRT_retreatment_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"][
                "Subject"
            ].values
            else "N"
        )

        ## Checking Study Status for NHL primary
        # Getting Study Status dataframe from SV, column Subject, Event Label and Event Date
        responseBRT_retreatment_SV_df = data["SV"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFU, column Subject, Event Label and Event Date
        responseBRT_retreatment_DSSVLTFU_df = data["DSSVLTFU"][["Subject", "Event Label", "Event Date"]]
        # Combine DSSVLTFU with SV dataframe vertically
        responseBRT_retreatment_SV_df = pd.concat([responseBRT_retreatment_SV_df, responseBRT_retreatment_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseBRT_retreatment_SV_df = responseBRT_retreatment_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseBRT_retreatment_SV_df = responseBRT_retreatment_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseBRT_retreatment_df = pd.merge(
            final_responseBRT_retreatment_df,
            responseBRT_retreatment_SV_df[["Subject", "Event Label"]],
            on="Subject",
            how="left",
        )

        # * Formatting the dataframe
        # Rename the column Event Label to Event Label (Study Status)
        final_responseBRT_retreatment_df["Event Label"] = final_responseBRT_retreatment_df["Event Label"].map(
            event_AB_dict
        )

        # Select the columns needed only
        final_responseBRT_retreatment_df = final_responseBRT_retreatment_df[
            [
                "Subject",
                "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
                "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
                "Retreatment Time Point (ig_RS1.RSTPTR)",
                "PET-Based Response",
                "Best PET Time Point",
                "CT-Based Response",
                "Best CT Time Point",
                "PET-Based ORR",
                "CT-Based ORR",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseBRT_retreatment_df = final_responseBRT_retreatment_df.replace([np.nan, np.inf, -np.inf], "")

    # TODO: RESPONSE LISTING ALL ONLY
    # Response data dataframe for NHL only
    responseC_df = data["RSALL"][
        [
            "Subject",
            "Event Date",
            "Study Phase (ig_RSALL1.STUDYPHS2)",
            "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)",
            "For Unscheduled Primary Treatment Time Point, Specify Day # (ig_RSALL1.UNSDAY)",
            "Retreatment Time Point (ig_RSALL1.RSALLTPTR)",
            "For Unscheduled Retreatment Time Point, Specify Day # (ig_RSALL1.UNSDAYR)",
            "Overall Disease Response (ig_RSALL2.RSALLCAT)",
            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)",
        ]
    ].copy()
    responseC_df = responseC_df.sort_values(by=["Subject", "Event Date"])
    # Replace Unknown/Not Assessed with Not Reported in all columns
    responseC_df = responseC_df.replace("Unknown/Not Assessed", "Not Reported")
    # Replace Not Assessed with Not Reported in all columns
    responseC_df = responseC_df.replace("Not Assessed", "Not Reported")

    # TODO: Cohort C - ALL Primary
    # Filter to only Primary Treatment
    responseC_primary_df = responseC_df[responseC_df["Study Phase (ig_RSALL1.STUDYPHS2)"] == "Primary Treatment"]
    # Replace value of "Unscheduled" in column Primary Treatment Time Point (ig_RS1.RSTPT) with value of "Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RS1.RSTPTDY)"
    temp_mask = responseC_primary_df["Primary Treatment Time Point (ig_RSALL1.RSALLTPT)"] == "Unscheduled"
    # Convert the column For Unscheduled Primary Treatment Time Point, Specify Day #  (ig_RSALL1.UNSDAY) to string
    responseC_primary_df = convert_integers_to_strings(
        responseC_primary_df, "For Unscheduled Primary Treatment Time Point, Specify Day # (ig_RSALL1.UNSDAY)"
    )
    responseC_primary_df.loc[temp_mask, "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)"] = (
        "Day "
        + responseC_primary_df.loc[
            temp_mask, "For Unscheduled Primary Treatment Time Point, Specify Day # (ig_RSALL1.UNSDAY)"
        ].astype(str)
    )
    # Remove rows with Pre-Treatment Safety Visit
    responseC_primary_df = responseC_primary_df[
        responseC_primary_df["Primary Treatment Time Point (ig_RSALL1.RSALLTPT)"] != "Pre-Treatment Safety Visit"
    ]
    # Remove rows with Post-Lymphodepleting Chemotherapy
    responseC_primary_df = responseC_primary_df[
        responseC_primary_df["Primary Treatment Time Point (ig_RSALL1.RSALLTPT)"] != "Post-Lymphodepleting Chemotherapy"
    ]
    # replace all rows with Extramedullary Disease Without Bone Marrow Involvement to Not Applicable
    responseC_primary_df["Overall Disease Response (ig_RSALL2.RSALLCAT)"] = responseC_primary_df[
        "Overall Disease Response (ig_RSALL2.RSALLCAT)"
    ].replace("Extramedullary Disease Without Bone Marrow Involvement", "Not Applicable")

    # Snapshot the responseC_primary_df
    responseC_primary_df_snapshot = responseC_primary_df.copy()
    # check the number of subject for cohort C - ALL Primary
    final_subject_C_prim_count = len(responseC_primary_df["Subject"].unique())

    # Check if there is any subject. If yes, then proceed, else skip
    if final_subject_C_prim_count > 0:
        # Fill NaN with "Not Applicable" in column Overall Disease Response (ig_RSALL2.RSALLCAT)
        responseC_primary_df["Overall Disease Response (ig_RSALL2.RSALLCAT)"].fillna("Not Applicable", inplace=True)
        # Fill NaN with "Not Applicable" in column Radiographic Tumor Response for Extramedullary Disease (ig_RSALL2.RADTURESEXDS)
        responseC_primary_df[
            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
        ].fillna("Not Applicable", inplace=True)
        # Convert PET-Based NHL Disease Response and CT-Based NHL Disease Response to numeric values
        responseC_primary_df["OV-Score"] = responseC_primary_df["Overall Disease Response (ig_RSALL2.RSALLCAT)"].map(
            DR_ALL_OV_dict
        )
        responseC_primary_df["ED-Score"] = responseC_primary_df[
            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
        ].map(DR_ALL_ED_dict)

        # * CURRENT RESPONSE
        # Get the indices of the rows with the maximum 'Event Date' for each 'Subject'
        idx = responseC_primary_df.groupby("Subject")["Event Date"].idxmax()
        # Select these rows for the current response
        responseC_primary_current_df = responseC_primary_df.loc[idx].copy()
        # Fill NaN with "Not Applicable" in column PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)
        responseC_primary_current_df["Overall Disease Response (ig_RSALL2.RSALLCAT)"].fillna(
            "Not Applicable", inplace=True
        )
        # Fill NaN with "Not Applicable" in column CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)
        responseC_primary_current_df[
            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
        ].fillna("Not Applicable", inplace=True)
        # rename the column Overall Disease Response (ig_RSALL2.RSALLCAT) to Current Response
        responseC_primary_current_df.rename(
            columns={
                "Overall Disease Response (ig_RSALL2.RSALLCAT)": "Current Overall Response",
                "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)": "Current ED Response",
                "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)": "Current Response Time Point",
            },
            inplace=True,
        )

        # * BEST RESPONSE
        ## Best OV ALL Disease Response primary
        # Get the indices of the rows with the minimum 'OV-Score' for each 'Subject'
        responseC_best_OV_idx = responseC_primary_df.groupby("Subject")["OV-Score"].idxmin()
        # Select these rows for the best OV response
        responseC_best_OV_df = responseC_primary_df.loc[responseC_best_OV_idx].copy()
        # Select the columns subject and OV-Based NHL Disease Response from responseC_best_OV_df
        responseC_best_OV_df = responseC_best_OV_df[
            [
                "Subject",
                "Overall Disease Response (ig_RSALL2.RSALLCAT)",
                "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)",
            ]
        ]
        # Rename the column OV-Based NHL Disease Response to OV-Based Response
        responseC_best_OV_df.rename(
            columns={
                "Overall Disease Response (ig_RSALL2.RSALLCAT)": "Best Overall Response",
                "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)": "Best Overall Response Time Point",
            },
            inplace=True,
        )
        # Replace Nan with Not Applicable
        responseC_best_OV_df["Best Overall Response"].fillna("Not Applicable", inplace=True)
        # Merge left with the primary current response dataframe
        final_responseC_primary_df = pd.merge(
            responseC_primary_current_df, responseC_best_OV_df, on="Subject", how="left"
        )

        ## Best ED ALL Disease Response primary
        # Get the indices of the rows with the minimum 'ED-Score' for each 'Subject'
        responseC_best_ED_idx = responseC_primary_df.groupby("Subject")["ED-Score"].idxmin()
        # Select these rows for the best CT-based response
        responseC_best_ED_df = responseC_primary_df.loc[responseC_best_ED_idx]
        # Select the columns subject and CT-Based NHL Disease Response from responseC_best_CT_df
        responseC_best_ED_df = responseC_best_ED_df[
            [
                "Subject",
                "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)",
                "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)",
            ]
        ]
        # Rename the column CT-Based NHL Disease Response to CT-Based Response
        responseC_best_ED_df.rename(
            columns={
                "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)": "Best ED Response",
                "Primary Treatment Time Point (ig_RSALL1.RSALLTPT)": "Best ED Time Point",
            },
            inplace=True,
        )
        # Replace Nan with Not Applicable
        responseC_best_ED_df["Best ED Response"].fillna("Not Applicable", inplace=True)
        # Merge left with the primary response dataframe
        final_responseC_primary_df = pd.merge(
            final_responseC_primary_df, responseC_best_ED_df, on="Subject", how="left"
        )

        ## * Overall ALL Disease Response at Day 28 primary
        # Filter responseC_primary_df to only Month 3
        responseC_primary_M3_df = responseC_df[
            responseC_df["Primary Treatment Time Point (ig_RSALL1.RSALLTPT)"] == "Day 28"
        ]
        # Selec the columns subject and OV, ED from responseC_primary_M3_df
        responseC_primary_M3_df = responseC_primary_M3_df[
            [
                "Subject",
                "Overall Disease Response (ig_RSALL2.RSALLCAT)",
                "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)",
            ]
        ]
        # Compare responseC_primary_M3_df with responseC_df, and add the subjects (do it once) that are not in responseC_primary_M3_df to responseC_primary_M3_df
        responseC_primary_M3_df = pd.concat(
            [
                responseC_primary_M3_df,
                responseC_df[~responseC_df["Subject"].isin(responseC_primary_M3_df["Subject"])][["Subject"]],
            ]
        )
        # Remove duplicates
        responseC_primary_M3_df = responseC_primary_M3_df.drop_duplicates(subset=["Subject"])
        # Copy snapshot of responseC_primary_df to a temporary dataframe
        temp_df = responseC_primary_df_snapshot
        # Sort the temporary dataframe by Subject and Event Date
        temp_df = temp_df.sort_values(by=["Subject", "Event Date"])
        # remove all the rows that have nan in Overall Disease Response (ig_RSALL2.RSALLCAT) and Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)
        temp_df = temp_df[
            temp_df["Overall Disease Response (ig_RSALL2.RSALLCAT)"].notna()
            | temp_df["Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"].notna()
        ]
        # Create a for loop that will check the Overall Disease Response (ig_RSALL2.RSALLCAT) and Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of each subject in responseC_primary_M3_df
        for index, row in responseC_primary_M3_df.iterrows():
            # check if the Overall Disease Response (ig_RSALL2.RSALLCAT) of the subject is nan, and check if the Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of the subject is nan
            if pd.isna(row["Overall Disease Response (ig_RSALL2.RSALLCAT)"]) and pd.isna(
                row["Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"]
            ):
                # if yes, check to see if the subject in in PD_df
                if row["Subject"] in PD_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT) date of the subject in PD_Primary_df
                    end_date = PD_df[PD_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_INITLF1.DSENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the Overall Disease Response (ig_RSALL2.RSALLCAT) of the subject with the Overall Disease Response (ig_RSALL2.RSALLCAT) of the temp_row
                        responseC_primary_M3_df.loc[index, "Overall Disease Response (ig_RSALL2.RSALLCAT)"] = temp_row[
                            "Overall Disease Response (ig_RSALL2.RSALLCAT)"
                        ]
                        # Replace the Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of the subject with the Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of the temp_row
                        responseC_primary_M3_df.loc[
                            index, "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
                        ] = temp_row[
                            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
                        ]
                # check if the subject is in Initiation of REtx before month 3
                elif row["Subject"] in DSINITRT_df["Subject"].values:
                    # get the End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT) date of the subject in DSINITRT_df
                    end_date = DSINITRT_df[DSINITRT_df["Subject"] == row["Subject"]][
                        "End of Primary Follow-Up Date (ig_DSINITRT1.DSRENPFUDAT)"
                    ].values[0]
                    # if yes, find the response of the same subject with the latest event date that is before the Month 3 event date
                    filtered_df = temp_df[(temp_df["Subject"] == row["Subject"]) & (temp_df["Event Date"] <= end_date)]
                    # Check if the filtered DataFrame is empty
                    if not filtered_df.empty:
                        # Access the last row if the DataFrame is not empty
                        temp_row = filtered_df.iloc[-1]
                        # Replace the Overall Disease Response (ig_RSALL2.RSALLCAT) of the subject with the Overall Disease Response (ig_RSALL2.RSALLCAT) of the temp_row
                        responseC_primary_M3_df.loc[index, "Overall Disease Response (ig_RSALL2.RSALLCAT)"] = temp_row[
                            "Overall Disease Response (ig_RSALL2.RSALLCAT)"
                        ]
                        # Replace the Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of the subject with the Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS) of the temp_row
                        responseC_primary_M3_df.loc[
                            index, "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
                        ] = temp_row[
                            "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)"
                        ]
        # Rename the column Overall Disease Response (ig_RSALL2.RSALLCAT) to OV ORR
        responseC_primary_M3_df.rename(
            columns={
                "Overall Disease Response (ig_RSALL2.RSALLCAT)": "OV ORR",
                "Radiographic Tumor Response for	Extramedullary Disease (ig_RSALL2.RADTURESEXDS)": "ED ORR",
            },
            inplace=True,
        )
        # Merge left with the current response dataframe
        final_responseC_primary_df = pd.merge(
            final_responseC_primary_df, responseC_primary_M3_df, on="Subject", how="left"
        )
        # Fill NaN with "Not Applicable" in column OV ORR
        final_responseC_primary_df["OV ORR"].fillna("Not Applicable", inplace=True)
        # Fill NaN with "Not Applicable" in column ED ORR
        final_responseC_primary_df["ED ORR"].fillna("Not Applicable", inplace=True)
        # Replace Extramedullary Disease Without Bone Marrow Involvement with Not Applicable
        final_responseC_primary_df["OV ORR"] = final_responseC_primary_df["OV ORR"].replace(
            "Extramedullary Disease Without Bone Marrow Involvement", "Not Applicable"
        )

        ## Checking AE and SAE for NHL primary
        # Getting AE and SAE dataframes
        responseC_primary_AE_df = data["AE"][
            ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
        ]
        # Check responseC_primary_AE_df if the subject of responseC_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseC_primary_df, else add 'N'
        final_responseC_primary_df["AE"] = final_responseC_primary_df["Subject"].apply(
            lambda x: "Y" if x in responseC_primary_AE_df["Subject"].values else "N"
        )
        # Check responseC_primary_AE_df if the subject of responseC_primary_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseC_primary_df, else add 'N'
        final_responseC_primary_df["SAE"] = final_responseC_primary_df["Subject"].apply(
            lambda x: "Y"
            if x
            in responseC_primary_AE_df[responseC_primary_AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"]["Subject"].values
            else "N"
        )

        ## Checking Study Status for NHL primary
        # Getting Study Status dataframe from SVALL, column Subject, Event Label and Event Date
        responseC_primary_SV_df = data["SVALL"][["Subject", "Event Label", "Event Date"]]
        # Getting Study Status dataframe from DSSVLTFUALL, column Subject, Event Label and Event Date
        responseC_primary_DSSVLTFU_df = data["DSSVLTFUALL"][["Subject", "Event Label", "Event Date"]]
        # Drop columns that are entirely NA from both DataFrames
        responseC_primary_SV_df = responseC_primary_SV_df.dropna(axis=1, how="all")
        responseC_primary_DSSVLTFU_df = responseC_primary_DSSVLTFU_df.dropna(axis=1, how="all")
        # Combine DSSVLTFU with SV dataframe vertically
        responseC_primary_SV_df = pd.concat([responseC_primary_SV_df, responseC_primary_DSSVLTFU_df])
        # Sort the dataframe by Subject and Event Date
        responseC_primary_SV_df = responseC_primary_SV_df.sort_values(by=["Subject", "Event Date"])
        # For each unique subject, get the last row of the dataframe
        responseC_primary_SV_df = responseC_primary_SV_df.groupby("Subject").tail(1)
        # Merge left with the current response dataframe
        final_responseC_primary_df = pd.merge(
            final_responseC_primary_df, responseC_primary_SV_df[["Subject", "Event Label"]], on="Subject", how="left"
        )
        # Rename the column Event Label to Event Label (Study Status)
        final_responseC_primary_df["Event Label"] = final_responseC_primary_df["Event Label"].map(event_C_dict)

        # Select the columns needed only
        final_responseC_primary_df = final_responseC_primary_df[
            [
                "Subject",
                "Current Overall Response",
                "Current ED Response",
                "Current Response Time Point",
                "Best Overall Response",
                "Best Overall Response Time Point",
                "Best ED Response",
                "Best ED Time Point",
                "OV ORR",
                "ED ORR",
                "AE",
                "SAE",
                "Event Label",
            ]
        ]
        final_responseC_primary_df = final_responseC_primary_df.replace([np.nan, np.inf, -np.inf], "")
        # Count number of subject of column 'Best Overall Response' that is not 'Not Applicable'
        final_subject_C_prim_OR_count = len(
            final_responseC_primary_df[final_responseC_primary_df["Best Overall Response"] != "Not Applicable"][
                "Subject"
            ].unique()
        )
        # Count number of subject of column 'Best ED Response' that is not 'Not Applicable'
        final_subject_C_prim_ED_count = len(
            final_responseC_primary_df[final_responseC_primary_df["Best ED Response"] != "Not Applicable"][
                "Subject"
            ].unique()
        )
    ### TODO: REPONSE STATS
    # TODO: SAFETY STATS

    # Gather all stats of Cohort A, B and C
    total_infused_df = final_infusion_df.copy()
    # Getting AE and SAE dataframes
    AE_df = data["AE"][
        ["Subject", "Form ILB Status", "AE or SAE? (ig_AE2.AESEV)", "Event Onset (ig_AE1.AEONSET)"]
    ].copy()
    # Check responseA_primary_AE_df if the subject of responseA_primary_df is in the AE dataframe. If yes, then add 'Y' to the column 'AE' in responseA_primary_df, else add 'N'
    total_infused_df["AE"] = total_infused_df["Subject"].apply(lambda x: "Y" if x in AE_df["Subject"].values else "N")
    # Check responseA_primary_AE_df if the subject of responseA_primary_df has SAE in column 'AE or SAE? (ig_AE2.AESEV)' . If yes, then add 'Y' to the column 'SAE' in responseA_primary_df, else add 'N'
    total_infused_df["SAE"] = total_infused_df["Subject"].apply(
        lambda x: "Y" if x in AE_df[AE_df["AE or SAE? (ig_AE2.AESEV)"] == "SAE"]["Subject"].values else "N"
    )

    # # Total number of subjects in cohort A, B, and C
    AE_total_count = get_stats_percentage("AE", total_infused_df).T
    SAE_total_count = get_stats_percentage("SAE", total_infused_df).T
    # merge AE and SAE dataframes
    final_safety_total_df = pd.concat([AE_total_count, SAE_total_count], axis=1)

    # # get stats for cohort A, B, and C
    AE_A_count = get_stats_percentage(
        "AE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort A: Non-Hodgkin Lymphoma (NHL)"]
    ).T
    SAE_A_count = get_stats_percentage(
        "SAE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort A: Non-Hodgkin Lymphoma (NHL)"]
    ).T
    # merge AE and SAE dataframes
    final_safety_A_df = pd.concat([AE_A_count, SAE_A_count], axis=1)

    AE_B_count = get_stats_percentage(
        "AE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort B: Chronic Lymphocytic Leukemia (CLL)"]
    ).T
    SAE_B_count = get_stats_percentage(
        "SAE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort B: Chronic Lymphocytic Leukemia (CLL)"]
    ).T
    # merge AE and SAE dataframes
    final_safety_B_df = pd.concat([AE_B_count, SAE_B_count], axis=1)

    AE_C_count = get_stats_percentage(
        "AE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort C: Acute Lymphoblastic Leukemia (ALL)"]
    ).T
    SAE_C_count = get_stats_percentage(
        "SAE", total_infused_df[total_infused_df["Cohort Assignment"] == "Cohort C: Acute Lymphoblastic Leukemia (ALL)"]
    ).T
    # merge AE and SAE dataframes
    final_safety_C_df = pd.concat([AE_C_count, SAE_C_count], axis=1)

    final_safety_df = pd.concat(
        [final_safety_total_df, final_safety_A_df, final_safety_B_df, final_safety_C_df], axis=0
    )

    # TODO: RESPONSE STATS
    if final_subject_A_prim_count > 0:
        responseA_stat = final_responseA_primary_df.copy()
        # replace 'Not Assessed' with 'Not Reported' for all columns in responseA_stat
        responseA_stat = responseA_stat.replace("Not Assessed", "Not Reported")
        final_response_stat_A_BOR_PET = get_stats_percentage("PET-Based Response", responseA_stat)
        final_response_stat_A_BOR_CT = get_stats_percentage("CT-Based Response", responseA_stat)
        final_response_stat_A_ORR_PET = get_stats_percentage("PET-Based ORR", responseA_stat)
        final_response_stat_A_ORR_CT = get_stats_percentage("CT-Based ORR", responseA_stat)

    if final_subject_B_prim_count > 0:
        responseB_stat = final_responseB_primary_df.copy()
        # replace 'Not Assessed' with 'Not Reported' for all columns in responseB_stat
        responseB_stat = responseB_stat.replace("Not Assessed", "Not Reported")
        responseB_stat = responseB_stat.replace("Not Evaluable", "Not Reported")
        final_response_stat_B_BOR_OV = get_stats_percentage("OV-Best Response", responseB_stat)
        final_response_stat_B_BOR_BM = get_stats_percentage("BM-Best Response", responseB_stat)
        final_response_stat_B_ORR_OV = get_stats_percentage("Overall Response", responseB_stat)
        final_response_stat_B_ORR_BM = get_stats_percentage("Bone Marrow Response", responseB_stat)

    if final_subject_C_prim_count > 0:
        responseC_stat = final_responseC_primary_df.copy()
        # replace 'Extramedullary Disease Without Bone Marrow Involvement', 'No Clinical Evidence of Relapse', 'Unknown/Not Assessed' with 'Not Reported' for all columns in responseC_stat
        responseC_stat = responseC_stat.replace(
            "Extramedullary Disease Without Bone Marrow Involvement", "Not Applicable"
        )
        responseC_stat = responseC_stat.replace("No Clinical Evidence of Relapse", "Not Reported")
        responseC_stat = responseC_stat.replace("Unknown/Not Assessed", "Not Reported")
        responseC_stat = responseC_stat.replace("Not Assessed", "Not Reported")
        final_response_stat_C_BOR_OV = get_stats_percentage("Best Overall Response", responseC_stat)
        final_response_stat_C_BOR_ED = get_stats_percentage("Best ED Response", responseC_stat)
        final_response_stat_C_ORR_OV = get_stats_percentage("OV ORR", responseC_stat)
        final_response_stat_C_ORR_ED = get_stats_percentage("ED ORR", responseC_stat)

    if final_subject_BRT_prim_count > 0:
        responseBRT_stat = final_responseBRT_primary_df.copy()
        # replace 'Not Assessed' with 'Not Reported' for all columns in responseBRT_stat
        responseBRT_stat = responseBRT_stat.replace("Not Assessed", "Not Reported")
        final_response_stat_BRT_BOR_PET = get_stats_percentage("PET-Based Response", responseBRT_stat)
        final_response_stat_BRT_BOR_CT = get_stats_percentage("CT-Based Response", responseBRT_stat)
        final_response_stat_BRT_ORR_PET = get_stats_percentage("PET-Based ORR", responseBRT_stat)
        final_response_stat_BRT_ORR_CT = get_stats_percentage("CT-Based ORR", responseBRT_stat)

    # TODO: UPDATE FORMAT FOR NHL, CLL, and ALL after getting the stats

    if final_subject_A_prim_count > 0:
        final_responseA_primary_df.loc[
            (final_responseA_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseA_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseA_primary_df["Event Label"] == "Pre-Retreatment"),
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
        ] = "Transitioned to Retreatment"
        final_responseA_primary_df.loc[
            (final_responseA_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseA_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseA_primary_df["Event Label"] == "Pre-Retreatment"),
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
        ] = "Transitioned to Retreatment"
        final_responseA_primary_df.loc[
            (final_responseA_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseA_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseA_primary_df["Event Label"] == "Pre-Retreatment"),
            "Primary Treatment Time Point (ig_RS1.RSTPT)",
        ] = "Transitioned to Retreatment"

    if final_subject_B_prim_count > 0:
        final_responseB_primary_df.loc[
            (final_responseB_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseB_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseB_primary_df["Event Label"] == "Pre-Retreatment"),
            "Overall CLL Disease Response (ig_RS2.RSCLLCAT)",
        ] = "Transitioned to Retreatment"
        final_responseB_primary_df.loc[
            (final_responseB_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseB_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseB_primary_df["Event Label"] == "Pre-Retreatment"),
            "CLL Bone Marrow Response (ig_RS2.RSCLLBMRESP)",
        ] = "Transitioned to Retreatment"
        final_responseB_primary_df.loc[
            (final_responseB_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseB_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseB_primary_df["Event Label"] == "Pre-Retreatment"),
            "Primary Treatment Time Point (ig_RS1.RSTPT)",
        ] = "Transitioned to Retreatment"

    if final_subject_C_prim_count > 0:
        final_responseC_primary_df.loc[
            (final_responseC_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseC_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseC_primary_df["Event Label"] == "Pre-Retreatment"),
            "Current Overall Response",
        ] = "Transitioned to Retreatment"
        final_responseC_primary_df.loc[
            (final_responseC_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseC_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseC_primary_df["Event Label"] == "Pre-Retreatment"),
            "Current ED Response",
        ] = "Transitioned to Retreatment"
        final_responseC_primary_df.loc[
            (final_responseC_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseC_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseC_primary_df["Event Label"] == "Pre-Retreatment"),
            "Current Response Time Point",
        ] = "Transitioned to Retreatment"

    if final_subject_BRT_prim_count > 0:
        final_responseBRT_primary_df.loc[
            (final_responseBRT_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseBRT_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseBRT_primary_df["Event Label"] == "Pre-Retreatment"),
            "PET-Based NHL Disease Response (ig_RS3.RSNHLPETCAT)",
        ] = "Transitioned to Retreatment"
        final_responseBRT_primary_df.loc[
            (final_responseBRT_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseBRT_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseBRT_primary_df["Event Label"] == "Pre-Retreatment"),
            "CT-Based NHL Disease Response (ig_RS3.RSNHLCTCAT)",
        ] = "Transitioned to Retreatment"
        final_responseBRT_primary_df.loc[
            (final_responseBRT_primary_df["Event Label"] == "Retreatment Long Term Follow-up")
            | (final_responseBRT_primary_df["Event Label"] == "Primary Retreatment")
            | (final_responseBRT_primary_df["Event Label"] == "Pre-Retreatment"),
            "Primary Treatment Time Point (ig_RS1.RSTPT)",
        ] = "Transitioned to Retreatment"

    if export:
        # print(final_output_dir  + '/' + final_output_file_name + '.xlsx')
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

            ## TODO: DSMB-Demo Stats Table
            # * WRITING DATA: LegalSex_df, Age_df, Race_df
            final_LegalSex_df.to_excel(writer, sheet_name="DSMB-Demo Stats Table", index=False, startcol=1, startrow=2)
            final_Age_df.to_excel(writer, sheet_name="DSMB-Demo Stats Table", index=False, startcol=1, startrow=7)
            final_Race_df.to_excel(writer, sheet_name="DSMB-Demo Stats Table", index=False, startcol=1, startrow=11)
            # assign worksheet to variable
            worksheet1 = writer.sheets["DSMB-Demo Stats Table"]

            # * FORMAT DATA
            for i in range(0, len(final_LegalSex_df)):
                for j in range(0, len(final_LegalSex_df.columns)):
                    worksheet1.write(i + 3, j + 1, final_LegalSex_df.iloc[i, j], normal_data_format)
                for k in range(0, 8 - len(final_Age_df.columns)):
                    worksheet1.write(i + 3, len(final_Age_df.columns) + k + 1, "", normal_data_format)
            for i in range(0, len(final_Age_df)):
                for j in range(0, len(final_Age_df.columns)):
                    worksheet1.write(i + 8, j + 1, final_Age_df.iloc[i, j], normal_data_format)
                for k in range(0, 8 - len(final_Age_df.columns)):
                    worksheet1.write(i + 8, len(final_Age_df.columns) + k + 1, "", normal_data_format)
            for i in range(0, len(final_Race_df)):
                for j in range(0, len(final_Race_df.columns)):
                    worksheet1.write(i + 12, j + 1, final_Race_df.iloc[i, j], normal_data_format)
                for k in range(0, 8 - len(final_Age_df.columns)):
                    worksheet1.write(i + 12, len(final_Age_df.columns) + k + 1, "", normal_data_format)

            # * WRITING HEADER AND FORMATTING
            Sex_order = ["Male", "Female", "Nonbinary (X)", "Not Reported"]
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
            for i in range(0, len(Sex_order)):
                worksheet1.write(i + 3, 0, Sex_order[i], bold_11_format)
            for i in range(0, len(Age_order)):
                worksheet1.write(i + 8, 0, Age_order[i], bold_11_format)
            for i in range(0, len(Race_order)):
                worksheet1.write(i + 12, 0, Race_order[i], bold_11_format)

            worksheet1.merge_range("D1:E1", "Cohort A (NHL)", bold_12_format)
            worksheet1.merge_range("F1:G1", "Cohort B (CLL)", bold_12_format)
            worksheet1.merge_range("H1:I1", "Cohort C (ALL)", bold_12_format)
            worksheet1.write(1, 0, "Status", bold_11_format)
            worksheet1.write(1, 1, "Total Screened\nN=" + str(final_status["Total Screened"]), bold_11_wrap_format)
            worksheet1.write(1, 2, "Screen Failed\nN=" + str(final_status["Screen Failed"]), bold_11_wrap_format)
            worksheet1.write(1, 3, "Eligible\nN=" + str(final_status["Cohort A Enrolled"]), bold_11_wrap_format)
            worksheet1.write(1, 4, "Infused\nN=" + str(final_status["Cohort A Infused"]), bold_11_wrap_format)
            worksheet1.write(1, 5, "Eligible\nN=" + str(final_status["Cohort B Enrolled"]), bold_11_wrap_format)
            worksheet1.write(1, 6, "Infused\nN=" + str(final_status["Cohort B Infused"]), bold_11_wrap_format)
            worksheet1.write(1, 7, "Eligible\nN=" + str(final_status["Cohort C Enrolled"]), bold_11_wrap_format)
            worksheet1.write(1, 8, "Infused\nN=" + str(final_status["Cohort C Infused"]), bold_11_wrap_format)
            worksheet1.merge_range("A3:I3", "Legal Sex", bold_11_format)
            worksheet1.merge_range("A8:I8", "Age at Consent", bold_11_format)
            worksheet1.merge_range("A12:I12", "Race", bold_11_format)
            worksheet1.autofit()

            ## TODO: Enrollment Listing
            # * WRITING DATA: enrollment_df
            final_enrollment_df.to_excel(writer, sheet_name="DSMB-Enrollment Listing", index=False, startcol=0)
            # assign worksheet to variable
            worksheet2 = writer.sheets["DSMB-Enrollment Listing"]

            # * FORMAT DATA
            for i in range(0, len(final_enrollment_df)):
                for j in range(0, len(final_enrollment_df.columns)):
                    worksheet2.write(i + 1, j, final_enrollment_df.iloc[i, j], normal_data_format)
            # Autofit
            worksheet2.autofit()

            ## TODO: DSMB-New Infusion Statistics
            # * WRITING DATA: new_infusion_df
            final_infusion_statA.to_excel(
                writer, sheet_name="DSMB-New Infusion Statistics", index=False, startrow=2, startcol=0
            )
            final_infusion_statB.to_excel(
                writer, sheet_name="DSMB-New Infusion Statistics", index=False, startrow=6, startcol=0
            )
            final_infusion_statC.to_excel(
                writer, sheet_name="DSMB-New Infusion Statistics", index=False, startrow=10, startcol=0
            )

            # assign worksheet to variable
            worksheet3 = writer.sheets["DSMB-New Infusion Statistics"]

            # * FORMATING DATA
            for i in range(0, len(final_infusion_statA)):
                for j in range(0, len(final_infusion_statA.columns)):
                    worksheet3.write(i + 3, j + 1, final_infusion_statA.iloc[i, j], normal_data_format)
            for i in range(0, len(final_infusion_statB)):
                for j in range(0, len(final_infusion_statB.columns)):
                    worksheet3.write(i + 7, j + 1, final_infusion_statB.iloc[i, j], normal_data_format)
            for i in range(0, len(final_infusion_statC)):
                for j in range(0, len(final_infusion_statC.columns)):
                    worksheet3.write(i + 11, j + 1, final_infusion_statC.iloc[i, j], normal_data_format)
            # * WRITING HEADER AND FORMATTING
            stat_order = ["Mean SD", "Median", "Range"]

            worksheet3.merge_range("B1:D1", "Cells Infused", bold_12_wrap_format)
            worksheet3.merge_range("E1:F1", "Transduction Efficiency", bold_12_wrap_format)
            worksheet3.write("B2", "Total Cells", bold_12_wrap_format)
            worksheet3.write("C2", "huCART19-IL18 Cells", bold_12_wrap_format)
            worksheet3.write("D2", "Met Target Dose", bold_12_wrap_format)
            worksheet3.write("E2", "%scFv Flow", bold_12_wrap_format)
            worksheet3.write("F2", "Met Target %scFv", bold_12_wrap_format)
            worksheet3.merge_range("A3:F3", "Cohort A (N=" + str(final_infusion_count[0]) + ")", bold_12_wrap_format)
            worksheet3.merge_range("A7:F7", "Cohort B (N=" + str(final_infusion_count[1]) + ")", bold_12_wrap_format)
            worksheet3.merge_range("A11:F11", "Cohort C (N=" + str(final_infusion_count[2]) + ")", bold_12_wrap_format)
            # Merge and format data
            worksheet3.merge_range("D4:D6", final_infusion_statA.iloc[0, 2], normal_data_format)
            worksheet3.merge_range("D8:D10", final_infusion_statB.iloc[0, 2], normal_data_format)
            worksheet3.merge_range("D12:D14", final_infusion_statC.iloc[0, 2], normal_data_format)
            worksheet3.merge_range("F4:F6", final_infusion_statA.iloc[0, 4], normal_data_format)
            worksheet3.merge_range("F8:F10", final_infusion_statB.iloc[0, 4], normal_data_format)
            worksheet3.merge_range("F12:F14", final_infusion_statC.iloc[0, 4], normal_data_format)

            for i in range(0, len(stat_order)):
                worksheet3.write(i + 3, 0, stat_order[i], bold_11_format)
            for i in range(0, len(stat_order)):
                worksheet3.write(i + 7, 0, stat_order[i], bold_11_format)
            for i in range(0, len(stat_order)):
                worksheet3.write(i + 11, 0, stat_order[i], bold_11_format)

            # * Autofit
            worksheet3.autofit()

            ## TODO: DSMB-Infusion Listing
            # * WRITING DATA: infusion_df, infusionR_df
            final_infusion_df.to_excel(writer, sheet_name="DSMB-Infusion Listing", index=False, startrow=1, startcol=0)
            final_infusionR_df.to_excel(
                writer, sheet_name="DSMB-Infusion Listing", index=False, startrow=1, startcol=14
            )
            # assign worksheet to variable
            worksheet4 = writer.sheets["DSMB-Infusion Listing"]
            # * FORMATING DATA
            for i in range(0, len(final_infusion_df)):
                for j in range(0, len(final_infusion_df.columns)):
                    worksheet4.write(i + 2, j, final_infusion_df.iloc[i, j], normal_data_format)
            for i in range(0, len(final_infusionR_df)):
                for j in range(0, len(final_infusionR_df.columns)):
                    worksheet4.write(i + 2, j + 14, final_infusionR_df.iloc[i, j], normal_data_format)
            # * WRITING HEADER AND FORMATTING
            worksheet4.merge_range("A1:A2", "Subject ID", bold_12_wrap_format)
            worksheet4.merge_range("B1:B2", "Study Day (Primary)", bold_12_wrap_format)
            worksheet4.merge_range("C1:C2", "Cohort Assignment", bold_12_wrap_format)
            worksheet4.merge_range("D1:D2", "Dose Level Assignment", bold_12_wrap_format)
            worksheet4.merge_range("E1:E2", "Lymphodepleting Chemotherapy Regimen", bold_12_wrap_format)
            worksheet4.merge_range("F1:F2", "Date of huCART19-IL18 Infusion", bold_12_wrap_format)
            worksheet4.merge_range("G1:J1", "Cells Infused", bold_12_wrap_format)
            worksheet4.merge_range("K1:L1", "Transduction Efficiency", bold_12_wrap_format)

            worksheet4.merge_range("O1:O2", "Subject ID", bold_12_wrap_format)
            worksheet4.merge_range("P1:P2", "Study Day (Retreatment)", bold_12_wrap_format)
            worksheet4.merge_range("Q1:Q2", "Cohort Assignment", bold_12_wrap_format)
            worksheet4.merge_range("R1:R2", "Lymphodepleting Chemotherapy Regimen", bold_12_wrap_format)
            worksheet4.merge_range("S1:S2", "Date of huCART19-IL18 Retreatment Infusion", bold_12_wrap_format)
            worksheet4.merge_range("T1:U1", "Cells Infused", bold_12_wrap_format)
            worksheet4.merge_range("V1:W1", "Transduction Efficiency", bold_12_wrap_format)

            # Autofit
            worksheet4.autofit()

            ## TODO: DSMB-Response Stats
            # * WRITING DATA
            # Safety Data
            final_safety_df.to_excel(writer, sheet_name="DSMB-Response Stats", index=False, startrow=2, startcol=1)
            # Response Data Cohort A
            if final_subject_A_prim_count > 0:
                final_response_stat_A_BOR_PET.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=10, startcol=1
                )
                final_response_stat_A_BOR_CT.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=10, startcol=3
                )
                final_response_stat_A_ORR_PET.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=17, startcol=1
                )
                final_response_stat_A_ORR_CT.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=17, startcol=3
                )
            # Resposne Data Cohort B
            if final_subject_B_prim_count > 0:
                final_response_stat_B_BOR_OV.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=10, startcol=5
                )
                final_response_stat_B_BOR_BM.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=10, startcol=7
                )
                final_response_stat_B_ORR_OV.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=17, startcol=5
                )
                final_response_stat_B_ORR_BM.to_excel(
                    writer, sheet_name="DSMB-Response Stats", index=False, startrow=17, startcol=7
                )
            # Resposne Data Cohort C

            # assign worksheet to variable
            worksheet5 = writer.sheets["DSMB-Response Stats"]

            # * FORMATING DATA
            # Safety Data
            for i in range(0, len(final_safety_df)):
                for j in range(0, len(final_safety_df.columns)):
                    worksheet5.write(i + 3, j + 1, final_safety_df.iloc[i, j], normal_data_format)
            # Response Data Cohort A
            for i in range(0, len(final_response_stat_A_BOR_PET)):
                for j in range(0, len(final_response_stat_A_BOR_PET.columns)):
                    worksheet5.write(i + 11, j + 1, final_response_stat_A_BOR_PET.iloc[i, j], normal_data_format)
            for i in range(0, len(final_response_stat_A_BOR_CT)):
                for j in range(0, len(final_response_stat_A_BOR_CT.columns)):
                    worksheet5.write(i + 11, j + 3, final_response_stat_A_BOR_CT.iloc[i, j], normal_data_format)
            for i in range(0, len(final_response_stat_A_ORR_PET)):
                for j in range(0, len(final_response_stat_A_ORR_PET.columns)):
                    worksheet5.write(i + 18, j + 1, final_response_stat_A_ORR_PET.iloc[i, j], normal_data_format)
            for i in range(0, len(final_response_stat_A_ORR_CT)):
                for j in range(0, len(final_response_stat_A_ORR_CT.columns)):
                    worksheet5.write(i + 18, j + 3, final_response_stat_A_ORR_CT.iloc[i, j], normal_data_format)
            if final_subject_B_prim_count == 0:
                for i in range(0, 6):
                    worksheet5.write(i + 11, 5, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 11, 7, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 18, 5, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 18, 7, "0 (0%)", normal_data_format)
            elif final_subject_B_prim_count > 0:
                # Response Data Cohort B
                for i in range(0, len(final_response_stat_B_BOR_OV)):
                    for j in range(0, len(final_response_stat_B_BOR_OV.columns)):
                        worksheet5.write(i + 11, j + 5, final_response_stat_B_BOR_OV.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_B_BOR_BM)):
                    for j in range(0, len(final_response_stat_B_BOR_BM.columns)):
                        worksheet5.write(i + 11, j + 7, final_response_stat_B_BOR_BM.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_B_ORR_OV)):
                    for j in range(0, len(final_response_stat_B_ORR_OV.columns)):
                        worksheet5.write(i + 18, j + 5, final_response_stat_B_ORR_OV.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_B_ORR_BM)):
                    for j in range(0, len(final_response_stat_B_ORR_BM.columns)):
                        worksheet5.write(i + 18, j + 7, final_response_stat_B_ORR_BM.iloc[i, j], normal_data_format)
            if final_subject_C_prim_count == 0:
                for i in range(0, 6):
                    worksheet5.write(i + 11, 9, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 11, 11, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 18, 9, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 18, 11, "0 (0%)", normal_data_format)
            elif final_subject_C_prim_count > 0:
                # Response Data Cohort C
                for i in range(0, len(final_response_stat_C_BOR_OV)):
                    for j in range(0, len(final_response_stat_C_BOR_OV.columns)):
                        worksheet5.write(i + 11, j + 9, final_response_stat_C_BOR_OV.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_C_BOR_ED)):
                    for j in range(0, len(final_response_stat_C_BOR_ED.columns)):
                        worksheet5.write(i + 11, j + 11, final_response_stat_C_BOR_ED.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_C_ORR_OV)):
                    for j in range(0, len(final_response_stat_C_ORR_OV.columns)):
                        worksheet5.write(i + 18, j + 9, final_response_stat_C_ORR_OV.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_C_ORR_ED)):
                    for j in range(0, len(final_response_stat_C_ORR_ED.columns)):
                        worksheet5.write(i + 18, j + 11, final_response_stat_C_ORR_ED.iloc[i, j], normal_data_format)
            if final_subject_BRT_prim_count == 0:
                for i in range(0, 6):
                    worksheet5.write(i + 28, 1, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 28, 3, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 35, 1, "0 (0%)", normal_data_format)
                    worksheet5.write(i + 35, 3, "0 (0%)", normal_data_format)
            elif final_subject_BRT_prim_count > 0:
                # Response Data Cohort B- RT
                for i in range(0, len(final_response_stat_BRT_BOR_PET)):
                    for j in range(0, len(final_response_stat_BRT_BOR_PET.columns)):
                        worksheet5.write(i + 28, j + 1, final_response_stat_BRT_BOR_PET.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_BRT_BOR_CT)):
                    for j in range(0, len(final_response_stat_BRT_BOR_CT.columns)):
                        worksheet5.write(i + 28, j + 3, final_response_stat_BRT_BOR_CT.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_BRT_ORR_PET)):
                    for j in range(0, len(final_response_stat_BRT_ORR_PET.columns)):
                        worksheet5.write(i + 35, j + 1, final_response_stat_BRT_ORR_PET.iloc[i, j], normal_data_format)
                for i in range(0, len(final_response_stat_BRT_ORR_CT)):
                    for j in range(0, len(final_response_stat_BRT_ORR_CT.columns)):
                        worksheet5.write(i + 35, j + 3, final_response_stat_BRT_ORR_CT.iloc[i, j], normal_data_format)
            # * WRITING HEADER AND FORMATTING
            # Safety Headers
            # number of subject of final_safety_total_df
            safety_total_df_subject_count = len(final_infusion_df["Subject"].unique())
            worksheet5.merge_range(
                "B1:E1", "Safety Statistics (N=" + str(safety_total_df_subject_count) + ")", bold_12_wrap_format
            )
            worksheet5.merge_range("B2:C2", "Adverse Events", bold_11_format)
            worksheet5.merge_range("D2:E2", "Serious Adverse Events ", bold_11_format)
            worksheet5.write("B3", "Yes", bold_11_format)
            worksheet5.write("C3", "No", bold_11_format)
            worksheet5.write("D3", "Yes", bold_11_format)
            worksheet5.write("E3", "No", bold_11_format)
            worksheet5.write("A4", "All Cohorts", bold_11_format)
            worksheet5.write("A5", "Cohort A", bold_11_format)
            worksheet5.write("A6", "Cohort B", bold_11_format)
            worksheet5.write("A7", "Cohort C", bold_11_format)

            # Response Headers
            worksheet5.merge_range(
                "A9:D9", "NHL Subject Response (N=" + str(final_subject_A_prim_count) + ")", bold_12_format
            )
            worksheet5.merge_range(
                "E9:H9", "CLL Subject Response (N=" + str(final_subject_B_prim_count) + ")", bold_12_format
            )
            worksheet5.merge_range("I9:L9", "ALL Subject Response", bold_12_format)
            worksheet5.merge_range(
                "A26:D26",
                "Richter's Transformation Subject Response (N=" + str(final_subject_BRT_prim_count) + ")",
                bold_12_format,
            )
            worksheet5.merge_range("A10:B10", "PET-Based Response", bold_11_format)
            worksheet5.merge_range("C10:D10", "CT-Based Response", bold_11_format)
            worksheet5.merge_range("E10:F10", "Overall Response", bold_11_format)
            worksheet5.merge_range("G10:H10", "Bone Marrow Response", bold_11_format)
            if final_subject_C_prim_count > 0:
                worksheet5.merge_range(
                    "I10:J10", "Overall Response (N=" + str(final_subject_C_prim_OR_count) + ")", bold_11_format
                )
                worksheet5.merge_range(
                    "K10:L10",
                    "Extramedullary Disease without Bone Marrow Involvement (N="
                    + str(final_subject_C_prim_ED_count)
                    + ")",
                    bold_11_format,
                )
            else:
                worksheet5.merge_range("I10:J10", "Overall Response (N=0)", bold_11_format)
                worksheet5.merge_range(
                    "K10:L10", "Extramedullary Disease without Bone Marrow Involvement (N=0)", bold_11_format
                )
            worksheet5.merge_range("A27:B27", "PET-Based Response", bold_11_format)
            worksheet5.merge_range("C27:D27", "CT-Based Response", bold_11_format)
            worksheet5.merge_range("A11:D11", "Best Overall Response (BOR)", bold_11_format)
            worksheet5.merge_range("E11:H11", "Best Overall Response (BOR)", bold_11_format)
            worksheet5.merge_range("I11:L11", "Best Overall Response (BOR)", bold_11_format)
            worksheet5.merge_range("A28:D28", "Best Overall Response (BOR)", bold_11_format)
            worksheet5.merge_range("A18:D18", "Overall Response Rate (ORR) at Month 3", bold_11_format)
            worksheet5.merge_range("E18:H18", "Overall Response Rate (ORR) at Month 3", bold_11_format)
            worksheet5.merge_range("I18:L18", "Overall Response Rate (ORR) at Day 28", bold_11_format)
            worksheet5.merge_range("A35:D35", "Overall Response Rate (ORR) at Month 3", bold_11_format)
            # Listing Response Criteria
            response_A_PET = [
                "Complete Metabolic Response (CMR)",
                "Partial Metabolic Response (PMR)",
                "No Metabolic Response (NMR)",
                "Indeterminate Response (IR)",
                "Progressive Metabolic Disease (PMD)",
                "Not Reported",
            ]
            response_A_CT = [
                "Complete Radiologic Response (CR)",
                "Partial Response (PR)",
                "Stable Disease (SD)",
                "Indeterminate Response (IR)",
                "Progressive Disease (PD)",
                "Not Reported",
            ]
            resposne_B_OV = [
                "Complete Remission (CR)",
                "Complete Remission with Incomplete Marrow Recovery (CRi)",
                "Partial Remission (PR)",
                "Stable Disease (SD)",
                "Progressive Disease (PD)",
                "Not Reported",
            ]
            response_B_BM = [
                "Complete Remission (CR)",
                "Partial Remission (PR)",
                "Progressive Disease (PD)",
                "Stable Disease (SD)",
                "Not Reported",
            ]
            response_C_OV = [
                "Complete Remission (CR)",
                "Complete Remission with Incomplete Blood Count Recovery (CRi)",
                "Complete Remission with Residual Mediastinal Disease (CRu)",
                "Treatment Failure (TF)",
                "Relapsed Disease (RD)",
                "Not Reported",
            ]
            response_C_ED = [
                "Complete Remission (CR)",
                "Partial Remission (PR)",
                "Stable Disease (SD)",
                "Indeterminate Response",
                "Progressive Disease (PD)",
                "Not Reported",
            ]
            for i in range(0, len(response_A_PET)):
                worksheet5.write(i + 11, 0, response_A_PET[i], bold_11_format)
                worksheet5.write(i + 18, 0, response_A_PET[i], bold_11_format)
                worksheet5.write(i + 28, 0, response_A_PET[i], bold_11_format)
                worksheet5.write(i + 35, 0, response_A_PET[i], bold_11_format)
            for i in range(0, len(response_A_CT)):
                worksheet5.write(i + 11, 2, response_A_CT[i], bold_11_format)
                worksheet5.write(i + 18, 2, response_A_CT[i], bold_11_format)
                worksheet5.write(i + 28, 2, response_A_CT[i], bold_11_format)
                worksheet5.write(i + 35, 2, response_A_CT[i], bold_11_format)
            for i in range(0, len(resposne_B_OV)):
                worksheet5.write(i + 11, 4, resposne_B_OV[i], bold_11_format)
                worksheet5.write(i + 18, 4, resposne_B_OV[i], bold_11_format)
            for i in range(0, len(response_B_BM)):
                worksheet5.write(i + 11, 6, response_B_BM[i], bold_11_format)
                worksheet5.write(i + 18, 6, response_B_BM[i], bold_11_format)
            worksheet5.write("G17", "black cell", black_cell)
            worksheet5.write("H17", "black cell", black_cell)
            worksheet5.write("G24", "black cell", black_cell)
            worksheet5.write("H24", "black cell", black_cell)
            for i in range(0, len(response_C_OV)):
                worksheet5.write(i + 11, 8, response_C_OV[i], bold_11_format)
                worksheet5.write(i + 18, 8, response_C_OV[i], bold_11_format)
            for i in range(0, len(response_C_ED)):
                worksheet5.write(i + 11, 10, response_C_ED[i], bold_11_format)
                worksheet5.write(i + 18, 10, response_C_ED[i], bold_11_format)

            worksheet5.autofit()

            ## TODO: Response Listing NHL
            # * WRITING DATA
            if final_subject_A_prim_count > 0:
                final_responseA_primary_df.to_excel(
                    writer, sheet_name="Response Listing NHL", index=False, header=False, startrow=3, startcol=0
                )
            if final_subject_A_retx_count > 0:
                final_responseA_retreatment_df.to_excel(
                    writer, sheet_name="Response Listing NHL", index=False, header=False, startrow=3, startcol=15
                )
            # assigning the worksheet
            worksheet6 = writer.sheets["Response Listing NHL"]
            # * FORMATING DATA
            if final_subject_A_prim_count > 0:
                for i in range(0, len(final_responseA_primary_df)):
                    for j in range(0, len(final_responseA_primary_df.columns)):
                        worksheet6.write(i + 3, j, final_responseA_primary_df.iloc[i, j], normal_data_format)
            if final_subject_A_retx_count > 0:
                for i in range(0, len(final_responseA_retreatment_df)):
                    for j in range(0, len(final_responseA_retreatment_df.columns)):
                        worksheet6.write(i + 3, j + 15, final_responseA_retreatment_df.iloc[i, j], normal_data_format)
            # * WRITING HEADER AND FORMATTING
            if final_subject_A_prim_count > 0:
                worksheet6.merge_range("A1:M1", "Cohort A (NHL)- Primary Follow-up", bold_12_format)
                worksheet6.merge_range("A2:A3", "Subject ID", bold_11_format)
                worksheet6.merge_range("B2:D2", "Current Response", bold_11_format)
                worksheet6.merge_range("E2:H2", "Best Response/Timepoint", bold_11_format)
                worksheet6.merge_range("I2:J2", "Overall Response/Month 3", bold_11_format)
                worksheet6.write("B3", "PET-Based Response", bold_11_format)
                worksheet6.write("C3", "CT-Based Response", bold_11_format)
                worksheet6.write("D3", "Study Timepoint", bold_11_format)
                worksheet6.write("E3", "PET-Based Response", bold_11_format)
                worksheet6.write("F3", "Study Timepoint", bold_11_format)
                worksheet6.write("G3", "CT-Based Response", bold_11_format)
                worksheet6.write("H3", "Study Timepoint", bold_11_format)
                worksheet6.write("I3", "PET-Based ORR", bold_11_format)
                worksheet6.write("J3", "CT-Based ORR", bold_11_format)
                worksheet6.merge_range("K2:K3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet6.merge_range("L2:L3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet6.merge_range("M2:M3", "Study Status", bold_11_wrap_format)
            if final_subject_A_retx_count > 0:
                worksheet6.merge_range("P1:AB1", "Cohort A (NHL)- Retreatment Follow-up", bold_12_format)
                worksheet6.merge_range("P2:P3", "Subject ID", bold_11_format)
                worksheet6.merge_range("Q2:S2", "Current Response", bold_11_format)
                worksheet6.merge_range("T2:W2", "Best Response/Timepoint", bold_11_format)
                worksheet6.merge_range("X2:Y2", "Overall Response/Month 3-R", bold_11_format)
                worksheet6.write("Q3", "PET-Based Response", bold_11_format)
                worksheet6.write("R3", "CT-Based Response", bold_11_format)
                worksheet6.write("S3", "Study Timepoint", bold_11_format)
                worksheet6.write("T3", "PET-Based Response", bold_11_format)
                worksheet6.write("U3", "Study Timepoint", bold_11_format)
                worksheet6.write("V3", "CT-Based Response", bold_11_format)
                worksheet6.write("W3", "Study Timepoint", bold_11_format)
                worksheet6.write("X3", "PET-Based ORR", bold_11_format)
                worksheet6.write("Y3", "CT-Based ORR", bold_11_format)
                worksheet6.merge_range("Z2:Z3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet6.merge_range("AA2:AA3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet6.merge_range("AB2:AB3", "Study Status", bold_11_wrap_format)
            # Autofit
            worksheet6.autofit()

            ## TODO:
            # * WRITING DATA
            if final_subject_B_prim_count > 0:
                final_responseB_primary_df.to_excel(
                    writer, sheet_name="Response Listing CLL", index=False, startrow=2, startcol=0
                )
                worksheet7 = writer.sheets["Response Listing CLL"]
                # * FORMATING DATA
                for i in range(0, len(final_responseB_primary_df)):
                    for j in range(0, len(final_responseB_primary_df.columns)):
                        worksheet7.write(i + 3, j, final_responseB_primary_df.iloc[i, j], normal_data_format)

                # * WRITING HEADER AND FORMATTING
                worksheet7.merge_range("A1:M1", "Cohort B (CLL) - Primary Follow-up", bold_12_format)
                worksheet7.merge_range("A2:A3", "Subject ID", bold_11_format)
                worksheet7.merge_range("B2:D2", "Current Response", bold_11_format)
                worksheet7.merge_range("E2:H2", "Best Response/Timepoint", bold_11_format)
                worksheet7.merge_range("I2:J2", "Overall Response/Month 3", bold_11_format)
                worksheet7.write("B3", "Overall Response", bold_11_format)
                worksheet7.write("C3", "Bone Marrow Response", bold_11_format)
                worksheet7.write("D3", "Study Timepoint", bold_11_format)
                worksheet7.write("E3", "Overall Response", bold_11_format)
                worksheet7.write("F3", "Study Timepoint", bold_11_format)
                worksheet7.write("G3", "Bone Marrow Response", bold_11_format)
                worksheet7.write("H3", "Study Timepoint", bold_11_format)
                worksheet7.write("I3", "Overall Response", bold_11_format)
                worksheet7.write("J3", "Bone Marrow Response", bold_11_format)
                worksheet7.merge_range("K2:K3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet7.merge_range("L2:L3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet7.merge_range("M2:M3", "Study Status", bold_11_wrap_format)

                if final_subject_B_retx_count > 0:
                    final_responseB_retreatment_df.to_excel(
                        writer, sheet_name="Response Listing CLL", index=False, startrow=2, startcol=0
                    )
                    # * FORMATING DATA
                    for i in range(0, len(final_responseB_retreatment_df)):
                        for j in range(0, len(final_responseB_retreatment_df.columns)):
                            worksheet7.write(
                                i + 3, j + 15, final_responseB_retreatment_df.iloc[i, j], normal_data_format
                            )

                    # * WRITING HEADER AND FORMATTING
                    if final_subject_B_retx_count > 0:
                        worksheet7.merge_range("P1:AB1", "Cohort B (CLL) - Retreatment Follow-up", bold_12_format)
                        worksheet7.merge_range("P2:P3", "Subject ID", bold_11_format)
                        worksheet7.merge_range("Q2:S2", "Current Response", bold_11_format)
                        worksheet7.merge_range("T2:W2", "Best Response/Timepoint", bold_11_format)
                        worksheet7.merge_range("X2:Y2", "Overall Response/Month 3", bold_11_format)
                        worksheet7.write("Q3", "Overall Response", bold_11_format)
                        worksheet7.write("R3", "Bone Marrow Response", bold_11_format)
                        worksheet7.write("S3", "Study Timepoint", bold_11_format)
                        worksheet7.write("T3", "Overall Response", bold_11_format)
                        worksheet7.write("U3", "Study Timepoint", bold_11_format)
                        worksheet7.write("V3", "Bone Marrow Response", bold_11_format)
                        worksheet7.write("W3", "Study Timepoint", bold_11_format)
                        worksheet7.write("X3", "Overall Response", bold_11_format)
                        worksheet7.write("Y3", "Bone Marrow Response", bold_11_format)
                        worksheet7.merge_range("Z2:Z3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                        worksheet7.merge_range("AA2:AA3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                        worksheet7.merge_range("AB2:AB3", "Study Status", bold_11_wrap_format)
                # Autofit
                worksheet7.autofit()

            ## TODO: Response Listing Richter's Transformation
            # * WRITING DATA
            if final_subject_BRT_prim_count > 0:
                final_responseBRT_primary_df.to_excel(
                    writer, sheet_name="Response Listing Richter", index=False, header=False, startrow=3, startcol=0
                )
                # assigning the worksheet
                worksheet8 = writer.sheets["Response Listing Richter"]
                for i in range(0, len(final_responseBRT_primary_df)):
                    for j in range(0, len(final_responseBRT_primary_df.columns)):
                        worksheet8.write(i + 3, j, final_responseBRT_primary_df.iloc[i, j], normal_data_format)
                worksheet8.merge_range(
                    "A1:M1", "Cohort B (Richter's Transformation)- Primary Follow-up", bold_12_format
                )
                worksheet8.merge_range("A2:A3", "Subject ID", bold_11_format)
                worksheet8.merge_range("B2:D2", "Current Response", bold_11_format)
                worksheet8.merge_range("E2:H2", "Best Response/Timepoint", bold_11_format)
                worksheet8.merge_range("I2:J2", "Overall Response/Month 3", bold_11_format)
                worksheet8.write("B3", "PET-Based Response", bold_11_format)
                worksheet8.write("C3", "CT-Based Response", bold_11_format)
                worksheet8.write("D3", "Study Timepoint", bold_11_format)
                worksheet8.write("E3", "PET-Based Response", bold_11_format)
                worksheet8.write("F3", "Study Timepoint", bold_11_format)
                worksheet8.write("G3", "CT-Based Response", bold_11_format)
                worksheet8.write("H3", "Study Timepoint", bold_11_format)
                worksheet8.write("I3", "PET-Based ORR", bold_11_format)
                worksheet8.write("J3", "CT-Based ORR", bold_11_format)
                worksheet8.merge_range("K2:K3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet8.merge_range("L2:L3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet8.merge_range("M2:M3", "Study Status", bold_11_wrap_format)
                worksheet8.autofit()
            if final_subject_BRT_retx_count > 0:
                final_responseBRT_retreatment_df.to_excel(
                    writer, sheet_name="Response Listing Richter", index=False, header=False, startrow=3, startcol=15
                )
                for i in range(0, len(final_responseBRT_retreatment_df)):
                    for j in range(0, len(final_responseBRT_retreatment_df.columns)):
                        worksheet8.write(i + 3, j + 15, final_responseBRT_retreatment_df.iloc[i, j], normal_data_format)
                worksheet8.merge_range("P1:AB1", "Cohort A (NHL)- Retreatment Follow-up", bold_12_format)
                worksheet8.merge_range("P2:P3", "Subject ID", bold_11_format)
                worksheet8.merge_range("Q2:S2", "Current Response", bold_11_format)
                worksheet8.merge_range("T2:W2", "Best Response/Timepoint", bold_11_format)
                worksheet8.merge_range("X2:Y2", "Overall Response/Month 3-R", bold_11_format)
                worksheet8.write("Q3", "PET-Based Response", bold_11_format)
                worksheet8.write("R3", "CT-Based Response", bold_11_format)
                worksheet8.write("S3", "Study Timepoint", bold_11_format)
                worksheet8.write("T3", "PET-Based Response", bold_11_format)
                worksheet8.write("U3", "Study Timepoint", bold_11_format)
                worksheet8.write("V3", "CT-Based Response", bold_11_format)
                worksheet8.write("W3", "Study Timepoint", bold_11_format)
                worksheet8.write("X3", "PET-Based ORR", bold_11_format)
                worksheet8.write("Y3", "CT-Based ORR", bold_11_format)
                worksheet8.merge_range("Z2:Z3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet8.merge_range("AA2:AA3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet8.merge_range("AB2:AB3", "Study Status", bold_11_wrap_format)
                worksheet8.autofit()

            ## TODO:
            # * WRITING DATA
            # * FORMATING DATA
            # * WRITING HEADER AND FORMATTING
            # Autofit
            if final_subject_C_prim_count > 0:
                final_responseC_primary_df.to_excel(
                    writer, sheet_name="Response Listing ALL", index=False, startrow=2, startcol=0
                )
                worksheet9 = writer.sheets["Response Listing ALL"]
                # * FORMATING DATA
                for i in range(0, len(final_responseC_primary_df)):
                    for j in range(0, len(final_responseC_primary_df.columns)):
                        worksheet9.write(i + 3, j, final_responseC_primary_df.iloc[i, j], normal_data_format)

                # * WRITING HEADER AND FORMATTING
                worksheet9.merge_range("A1:M1", "Cohort C (ALL) - Primary Follow-up", bold_12_format)
                worksheet9.merge_range("A2:A3", "Subject ID", bold_11_format)
                worksheet9.merge_range("B2:D2", "Current Response", bold_11_format)
                worksheet9.merge_range("E2:H2", "Best Response/Timepoint", bold_11_format)
                worksheet9.merge_range("I2:J2", "Overall Response/Day 28", bold_11_format)
                worksheet9.write("B3", "Overall Response", bold_11_format)
                worksheet9.write("C3", "Extramedullary Disease without Bone Marrow Involvement", bold_11_format)
                worksheet9.write("D3", "Study Timepoint", bold_11_format)
                worksheet9.write("E3", "Overall Response", bold_11_format)
                worksheet9.write("F3", "Study Timepoint", bold_11_format)
                worksheet9.write("G3", "Extramedullary Disease without Bone Marrow Involvement", bold_11_format)
                worksheet9.write("H3", "Study Timepoint", bold_11_format)
                worksheet9.write("I3", "Overall Response", bold_11_format)
                worksheet9.write("J3", "Extramedullary Disease without Bone Marrow Involvement", bold_11_format)
                worksheet9.merge_range("K2:K3", "Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet9.merge_range("L2:L3", "Serious Adverse Events \n(Y/N)", bold_11_wrap_format)
                worksheet9.merge_range("M2:M3", "Study Status", bold_11_wrap_format)

                # if final_subject_C_retx_count > 0:
                #     final_responseB_retreatment_df.to_excel(writer, sheet_name='Response Listing CLL', index = False, startrow=2, startcol=0)
                #     # * FORMATING DATA
                #     for i in range(0, len(final_responseC_retreatment_df)):
                #         for j in range(0, len(final_responseB_retreatment_df.columns)):
                #             worksheet7.write(i + 3, j + 15, final_responseB_retreatment_df.iloc[i, j], normal_data_format)

                # Autofit
                worksheet9.autofit()
