#!/usr/bin/env python3
import pandas as pd
import numpy as np
from util import *
from DSMB.DSMB_util import *
from dateutil.relativedelta import *
from datetime import datetime, date
from typing import Optional


def DSMB15122(
    data,
    export,
    output_dir,
    output_file_name,
    debug,
):
    # TODO: DEMO ENROLLMENT LISTING
    if not data["DM"].empty:
        # Subject
        enrollment_df = data["DM"][["Subject"]].copy()

        # Cohort Assignment
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DSCA",
            "Cohort Assignment (IG_NS_NA_DSCA1.CL_NS_YH_CACHASCOD_cl_NS_COHORT1)",
            "Cohort Assignment",
        )
        # Dose Level
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DSDLA",
            "Dose Level Assignment (IG_NS_NA_DSDLA1.CL_NS_YH_DLADOSELV_cl_NS_DOSELV1)",
            "Dose Level",
        )
        # get study treatment administered data from EXINF for subjects did not end of study before infusion
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "EXINF",
            "Was study treatment administered? (IG_NS_NA_EXINF1.CL_NS_NH_INFADMIN_cl_YS_YN1)",
            "Study Treatment Administered1",
        )
        # Diagnosis
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "MHDIAG",
            "Primary Diagnosis (IG_NS_NA_MHDIAG1.CL_NS_YH_PRMDIAG_cl_NS_PRMDIAG)",
            "Disease",
        )
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "MHDIAG",
            "Specify Other Diagnosis (IG_NS_NA_MHDIAG1.TX_NS_NH_PRMDIAGOTH)",
            "Disease other",
        )
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "MHDIAG",
            "Event Date",
            "Event Date DIAG",
        )
        enrollment_df["Disease Type"] = None
        # List the columns in the order you want to use them for filling 'Disease'
        columns_to_fill_from = ["Disease", "Disease other"]
        # Use fillna() in a loop to fill 'Disease' from the specified columns
        for col in columns_to_fill_from:
            enrollment_df["Disease Type"] = enrollment_df["Disease Type"].fillna(
                enrollment_df[col]
            )
        enrollment_df = enrollment_df.drop(columns=columns_to_fill_from)

        # Legal Sex
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Legal Sex (IG_NS_NA_DM1.CL_NS_YH_SEX_cl_NS_DMSEX1)",
            "Legal Sex",
        )
        # Sex Assigned at Birth
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Sex Assigned at Birth (IG_NS_NA_DM1.CL_NS_NH_BRTHSEX_cl_NS_DMSEX3)",
            "Sex Assigned at Birth",
        )
        # Gender Identity
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Gender Identity (IG_NS_NA_DM1.CL_NS_NH_GENDERID_cl_NS_DMSEX2)",
            "Gender Identity",
        )
        # Age at Consent
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Date of Birth (IG_NS_NA_DM1.DT_NS_NH_BRTHDAT)",
            "Date of Birth",
        )
        # Pre-screening consent
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Pre-Screening Consent Date (IG_NS_NA_DM1.DT_NS_YH_RFICDAT)",
            "Consent Date",
        )
        enrollment_df["Consent Date"] = pd.to_datetime(enrollment_df["Consent Date"])
        enrollment_df["Date of Birth"] = pd.to_datetime(enrollment_df["Date of Birth"])
        mask = ~enrollment_df[["Consent Date", "Date of Birth"]].isnull().any(axis=1)
        enrollment_df.loc[mask, "Age at Consent"] = enrollment_df[mask].apply(
            lambda x: relativedelta(x["Consent Date"], x["Date of Birth"]).years, axis=1
        )
        # main consent
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "IE",
            "Main Consent Date (IG_NS_NA_IE1.DT_NS_NH_MAINCDAT)",
            "Main Consent Date",
        )
        # for two IEs, use the latest IE
        enrollment_df["Main Consent Date"] = pd.to_datetime(
            enrollment_df["Main Consent Date"]
        )
        enrollment_df = enrollment_df.sort_values(["Subject", "Main Consent Date"])
        enrollment_df = enrollment_df.drop_duplicates(subset=["Subject"], keep="last")
        enrollment_df = enrollment_df.drop_duplicates()
        #      enrollment_df["Main Consent Date"] = pd.to_datetime(
        #           enrollment_df["Main Consent Date"]
        #       )
        # for rows that 'Pre-Screening Consent Date' isnull but 'Main Consent Date' is not null, then use 'Main Consent Date' instead to calculate age
        enrollment_df.loc[
            (
                enrollment_df["Consent Date"].isnull()
                & enrollment_df["Main Consent Date"].notnull()
            ),
            "Age at Consent",
        ] = enrollment_df.loc[
            (
                enrollment_df["Consent Date"].isnull()
                & enrollment_df["Main Consent Date"].notnull()
            )
        ].apply(
            lambda x: relativedelta(x["Main Consent Date"], x["Date of Birth"]).years,
            axis=1,
        )

        # Race
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Race (IG_NS_NA_DM1.CL_NS_YH_RACE_cl_NS_DMRACE1)",
            "Race",
        )
        # Ethnicity
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Ethnicity (IG_NS_NA_DM1.CL_NS_NH_ETHNIC_cl_NS_DMETHNIC1)",
            "Ethnicity",
        )
        # Event Date for DM
        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DM",
            "Event Date",
            "Event Date DM",
        )
        # when there are two DM entered, use the latest DM
        enrollment_df["Event Date DM"] = pd.to_datetime(enrollment_df["Event Date DM"])
        enrollment_df = enrollment_df.sort_values(["Subject", "Event Date DM"])
        enrollment_df = enrollment_df.drop_duplicates(subset=["Subject"], keep="last")
        enrollment_df = enrollment_df.drop(columns=["Event Date DM"])
        enrollment_df = enrollment_df.drop_duplicates()
        # when there are two diagnosis entered, use the latest
        enrollment_df["Event Date DIAG"] = pd.to_datetime(
            enrollment_df["Event Date DIAG"]
        )
        enrollment_df = enrollment_df.sort_values(["Subject", "Event Date DIAG"])
        enrollment_df = enrollment_df.drop_duplicates(subset=["Subject"], keep="last")
        enrollment_df = enrollment_df.drop(columns=["Event Date DIAG"])
        enrollment_df = enrollment_df.drop_duplicates()

        # Subject meets all study eligibility? Only get data eligibility data from IE when IE is entered, otherwise check DSEOS
        if not data["IE"].empty:
            #  print("hi IE")
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "IE",
                "Subject Meets All Study Eligibility (IG_NS_NA_IE3.CL_NS_YH_ELIGYN_cl_YS_YN1)",
                "Subject meets all study eligibility?1",
            )
            # Reason for Screen Failure
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "IE",
                "Other Screen Fail Reason (IG_NS_NA_IE4.TX_NS_YH_OTHRSFREAS)",
                "Reason for Screen Failure1",
            )
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "IE",
                "Screen Failure Reason (IG_NS_NA_IE4.CL_NS_YH_IECAT_cl_NS_IEREASSF1)",
                "SF1",
            )
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "IE",
                "Select the Primary Inclusion Criterion Excluding this Subject (IG_NS_NA_IE4.CL_NS_NH_ITESTCD_cl_NS_IEINCL1)",
                "SF2",
            )
            enrollment_df["SF2"] = enrollment_df[enrollment_df["SF2"].notna()][
                "SF2"
            ].astype(str)
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "IE",
                "Select the Primary Exclusion Criterion Excluding this Subject (IG_NS_NA_IE4.CL_NS_NH_ETESTCD_cl_NS_IEEXCL1)",
                "SF3",
            )
            enrollment_df["SF3"] = enrollment_df[enrollment_df["SF3"].notna()][
                "SF3"
            ].astype(str)
            enrollment_df["SF4"] = (
                enrollment_df["SF1"].fillna("")
                + " "
                + enrollment_df["SF2"].fillna("")
                + enrollment_df["SF3"].fillna("")
            )
            enrollment_df["Reason for Screen Failure1"].fillna(
                enrollment_df["SF4"], inplace=True
            )
            enrollment_df = enrollment_df.drop(columns=["SF1", "SF2", "SF3", "SF4"])

        if not data["DSEOS"].empty:
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "DSEOS",
                "Did the Subject receive the investigational product? (IG_NS_NA_DSEOS1.CL_NS_NH_EOSRIP_cl_YS_YN1)",
                "Study Treatment Administered2",
            )
            enrollment_df["Study Treatment Administered2"] = enrollment_df[
                enrollment_df["Study Treatment Administered2"].notna()
            ]["Study Treatment Administered2"].astype(str)
            # if subject did not sign main consent form, implies subject screen failured before IE is entered
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "DSEOS",
                "Did the Subject sign the main consent form? (IG_NS_NA_DSEOS1.CL_NS_NH_MCNSNT_cl_YS_YN1)",
                "Subject meets all study eligibility?2",
            )
            enrollment_df = add_rename_column_corelisting(
                enrollment_df,
                data,
                "DSEOS",
                "Provide Supportive Information (IG_NS_NA_DSEOS2.TX_NS_YH_EOSTERM)",
                "Reason for Screen Failure2",
            )
            # when subject sign main consent is not "No" in DSEOS or subject is eligible in IE, set the following 3 data from DSEOS to ""
            enrollment_df.loc[
                (enrollment_df["Subject meets all study eligibility?2"] != "No")
                | (enrollment_df["Subject meets all study eligibility?1"] == "Yes"),
                "Subject meets all study eligibility?2",
            ] = ""
            enrollment_df.loc[
                (enrollment_df["Subject meets all study eligibility?2"] != "No")
                | (enrollment_df["Subject meets all study eligibility?1"] == "Yes"),
                "Reason for Screen Failure2",
            ] = ""
            enrollment_df.loc[
                (enrollment_df["Subject meets all study eligibility?2"] != "No")
                | (enrollment_df["Subject meets all study eligibility?1"] == "Yes"),
                "Study Treatment Administered2",
            ] = ""

        enrollment_df = add_rename_column_corelisting(
            enrollment_df,
            data,
            "DSEOS",
            "End of Study Date (IG_NS_NA_DSEOS1.DT_NS_YH_EOSDAT)",
            "End of Study Date",
        )
        # combine the data from IE and DSEOS
        enrollment_df["Subject meets all study eligibility?"] = (
            enrollment_df["Subject meets all study eligibility?1"].fillna("")
            + " "
            + enrollment_df["Subject meets all study eligibility?2"].fillna("")
        )
        enrollment_df["Subject meets all study eligibility?"].fillna(
            enrollment_df["Subject meets all study eligibility?"], inplace=True
        )
        enrollment_df = enrollment_df.drop(
            columns=[
                "Subject meets all study eligibility?1",
                "Subject meets all study eligibility?2",
            ]
        )
        enrollment_df["Reason for Screen Failure"] = (
            enrollment_df["Reason for Screen Failure1"].fillna("")
            + " "
            + enrollment_df["Reason for Screen Failure2"].fillna("")
        )
        enrollment_df["Reason for Screen Failure"].fillna(
            enrollment_df["Reason for Screen Failure"], inplace=True
        )
        enrollment_df = enrollment_df.drop(
            columns=[
                "Reason for Screen Failure1",
                "Reason for Screen Failure2",
            ]
        )
        enrollment_df["Study Treatment Administered"] = (
            enrollment_df["Study Treatment Administered1"].fillna("")
            + " "
            + enrollment_df["Study Treatment Administered2"].fillna("")
        )
        enrollment_df["Study Treatment Administered"].fillna(
            enrollment_df["Study Treatment Administered"], inplace=True
        )
        enrollment_df = enrollment_df.drop(
            columns=[
                "Study Treatment Administered1",
                "Study Treatment Administered2",
            ]
        )
        #    print(enrollment_df["Subject"] + enrollment_df["Study Treatment Administered"])
        enrollment_df.loc[
            (enrollment_df["Study Treatment Administered"] != "Yes")
            & (enrollment_df["End of Study Date"].isnull()),
            "Study Treatment Administered",
        ] = "Pending"
        # enrollment_df.loc[
        #     (enrollment_df["Study Treatment Administered"] != "Yes")
        #     & (~enrollment_df["End of Study Date"].isnull()),
        #     "Study Treatment Administered",
        # ] = "No"
        enrollment_df["End of Study Date"] = pd.to_datetime(
            enrollment_df["End of Study Date"]
        )
        enrollment_df = enrollment_df.sort_values(["Subject", "End of Study Date"])
        enrollment_df = enrollment_df.drop_duplicates(subset=["Subject"], keep="last")
        enrollment_df = enrollment_df.drop(columns=["End of Study Date"])
        enrollment_df = enrollment_df.drop_duplicates()
        #   print(enrollment_df["Subject"] + enrollment_df["Study Treatment Administered"])
        ### TODO: Demo Stats Table
        # !Update this filter options to each cohort
        filter_options = [
            enrollment_df["Consent Date"].notna()
            | enrollment_df["Main Consent Date"].notna(),
            enrollment_df["Cohort Assignment"] == "Cohort 1",
        ]
        status_list = []
        LegalSex_list = []
        Age_at_Consent_list = []
        Race_list = []
        Ethnicity_list = []

        if debug:
            print(len(enrollment_df), enrollment_df.index)
            for filter_option in filter_options:
                print(len(filter_option), filter_option.index)
                print(filter_option.equals(enrollment_df.index))

        for filter_index, filter_option in enumerate(filter_options):
            # Apply the filter to the dataframe
            filtered_df = enrollment_df[filter_option].copy()
            filtered_df = filtered_df[
                (filtered_df["Consent Date"].notna())
                | (filtered_df["Main Consent Date"].notna())
            ]
            # Calculate the stats
            ## Total Consented
            TT_df = filtered_df.copy()
            TT = filtered_df["Subject"].count()
            ## Screen Failed
            SF_df = filtered_df[
                filtered_df["Subject meets all study eligibility?"]
                .fillna("")
                .astype(str)
                .str.strip()
                == "No"
            ].copy()
            SF = SF_df["Subject"].count()
            ## Eligible, convert to str and strip the space
            EL_df = filtered_df[
                filtered_df["Subject meets all study eligibility?"]
                .fillna("")
                .astype(str)
                .str.strip()
                == "Yes"
            ].copy()
            EL = EL_df["Subject"].count()
            ## Study Treatment Administered
            INF_df = filtered_df[
                filtered_df["Study Treatment Administered"]
                .fillna("")
                .astype(str)
                .str.strip()
                == "Yes"
            ].copy()
            INF = INF_df["Subject"].count()

            # Define a dictionary containing the status of each variable
            status_list.append(
                {
                    "Total Consented": TT,
                    "Screen Failed": SF,
                    "Eligible": EL,
                    "Study Treatment Administered": INF,
                }
            )

            # Calculate the stats for the filtered dataframe
            LegalSex_list.append(
                get_stats_percentage("Legal Sex", TT_df, SF_df, EL_df, INF_df)
            )
            Age_at_Consent_list.append(
                get_stats_df("Age at Consent", TT_df, SF_df, EL_df, INF_df)
            )
            Race_list.append(get_stats_percentage("Race", TT_df, SF_df, EL_df, INF_df))
            Ethnicity_list.append(
                get_stats_percentage("Ethnicity", TT_df, SF_df, EL_df, INF_df)
            )

        if debug:
            # 0: All Cohorts, 1: Cohort 1
            print(status_list)
            print(LegalSex_list)
            print(Age_at_Consent_list)
            print(Race_list)
            print(Ethnicity_list)

        # *: remove after calculating the stats
        enrollment_df = enrollment_df.drop(
            columns=["Consent Date", "Date of Birth", "Main Consent Date"]
        )

    if export:
        with pd.ExcelWriter(
            output_dir + "/" + output_file_name + ".xlsx", engine="xlsxwriter"
        ) as writer:
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
                if enrollment_df["Subject"].count() > 0:
                    # * WRITING DATA: LegalSex_list, Age_at_Consent_list, Race_list, Ethnicity_list
                    worksheet1 = writer.book.add_worksheet("DSMB-Demo Stats Table")

                    # * FORMAT DATA
                    for i in range(0, len(status_list)):
                        for j in range(0, len(LegalSex_list[i])):
                            for k in range(0, len(LegalSex_list[i].columns)):
                                worksheet1.write(
                                    j + 3,
                                    k + 1 + i * 4,
                                    LegalSex_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(Age_at_Consent_list[i])):
                            for k in range(0, len(Age_at_Consent_list[i].columns)):
                                worksheet1.write(
                                    j + 8,
                                    k + 1 + i * 4,
                                    Age_at_Consent_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(Race_list[i])):
                            for k in range(0, len(Race_list[i].columns)):
                                worksheet1.write(
                                    j + 12,
                                    k + 1 + i * 4,
                                    Race_list[i].iloc[j, k],
                                    normal_data_format,
                                )
                        for j in range(0, len(Ethnicity_list[i])):
                            for k in range(0, len(Ethnicity_list[i].columns)):
                                worksheet1.write(
                                    j + 22,
                                    k + 1 + i * 4,
                                    Ethnicity_list[i].iloc[j, k],
                                    normal_data_format,
                                )

                    # Apply the format to a range of cells
                    # worksheet1.set_column('B:I', None, normal_data_format)

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
                    Ethnicity_order = ["Hispanic", "Non-Hispanic", "Unknown"]

                    for i in range(0, len(Sex_order)):
                        worksheet1.write(i + 3, 0, Sex_order[i], bold_11_format)
                    for i in range(0, len(Age_order)):
                        worksheet1.write(i + 8, 0, Age_order[i], bold_11_format)
                    for i in range(0, len(Race_order)):
                        worksheet1.write(i + 12, 0, Race_order[i], bold_11_format)
                    for i in range(0, len(Ethnicity_order)):
                        worksheet1.write(i + 22, 0, Ethnicity_order[i], bold_11_format)

                    worksheet1.merge_range(
                        "B1:E1", "Overall Study Enrollment", bold_12_format
                    )
                    worksheet1.merge_range(
                        "F1:I1", "Cohort 1 Enrollment", bold_12_format
                    )
                    worksheet1.write(1, 0, "Status", bold_11_format)
                    for i in range(len(status_list)):
                        worksheet1.write(
                            1,
                            1 + i * 4,
                            "Total Consented\nN="
                            + str(status_list[i]["Total Consented"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            2 + i * 4,
                            "Screen Failed\nN=" + str(status_list[i]["Screen Failed"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            3 + i * 4,
                            "Eligible\nN=" + str(status_list[i]["Eligible"]),
                            bold_11_wrap_format,
                        )
                        worksheet1.write(
                            1,
                            4 + i * 4,
                            "Study Treatment Administered\nN="
                            + str(status_list[i]["Study Treatment Administered"]),
                            bold_11_wrap_format,
                        )

                    worksheet1.merge_range("A3:I3", "Legal Sex", bold_11_format)
                    worksheet1.merge_range("A8:I8", "Age at Consent", bold_11_format)
                    worksheet1.merge_range("A12:I12", "Race", bold_11_format)
                    worksheet1.merge_range("A22:I22", "Ethnicity", bold_11_format)
                    worksheet1.autofit()

                    ## TODO: Enrollment Listing
                    # * WRITING DATA: enrollment_df
                    worksheet2 = writer.book.add_worksheet("DSMB-Enrollment Listing")
                    # * WRITING HEADER AND FORMATTING
                    # Assuming 'enrollment_df' is your DataFrame
                    enrollment_df.replace(
                        [np.inf, -np.inf], np.nan, inplace=True
                    )  # Replace INF with NaN
                    enrollment_df.fillna(
                        "", inplace=True
                    )  # Replace NaN with a placeholder
                    for i in range(0, len(enrollment_df.columns)):
                        worksheet2.write(0, i, enrollment_df.columns[i], bold_11_format)
                    # * FORMAT DATA
                    for i in range(0, len(enrollment_df)):
                        for j in range(0, len(enrollment_df.columns)):
                            worksheet2.write(
                                i + 1, j, enrollment_df.iloc[i, j], normal_data_format
                            )
                    # Autofit
                    worksheet2.autofit()
