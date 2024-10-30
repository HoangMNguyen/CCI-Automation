import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import os
import zipfile
import pandas as pd


def read_data_dict_zip_corelisting(input_dir: str, cut_off_date=None) -> dict:
    """Read all data from csv files within the corelisting from a zip file into a dictionary of dataframes

    - Apply the cut off date if it's not None
    - Convert the subject ID to the format of 12345-67 instead of 100-12345-67
    - Filtered data to only submitted status
    - Filtered data to only non-ILB status

    Args:
        input_dir (string): string of input directory of the zip file
        cut_off_date (datetime): the cut off date for the data

    Returns:
        data (dict) : the dictionary of dataframe, with keys are form names

    """

    data: dict = {}
    if cut_off_date is not None:
        cut_off_date = pd.to_datetime(cut_off_date)

    # Open the zip file
    with zipfile.ZipFile(input_dir, "r") as z:
        for file_name in z.namelist():
            if file_name.endswith(".csv"):
                file_name_noCSV = file_name.split(".")[0]

                # Load the CSV file into a Pandas data frame
                with z.open(file_name) as f:
                    df = pd.read_csv(f)
                    # check if df has Event Date column and none of the data is blank
                    if (
                        "Event Date" in df.columns and df["Event Date"].isnull().sum() != len(df["Event Date"])
                    ) or "EOS" in file_name_noCSV.split("_")[-1]:
                        df["Event Date"] = pd.to_datetime(df["Event Date"])
                        # if cut_off_date is not None, then filter the data based on the cut_off_date
                        if cut_off_date is not None and "EOS" not in file_name_noCSV.split("_")[-1]:
                            df = df[df["Event Date"] <= cut_off_date]
                        elif cut_off_date is not None and "EOS" in file_name_noCSV.split("_")[-1]:
                            # find the column header that contains "End of Study Date"
                            column_name = df.columns[df.columns.str.contains("End of Study Date", case=False)][0]
                            # Convert the relevant column to datetime
                            df[column_name] = pd.to_datetime(df[column_name])
                            # filter out the data that is after the cut off date. If the date is null, keep it
                            df = df[(df[column_name] <= cut_off_date) | (df[column_name].isnull())]

                    # Replace '100-' prefix with an empty string
                    df["Subject"] = df["Subject"].str.replace("^100-", "", regex=True)
                    # only get data that is submitted status
                    df = df[df["Form Status"] == "Submitted"]
                    # only get data that is not ILB status
                    df = df[df["Form ILB Status"] == False]
                data[file_name_noCSV.split("_")[-1]] = df

    return data


def get_dose_level(dose, power, dose_level_mapping):
    """
    Helper function to compute dose level administered.

    Parameters:

    Returns:
    str: The correct Dose Level (e.g., DL1, DL2, etc.) or 'Not Matching' if no match found.
    """
    # Calculate the actual dose administered
    actual_dose = dose * (10**power)

    # Determine the correct dose level based on the mapping provided
    for dose_level, (p_dose, p_power) in dose_level_mapping.items():
        if actual_dose == p_dose * (10**p_power):
            return dose_level

    # if no match found, return 'Not Matching'
    return "Not Matching"


def get_study_name(input_dir) -> str:
    """
    Extracts the study name from the input directory path.

    The function assumes that the study name is located in the third
    segment of the folder name after splitting it by underscores ('_').
    If the extracted segment is not a digit, an error message is printed.

    Returns:
        str: The study name if it is a digit; otherwise, None.
    """
    # Extract the folder name from the input directory path
    folder_name: str = input_dir.split("/")[-1]

    # Extract the study name from the folder name
    study_name: str = folder_name.split("_")[2][0:5]

    # Check if the extracted study name is a digit
    if not study_name.isdigit():
        print(f"It's not a study name: {study_name}")
        return None
    else:
        return study_name
