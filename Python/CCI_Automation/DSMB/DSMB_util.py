import pandas as pd
import requests
import json
from datetime import datetime, timedelta
import numpy as np
import os
import math
import zipfile
from typing import Optional, List, Dict, Tuple, Union

# TODO: implement type hints for all functions


def add_rename_column_corelisting(
    current_df,
    input_data,
    form_name,
    column_name,
    new_column_name,
    key1="Subject",
    key2=None,
):
    """Adding a new column (merge left) to the current dataframe with an updated new name from the data of corelisting

    Args:
        current_df (dataframe): current dataframe with subject ID list
        input_data (dict): dictionary of dataframe, with keys are form names
        form_name (string): form name
        column_name (string): exact column name in corelisting
        new_column_name (string): the name of the column that would be updated
        key (string): usually it's subject ID of the

    Returns:
        new_df (dataframe): the new dataframe that is added with the new column and renamed
    """
    if key2 == None:
        df = input_data[form_name][[key1, "Form ILB Status", column_name]].copy()
    elif key2 != None:
        df = input_data[form_name][[key1, key2, "Form ILB Status", column_name]].copy()
    df = df[df["Form ILB Status"] == False]
    new_col_name = {column_name: new_column_name}
    df = df.rename(columns=new_col_name)
    df = df.drop(columns=["Form ILB Status"])
    if key2 == None:
        new_df = pd.merge(current_df, df, on=key1, how="left")
    elif key2 != None:
        new_df = pd.merge(current_df, df, on=[key1, key2], how="left")
    return new_df


def add_rename_column_df(
    current_df,
    input_df,
    form_name,
    column_name,
    new_column_name,
    key1="Subject",
    key2=None,
):
    """Adding a new column (merge left) to the current dataframe with an updated new name from the data of corelisting

    Args:
        current_df (dataframe): current dataframe with subject ID list
        input_df (dataframe):dataframe, with keys are form names
        form_name (string): form name
        column_name (string): exact column name in corelisting
        new_column_name (string): the name of the column that would be updated
        key (string): usually it's subject ID of the

    Returns:
        new_df (dataframe): the new dataframe that is added with the new column and renamed
    """
    if key2 == None:
        df = input_df[[key1, column_name]].copy()
    elif key2 != None:
        df = input_df[[key1, key2, column_name]].copy()
    # df = df[df['Form ILB Status'] == False]
    # df = df[df['Form Status'] != 'Blank']
    new_col_name = {column_name: new_column_name}
    df = df.rename(columns=new_col_name)
    # df = df.drop(columns = ['Form ILB Status'])
    if key2 == None:
        new_df = pd.merge(current_df, df, on=key1, how="left")
    elif key2 != None:
        new_df = pd.merge(current_df, df, on=[key1, key2], how="left")
    return new_df


def get_stats_percentage(column, *args):
    """
    Compute the percentage of each category in a specified column across multiple DataFrames.
    If the sum of the counts for a category is 0, the percentage is 0.0%.
    No
    Returns:
        dataframe: merged dataframe for stats
    """
    if column == "Legal Sex":
        groupby = ["Male", "Female", "X (Nonbinary)", "Not Reported"]
    elif column == "Race":
        groupby = [
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
    elif column == "Ethnicity":
        groupby = ["Hispanic", "Non-Hispanic", "Unknown"]
    elif column == "AE":
        groupby = ["Y", "N"]
    elif column == "SAE":
        groupby = ["Y", "N"]
    # NHL for 15420
    elif column == "PET-Based Response" or column == "PET-Based ORR":
        groupby = [
            "Complete Metabolic Response (CMR)",
            "Partial Metabolic Response (PMR)",
            "No Metabolic Response (NMR)",
            "Indeterminate Response (IR)",
            "Progressive Metabolic Disease (PMD)",
            "Not Reported",
        ]
    elif column == "CT-Based Response" or column == "CT-Based ORR":
        groupby = [
            "Complete Radiologic Response (CR)",
            "Partial Response (PR)",
            "Stable Disease (SD)",
            "Indeterminate Response (IR)",
            "Progressive Disease (PD)",
            "Not Reported",
        ]
    # NHL for 12423
    elif column == "PET-Based NHL Response" or column == "PET-Based NHL ORR":
        groupby = [
            "Complete Metabolic Response (CMR)",
            "Partial Metabolic Response (PMR)",
            "No Metabolic Response (NMR)",
            #     "Indeterminate Response (IR)",
            "Progressive Metabolic Disease (PMD)",
            "Not Reported",
        ]
    elif column == "CT-Based NHL Response" or column == "CT-Based NHL ORR":
        groupby = [
            "Complete Radiologic Response (CR)",
            "Partial Response (PR)",
            "Stable Disease (SD)",
            #     "Indeterminate Response (IR)",
            "Progressive Disease (PD)",
            "Not Reported",
        ]
    # CLL
    elif column == "OV-Best Response" or column == "Overall Response":
        groupby = [
            "Complete Remission (CR)",
            "Complete Remission with Incomplete Marrow Recovery (CRi)",
            "Partial Remission (PR)",
            "Stable Disease (SD)",
            "Progressive Disease (PD)",
            "Not Reported",
        ]
    elif column == "BM-Best Response" or column == "Bone Marrow Response":
        groupby = [
            "Complete Remission (CR)",
            "Partial Remission (PR)",
            "Progressive Disease (PD)",
            "Stable Disease (SD)",
            "Not Reported",
        ]
    # ALL
    elif column == "Best Overall Response" or column == "OV ORR":
        groupby = [
            "Complete Remission (CR)",
            "Complete Remission with Incomplete Blood Count Recovery (CRi)",
            "Complete Remission with Residual Mediastinal Disease (CRu)",
            "Treatment Failure (TF)",
            "Relapsed Disease (RD)",
            "Not Reported",
        ]
    elif column == "Best ED Response" or column == "ED ORR":
        groupby = [
            "Complete Remission (CR)",
            "Partial Remission (PR)",
            "Stable Disease (SD)",
            "Indeterminate Response",
            "Progressive Disease (PD)",
            "Not Reported",
        ]
    else:
        groupby = None
    # create a dataframe to store the stats
    main_df = pd.DataFrame()

    for arg in args:
        if groupby != None:
            temp_df = (
                arg.groupby(column)
                .agg({"Subject": "count"})
                .reindex(groupby, fill_value=0)
            )
        else:
            temp_df = arg.groupby(column).agg({"Subject": "count"})
        # combine the percentage and count column into 1 column for variables different than 0
        temp_df = temp_df.apply(
            lambda x: x.astype(str)
            + " ("
            + (x / sum(x)).apply(lambda y: "{:.1%}".format(y))
            + ")"
            if x.sum() != 0
            else x.astype(str) + " (0.0%)"
        )
        # merge the temp_df to the main df
        main_df = pd.concat([main_df, temp_df], axis=1)
    return main_df


def get_stats_df(column, *dfs):
    """
    Compute and aggregate statistical measures for a specified column across multiple DataFrames.

    This function calculates various statistical measures for a given column name that is present
    across multiple pandas DataFrames. The statistics are calculated for each DataFrame and then aggregated.

    Args:
        column (str): The name of the column for which statistical measures are to be calculated.
        *dfs (tuple of DataFrame): Variable number of DataFrame objects on which statistics are to be computed.

    Returns:
        DataFrame: A pandas DataFrame aggregating the calculated statistical measures from each input DataFrame.
    """
    # Initialize an empty DataFrame
    main_df = pd.DataFrame()

    for i, df in enumerate(dfs):
        # Calculate the mean, standard deviation, median, and range of 'Age at Consent'
        df.loc[:, column] = pd.to_numeric(df[column], errors="coerce")
        mean = df[column].mean()
        std = df[column].std()
        median = df[column].median()
        minimum = df[column].min()
        maximum = df[column].max()

        if np.isnan(std):
            std = 0

        # Create a DataFrame for the current stats if non of the stats are NaN
        if (
            not np.isnan(mean)
            and not np.isnan(median)
            and not np.isnan(minimum)
            and not np.isnan(maximum)
        ):
            # Format the mean and standard deviation
            if mean < 100:
                mean_std = f"{mean:.2f} ({std:.2f})"
            else:
                mean_std = f"{convert_float_2_sci_notation(int(mean))} ({convert_float_2_sci_notation(int(std))})"
            if median > 100:
                median = f"{convert_float_2_sci_notation(int(median))}"
            if minimum < 100 and maximum < 100:
                range = f"{minimum:.2f} - {maximum:.2f}"
            else:
                range = f"{convert_float_2_sci_notation(int(minimum))} - {convert_float_2_sci_notation(int(maximum))}"
            stats_df = pd.DataFrame(
                {"Mean ± SD": mean_std, "Median": median, "Range": range}, index=[i]
            )
        else:
            stats_df = pd.DataFrame(index=[i])
        # Append the stats DataFrame to the main DataFrame if it is not empty
        main_df = pd.concat([main_df, stats_df])
    # Transpose the DataFrame to switch the axes
    main_df = main_df.T
    # if NaN, replace with '0 (0.0%)'
    main_df = main_df.fillna("0 (0.0%)")
    return main_df


def get_stats_perc_df(column, *dfs):
    """
    Compute and aggregate percentage-based statistical measures for a specified column across multiple DataFrames.

    This function calculates statistics related to the percentage distribution of values for a given column
    that exists across multiple pandas DataFrames. It computes these statistics for each DataFrame individually
    and then aggregates the results.

    Args:
        column (str): The name of the column for which percentage-based statistical measures are to be calculated.
        *dfs (tuple of DataFrame): A variable number of DataFrame objects for which the statistics are to be computed.

    Returns:
        DataFrame: A pandas DataFrame aggregating the percentage-based statistical measures from each input DataFrame.
    """
    # Initialize an empty DataFrame
    main_df = pd.DataFrame()

    for i, df in enumerate(dfs):
        # Calculate the mean, standard deviation, median, and range of 'Age at Consent'
        df.loc[:, column] = pd.to_numeric(df[column], errors="coerce")
        mean = df[column].mean()
        std = df[column].std()
        median = df[column].median()
        minimum = df[column].min()
        maximum = df[column].max()

        if np.isnan(std):
            std = 0

        # Create a DataFrame for the current stats if non of the stats are NaN
        if (
            not np.isnan(mean)
            and not np.isnan(median)
            and not np.isnan(minimum)
            and not np.isnan(maximum)
        ):
            # Format the mean and standard deviation
            if mean < 100:
                mean_std = f"{mean:.2f}% ({std:.2f}%)"
            else:
                mean_std = f"{convert_float_2_sci_notation(int(mean))}% ({convert_float_2_sci_notation(int(std))}%)"
            if median > 100:
                median = f"{convert_float_2_sci_notation(int(median))}%"
            else:
                median = f"{median:.2f}%"
            if minimum < 100 and maximum < 100:
                range = f"{minimum:.2f}% - {maximum:.2f}%"
            else:
                range = f"{convert_float_2_sci_notation(int(minimum))}% - {convert_float_2_sci_notation(int(maximum))}%"
            stats_df = pd.DataFrame(
                {"Mean ± SD": mean_std, "Median": median, "Range": range}, index=[i]
            )
        else:
            stats_df = pd.DataFrame(index=[i])
        # Append the stats DataFrame to the main DataFrame if it is not empty
        main_df = pd.concat([main_df, stats_df])
    # Transpose the DataFrame to switch the axes
    main_df = main_df.T
    return main_df

    """

    Returns:
        dataframe: merged dataframe for Race stats
    """
    # create a dataframe to store the stats
    main_df = pd.DataFrame()
    # Define the order of Race categories
    race_order = [
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
    for arg in args:
        temp_df = (
            arg.groupby("Race")
            .agg({"Subject": "count"})
            .reindex(race_order, fill_value=0)
        )
        # combine the percentage and count column into 1 column for variables different than 0
        temp_df = temp_df.apply(
            lambda x: x.astype(str)
            + " ("
            + (x / sum(x)).apply(lambda y: "{:.1%}".format(y))
            + ")"
            if x.sum() != 0
            else x.astype(str)
        )
        # merge the temp_df to the main df
        main_df = pd.concat([main_df, temp_df], axis=1)
    return main_df


def convert_float_2_sci_notation(n):
    """
    Convert a floating-point number to its scientific notation as a string.

    This function takes a floating-point number and converts it into a string that represents its scientific notation.
    Scientific notation expresses the number as a mantissa multiplied by 10 raised to an exponent. The mantissa is
    rounded to two decimal places unless it is an integer. If the input number is 0, it simply returns '0'.

    Args:
        n (float): The floating-point number to be converted into scientific notation.

    Returns:
        string: A string representing the scientific notation of the input number. The mantissa is formatted to
                have two decimal places unless it is an integer, followed by 'x10' and the exponent.
    """
    if n == 0:
        return "0"
    else:
        exponent = int(math.floor(math.log10(abs(n))))
        mantissa = n / 10**exponent
        if mantissa.is_integer():
            mantissa = int(mantissa)
            return f"{mantissa}x10^{exponent}"
        else:
            return f"{mantissa:.2f}x10^{exponent}"


def convert_sci_notation_2_float(s):
    """
    Convert a string representing a number in scientific notation to a floating-point number.

    This function takes a string in the format 'mantissax10^exponent', where the mantissa and exponent are
    separated by 'x10^', and converts it into a floating-point number. If the string is not in the correct format
    or contains values that cannot be converted to floats, the function handles the error and returns NaN.

    Args:
        s (str): A string representing a number in scientific notation.

    Returns:
        float: The floating-point equivalent of the input string. If the string is not properly formatted or
               contains invalid values, NaN (Not a Number) is returned.
    """
    try:
        parts = s.split("x10^")
        mantissa = float(parts[0])
        exponent = float(parts[1])
        return mantissa * 10**exponent

    except ValueError:
        return np.nan


# if __name__ == '__main__':
# api_key = 'ZDE0ZmVlNzAtMTFiZi00YzUzLWJhZjctM2Q0M2I3M2VhYjdh'
# workspace_id = '63175caa4723ae23e827f33a'
# project_name = '15420 Amendment V4'

# dict_cl = clockify_get_dict_projects(api_key, workspace_id)
# clockify_create_tasks('ZDE0ZmVlNzAtMTFiZi00YzUzLWJhZjctM2Q0M2I3M2VhYjdh', '63175caa4723ae23e827f33a', 'test')
