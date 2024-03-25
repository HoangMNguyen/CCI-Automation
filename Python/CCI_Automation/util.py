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

def clockify_sort_tasks(clockify_df, template_df):
    """sorting tasks of clockify df based on the template df

    Args:
        clockify_df (_type_): clockify dataframe
        template_df (_type_): template dataframe 

    Returns:
        dataframe : sorted dataframe
    """
    # Create a dictionary mapping task names to their position in the template list
    template_order = {task: i for i, task in enumerate(template_df['Task'])}
    
    # Define a custom sorting function that uses the template order
    def sort_key(task_name):
        """
        Custom sorting function for task names.

        Args:
            task_name (str): The name of the task to be sorted.

        Returns:
            int: The sorting index for the given task name.
        """
        if task_name in template_order:
            return template_order[task_name]
        else:
            return float('inf')
    
    # Compute the sort order using the custom sorting function
    sort_order = clockify_df['Task ID'].map(sort_key).argsort()
    
    # Sort the Clockify DataFrame using the computed sort order
    sorted_df = clockify_df.iloc[sort_order]
    
    return sorted_df

def clockify_get_task_df(api_key, workspace_id, project_name):
    """
    Retrieve a DataFrame of tasks from Clockify.

    Returns:
        DataFrame: A pandas DataFrame containing Clockify tasks data.
    """
    projects = clockify_get_dict_projects(api_key, workspace_id)        
    project_id = projects[project_name]

    headers = {
        'X-Api-Key': api_key,
        'Content-Type': 'application/json'
    }
    
    params = {'is-active' : 'true'}

    url = f'https://api.clockify.me/api/v1/workspaces/{workspace_id}/projects/{project_id}/tasks'

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        tasks = response.json()
        task_names = [task['name'] for task in tasks]
        df = pd.DataFrame({'Task ID': task_names})
        return df
    else:
        print(f"Error: {response.status_code}, {response.text}")

def clockify_get_project_id(api_key, workspace_id, project_name):
    """
    Fetch the project ID from Clockify.

    Returns:
        str: The project ID.
    """
    projects = clockify_get_dict_projects(api_key, workspace_id)        
    project_id = projects[project_name]
    return project_id
        
def clockify_get_dict_projects(api_key, workspace_id):
    """
    Retrieve a dictionary of projects from Clockify.

    Returns:
        dict: A dictionary where keys are project names and values are project details.
    """
    headers = {
        'content-type': 'application/json',
        'X-Api-Key': api_key
    }
    url = f'https://api.clockify.me/api/v1/workspaces/{workspace_id}/projects'
    
    response = requests.get(url, headers=headers)
    projects_data = response.json()
    
    # Create a dictionary mapping project names to project IDs
    projects_dict = {project['name']: project['id'] for project in projects_data}
    
    return projects_dict

def clockify_get_list_projects(api_key, workspace_id):
    """
    Fetch a list of projects from Clockify.

    Returns:
        list: A list of projects.
    """
    headers = {
        'content-type': 'application/json',
        'X-Api-Key': api_key
    }
    url = f'https://api.clockify.me/api/v1/workspaces/{workspace_id}/projects'
    
    params = {'archived': 'false'}
    response = requests.get(url, headers=headers, params=params)
    projects_data = response.json()
    
    # Create a dictionary mapping project names to project IDs
    projects_list = [project['name'] for project in projects_data if project['archived'] == False]
    
    return projects_list

def clockify_get_api_key():
    """
    Retrieve the API key for Clockify access.

    Returns:
        str: The API key.
    """
    api_key = 'ZDE0ZmVlNzAtMTFiZi00YzUzLWJhZjctM2Q0M2I3M2VhYjdh'
    return api_key

def clockify_get_workplace_id():
    """
    Obtain the workplace ID from Clockify.

    Returns:
        str: The workplace ID.
    """
    workspace_id = '63175caa4723ae23e827f33a'
    return workspace_id

def clockify_get_detailed_report(api_key, workspace_id, project_name):
    """
    Generate a detailed report from Clockify data.

    Returns:
        DataFrame: A pandas DataFrame containing the detailed report.
    """
    headers = {
        'content-type': 'application/json',
        'X-Api-Key': api_key
    }
    url = f'https://reports.api.clockify.me/v1/workspaces/{workspace_id}/reports/detailed'
    
    # Set the start date to the earliest possible date and the end date to today
    start_date = (datetime.utcnow() - timedelta(days=365)).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
    end_date = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
    project_id = None
    project_id = clockify_get_project_id(clockify_get_api_key(), clockify_get_workplace_id(), project_name)
    
    data = {
        'dateRangeEnd': end_date,
        'dateRangeStart': start_date,
        'detailedFilter': {
            'pageSize': 400
        },
        'projects': {
            'contains' : "CONTAINS_ONLY",
            'ids' : [project_id]
        },
        'exportType': 'JSON'
    }
    
    response = requests.post(url, headers=headers, data=json.dumps(data))
    report_data = response.json()
    
    # Print the report data for debugging purposes
    # print(report_data)
    
    # Convert the report data into a pandas DataFrame
    report_df = pd.json_normalize(report_data['timeentries'])
    if 'tags' in report_df.columns:
        report_df = report_df[['projectName', 'taskName', 'userName', 'tags', 'timeInterval.start', 'timeInterval.end']]
        report_df['tags'] = report_df['tags'].apply(lambda x: x[0]['name'] if isinstance(x, list) and len(x) > 0 and isinstance(x[0], dict) and 'name' in x[0] else None)
    else:
        report_df = report_df[['projectName', 'taskName', 'userName', 'timeInterval.start', 'timeInterval.end']]
        report_df['tags'] = np.nan
    # print(report_df['tags'].dtypes)
    report_df['timeInterval.start'] = pd.to_datetime(report_df['timeInterval.start'])
    report_df['timeInterval.end'] = pd.to_datetime(report_df['timeInterval.end'])
    report_df['Duration (decimal)'] = (report_df['timeInterval.end'] - report_df['timeInterval.start']) / pd.Timedelta(hours=1)
    
    report_df['Duration (h)'] = pd.to_timedelta(report_df['Duration (decimal)'], unit='h').apply(format_timedelta)
    report_df['Duration (decimal)'] = report_df['Duration (decimal)'].round(2)
    new_report_columns_name = {'projectName':'Project', 'taskName':'Task', 'userName': 'User', 'tags': 'Tags'}
    report_df = report_df.rename(columns=new_report_columns_name)
    report_df = report_df.drop('timeInterval.end', axis = 1)
    report_df = report_df.drop('timeInterval.start', axis = 1)
    return report_df

def format_timedelta(td):
    """
    Format a timedelta object into a human-readable string.

    Args:
        timedelta_obj (timedelta): The timedelta object to format.

    Returns:
        str: A formatted string representing the timedelta.
    """
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f'{hours:02d}:{minutes:02d}:{seconds:02d}'

def clockify_create_tasks(api_key, workspace_id, project_name):
    """
    Create tasks in Clockify for a given project.

    This function reads task names from a CSV file, constructs a list of tasks, and then uses the Clockify API
    to create these tasks under a specified project in a specified workspace.

    Args:
        api_key (str): The API key for authenticating with the Clockify API.
        workspace_id (str): The ID of the workspace in Clockify where the tasks will be created.
        project_name (str): The name of the project in Clockify under which the tasks will be created.

    Returns:
        list: A list containing the responses from the Clockify API for each task creation request.
    """
    current_dir = os.getcwd()
    template_tasks = pd.read_csv(os.path.join(current_dir, "Clockify/Tasks.csv"))
    task_list = template_tasks['Task'].tolist()
    # print(task_list)
    project_id = clockify_get_project_id(clockify_get_api_key(), clockify_get_workplace_id(), project_name)
    headers = {
        'content-type': 'application/json',
        'X-Api-Key': api_key
    }
    url = f'https://api.clockify.me/api/v1/workspaces/{workspace_id}/projects/{project_id}/tasks'
    for task_name in task_list:
        data = {
            'name': task_name
        }
        response = requests.post(url, headers=headers, json=data)
        print(response.json())

def read_data_dict_zip_corelisting(input_dir : str, cut_off_date = None) -> dict:
    """Read all data from csv files within the corelisting from a zip file into a dictionary of dataframes
    
    - Apply the cut off date if it's not None
    - Convert the subject ID to the format of 12345-67 instead of 100-12345-67
    - Filtered data to only submitted status
    
    Args:
        input_dir (string): string of input directory of the zip file
        cut_off_date (datetime): the cut off date for the data

    Returns:
        data (dict) : the dictionary of dataframe, with keys are form names
    
    """

    data : dict = {}
    if cut_off_date != None:
        cut_off_date = pd.to_datetime(cut_off_date)
    
    # Open the zip file
    with zipfile.ZipFile(input_dir, 'r') as z:
        for file_name in z.namelist():
            if file_name.endswith('.csv'):
                file_name_noCSV = file_name.split('.')[0]
                
                # Load the CSV file into a Pandas data frame
                with z.open(file_name) as f:
                    df = pd.read_csv(f)
                    #check if df has Event Date column and none of the data is blank
                    if 'Event Date' in df.columns and df['Event Date'].isnull().sum() != len(df['Event Date']):
                        df['Event Date'] = pd.to_datetime(df['Event Date'])
                        # if cut_off_date is not None, then filter the data based on the cut_off_date
                        if cut_off_date != None:
                            df = df[df['Event Date'] <= cut_off_date]
                    
                    # Replace '100-' prefix with an empty string
                    df['Subject'] = df['Subject'].str.replace('^100-', '', regex=True)
                    # only get data that is submitted status
                    df = df[df['Form Status'] == 'Submitted']
                data[file_name_noCSV.split("_")[-1]] = df
                
                
    return data

def add_rename_column_corelisting(current_df, input_data, form_name, column_name, new_column_name, key1 ='Subject', key2 = None):
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
        df = input_data[form_name][[key1, 'Form ILB Status', column_name]].copy()
    elif key2 != None:
        df = input_data[form_name][[key1, key2, 'Form ILB Status', column_name]].copy()
    df = df[df['Form ILB Status'] == False]
    new_col_name = {column_name:new_column_name}
    df = df.rename(columns=new_col_name)
    df = df.drop(columns = ['Form ILB Status'])
    if key2 == None:
        new_df = pd.merge(current_df, df, on=key1, how='left')
    elif key2 != None:
        new_df = pd.merge(current_df, df, on=[key1, key2], how='left')
    return new_df

def add_rename_column_df(current_df, input_df, form_name, column_name, new_column_name, key1 ='Subject', key2 = None):
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
    new_col_name = {column_name:new_column_name}
    df = df.rename(columns=new_col_name)
    # df = df.drop(columns = ['Form ILB Status'])
    if key2 == None:
        new_df = pd.merge(current_df, df, on=key1, how='left')
    elif key2 != None:
        new_df = pd.merge(current_df, df, on=[key1, key2], how='left')
    return new_df

def get_stats_df(column,*dfs):
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
        df.loc[:, column] = pd.to_numeric(df[column], errors='coerce')
        mean = df[column].mean()
        std = df[column].std()
        median = df[column].median()
        minimum = df[column].min()
        maximum = df[column].max()
        
        if np.isnan(std):
            std = 0

        # Create a DataFrame for the current stats if non of the stats are NaN
        if not np.isnan(mean) and not np.isnan(median) and not np.isnan(minimum) and not np.isnan(maximum):
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
            stats_df = pd.DataFrame({
                'Mean ± SD': mean_std,
                'Median': median,
                'Range': range
            }, index=[i])
        else:
            stats_df = pd.DataFrame(index=[i])
        # Append the stats DataFrame to the main DataFrame if it is not empty
        main_df = pd.concat([main_df, stats_df])
    # Transpose the DataFrame to switch the axes
    main_df = main_df.T
    return main_df

def get_stats_perc_df(column,*dfs):
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
        df.loc[:, column] = pd.to_numeric(df[column], errors='coerce')
        mean = df[column].mean()
        std = df[column].std()
        median = df[column].median()
        minimum = df[column].min()
        maximum = df[column].max()

        if np.isnan(std):
            std = 0

        # Create a DataFrame for the current stats if non of the stats are NaN
        if not np.isnan(mean) and not np.isnan(median) and not np.isnan(minimum) and not np.isnan(maximum):
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
            stats_df = pd.DataFrame({
                'Mean ± SD': mean_std,
                'Median': median,
                'Range': range
            }, index=[i])
        else:
            stats_df = pd.DataFrame(index=[i])
        # Append the stats DataFrame to the main DataFrame if it is not empty
        main_df = pd.concat([main_df, stats_df])
    # Transpose the DataFrame to switch the axes
    main_df = main_df.T
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
        return '0'
    else:
        exponent = int(math.floor(math.log10(abs(n))))
        mantissa = n / 10**exponent
        if mantissa.is_integer():
            mantissa = int(mantissa)
            return f'{mantissa}x10^{exponent}'
        else:
            return f'{mantissa:.2f}x10^{exponent}'

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
        parts = s.split('x10^')
        mantissa = float(parts[0])
        exponent = float(parts[1])
        return mantissa * 10 ** exponent
    
    except ValueError:
        return np.nan

import pandas as pd

def convert_integers_to_strings(df, column_name):
    """
    This function takes a dataframe and a column name as input.
    It converts all integers in the specified column to strings, while keeping strings unchanged.
    
    :param df: Pandas DataFrame
    :param column_name: Name of the column to process
    :return: DataFrame with the updated column
    """
    # Check if column exists in dataframe
    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' does not exist in the dataframe")

    # Convert integers to strings
    df[column_name] = df[column_name].apply(lambda x: str(x) if isinstance(x, int) else x)

    return df


#if __name__ == '__main__':
    # api_key = 'ZDE0ZmVlNzAtMTFiZi00YzUzLWJhZjctM2Q0M2I3M2VhYjdh'
    # workspace_id = '63175caa4723ae23e827f33a'
    # project_name = '15420 Amendment V4'

    # dict_cl = clockify_get_dict_projects(api_key, workspace_id)
    # clockify_create_tasks('ZDE0ZmVlNzAtMTFiZi00YzUzLWJhZjctM2Q0M2I3M2VhYjdh', '63175caa4723ae23e827f33a', 'test')