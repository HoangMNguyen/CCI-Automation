import pytest
from EnrollmentLog.EnrollmentLog12423 import EnrollmentLog12423
from util import *
import os

current_dir = os.path.dirname(os.path.abspath(__file__))

# Define a fixture for setup procedures
@pytest.fixture
def setup_data_1():
    input_path = os.path.join(current_dir, 'data', 'Core_Listings_12423-TmCD19-IL18_2024_04_10_13_09_EDT.zip')
    valid_input = read_data_dict_zip_corelisting(input_path)
    actual_output = EnrollmentLog12423(valid_input)
    output_path = os.path.join(current_dir, 'data', '240410-12423 Enrollment Log.csv')
    # Read expected output from .csv file into pandas dataframe
    expected_output = pd.read_csv(output_path, keep_default_na=False, na_values=[''])
    return expected_output, actual_output

def test_headers(setup_data_1):
    expected_output, actual_output = setup_data_1
    # Test that the DataFrames have the same columns (header)
    assert list(expected_output.columns) == list(actual_output.columns)

def test_row_numbers(setup_data_1):
    expected_output, actual_output = setup_data_1
    # Test that the DataFrames have the same number of rows
    assert len(expected_output) == len(actual_output)

def test_values(setup_data_1):
    expected_output, actual_output = setup_data_1
    # Test each row and column
    for (index1, row1), (index2, row2) in zip(expected_output.iterrows(), actual_output.iterrows()):
        for col in expected_output.columns:
            val1 = row1[col]
            val2 = row2[col]
            assert val1 == val2 or (pd.isnull(val1) and pd.isnull(val2)), f"Difference at row {index1 + 2}, column {col}"