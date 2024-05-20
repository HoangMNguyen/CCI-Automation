import pandas as pd
import os
import re


def change_type(input_path, output_dir, output_file_name):
    # Check if the input path exists
    if not os.path.exists(input_path):
        raise FileNotFoundError("The specified input path does not exist.")
    else:
        # Determine the file extension
        _, file_extension = os.path.splitext(input_path)

    # Load the file based on its extension
    if file_extension.lower() == ".csv":
        data = pd.read_csv(input_path)
    else:
        raise ValueError("Unsupported file type. Please provide a'.csv' file.")

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(
        output_dir + "/" + output_file_name + ".xlsx", engine="xlsxwriter"
    )

    # Convert the dataframe to an XlsxWriter Excel object
    data.to_excel(writer, index=False, sheet_name="Sheet1")

    # Get the xlsxwriter workbook and worksheet objects
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]

    # Create a format to set matched columns to text type
    text_format = workbook.add_format({"num_format": "@"})

    # Function to check if any cell in the column matches the specific pattern
    def has_pattern(series):
        return series.apply(lambda x: bool(re.match(r"\d+-\d+", str(x)))).any()

    # Apply formatting only to columns that contain the pattern
    for col_num, col_name in enumerate(data.columns):
        if has_pattern(data[col_name]):
            worksheet.set_column(col_num, col_num, None, text_format)

    # Close the Pandas Excel writer and output the Excel file
    writer.close()
