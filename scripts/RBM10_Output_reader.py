import os, subprocess
import shutil
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
import time


def convert_to_numeric(s):
    # todo: doc string
    # todo: comments
    num = float(s)
    if num.is_integer():
        return int(num)
    return num

 # Function to convert decimal year to datetime
def decimal_year_to_datetime(decimal_year):
    #todo: doc string

    year = int(decimal_year)  # Extract the integer part (year)
    remainder = decimal_year - year  # Extract the decimal part

    # Check if the year is a leap year
    if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
        days_in_year = 366
    else:
        days_in_year = 365

    # Convert the remainder to days and time
    days = remainder * days_in_year
    date_time = datetime(year, 1, 1) + timedelta(days=days)
    # date.replace(hour=0, minute=0, second=0, microsecond=0)
    return date_time.replace(hour=0, minute=0, second=0, microsecond=0)


def decimal_year_to_date(dy):
    year = int(dy)
    remainder = dy - year
    start = datetime(year, 1, 1)
    end = datetime(year + 1, 1, 1)
    delta = end - start
    return start + timedelta(days=remainder * delta.days)

def overwrite_excel_with_df(filename, df, sheet_name='Sheet1'):
    """
    Opens an Excel file with the given filename, clears all data,
    and overwrites it with the contents of the provided DataFrame.

    Parameters:
        filename (str): Name of the Excel file (e.g., 'data.xlsx').
        df (pd.DataFrame): DataFrame to write into the Excel file.
        sheet_name (str): Name of the Excel sheet (default is 'Sheet1').
    """
    # Optional: check if file exists before writing (can also skip this)
    if os.path.exists(filename):
        print(f"Overwriting existing file: {filename}")
    else:
        print(f"Creating new file: {filename}")

    # Write the new DataFrame to the file, overwriting existing content
    with pd.ExcelWriter(filename, engine='openpyxl', mode='w') as writer:
        df.to_excel(writer, index=True, sheet_name=sheet_name)

    print(f"File '{filename}' has been updated.")


def read_output(s_RBMfilelocations,s_scenario_labels,s_excel_file_location):
    """
    Read simulation output files (fort.41) from River Basin Model runs, extract and convert
    temperature data, and export the results to an Excel file.

    Parameters
    ----------
    s_RBMfilelocations : list of str
        List of file paths to each scenario's working directory, where 'fort.41' files are located.

    s_scenario_labels : list of str
        List of labels for each scenario. Must be the same length as `s_RBMfilelocations`.

    s_excel_file_location : str
        Path to the output Excel file where processed results will be written.

    Returns
    -------
    None
        This function does not return a value. It writes the processed DataFrame to an Excel file.

    """

    # define the number of scenarios
    n_scenarios = len(s_scenario_labels)
    for i in range(n_scenarios):
        # working directory for the RBM10 results
        s_working_path = s_RBMfilelocations[i]
        # scenario name
        scenario_name = s_scenario_labels[i]

        # combine the strings fro the working directory and file name
        output_data_path = os.path.join(s_working_path, "fort.41")
        with open(output_data_path, 'r') as file:
            lines = file.readlines()

        # Put the temperature and flow data in the RBM10 file in a dataframe
        sf_result_dataframe = pd.DataFrame(lines, columns=['values'])
        sf_result_dataframe['values'] = sf_result_dataframe['values'].str.replace(r'\s+', ' ', regex=True)
        sf_split_dataframe = sf_result_dataframe['values'].str.split(expand=True)


        # Convert all strings to numeric values
        for column in sf_split_dataframe.columns:
            sf_split_dataframe[column] = sf_split_dataframe[column].apply(convert_to_numeric)

        df_split_dataframe = sf_split_dataframe

        # Assign names to the columns in the dataframe
        i_num_column = int((df_split_dataframe.shape[1] - 1) / 3)
        sl_name1 = ['Rivermile', 'Tmean', 'Qmean'] * i_num_column
        il_row = df_split_dataframe.iloc[0, range(1, df_split_dataframe.shape[1], 3)].tolist()
        il_name2 = [s for s in il_row for _ in range(3)]
        sl_column_name = ['Date'] + [f"{name}{number}" for name, number in zip(sl_name1, il_name2)]
        df_split_dataframe.columns = sl_column_name

        # Set the 'Date' column as the index
        df_split_dataframe.set_index('Date', inplace=True)

        # Select the columns that show the temperature at Douglas city and NF Trinity and Rename the columns
        # df_split_dataframe1 = df_split_dataframe[['Tmean92.6', 'Qmean92.6',
        #                                           'Tmean72.6', 'Qmean72.6',
        #                                           'Tmean'
        #                                           ]]

        #Select columns for temperatures at RBM10 node 0.5, 31.6, 72.6, 92.6, and 112. (For BDO biologists)
        df_split_dataframe1 =df_split_dataframe[
            ['Tmean92.6', 'Qmean92.6', 'Tmean72.6', 'Qmean72.6', 'Tmean0.5', 'Qmean0.5', 'Tmean31.6', 'Qmean31.6',
             'Tmean112.0', 'Qmean112.0']]

        # Rename columns
        df_split_dataframe1 = df_split_dataframe1.rename(columns={
                                                                    'Tmean92.6': 'temperature_douglas_city',
                                                                    'Qmean92.6': 'flow_douglas_city',
                                                                    'Tmean72.6': 'temperature_north_fork_trinity',
                                                                    'Qmean72.6': 'flow_north_fork_trinity',
                                                                    'Tmean0.5':"temperature_0.5",
                                                                    'Qmean0.5' :"flow_0.5",
                                                                    "Tmean31.6":"temperature_31.6",
                                                                    "Qmean31.6": "flow_31.6",
                                                                    'Tmean112.0': "temperature_112.0",
                                                                    'Qmean112.0': "flow_112.0",
                                                                })
        # Subset the temperature columns
        sl_temperature_columns = ["temperature_0.5", "temperature_31.6",'temperature_douglas_city','temperature_north_fork_trinity',"temperature_112.0"]
        df_split_dataframe2 = df_split_dataframe1[sl_temperature_columns]

        # convert the temperature data from deg C to deg F
        for s_col in sl_temperature_columns:
            df_split_dataframe2[s_col] = df_split_dataframe1[s_col]*(9/5) + 32

        # Create columns for year, month and day from the date time index
        #df_split_dataframe2['datetime'] = df_split_dataframe2.index.to_series().apply(decimal_year_to_datetime)
        aa = df_split_dataframe2.index.to_series().apply(decimal_year_to_datetime)

        # Extract year, month, day
        df_split_dataframe2['Date'] = df_split_dataframe2.index.to_series().apply(decimal_year_to_date).dt.strftime('%Y-%m-%d')
        df_split_dataframe2['Scenario'] = scenario_name
        df_split_dataframe2['Year'] = aa.dt.year
        df_split_dataframe2['Month'] = aa.dt.month
        df_split_dataframe2['Day'] = aa.dt.day

        # Compute water year (starts Oct 1, ends Sep 30)
        df_split_dataframe2['WY'] = aa.apply(
        lambda dt: dt.year + 1 if dt.month >= 10 else dt.year
         )

        # Compute delivery contract year (starts Mar 1, ends Feb 29)
        df_split_dataframe2['DY'] = aa.apply(
        lambda dt: dt.year if dt.month >= 3 else dt.year - 1
        )

        # Reset the index to row numbers
        df_split_dataframe2 = df_split_dataframe2.reset_index(drop=True)
        df_split_dataframe2.index.name = "Index"

        if i == 0:
            sl_all_columns = ['Date','Scenario','Year','Month','Day','WY','DY']
            sl_all_columns.extend(sl_temperature_columns)
            df_split_dataframe3 = df_split_dataframe2[sl_all_columns]
        else:
            sl_all_columns = ['Date', 'Scenario', 'Year', 'Month', 'Day', 'WY', 'DY']
            sl_all_columns.extend(sl_temperature_columns)
            df_split_dataframe3 = pd.concat([df_split_dataframe3, df_split_dataframe2[sl_all_columns]], ignore_index=True)

    overwrite_excel_with_df(s_excel_file_location, df_split_dataframe3)

    return df_split_dataframe3

if __name__ == "__main__":
    ############## ################## User Input Needed ###################################
    # RBM10 Runs
    runs = [
        ["Baseline",
         (r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-07 naa\full\rbm10")],
        ["Alt 1", (
            r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-07 trinity alt1\full\rbm10")],
        ["Alt 2a", (
            r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-31 Alt2a_2022MED_SLR15_03302025\full\rbm10")],
        ["Alt 2b", (
            r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-11 Alt2b_121924_flowadj16\full\rbm10")],
        ["Alt 3", (
            r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-31 Alt3_2022MED_SLR15_03302025\full\rbm10")],
        ["Alt 4",
         (r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-06 alt4\full\rbm10")],
        ["Alt 6", (
            r"C:\Users\cyu\trinity_hec5q_github\alt6_20250508\2025-05-09 Alt6_wTUCP_2022MED_CCWD\full\rbm10")],
        ["Alt 7", (
            r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\2025-03-14 alt7_02272025_numerical_revision\full\rbm10")],
    ]

    s_scenario_labels, s_RBMfilelocations = zip(*runs)
    s_excel_file_location = r"C:\calsim_gits\dss_reader_git\calsim_dss_reader\DSS_contents_rbm10_additionalNodes.xlsx"
        #r"C:\Users\cyu\trinity_hec5q_github\alt_temperature_runs\DSS_contents_Alt1_NAA.xlsx"

    o_rbm10_results = read_output(s_RBMfilelocations,s_scenario_labels,s_excel_file_location)

################################# No user input needed after this ######################################




