"""
Process Shastabins_ for alternatives that have them.
Uses output DSS_contents.xlsx file from the DSS Reader
"""

import pandas as pd
# Excel file with SHASTABIN_ timeseries from Calsim Runs.
# Must include all relevant scenarios.
# This is the output from reading in the dss file using the dssreader script.
s_input_data = r""
# output file
s_output = r"..\inputs\shasta_bin_info.xlsx"

#Read in the dss reader excel file and subset the dataframe to only include relevant columns
dss_pathb = ['SHASTABIN_']
sl_columns = ['Date', 'Month', 'WY', 'Scenario']
sl_columns.extend(dss_pathb)

#Read in monthly timeseries for SHASTABIN_
di_monthly = pd.read_excel(s_input_data, index_col = 0)[sl_columns]

#Subset timeseries to only the final shastabin_, which is set in May.
#Note that the final SHASTABIN_ type set in May of the calendar year is used to determine temperature logic for the full calendar year.
#Ex: final May 1990 bin type is used from Jan 1, 1990 to Dec 31, 1990.
di_annual = di_monthly.loc[di_monthly.Month ==5]
di_annual['calendar_yr'] =di_annual.Date.dt.year
di_annual = di_annual[['calendar_yr', 'SHASTABIN_', "Scenario"]]
di_annual.set_index('calendar_yr', inplace = True)

#Save shastabin_ data to excel file.
di_annual.to_excel(s_output)