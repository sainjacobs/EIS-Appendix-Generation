# -*- coding: utf-8 -*-
"""
Created on Fri May 16 09:46:37 2025

Process WYTypes to create the wyflags.xlsx file. 
Uses output DSS_contents.xlsx file from the DSS Reader 
(Fields: WYT_TRIN_, WYT_SAC_, WYT_SJR_).
For WY 1921, grab Oct-1921's WYType.

@author: cyu
"""

import os, sys
import pandas as pd
from datetime import datetime

s_dvfile = r"C:\calsim_gits\dss_reader_git\calsim_dss_reader\DSS_contents_wyt.xlsx"
s_output = r"..\inputs\wy_flags.xlsx"

#Dictionary specifying the dss part b name containing wytype information, the 
#corresponding name to give to the final dataset column, and the month in which 
#the wytype is set (These are CALENDAR months, so 10 = Oct.)
sc_wytypes = {'WYT_TRIN_': ("TRIN",4), 
              "WYT_SAC_": ("40-30-30",5), 
              "WYT_SJR_": ("60-20-20",5), 
              
              }

#Read in the dss reader excel file and subset the dataframe to only include relevant columns
dss_pathb = [k for k in sc_wytypes.keys()]
sl_columns = ['Date', 'Month', 'WY']
sl_columns.extend(dss_pathb)
di_monthly = pd.read_excel(s_dvfile, index_col = 0)[sl_columns]

#Create dataframe to hold the final processed data. 
df_wytypes = pd.DataFrame(index = range(di_monthly.WY.min() -1, di_monthly.WY.max()+1))
#Loop through each of the wytypes 
for s_wytype in sc_wytypes:

    #Grab the data corresponding the month where water year type is defined
    i_month_defined = sc_wytypes[s_wytype][-1]

    #Grab the WYtype name.
    s_name_column = sc_wytypes[s_wytype][0]

    #Pull the column for this water year type's data from the Calsim DSS reader output file, for the month it is defined.
    df_defined = pd.DataFrame(di_monthly.loc[di_monthly.Month == i_month_defined].set_index('WY')[s_wytype])
    df_defined['WY'] = df_defined.index
    #For 1921, use the first date's wyt
    if di_monthly.Date.min() == datetime(1921, 10,31):
        df_defined.loc[1921, s_wytype] =di_monthly.loc[di_monthly.Date == di_monthly.Date.min()][s_wytype].unique()[0]

    #Clean up the data, removing duplicates (WY type is the same across different alternatives)
    df_defined.drop_duplicates(inplace = True)
    df_defined.sort_index(inplace = True)
    df_defined.drop(columns =[ 'WY'], inplace = True)

    #Store this data in the final processed data dataframe.
    df_wytypes[s_name_column] = pd.DataFrame(df_defined.copy(deep = True))

#Save wytype data to excel file.
df_wytypes.to_excel(s_output)
