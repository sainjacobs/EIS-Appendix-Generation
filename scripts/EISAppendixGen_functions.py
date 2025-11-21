import pandas as pd
import numpy as np
import docx
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_ROW_HEIGHT_RULE
import matplotlib.pyplot as plt
import calendar
import os
from time import strptime
from storage_to_elevation import storage_to_elevation
from ec_to_cl import ec_to_cl
from math import floor

def get_locations(location_crosswalk_path, fields):
    """
    Gets location names from field codes passed

    Parameters
    ----------
    location_crosswalk_path: string
        Path and file name for xlsx file containing location names and field codes
    fields: list of strings
        Names of the fields to be processed

    Returns
    ----------
    locations : list of str
        List of the location titles from the crosswalk file given in location_crosswalk_path

    """
    #Read in crosswalk as a df
    crosswalk = pd.read_excel(location_crosswalk_path)

    #Look up each field code's corresponding location title and add to a list
    locations = []
    for i, field in enumerate(fields):
        if type(field) ==str:
            locations.append(crosswalk.loc[crosswalk["DSSPartB"] == field, "Location (Title)"].values[0])
        elif type(field)==tuple:
            if i>0 and field == fields[i-1]:
                use_index = 1
            else:
                use_index = 0
            #If filter_search is provided, then filter the parameter column of the dataframe. This is used to get the elevation figure/table titles, since they're under the same dsspartb as the storage titles.
            locations.append(crosswalk.loc[(crosswalk['DSSPartB'] == field[0]) &(crosswalk.Parameter == field[1]), "Location (Title)"].values[use_index])

    return locations

def get_locations_params(location_crosswalk_path, fields):
    """
    Gets location names from field codes passed

    Parameters
    ----------
    location_crosswalk_path: string
        Path and file name for xlsx file containing location names and field codes
    fields: list of strings
        Names of the fields to be processed

    Returns
    ----------
    locations: list of str
        list of the parameter names corresponding to the locations, based on what's in the crosswalk file.

    """
    #Read in crosswalk as a df
    crosswalk = pd.read_excel(location_crosswalk_path)

    #Look up each field code's corresponding location title and add to a list
    locations = []
    for i, field in enumerate(fields):
        if type(field) ==str:
            locations.append(crosswalk.loc[crosswalk["DSSPartB"] == field, "Parameter"].values[0])
        elif type(field)==tuple:
            if i>0 and field == fields[i-1]:
                use_index = 1
            else:
                use_index = 0
            #If filter_search is provided, then filter the parameter column of the dataframe. This is used to get the elevation figure/table titles, since they're under the same dsspartb as the storage titles.
            locations.append(crosswalk.loc[(crosswalk['DSSPartB'] == field[0]) &(crosswalk.Parameter == field[1]), "Parameter"].values[use_index])

    return locations

def get_location_wytypes(location_crosswalk_path, fields):
    """
    Gets location names from field codes passed

    Parameters
    ----------
    location_crosswalk_path: string
        Path and file name for xlsx file containing location names and field codes
    fields: list of strings
        Names of the fields to be processed

    Returns
    ---------
    wytype_list:  list of str
        list of water year types corresponding to each location, based on what's in the crosswalk file.

    """
    #Read in crosswalk as a df
    crosswalk = pd.read_excel(location_crosswalk_path)

    #Look up each field code's corresponding location title and add to a list
    wytype_list = []
    for i, field in enumerate(fields):
        if type(field) ==str:
            wytype_list.append(crosswalk.loc[crosswalk["DSSPartB"] == field, "Water Year Type Index"].values[0])
        elif type(field)==tuple:
            if i>0 and field == fields[i-1]:
                use_index = 1
            else:
                use_index = 0
            #If filter_search is provided, then filter the parameter column of the dataframe. This is used to get the elevation figure/table titles, since they're under the same dsspartb as the storage titles.
            wytype_list.append(crosswalk.loc[(crosswalk['DSSPartB'] == field[0]) &(crosswalk.Parameter == field[1]), "Water Year Type Index"].values[use_index])

    return wytype_list

def calculate_supply_fields(s_inputs, s_formulas, s_wy_flags_path):
    """
    Reads in data and calculated the specific categories for the water supply appendix
    Parameters
    ----------
    s_inputs: str
        Path for DSS inputs
    s_formulas: str
        Path for the excel sheet with all the formulas
    s_wy_flags_path: path
        Path to the WYTs

    Returns
    -------
    df_final: dataframe
        Data for the tables
    df_exceedances: dataframe
        Dataframe with the exceedances
    """

    # DSS data
    df_inputs = pd.read_excel(s_inputs)

    # replace 'Baseline' with 'NAA'
    df_inputs.replace({'Baseline': 'NAA'}, inplace=True)

    # Formulas for calculated fields
    df_formulas = pd.read_excel(s_formulas, sheet_name='annual')

    # Will hold the outputs
    df_output = pd.DataFrame(index=pd.MultiIndex.from_product([df_inputs['Scenario'].unique(), df_inputs['Year'].unique()]))

    # Go through each sub field and calculate it
    for row_index, row in df_formulas.iterrows():
        s_formula = row['Formula']
        sl_add_fields = [field.strip() for field in s_formula.split(' + ')]
        s_stat = 'sum' if row['Statistic'] == 'Sum' else 'mean'
        if row['Resolution'] == 'Annual':
            if row['Annual_ Month_ Range'] == 'MarFeb':
                df_output.loc[:, row['Title']] = df_inputs.groupby(['Scenario', 'DY'])[sl_add_fields].agg(s_stat).agg(s_stat, axis=1)
            elif row['Annual_ Month_ Range'] == 'JanDec':
                df_output.loc[:, row['Title']] = df_inputs.groupby(['Scenario', 'Year'])[sl_add_fields].agg(s_stat).agg(s_stat, axis=1)
            elif row['Annual_ Month_ Range'] == 'OctSep':
                df_output.loc[:, row['Title']] = df_inputs.groupby(['Scenario', 'WY'])[sl_add_fields].agg(s_stat).agg(s_stat, axis=1)
        elif row['Resolution'] == 'SingleMonth':
            i_month = list(calendar.month_abbr).index(row['Annual_ Month_ Range'])
            df_output.loc[:, row['Title']] = df_inputs[df_inputs['Month'] == i_month].groupby(['Scenario', 'Year'])[sl_add_fields].agg(s_stat).agg(s_stat, axis=1)

    # drop 1921 partial year
    df_output.drop(index=[1921], level=1, inplace=True)

    # formulas for the final categories
    df_final_formulas = pd.read_excel(s_formulas, sheet_name='final', index_col=[0, 1])

    # data to hold the table data, skip 2021 since none have that full year
    df_final = pd.DataFrame(index=df_output.drop(index=[2021], level=1).index, columns=df_final_formulas.index)

    # go through each final formula and calculate it
    for row_index, row in df_final_formulas.iterrows():

        # Pull out description and units
        s_description = row['Description']
        s_units = row['Units']

        # remove description and units
        row = row.iloc[2:]

        # Fields to add up
        ls_fields = row[~row.isna()].values
        if len(ls_fields) == 0:
            # add in description and units
            df_final.loc['Description', row_index] = s_description
            df_final.loc['Units', row_index] = s_units
            continue

        # add them up and insert into final data frame
        df_final[row_index] = df_output[ls_fields].sum(axis=1)

        # add in description and units
        df_final.loc['Description', row_index] = s_description
        df_final.loc['Units', row_index] = s_units

    # save the descriptions and fields
    df_temp = df_final.loc[['Description', 'Units'], :]

    # calculate the fields that are not in the formulas
    df_final[('Total For All Regions', 'Total Supplies')] = df_final[['Sacramento River Hydrologic Region', 'San Joaquin River Hydrologic Region (not including Friant-Kern and Madera Canal water users)',
                                                                      'San Francisco Bay Hydrologic Region', 'Central Coast Hydrologic Region', 'Tulare Lake Hydrologic Region (not including Friant-Kern Canal water users)',
                                                                      'South Lahontan Hydrologic Region', 'South Coast Hydrologic Region']].sum(axis=1)

    df_final[('North of Delta', 'SWP Ag')] = df_final[('North of Delta', 'SWP Ag')].fillna(0)
    df_final[('Total CVP North of Delta', 'Total CVP Ag and M&I')] = df_final[[('North of Delta', 'CVP Ag'), ('North of Delta', 'CVP M&I')]].sum(axis=1)
    df_final[('Total SWP North of Delta', 'Total SWP Ag and M&I')] = df_final[[('North of Delta', 'SWP Ag'), ('North of Delta', 'SWP M&I')]].sum(axis=1)
    df_final[('Total North of Delta', 'Total Ag and M&I Deliveries')] = df_final[[('Total CVP North of Delta', 'Total CVP Ag and M&I'), ('Total SWP North of Delta', 'Total SWP Ag and M&I')]].sum(axis=1)
    df_final[('Total CVP South of Delta', 'Total CVP Ag and M&I')] = df_final[[('South of Delta', 'CVP Ag'), ('South of Delta', 'CVP M&I')]].sum(axis=1)
    df_final[('Total SWP South of Delta', 'Total SWP Ag and M&I')] = df_final[[('South of Delta', 'SWP Ag'), ('South of Delta', 'SWP M&I')]].sum(axis=1)
    df_final[('Total South of Delta', 'Total Ag and M&I Deliveries')] = df_final[[('Total CVP South of Delta', 'Total CVP Ag and M&I'), ('Total SWP South of Delta', 'Total SWP Ag and M&I')]].sum(axis=1)

    # replace the descriptions and units
    df_final.loc[['Description', 'Units'], :] = df_temp

    # the WYTs for each year
    wy_flags_all = pd.read_excel(s_wy_flags_path, index_col=0)

    # add in long term average and dry and critical average
    for scenario in df_final.index.get_level_values(0).unique():
        if scenario in ['Description', 'Units']:
            continue

        # Long term average
        df_final.loc[(scenario, 'Long Term'), :] = df_final.loc[scenario, :].mean()
        li_dry_crit_years = wy_flags_all[wy_flags_all['40-30-30'].isin([4, 5])].index
        if 2021 in li_dry_crit_years:
            li_dry_crit_years = li_dry_crit_years.drop(2021)

            # dry and crit years
        df_final.loc[(scenario, 'Dry and Critical'), :] = df_final.loc[scenario, :].loc[li_dry_crit_years, :].mean()

    # read in exceedance plot formulas
    df_exceedance_formulas = pd.read_excel(s_formulas, sheet_name='exceedance', index_col=0)

    df_exceedances = pd.DataFrame(index=pd.MultiIndex.from_product([df_inputs['Scenario'].unique(), df_inputs['Year'].unique()]), columns=df_exceedance_formulas.index)

    # calculate each exceedance field
    for row_index, row in df_exceedance_formulas.iterrows():

        # fields to add up
        ls_fields = row[~row.isna()].values
        ls_add_fields = [field for field in ls_fields if field[0] != '-']
        ls_subtract_fields = [field[1:] for field in ls_fields if field[0] == '-']

        # If they are all water year, inclue 2021
        if np.all(df_formulas[df_formulas['Title'].isin(ls_fields)]['Annual_ Month_ Range'] == 'OctSep'):
            # add them up and insert into final data frame
            df_exceedances[row_index] = df_output[ls_add_fields].sum(axis=1) - df_output[ls_subtract_fields].sum(axis=1)
        else:
            # add them up and insert into final data frame
            df_exceedances[row_index] = df_output.drop(index=[2021], level=1)[ls_add_fields].sum(axis=1) - df_output.drop(index=[2021], level=1)[ls_subtract_fields].sum(axis=1)

    return df_final, df_exceedances


def parse_dssReader_output(dss_path, runs, field, report_type, convert_to_elevation= False, convert_to_cl=False,  orig_unit = 'TAF', storage_elevation_fn = ''):
    """
    Reads DSS output from reader for desired runs and field

    Parameters
    ----------
    dss_path: string
        Path and file name for xlsx file containing DSSReader Output
    runs: list of strings
        Names of the runs to be processed
    report_type: string
        Type of report being generated. Used to check whether or not it's a temperature report
    field: string
        Current field being processed
    convert_to_elevation: bool
        True if you are converting storage to elevation. Need to also set the orig_unit field to the original storage
        unit
    orig_unit: str
        Original storage unit (Currently only have "TAF" implemented). Used for storage to elevation conversion.

    storage_elevation_fn: str
        Optional. Default is "". Filename of Excel file containing storage-elev relationships for CalSim. (Storage-elev
        tables taken from lookup/res_info.table in CalSim wresl code.

    :returns
    t_dfs: list of pandas dataframes
        List of dataframe containing data for this location. Each dataframe corresponds to a run. Dataframe's has columns
        for WY, and each month.

        For temperature, daily values will be averaged to monthly.

    """
    #Read DSS Output from specified path for specified field
    dss_output = pd.read_excel(dss_path)

    dss_output = dss_output[["Month", "Scenario", "WY", field]]

    if report_type in ["temperature"]:
        #If temperature or DSM2 data is being read, convert daily data to monthly by averaging
        #scenario = dss_output.loc[0, "Scenario"]
        #dss_output.drop(columns = ["Scenario"], inplace = True)

        monthly_data = dss_output.groupby(["Scenario", "WY", "Month"]).mean()
        monthly_data.reset_index(inplace=True)
        dss_output = monthly_data

        #Drop rows with flag value for missing data
        rows_to_drop = (dss_output[dss_output.columns[3:]] < -100).any(axis=1)
        dss_output = dss_output[~rows_to_drop]

        #dss_output["Scenario"] = scenario
    #If we want elevation, need to convert from storage
    if convert_to_elevation:
        #Convert storage to elevation
        df_elevations = storage_to_elevation(dss_output, field, storage_elevation_fn, orig_unit = orig_unit)
        #Replace the dss_output dataframe and continue formatting the tables.
        dss_output = df_elevations
    if convert_to_cl:
        #Convert EC (microsiemens/cm) to mg/L Cl using the regression relationship defined as eqn 2 in
        #https://www.waterboards.ca.gov/waterrights/water_issues/programs/bay_delta/california_waterfix/exhibits/docs/ccc_cccwa/CCC-SC_25.pdf
        df_cl = ec_to_cl(dss_output, field, orig_unit = orig_unit)
        #Replace the dss_output dataframe and continue formatting the tables.
        dss_output = df_cl

    # Create df for each alternative/run and reformat
    run_dfs = []
    for run in runs:
        if run == "NAA":
            run_df = dss_output.loc[dss_output["Scenario"] == "Baseline"]
        else:
            run_df = dss_output.loc[dss_output["Scenario"] == run]

        run_df["month_name"] = " "

        #Add abbrievated month name to df for tables and plotting later
        for index, row in run_df.iterrows():
            run_df.loc[index, "month_name"] = calendar.month_abbr[int(row["Month"])]
        #Drop unneeded columns
        run_df.drop(columns=["Month", "Scenario"], inplace=True)
        run_dfs.append(run_df)

    #Transpose dfs to be in correct format for tables
    t_dfs = []
    for run_df in run_dfs:
        transposed_df = pd.DataFrame(
            columns=["WY", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"])
        #One row for each WY consisting of a column for each monthly EC value
        for year in np.unique(run_df["WY"]):
            year_t = run_df.loc[run_df["WY"] == year]
            year_t.set_index("month_name", inplace=True)
            year_t = year_t.transpose()
            year_t.insert(0, "WY", year)
            year_t.reset_index(drop=True, inplace=True)
            #Add each year as new row to df
            transposed_df = pd.concat([transposed_df, year_t.iloc[1:2]], axis=0, ignore_index=True)

        t_dfs.append(transposed_df)

    return t_dfs

def create_exceedance_tables(t_dfs, wy_flags_path, use_wytype, report_type, use_calendar_yr = False):
    """
    Creates exceedance tables formatted for appendix report from transposed DSSReader Output

    Parameters
    ----------
    t_dfs: list of dataframes
        Dataframe outputs from DSSReader that have been transposed to be formatted for table
    wy_flags_path: str
        excel file with the water year types
    use_wytype: str
        water year type to use.
        "40-30-30" to use WYT_SAC_
        "60-20-20" to use WYT_SJR_
        "TRIN" to use WYT_TRIN_
    report_type: str
        type of report (Calsim, temperature, etc)
    use_calendar_yr: bool
        Optional. Default is false. True indicates months should be sorted into calendar year instead of water year
        when calculating the water year type averages. False indicated that months will be sorted into water years when
        determining water year type.


    Returns
    ----
    exc_tables: list of pandas DataFrames containing 1,10,20,..., 90,99% exceedance probability data and the WYType summary statistics
    exc_probs: pandas dataseries of exceedance probabilities that are represented in the exc_tables
    fig_tables: list of pandas DataFrames of full list of exceedance probabilities and corresponding values. Used for plotting.
    il_num_years: list of pandas DataFrames that record how many years of data are available for each month

    """
    exc_tables = []
    fig_tables = []
    il_num_years = []
    wy_list = t_dfs[0].WY.tolist()
    for table in t_dfs:
        table = table.drop(columns = ["WY"]).copy()
        table = table.apply(lambda x: x.sort_values(ascending=False).values)
        #Remove first and last rows
        #table.iloc[::-1, ::1]
        #Rank ECs from 1 to 100 - No longer using this, since this produces inaccurate exceedance probabilities if simulation periods don't start Oct and end in Sept.
        #table.insert(0, "Rank", range(1,table.shape[0] + 1))
        ##Calculate exceedance probability and add column to table
        #table.insert(1, "Exc Prob", (table["Rank"]) / (table.shape[0] + 1) * 100) # m/(N+1)

        # Create dataframe for 10%, 20%, 30%, etc. exceedance values by linearly interpolating between the table values for each month
        table_interp = pd.DataFrame(index = range(10,100, 10))
        table_interp.index.name = 'Exc Prob'
        #Create dataframe for 1,2,3..., 99% exceedance values by linearly interpolating between the table values for each month. Used for plotting figures.
        table_all = pd.DataFrame(index = range(1,101,1))
        table_all.index.name = 'Exc Prob'

        for m_name in table.columns[-12:]:
            #Subset the table data column corresponding to this month.
            df_month = table[[m_name]]
            #Remove any null entries
            df_month.dropna(inplace = True)
            #Calculate the rank for the remaining (non-null) entries
            df_month['Rank'] = range(1,len(df_month)+1)
            df_month['Exc Prob'] = df_month["Rank"]/ (df_month.shape[0] + 1) * 100 # m/(N+1)

            #Interpolate to get the 10,20,...,90% exc prob values and place in dataframe.
            exceedance_values = np.interp(table_interp.index.values, df_month['Exc Prob'].values, df_month[m_name].values)
            table_interp[m_name]= exceedance_values

            # Interpolate to get the 1,2,3,4..., 99% exc prob values and place in dataframe.
            exceedance_values_all = np.interp(table_all.index.values, df_month['Exc Prob'].values,
                                              df_month[m_name].values)
            df_exceedance_values_all = pd.DataFrame(exceedance_values_all, index = table_all.index.values, columns = [m_name])
            #Only include values for exp probabilities between min(m/(N+1)) and max(m/(N+1))
            df_exceedance_values_all = df_exceedance_values_all.loc[(df_exceedance_values_all.index>= df_month['Exc Prob'].min()) &(df_exceedance_values_all.index<=df_month['Exc Prob'].max())]
            table_all[m_name]=  df_exceedance_values_all
        #Reset index for the table_all
        table_all.reset_index(inplace = True, drop = False)

        #Add the lowest and highest exceedance probability rows
        table_interp.loc[table_all.iloc[0]['Exc Prob']] = table_all.iloc[0][table_all.columns[-12:]].values
        table_interp.loc[table_all.iloc[-1]['Exc Prob']] = table_all.iloc[-1][table_all.columns[-12:]].values

        #Sort by exceedance probability.
        table_interp.sort_index(inplace = True)
        table_interp.reset_index(drop = False, inplace = True)

        #Store the exceedance table
        exc_tables.append(table_interp)
        #Store the exc probability table for plotting the figures
        fig_tables.append(table_all.copy(deep = True))

        #Calculate the sample size used to calculate statistics in each month
        df_num_years = table.count(axis = 0)
        il_num_years.append(df_num_years)

    #Calculate full simulation period average for each run and format to be added to exceedance table as one row
    stats_dfs = []
    for run in t_dfs:
        period_ave = run.drop(columns = ['WY']).mean(axis=0)
        stats_df = pd.DataFrame(period_ave)
        stats_df = stats_df.transpose()

        stats_df["Exc Prob"] = ["Full Simulation Period Average"]

        stats_dfs.append(stats_df)

    #Read in water year typing flags
    wy_flags_all = pd.read_excel(wy_flags_path, index_col = 0)
    #Subset to only the wytype that is specified by use_wytype
    wy_flags = wy_flags_all[[use_wytype]]

    if use_wytype == 'TRIN': #Names for Trinity WYType
        year_types = ["Extremely Wet", "Wet", "Normal", "Dry", "Critically Dry"]
    else: #Names for the Sacramento and SJR WYType
        year_types = ["Wet", "Above Normal", "Below Normal", "Dry", "Critical"]
    # make a copy of exc probabilities to use with figures after deleting from tables df
    exc_probs = exc_tables[0]["Exc Prob"]

    #Create empty dataframe to store percentages for each of the wytypes
    wytype_percents = pd.DataFrame(index = range(1, 6), columns = ['percentage'])

    # calculate wet, above normal, dry, etc water years (sum for year type/ count of year type)
    for table_index in range(len(t_dfs)):

        t_dfs[table_index].set_index('WY', inplace = True)
        t_dfs[table_index]["flag"] = wy_flags[use_wytype]
        if use_calendar_yr:
            df_monthly_ts = pd.melt(t_dfs[table_index].reset_index(drop=False), id_vars='WY',
                    value_vars=t_dfs[table_index].columns[:-1])
            df_monthly_ts['i_month']= df_monthly_ts.apply(lambda l: datetime.strptime(l.variable+ "-01-1900","%b-%d-%Y" ).month,axis = 1)
            df_monthly_ts['dates'] = df_monthly_ts.apply(lambda l: datetime(l.WY -1 , l.i_month, 1) if l.i_month>=10 else datetime(l.WY, l.i_month, 1), axis = 1)
            df_monthly_ts['calendar_yr'] = df_monthly_ts.apply(lambda l: l.WY -1  if l.i_month>=10 else l.WY, axis = 1)

            #Calendar year dataframe (rows are calendar years, columns are months, values are monthly values
            t_dfs_calendar_yr = df_monthly_ts.pivot(columns = 'variable', index = 'calendar_yr', values = 'value')
            #Reorder the columns so months are in order.
            t_dfs_calendar_yr = t_dfs_calendar_yr [['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']]

            #Grab the wy that is associated with each calendar year and add as a column. This means that calendar yr 1980 will be assigned the flag for wytype associated with wy1980.
            t_dfs_calendar_yr['flag'] = wy_flags[use_wytype]

        exc_probs_i = exc_tables[table_index]["Exc Prob"]
        month_vals = {}
        # Also add full sim period average as a row in exceedance table
        exc_tables[table_index] = pd.concat([exc_tables[table_index], stats_dfs[table_index].iloc[0:1]], ignore_index=True)

        #Iterate through each type of year (wet, above normal, etc) to compute sums
        for year_type in range(len(year_types)):

            # Calculate the percentage of total water years that have this wytype
            if use_calendar_yr:
                d_percent_wytype = round(len(t_dfs_calendar_yr.loc[t_dfs_calendar_yr['flag'] == year_type + 1]) / len(t_dfs_calendar_yr) * 100, 1)
            else:
                d_percent_wytype = round(len(t_dfs[table_index].loc[t_dfs[table_index]['flag'] == year_type + 1]) / len(t_dfs[table_index]) * 100, 1)
            wytype_percents.at[year_type+1, 'percentage'] = d_percent_wytype
            for month in t_dfs[table_index].columns[:-1]:
                #Flags are 1 - 5 to specify which type of year
                #Calculate mean of months classified as current year type based on flag

                ## Using water years
                if not use_calendar_yr:
                    month_vals[month] = [t_dfs[table_index].loc[t_dfs[table_index]['flag'] == (year_type + 1), month].mean()]
                else:
                    # Using calendar years
                    month_vals[month] =[t_dfs_calendar_yr.loc[t_dfs_calendar_yr['flag'] == (year_type + 1), month].mean()]




            month_vals["Exc Prob"] = year_types[year_type]

            exc_tables[table_index] = pd.concat([exc_tables[table_index], pd.DataFrame.from_dict(month_vals)], ignore_index=True)


        #Create list of desired row labels
        row_labels = [f"{round(value)}% Exceedance" for value in exc_probs_i.values]
        row_labels.append('Full Simulation Period Average')
        wy_type_labels = [f"{year_types[i]} Years ({wytype_percents.loc[i+1].item():.0f}%)" if wytype_percents.loc[i+1].item() == int(wytype_percents.loc[i+1].item())  else
                          f"{year_types[i]} Years ({wytype_percents.loc[i + 1].item():.1f}%)" for i in range(len(year_types))]
        row_labels.extend(wy_type_labels)

        #Remove extra columns
        exc_tables[table_index].drop(columns=["Exc Prob"], inplace=True)

        #Round table values
        if report_type == 'temperature':
            exc_tables[table_index] = exc_tables[table_index].astype(float)#.round(1)
        else:
            exc_tables[table_index] = exc_tables[table_index].astype(float)#.round(0)

        # Add row labels for report tables in first column
        exc_tables[table_index].insert(0, "Statistic", row_labels)

        # Move new header names to first row
        exc_tables[table_index].index = exc_tables[table_index].index + 1  # shifting index
        exc_tables[table_index] = exc_tables[table_index].sort_index()

    return exc_tables, exc_probs, fig_tables, il_num_years


def make_rows_bold(*rows):
    """
    Makes text in specified table rows bold.

    Parameters
    ----------
    rows: row attributes from docx table object
        1 or more rows that will be converted to bold text (ex table.rows[0])

    """

    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def set_cell_border(cell: _Cell, **kwargs):
    """
    Set cell's border.
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )

     Parameters
    ----------
    cell: cell attribute from docx table object
        1 cell that will be converted to bold text (ex table.rows[0].cells[0])

    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def set_cell_color(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    tcPr.append(shading)


def change_table_font_size(document, font_size):
    """
    Changes the font size of all text in all tables within a document.

    Parameters
    ----------
    document: docx document object
        Document that will have font size adjusted
    font_size: int
        New font size

    """

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = docx.shared.Pt(font_size)


def add_commas_to_table(doc):
    """
    Adds commas to numbers in all tables of a docx document.

    Parameters
    ----------
    doc: docx document object
        Document that will have commas added to numeric values in all tables

    """

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        try:
                            # Check if the text is a number
                            number = float(run.text)
                            # Format the number with commas
                            formatted_number = f"{number:,}"
                            formatted_number = formatted_number.rsplit(".", 1)[0]
                            run.text = formatted_number
                        except ValueError:
                            # If the text is not a number, do nothing
                            pass

def format_decimals(doc):
    """
    Adds commas to numbers in all tables of a docx document.

    Parameters
    ----------
    doc: docx document object
        Document that will have commas added to numeric values in all tables

    """

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        try:
                            # Check if the text is a number
                            number = float(run.text)
                            # Format the number with commas
                            #formatted_number = f"{number:,}"
                            formatted_number = "{:.1f}".format(number)
                            #formatted_number = formatted_number.rsplit(".", 1)[0]
                            run.text = formatted_number
                        except ValueError:
                            # If the text is not a number, do nothing
                            pass
def change_orientation(doc, new_orientation):
    """
    Changes section orientation from portrait to landscape or vice versa

    Parameters
    ----------
    doc: docx document object
        Document that will have commas added to numeric values in all tables
    new_orientation: string
        Either "landscape" or "portrait" to indicate the desire page orientation for the new section

    """

    current_section = doc.sections[-1]
    new_width, new_height = current_section.page_height, current_section.page_width
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    if new_orientation == "landscape":
        new_section.orientation = WD_ORIENT.LANDSCAPE
    else:
        new_section.orientation = WD_ORIENT.PORTRAIT

    new_section.page_width = new_width
    new_section.page_height = new_height

    return new_section

def format_table(doc_table, table_df, doc, report_type):
    """
    Creates tables formatted for appendix report from exceedance tables

    Parameters
    ----------
    t: docx table object
        Exceedance table to be formatted for report
    table_df: dataframe
        Dataframe containing data to go into report table
    doc: docx object
        Docx object containing table to be formatted
    """
    # Change font size to fit on page better
   # change_table_font_size(doc, 8)

    # add the header rows.
    for column_index in range(table_df.shape[1]):
        doc_table.cell(0, column_index).text = table_df.columns[column_index]

    # add the rest of the data frame
    for row_index in range(table_df.shape[0]):
        for column_index in range(table_df.shape[1]):
            if column_index == 0: #Add the exceedance percentage row labels to the table
                doc_table.cell(row_index + 1, column_index).text = str(table_df.values[row_index, column_index])
            else:
                #For all other columns, add the CalSim, temperature, or salinity values in the table. Round values.
                if report_type == 'temperature':
                    #For temperature tables, round values to nearest 10th place
                    doc_table.cell(row_index + 1, column_index).text = str(round(table_df.values[row_index, column_index],1))
                else:
                    #For CalSim or DSM2 tables, round values to the nearest integer
                    doc_table.cell(row_index + 1, column_index).text = str(round(table_df.values[row_index, column_index]))
    # Set table top and bottom borders
    borders = OxmlElement('w:tblBorders')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), '4')
    borders.append(bottom_border)
    top_border = OxmlElement('w:top')
    top_border.set(qn('w:val'), 'single')
    top_border.set(qn('w:sz'), '4')
    borders.append(top_border)

    doc_table._tbl.tblPr.append(borders)

    # Make headers bold
    make_rows_bold(doc_table.rows[0])

    # Make first column bold
    bolding_columns = [0]
    for row in list(range(table_df.shape[0] + 1)):
        for column in bolding_columns:
            doc_table.rows[row].cells[column].paragraphs[0].runs[0].font.bold = True

    # Add superscript to Full Simulation Period Average cell
    script_cell = doc_table.cell(10, 0).paragraphs[0]
    run = script_cell.add_run("a")
    run.font.superscript = True

    # Add borders to middle row and under header
    for cell in doc_table.rows[0].cells:
        set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    for cell in doc_table.rows[10].cells:
        set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    for cell in doc_table.rows[10].cells:
        set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    # Widen margins of table
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # Widen cell size in first column
    for cell in doc_table.columns[0].cells:
        cell.width = Inches(3.4)

    if report_type == "temperature":
        #format numbers to one decimal pt
        format_decimals(doc)
    else:
        # Add commas to values in table
        add_commas_to_table(doc)
        #Commas won't be needed for temperature values

    # Align values in center of cells
    for row in doc_table.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Decrease row spacing for table
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(0.45)  # 2 cm

    # Change font size to fit on page better
    change_table_font_size(doc, 8)


def format_table_supply(doc_table, df_table, doc, comparison, il_page_breaks):
    """
    Creates table for water supply data

    Parameters
    ----------
    doc_table: docx table object
        Table to be formatted for report
    df_table: dataframe
        Dataframe containing data to go into report table
    doc: docx object
        Docx object containing table to be formatted
    comparison: list
        List of the comparison names
    il_page_breaks: list
        Rows that need a page break header

    Returns
    -------
    none
    """

    doc_table.autofit = False
    # set consistent borders over whole table
    for row in doc_table.rows:
        for cell in row.cells:
            set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                            bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                            start={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                            end={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Decrease row spacing for table
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Inches(0.25)

    # Create the Header Row
    doc_table.cell(0, 0).merge(doc_table.cell(0, 3))
    doc_table.cell(0, 0).text = 'Water Supply Reliability'
    doc_table.cell(0, 4).text = comparison[1]
    doc_table.cell(0, 5).text = comparison[0]
    doc_table.cell(0, 6).text = comparison[1] + ' minus ' + comparison[0]
    doc_table.rows[0].height = Inches(0.7)
    doc_table.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

    # Make headers bold
    make_rows_bold(doc_table.rows[0])

    curr_row = 1

    # Loops through each section and subsetion and add into table
    for section_name in df_table.columns.get_level_values(0).unique():

        # if we hit a page break, recreate the header
        if curr_row in il_page_breaks:

            # add row to bottom and format it
            row = doc_table.add_row()
            for cell in row.cells:
                set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                start={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                end={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            # Decrease row spacing for table
            row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            row.height = Inches(0.25)

            # Create the Header Row
            doc_table.cell(curr_row, 0).merge(doc_table.cell(curr_row, 3))
            doc_table.cell(curr_row, 0).text = 'Water Supply Reliability'
            doc_table.cell(curr_row, 4).text = comparison[1]
            doc_table.cell(curr_row, 5).text = comparison[0]
            doc_table.cell(curr_row, 6).text = comparison[1] + ' minus ' + comparison[0]
            doc_table.rows[curr_row].height = Inches(0.7)
            doc_table.rows[curr_row].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            # Make headers bold
            make_rows_bold(doc_table.rows[curr_row])
            curr_row += 1

        # Create the section
        doc_table.cell(curr_row, 0).merge(doc_table.cell(curr_row, 6))
        doc_table.cell(curr_row, 0).text = section_name.upper()

        # Make headers bold
        make_rows_bold(doc_table.rows[curr_row])

        # Make the background grey
        set_cell_color(doc_table.cell(curr_row, 0), "#E8E8E8")

        curr_row += 1

        # go through the subsections
        for sub_section in df_table[section_name].columns:

            # if we hit a page break, recreate the header
            if curr_row in il_page_breaks:
                row = doc_table.add_row()
                for cell in row.cells:
                    set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                    bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                    start={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"},
                                    end={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                # Decrease row spacing for table
                row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
                row.height = Inches(0.25)

                # Create the Header Row
                doc_table.cell(curr_row, 0).merge(doc_table.cell(curr_row, 3))
                doc_table.cell(curr_row, 0).text = 'Water Supply Reliability'
                doc_table.cell(curr_row, 4).text = comparison[1]
                doc_table.cell(curr_row, 5).text = comparison[0]
                doc_table.cell(curr_row, 6).text = comparison[1] + ' minus ' + comparison[0]
                doc_table.rows[curr_row].height = Inches(0.7)
                doc_table.rows[curr_row].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

                # Make headers bold
                make_rows_bold(doc_table.rows[curr_row])
                curr_row += 1

            # add the name of the subsection
            doc_table.cell(curr_row, 0).merge(doc_table.cell(curr_row + 1, 0))
            doc_table.cell(curr_row, 0).text = sub_section

            # add description
            doc_table.cell(curr_row, 1).merge(doc_table.cell(curr_row + 1, 1))
            doc_table.cell(curr_row, 1).text = df_table.loc['Description', (section_name, sub_section)]

            # Add units
            doc_table.cell(curr_row, 2).merge(doc_table.cell(curr_row+1, 2))
            doc_table.cell(curr_row, 2).text = df_table.loc['Units', (section_name, sub_section)]

            # Add long term numbers
            doc_table.cell(curr_row, 3).text = 'Long Term'
            doc_table.cell(curr_row, 4).text = str(round(df_table.loc[(comparison[1], 'Long Term'), (section_name, sub_section)]))
            doc_table.cell(curr_row, 5).text = str(round(df_table.loc[(comparison[0], 'Long Term'), (section_name, sub_section)]))
            doc_table.cell(curr_row, 6).text = str(round(df_table.loc[(comparison[1], 'Long Term'), (section_name, sub_section)]- df_table.loc[(comparison[0], 'Long Term'), (section_name, sub_section)]))

            curr_row += 1
            # add dry and crit numbers
            doc_table.cell(curr_row, 3).text = 'Dry and Critical'
            doc_table.cell(curr_row, 4).text = str(round(df_table.loc[(comparison[1], 'Dry and Critical'), (section_name, sub_section)]))
            doc_table.cell(curr_row, 5).text = str(round(df_table.loc[(comparison[0], 'Dry and Critical'), (section_name, sub_section)]))
            doc_table.cell(curr_row, 6).text = str(
                round(df_table.loc[(comparison[1], 'Dry and Critical'), (section_name, sub_section)] - df_table.loc[(comparison[0], 'Dry and Critical'), (section_name, sub_section)]))

            curr_row += 1

    # Formatting for table
    borders = OxmlElement('w:tblBorders')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), '4')
    borders.append(bottom_border)
    top_border = OxmlElement('w:top')
    top_border.set(qn('w:val'), 'single')
    top_border.set(qn('w:sz'), '4')
    borders.append(top_border)

    doc_table._tbl.tblPr.append(borders)

    # Add commas to values in table
    add_commas_to_table(doc)

    # Adjust the width of the columns
    for cell in doc_table.columns[0].cells:
        cell.width = Inches(1.45)
    for cell in doc_table.columns[1].cells:
        cell.width = Inches(3.2)
    for cell in doc_table.columns[2].cells:
        cell.width = Inches(.75)
    for cell in doc_table.columns[3].cells:
        cell.width = Inches(1.2)
    for cell in doc_table.columns[4].cells:
        cell.width = Inches(0.8)
    for cell in doc_table.columns[5].cells:
        cell.width = Inches(0.8)
    for cell in doc_table.columns[6].cells:
        cell.width = Inches(0.8)

    # Change font size to fit on page better
    change_table_font_size(doc, 10)


def create_month_plot(dfs, fig_value, month, month_directory, alts, line_styles, line_colors, report_type=''):
    """
    Generates and saves individual month plots

    Parameters
    ----------
    dfs: list of dataframes
        List of dataframes with monthly values as one of the columns
    fig_value: str
        y-axis label
    month: string
        Current month to be evaluated
    month_directory: string
        Directory to save month plots in
    alts: list of strings
        Names of the runs being compared in report
    line_styles: list of strings
        Styles for lines on plots
    line_colors: list of strings
        Colors for lines on plots
    report_type: str
        Type of report, only really matters if its water supply

    Returns
    --------
    None

    """
    # Check for/create directory to store monthly exceedance plots
    if not os.path.exists(month_directory):
        os.makedirs(month_directory)

    # define size and borders
    if report_type == 'water supply':
        fig, axs = plt.subplots(figsize=(9, 5.5), linewidth=1, edgecolor="black")
    else:
        fig, axs = plt.subplots(figsize=(10, 5), linewidth=3, edgecolor="black")

    for fig_index in range(len(dfs)):
        # Dataset for this alt
        df_alt_data = dfs[fig_index].copy(deep=True)

        # Subset to only the month of interest
        df_month = df_alt_data[[month]]

        # Now calculate exceedance values using this month's data
        df_month.sort_values(by=month, inplace=True, ascending=False)
        df_month.dropna(subset=[month], inplace=True)
        df_month['Rank'] = range(1, len(df_month) + 1)
        df_month['Exc Prob'] = df_month["Rank"] / (df_month.shape[0] + 1) * 100  # m/(N+1)

        # plot exceedance probability vs monthly EC
        percentages = range(0, 101, 10)
        percentage_labels = [f"{int(i)}%" for i in percentages]

        axs.plot(df_month['Exc Prob'].values, df_month[month].values, color=line_colors[fig_index],
                 linestyle=line_styles[fig_index], label=alts[fig_index])
        axs.set_xticks(percentages)
        axs.set_xticklabels(percentage_labels)


        axs.set_ylabel(fig_value)
        axs.set_xlabel("Exceedance Probability")

        # Save this parameter to orient the legend correctly
        axbox = axs.get_position()

        # Add gridlines
        plt.grid(color='gray', linestyle='--', linewidth=0.8)

        # Add a legend
        plt.legend(loc='center', ncol=4, bbox_to_anchor=[axbox.x0 + 0.5 * axbox.width, 1.08])

    if report_type != 'water supply':
        # Add month number at beginning so that figures can be easily inserted in CY order to document later
        month_number = str(strptime(month, '%b').tm_mon)
        # Add leading zeros to month numbers
        if len(month_number) < 2:
            month_number = str(0) + month_number

    # flip x-axis
    axs.invert_xaxis()

    if report_type == 'water supply':
        # Save figure to directory
        plt.savefig(month_directory + "/" + month + ".png")
    else:
        # Save figure to month directory
        plt.savefig(month_directory + "/" + month_number + "_" + month + "_monthly_exceedance" + ".png")

    plt.close()


def create_stat_plot(stat_fig_dfs, fig_value, stat, stat_directory, alts, line_styles, line_colors):
    """
    Generates and saves individual month plots

    Parameters
    ----------
    stat_fig_dfs: list of dataframes
        Dataframes with average values by year type
    fig_value: str
        plot ylabel
    stat: string
        Current type of year to be evaluated
    stat_directory: string
        Directory to save stat plots in
    alts: list of strings
        Names of the runs being compared in report
    line_styles: list of strings
        Styles for lines on plots
    line_colors: list of strings
        Colors for lines on plots
    Returns
    ----------
    None
    """
    if not os.path.exists(stat_directory):
        os.makedirs(stat_directory)

    fig, axs = plt.subplots(figsize=(10, 5), linewidth=3, edgecolor="black")
    for fig_index in range(len(stat_fig_dfs)):
        if stat == "Full Simulation Period":
            axs.plot(stat_fig_dfs[fig_index]["month"], stat_fig_dfs[fig_index]["Full Simulation Period Average"], color=line_colors[fig_index],
                     linestyle=line_styles[fig_index])
        else:
            axs.plot(stat_fig_dfs[fig_index]["month"], stat_fig_dfs[fig_index][stat], color=line_colors[fig_index],
                     linestyle=line_styles[fig_index])

        # Save this to position legend correctly
        axbox = axs.get_position()

        axs.set_ylabel(fig_value)

        # Add gridlines
        plt.grid(color='gray', linestyle='--', linewidth=0.8)
        # Add legend
        plt.legend(labels=alts, loc='center', ncol=4, bbox_to_anchor=[axbox.x0 + 0.5 * axbox.width, 1.08])

    # Save stat fig to directory
    plt.savefig(stat_directory + "/" + stat[:5] + "_exceedance" + ".png")
    plt.close()

def order_elevation_storage_fields(fields):
    """
    Generates list of tuples where each tuple is (input field, storage or elevation), based on the fields provided.
    This is used to preprocess the fields for the storage and elevation appendix, since some fields have both storage
    and elevation, while others just have storage (Ex: S_SLUIS_CVP).

    Note that this function will check for whether the field is in the list or not. If it isn't it'll raise an error
    telling the user they need to update the master list.

    Parameters
    ----------
    fields: list of str
        List of the fields being included in this appendix. This function is only intended to take reservoir fields as
        inputs.
    """
    # List of all the location and elevation fields in the desired order
    master_list = [("S_TRNTY", 'Storage'),
                   ("S_TRNTY", 'Elevation'),
                   ("S_SHSTA", 'Storage'),
                   ("S_SHSTA", 'Elevation'),
                   ("S_OROVL", 'Storage'),
                   ("S_OROVL", 'Elevation'),
                   ("S_FOLSM", 'Storage'),
                   ("S_FOLSM", 'Elevation'),
                   ("S_SLUIS", 'Storage'),
                   ("S_SLUIS", 'Elevation'),
                   ("S_SLUIS_CVP", 'Storage'),
                   ("S_SLUIS_SWP", 'Storage'),
                   ("S_MELON", "Storage"),
                   ("S_MELON", "Elevation"),
                   ("S_MLRTN", "Storage"),
                   ("S_MLRTN", "Elevation"),
                    ] #List of all the location and elevation fields in the desired order

    #Subset the list based on the fields.
    subset_list = [ordered_field for ordered_field in master_list if ordered_field[0] in fields]

    #Check to make sure that there's no new field names that aren't in the master list
    fields_in_master = [ordered_field[0] for ordered_field in master_list]
    for field in fields:
        if field not in fields_in_master:
            #If any field is not included in the master_list, then throw an error telling the user to add it to the master list. Otherwise, it will be excluded from the generated appendix.
            raise ValueError(f"Need to update master list with new field {field}.")

    return subset_list


