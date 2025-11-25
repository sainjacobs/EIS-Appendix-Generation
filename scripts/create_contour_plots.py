"""
Creates contour plots showing temperatures at several points donwstream on the Sacramento River. Each contour plot
shows 1yr of data.

"""

import pandas as pd
import numpy as np
import os
from datetime import date
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib
from matplotlib.colors import LinearSegmentedColormap
matplotlib.use('agg')

#Set up plot formatting to match Visual Identity Online Manual
matplotlib.rcParams['font.family'] = 'sans-serif'
matplotlib.rcParams['font.sans-serif'] = ['Segoe UI'] + matplotlib.rcParams['font.sans-serif']
matplotlib.rcParams['font.weight'] = 'normal'
matplotlib.rcParams['font.size'] = 14
matplotlib.rcParams['grid.linestyle'] = '--'
matplotlib.rcParams['axes.grid'] = True

#Reclamation color palette
sl_colors = ['#003e51', '#00809e', '#C69214','#e0d6b5' , '#F56600', '#144d29' ,'#674da1','#9A3324']

#Create colormaps from Reclamation color palette.
#Sequential colormap
o_cm = LinearSegmentedColormap.from_list('usbr', sl_colors[:4], N=8)
#Diverging colormap
o_cm_diverging = LinearSegmentedColormap.from_list('usbr',
                                                   [sl_colors [0], sl_colors[1],  sl_colors[-4], sl_colors[-1]],
                                                   N=4*2)


def generate_contour_plot(i_year, s_yrly_contour_dir, df_results_temperature, da_river_miles, o_colormap,
                          d_contour_level_start=40, d_contour_level_end = 75, i_levels = 8, alt_name = ''):
    """
    Generates contour plot of water temperatures using temperature outputs at multiple locations downstream.
    Uses 1 (calendar) year of data.

    Parameters
    ----------
    i_year: integer
        Calendar year
    s_yrly_contour_dir: str
        Directory where we save the contour plot
    df_results_temperature: float dataframe
        Dataframe of simulated temperatures (Rows are dates, columns are locations)
        Note: columns in df_results_temperature must be in order from upstream to downstream - Ex: Below Trinity,
        Above Lewiston, Below Lewiston, Douglas City, and North Fork Trinity.
    o_colormap: object
        Colormap object
    d_contour_level_start: float
        Minimum contour value
    d_contour_level_end: float
        Maximum contour value
    i_levels: int
        Number of contour levels
    alt_name: str
        Optional. Defaults to "". Alternative name associated with data used to generate contour plot. Is annotated on
         figure as part of the colorbar label.

    Returns
    -------
    None

    """

    # Set start and end dates for the year
    o_year_start = date(i_year, 1, 1)
    o_year_end = date(i_year, 12, 31)

    # Subset this year's data and reshape temperatures so rows are locations from upstream to downstream and columns are days of the year.
    da_temps = df_results_temperature.loc[o_year_start: o_year_end].T.values

    # Approximate river miles for the locations. Not exact/only meant for visualization.
    # Note that columns in df_results_temperature must be in order from upstream to downstream: Below Trinity, Above Lewiston, Below Lewiston, Douglas City, and North Fork Trinity.
    # da_river_miles = np.array([105, 101, 99, 92.6, 72.6]) #This is now an input.

    # Get dates in the year
    oa_year_dates = df_results_temperature.loc[o_year_start:o_year_end].index#pd.date_range(o_year_start, o_year_end, freq='1D')
    print( f"{alt_name}####\n")
    print(df_results_temperature.loc[o_year_start:o_year_end])
    # Set the contour levels to use in the plot
    dl_contour_levels = np.linspace(d_contour_level_start, d_contour_level_end, i_levels)

    # Create water temperature contour plot. (y-axis is location, x-axis is date of year)
    o_fig, o_ax = plt.subplots(figsize=(16, 5))
    o_contour_plot = o_ax.contourf(oa_year_dates, da_river_miles, da_temps, cmap=o_colormap,
                                   levels=dl_contour_levels, extend='both')

    #Add a contour for 53.5degrees F (this is the target temperature for the mixed compliance location logic).
    # o_compliance_line = o_ax.contour(oa_year_dates, da_river_miles, da_temps, levels = [53.5] , color = 'r')

    # Draw horizontal dashed line for each of the locations
    for d_river_mile in da_river_miles:
        o_ax.axhline(y=d_river_mile, linestyle='--', color='white')

    # Set x-axis tick format and label
    o_ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    #Label months on the month-end date (b/c this is how the datetime index for monthly outputs is for dss outputs)
    o_ax.set_xticks(pd.date_range(o_year_start, o_year_end, freq = "1ME"))
    o_ax.set_xlabel("Date")

    # Set y-axis ticks and labels
    o_ax.set_yticks(da_river_miles)
    o_ax.set_yticklabels(df_results_temperature.columns)

    # Draw arrow/label downstream direction
    o_ax.annotate("", xytext=(-.2, 1), xy=(-.2, 0), arrowprops=dict(arrowstyle="->"), xycoords='axes fraction',
                  va='center', ha='center')
    o_ax.annotate("Location, Moving Downstream", xy=(-.22, .5), rotation=90, xycoords='axes fraction',
                  va='center', ha='center')

    # Create colorbar
    o_fig.colorbar(o_contour_plot, label=f'{alt_name} Temperature (Degrees F)')

    # Save water temperature contour figure for this year
    o_fig.savefig(os.path.join(s_yrly_contour_dir, rf"{i_year}_TemperatureContourMap.png"), bbox_inches='tight')
    plt.close(o_fig)
    return

if __name__ == '__main__':
    #DSS contents file from the dss reader
    input_dss_fn = r"C:\Users\fnufferrodriguez\OneDrive - DOI\Desktop\calsim_dss_reader\DSS_contents_temp_test.xlsx"

    #Output directory for plots
    outdir = r'./contour_plots'
    if not os.path.exists(outdir):os.mkdir(outdir)

    #Calendar Years to generate the plots
    i_calendar_yrs = [1977,2014,2021]

    #Read in data
    df_input = pd.read_excel(input_dss_fn, index_col = 0)

    #Format the data
    df_input.replace(value = np.nan,to_replace = -340282346638528897590636046441678635008, inplace = True) #Replace no data placeholder (-340282346638528897590636046441678635008) with np.nan

    #Loop through each scenario and create the contour plot.
    for scenario in df_input.Scenario.unique():
        #Create output directory for this scenario
        outdir_scenario = os.path.join(outdir, scenario)
        if not os.path.exists(outdir_scenario):os.mkdir(outdir_scenario)

        #Grab data for this scenario
        df_contour_input = df_input.loc[df_input.Scenario == scenario]

        #Format data with date as index, and subset to only include the locations of interest. Columns correspond to locations from upstream to downstream.
        df_contour_input.set_index("Date", inplace= True)
        df_contour_input = df_contour_input[[ "BLW SHASTA", 'BLW KESWICK','HWY44','BLW CLEAR CREEK','AIRPORT',]] #From upstream to downstream.
        da_river_miles = [150,125, 100,75,50] #Mock array of river miles corresponding to Airport, Blw Clear Creek, and HWY 44.

        #Generate contour plots for 1977, 2014, and 2021
        for i_year in i_calendar_yrs:
        # for i_year in range(1923, 2022):
            alt_name = "NAA" if scenario == 'Baseline' else scenario
            generate_contour_plot(i_year, outdir_scenario, df_contour_input, da_river_miles, o_cm,
                              d_contour_level_start=48, d_contour_level_end=62, i_levels=8, alt_name = alt_name)



