"""
Functions to convert storage to elevation. May end up moving this to the EIS Appendix Gen file instead.
"""

import pandas as pd
import numpy as np


#Function for storage to elevation conversion
def storage_to_elevation(df_storage, field, storage_elevation_fn , orig_unit = 'TAF' ):
    """
    Converts storage to elevation, for the field specified. Uses the storage-elevation data stored in storage_elevation_fn.

    Note that storage_elevation_fn is the data included in the Calsim's wresl code Run/Lookup/res_info.table file. San
    Luis reservoir storage-elevation curve is derived from summing the storages for the San Luis CVP and SWP storages in
    the storage elevation table.

    Parameters
    ----------
    df_storage: pandas dataframe
        Dataframe containing storage timeseries from CalSim output, as passed from parse_dssReader_output.
        df_storage's {field} column must store the storage timeseries.
    field: str
        column name in df_storage that contains the storage timeseries.
    storage_elevation_fn: str
        file name of storage-elevation table data. See note in function description.
    orig_unit: str
        original field's unit. Conversion to acre-ft is implemented in function.

    Returns
    -------
    df_elev: pandas dataframe
        Elevation timeseries corresponding to CalSim storage timeseries.
    """
    da_storage = df_storage[field].values #array of storages
    if orig_unit == 'TAF':
        da_storage_af = df_storage[field].values*1000
    else: 
        raise ValueError("Need to implement conversion of storage units to AF")
    #Open the storage elevation tables
    df_s_to_el = pd.read_excel(storage_elevation_fn, sheet_name = None)
    df_res_name = df_s_to_el['res_name']
    df_s_e = df_s_to_el['res_info'].set_index('res_num')
    #Get the correct storage elevation table

    df_table = df_s_e.loc[df_res_name.loc[df_res_name.calsim_name == field].res_num][['storage (AF)', 'elevation (ft)']]

    #Check that the CALSIM storages do not fall outside of the storage-elevation table values.
    if da_storage_af.max()> df_table['storage (AF)'].values.max():
        raise ValueError(f"CALSIM Storage {da_storage_af.max()}AF in {field} is greater than max value in storage-elevation tables.")
    elif da_storage_af.min()< df_table['storage (AF)'].values.min():
        raise ValueError(f"CALSIM Storage in {field} is less than min value in storage-elevation tables.")

    #Convert storage to elevation using linear interpolation.
    da_elev = np.interp(da_storage_af, df_table['storage (AF)'].values, df_table['elevation (ft)'])
    
    #return a dataframe with the elevations
    df_elev = df_storage.copy(deep = True)
    df_elev[field] = da_elev
    #For debugging purposes, save to csv
    df_elev.to_csv(rf"_debug_{field}_converted_elevs.csv")
    return df_elev