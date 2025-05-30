import pandas as pd
import numpy as np


# Function for storage to elevation conversion
def ec_to_cl(df_ec, field, orig_unit=''):
    """
    Converts EC to mg/L Cl, for the field specified. Uses the regression relationship defined as equation 2 in
    https://www.waterboards.ca.gov/waterrights/water_issues/programs/bay_delta/california_waterfix/exhibits/docs/ccc_cccwa/CCC-SC_25.pdf

    Cl (mg/L) = max(0.15*EC-12, 0.285*EC-50)

    Parameters
    ----------
    df_ec: pandas dataframe
        Dataframe containing storage timeseries from CalSim output, as passed from parse_dssReader_output.
        df_storage's {field} column must store the storage timeseries.
    field: str
        column name in df_storage that contains the storage timeseries.
    storage_elevation_fn: str
        file name of storage-elevation table data. See note in function description.
    orig_unit: str
        original field's unit. Acceptable orig_unit is uS/cm (microSiemens/cm).

    Returns
    -------
    df_elev: pandas dataframe
        Elevation timeseries corresponding to CalSim storage timeseries.
    """
    da_ec = df_ec[field].values  # array of EC in units defined by orig_unit
    if orig_unit == "uS/cm":
        #Convert EC in uS/cm to mg/L Cl using the equation: Cl (mg/L) = max(0.15*EC-12, 0.285*EC-50)
        da_Cl_option1 = 0.15*da_ec -12
        da_Cl_option2 = 0.285*da_ec-50
        da_Cl = np.max([da_Cl_option1, da_Cl_option2], axis = 0)
    else:
        raise ValueError("Need to implement conversion of storage units to AF")

    # return a dataframe with the mg/L Cl concentration
    df_Cl = df_ec.copy(deep = True)
    df_Cl[field] = da_Cl

    #For debug purposes, output the converted Cl concentration.
    df_Cl.to_csv(rf"_debug_{field}_converted_Cl.csv")

    return df_Cl