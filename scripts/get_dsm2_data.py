### This file is meant to recreate what get_dsm2_comp_data_20240221.r from Jacobs does
### to get the full appendix, more steps must be followed. See Jacobs Attachment_F2-08_Scripts/ReadMe.txt
### This for the water quality compliance attachment




import os
import re
import pandas as pd
import datetime
from pydsstools.heclib.dss import HecDss
import numpy as np

def get_stations():
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    stations_df = None

    # Look for the DSM2ComplianceLocations.csv file
    for j in range(len(fnames)):
        if fnames[j] == "DSM2ComplianceLocations.csv":
            stations_df = pd.read_csv(fpaths[j], header=None, sep=",")
            break  # Exit loop once the file is found

    if stations_df is not None:
        stations = stations_df.iloc[:, 0].astype(str).tolist()
        return stations
    else:
        return []  # Return empty list if file not found


def get_locations():
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    stations_df = None

    # Look for the DSM2ComplianceLocations.csv file
    for j in range(len(fnames)):
        if fnames[j] == "DSM2ComplianceLocations.csv":
            stations_df = pd.read_csv(fpaths[j], header=None, sep=",")
            break  # Exit loop once the file is found

    if stations_df is not None:
        stations = stations_df.iloc[:, 1].astype(str).tolist()
        return stations
    else:
        return []  # Return empty list if file not found


def get_stats():
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    stations_df = None

    # Look for the DSM2ComplianceLocations.csv file
    for j in range(len(fnames)):
        if fnames[j] == "DSM2ComplianceLocations.csv":
            stations_df = pd.read_csv(fpaths[j], header=None, sep=",")
            break  # Exit loop once the file is found

    if stations_df is not None:
        stations = stations_df.iloc[:, 2].astype(str).tolist()
        return stations
    else:
        return []  # Return empty list if file not found


def get_sri_current_condition():
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    wyts_df = None

    # Look for the SacRiverIndex.csv file
    for j in range(len(fnames)):
        if fnames[j] == "SacRiverIndex.csv":
            print("\nFound file:", fnames[j])
            wyts_df = pd.read_csv(fpaths[j], header=0, sep=",")
            print("\n")
            break  # Exit loop once the file is found

    return wyts_df


def get_wyts_2022():
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    wyts_df = None

    # Look for the WYT_2022MED.csv file
    for j in range(len(fnames)):
        if fnames[j] == "WYT_2022MED.csv":
            wyts_df = pd.read_csv(fpaths[j], header=0, sep=",")
            break  # Exit loop once the file is found

    return wyts_df


def get_wyts_current_condition():
    # Define the path to the directory
    dir_path = "./../stations"

    # List all files in the directory
    fnames = os.listdir(dir_path)
    fpaths = [os.path.join(dir_path, fname) for fname in fnames]

    # Search for the target file
    for fname, fpath in zip(fnames, fpaths):
        if fname == "WYT_CurrentConditions.csv":
            print("\nFound file:", fname)
            wyts_df = pd.read_csv(fpath, )
            print()
            return wyts_df


def get_specified_table(s_csv_name):
    # Define the path to the stations directory
    stations_dir = "./../stations"

    # Get full paths and file names
    fpaths = [os.path.join(stations_dir, fname) for fname in os.listdir(stations_dir)]
    fnames = os.listdir(stations_dir)

    std_df = None

    # Look for the D1641_AG.csv file
    for j in range(len(fnames)):
        if fnames[j] == s_csv_name:
            std_df = pd.read_csv(fpaths[j], header=0, sep=",")
            break  # Exit loop once the file is found

    return std_df


def get_dsm2_timeseries_data(input_model_name):
    # Construct the path to the model's directory
    file_path = os.path.join("..", "studies", input_model_name)
    print("Searching in:", file_path)

    # List all .dss files in the directory (case-insensitive)
    files = [f for f in os.listdir(file_path) if f.lower().endswith(".dss")]

    print("Model Name:", input_model_name)
    print("DSS Files Found:", files)
    for file_index, file in enumerate(files):
        if "2022" in input_model_name:
            df_sri = get_sri_current_condition()
            df_wyt = get_wyts_2022()
        else:
            print('2022 not in name but still using')
            df_sri = get_sri_current_condition()
            df_wyt = get_wyts_2022()

    # Load csv files for compliance standards
        df_D1641AG = get_specified_table("D1641_AG.csv")
        df_D1641FWS = get_specified_table("D1641_FWS.csv")
        df_D1641MI = get_specified_table("D1641_MI.csv")
        df_D1641MID = get_specified_table("D1641_MID.csv", )
        df_MI_Antioch = get_specified_table("MI_Antioch.csv")
        df_MI_Other  = get_specified_table("MI_Other.csv")
        comp_files = [df_D1641AG,
                           df_D1641FWS,
                           df_D1641MI,
                           df_D1641MID,
                           df_MI_Antioch,
                           df_MI_Other]
        comp_names = ["D1641 AG",
                        "D1641 FWS",
                        "D1641 MI",
                        "D1641 MID",
                        "MI Antioch",
                        "MI Other"]

        # Split the filename by underscores and get the 4th element (index 3)
        parts = file.split("_")
        var_type = parts[3]

        # Construct the full path to the file
        infile = os.path.join(file_path, file)

        with HecDss.Open(infile) as input_file:
            # Retreive stations and locations
            out_statns = get_stations()
            out_locs = get_locations()
            out_stats = get_stats()


            outfile_location = "/".join(infile.split("\\")[:3])

            # Construct the output filename
            tsfilename = f"DSM2ComplianceData_{input_model_name}.csv"
            tsfilename = os.path.join(outfile_location, tsfilename)

            # Open the file for writing
            with open(tsfilename, 'w') as tsfile:
                # Write the header line
                header = "Var Name,Location,Var type,Date,ValueEC,ValueCl,Study Scenario,Study Type,UnitsEC,UnitsCl,D1641AG,D1641FWS,D1641MI,D1641MIDNumDays,D1641MIDThreshold,MIAntiochNumDays,MIAntiochThreshold,MIOther,SAC INDEX\n"

                tsfile.write(header)

            nd_flags = []

            # Loop through locations
            for station_index, station in enumerate(out_statns):
                print(station)

                cl_flag = 0
                nd_flag = 0
                def_stn_flag = 0
                sjr_fws_stn_flag = 0

                bpart = station

                # Set flags based on station name
                if bpart in ["SLCBN002", "SLSUS012"]:
                    def_stn_flag = 1
                if bpart in ["RSAN018", "RSAN032", "RSAN037"]:
                    sjr_fws_stn_flag = 1
                # Create search path
                cpart = f"EC-{out_stats[station_index]}"
                path = f"/*/{bpart}/{cpart}/*/1DAY/*/"
                path_list = input_file.getPathnameList(path, sort=1)

                if path_list == []:
                    continue

                expanded_path = path_list[0].split('/')
                expanded_path[4] = ''
                path = "/".join(expanded_path)
                o_timeseries = input_file.read_ts(path)

                if o_timeseries.empty:
                    continue

                dates = np.array(o_timeseries.pytimes)
                dates = dates + datetime.timedelta(days=-2)

                years = [(date + datetime.timedelta(days=-1)).strftime("%Y") for date in dates]
                months = [(date + datetime.timedelta(days=-1)).strftime("%m") for date in dates]
                days = [(date + datetime.timedelta(days=-1)).strftime("%d") for date in dates]

                values = np.round(o_timeseries.values, 5)

                unitsEC = ["mmhos/cm"] * len(values)
                unitsCl = ["mg/L"] * len(values)
                stations = [out_statns[station_index]] * len(values)
                locations = [out_locs[station_index]] * len(values)
                cpart = [cpart] * len(values)
                study_scenario = [input_model_name] * len(values)
                study_type = ["DSM2"] * len(values)

                # Initialize lists to store results
                wyts = []
                wy_inds = []
                prev_sri = []
                curr_sri = []

                for year_index in range(len(years)):
                    if int(months[year_index]) < 10:
                        wy = years[year_index]

                    else:
                        wy = str(int(years[year_index]) + 1)

                    # Look up water year type and index
                    wyt_row = df_wyt[df_wyt.iloc[:, 0] == int(wy)]
                    wyt = str(wyt_row.iloc[0, 2]) if not wyt_row.empty else None
                    wy_ind = float(wyt_row.iloc[0, 1]) if not wyt_row.empty else None

                    # Look up previous and current SRI
                    p_sri_row = df_sri[df_sri.iloc[:, 0] == int(years[year_index]) - 1]
                    sri_row = df_sri[df_sri.iloc[:, 0] == int(years[year_index])]

                    p_sri = float(p_sri_row.iloc[0, 3]) if not p_sri_row.empty else None
                    sri = float(sri_row.iloc[0, 3]) if not sri_row.empty else None

                    # Append to lists
                    wyts.append(wyt)
                    wy_inds.append(wy_ind)
                    prev_sri.append(p_sri)
                    curr_sri.append(sri)
                # Append the last wyt value to the list
                wyts.append(wyt)
                wyts = wyts[1:]

                # Print unique values
                print("Unique WYT values:", list(set(wyts)))

                # Initialize list for compliance DataFrames
                loc_comp_dfs = [None] *  len(comp_names)

                # Initialize compliance count list with zeros
                comp_count = [0] * len(comp_names)

                for compliance_index in range(len(comp_files)):
                    comp_df = comp_files[compliance_index]

                    # Filter rows where Var Name matches the current station
                    loc_comp_df = comp_df[comp_df["Var Name"] == out_statns[station_index]]

                    if not loc_comp_df.empty:
                        if out_statns[station_index] == "RSAN007":
                            if out_stats[station_index] == "MAX" and comp_names[compliance_index] == "MI Antioch":
                                print(comp_names[compliance_index])
                                comp_count[compliance_index] = 1
                                loc_comp_dfs[compliance_index] = loc_comp_df
                            elif out_stats[station_index] == "MEAN" and comp_names[compliance_index] != "MI Antioch":
                                print(comp_names[compliance_index])
                                comp_count[compliance_index] = 1
                                loc_comp_dfs[compliance_index] = loc_comp_df
                        else:
                            print(comp_names[compliance_index])
                            comp_count[compliance_index] = 1
                            loc_comp_dfs[compliance_index] = loc_comp_df

                # Loop through time series and set standard
                dates_copy = dates.copy()
                std_nms = []
                # std_ts_df = pd.DataFrame()
                count = 0
                for compliance_index in range(len(comp_count)):
                    print('comp_count[compliance_index]: ' + str(comp_count[compliance_index]))
                    if comp_count[compliance_index] < 1:
                        print('less than 1')
                        std_ts = [np.nan] * len(years)
                        std_nm = comp_names[compliance_index]

                        if std_nm in ["D1641 MID", "MI Antioch"]:
                            std_ts2 = std_ts

                            if count < 1:
                                std_ts_df = pd.DataFrame({
                                    f"{std_nm}NumDays": std_ts,
                                    f"{std_nm}Threshold": std_ts2
                                })
                                std_nms.append(str(std_nm))
                                count = 1
                            else:
                                std_ts_df[f"{std_nm}NumDays"] = std_ts
                                std_ts_df[f"{std_nm}Threshold"] = std_ts2
                                std_nms.append(str(std_nm))


                        else:
                            if count < 1:
                                std_ts_df = pd.DataFrame({std_nm: std_ts})
                                std_nms.append(str(std_nm))
                                count = 1
                            else:
                                std_ts_df[std_nm] = std_ts
                                std_nms.append(str(std_nm))

                        print("std_ts_df")
                        print(std_ts_df)
                        print("\n")
                    else:
                        print("\ncomp_count[compliance_index] >= 1")
                        df_comp = loc_comp_dfs[compliance_index]
                        print(df_comp)
                        std_nm = comp_names[compliance_index]

                        std_ts = []
                        std_ts2 = []

                        if not pd.isna(df_comp["NumDays"].iloc[0]):
                            print("NumDays")
                            cl_flag = 1
                            nd_flag = 1

                            for year_index2 in range(len(years)):
                                if year_index2 == 0:
                                    wyt_data = df_comp[df_comp["Sac Index Val"] == wy_inds[year_index2]]
                                    NumDays = wyt_data["NumDays"].iloc[0]
                                    Threshold = wyt_data["Val1"].iloc[0]
                                    NumDays_next = NumDays
                                    Threshold_next = Threshold

                                    if int(months[year_index2]) < 10:
                                        wy_prev = years[year_index2]
                                        yr_prev = years[year_index2]
                                    else:
                                        wy_prev = str(int(years[year_index2]) + 1)
                                        yr_prev = years[year_index2]

                                yr = years[year_index2]
                                wy = years[year_index2] if int(months[year_index2]) < 10 else str(int(years[year_index2]) + 1)

                                if wy != wy_prev:
                                    wyt_data = df_comp[df_comp["Sac Index Val"] == wy_inds[year_index2]]
                                    NumDays_next = wyt_data["NumDays"].iloc[0] if not wyt_data.empty else np.nan
                                    Threshold_next = wyt_data["Val1"].iloc[0] if not wyt_data.empty else np.nan
                                    wy_prev = wy

                                if yr != yr_prev:
                                    NumDays = NumDays_next
                                    Threshold = Threshold_next
                                    yr_prev = yr

                                std_ts.append(NumDays)
                                std_ts2.append(Threshold)

                            # Remove first placeholder if needed (mimicking R's append + [-1])
                            std_ts = std_ts[1:] if len(std_ts) > len(years) else std_ts
                            std_ts2 = std_ts2[1:] if len(std_ts2) > len(years) else std_ts2

                            # Add to std_ts_df
                            if count < 1:
                                std_ts_df = pd.DataFrame({
                                    f"{std_nm}NumDays": std_ts,
                                    f"{std_nm}Threshold": std_ts2
                                })
                                std_nms.append(std_nm)
                                count = 1
                            else:
                                std_ts_df[f"{std_nm}NumDays"] = std_ts
                                std_ts_df[f"{std_nm}Threshold"] = std_ts2
                                std_nms.append(std_nm)

                        else:
                            if std_nm == "D1641 MI":
                                cl_flag = 1
                            for year_index2 in range(len(years)):
                                yr_count = 0
                                if year_index2 == 0:
                                    # Reset deficiency flag
                                    def_flag = 0

                                    # Filter compliance data for the current water year index
                                    wyt_data = df_comp[df_comp["Sac Index Val"] == wy_inds[year_index2]]

                                    # Parse start and end dates for each compliance window

                                    start_1 = datetime.datetime.strptime(wyt_data['Start Date 1'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 1'].iloc[0]) else None
                                    end_1 = datetime.datetime.strptime(wyt_data['End Date 1'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 1'].iloc[0]) else None
                                    start_2 = datetime.datetime.strptime(wyt_data['Start Date 2'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 2'].iloc[0]) else None
                                    end_2 = datetime.datetime.strptime(wyt_data['End Date 2'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 2'].iloc[0]) else None
                                    start_3 = datetime.datetime.strptime(wyt_data['Start Date 3'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 3'].iloc[0]) else None
                                    end_3 = datetime.datetime.strptime(wyt_data['End Date 3'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 3'].iloc[0]) else None
                                    start_4 = datetime.datetime.strptime(wyt_data['Start Date 4'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 4'].iloc[0]) else None
                                    end_4 = datetime.datetime.strptime(wyt_data['End Date 4'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 4'].iloc[0]) else None
                                    start_5 = datetime.datetime.strptime(wyt_data['Start Date 5'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 5'].iloc[0]) else None
                                    end_5 = datetime.datetime.strptime(wyt_data['End Date 5'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 5'].iloc[0]) else None
                                    start_6 = datetime.datetime.strptime(wyt_data['Start Date 6'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 6'].iloc[0]) else None
                                    end_6 = datetime.datetime.strptime(wyt_data['End Date 6'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 6'].iloc[0]) else None

                                    # Store previous year and water year index values
                                    year_prev = int(years[year_index2])
                                    wy_ind_prev = int(wy_inds[year_index2])
                                    wy_ind_prev2 = int(wy_inds[year_index2])
                                year = years[year_index2]
                                if int(year) != year_prev:
                                    print(f"\nYear: {years[year_index2]}")
                                    yr_count += 1
                                    def_flag = 0

                                    # Get new WYT data for this year
                                    wyt_data = df_comp[df_comp["Sac Index Val"] == wy_inds[year_index2]]
                                    print(f"WY index: {wy_inds[year_index2]}")
                                    print(f"Prev Sac River index: {prev_sri[year_index2]}\n")

                                    # Deficiency logic
                                    if wy_inds[year_index2] == 5 and wy_ind_prev >= 4:
                                        def_flag = 1
                                    elif wy_inds[year_index2] == 4 and prev_sri[year_index2] < 11.35:
                                        def_flag = 1
                                    elif wy_inds[year_index2] == 4 and wy_ind_prev >= 4 and wy_ind_prev2 == 5 and yr_count > 1:
                                        def_flag = 1

                                    # SJR FWS logic
                                    sjr_fws_flag = 1 if wy_inds[year_index2] == 4 and curr_sri[year_index2] < 8.1 else 0

                                    # Set compliance windows
                                    if def_flag == 1 and def_stn_flag == 1:
                                        print("Deficiency")
                                        start_1 = datetime.datetime.strptime(f"1-Jan", "%d-%b")
                                        end_1 = datetime.datetime.strptime(f"31-Mar", "%d-%b")
                                        start_2 = datetime.datetime.strptime(f"1-Apr", "%d-%b")
                                        end_2 = datetime.datetime.strptime(f"30-Apr", "%d-%b")
                                        start_3 = datetime.datetime.strptime(f"1-May", "%d-%b")
                                        end_3 = datetime.datetime.strptime(f"31-May", "%d-%b")
                                        start_4 = datetime.datetime.strptime(f"1-Oct", "%d-%b")
                                        end_4 = datetime.datetime.strptime(f"31-Oct", "%d-%b")
                                        start_5 = datetime.datetime.strptime(f"1-Nov", "%d-%b")
                                        end_5 = datetime.datetime.strptime(f"30-Nov", "%d-%b")
                                        start_6 = datetime.datetime.strptime(f"1-Dec", "%d-%b")
                                        end_6 = datetime.datetime.strptime(f"31-Dec", "%d-%b")
                                    else:

                                        start_1 = datetime.datetime.strptime(wyt_data['Start Date 1'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 1'].iloc[0]) else None
                                        end_1 = datetime.datetime.strptime(wyt_data['End Date 1'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 1'].iloc[0]) else None
                                        start_2 = datetime.datetime.strptime(wyt_data['Start Date 2'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 2'].iloc[0]) else None
                                        end_2 = datetime.datetime.strptime(wyt_data['End Date 2'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 2'].iloc[0]) else None
                                        start_3 = datetime.datetime.strptime(wyt_data['Start Date 3'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 3'].iloc[0]) else None
                                        end_3 = datetime.datetime.strptime(wyt_data['End Date 3'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 3'].iloc[0]) else None
                                        start_4 = datetime.datetime.strptime(wyt_data['Start Date 4'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 4'].iloc[0]) else None
                                        end_4 = datetime.datetime.strptime(wyt_data['End Date 4'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 4'].iloc[0]) else None
                                        start_5 = datetime.datetime.strptime(wyt_data['Start Date 5'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 5'].iloc[0]) else None
                                        end_5 = datetime.datetime.strptime(wyt_data['End Date 5'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 5'].iloc[0]) else None
                                        start_6 = datetime.datetime.strptime(wyt_data['Start Date 6'].iloc[0], "%d-%b") if pd.notna(wyt_data['Start Date 6'].iloc[0]) else None
                                        end_6 = datetime.datetime.strptime(wyt_data['End Date 6'].iloc[0], "%d-%b") if pd.notna(wyt_data['End Date 6'].iloc[0]) else None

                                    # Adjust end_1 if SJR FWS condition is met
                                    if sjr_fws_flag == 1 and sjr_fws_stn_flag == 1:
                                        end_1 = datetime.datetime.strptime(f"30-Apr", "%d-%b")

                                    # Update previous year/index trackers
                                    year_prev = int(years[year_index2])
                                    wy_ind_prev2 = wy_ind_prev
                                    wy_ind_prev = int(wy_inds[year_index2])

                                # Add 8 hours to start
                                start_plus = datetime.timedelta(hours=8)

                                # Add one day for days marked 12-31
                                end_plus = datetime.timedelta(hours=32)

                                # Current date being evaluated
                                now = dates_copy[year_index2]

                                val_found = 0
                                reg_val = None

                                if start_1 is not None:
                                    start_1 = start_1.replace(year = int(years[year_index2]))
                                    end_1 = end_1.replace(year = int(years[year_index2]))
                                    if now >= start_1 + start_plus and now < end_1 + end_plus:
                                        reg_val = wyt_data["Val1"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 15.6
                                        val_found = 1
                                    end_1 = end_1.replace(year=2017)
                                if start_2 is not None and val_found == 0:
                                    start_2 = start_2.replace(year = int(years[year_index2]))
                                    end_2 = end_2.replace(year = int(years[year_index2]))
                                    if now >= start_2 + start_plus and now < end_2 + end_plus:
                                        reg_val = wyt_data["Val2"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 14.0
                                        val_found = 1
                                    end_2 = end_2.replace(year=2017)

                                if start_3 is not None and val_found == 0:
                                    start_3 = start_3.replace(year=int(years[year_index2]))
                                    end_3 = end_3.replace(year=int(years[year_index2]))
                                    if now >= start_3 + start_plus and now < end_3 + end_plus:
                                        reg_val = wyt_data["Val3"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 12.5
                                        val_found = 1
                                    end_3 = end_3.replace(year=2017)

                                if start_4 is not None and val_found == 0:
                                    start_4 = start_4.replace(year=int(years[year_index2]))
                                    end_4 = end_4.replace(year=int(years[year_index2]))
                                    if now >= start_4 + start_plus and now < end_4 + end_plus:
                                        reg_val = wyt_data["Val4"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 19.0
                                        val_found = 1
                                    end_4 = end_4.replace(year=2017)

                                if start_5 is not None and val_found == 0:
                                    start_5 = start_5 = start_5.replace(year=int(years[year_index2]))
                                    end_5 = end_5.replace(year=int(years[year_index2]))
                                    if now >= start_5 + start_plus and now < end_5 + end_plus:
                                        reg_val = wyt_data["Val5"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 16.5
                                        val_found = 1
                                    end_5 = end_5.replace(year=2017)

                                if start_6 is not None and val_found == 0:
                                    start_6 = start_6.replace(year=int(years[year_index2]))
                                    end_6 = end_6.replace(year=int(years[year_index2]))
                                    if now >= start_6 + start_plus and now < end_6 + end_plus:
                                        reg_val = wyt_data["Val6"].iloc[0]
                                        if def_flag == 1 and def_stn_flag == 1:
                                            reg_val = 15.6
                                        val_found = 1
                                    end_6 = end_6.replace(year=2017)

                                std_ts.append(reg_val)
                            std_ts = std_ts[1:]
                            std_ts.append(std_ts[-1])

                            if count < 1:
                                std_ts_df = pd.DataFrame({std_nm: std_ts})
                                std_nms = [str(std_nm)]
                                count = 1
                            else:
                                std_ts_df[std_nm] = std_ts
                                std_nms.append(str(std_nm))

                ThirtyDay_Stns = ["RSAN112", "RSAN072", "OLDR_MIDR", "ROLD059"]
                Monthly_Stns = ["RSAC081", "SLMZU025", "SLMZU011", "SLCBN002", "SLSUS012", "CHWST000", "CHDMC004"]
                new_EC_vals = []

                if out_statns[station_index] in ThirtyDay_Stns:
                    for k in range(len(values)):
                        if k < 29:
                            new_val = np.mean(values[:k+1])
                        else:
                            sub_vals = []
                            for k2 in range(0, 30):
                                val = values[k-k2]
                                if val > -1000:
                                    sub_vals.append(val)

                            new_val = np.mean(sub_vals)
                        new_val = new_val / 1000
                        new_EC_vals.append(new_val)
                elif out_statns[station_index] in Monthly_Stns:

                    daily_df = pd.DataFrame({'Date': dates, 'Value': values})
                    daily_df['Month'] = [date.month for date in dates]
                    daily_df['Year'] = [date.year for date in dates]

                    month_df = daily_df.groupby(['Year', 'Month'], as_index=False).mean()

                    month_df['Date'] = [datetime.datetime(1921, 3, 1) + pd.DateOffset(months=i) + datetime.timedelta(days=-1) for i in range(len(month_df))]
                    for k in range(len(values)):
                         m = daily_df['Month'][k]
                         y = daily_df['Year'][k]
                         sub_df = month_df[month_df['Month'] == m]
                         sub_df = sub_df[sub_df['Year'] == y]
                         new_val = sub_df['Value'].iloc[0] / 1000
                         new_EC_vals.append(new_val)
                else: # 14 days
                    for k in range(len(values)):
                        if k < 13:
                            new_val = np.mean(values[:k+1])
                        else:
                            sub_vals = []
                            for k2 in range(0, 14):
                                val = values[k-k2]
                                if val > -1000:
                                    sub_vals.append(val)

                            new_val = np.mean(sub_vals)
                        new_val = new_val / 1000
                        new_EC_vals.append(new_val)

                if cl_flag > 0:
                    new_Cl_vals = []
                    for val in values:
                        if val < -1000:
                            new_val = None
                        elif val <=280:
                            new_val = val*0.15 - 12
                        else:
                            new_val = val*0.285 - 50
                        new_Cl_vals.append(new_val)
                else:
                    new_Cl_vals = [None] * len(values)

                nd_flags.append(nd_flag)

                df = pd.DataFrame({
                    "Var Name": stations,
                    "Location": locations,
                    "Var type": cpart,
                    "Date": dates,
                    "ValueEC": new_EC_vals,
                    "ValueCl": new_Cl_vals,
                    "Study Scenario": study_scenario,
                    "Study Type": study_type,
                    "UnitsEC": unitsEC,
                    "UnitsCl": unitsCl
                })
                df = pd.concat([df, std_ts_df], axis=1)

                # Add water year type column
                df["WYT"] = wyts

                # Filter to start in WY 1921 (i.e., Oct 1921 or later)
                df = df[(df["Date"].dt.year > 1921) | ((df["Date"].dt.year == 1921) & (df["Date"].dt.month > 9))]

                # Remove rows with invalid EC values
                df = df[df["ValueEC"] > -1000000]

                # Remove rows where all values are NA
                df = df.dropna(how='all')

                # Rename columns to match final output
                df.columns = [
                    "Var Name", "Location", "Var type", "Date", "ValueEC", "ValueCl",
                    "Study Scenario", "Study Type", "UnitsEC", "UnitsCl",
                    "D1641AG", "D1641FWS", "D1641MI", "D1641MIDNumDays", "D1641MIDThreshold",
                    "MIAntiochNumDays", "MIAntiochThreshold", "MIOther", "SAC INDEX"
                ]

                # Write to file (no header, no row names, comma-separated)
                df.to_csv(tsfilename, index=False, header=False, mode='a')  # quoting=3 disables quoting (like quote = FALSE in R)

                # Optionally clear variables (not necessary in Python, but for memory cleanup)
                del df, stations, locations, cpart, dates, new_EC_vals, new_Cl_vals, study_scenario, study_type, unitsEC, unitsCl, values

        file_path = os.path.join("../studies",input_model_name)
        outfile_location = file_path
        tsfilename = "DSM2ComplianceData_" + input_model_name + ".csv"
        tsfilename = os.path.join(file_path, tsfilename)

        # Read the output ts file
        df_pre = pd.read_csv(tsfilename, skiprows=1, header=None)
        df_pre_names = pd.read_csv(tsfilename, nrows=1, header=None)

        # Assign column names from the first row
        df_pre.columns = df_pre_names.iloc[0]

        # Calculate differences
        ECval = df_pre['ValueEC']
        Clval = df_pre['ValueCl']
        AG = df_pre['D1641AG']
        FWS = df_pre['D1641FWS']
        MI = df_pre['D1641MI']

        df_pre['AG_diff'] = ECval - AG
        df_pre['FWS_diff'] = ECval - FWS
        df_pre['MI_diff'] = Clval - MI

        ndblw_arr = []
        ndvio_arr = []
        ndblw_arrA = []
        ndvio_arrA = []

        for ind in range(len(out_statns)):
            stn = out_statns[ind]
            v_type = f"EC-{out_stats[ind]}"

            df_stn2 = df_pre[df_pre['Var Name'] == stn]
            df_stn = df_stn2[df_stn2['Var type'] == v_type]

            if df_stn.empty:
                continue

            MIDT = df_stn['D1641MIDThreshold'].values
            AntiochT = df_stn['MIAntiochThreshold'].values

            if not (pd.isna(MIDT[0]) or not pd.isna(AntiochT[0])):
                dt = pd.to_datetime(df_stn['Date'])
                years = [date.year for date in dt]
                months = [date.month for date in dt]
                Clval = df_stn['ValueCl'].values

                if out_stats[ind] == 'MEAN':
                    MIDND = df_stn['D1641MIDNumDays'].values
                    itr = 0
                    yrtot = 0
                    for k in range(len(df_stn)):
                        if k == 0:
                            yr_prev = years[k]
                        yr = years[k]

                        if yr != yr_prev:
                            yrtot += itr
                            ndblw_arr[-1] = yrtot
                            if np.isnan(yrtot) or np.isnan(MIDND[k - 1]):
                                ndvio_arr[-1] = 0
                            elif yrtot < MIDND[k - 1]:
                                ndvio_arr[-1] = 1
                            else:
                                ndvio_arr[-1] = 0
                            ndblw_arr.append(np.nan)
                            ndvio_arr.append(0)
                            yrtot = 0
                            itr = 0
                            yr_prev = yr
                        elif k == len(df_stn) - 1:
                            yrtot += itr
                            ndblw_arr.append(yrtot)
                            if np.isnan(yrtot) or np.isnan(MIDND[k]):
                                ndvio_arr.append(0)
                            elif yrtot < MIDND[k]:
                                ndvio_arr.append(1)
                            else:
                                ndvio_arr.append(0)
                        else:
                            ndblw_arr.append(np.nan)
                            ndvio_arr.append(0)

                        if np.isnan(MIDT[k]) or np.isnan(Clval[k]):
                            continue
                        elif Clval[k] < MIDT[k]:
                            itr += 1
                        elif itr >= 14:
                            yrtot += itr
                            itr = 0
                        else:
                            itr = 0

                    ndblw_add = [np.nan] * len(df_stn)
                    ndblw_arrA.extend(ndblw_add)
                    ndvio_arrA.extend(ndblw_add)

                else:
                    AntiochND = df_stn['MIAntiochNumDays'].values
                    itrA = 0
                    for k in range(len(df_stn)):
                        if k == 0:
                            yr_prev = years[k]
                        yr = years[k]

                        if yr != yr_prev:
                            ndblw_arrA[-1] = itrA
                            if np.isnan(itrA) or np.isnan(AntiochND[k - 1]):
                                ndvio_arrA[-1] = 0
                            elif itrA < AntiochND[k - 1]:
                                ndvio_arrA[-1] = 1
                            else:
                                ndvio_arrA[-1] = 0
                            ndblw_arrA.append(np.nan)
                            ndvio_arrA.append(0)
                            itrA = 0
                            yr_prev = yr
                        elif k == len(df_stn) - 1:
                            ndblw_arrA.append(itrA)
                            if np.isnan(itrA) or np.isnan(AntiochND[k]):
                                ndvio_arrA.append(0)
                            elif itrA < AntiochND[k]:
                                ndvio_arrA.append(1)
                            else:
                                ndvio_arrA.append(0)
                        else:
                            ndblw_arrA.append(np.nan)
                            ndvio_arrA.append(0)

                        if np.isnan(AntiochT[k]) or np.isnan(Clval[k]):
                            continue
                        elif Clval[k] < AntiochT[k]:
                            itrA += 1

                    ndblw_arr.extend([np.nan] * len(df_stn))
                    ndvio_arr.extend([np.nan] * len(df_stn))

            else:
                ndblw_add = [np.nan] * len(df_stn)
                ndblw_arr.extend(ndblw_add)
                ndvio_arr.extend(ndblw_add)
                ndblw_arrA.extend(ndblw_add)
                ndvio_arrA.extend(ndblw_add)

        # Assign new columns to df_pre
        df_pre['NumDaysBlw'] = ndblw_arr
        df_pre['NumDaysViolated'] = ndvio_arr
        df_pre['AntiochNumDaysBlw'] = ndblw_arrA
        df_pre['AntiochNumDaysViolated'] = ndvio_arrA
        # Construct output filename
        tsfilenamediff = f"DSM2ComplianceDiffData_{input_model_name}.csv"
        tsfilenamediff = os.path.join(outfile_location, tsfilenamediff)

        # Define custom header
        custom_header = [
            "Var Name", "Location", "Var type", "Date", "ValueEC", "ValueCl", "Study Scenario", "Study Type",
            "UnitsEC", "UnitsCl", "D1641AG", "D1641FWS", "D1641MI", "D1641MIDNumDays", "D1641MIDThreshold",
            "MIAntiochNumDays", "MIAntiochThreshold", "MIOther", "SAC INDEX", "DiffAG", "DiffFWS", "DiffMI",
            "NumDaysBlw", "NumDaysViolated", "AntiochNumDaysBlw", "AntiochNumDaysViolated"
        ]

        # Write to CSV with custom header
        with open(tsfilenamediff, 'w') as f:
            f.write(','.join(custom_header) + '\n')
            df_pre.to_csv(f, index=False, header=False, na_rep='NA')


if __name__ == '__main__':
    # Set working directory to the script's location
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    # List all subdirectories in ../studies that contain "ALT" or "NAA"
    studies_path = os.path.abspath(os.path.join(script_dir, "../studies"))
    dirnames = [d for d in os.listdir(studies_path)
                if os.path.isdir(os.path.join(studies_path, d)) and re.search(r"(ALT|NAA)", d)]

    print("Model directories found:")
    print(dirnames)

    # Loop through each model directory and call the processing function
    for model_name in dirnames:
        print(f"Processing model: {model_name}")
        get_dsm2_timeseries_data(model_name)

