# EIS Appendix Generation

## Name
EIS Appendix Generation

## Description
Automatically generate appendices for EIS reports containing VI/508 Compliant tables and figures using the CSV output from the DSS Reader.

Note: 
- All WY type statistics are calculated by sorting monthly data into years determined by the `use_calendar_yr` flag. 
- "Calendar year-year type sorting" means that instead of classifying October 1980 as the same WYType as WY 1981 (the water year it is part of), 
you would classify it by the same WYType as WY1980. This affects the WYType average plots (last 5 plots for each location)
and the WYType averages in the last 5 rows of each table.
- If generating storage-elevation appendix, check that the ./inputs/storage_elevation_table.xlsx is up-to-date with what you want to use. The current version of the storage-elevation table is from the Sac 2021 LTO. A more up-to-date Oroville bathymetry is used in post-LTO modeling, but is **not** the version included in this repo. 
- Check that the ./inputs/wy_flags.xlsx table matches your CalSim run's hydrology. If not, follow step 3.0) below to update it. 
- Automatic caption numbering for figure and tables inherits the Heading 2 numbering. 

## Installation
You will need local copies of this repo and the CalSim DSS Reader: https://gitlab.bor.doi.net/usbr-cvp-modeling/calsim_dss_reader

Use appendix_gen.yml to create a conda environment with necessary packages for running the scripts using the command below:

conda env create -f appendix_gen.yml

## Usage
1. Run the DSS Reader 
	1) Open dssReader.py in your local version of the CalSim DSS Reader. 
	2) In line 35, set the `model` variable equal to the string name of the type of model you wish to interpret data from
	(options are "CALSIM", "HEC5Q" or "DSM2")
	3) Beginning in line 37 in the `runs` list, for each list entry in `runs`, enter the name of each of your dss files in the parentheses along with the name of the run 
	(such as Baseline, Alt1, etc.). Write the file names without using quotation marks. Refer to the NAA scenario as "Baseline". A NAA/Baseline scenario must be included in the runs for 
	the appendix generation script to function properly down the line. Don't forget the ".dss" file extension when you are specifying file names.
	4) Beginning line 56, in the `add_field_list`, specify the field variables that you want to retrieve from the DSS files. These correspond to the B part in the DSS pathname.
	5) Run dssReader.py.
	6) When the DSS Reader has finished running, open the calsim_dss_reader directory and find the DSS Reader outputs. There should
	be three files: DSS_contents.xlsx, DSS_contents_CFS.xlsx, and DSS_contents_TAF.xlsx. The first output file preserves all units
	from the input dss file, the second converts relevant columns to CFS, and the third converts to TAF. 
2. Copy the DSS Reader output file with the desired units for your model and paste the output into the eis_appendix_generation directory in the 
"inputs" folder.
3. If running a temperature appendix, run the DSS Reader to extract the "SHASTABIN_" variable from the HEC-5Q inputs file "CALSIMII_HEC5Q.DSS" files for each alt. 
   1) Use "CALSIM" for the `model` variable. 
   2) Use the corresponding "CALSIMII_HEC5Q.DSS" files for each alt in the `runs` list. Note that only alternatives including the Shasta action will have the SHASTABIN_ variable. 
   3) Assign `list = ['SHASTABIN_']` to `add_field_list`. 
   4) Run dssReader.py
4. Run the EIS Appendix Generation scripts 

	0) **Optional**: Preprocess the WY type datasets to get the final water year type determination for each water year. **This only needs to be done once if your climate scenario is the same.** 
       1) Open process_wytypes.py and assign `s_dvfile` to the file path for the DSS_contents.xlsx file containing CalSim output Water year type variables. Assign `s_output` to the path corresponding to /inputs/wy_flags.xlsx. 
       2) Run process_wytypes.py.

	1) Open the EISAppendixGen.py script. 
	2) In the `fields` list, specify the same field variables that you specified in the DSS Reader to retrieve from the DSS files. These correspond to the B part
	in the DSS pathname.
	3) Define the `alts` list, specify the same run names that you provided in the DSSReader (such as NAA, Alt1, etc.) All names should be exactly the same, except that
	"Baseline" should be referred to as "NAA" in these scripts. Write the file names without quotation marks.
	4) Define the `report_type` variable as either "flow", "elevation", "diversion", or "water supply" for a CalSim appendices, "temperature" for a HEC5Q appendix, 
	or "EC", "Cl", or "Position" for a DSM2 (salinity) appendices.
   5) If running a temperature report:
      1) In the `compliance_fields` list, list the compliance location fields. For the 2021LTO mixed compliance location logic, this should be set to ['AIRPORT', 'BLW CLEAR CREEK', 'HWY44']
      2) In the `compliance_dict`, set the keys to the SHASTABIN_ and the values to the corresponding compliance locations used for those SHASTABIN_ types. For the 2021 LTO mixed compliance location logic, use the following: 
   ```
     compliance_dict = {
        1: 'AIRPORT',          #Shasta Bin1A has the most downstream compliance location
        2: 'AIRPORT',          #Shasta Bin1B has the most downstream compliance location
        3: 'BLW CLEAR CREEK',  #Shasta Bin2A has the middle compliance location
        4: 'BLW CLEAR CREEK',  #Shasta Bin2B has the middle compliance location
        5: 'HWY44',            #Shasta Bin3A has the most upstream compliance location
        6: 'HWY44',            #Shasta Bin3B has the most upstream compliance location
    }
   ```
	6) Define the `appendix_prefix` variable with the prefix you want for all appendix tables and figures in your report. Include a leading space.
	7) Make sure `dss_path` variable correctly references the DSS_contents output file you copied over. Also, make sure that the parent directory is correct
	for where you have your eis_appendix_gen local directory stored. 
    8) Make sure that `wy_flags_path`, `doc_name`, `location_cw_path`, `storage_elevation_table`, and `new_doc` contain the correct parent directory for your local copy of eis_appendix_gen.
	9) Set `use_calendar_yr` to True to use calendar year-year type sorting. Set `use_calendar_yr` to False to use water year-year type sorting. 
   10) Run the EISAppendixGen.py script
5. The EIS Appendix output will be a Microsoft Word Document in the eis_appendix_gen directory under the name f"appendix_final_{report_type}.docx". 
6. After the script finishes running, open the Word document and **Ctrl+A** to select all. Then press **F9** to generate the table and figure numbers. 

## Water Temperature Contour Plots Generation
Action 5 documentation also included contour plots of temperatures along the Sacramento River, at 5 selected locations. Distances of locations downstream are approximate. This script also uses monthly inputs. 
1) Open create_contour_plots.py
2) Edit `input_dss_fn` to be the excel file name that contains the temperature at locations you want included in the contour plots. This file must be in the format outputted by the DSS reader, in a monthly timestep. 
3) Modify `outdir` to be your desired output directory. 
4) Modify `i_calendar_yrs` to be the years you want to generate contour plots.
5) Modify locations to include in the contour plot (`df_contour_input` subset in line 153) and their corresponding river miles (`da_river_miles` in line 154) as needed. 
6) Run create_contour_plots.py

## Usage (Water Quality Compliance)
1. Create a folder called `studies` and place all DSM2 output files in the folder. These will typically be the files that end with 'EC_p'.
2. Ensure that each alternative has the correct hydrology in the file name (ex: 2022MED).
3. Open get_dsm2_data.py
4. Update `scenario_names` to contain the names and files for each alternative. The structure of this dictionary is '<display name>': '<file name>'.
5. Make sure that `doc_name` and `new_doc` contain the paths (not on OneDrive) for the temporary and final doc.
6. Run get_dsm2_data.py
7. The output will be a Microsoft Word Document with the name specified in `new_doc`.
8. After the script finishes running, open the Word document and **Ctrl+A** to select all. Then press **F9** to generate the figure numbers. 

## Support
Please contact rlucas@usbr.gov for support.
