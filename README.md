# EIS Appendix Generation

## Name
EIS Appendix Generation

## Description
Automatically generate appendices for EIS reports containing VI/508 Compliant tables and figures using the CSV output from the DSS Reader.

Note: 
- All WY type statistics are calculated by sorting monthly data into corresponding water years. Ex: Oct 1980 data is assigned the WY type associated with WY 1981. 
- If generating storage-elevation appendix, check that the ./inputs/storage_elevation_table.xlsx is up-to-date with what you want to use. The current version of the storage-elevation table is from the Sac 2021 LTO. A more up-to-date Oroville bathymetry is used in post-LTO modeling, but is **not** the version included in this repo. 
- Check that the ./inputs/wy_flags.xlsx table matches your CalSim run's hydrology. If not, follow step 3.0) below to update it. 
- Automatic caption numbering for figure and tables inherits the Heading 2 numbering.
- Input to the EIS Appendix Generation script is assumed to be monthly, unless the report type is temperature. In the case of temperature, values are averaged for each month.  

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
3. Run the EIS Appendix Generation scripts 

	0) **Optional**: Preprocess the WY type datasets to get the final water year type determination for each water year. **This only needs to be done once if your climate scenario is the same.** 
       1) Open process_wytypes.py and assign `s_dvfile` to the file path for the DSS_contents.xlsx file containing CalSim output Water year type variables. Assign `s_output` to the path corresponding to /inputs/wy_flags.xlsx. 
       2) Run process_wytypes.py.

	1) Open the EISAppendixGen.py script. 
	2) In the `fields` list, specify the same field variables that you specified in the DSS Reader to retrieve from the DSS files. These correspond to the B part
	in the DSS pathname.
	3) Define the `alts` list, specify the same run names that you provided in the DSSReader (such as NAA, Alt1, etc.) All names should be exactly the same, except that
	"Baseline" should be referred to as "NAA" in these scripts. Write the file names without quotation marks.
	4) Define the `report_type` variable as either "flow", "elevation", or "diversion" for a CalSim appendices, "temperature" for a HEC5Q appendix, 
	or "EC", "Cl", or "X2" for a DSM2 (salinity) appendices.
	5) Define the `appendix_prefix` variable with the prefix you want for all appendix tables and figures in your report. Include a leading space.
	7) Make sure `dss_path` variable correctly references the DSS_contents output file you copied over. Also, make sure that the parent directory is correct
	for where you have your eis_appendix_gen local directory stored. 
    8) Make sure that `wy_flags_path`, `doc_name`, `location_cw_path`, `storage_elevation_table`, and `new_doc` contain the correct parent directory for your local copy of eis_appendix_gen.
	9) Run the EISAppendixGen.py script
4. The EIS Appendix output will be a Microsoft Word Document in the eis_appendix_gen directory under the name f"appendix_final_{report_type}.docx". 
5. After the script finishes running, open the Word document and **Ctrl+A** to select all. Then press **F9** to generate the table and figure numbers. 

## Support
Please contact emadonna@usbr.gov for support.
