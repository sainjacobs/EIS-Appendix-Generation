# EIS Appendix Generation

## Name
EIS Appendix Generation

## Description
Automatically generate appendices for EIS reports containing VI/508 Compliant tables and figures using the CSV output from the DSS Reader.


## Installation
You will need local copies of this repo and the CalSim DSS Reader: https://gitlab.bor.doi.net/usbr-cvp-modeling/calsim_dss_reader

Use appendix_gen.yml to create a conda environment with necessary packages for running the scripts using the command below:

conda env create -f apendix_gen.yml

## Usage
1. Run the DSS Reader
	1) Open dssReader.py in your local version of the CalSim DSS Reader. 
	2) In line 35, set the "model" variable equal to the string name of the type of model you wish to interpret data from
	(options are "CALSIM", "HEC5Q" or "DSM2")
	3) Beginning in line 37 in the "runs" list, for each list entry in "runs", enter the name of each of your dss files in the parentheses along with the name of the run 
	(such as Baseline, Alt1, etc.). Write the file names without using quotation marks. Refer to the NAA scenario as "Baseline". A NAA/Baseline scenario must be included in the runs for 
	the appendix generation script to function properly down the line. . Don't forget the ".dss" file extension when you are specifying file names.
	4) Beginning line 56, in the "add_field_list", specify the field variables that you want to retrieve from the DSS files. These correspond to the B part in the DSS pathname.
	5) Run dssReader.py.
	6) When the DSS Reader has finished running, open the calsim_dss_reader directory and find the DSS Reader outputs. There should
	be three files: DSS_contents.xlsx, DSS_contents_CFS.xlsx, and DSS_contents_TAF.xlsx. The first output file preserves all units
	from the input dss file, the second converts relevant columns to CFS, and the third converts to TAF. 
2. Copy the DSS Reader output file with the desired units for your model and paste the output into the eis_appendix_generation directory in the 
"inputs" folder.
3. Run the EIS Appendix Generation scripts
	1) Open the EISAppendixGen py script that corresponds to the type of report you want to generate an appendix for: EISAppendixGen.py for a 
	flow, elevation, or diversion report using CalSim model runs; EISAppendixGenSalinity.py for a salinity report using HEC5Q runs; EISAppendixGenTemp.py
	for a temperature report using DSM2 model runs.
	2) In the "fields" list on line 15, specify the same field variables that you specified in the DSS Reader to retrieve from the DSS files. These correspond to the B part
	in the DSS pathname.
	3) In line 18 in the "alts" list, specify the same run names that you provided in the DSSReader (such as NAA, Alt1, etc.) All names should be exactly the same, except that
	"Baseline" should be referred to as "NAA" in these scripts. Write the file names without quotation marks.
	4) In line 20, define the "report_type" variable as either "flow", "elevation", or "diversion" for a CalSim appendix, "temperature" for a HEC5Q appendix, 
	or "salinity" for a DSM2 appendix.
	5) In line 24, define the "appendix_prefix" with the prefix you want for all appendix tables and figures in your report. Include a leading space.
	6) In line 27, make sure the correct crosswalk file is referenced in "location_cw_path" depending on the type of report you are generating. The file
	name should include "CalSim", "salinity", or "Temp". Also, make sure that you change the parent directory to reflect the absolute path to your eis_appendix_gen
	local directory.
	7) In line 30, make sure "dss_path" correctly references the DSS_contents output file you copied over. Also, make sure that the parent directory is correct
	for where you have your eis_appendix_gen local directory stored.
	8) In lines 32 - 39, make sure that "wy_flags_path", "doc_name", and "new_doc" contain the correct parent directory for your local copy of eis_appendix_gen.
	9) Run the EISAppendixGen py script
4. The EIS Appendix output will be a Microsoft Word Document in the eis_appendix_gen directory under the name "appendix_final.docx".

## Support
Please contact emadonna@usbr.gov for support.
