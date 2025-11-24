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
You will need local copies of this repo and the CalSim DSS Reader: https://gitlab.bor.doi.net/bdo-modeling/calsim_dss_reader

## Environment Set Up
Set up an environment by running `conda env create -f appendix_gen.yml`

## Usage
1. Run the DSS Reader 
	1) Open dssReader.py in your local version of the CalSim DSS Reader. 
	2) In line 35, set the `model` variable equal to the string name of the type of model you wish to interpret data from
	(options are "CALSIM", "HEC5Q" or "DSM2")
	3) Beginning in line 37 in the `runs` list, for each list entry in `runs`, enter the name of each of your dss files in the parentheses along with the name of the run 
	(such as Baseline, Alt1, etc.). Refer to the NAA scenario as "Baseline". A NAA/Baseline scenario must be included in the runs for 
	the appendix generation script to function properly down the line. Don't forget the ".dss" file extension when you are specifying file names.
	4) Beginning line 56, in the `add_field_list`, specify the field variables that you want to retrieve from the DSS files. These correspond to the B part in the DSS pathname. Suggested fields are below in the Suggested Fields section.
	5) Run dssReader.py.
	6) When the DSS Reader has finished running, open the calsim_dss_reader directory and find the DSS Reader outputs. There should
	be three files: DSS_contents.xlsx, DSS_contents_CFS.xlsx, and DSS_contents_TAF.xlsx. The first output file preserves all units
	from the input dss file, the second converts relevant columns to CFS, and the third converts to TAF. 
   7) Save the file you want to use for the appendix.
2. If running a temperature appendix, run the DSS Reader again to extract the "SHASTABIN_" variable from the HEC-5Q inputs file "CALSIMII_HEC5Q.DSS" files for each alt. 
   1) Use "CALSIM" for the `model` variable. 
   2) Use the corresponding "CALSIMII_HEC5Q.DSS" files for each alt in the `runs` list. Note that only alternatives including the Shasta action will have the SHASTABIN_ variable. 
   3) Assign `list = ['SHASTABIN_']` to `add_field_list`. 
   4) Run dssReader.py
   5) In the process_shastabin.py file if this repository, change `s_input_data` to point to the DSS_contents.xlsx output of dssReader.py
   6) Run process_shastabin.py
3. Run the EIS Appendix Generation scripts
   0) **Optional**: Preprocess the WY type datasets to get the final water year type determination for each water year. **This only needs to be done once if your climate scenario is the same.** 
      1) Open process_wytypes.py and assign `s_dvfile` to the file path for the DSS_contents.xlsx file containing CalSim output Water year type variables. Assign `s_output` to the path corresponding to /inputs/wy_flags.xlsx. 
      2) Run process_wytypes.py.
   1) Open the EIS_appendix_gen_....py file corresponding to the type of report you wish to create. 
   2) In the `fields` list, specify the same field variables that you specified in the DSS Reader to retrieve from the DSS files. These correspond to the B part
   in the DSS pathname. Suggested fields are below in the Suggested Fields section and provided in the python file.
   3) Define the `alts` list, specify the same run names that you provided in the DSSReader (such as NAA, Alt1, etc.) All names should be exactly the same, except that
   "Baseline" should be referred to as "NAA" in these scripts.
   4) If creating a CalSim or salinity appendix, define the `report_type` variable as either "flow", "elevation", or "diversion" for a CalSim appendices or "EC", "Cl", or "Position" for a DSM2 (salinity) appendix.
   5) if creating a compliance appenddix, set `scenario_names` to contain the names and files for each alternative. The structure of this dictionary is '<display name>': '<file name>'.
   6) Define the `appendix_prefix` variable with the prefix you want for all appendix tables and figures in your report. Include a leading space. Recommended options are provided.
   7) Make sure `dss_path` variable correctly references the DSS_contents output file. Use the original units for temperature, salinity, and compliance, use TAF for elevations and water supply, and use CFS for diversions and flow.
   8) Set `doc_name` and `new_doc` to full paths to the locations to hold a temporary document and final document. These must be full paths with no spaces (no OneDrive) for VBA to work.
   9) Make sure that any other variables (`template`, `wy_flags_path`, `location_cw_path`, `storage_elevation_table`, `use_lumped_table_captions`, `shastabin_data_path` (temperature only), `compliance_fields` (temperature only), and `compliance_dict` (temperature only)) are set the correct values, but these should not need to change.
   For all appendices except for compliance and water supply, set `use_calendar_yr` to True to use calendar year-year type sorting. Set `use_calendar_yr` to False to use water year-year type sorting. 
   11) Run the EIS_appendix_gen_....py script
5. The EIS Appendix output will be a Microsoft Word Document with the `new_doc` location/name. 
6. After the script finishes running:
   1. Open the Word document and **Ctrl+A** to select all. Then press **F9** to generate the table and figure numbers. 
   2. For the Heading 2 Numbering, you may have to adjust it to match the appendix_prefix variable (Ex: 'F.2.2') by right-clicking and selecting 'Adjust List Indents'. Then modify the numbering to match appendix_prefix under 'Enter formatting for number:'

## Water Temperature Contour Plots Generation
Action 5 documentation also included contour plots of temperatures along the Sacramento River, at 5 selected locations. Distances of locations downstream are approximate. This script also uses monthly inputs. 
1) Open create_contour_plots.py
2) Edit `input_dss_fn` to be the excel file name that contains the temperature at locations you want included in the contour plots. This file must be in the format outputted by the DSS reader, in a monthly timestep. 
3) Modify `outdir` to be your desired output directory. 
4) Modify `i_calendar_yrs` to be the years you want to generate contour plots.
5) Modify locations to include in the contour plot (`df_contour_input` subset in line 153) and their corresponding river miles (`da_river_miles` in line 154) as needed. 
6) Run create_contour_plots.py


## Suggested Fields

These are the fields for each attachment in the correct order. Determined when running the Action 5 appendices to match what was done in the LTO.

### Elevation:
"S_TRNTY","S_SHSTA","S_OROVL","S_FOLSM","S_SLUIS","S_SLUIS_CVP","S_SLUIS_SWP","S_MELON","S_MLRTN"

### Flow:
'C_LWSTN','C_CLR011','C_KSWCK','C_SAC257','C_SAC240','C_SAC201','C_SAC120','C_FTR059','C_FTR003','SP_SAC083_YBP037','C_YBP020','C_NTOMA','C_AMR004','C_SAC048','C_SAC007','C_SJR225','C_SJR180','C_SJR115','C_STS059','C_STS004','C_SJR070','C_OMR014','NDO'

### Diversion:
"D_LWSTN_CCT011","D_SAC240_TCC001","D_SAC207_GCC007","D_NTOMA_FSC003","D_MLRTN_FRK000","D_MLRTN_MDC006","D_SAC030_MOK014","TOTAL_EXP", "C_DMC003","C_CAA003_CVP","C_CAA003_SWP","D_DMC007_CAA009"

### Water Supply (specifically for the DSS reader):
'D_SBP028_17S_PR', 'D_JBC002_17N_PR','D_CRK005_17N_NR', 'D_SAC294_03_PA', 'D_CAA143_90_PA2','D_WTPCSD_02_PA', 'D_DMC034_71_PA2', 'D_XCC025_72_PA','D_WTPBLV_03_PU2', 'D_WTPCSD_02_PU', 'D_WKYTN_02_PU','D_WTPJMS_03_PU1', 'D_FOLSM_WTPSJP_CVP', 'D_NFA016_WTPWAL_CVP','D_WTPJJO_50_PU', 'D_WTPFTH_02_SU', 'D_WTPFTH_03_SU','D_WTPBUK_03_SU', 'D_SAC296_02_SA', 'D_SAC289_03_SA','D_GCC027_08N_SA2', 'D_GCC056_08S_SA2', 'D_SAC178_08N_SA1','D_SAC159_08N_SA1', 'D_SAC159_08S_SA1', 'D_SAC121_08S_SA3','D_SAC109_08S_SA3', 'D_SAC136_18_SA', 'D_SAC122_19_SA','D_SAC115_19_SA', 'D_SAC099_19_SA', 'D_SAC091_19_SA','D_MTC000_09_SA1', 'D_SAC082_22_SA1', 'D_SAC078_22_SA1','D_SAC083_21_SA', 'D_SAC074_21_SA', 'D_MDOTA_73_XA','D_DMC111_73_XA', 'D_MDOTA_64_XA', 'D_XCC010_72_XA2','D_ARY010_72_XA1', 'D_XCC054_72_XA3', 'D_GCC027_08N_PR1','D_MTC000_09_PR', 'D_GCC056_08S_PR', 'D_CBD037_08S_PR','D_GCC039_08N_PR2', 'D_MDOTA_91_PR', 'D_ARY010_72_PR6','D_XCC025_72_PR6', 'D_XCC054_72_PR5', 'D_VLW008_72_PR5','D_ARY010_72_PR3', 'D_XCC033_72_PR4', 'D_ARY010_72_PR4','D_XCC033_72_PR2', 'D_VLW008_72_PR1', 'D_CAA046_71_PA7_PAG','D_THRMF_11_NU1_PMI', 'D_WTPCYC_16_PU', 'D_CSB038_OBISPO_PCO','D_CSB103_BRBRA_PMI', 'D_CSB103_BRBRA_PCO', 'D_CSB103_BRBRA_PIN','D_ESB324_AVEK_PCO', 'D_ESB324_AVEK_PIN', 'D_ESB433_MWDSC_PMI','D_ESB433_MWDSC_PCO', 'D_ESB433_MWDSC_PIN', 'D_THRMA_WEC000','D_THRMA_RVC000', 'D_FTR039_SEC009', 'D_THRMA_JBC000','D_FTR021_16_SA', 'D_FTR014_16_SA', 'D_FTR018_15S_SA','D_FTR018_16_SA', 'D_WTPMNR_13_NU1', 'D_OROVL_13_NU1','D_FPT013_WTPVNY_CVP', 'D_WTPSAC_26S_PU4_CVP','D_FPT013_FSC013_EBMUD', 'D_FPT013_FSC013_CCWD', 'D_CCL005_04_PA1','D_TCC022_04_PA2', 'D_TCC036_07N_PA', 'D_TCC081_07S_PA','D_CBD028_08S_PA', 'D_CBD049_08N_PA', 'D_KLR005_21_PA','DG_08N_PR2', 'DG_12_NU1', 'DG_11_NU1', 'DG_16_PU','D_FTR021_16_PA', 'DG_16_SA', 'DG_15S_SA', 'D_SVRWD_CSTLN_PCO','D_SVRWD_CSTLN_PMI', 'D_PRRIS_MWDSC', 'D_PRRIS_MWDSC_PCO','D_PRRIS_MWDSC_PIN', 'D_PRRIS_MWDSC_PMI', 'D_PYRMD_VNTRA_LCP','D_PYRMD_VNTRA_PCO', 'D_PYRMD_VNTRA_PMI', 'D_CSTIC_VNTRA','D_CSTIC_VNTRA_PCO', 'D_CSTIC_VNTRA_PIN', 'D_CSTIC_VNTRA_PMI','D_BKR004_NBA009_NAPA_PMI', 'D_BKR004_NBA009_NAPA_PCO','D_BKR004_NBA009_NAPA_PIN', 'D_BKR004_NBA009_SCWA_PMI','D_BKR004_NBA009_SCWA_PCO', 'D_BKR004_NBA009_SCWA_PIN','D_CCC019_CCWD', 'DG_13_NU1', 'D_MDOTA_90_PA1', 'D_CAA109_90_PA1','D_CAA143_90_PA1', 'D_CAA155_90_PA1', 'D_CAA172_90_PA1','D_MDOTA_91_PA', 'DG_73_XA', 'DG_64_XA', 'DG_72_XA2', 'DG_72_XA1','DG_91_PR', 'DG_72_PR6', 'DG_72_PR5', 'DG_72_PR3', 'DG_63_PR3','DG_72_PR4', 'D_DMC011_71_PA8', 'D_DMC030_71_PA1','D_DMC034_71_PA3', 'D_DMC044_71_PA4', 'D_DMC044_71_PA5','D_DMC064_71_PA6', 'D_DMC021_50_PA1', 'D_DMC070_73_PA1','D_CAA087_73_PA1', 'D_DMC105_73_PA2', 'D_CAA109_73_PA3','D_DMC091_73_PA3', 'DG_72_XA3', 'DG_72_PR2', 'DG_72_PR1','D_PCH000_SBCWD_PA', 'D_PCH000_SCVWD_PA', 'D_PCH000_SBCWD_PU','D_PCH000_SCVWD_PU', 'DG_11_SA1', 'DG_11_SA3','D_FOLSM_WTPEDH_CVP', 'D_FOLSM_EDCOCA_CVP', 'D_FOLSM_PCWA_CVP','D_FOLSM_WTPFOL_CVP', 'D_FOLSM_WTPRSV_CVP', 'D_CAA046_71_PA7_PIN','D_CAA046_71_PA7_PCO', 'D_SBA009_ACFC_PMI', 'D_SBA009_ACFC_PCO','D_SBA009_ACFC_PIN', 'D_SBA020_ACFC_PMI', 'D_SBA020_ACFC_PCO','D_SBA029_ACWD_PMI', 'D_SBA029_ACWD_PCO', 'D_SBA029_ACWD_PIN','D_SBA036_SCVWD_PMI', 'D_SBA036_SCVWD_PCO', 'D_SBA036_SCVWD_PIN','D_CAA143_90_PU', 'D_CAA156_90_PU', 'D_CAA165_90_PU','D_CAA173_EMPIRE_PAG', 'D_CAA173_EMPIRE_PCO','D_CAA173_EMPIRE_PIN', 'D_CAA181_KINGS_PAG', 'D_CAA181_KINGS_PCO','D_CAA181_KINGS_PIN', 'D_CAA183_TULARE_PAG', 'D_CAA183_TULARE_PCO','D_CAA183_TULARE_PIN', 'D_CAA184_DUDLEY_PAG','D_CAA184_DUDLEY_PCO', 'D_CAA184_DUDLEY_PIN', 'D_CAA194_KERN_PAG','D_CAA194_KERN_PCO', 'D_CAA194_KERNA_PMI', 'D_CAA194_KERNB_PMI','D_CAA238_CVPCV', 'D_CAA239_CVPRF', 'D_CAA242_KERN_PAG','D_CAA242_KERN_PCO', 'D_CAA242_KERN_PIN', 'D_CAA279_KERN','D_CSB015_KERN_BMWD_PAG', 'D_CSB015_KERN_BMWD_PCO','D_CSB009_CLRTA1_DDWD_PAG', 'D_CSB009_CLRTA1_DDWD_PCO','D_CSB009_CLRTA1_DDWD_PIN', 'D_CSB038_OBISPO_PCO','D_CSB038_OBISPO_PIN', 'D_CSB038_OBISPO_PMI', 'D_CSB103_BRBRA_PCO','D_CSB103_BRBRA_PIN', 'D_CSB103_BRBRA_PMI', 'D_ESB324_AVEK_PCO','D_ESB324_AVEK_PIN', 'D_ESB324_AVEK_PMI', 'D_ESB347_PLMDL_PCO','D_ESB347_PLMDL_PIN', 'D_ESB347_PLMDL_PMI', 'D_ESB355_LROCK_PCO','D_ESB355_LROCK_PMI', 'D_ESB403_MOJVE_PCO', 'D_ESB403_MOJVE_PMI','D_ESB407_CCHLA_PCO', 'D_ESB407_CCHLA_PIN', 'D_ESB407_CCHLA_PMI','D_ESB408_DESRT_PCO', 'D_ESB408_DESRT_PIN', 'D_ESB408_DESRT_PMI','D_ESB413_MWDSC_LCP', 'D_ESB413_MWDSC_PCO', 'D_ESB413_MWDSC_PIN','D_ESB413_MWDSC_PMI', 'D_ESB414_BRDNO_LCP', 'D_ESB414_BRDNO_PCO','D_ESB414_BRDNO_PIN', 'D_ESB414_BRDNO_PMI', 'D_ESB415_GABRL_LCP','D_ESB415_GABRL_PCO', 'D_ESB415_GABRL_PIN', 'D_ESB415_GABRL_PMI','D_ESB420_GRGNO_LCP', 'D_ESB420_GRGNO_PCO', 'D_ESB420_GRGNO_PIN','D_ESB420_GRGNO_PMI', 'D_WSB031_MWDSC_LCP', 'D_WSB031_MWDSC_PCO','D_WSB031_MWDSC_PIN', 'D_WSB031_MWDSC_PMI', 'D_WSB032_CLRTA2_LCP','D_WSB032_CLRTA2_PCO', 'D_WSB032_CLRTA2_PMI', 'D_FSC015_60N_NA2','D_FSC025_60N_PU_CVP', 'PERDV_CVPAG_S', 'PERDV_CVPAG_SYS','PERDV_CVPEX_S', 'PERDV_CVPMI_S', 'PERDV_CVPMI_SYS','PERDV_CVPRF_SYS', 'PERDV_CVPSC_SYS', 'PERDV_SWP_AG1','PERDV_SWP_FSC', 'PERDV_SWP_MWD1', 'SWP_CO_TOTAL', 'SWP_IN_TOTAL','SWP_TA_FEATH', 'SWP_TA_NBAY', 'SWP_TA_TOTAL', 'TOTAL_EXP','D_MLRTN_FRK_C1', 'D_MLRTN_FRK_C2', 'D_MLRTN_MDC_C1','D_MLRTN_MDC_C2', 'D_TCC111_07S_PA', 'D_SAC162_09_SA2','DEL_CVP_PSC_N', 'DEL_CVP_PRF_N', 'DEL_CVP_PMI_N_WAMER','DEL_CVP_PAG_N', 'SWP_FRSA', 'DEL_CVP_PEX_S','DEL_CVP_PRF_S','DEL_CVP_PMI_S','DEL_CVP_PAG_S',
### Temperature:
"BLW LEWISTON","WHISKEYTOWN","IGO","ABV SACRAMENTO","BLW KESWICK","BLW CLEAR CREEK","BALLS FERRY","JELLYS FERRY","BEND BRIDGE","RED_BLUFF","RED BLUFF DAM","HAMILTON CITY","BLW NIMBUS(HAZEL AVE)","WATT AVE","ABV CONFLUENCE"

### EC:
"SAC_DS_STMBTSL","CACHE_RYER","RSAC123","RSAC092","RSAC101","RSAN112","RSAN112","RSAN018","ROLD024","RSAN007","RSAC075","RSAC081","CHIPS_N_437","CHIPS_S_442","RSAC064","CHDMC006","CLIFTONCOURT","ROLD034","CHVCT000"

### Position:
'X2'

### Cl:
'ROLD024','RSAN007','CLIFTONCOURT','CHDMC006','SLBAR002'

### Water Temperature Contour Plots (for the DSS reader):
"BLW SHASTA", 'BLW KESWICK','HWY44','BLW CLEAR CREEK','AIRPORT'

## Support
Please contact rlucas@usbr.gov for support.
