from EISAppendixGen_functions import create_water_supply_appendix, create_appendix

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    #Fields to use from DSS Reader

    # Use for running "elevations" report type, in desired order.
    # fields = ["S_TRNTY","S_SHSTA","S_OROVL","S_FOLSM","S_SLUIS","S_SLUIS_CVP","S_SLUIS_SWP","S_MELON","S_MLRTN"]

    # #Use for running "flow" report type
    fields = ["C_LWSTN", "C_CLR011", "C_KSWCK", "C_SAC257", "C_SAC240", "C_SAC201",
                            "C_SAC120", "C_FTR059", "C_FTR003", "SP_SAC083_YBP037", "C_YBP020",
                            "C_NTOMA", "C_AMR004", "C_SAC048", "C_SAC007", "C_SJR225", "C_SJR180",
                            "C_SJR115", "C_STS059", "C_STS004", "C_SJR070", "C_OMR014", "NDO"]

    # #Used for running "diversions" report type
    # fields = [ "D_LWSTN_CCT011","D_SAC240_TCC001"
    # ,"D_SAC207_GCC007","D_NTOMA_FSC003","D_MLRTN_FRK000","D_MLRTN_MDC006",
    # "D_SAC030_MOK014","TOTAL_EXP", "C_DMC003","C_CAA003_CVP","C_CAA003_SWP","D_DMC007_CAA009"]

    #Temperature
    # alts = ['NAA', "Action 5"]
    # fields = [
    #     "AIRPORT",  # Compliance location (most downstream) - Sac Rv along Airport Rd
    #     "BLW CLEAR CREEK",  # Compliance location (middle) - Sac River below Clear Creek
    #     "HWY44",  # Compliance location (most upstream) - Sac River at HWY 44
    #
    #     # Other locations (Not compliance locations, but still included in documentation).
    #     "BLW LEWISTON",
    #     "WHISKEYTOWN",
    #     "IGO",
    #     "ABV SACRAMENTO",
    #     'BLW SHASTA',
    #     "BLW KESWICK",
    #     "BALLS FERRY",
    #     "JELLYS FERRY",
    #     "BEND BRIDGE",
    #     "RED_BLUFF",
    #     "RED BLUFF DAM",
    #     "HAMILTON CITY",
    #     "BLW NIMBUS(HAZEL AVE)",
    #     "WATT AVE",
    #     "ABV CONFLUENCE"
    # ]

    # Water supply fields, order doesn't matter
    alts = ['NAA', "ALT2v1"]

    # Compliance fields for the mixed compliance location logic used in Alt2v3 and Action 5.
    compliance_fields = ['AIRPORT', 'BLW CLEAR CREEK', 'HWY44']

    ##### Mixed Compliance Location Logic #####
    # Shastabin_ == 1 or 2 means compliance location is at Sac Rv at AIRPORT RD. (Most downstream location)
    # Shastabin_ == 3 or 4 means compliance location is  Sac Rv blw Clear Creek.
    # Shastabin_ == 5 or 6 means compliance location is at Sac Rv at HWY 44. (Most upstream location)
    compliance_dict = {
        1: 'AIRPORT',
        2: 'AIRPORT',
        3: 'BLW CLEAR CREEK',
        4: 'BLW CLEAR CREEK',
        5: 'HWY44',
        6: 'HWY44',
    }
    # fields = ['D_SBP028_17S_PR', 'D_JBC002_17N_PR',
    #    'D_CRK005_17N_NR', 'D_SAC294_03_PA', 'D_CAA143_90_PA2',
    #    'D_WTPCSD_02_PA', 'D_DMC034_71_PA2', 'D_XCC025_72_PA',
    #    'D_WTPBLV_03_PU2', 'D_WTPCSD_02_PU', 'D_WKYTN_02_PU',
    #    'D_WTPJMS_03_PU1', 'D_FOLSM_WTPSJP_CVP', 'D_NFA016_WTPWAL_CVP',
    #    'D_WTPJJO_50_PU', 'D_WTPFTH_02_SU', 'D_WTPFTH_03_SU',
    #    'D_WTPBUK_03_SU', 'D_SAC296_02_SA', 'D_SAC289_03_SA',
    #    'D_GCC027_08N_SA2', 'D_GCC056_08S_SA2', 'D_SAC178_08N_SA1',
    #    'D_SAC159_08N_SA1', 'D_SAC159_08S_SA1', 'D_SAC121_08S_SA3',
    #    'D_SAC109_08S_SA3', 'D_SAC136_18_SA', 'D_SAC122_19_SA',
    #    'D_SAC115_19_SA', 'D_SAC099_19_SA', 'D_SAC091_19_SA',
    #    'D_MTC000_09_SA1', 'D_SAC082_22_SA1', 'D_SAC078_22_SA1',
    #    'D_SAC083_21_SA', 'D_SAC074_21_SA', 'D_MDOTA_73_XA',
    #    'D_DMC111_73_XA', 'D_MDOTA_64_XA', 'D_XCC010_72_XA2',
    #    'D_ARY010_72_XA1', 'D_XCC054_72_XA3', 'D_GCC027_08N_PR1',
    #    'D_MTC000_09_PR', 'D_GCC056_08S_PR', 'D_CBD037_08S_PR',
    #    'D_GCC039_08N_PR2', 'D_MDOTA_91_PR', 'D_ARY010_72_PR6',
    #    'D_XCC025_72_PR6', 'D_XCC054_72_PR5', 'D_VLW008_72_PR5',
    #    'D_ARY010_72_PR3', 'D_XCC033_72_PR4', 'D_ARY010_72_PR4',
    #    'D_XCC033_72_PR2', 'D_VLW008_72_PR1', 'D_CAA046_71_PA7_PAG',
    #    'D_THRMF_11_NU1_PMI', 'D_WTPCYC_16_PU', 'D_CSB038_OBISPO_PCO',
    #    'D_CSB103_BRBRA_PMI', 'D_CSB103_BRBRA_PCO', 'D_CSB103_BRBRA_PIN',
    #    'D_ESB324_AVEK_PCO', 'D_ESB324_AVEK_PIN', 'D_ESB433_MWDSC_PMI',
    #    'D_ESB433_MWDSC_PCO', 'D_ESB433_MWDSC_PIN', 'D_THRMA_WEC000',
    #    'D_THRMA_RVC000', 'D_FTR039_SEC009', 'D_THRMA_JBC000',
    #    'D_FTR021_16_SA', 'D_FTR014_16_SA', 'D_FTR018_15S_SA',
    #    'D_FTR018_16_SA', 'D_WTPMNR_13_NU1', 'D_OROVL_13_NU1',
    #    'D_FPT013_WTPVNY_CVP', 'D_WTPSAC_26S_PU4_CVP',
    #    'D_FPT013_FSC013_EBMUD', 'D_FPT013_FSC013_CCWD', 'D_CCL005_04_PA1',
    #    'D_TCC022_04_PA2', 'D_TCC036_07N_PA', 'D_TCC081_07S_PA',
    #    'D_CBD028_08S_PA', 'D_CBD049_08N_PA', 'D_KLR005_21_PA',
    #    'DG_08N_PR2', 'DG_12_NU1', 'DG_11_NU1', 'DG_16_PU',
    #    'D_FTR021_16_PA', 'DG_16_SA', 'DG_15S_SA', 'D_SVRWD_CSTLN_PCO',
    #    'D_SVRWD_CSTLN_PMI', 'D_PRRIS_MWDSC', 'D_PRRIS_MWDSC_PCO',
    #    'D_PRRIS_MWDSC_PIN', 'D_PRRIS_MWDSC_PMI', 'D_PYRMD_VNTRA_LCP',
    #    'D_PYRMD_VNTRA_PCO', 'D_PYRMD_VNTRA_PMI', 'D_CSTIC_VNTRA',
    #    'D_CSTIC_VNTRA_PCO', 'D_CSTIC_VNTRA_PIN', 'D_CSTIC_VNTRA_PMI',
    #    'D_BKR004_NBA009_NAPA_PMI', 'D_BKR004_NBA009_NAPA_PCO',
    #    'D_BKR004_NBA009_NAPA_PIN', 'D_BKR004_NBA009_SCWA_PMI',
    #    'D_BKR004_NBA009_SCWA_PCO', 'D_BKR004_NBA009_SCWA_PIN',
    #    'D_CCC019_CCWD', 'DG_13_NU1', 'D_MDOTA_90_PA1', 'D_CAA109_90_PA1',
    #    'D_CAA143_90_PA1', 'D_CAA155_90_PA1', 'D_CAA172_90_PA1',
    #    'D_MDOTA_91_PA', 'DG_73_XA', 'DG_64_XA', 'DG_72_XA2', 'DG_72_XA1',
    #    'DG_91_PR', 'DG_72_PR6', 'DG_72_PR5', 'DG_72_PR3', 'DG_63_PR3',
    #    'DG_72_PR4', 'D_DMC011_71_PA8', 'D_DMC030_71_PA1',
    #    'D_DMC034_71_PA3', 'D_DMC044_71_PA4', 'D_DMC044_71_PA5',
    #    'D_DMC064_71_PA6', 'D_DMC021_50_PA1', 'D_DMC070_73_PA1',
    #    'D_CAA087_73_PA1', 'D_DMC105_73_PA2', 'D_CAA109_73_PA3',
    #    'D_DMC091_73_PA3', 'DG_72_XA3', 'DG_72_PR2', 'DG_72_PR1',
    #    'D_PCH000_SBCWD_PA', 'D_PCH000_SCVWD_PA', 'D_PCH000_SBCWD_PU',
    #    'D_PCH000_SCVWD_PU', 'DG_11_SA1', 'DG_11_SA3',
    #    'D_FOLSM_WTPEDH_CVP', 'D_FOLSM_EDCOCA_CVP', 'D_FOLSM_PCWA_CVP',
    #    'D_FOLSM_WTPFOL_CVP', 'D_FOLSM_WTPRSV_CVP', 'D_CAA046_71_PA7_PIN',
    #    'D_CAA046_71_PA7_PCO', 'D_SBA009_ACFC_PMI', 'D_SBA009_ACFC_PCO',
    #    'D_SBA009_ACFC_PIN', 'D_SBA020_ACFC_PMI', 'D_SBA020_ACFC_PCO',
    #    'D_SBA029_ACWD_PMI', 'D_SBA029_ACWD_PCO', 'D_SBA029_ACWD_PIN',
    #    'D_SBA036_SCVWD_PMI', 'D_SBA036_SCVWD_PCO', 'D_SBA036_SCVWD_PIN',
    #    'D_CAA143_90_PU', 'D_CAA156_90_PU', 'D_CAA165_90_PU',
    #    'D_CAA173_EMPIRE_PAG', 'D_CAA173_EMPIRE_PCO',
    #    'D_CAA173_EMPIRE_PIN', 'D_CAA181_KINGS_PAG', 'D_CAA181_KINGS_PCO',
    #    'D_CAA181_KINGS_PIN', 'D_CAA183_TULARE_PAG', 'D_CAA183_TULARE_PCO',
    #    'D_CAA183_TULARE_PIN', 'D_CAA184_DUDLEY_PAG',
    #    'D_CAA184_DUDLEY_PCO', 'D_CAA184_DUDLEY_PIN', 'D_CAA194_KERN_PAG',
    #    'D_CAA194_KERN_PCO', 'D_CAA194_KERNA_PMI', 'D_CAA194_KERNB_PMI',
    #    'D_CAA238_CVPCV', 'D_CAA239_CVPRF', 'D_CAA242_KERN_PAG',
    #    'D_CAA242_KERN_PCO', 'D_CAA242_KERN_PIN', 'D_CAA279_KERN',
    #    'D_CSB015_KERN_BMWD_PAG', 'D_CSB015_KERN_BMWD_PCO',
    #    'D_CSB009_CLRTA1_DDWD_PAG', 'D_CSB009_CLRTA1_DDWD_PCO',
    #    'D_CSB009_CLRTA1_DDWD_PIN', 'D_CSB038_OBISPO_PCO',
    #    'D_CSB038_OBISPO_PIN', 'D_CSB038_OBISPO_PMI', 'D_CSB103_BRBRA_PCO',
    #    'D_CSB103_BRBRA_PIN', 'D_CSB103_BRBRA_PMI', 'D_ESB324_AVEK_PCO',
    #    'D_ESB324_AVEK_PIN', 'D_ESB324_AVEK_PMI', 'D_ESB347_PLMDL_PCO',
    #    'D_ESB347_PLMDL_PIN', 'D_ESB347_PLMDL_PMI', 'D_ESB355_LROCK_PCO',
    #    'D_ESB355_LROCK_PMI', 'D_ESB403_MOJVE_PCO', 'D_ESB403_MOJVE_PMI',
    #    'D_ESB407_CCHLA_PCO', 'D_ESB407_CCHLA_PIN', 'D_ESB407_CCHLA_PMI',
    #    'D_ESB408_DESRT_PCO', 'D_ESB408_DESRT_PIN', 'D_ESB408_DESRT_PMI',
    #    'D_ESB413_MWDSC_LCP', 'D_ESB413_MWDSC_PCO', 'D_ESB413_MWDSC_PIN',
    #    'D_ESB413_MWDSC_PMI', 'D_ESB414_BRDNO_LCP', 'D_ESB414_BRDNO_PCO',
    #    'D_ESB414_BRDNO_PIN', 'D_ESB414_BRDNO_PMI', 'D_ESB415_GABRL_LCP',
    #    'D_ESB415_GABRL_PCO', 'D_ESB415_GABRL_PIN', 'D_ESB415_GABRL_PMI',
    #    'D_ESB420_GRGNO_LCP', 'D_ESB420_GRGNO_PCO', 'D_ESB420_GRGNO_PIN',
    #    'D_ESB420_GRGNO_PMI', 'D_WSB031_MWDSC_LCP', 'D_WSB031_MWDSC_PCO',
    #    'D_WSB031_MWDSC_PIN', 'D_WSB031_MWDSC_PMI', 'D_WSB032_CLRTA2_LCP',
    #    'D_WSB032_CLRTA2_PCO', 'D_WSB032_CLRTA2_PMI', 'D_FSC015_60N_NA2',
    #    'D_FSC025_60N_PU_CVP', 'PERDV_CVPAG_S', 'PERDV_CVPAG_SYS',
    #    'PERDV_CVPEX_S', 'PERDV_CVPMI_S', 'PERDV_CVPMI_SYS',
    #    'PERDV_CVPRF_SYS', 'PERDV_CVPSC_SYS', 'PERDV_SWP_AG1',
    #    'PERDV_SWP_FSC', 'PERDV_SWP_MWD1', 'SWP_CO_TOTAL', 'SWP_IN_TOTAL',
    #    'SWP_TA_FEATH', 'SWP_TA_NBAY', 'SWP_TA_TOTAL', 'TOTAL_EXP',
    #    'D_MLRTN_FRK_C1', 'D_MLRTN_FRK_C2', 'D_MLRTN_MDC_C1',
    #    'D_MLRTN_MDC_C2', 'D_TCC111_07S_PA', 'D_SAC162_09_SA2']

    #alts = ['NAA', 'Alternative 1', 'Alternative 2a', 'Alternative 2b', 'Alternative 3', 'Alternative 4', 'Alternative 6', 'Alternative 7']

    # Scenarios to compare
    #alts = ["NAA", "Alt1", "Alt2a", 'Alt2b', 'Alt3', 'Alt4', 'Alt6', 'Alt7']
    #alts = ['NAA', "SACLTO_Alt4"]
    #Temperature test
    #fields = ["BLW CLEAR CREEK"]
    #alts = ["NAA", "NAA"]

    #Salinity Test
    #fields = ["SAC_DS_STMBTSL","RSAN007","RSAC075", "RSAC081"] #Test fields for EC DSM2 appendix
    #fields = ['ROLD024','RSAN007'] #Test fields for Cl DSM2 appendix
    # fields = ['X2']
    #alts = ["NAA", "ALT1"]
    """
    Specify whether report is "flow", "elevation', "diversion", or "water supply" (CalSim appendices), "temperature" (HEC-5Q appendix), 
    "EC", "Cl", "Position" (salinity/DSM2 appendices). 
    
    Note 1: "elevation" option also includes storages. "Position" is the X2 position.
    Note 2: Conversion from microSiemens/cm to mg/L Cl uses equation 2 of https://www.waterboards.ca.gov/waterrights/water_issues/programs/bay_delta/california_waterfix/exhibits/docs/ccc_cccwa/CCC-SC_25.pdf
    
    """
    report_type = "flow"

    #For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    #Select whether to use the calendar year to group data.
    use_calendar_yr = True  #Note: For Trinity LTO tables/figures, use False.

    #TODO Add selection of units (elevation, temperature, provide both cfs and taf?)
    #Get more info from crosswalk, units, description, etc
    #Add salinity and temperature - could break into separate scripts

    # Prefix for tables and figures in appendix
    appendix_prefix = " F.2.2" #F.2.1 is elevation; F.2.2 is flow; F.2.3 is diversion; F.2.4 is water supply
                                #F.2.5 is DSM2-EC ; F.2.6 is DSM2-X2 (position); F.2.7 is DSM2 - Chloride; F.2.8 is DSM2

    # Path to file with location code crosswalk
    if report_type in ["flow", "elevation", "diversion", "water supply"]:
        #CalSim related appendices use the calsim crosswalk
        location_cw_path = r"..\inputs\location_code_crosswalk_CalSim.xlsx"
    elif report_type in ["EC", "Cl", "Position"]:
        #DSM2 related reports use the salinity crosswalk
        location_cw_path = r"..\inputs\location_code_crosswalk_salinity.xlsx"
    elif report_type == 'temperature':
        #Temperature related appendices use the temperature crosswalk
        location_cw_path = r"..\inputs\location_code_crosswalk_Temp.xlsx"

    s_supply_formulas = r"..\inputs\water_supply_formulas.xlsx"

    #Path to file with DSSReader output
    # for water supply, must be the _TAF output
    #Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
    dss_path = r"C:\Users\fnufferrodriguez\OneDrive - DOI\Desktop\calsim_dss_reader\DSS_contents.xlsx"

    #File containing shasta bin information (By calendar yr) for each of the alternatives
    shastabin_data_path  = r"..\inputs\shasta_bin_info.xlsx"

    #Path to file with WY Typing data
    wy_flags_path = "..\inputs\wy_flags.xlsx"

    #Path to storage-elevation table data
    storage_elevation_table = r"..\inputs\storage_elevation_table.xlsx"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    #Name of intermediate word doc - update parent directory
    template = r"..\inputs\template_v2-fonts.docx"
    doc_name = r"..\appendix_temp.docx"
    #Name of final word doc
    new_doc = rf"..\appendix_final_{report_type}.docx"

####END OF USER INPUTS #######

    # call the corresponding function for the appendix
    if report_type == 'water supply':
        create_water_supply_appendix(alts, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, s_supply_formulas)
    else:
        create_appendix(report_type, alts, fields, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, location_cw_path, use_calendar_yr=use_calendar_yr, use_lumped_table_captions=use_lumped_table_captions,
                            storage_elevation_table=storage_elevation_table, compliance_fields=compliance_fields, compliance_dict=compliance_dict, shastabin_data_path=shastabin_data_path)
