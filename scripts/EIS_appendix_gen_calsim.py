from EISAppendixGen_functions import create_appendix
import os

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    # Fields to use from DSS Reader

    # Use for running "elevations" report type.
    # fields = ["S_TRNTY","S_SHSTA","S_OROVL","S_FOLSM","S_SLUIS","S_SLUIS_CVP","S_SLUIS_SWP","S_MELON","S_MLRTN"]

    # Use for running "flow" report type
    fields = ['C_LWSTN','C_CLR011','C_KSWCK','C_SAC257','C_SAC240','C_SAC201','C_SAC120','C_FTR059','C_FTR003','SP_SAC083_YBP037', 'C_YBP020',
              'C_NTOMA','C_AMR004','C_SAC048','C_SAC007','C_SJR225','C_SJR180','C_SJR115','C_STS059','C_STS004','C_SJR070','C_OMR014','NDO']

    # Used for running "diversions" report type
    # fields = ["D_LWSTN_CCT011","D_SAC240_TCC001","D_SAC207_GCC007","D_NTOMA_FSC003","D_MLRTN_FRK000","D_MLRTN_MDC006",
    # "D_SAC030_MOK014","TOTAL_EXP", "C_DMC003","C_CAA003_CVP","C_CAA003_SWP","D_DMC007_CAA009"]

    # alternatives to include
    alts = ['NAA', 'Alt2v2_woTUCP']

    """
    Specify whether report is "flow", "elevation", or "diversion"

    Note 1: "elevation" option also includes storages.
    """
    report_type = "flow"

    # For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    # Select whether to use the calendar year to group data.
    use_calendar_yr = True  # Note: For Trinity LTO tables/figures, use False.

    # Prefix for tables and figures in appendix
    appendix_prefix = " F.2.2"  # F.2.1 is elevation; F.2.2 is flow; F.2.3 is diversion

    # Define base working directory for reference
    base_dir = r"C:\20251211_BA_Modeling_Appendix\EIS_Appendix_Generation\EIS-Appendix-Generation-main"

    # Change directory to scripts: SN 20260303
    os.chdir(os.path.join(base_dir, "scripts"))

    # Path to file with location code crosswalk
    location_cw_path = os.path.join(base_dir, r"inputs\location_code_crosswalk_CalSim.xlsx")

    # Path to file with DSSReader output
    # for water supply, must be the _TAF output
    # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
    dss_path =  os.path.join(base_dir, r"inputs\DSS_contents.xlsx")

    # Path to file with WY Typing data
    wy_flags_path = os.path.join(base_dir, r"inputs\wy_flags.xlsx")

    # Path to storage-elevation table data
    storage_elevation_table = os.path.join(base_dir, r"inputs\storage_elevation_table.xlsx")

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc - update parent directory
    template = os.path.join(base_dir, r"inputs\template_v2-fonts.docx")
    doc_name = r"appendix_temp.docx"
    # Name of final word doc
    new_doc = rf"appendix_final_{report_type}.docx"

    ####END OF USER INPUTS #######

    # call the corresponding function for the appendix
    create_appendix(report_type, alts, fields, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, location_cw_path, use_calendar_yr=use_calendar_yr,
                        use_lumped_table_captions=use_lumped_table_captions, storage_elevation_table=storage_elevation_table)
