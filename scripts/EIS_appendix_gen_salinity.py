from EISAppendixGen_functions import create_appendix

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    # Fields to use from DSS Reader
    # EC
    fields = ["SAC_DS_STMBTSL", "CACHE_RYER", "RSAC123", "RSAC092", "RSAC101", "RSAN112", "RSAN112", "RSAN018", "ROLD024", "RSAN007",
              "RSAC075", "RSAC081", "CHIPS_N_437", "CHIPS_S_442", "RSAC064", "CHDMC006", "CLIFTONCOURT", "ROLD034", "CHVCT000"]
    # Cl
    # fields = ['ROLD024', 'RSAN007', 'CLIFTONCOURT', 'CHDMC006', 'SLBAR002']
    # Position
    # fields = ['X2']

    # alternatives to include
    alts = ['NAA', "Action 5"]

    """
    Specify whether report is "EC", "Cl", or "Position" 

    Note 1: "Position" is the X2 position.
    Note 2: Conversion from microSiemens/cm to mg/L Cl uses equation 2 of https://www.waterboards.ca.gov/waterrights/water_issues/programs/bay_delta/california_waterfix/exhibits/docs/ccc_cccwa/CCC-SC_25.pdf

    """
    report_type = "EC"

    # For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    # Select whether to use the calendar year to group data.
    use_calendar_yr = True  # Note: For Trinity LTO tables/figures, use False.

    # Prefix for tables and figures in appendix
    appendix_prefix = " F.2.5"  # F.2.5 is DSM2-EC ; F.2.6 is DSM2-X2 (position); F.2.7 is DSM2 - Chloride

    # Path to file with location code crosswalk
    location_cw_path = r"..\inputs\location_code_crosswalk_salinity.xlsx"

    # Path to file with DSSReader output
    # for water supply, must be the _TAF output
    # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
    dss_path = r""

    # Path to file with WY Typing data
    wy_flags_path = "..\inputs\wy_flags.xlsx"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc - update parent directory
    template = r"..\inputs\template_v2-fonts.docx"
    doc_name = r"appendix_temp.docx"
    # Name of final word doc
    new_doc = rf"appendix_final_{report_type}.docx"

    ####END OF USER INPUTS #######

    create_appendix(report_type, alts, fields, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, location_cw_path, use_calendar_yr=use_calendar_yr,
                        use_lumped_table_captions=use_lumped_table_captions)
