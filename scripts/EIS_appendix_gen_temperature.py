from EISAppendixGen_functions import create_appendix

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    # Fields to use from DSS Reader
    fields = [
        "AIRPORT",  # Compliance location (most downstream) - Sac Rv along Airport Rd
        "BLW CLEAR CREEK",  # Compliance location (middle) - Sac River below Clear Creek
        "HWY44",  # Compliance location (most upstream) - Sac River at HWY 44

        # Other locations (Not compliance locations, but still included in documentation).
        "BLW LEWISTON",
        "WHISKEYTOWN",
        "IGO",
        "ABV SACRAMENTO",
        'BLW SHASTA',
        "BLW KESWICK",
        "BALLS FERRY",
        "JELLYS FERRY",
        "BEND BRIDGE",
        "RED_BLUFF",
        "RED BLUFF DAM",
        "HAMILTON CITY",
        "BLW NIMBUS(HAZEL AVE)",
        "WATT AVE",
        "ABV CONFLUENCE"
    ]

    # alternatives to include
    alts = ['NAA', "Action 5"]

    # Compliance fields for the mixed compliance location.
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

    # For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    # Select whether to use the calendar year to group data.
    use_calendar_yr = True  # Note: For Trinity LTO tables/figures, use False.

    # Prefix for tables and figures in appendix
    # ' F.2.11' is the common one for temperature
    appendix_prefix = " F.2.11"

    # Path to file with location code crosswalk
    location_cw_path = r"..\inputs\location_code_crosswalk_Temp.xlsx"

    s_supply_formulas = r"..\inputs\water_supply_formulas.xlsx"

    # Path to file with DSSReader output
    # for water supply, must be the _TAF output
    # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
    dss_path = r"C:\Users\fnufferrodriguez\OneDrive - DOI\Desktop\calsim_dss_reader\DSS_contents_temp_test.xlsx"

    # File containing shasta bin information (By calendar yr) for each of the alternatives
    shastabin_data_path = r"..\inputs\shasta_bin_info.xlsx"

    # Path to file with WY Typing data
    wy_flags_path = "..\inputs\wy_flags.xlsx"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc - update parent directory
    template = r"..\inputs\template_v2-fonts.docx"
    doc_name = r"..\appendix_temp.docx"
    # Name of final word doc
    new_doc = rf"..\appendix_final_temperature.docx"

    ####END OF USER INPUTS #######

    # call the corresponding function for the appendix
    create_appendix('temperature', alts, fields, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, location_cw_path, use_calendar_yr=use_calendar_yr,
                    use_lumped_table_captions=use_lumped_table_captions, compliance_fields=compliance_fields, compliance_dict=compliance_dict, shastabin_data_path=shastabin_data_path)
