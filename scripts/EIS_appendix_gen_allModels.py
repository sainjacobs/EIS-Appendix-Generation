"""
Universal appendix-generation driver.

EISAppendixGen_functions.py already implements one engine (`create_appendix`)
that is shared by the CalSim ("flow"/"elevation"/"diversion"), DSM2 salinity
("EC"/"Cl"/"Position"), and temperature ("temperature") appendices - the only
difference between those three drivers was which input values they set. The
water supply and water quality compliance appendices call their own engine
functions (`create_water_supply_appendix` / `create_compliance_appendix`)
with a different set of inputs.

This driver merges all five of those single-purpose scripts
(EIS_appendix_gen_calsim.py, EIS_appendix_gen_salinity.py,
EIS_appendix_gen_temperature.py, EIS_appendix_gen_water_supply.py,
EIS_appendix_gen_compliance.py) into one file. Set REPORT_TYPE below to pick
which appendix to build, edit the matching configuration block, and run the
script - the dispatch logic at the bottom calls the right engine function
automatically.
"""

from EISAppendixGen_functions import (
    create_appendix,
    create_water_supply_appendix,
    create_compliance_appendix,
)
import os

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    # Choose which appendix to generate. Must be one of:
    #   "flow", "elevation", "diversion"   -> CalSim appendices
    #   "EC", "Cl", "Position"             -> DSM2 salinity appendices
    #   "temperature"                      -> HEC-5Q temperature appendix
    #   "water_supply"                     -> Water supply appendix
    #   "compliance"                       -> Water quality compliance appendix
    REPORT_TYPE = "flow"

    CALSIM_REPORT_TYPES = {"flow", "elevation", "diversion"}
    SALINITY_REPORT_TYPES = {"EC", "Cl", "Position"}

    # Define base working directory for reference
    base_dir = r"C:\Github\EIS-Appendix-Generation"

    # Change directory to scripts so relative paths behave the same regardless
    # of where this script is launched from.
    os.chdir(os.path.join(base_dir, "scripts"))

    # Output directory for generated Word docs and plot folders.
    output_folder = r"C:\\20251211_BA_Modeling_Appendix\\outputs_gitRepo\\trial"

    # Path to the template doc, shared by every appendix type.
    template = os.path.join(base_dir, r"inputs\template_v2-fonts.docx")

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc
    doc_name = os.path.join(output_folder, "appendix_temp.docx")
    # Name of final word doc
    new_doc = os.path.join(output_folder, f"appendix_final_{REPORT_TYPE}.docx")

    # Formatting for pages containing tables. Measurements ending in "_pt"
    # are in points; row_height_cm is in centimeters. Only used by the
    # CalSim/salinity/temperature appendices (create_appendix).
    table_page_format = {
        "appendix_heading_font_size": 21,
        "location_heading_font_size": 16,
        "table_font_size": 8,
        "caption_font_size": 10,
        "footnote_font_size": 8,
        "row_height_cm": 0.42,
        "cell_space_before_pt": 1,
        "cell_space_after_pt": 1,
        "caption_space_before_pt": 4,
        "caption_space_after_pt": 2,
        "footnote_space_before_pt": 2,
        "footnote_space_after_pt": 2,
    }

    # Formatting for the Word pages containing plots.
    plot_page_format = {
        "caption_font_size": 12,
        "caption_space_before_pt": 1,
        "caption_space_after_pt": 1,
        "footnote_font_size": 9,
        "footnote_space_before_pt": 1,
        "footnote_space_after_pt": 1,
        "top_blank_lines": 2,
    }

    # Formatting inside the generated plots. Matplotlib color names, hex color
    # codes, and line-style strings are accepted. Supply at least one color and
    # line style for every alternative included below.
    plot_format = {
        "line_colors": ["k", "b", "m", "orange", "y", "r", "purple", "g", "c"],
        "line_styles": ["-", "-.", "--", "-.", "-.", "--", "-.", "-.", ":"],
        "line_width": 1.5,
        "figure_size": (10, 5),
        "figure_border_width": 3,
        "figure_border_color": "black",
        "axis_label_font_size": 10,
        "tick_label_font_size": 10,
        "legend_font_size": 10,
        "legend_columns": 4,
        "compliance_legend_columns": 3,
        "legend_location": "center",
        "legend_y": 1.08,
        "legend_frame": False,
        "grid_color": "gray",
        "grid_style": "--",
        "grid_line_width": 0.8,
        "compliance_marker": "o",
        "compliance_marker_size": 3,
        "save_dpi": 300,
    }

    # Select whether to use the calendar year to group data.
    use_calendar_yr = True  # Note: For Trinity LTO tables/figures, use False.

    # For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    if REPORT_TYPE in CALSIM_REPORT_TYPES:
        ##### CalSim (flow / elevation / diversion) inputs #####

        # Fields to use from DSS Reader, one list per report type.
        fields_by_report_type = {
            # Use for running "elevation" report type.
            "elevation": ["S_TRNTY", "S_SHSTA", "S_OROVL", "S_FOLSM", "S_SLUIS", "S_SLUIS_CVP", "S_SLUIS_SWP", "S_MELON", "S_MLRTN"],
            # Use for running "flow" report type.
            "flow": ['C_LWSTN', 'C_CLR011', 'C_KSWCK', 'C_SAC257', 'C_SAC240', 'C_SAC201', 'C_SAC120', 'C_FTR059', 'C_FTR003', 'SP_SAC083_YBP037', 'C_YBP020',
                     'C_NTOMA', 'C_AMR004', 'C_SAC048', 'C_SAC007', 'C_SJR225', 'C_SJR180', 'C_SJR115', 'C_STS059', 'C_STS004', 'C_SJR070', 'C_OMR014', 'NDO'],
            # Use for running "diversion" report type.
            "diversion": ["D_LWSTN_CCT011", "D_SAC240_TCC001", "D_SAC207_GCC007", "D_NTOMA_FSC003", "D_MLRTN_FRK000", "D_MLRTN_MDC006",
                          "D_SAC030_MOK014", "TOTAL_EXP", "C_DMC003", "C_CAA003_CVP", "C_CAA003_SWP", "D_DMC007_CAA009"],
        }
        fields = fields_by_report_type[REPORT_TYPE]

        # alternatives to include
        # Map the DSS/model run names to the labels used in the appendix.
        # The keys remain the short names used for data lookup.
        alts = {
            'NAA': 'No Action Alternative',
            'Alt2v2_woTUCP': 'Alternative 2v2 without TUCP',
        }

        # Specify whether to use long names for alternatives in the appendix.
        use_long_name = False  # True to use long names for alternatives in the appendix, False to use short names

        # Prefix for tables and figures in appendix
        appendix_prefix_by_report_type = {"elevation": " F.2.1", "flow": " F.2.2", "diversion": " F.2.3"}
        appendix_prefix = appendix_prefix_by_report_type[REPORT_TYPE]

        # Path to file with location code crosswalk
        location_cw_path = os.path.join(base_dir, r"inputs\location_code_crosswalk_CalSim.xlsx")

        # Path to file with DSSReader output
        # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
        # WYT flags are read from monthly WYT_SAC_ and WYT_SJR_ columns in this file.
        dss_path = os.path.join(base_dir, r"inputs\DSS_contents.xlsx")

        # CalSim reports read WYT flags from dss_path instead of a separate file.
        wy_flags_path = None

        # Path to storage-elevation table data (only used for elevation).
        storage_elevation_table = os.path.join(base_dir, r"inputs\storage_elevation_table.xlsx")

        # Not used for CalSim appendices.
        compliance_fields = []
        compliance_dict = {}
        shastabin_data_path = ""

    elif REPORT_TYPE in SALINITY_REPORT_TYPES:
        ##### DSM2 salinity (EC / Cl / Position) inputs #####

        fields_by_report_type = {
            "EC": ["SAC_DS_STMBTSL", "CACHE_RYER", "RSAC123", "RSAC092", "RSAC101", "RSAN112", "RSAN112", "RSAN018", "ROLD024", "RSAN007",
                   "RSAC075", "RSAC081", "CHIPS_N_437", "CHIPS_S_442", "RSAC064", "CHDMC006", "CLIFTONCOURT", "ROLD034", "CHVCT000"],
            "Cl": ['ROLD024', 'RSAN007', 'CLIFTONCOURT', 'CHDMC006', 'SLBAR002'],
            "Position": ['X2'],
        }
        fields = fields_by_report_type[REPORT_TYPE]

        # alternatives to include
        # Map the DSS/model run names to the labels used in the appendix.
        # The keys remain the short names used for data lookup.
        alts = {
            'NAA': 'No Action Alternative',
            'Alt2v2_woTUCP': 'Alternative 2v2 without TUCP',
        }

        # Specify whether to use long names for alternatives in the appendix.
        use_long_name = False  # True to use long names for alternatives in the appendix, False to use short names

        # Prefix for tables and figures in appendix
        appendix_prefix_by_report_type = {"EC": " F.2.5", "Position": " F.2.6", "Cl": " F.2.7"}
        appendix_prefix = appendix_prefix_by_report_type[REPORT_TYPE]

        # Path to file with location code crosswalk
        location_cw_path = os.path.join(base_dir, r"inputs\location_code_crosswalk_salinity.xlsx")

        # Path to file with DSSReader output
        # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
        dss_path = os.path.join(base_dir, r"inputs\DSS_contents_CFS.xlsx")

        # Path to file with WY Typing data
        wy_flags_path = os.path.join(base_dir, r"inputs\wy_flags.xlsx")

        # Not used for salinity appendices.
        storage_elevation_table = ''
        compliance_fields = []
        compliance_dict = {}
        shastabin_data_path = ""

    elif REPORT_TYPE == "temperature":
        ##### Temperature (HEC-5Q) inputs #####

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
            "ABV CONFLUENCE",
        ]

        # alternatives to include
        # Map the DSS/model run names to the labels used in the appendix.
        # The keys remain the short names used for data lookup.
        alts = {
            'NAA': 'No Action Alternative',
            'Action 5': 'Action 5',
        }

        # Specify whether to use long names for alternatives in the appendix.
        use_long_name = False  # True to use long names for alternatives in the appendix, False to use short names

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

        # Prefix for tables and figures in appendix
        appendix_prefix = " F.2.11"

        # Path to file with location code crosswalk
        location_cw_path = os.path.join(base_dir, r"inputs\location_code_crosswalk_Temp.xlsx")

        # Path to file with DSSReader output
        dss_path = ""

        # File containing shasta bin information (By calendar yr) for each of the alternatives
        shastabin_data_path = os.path.join(base_dir, r"inputs\shasta_bin_info.xlsx")

        # Path to file with WY Typing data
        wy_flags_path = os.path.join(base_dir, r"inputs\wy_flags.xlsx")

        # Not used for the temperature appendix.
        storage_elevation_table = ''

    elif REPORT_TYPE == "water_supply":
        ##### Water supply inputs #####

        # alternatives to include
        alts = ['NAA', "Action 5"]

        # Prefix for tables and figures in appendix
        appendix_prefix = " F.2.4"

        # Path to file with the water supply calculation formulas
        s_supply_formulas = os.path.join(base_dir, r"inputs\water_supply_formulas.xlsx")

        # Path to file with DSSReader output
        dss_path = ""

        # Path to file with WY Typing data
        wy_flags_path = os.path.join(base_dir, r"inputs\wy_flags.xlsx")

    elif REPORT_TYPE == "compliance":
        ##### Water quality compliance inputs #####

        # This dictionary should hold the display name and the full DSS file for each alternative in the order you want
        # them displayed. All of these dss files should be in the studies folder.
        # Note that the hydrology should be in the name (ex: '2022MED').
        scenario_names = {
            'NAA': "NAA_2022Med_090723_EC_p.dss",
            "ALT1": "ALT1_2022Med_090923_EC_p.dss",
            "Alt2woTUCPwoVA": "ALT2v1_woTUCP_2022Med_091324_EC_p.dss",
            "Alt2wTUCPwoVA": "ALT2v1_wTUCP_2022Med_091324_EC_p.dss",
            "Alt2woTUCPDeltaVA": "ALT2v2_woTUCP_2022Med_091324_EC_p.dss",
            "Alt2woTUCPAllVA": "ALT2v3_woTUCP_2022Med_091324_EC_p.dss",
            "ALT3": "ALT3_2022Med_101323_EC_p.dss",
            "ALT4": "ALT4_2022MED_091624_EC_p.dss",
            "Action 5": "ALT5_wTUCP_2022Med_052125_EC_p.dss",
        }

    else:
        raise ValueError(
            f"Unrecognized REPORT_TYPE {REPORT_TYPE!r}. Expected one of "
            f"{sorted(CALSIM_REPORT_TYPES | SALINITY_REPORT_TYPES)}, 'temperature', 'water_supply', or 'compliance'."
        )

    ####END OF USER INPUTS #######

    os.makedirs(output_folder, exist_ok=True)

    # Dispatch to the engine function that matches the selected appendix.
    if REPORT_TYPE in CALSIM_REPORT_TYPES or REPORT_TYPE in SALINITY_REPORT_TYPES or REPORT_TYPE == "temperature":
        create_appendix(REPORT_TYPE, alts, fields, appendix_prefix, dss_path,
                        doc_name, new_doc, wy_flags_path=wy_flags_path, template=template,
                        location_cw_path=location_cw_path, use_calendar_yr=use_calendar_yr,
                        use_lumped_table_captions=use_lumped_table_captions,
                        storage_elevation_table=storage_elevation_table,
                        compliance_fields=compliance_fields,
                        compliance_dict=compliance_dict,
                        shastabin_data_path=shastabin_data_path,
                        use_long_name=use_long_name,
                        table_page_format=table_page_format,
                        plot_page_format=plot_page_format,
                        plot_format=plot_format)
    elif REPORT_TYPE == "water_supply":
        create_water_supply_appendix(alts, appendix_prefix, dss_path, doc_name, new_doc,
                                     wy_flags_path, template, s_supply_formulas)
    elif REPORT_TYPE == "compliance":
        create_compliance_appendix(scenario_names, template, doc_name, new_doc)
