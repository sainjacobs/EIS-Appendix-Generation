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
    # Map the DSS/model run names to the labels used in the appendix.
    # The keys remain the short names used for data lookup.
    alts = {
        'NAA': 'No Action Alternative',
        'Alt2v2_woTUCP': 'Alternative 2v2 without TUCP',
    }

    # Specify whether to use long names for alternatives in the appendix.
    use_long_name = False      # True to use long names for alternatives in the appendix, False to use short names

    # Formatting for pages containing tables. Measurements ending in "_pt"
    # are in points; row_height_cm is in centimeters.
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
    # line style for every alternative included above.
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

    # Specify whether report is "flow", "elevation", or "diversion"
    # Note 1: "elevation" option also includes storages.
    report_type = "flow"

    # For NAA vs alternative comparison tables, specify whether you want the table captions lumped or not.
    use_lumped_table_captions = False

    # Select whether to use the calendar year to group data.
    use_calendar_yr = True  # Note: For Trinity LTO tables/figures, use False.

    # Prefix for tables and figures in appendix
    appendix_prefix = " F.2.2"  # F.2.1 is elevation; F.2.2 is flow; F.2.3 is diversion

    # Define base working directory for reference
    base_dir = r"C:\Github\EIS-Appendix-Generation"

    # Change directory to scripts: SN 20260303
    os.chdir(os.path.join(base_dir, "scripts"))

    # Path to file with location code crosswalk
    location_cw_path = os.path.join(base_dir, r"inputs\location_code_crosswalk_CalSim.xlsx")

    # Path to file with DSSReader output
    # for water supply, must be the _TAF output
    # Use output from DSS reader in desired units (CFS or TAF). Use TAF for elevation/storage and CFS for the flow and diversion appendices.
    # WYT flags are read from monthly WYT_SAC_ and WYT_SJR_ columns in this file.
    # The appendix code uses the May value from each water year as the annual WYT.
    dss_path =  os.path.join(base_dir, r"inputs\DSS_contents.xlsx")

    # Path to storage-elevation table data
    storage_elevation_table = os.path.join(base_dir, r"inputs\storage_elevation_table.xlsx")

    # Output directory for generated Word docs and plot folders.
    output_folder = r"C:\\20251211_BA_Modeling_Appendix\\outputs_gitRepo\\trial"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc - update parent directory
    template = os.path.join(base_dir, r"inputs\template_v2-fonts.docx")
    doc_name = os.path.join(output_folder, "appendix_temp.docx")
    # Name of final word doc
    new_doc = os.path.join(output_folder, f"appendix_final_{report_type}.docx")

    ####END OF USER INPUTS #######

    os.makedirs(output_folder, exist_ok=True)

    # call the corresponding function for the appendix
    create_appendix(report_type, alts, fields, appendix_prefix, dss_path,
                    doc_name, new_doc, wy_flags_path=None, template=template,
                    location_cw_path=location_cw_path, use_calendar_yr=use_calendar_yr,
                    use_lumped_table_captions=use_lumped_table_captions,
                    storage_elevation_table=storage_elevation_table,
                    use_long_name=use_long_name,
                    table_page_format=table_page_format,
                    plot_page_format=plot_page_format,
                    plot_format=plot_format)
