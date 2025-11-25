from EISAppendixGen_functions import create_water_supply_appendix

if __name__ == "__main__":

    ###USER INPUTS BELOW#####

    # alternatives to include
    alts = ['NAA', "Action 5"]

    # Prefix for tables and figures in appendix
    appendix_prefix = " F.2.4"

    # Path to file with location code crosswalk
    s_supply_formulas = r"..\inputs\water_supply_formulas.xlsx"

    # Path to file with DSSReader output
    # for water supply, must be the _TAF output
    dss_path = r""

    # Path to file with WY Typing data
    wy_flags_path = "..\inputs\wy_flags.xlsx"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # Name of intermediate word doc - update parent directory
    template = r"..\inputs\template_v2-fonts.docx"
    doc_name = r"appendix_temp.docx"
    # Name of final word doc
    new_doc = r"appendix_final_water_supply.docx"

    ####END OF USER INPUTS #######

    # call the corresponding function for the appendix
    create_water_supply_appendix(alts, appendix_prefix, dss_path, doc_name, new_doc, wy_flags_path, template, s_supply_formulas)
