from EISAppendixGen_functions import create_compliance_appendix

if __name__ == '__main__':

    ###USER INPUTS BELOW#####

    # this dictionary should hold the display name and the full DSS file for each alternative in the order you want them displayed
    # all of these dss files should be in the studies folder
    # note that the hydrology should be in the name (ex: '2022MED')
    # ex: {'NAA':"NAA_2022Med_090723_EC_p.dss",.... }
    scenario_names = {'NAA':"NAA_2022Med_090723_EC_p.dss",
                      "ALT1":"ALT1_2022Med_090923_EC_p.dss",
                      "Alt2woTUCPwoVA": "ALT2v1_woTUCP_2022Med_091324_EC_p.dss",
                      "Alt2wTUCPwoVA": "ALT2v1_wTUCP_2022Med_091324_EC_p.dss",
                      "Alt2woTUCPDeltaVA": "ALT2v2_woTUCP_2022Med_091324_EC_p.dss",
                      "Alt2woTUCPAllVA": "ALT2v3_woTUCP_2022Med_091324_EC_p.dss",
                      "ALT3": "ALT3_2022Med_101323_EC_p.dss",
                      "ALT4": "ALT4_2022MED_091624_EC_p.dss",
                      "Action 5": "ALT5_wTUCP_2022Med_052125_EC_p.dss"
    }

    # this is the template for the Word doc, generally doesn't need to change
    template = r"..\inputs\template_v2-fonts.docx"

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    # name of the temporary document
    doc_name = rf"temp_appendix.docx"
    # Name of final word doc
    new_doc = rf"appendix_water_quality_compliance_supply.docx"

    ####END OF USER INPUTS #######

    create_compliance_appendix(scenario_names, template, doc_name, new_doc)
