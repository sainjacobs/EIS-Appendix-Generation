import pandas as pd
import numpy as np
import docx
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_ROW_HEIGHT_RULE
import matplotlib.pyplot as plt
import calendar
import os
from time import strptime

def get_locations(location_crosswalk_path, fields):
    """
    Gets location names from field codes passed

    Parameters
    ----------
    location_crosswalk_path: string
        Path and file name for xlsx file containing location names and field codes
    fields: list of strings
        Names of the fields to be processed

    """
    #Read in crosswalk as a df
    crosswalk = pd.read_excel(location_crosswalk_path)

    #Look up each field code's corresponding location title and add to a list
    locations = []
    for f in fields:
        locations.append(crosswalk.loc[crosswalk["DSSPartB"] == f, "Location (Title)"].values[0])

    return locations
def parse_dssReader_output(dss_path, runs, field):
    """
    Reads DSS output from reader for desired runs and field

    Parameters
    ----------
    dss_path: string
        Path and file name for xlsx file containing DSSReader Output
    runs: list of strings
        Names of the runs to be processed
    field: string
        Current field being processed

    """
    #Read DSS Output from specified path for specified field
    dss_output = pd.read_excel(dss_path)

    dss_output = dss_output[["Month", "Scenario", "WY", field]]

    # Create df for each alternative/run and reformat
    run_dfs = []
    for run in runs:
        if run == "NAA":
            run_df = dss_output.loc[dss_output["Scenario"] == "Baseline"]
        else:
            run_df = dss_output.loc[dss_output["Scenario"] == run]

        run_df["month_name"] = " "

        #Add abbrievated month name to df for tables and plotting later
        for index, row in run_df.iterrows():
            run_df.loc[index, "month_name"] = calendar.month_abbr[row["Month"]]
        #Drop unneeded columns
        run_df.drop(columns=["Month", "Scenario"], inplace=True)
        run_dfs.append(run_df)

    #Transpose dfs to be in correct format for tables
    t_dfs = []
    for run_df in run_dfs:
        transposed_df = pd.DataFrame(
            columns=["WY", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"])
        #One row for each WY consisting of a column for each monthly EC value
        for year in np.unique(run_df["WY"]):
            year_t = run_df.loc[run_df["WY"] == year]
            year_t.set_index("month_name", inplace=True)
            year_t = year_t.transpose()
            year_t.insert(0, "WY", year)
            year_t.reset_index(drop=True, inplace=True)
            #Add each year as new row to df
            transposed_df = pd.concat([transposed_df, year_t.iloc[1:2]], axis=0, ignore_index=True)
        t_dfs.append(transposed_df)

    return t_dfs

def create_exceedance_tables(t_dfs, wy_flags_path):
    """
    Creates exceedance tables formatted for appendix report from transposed DSSReader Output

    Parameters
    ----------
    t_dfs: list of dataframes
        Dataframe outputs from DSSReader that have been transposed to be formatted for table

    """
    exc_tables = []
    for table in t_dfs:
        table.drop(columns = ["WY"], inplace = True)
        #Sort df from highest monthly EC to lowest
        #table = table.sort_values(by=table.columns.tolist(), ascending=False)
        table = table.apply(lambda x: x.sort_values(ascending=False).values)
        #Remove first and last rows
        #table.iloc[::-1, ::1]
        #Rank ECs from 1 to 100
        table.insert(0, "Rank", range(1,101))
        #Calculate exceedance probability and add column to table
        table.insert(1, "Exc Prob", (table["Rank"] - 1)/(table.shape[0]-1)*100)
        #Round all table values to 1 decimal
        table = table.round(1)
        #Keep only every 10th row so that only 10, 20, 30, etc. summary percents are shown in table
        table = table.iloc[::10]
        #Drop first row so that table starts at 10% exceedance prob
        table.drop(index=table.index[0], axis=0, inplace=True)

        exc_tables.append(table)

    #Calculate full simulation period average for each run and format to be added to exceedance table as one row
    stats_dfs = []
    for t in t_dfs:
        period_ave = t.mean(axis=0)
        stats_df = pd.DataFrame(period_ave)
        stats_df = stats_df.transpose()

        stats_df["Exc Prob"] = ["Full Simulation Period Average"]

        stats_dfs.append(stats_df)

    #Read in year typing flags
    wy_flags = pd.read_excel(wy_flags_path)
    year_types = ["Wet", "Above Normal", "Below Normal", "Dry", "Critical"]

    # make a copy of exc probabilities to use with figures after deleting from tables df
    exc_probs = exc_tables[0]["Exc Prob"]

    # calculate wet, above normal, dry, etc water years (EC sum for year type/ count of year type)
    for i in range(len(t_dfs)):
        t_dfs[i]["flag"] = wy_flags["Year Type"]
        month_vals = {}
        # Also add full sim period average as a row in exceedance table
        exc_tables[i] = pd.concat([exc_tables[i], stats_dfs[i].iloc[0:1]], ignore_index=True)

        #Iterate through each type of year (wet, above normal, etc) to compute sums
        for y in range(len(year_types)):
            for month in t_dfs[i].columns[:-1]:
                #Flags are 1 - 5 to specify which type of year
                #Calculate mean of months classified as current year type based on flag
                month_vals[month] = [t_dfs[i].loc[t_dfs[i]['flag'] == (y + 1), month].mean()]

            month_vals["Exc Prob"] = year_types[y]

            exc_tables[i] = pd.concat([exc_tables[i], pd.DataFrame.from_dict(month_vals)], ignore_index=True)
        exc_tables[i].drop(columns=["Rank", "Exc Prob"], inplace=True)
        exc_tables[i] = exc_tables[i].astype(float).round(1)
        # Add row labels for report tables in first column
        exc_tables[i].insert(0, "Statistic",
                             ["10% Exceedance", "20% Exceedance", "30% Exceedance", "40% Exceedance", "50% Exceedance",
                              "60% Exceedance", "70% Exceedance", "80% Exceedance", "90% Exceedance",
                              "Full Simulation Period Average", "Wet Water Years (28%)",
                              "Above Normal Water Years (14%)", "Below Normal Water Years (18%)",
                              "Dry Water Years (24%)", "Critical Water Years (16%)"])
        # Move new header names to first row
        exc_tables[i].index = exc_tables[i].index + 1  # shifting index
        exc_tables[i] = exc_tables[i].sort_index()

    return exc_tables, exc_probs


def make_rows_bold(*rows):
    """
    Makes text in specified table rows bold.

    Parameters
    ----------
    rows: row attributes from docx table object
        1 or more rows that will be converted to bold text (ex table.rows[0])

    """

    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def set_cell_border(cell: _Cell, **kwargs):
    """
    Set cell's border.
    Usage:

    set_cell_border(
        cell,
        top={"sz": 12, "val": "single", "color": "#FF0000", "space": "0"},
        bottom={"sz": 12, "color": "#00FF00", "val": "single"},
        start={"sz": 24, "val": "dashed", "shadow": "true"},
        end={"sz": 12, "val": "dashed"},
    )

     Parameters
    ----------
    cell: cell attribute from docx table object
        1 cell that will be converted to bold text (ex table.rows[0].cells[0])

    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # check for tag existnace, if none found, then create one
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # list over all available tags
    for edge in ('start', 'top', 'end', 'bottom', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # check for tag existnace, if none found, then create one
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)

            # looks like order of attributes is important
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


def change_table_font_size(document, font_size):
    """
    Changes the font size of all text in all tables within a document.

    Parameters
    ----------
    document: docx document object
        Document that will have font size adjusted
    font_size: int
        New font size

    """

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = docx.shared.Pt(font_size)


def add_commas_to_table(doc):
    """
    Adds commas to numbers in all tables of a docx document.

    Parameters
    ----------
    doc: docx document object
        Document that will have commas added to numeric values in all tables

    """

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        try:
                            # Check if the text is a number
                            number = float(run.text)
                            # Format the number with commas
                            formatted_number = f"{number:,}"
                            formatted_number = formatted_number.rsplit(".", 1)[0]
                            run.text = formatted_number
                        except ValueError:
                            # If the text is not a number, do nothing
                            pass


def change_orientation(doc, new_orientation):
    """
    Changes section orientation from portrait to landscape or vice versa

    Parameters
    ----------
    doc: docx document object
        Document that will have commas added to numeric values in all tables
    new_orientation: string
        Either "landscape" or "portrait" to indicate the desire page orientation for the new section

    """

    current_section = doc.sections[-1]
    new_width, new_height = current_section.page_height, current_section.page_width
    new_section = doc.add_section(WD_SECTION.NEW_PAGE)
    if new_orientation == "landscape":
        new_section.orientation = WD_ORIENT.LANDSCAPE
    else:
        new_section.orientation = WD_ORIENT.PORTRAIT

    new_section.page_width = new_width
    new_section.page_height = new_height

    return new_section

def format_table(t, table, doc):
    """
    Creates tables formatted for appendix report from exceedance tables

    Parameters
    ----------
    t: docx table object
        Exceedance table to be formatted for report
    doc: docx object
        Docx object containing table to be formatted
    """
    # Change font size to fit on page better
    change_table_font_size(doc, 8)

    # add the header rows.
    for j in range(table.shape[-1]):
        t.cell(0, j).text = table.columns[j]

    # add the rest of the data frame
    for g in range(table.shape[0]):
        for j in range(table.shape[-1]):
            t.cell(g + 1, j).text = str(table.values[g, j])

    # Set table top and bottom borders
    borders = OxmlElement('w:tblBorders')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), '4')
    borders.append(bottom_border)
    top_border = OxmlElement('w:top')
    top_border.set(qn('w:val'), 'single')
    top_border.set(qn('w:sz'), '4')
    borders.append(top_border)

    t._tbl.tblPr.append(borders)

    # Make headers bold
    make_rows_bold(t.rows[0])

    # Make first column bold
    bolding_columns = [0]
    for row in list(range(table.shape[0] + 1)):
        for column in bolding_columns:
            t.rows[row].cells[column].paragraphs[0].runs[0].font.bold = True

    # Add superscript to Full Simulation Period Average cell
    script_cell = t.cell(10, 0).paragraphs[0]
    run = script_cell.add_run("a")
    run.font.superscript = True

    # Add borders to middle row and under header
    for cell in t.rows[0].cells:
        set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    for cell in t.rows[10].cells:
        set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    for cell in t.rows[10].cells:
        set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

    # Widen margins of table
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(0.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # Widen cell size in first column
    for cell in t.columns[0].cells:
        cell.width = Inches(3.4)

    # Add commas to values in table
    add_commas_to_table(doc)

    # Align values in center of cells
    for row in t.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Decrease row spacing for table
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        row.height = Cm(0.45)  # 2 cm

def create_month_plot(fig_dfs, month, month_directory, alts, line_styles, line_colors):
    """
    Generates and saves individual month plots

    Parameters
    ----------
    fig_dfs: list of dataframes
        Dataframes with exceedance values by month
    month: string
        Current month to be evaluated
    month_directory: string
        Directory to save month plots in
    alts: list of strings
        Names of the runs being compared in report
    line_styles: list of strings
        Styles for lines on plots
    line_colors: list of strings
        Colors for lines on plots
    """
    # Check for/create directory to store monthly exceedance plots
    if not os.path.exists(month_directory):
        os.makedirs(month_directory)

    fig, axs = plt.subplots(figsize=(10, 5), linewidth=3, edgecolor="black")
    for s in range(len(fig_dfs)):
        # plot exceedance probability vs monthly EC
        axs.plot(fig_dfs[s]["exc_prob"], fig_dfs[s][month], color=line_colors[s], linestyle=line_styles[s])

        # flip x-axis
        plt.gca().invert_xaxis()

        axs.set_ylabel("Monthly Flow (cfs)")
        axs.set_xlabel("Exceedance Probability")

        # Save this parameter to orient the legend correctly
        axbox = axs.get_position()

        # Add gridlines
        plt.grid(color='gray', linestyle='--', linewidth=0.8)

        # Add a legend
        plt.legend(labels=alts, loc='center', ncol=4, bbox_to_anchor=[axbox.x0 + 0.5 * axbox.width, 1.08])

    # Add month number at beginning so that figures can be easily inserted in CY order to document later
    month_number = str(strptime(month, '%b').tm_mon)
    # Add leading zeros to month numbers
    if len(month_number) < 2:
        month_number = str(0) + month_number
    # Save figure to month directory
    plt.savefig(month_directory + "/" + month_number + "_" + month + "_monthly_exceedance" + ".png")
    plt.close()

def create_stat_plot(stat_fig_dfs, stat, stat_directory, alts, line_styles, line_colors):
    """
    Generates and saves individual month plots

    Parameters
    ----------
    stat_fig_dfs: list of dataframes
        Dataframes with average values by year type
    stat: string
        Current type of year to be evaluated
    stat_directory: string
        Directory to save stat plots in
    alts: list of strings
        Names of the runs being compared in report
    line_styles: list of strings
        Styles for lines on plots
    line_colors: list of strings
        Colors for lines on plots
    """
    if not os.path.exists(stat_directory):
        os.makedirs(stat_directory)

    fig, axs = plt.subplots(figsize=(10, 5), linewidth=3, edgecolor="black")
    for s in range(len(stat_fig_dfs)):
        if stat == "Full Simulation Period":
            axs.plot(stat_fig_dfs[s]["month"], stat_fig_dfs[s]["Full Simulation Period Average"], color=line_colors[s],
                     linestyle=line_styles[s])
        else:
            axs.plot(stat_fig_dfs[s]["month"], stat_fig_dfs[s][stat], color=line_colors[s],
                     linestyle=line_styles[s])

        # Save this to position legend correctly
        axbox = axs.get_position()

        axs.set_ylabel("Monthly Flow (cfs)")

        # Add gridlines
        plt.grid(color='gray', linestyle='--', linewidth=0.8)
        # Add legend
        plt.legend(labels=alts, loc='center', ncol=4, bbox_to_anchor=[axbox.x0 + 0.5 * axbox.width, 1.08])

    # Save stat fig to directory
    plt.savefig(stat_directory + "/" + stat[:5] + "_exceedance" + ".png")
    plt.close()