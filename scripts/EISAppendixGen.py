import pandas as pd
import numpy as np
import docx
from docx.table import _Cell
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Cm
from docx.shared import Pt
import subprocess
import os

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


if __name__ == "__main__":

    num_tables = 3

    # Test captions and alt text to add to tables
    captions = ['Table F.2.2-2-1c. Clear Creek below Whiskeytown Dam Flow, ALT1 minus NAA, Monthly Flow (cfs)',
                'Table F.2.2-2-4c. Clear Creek below Whiskeytown Dam Flow, Alt2woTUCPAllVA minus NAA, Monthly Flow (cfs)',
                'Table F.2.2-2-5a. Clear Creek below Whiskeytown Dam Flow, NAA, Monthly Flow (cfs)']

    alt_text = ["Alt text example 1", "Alt text example 2", "Alt text example 3"]

    # Windows command prompt can't save to OneDrive bc of the space in the file path, save locally instead
    # Pass absolute paths to VBS
    #Name of intermediate word doc
    doc_name = "C:/Users/emadonna/OneDrive - DOI/emadonna/baseline/eis_appendices/appendix_temp.docx"
    #Name of final word doc
    new_doc = "C:/Users/emadonna/OneDrive - DOI/emadonna/baseline/eis_appendices/appendix_final.docx"

    ##### Generate dfs with dummy data ########

    dfs = []
    for i in range(num_tables):
        df = pd.DataFrame(np.random.randint(0, 5000, size=(15, 12)),
                          columns=["Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"])

        # Add row labels in first column
        df.insert(0, "Statistic",
                  ["10% Exceedance", "20% Exceedance", "30% Exceedance", "40% Exceedance", "50% Exceedance",
                   "60% Exceedance", "70% Exceedance", "80% Exceedance", "90% Exceedance",
                   "Full Simulation Period Average", "Wet Water Years (28%)", "Above Normal Water Years (14%)",
                   "Below Normal Water Years (18%)", "Dry Water Years (24%)", "Critical Water Years (16%)"])

        # Move new header names to first row
        df.index = df.index + 1  # shifting index
        df = df.sort_index()
        dfs.append(df)

    ##### Use docx package to create a document with formatted table objects and save to Word .docx file ###########

    # Create an instance of a word document
    doc = docx.Document()

    ## Wrap all code in loop to add multiple tables to one doc
    for i in range(len(dfs)):
        # Add caption above table
        p = doc.add_paragraph()
        run = p.add_run(captions[i])
        run.font.bold = True
        run.font.size = Pt(12)

        # add a table to the end and create a reference variable
        # extra row is so we can add the header row
        t = doc.add_table(dfs[i].shape[0] + 1, dfs[i].shape[1])

        # add the header rows.
        for j in range(dfs[i].shape[-1]):
            t.cell(0, j).text = dfs[i].columns[j]

        # add the rest of the data frame
        for k in range(dfs[i].shape[0]):
            for j in range(dfs[i].shape[-1]):
                t.cell(k + 1, j).text = str(dfs[i].values[k, j])

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
        for row in list(range(dfs[i].shape[0] + 1)):
            for column in bolding_columns:
                t.rows[row].cells[column].paragraphs[0].runs[0].font.bold = True

        # Add borders to middle row and under header
        for cell in t.rows[0].cells:
            set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

        for cell in t.rows[10].cells:
            set_cell_border(cell, bottom={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

        for cell in t.rows[10].cells:
            set_cell_border(cell, top={"sz": 7, "color": "#000a00", "space": 0.5, "val": "single"})

        # Change font size to fit on page better
        change_table_font_size(doc, 9)

        # Widen margins of table
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(0.5)
            section.bottom_margin = Cm(0.5)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)

        # Widen cell size in first column
        for cell in t.columns[0].cells:
            cell.width = Inches(3.2)

        # Add commas to values in table
        add_commas_to_table(doc)

        # Align values in center of cells
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_ALIGN_VERTICAL

        for row in t.rows:
            for cell in row.cells:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Add footnotes to one of the tables as proof of concept
        if i == 1:
            # Add footnotes at end of table
            f = doc.add_paragraph()
            run = f.add_run('* All scenarios are simulated at 2022 Median climate condition and 15 cm sea level rise.')
            run.font.size = Pt(9)
            f.paragraph_format.space_before = Pt(1)
            f.paragraph_format.space_after = Pt(1)

            f1 = doc.add_paragraph()
            run = f1.add_run(
                '* Water Year Types defined by the Sacramento Valley 40-30-30 Index Water Year Hydrologic Classification (SWRCB D-1641, 1999).')
            run.font.size = Pt(9)
            f1.paragraph_format.space_before = Pt(1)
            f1.paragraph_format.space_after = Pt(1)

            f2 = doc.add_paragraph()
            run = f2.add_run('* Water Year Types results are displayed with calendar year – year type sorting.')
            run.font.size = Pt(9)
            f2.paragraph_format.space_before = Pt(1)
            # paragraph.paragraph_format.space_after = Pt(5)

        doc.add_page_break()

    # Save table to word doc
    doc.save(doc_name)

    ##### Use Python to Run VBS Script that adds alt text to table in saved docx file #######

    # Format alt text for all tables as one string to be passed to vbs
    alt_text_string = ("+").join(alt_text)
    alt_text_string = alt_text_string.replace(" ", "_")

    #Run vbs script
    #Arguments are existing document, new document to be saved to, alternative text for all tables, number of tables
    #This will fail if Microsoft Word has document open in the background
    #try opening Task Manager and Ending MS Word Background Task, then rerun
    result = subprocess.call("cscript.exe add_alt_text.vbs " + doc_name + " " + new_doc + " " + alt_text_string + " " + str(num_tables))

    #Remove temporary doc if process ran successfully
    if result == 0:
        os.remove(doc_name)
    else:
        print("VBS script did not run successfully. Try using task manager to end MS Word Background Task and then rerun")