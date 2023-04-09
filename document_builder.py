import pandas as pd
from docxtpl import DocxTemplate
import analyzer
import os
from PyPDF2 import PdfMerger, PdfReader


def create_word(text_pdf, graphics_pdf):
    assert isinstance(text_pdf, str)
    assert isinstance(graphics_pdf, str)

    pages = analyzer.pdf_to_text(text_pdf)
    source_dict = analyzer.clean_text(pages)

    date = source_dict["date"]
    building = source_dict["building"]

    df_text = pd.DataFrame(data=source_dict["piles"],
                           columns=["Pile Name", "v peak [mm/s]", "a peak [m/s2]", "t50% [ms]", "Measured Length [m]"])
    df_database = pd.read_excel("source/database/BMW_report_buildings_data.xlsx", sheet_name=building, header=1).dropna(
        axis=1)

    full_building_name = df_database["Épület neve"][0]

    table_rows = []
    faulty_piles = []
    missing_piles = []
    type_of_piles, diameter_of_piles = set(), set()

    # goes through the piles found in the pdf and extracts relevant information from the database for each of the piles
    for index, pile in enumerate(df_text["Pile Name"]):
        pile_no = int(pile[:pile.index("-")].replace("P", ""))
        row_select = df_database["Cölöp jele"] == pile_no

        if not row_select.any():
            missing_piles.append(pile)
            continue


        ccs = round(float(df_database[row_select]["Cölöpcsúcs"].iloc[0]), 2)
        vvs = round(float(df_database[row_select]["Visszavésési szint"].iloc[0]), 2)
        measured_pile_length = round(float(df_text["Measured Length [m]"][index].replace(",", ".")), 2)
        measure_pile_toe_level = round(vvs - measured_pile_length, 2)

        # keeping track of all the types and diameters encountered
        type_of_piles.add(df_database[row_select]["Cölöp típus"].iloc[0])
        diameter_of_piles.add(df_database[row_select]["Cölöp átmérő"].iloc[0])

        # checks if a pile is faulty or not
        if not ccs - 0.2 <= measure_pile_toe_level <= ccs:
            faulty_piles.append(pile)

        # initializes min_length and max_length with the first values it can find, related to them
        if index == 0:
            min_length, max_length = measured_pile_length, measured_pile_length
        # checks if there is a value smaller than min_length or larger than max_length
        # and updates these values if any of the above conditions are met
        else:
            if measured_pile_length < min_length:
                min_length = measured_pile_length
            if measured_pile_length > max_length:
                max_length = measured_pile_length

        # compiles a list of dictionaries that will make up the table in the docx
        table_rows.append({"pile": pile_no,
                           "ccs": f"{ccs:.2f}",
                           "vvs": f"{vvs:.2f}",
                           "measured_pile_length": f"{measured_pile_length:.2f}",
                           "measured_pile_toe_level": f"{measure_pile_toe_level:.2f}",
                           "min_length": f"{round(min_length, 2):.2f}",
                           "max_length": f"{round(max_length, 2):.2f}"})

    type_of_piles_str = ""
    for index, pile_type in enumerate(sorted(list(type_of_piles))):
        type_of_piles_str += pile_type
        if index + 1 < len(type_of_piles):
            type_of_piles_str += ", "

    diameter_of_piles_str = ""
    for index, pile_diameter in enumerate(sorted(list(diameter_of_piles))):
        diameter_of_piles_str += pile_diameter
        if index + 1 < len(diameter_of_piles):
            diameter_of_piles_str += ", "

    # get the sum of pages of all attachments
    attachment_length = analyzer.length_of_pdf(text_pdf) + analyzer.length_of_pdf(graphics_pdf)

    # determine the no of pages in the expert advice section based on the no of piles
    # (and thus based on the size of the table in the generated docx file)
    no_of_piles = len(df_text["Pile Name"]) + 1
    if no_of_piles <= 25:
        expert_advice_length = 4
    else:
        expert_advice_length = int((no_of_piles - 25) / 46) + 4 + 1

    # building context dict
    context = {"table_rows": table_rows,
               "date": date,
               "building_notation": building,
               "full_building_name": full_building_name,
               "min_length": min_length,
               "max_length": max_length,
               "type_of_piles": type_of_piles_str,
               "diameter_of_piles": diameter_of_piles_str,
               "expert_advice_length": expert_advice_length,
               "attachment_length": attachment_length
               }

    template = DocxTemplate('source/template/BMW_report_template.docx')
    template.render(context)
    if not os.path.exists(f"output/{building}"):
        os.makedirs(f"output/{building}")
    if not os.path.exists(f"output/{building}/{date}-i mérés"):
        os.makedirs(f"output/{building}/{date}-i mérés")

    path_to_docx = f'output/{building}/{date}-i mérés/FCH-20091_SIT_DEBRECEN_BMW - {building}_{date.replace(".", "")}_HBM_report.docx'

    template.save(path_to_docx)

    if faulty_piles or missing_piles:
        with open(f"output/{building}/{date}-i mérés/ERROR_FAULTY_OR_MISSING_PILES_FOUND.txt", "w", encoding="utf-8") as file:
            for pile in faulty_piles:
                file.write(f"Faulty: {pile}\n")
            for pile in missing_piles:
                file.write(f"Missing: {pile}\n")
    else:
        with open(f"output/{building}/{date}-i mérés/OK.txt", "w", encoding="utf-8") as file:
            pass

    return faulty_piles, missing_piles, path_to_docx


def merge_pdfs(input_paths, output_path):
    """
    Merge multiple PDFs into one in the order they are given in the input_paths list.
    :param input_paths: A list of file paths of the input PDF files.
    :param output_path: The file path of the output PDF file.
    """
    merger = PdfMerger()

    for path in input_paths:
        with open(path, 'rb') as f:
            merger.append(PdfReader(f))

    with open(output_path, 'wb') as f:
        merger.write(f)
