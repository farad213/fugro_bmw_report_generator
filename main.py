import os
import re
import shutil

from docx2pdf import convert
from tqdm import tqdm

import document_builder


def pair_files(input_folder):
    assert isinstance(input_folder, str)

    file_pairs = []

    # Get a list of all files in the input folder
    all_files = os.listdir(input_folder)

    # Define a regex pattern to match the shared part of the filenames
    pattern = re.compile(r"(text|graphics)_(.*?)\.pdf")

    # Loop through all files and group them by shared part of the filename
    for filename in all_files:
        # Try to match the regex pattern to the filename
        match = pattern.match(filename)
        if match:
            # Extract the shared part of the filename
            file_type = match.group(1)
            shared_part = match.group(2)
            # If this is the text file of a pair, create a new dict entry
            if file_type == "text":
                file_pair = {"text": filename}
                # If the graphics file exists, add it to the dict entry
                graphics_filename = f"graphics_{shared_part}.pdf"
                if graphics_filename in all_files:
                    file_pair["graphics"] = graphics_filename
                file_pairs.append(file_pair)

    return file_pairs


def run():
    file_pairs = pair_files("input")
    master_check = {"faulty": {}, "missing": {}, "blow_counter": {}}
    master_faulty_check_list = []
    master_missing_check_list = []
    for pair in tqdm(file_pairs, desc="Processing", bar_format="{l_bar}{bar} ETA: {remaining} Elapsed: {elapsed}"):
        text_path = f"input/{pair['text']}"
        graphics_path = f"input/{pair['graphics']}"

        faulty_piles, missing_piles, path_to_docx, no_of_piles, no_of_blows = document_builder.create_word(
            text_pdf=text_path,
            graphics_pdf=graphics_path)

        path_to_pdf = path_to_docx.replace("docx", "pdf")

        convert(input_path=path_to_docx, output_path=path_to_pdf)
        document_builder.merge_pdfs(input_paths=[path_to_pdf, text_path, graphics_path], output_path=path_to_pdf)

        base_dir = path_to_pdf[:path_to_pdf.rindex('/')]

        if faulty_piles:
            master_check["faulty"] = master_check["faulty"] | {base_dir: faulty_piles}
            master_faulty_check_list.append(faulty_piles)
        else:
            master_check["faulty"] = master_check["faulty"] | {base_dir: "OK"}
        if missing_piles:
            master_check["missing"] = master_check["missing"] | {base_dir: missing_piles}
            master_missing_check_list.append(missing_piles)
        else:
            master_check["missing"] = master_check["missing"] | {base_dir: "OK"}
        master_check["blow_counter"] = master_check["blow_counter"] | {base_dir: f"Piles: {no_of_piles}/{no_of_blows}"}

        os.makedirs(f"{base_dir}/pdf")
        shutil.move(text_path, f"{base_dir}/pdf/{pair['text']}")
        shutil.move(graphics_path, f"{base_dir}/pdf/{pair['graphics']}")

    if master_faulty_check_list or master_missing_check_list or no_of_piles != no_of_blows:
        with open("output/MASTER_CHECK_ERROR_FAULTY_OR_MISSING_PILES_FOUND.txt", "w", encoding="utf-8") as file:
            for key in master_check["faulty"].keys():
                _, building, date = key.split("/")
                date = date.removesuffix("-i mérés")
                value = master_check["blow_counter"][key]
                file.write(f"{building:6}{date + ':':15} {value}\t\t")
                value = master_check["missing"][key]
                file.write(f"Missing: {value}\t\t")
                value = master_check["faulty"][key]
                file.write(f"Faulty: {value}\n")

    else:
        with open("output/MASTER_CHECK_OK.txt", 'w', encoding="utf-8") as file:
            for key in master_check["faulty"].keys():
                _, building, date = key.split("/")
                date = date.removesuffix("-i mérés")
                value = master_check["blow_counter"][key]
                file.write(f"{building:6}{date + ':':15} {value}\t\t")
                value = master_check["missing"][key]
                file.write(f"Missing: {value}\t\t")
                value = master_check["faulty"][key]
                file.write(f"Faulty: {value}\n")

    return None


if __name__ == "__main__":
    run()
