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


file_pairs = pair_files("input")
for pair in tqdm(file_pairs, desc="Processing", bar_format="{l_bar}{bar} ETA: {remaining} Elapsed: {elapsed}"):
    text_path = f"input/{pair['text']}"
    graphics_path = f"input/{pair['graphics']}"

    path_to_docx = document_builder.create_word(text_pdf=text_path, graphics_pdf=graphics_path)

    path_to_pdf = path_to_docx.replace("docx", "pdf")

    convert(input_path=path_to_docx, output_path=path_to_pdf)
    document_builder.merge_pdfs(input_paths=[path_to_pdf, text_path, graphics_path], output_path=path_to_pdf)

    base_dir = path_to_pdf[:path_to_pdf.rindex('/')]
    os.makedirs(f"{base_dir}/pdf")
    shutil.move(text_path, f"{base_dir}/pdf/{pair['text']}")
    shutil.move(graphics_path, f"{base_dir}/pdf/{pair['graphics']}")
