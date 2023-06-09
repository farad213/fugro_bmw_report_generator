import io

from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument


def pdf_to_text(path):
    """Converts a PDF file to plain text using the PyPDF2 library.

    Args:
        path (str): The path of the PDF file to be converted.

    Returns:
        list: A list of strings, where each string corresponds to the text extracted from a page in the PDF file.

    Raises:
        FileNotFoundError: If the specified path does not exist.
        ValueError: If the specified path does not point to a PDF file."""

    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    pages_text = []

    with open(path, 'rb') as fp:
        for page in PDFPage.get_pages(fp):
            outfp = io.StringIO()
            device = TextConverter(rsrcmgr, outfp, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            interpreter.process_page(page)
            page_text = outfp.getvalue()
            pages_text.append(page_text)
            outfp.close()

    return pages_text


def clean_text(pages):
    """Cleans and extracts data from a list of strings representing pages of a report generated by a pile testing software.

    Args:
        pages (list): A list of strings, where each string corresponds to the text extracted from a page of a pile testing report.

    Returns:
        dict: A dictionary containing the cleaned data extracted from the pages. The dictionary has the following keys:
            - building: A string representing the name of the building where the piles were tested.
            - pages_no: A string representing the total number of pages in the pile testing report.
            - date: A string representing the date when the pile testing was performed.
            - piles: A list of lists representing the measured values of different parameters for each pile. Each inner list contains the values for one pile.


    Raises:
        AssertionError: If the pages argument is not of type list.
        ValueError: If the pages list is empty or if any page in the list is empty or does not contain the expected data.
        IndexError: If any of the expected data fields (such as building, pages_no, or date) cannot be extracted from the first page of the report.
    """

    assert isinstance(pages, list), f"pages argument is not of type list, instead is type of {type(list)}"

    page = pages[0]
    misc_data = page.split("\n")

    building = misc_data[0].split("-")[1].strip()
    pages_no = misc_data[-3].removeprefix("1 of ").strip()
    date = misc_data[-5].removesuffix("-i mérés").strip()
    no_of_blows = int(misc_data[4].split(":")[1].strip())

    misc_data = {"building": building, "pages_no": pages_no, "date": date, "no_of_blows": no_of_blows}

    data = []
    for index, page in enumerate(pages):

        cutoff = page.index("Measured Length [m]") + len("Measured Length [m]")
        page = page[cutoff:]
        page_list = page.split("\n\n")
        page_list = page_list[:-3]
        page_list = [element for element in page_list if element]
        page_columns = [column.split("\n") for column in page_list]
        zipped_columns = zip(*page_columns)
        rows = [list(x) for x in zipped_columns]
        data.extend(rows)

    cleaned_data = misc_data | {"piles": data}
    return cleaned_data

def length_of_pdf(path):
    """
    Get the number of pages in a PDF file using PDFMiner.

    Parameters:
    path (str): The file path of the PDF file.

    Returns:
    int: The number of pages in the PDF file.

    Raises:
    PDFSyntaxError: If the PDF file is malformed or corrupt.

    """
    with open(path, 'rb') as file:
        parser = PDFParser(file)
        document = PDFDocument(parser)
        return len(document.catalog['Pages'].resolve()['Kids'])


