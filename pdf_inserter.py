import gc
import pypdf
import pdfplumber
from pathlib import Path
import win32com.client

# Constants for Word-to-PDF conversion
WD_PDF_FORMAT = 17


def create_pdf_from_word(word_path: Path, output_path: Path) -> None:
    """
    Converts a Word document to a PDF.

    :param Path word_path: The path to the Word document.
    :param Path output_path: The path where the resulting PDF will be saved.
    """
    word = win32com.client.Dispatch('Word.Application')
    word.DisplayAlerts = False
    doc = None
    try:
        doc = word.Documents.Open(str(word_path))
        doc.ShowRevisions = False
        doc.PrintRevisions = False
        doc.SaveAs(str(output_path), FileFormat=WD_PDF_FORMAT)
        doc.Close()
    except Exception as e:
        if doc is not None:
            doc.Close()
        raise RuntimeError(f"{e}") from e
    finally:
        word.Quit()
        gc.collect()


def insert_pdfs_at_keywords(main_pdf_path: Path, insert_pdfs: dict[str, list[Path]], keyword_page_map: dict[str, int],
                            output_pdf_path: Path) -> None:
    """
    Inserts PDFs at specified keyword locations within a main PDF.

    :param Path main_pdf_path: Path to the main PDF where the inserts will be made.
    :param dict insert_pdfs: A dictionary mapping keywords to lists of PDF paths to insert.
    :param dict keyword_page_map: A dictionary mapping keywords to the page numbers they were found on.
    :param Path output_pdf_path: Path where the output PDF will be saved.
    """
    main_reader = pypdf.PdfReader(main_pdf_path)
    output_writer = pypdf.PdfWriter()

    # Iterate through the pages of the main PDF and add each to the output PDF.
    for page_num in range(len(main_reader.pages)):
        page = main_reader.pages[page_num]
        output_writer.add_page(page)

        # Check if there are PDFs to insert after this page.
        for keyword, insert_pdf_paths in insert_pdfs.items():
            if keyword in keyword_page_map and keyword_page_map[keyword] == page_num:
                # Insert the associated PDFs if the keyword is found on this page.
                for insert_pdf_path in insert_pdf_paths:
                    with open(insert_pdf_path, 'rb') as insert_file:
                        insert_reader = pypdf.PdfReader(insert_file)
                        for insert_page in insert_reader.pages:
                            output_writer.add_page(insert_page)

    # Write the modified content to the output file.
    with open(output_pdf_path, 'wb') as output_file:
        output_writer.write(output_file)


def find_keywords_in_pdf(main_pdf_path: Path, keywords: list[str]) -> dict[str, int]:
    """
    Finds specified keywords in a PDF and returns their corresponding page numbers.

    :param Path main_pdf_path: Path to the main PDF file to search.
    :param list[str] keywords: List of keywords to find in the PDF.
    :return: A dictionary mapping each keyword to the page number it was found on.
    :rtype: dict[str, int]
    """
    keyword_page_map = {}

    # Open the main PDF with pdfplumber for text extraction.
    with pdfplumber.open(main_pdf_path) as pdf:
        pdf_pages_len = len(pdf.pages)
        # Iterate through the pages in reverse order for optimal keyword detection.
        for i, page in enumerate(reversed(pdf.pages)):
            text = page.extract_text()
            if text:
                for keyword in keywords:
                    if keyword in text and keyword not in keyword_page_map:
                        # Record the page number where the keyword is found.
                        keyword_page_map[keyword] = pdf_pages_len - 1 - i

    return keyword_page_map


def merge_pdfs_in_folder(folder_path: Path, output_pdf_path: Path) -> None:
    """
    Merges all PDFs in a specified folder into a single output PDF.

    :param Path folder_path: Path to the folder containing the PDFs to merge.
    :param Path output_pdf_path: Path where the merged PDF will be saved.
    """
    merger = pypdf.PdfWriter()
    folder = Path(folder_path)
    pdf_files = sorted(folder.glob("*.pdf"))

    if not pdf_files:
        return

    # Iterate through the sorted PDF files and merge them.
    for pdf_file in pdf_files:
        merger.append(pdf_file)

    # Write the merged content to the output PDF.
    with open(output_pdf_path, 'wb') as output_file:
        merger.write(output_file)
    merger.close()
