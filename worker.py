from PyQt5.QtCore import QThread, pyqtSignal
from pdf_inserter import create_pdf_from_word, find_keywords_in_pdf, insert_pdfs_at_keywords
from pathlib import Path


class Worker(QThread):
    """
    Worker class for performing PDF insertion tasks in a separate thread.

    Attributes:
        error_occurred (pyqtSignal): Signal emitted when an error occurs, carrying the error message.
        process_finished (pyqtSignal): Signal emitted when a process finishes, carrying the output path.
    """
    error_occurred = pyqtSignal(str)
    process_finished = pyqtSignal(Path)

    def __init__(self, word_path: Path, pdf_path: Path, insert_pdfs: dict[str, list[Path]], *args, **kwargs):
        """
        Initialize the Worker thread with paths and insertion details.

        :param Path word_path: The path to the Word document.
        :param Path pdf_path: The path where the output PDF will be saved.
        :param dict insert_pdfs: A dictionary with keywords as keys and PDF paths to be inserted as values.
        """
        super().__init__(*args, **kwargs)
        self.word_path = word_path
        self.pdf_path = pdf_path
        self.insert_pdfs = insert_pdfs

    def run(self):
        """
        Executes the PDF insertion process in a separate thread.
        Converts a Word document to PDF, finds keywords, and inserts other PDFs at specified keyword positions.
        Emits `process_finished` on success or `error_occurred` on failure.
        """
        try:
            self._create_pdf()
            keyword_page_map = self._find_keywords()
            self._insert_pdfs(keyword_page_map)
            self.process_finished.emit(self.pdf_path)
        except Exception as e:
            self.error_occurred.emit(str(e))

    def _create_pdf(self):
        """
        Converts the Word document to a PDF at the specified output path.

        :raises Exception: If PDF creation fails.
        """
        create_pdf_from_word(self.word_path, self.pdf_path)

    def _find_keywords(self) -> dict[str, int]:
        """
        Finds the page numbers for each keyword in the generated PDF.

        :return: A mapping of keywords to their respective page numbers.
        :rtype: dict[str, int]
        :raises Exception: If keyword search fails.
        """
        keywords = list(self.insert_pdfs.keys())
        return find_keywords_in_pdf(self.pdf_path, keywords)

    def _insert_pdfs(self, keyword_page_map: dict[str, int]):
        """
        Inserts PDFs at the locations specified by the keyword page map.

        :param dict keyword_page_map: A mapping of keywords to their respective page numbers.
        :raises Exception: If PDF insertion fails.
        """
        insert_pdfs_at_keywords(self.pdf_path, self.insert_pdfs, keyword_page_map, self.pdf_path)
