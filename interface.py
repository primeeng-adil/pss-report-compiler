import os
from pathlib import Path
from typing import Callable
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QDialog, QMessageBox
from add_more import AddMoreWindow
from interface_ui import Ui_MainWindow
from progress_ui import Ui_Dialog
from worker import Worker

SC_DIR_NAME = "SC"
COOR_DIR_NAME = "TCC"
UTILITY_DIR_NAME = "REF"
REF_DIR_NAME = "SLD"
SORT_SC_KEY = ['PRES', 'ULT', 'GEN']


class Interface(QMainWindow, Ui_MainWindow):
    """
    Main interface for the application, handling user interactions and executing PDF generation.
    """

    def __init__(self, app_path: Path, *args, **kwargs):
        """
        Initializes the interface with the given application path and optional arguments.

        :param Path app_path: Path to the application directory.
        :param args: Additional positional arguments.
        :param kwargs: Additional keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.additional_data = {}
        self.worker = None
        self.add_more_window = None
        self.icon_path = str(Path(app_path, './res', 'icon.ico'))
        self.setWindowIcon(QIcon(self.icon_path))
        self.progress_dialog = Dialog(self.icon_path)
        self.setupUi(self)
        self._set_logo(app_path)
        self._connect_buttons()
        self._add_connections()
        self.show()

    def _set_logo(self, app_path):
        logo_path = str(Path(app_path, './res', 'logo_title.png'))
        logo = QPixmap(logo_path)
        self.title.setPixmap(logo)

    def _initialize_add_more(self):
        """
        Initializes the Add More window and connects its signals.
        """
        self.add_more_window = AddMoreWindow(self.icon_path, self.additional_data)
        self.add_more_window.window_closed.connect(self._handle_add_more_close)
        self.add_more_window.window_confirmed.connect(self._handle_add_more_confirm)
        self.add_more_window.show()

    def _handle_add_more_close(self):
        """
        Handles the event when the Add More window is closed.
        """
        del self.add_more_window

    def _handle_add_more_confirm(self):
        """
        Updates additional data with the mapping from the Add More window and closes it.
        """
        self.additional_data = self.add_more_window.get_heading_path_mapping()
        del self.add_more_window

    def _connect_buttons(self) -> None:
        """
        Connects buttons in the UI to their respective event handlers.
        """
        self._toggle_initial_states()
        self.report_browse_btn.clicked.connect(self._show_file_browser_report)
        self.sc_browse_btn.clicked.connect(lambda: self._show_directory_browser(self.sc_input))
        self.coor_browse_btn.clicked.connect(lambda: self._show_directory_browser(self.coor_input))
        self.utility_browse_btn.clicked.connect(lambda: self._show_directory_browser(self.utility_input))
        self.ref_browse_btn.clicked.connect(lambda: self._show_directory_browser(self.ref_input))
        self.generate_btn.clicked.connect(self._execute_worker_thread)
        self.add_more_btn.clicked.connect(self._initialize_add_more)

    def _toggle_initial_states(self):
        """
        Sets the initial disabled states for browse buttons and input fields.
        """
        for button, input_field in [
            (self.sc_browse_btn, self.sc_input),
            (self.coor_browse_btn, self.coor_input),
            (self.utility_browse_btn, self.utility_input),
            (self.ref_browse_btn, self.ref_input)
        ]:
            button.setDisabled(True)
            input_field.setDisabled(True)

    def _execute_worker_thread(self):
        """
        Prepares paths and initializes a Worker thread to process PDF generation.
        """
        report_doc_path = Path(self.report_input.text())
        final_pdf_path = report_doc_path.with_suffix('.pdf')
        insert_pdfs = self._prepare_insert_pdfs()

        self.worker = Worker(report_doc_path, final_pdf_path, insert_pdfs)
        self.worker.error_occurred.connect(self._handle_error)
        self.worker.process_finished.connect(self._handle_finished)
        self.progress_dialog.show()
        self.worker.start()

    def _prepare_insert_pdfs(self) -> dict[str, list[Path]]:
        """
        Prepares a dictionary of PDFs to be inserted based on user selections.

        :return: A dictionary with section titles as keys and lists of PDF paths as values.
        :rtype: dict[str, list[Path]]
        """
        insert_pdfs = {}
        self._add_section_pdfs('– Short Circuit Results', self.sc_input, insert_pdfs, self._sc_sorting_func)
        self._add_section_pdfs('– Coordination Curves', self.coor_input, insert_pdfs)
        self._add_section_pdfs('– Utility Fault Data', self.utility_input, insert_pdfs)
        self._add_section_pdfs('– System Model Data', self.ref_input, insert_pdfs)
        insert_pdfs.update(self._get_additional_pdfs())
        return insert_pdfs

    @staticmethod
    def _sc_sorting_func(item: str) -> int:
        """
        Determines the sorting order of an item based on predefined keys.

        :param str item: The item to be sorted, typically a string representation of a file or path.
        :return: The index of the key found in the item, which dictates its sorting order.
                 If no key is found, returns the length of `SORT_SC_KEY` to place the item at the end.
        :rtype: int
        """
        for index, key in enumerate(SORT_SC_KEY):
            if key in str(item):
                return index
        return len(SORT_SC_KEY)

    @staticmethod
    def _add_section_pdfs(title: str, input_field, insert_pdfs: dict, sorting_func: Callable = None):
        """
        Adds PDFs for a specific section.

        :param str title: The title of the section.
        :param QLineEdit input_field: The input field containing the directory path.
        :param dict insert_pdfs: The dictionary to update with the section PDFs.
        :param Callable sorting_func: Function to sort with. Default is None.
        """
        dir_path = Path(input_field.text())
        insert_pdfs[title] = sorted(dir_path.glob("*.pdf"), key=sorting_func)

    def _get_additional_pdfs(self) -> dict[str, list[Path]]:
        """
        Retrieves additional PDFs from the additional data mapping.

        :return: A dictionary with headings as keys and sorted lists of PDF paths as values.
        :rtype: dict[str, list[Path]]
        """
        return {key: sorted(Path(value).glob("*.pdf")) for key, value in self.additional_data.items()}

    def _handle_error(self, error_message: str):
        """
        Handles errors emitted by the worker.

        :param str error_message: The error message emitted by the worker.
        """
        self.progress_dialog.hide()
        self._show_error_msg('Runtime Error', 'Compilation Failed', error_message)

    def _handle_finished(self, pdf_path: Path):
        """
        Handles successful completion of the worker process.

        :param Path pdf_path: The path to the generated PDF.
        """
        self.progress_dialog.hide()
        os.startfile(pdf_path)

    def _add_connections(self):
        """
        Adds connections for checkbox toggles and input text changes.
        """
        self.report_input.textChanged.connect(self._set_default_paths)
        self.sc_default_cb.toggled.connect(
            lambda: self._set_default_path(self.sc_default_cb, self.sc_input, SC_DIR_NAME))
        self.coor_default_cb.toggled.connect(
            lambda: self._set_default_path(self.coor_default_cb, self.coor_input, COOR_DIR_NAME))
        self.utility_default_cb.toggled.connect(
            lambda: self._set_default_path(self.utility_default_cb, self.utility_input, UTILITY_DIR_NAME))
        self.ref_default_cb.toggled.connect(
            lambda: self._set_default_path(self.ref_default_cb, self.ref_input, REF_DIR_NAME))

    def _set_default_path(self, checkbox, line, name):
        """
        Sets the default path for an input field based on a checkbox state.

        :param QCheckBox checkbox: The checkbox to monitor.
        :param QLineEdit line: The line edit field to update.
        :param str name: The folder name to set as default.
        """
        report_path = Path(self.report_input.text())
        if report_path.is_file() and checkbox.isChecked():
            line.setText(str(report_path.parent / name))

    def _set_default_paths(self):
        """
        Sets default paths for all sections based on the report input path.
        """
        self._set_default_path(self.sc_default_cb, self.sc_input, SC_DIR_NAME)
        self._set_default_path(self.coor_default_cb, self.coor_input, COOR_DIR_NAME)
        self._set_default_path(self.utility_default_cb, self.utility_input, UTILITY_DIR_NAME)
        self._set_default_path(self.ref_default_cb, self.ref_input, REF_DIR_NAME)

    def _show_file_browser_report(self) -> None:
        """
        Opens a file dialog to select a Word document for the report.
        """
        filepath, _ = QFileDialog().getOpenFileName(self, 'Open', None, 'Word Documents (*.docx)')
        self.report_input.setText(filepath)

    def _show_directory_browser(self, input_field) -> None:
        """
        Opens a directory browser and sets the selected directory to the provided input field.

        :param QLineEdit input_field: The input field to set the directory path.
        """
        dir_path = QFileDialog().getExistingDirectory(self)
        input_field.setText(dir_path)

    def _show_error_msg(self, title: str, text: str, desc: str):
        """
        Show an error message box with the specified title, text, and description.

        :param str title: The title of the message box.
        :param str text: The main text of the message box.
        :param str desc: Additional description for the error.
        """
        msg = QMessageBox(self)
        msg.setWindowIcon(QIcon(self.icon))
        msg.setIcon(QMessageBox.Critical)
        msg.setInformativeText(desc)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.exec_()


class Dialog(QDialog, Ui_Dialog):
    """
    Custom dialog for showing progress during long-running tasks.
    """

    def __init__(self, icon_path: str, *args, **kwargs):
        """
        Initializes the progress dialog with the given icon.

        :param str icon_path: Path to the icon for the dialog.
        :param args: Additional positional arguments.
        :param kwargs: Additional keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.setWindowIcon(QIcon(icon_path))
