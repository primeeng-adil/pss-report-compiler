import os
import pickle
from pathlib import Path
from typing import Callable
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QDialog, QMessageBox
from add_more import AddMoreWindow
from interface_ui import Ui_MainWindow
from progress_ui import Ui_Dialog
from worker import Worker

SC_NAME = 'Short Circuit'
COOR_NAME = 'Coordination'
UTIL_NAME = 'Utility'
REF_NAME = 'Reference'
SORT_SC_KEY = ['PRES', 'ULT', 'GEN']
ICON_NAME = 'icon.ico'
LOGO_NAME = 'logo_title.png'
DATA_NAME = 'data.pickle'
DEFAULTS_NAME = 'defaults.txt'
INPUT_ERROR_TITLE = 'Input Error'
INPUT_ERROR_MSG = 'Data Validation Failed'
INPUT_ERROR_DESC = ('The entered path to the Word document does not exist. '
                    'Please select a valid path from the Browse option.')


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
        self.app_path = app_path
        self.default_path = Path.home() / 'Documents' / 'PRC Data'
        self.icon = self._get_icon()
        self.additional_data = {}
        self.headings = {}
        self.default_dirs = {}
        self._get_defaults()
        self.add_more_window = None
        self.progress_dialog = None
        self.worker = None
        self.setupUi(self)
        self._set_logo()
        self._set_window_icon()
        self._connect_buttons()
        self._add_connections()
        self.show()

    def _get_defaults(self):
        defaults_file = self.default_path / DEFAULTS_NAME
        with open(defaults_file, 'r', encoding='UTF-8') as f:
            for line in f.readlines():
                _name, _dir, _heading = line.strip().split(',')
                self.headings.update({_name: _heading.strip()})
                self.default_dirs.update({_name: _dir.strip()})

    def _set_window_icon(self) -> None:
        """
        Sets the window icon for the application using the stored icon.
        """
        self.setWindowIcon(self.icon)

    def _get_icon(self) -> QIcon:
        """
        Retrieves the application icon.

        :return: A QIcon object representing the application icon.
        :rtype: QIcon
        """
        icon_path = str(Path(self.app_path, './res', ICON_NAME))
        return QIcon(icon_path)

    def _set_logo(self) -> None:
        """
        Sets the logo image for the application's title widget.
        """
        logo_path = str(Path(self.app_path, './res', LOGO_NAME))
        self.title.setPixmap(QPixmap(logo_path))

    def _initialize_add_more(self):
        """
        Initializes the Add More window and connects its signals.
        """
        self.add_more_window = AddMoreWindow(self.icon, self.additional_data)
        self.add_more_window.window_closed.connect(self._handle_add_more_close)
        self.add_more_window.window_confirmed.connect(self._handle_add_more_confirm)
        self.add_more_window.setWindowModality(Qt.ApplicationModal)
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
        self.util_browse_btn.clicked.connect(lambda: self._show_directory_browser(self.util_input))
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
            (self.util_browse_btn, self.util_input),
            (self.ref_browse_btn, self.ref_input)
        ]:
            button.setDisabled(True)
            input_field.setDisabled(True)

    def _execute_worker_thread(self):
        """
        Prepares paths and initializes a Worker thread to process PDF generation.
        """
        if not self._verify_inputs():
            return
        report_doc_path = Path(self.report_input.text())
        final_pdf_path = report_doc_path.with_suffix('.pdf')
        insert_pdfs = self._prepare_insert_pdfs()

        self.worker = Worker(report_doc_path, final_pdf_path, insert_pdfs)
        self.worker.error_occurred.connect(self._handle_error)
        self.worker.process_finished.connect(self._handle_finished)
        self._show_progress_dialog()
        self.worker.start()

    def _verify_inputs(self):
        report_path = Path(self.report_input.text())
        if not self.report_input.text() or not report_path.is_file():
            self._show_error_msg(INPUT_ERROR_TITLE, INPUT_ERROR_MSG, INPUT_ERROR_DESC)
            return False
        return True

    def _prepare_insert_pdfs(self) -> dict[str, list[Path]]:
        """
        Prepares a dictionary of PDFs to be inserted based on user selections.

        :return: A dictionary with section titles as keys and lists of PDF paths as values.
        :rtype: dict[str, list[Path]]
        """
        insert_pdfs = {}
        self._add_section_pdfs(self.headings[SC_NAME], self.sc_input, insert_pdfs, self._sc_sorting_func)
        self._add_section_pdfs(self.headings[COOR_NAME], self.coor_input, insert_pdfs)
        self._add_section_pdfs(self.headings[UTIL_NAME], self.util_input, insert_pdfs)
        self._add_section_pdfs(self.headings[REF_NAME], self.ref_input, insert_pdfs)
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
            lambda: self._set_default_path(self.sc_default_cb, self.sc_input, self.default_dirs[SC_NAME]))
        self.coor_default_cb.toggled.connect(
            lambda: self._set_default_path(self.coor_default_cb, self.coor_input, self.default_dirs[COOR_NAME]))
        self.util_default_cb.toggled.connect(
            lambda: self._set_default_path(self.util_default_cb, self.util_input, self.default_dirs[UTIL_NAME]))
        self.ref_default_cb.toggled.connect(
            lambda: self._set_default_path(self.ref_default_cb, self.ref_input, self.default_dirs[REF_NAME]))

        data_filepath = self.default_path / DATA_NAME
        defaults_filepath = self.default_path / DEFAULTS_NAME
        self.action_defaults.triggered.connect(lambda: os.startfile(defaults_filepath))
        self.action_save_inputs.triggered.connect(lambda: self.save_data(data_filepath))
        self.action_load_inputs.triggered.connect(lambda: self.load_data(data_filepath))
        self.action_save_as_inputs.triggered.connect(self.save_as_data)
        self.action_load_as_inputs.triggered.connect(self.load_as_data)
        self.action_exit.triggered.connect(self.close)

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
        self._set_default_path(self.sc_default_cb, self.sc_input, self.default_dirs[SC_NAME])
        self._set_default_path(self.coor_default_cb, self.coor_input, self.default_dirs[COOR_NAME])
        self._set_default_path(self.util_default_cb, self.util_input, self.default_dirs[UTIL_NAME])
        self._set_default_path(self.ref_default_cb, self.ref_input, self.default_dirs[REF_NAME])

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

    def _show_progress_dialog(self) -> None:
        """
        Creates and displays a progress dialog until the process is completed.
        """
        self.progress_dialog = Dialog(self.icon)
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.show()

    def _get_save_obj(self):
        report_doc = self.report_input.text()
        sc_input = [self.sc_input.text(), self.sc_default_cb.isChecked()]
        coor_input = [self.coor_input.text(), self.coor_default_cb.isChecked()]
        util_input = [self.util_input.text(), self.util_default_cb.isChecked()]
        ref_input = [self.ref_input.text(), self.ref_default_cb.isChecked()]
        return [report_doc, sc_input, coor_input, util_input, ref_input, self.additional_data]

    def _set_load_obj(self, data_obj):
        self.report_input.setText(data_obj[0])
        self.sc_default_cb.setChecked(data_obj[1][1])
        self.coor_default_cb.setChecked(data_obj[2][1])
        self.util_default_cb.setChecked(data_obj[3][1])
        self.ref_default_cb.setChecked(data_obj[4][1])
        self.sc_input.setText(data_obj[1][0])
        self.coor_input.setText(data_obj[2][0])
        self.util_input.setText(data_obj[3][0])
        self.ref_input.setText(data_obj[4][0])
        self.additional_data = data_obj[-1]

    def save_data(self, save_file):
        try:
            save_file.parent.mkdir(exist_ok=True)
            with open(save_file, 'wb') as f:
                save_obj = self._get_save_obj()
                pickle.dump(save_obj, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as ex:
            self._show_error_msg('Save Error', 'Unable to Save', str(ex))

    def load_data(self, load_file):
        try:
            with open(load_file, 'rb') as f:
                data_obj = pickle.load(f)
                self._set_load_obj(data_obj)
        except Exception as ex:
            self._show_error_msg('Load Error', 'Unable to Load', str(ex))

    def save_as_data(self):
        dir_path = QFileDialog().getExistingDirectory(self)
        filepath = Path(dir_path) / DATA_NAME
        if dir_path and filepath.parent.is_dir():
            self.save_data(filepath)

    def load_as_data(self):
        filepath, _ = QFileDialog().getOpenFileName(self, 'Open', None, 'PICKLE (*.pickle)')
        if filepath and Path(filepath).is_file():
            self.load_data(filepath)


class Dialog(QDialog, Ui_Dialog):
    """
    Custom dialog for showing progress during long-running tasks.
    """

    def __init__(self, icon: QIcon, *args, **kwargs):
        """
        Initializes the progress dialog with the given icon.

        :param QIcon icon: Icon for the dialog.
        :param args: Additional positional arguments.
        :param kwargs: Additional keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.setWindowIcon(icon)
