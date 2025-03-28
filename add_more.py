from pathlib import Path
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QFileDialog, QMessageBox, QListWidget
from add_more_ui import Ui_Form

# Constants for messages and titles
ERROR_TITLE = "Input Error"
ERROR_DIR_NOT_FOUND = "Directory Not Found"
ERROR_DIR_ALREADY_EXISTS = "Directory Already Exists"
ERROR_DIR_NOT_FOUND_DESC = "The entered path is not a directory. Please use the browse button to select a directory."
ERROR_DIR_ALREADY_EXISTS_DESC = "The selected directory has already been added. Please choose a different directory."


class AddMoreWindow(QWidget, Ui_Form):
    window_closed = pyqtSignal()
    window_confirmed = pyqtSignal()

    def __init__(self, icon: QIcon, initial_data: dict[str, str], *args, **kwargs):
        """
        Initialize the AddMoreWindow with icon and initial data.

        :param QIcon icon: Window icon.
        :param dict initial_data: Initial data with headings as keys and paths as values.
        """
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.icon = icon
        self.setWindowIcon(self.icon)
        self._add_btn_connections()
        self._add_selection_rules()
        self._initialize_with_data(initial_data)

    def _initialize_with_data(self, initial_data: dict[str, str]):
        """
        Populate the lists with initial data.

        :param dict initial_data: A dictionary of headings and directory paths.
        """
        for heading, path in initial_data.items():
            self.heading_list.addItem(heading)
            self.dir_list.addItem(path)

    def _add_btn_connections(self):
        """Connect button signals to their respective functions."""
        self.add_heading_btn.clicked.connect(self.add_heading)
        self.add_dir_btn.clicked.connect(self.add_directory)
        self.browse_btn.clicked.connect(self.browse_dir)
        self.remove_btn.clicked.connect(self.remove_items)
        self.confirm_btn.clicked.connect(self.confirm)
        self.cancel_btn.clicked.connect(self.cancel)
        self.move_up_btn.clicked.connect(self._handle_move_up)
        self.move_down_btn.clicked.connect(self._handle_move_down)

    def _add_selection_rules(self):
        """Only allow mutually exclusive selection between both list widgets."""
        self.dir_list.clicked.connect(self.heading_list.clearSelection)
        self.heading_list.clicked.connect(self.dir_list.clearSelection)

    def _handle_move_up(self):
        """Handle moving selected items up in their respective lists."""
        self._move_selected_item(self.heading_list, up=True)
        self._move_selected_item(self.dir_list, up=True)

    def _handle_move_down(self):
        """Handle moving selected items down in their respective lists."""
        self._move_selected_item(self.heading_list, up=False)
        self._move_selected_item(self.dir_list, up=False)

    @staticmethod
    def _move_selected_item(list_widget: QListWidget, up: bool):
        """
        Move the selected item up or down in the given list widget.

        :param QListWidget list_widget: The list widget containing the items.
        :param bool up: If True, moves the item up; otherwise, moves it down.
        """
        if not list_widget.selectedItems():
            return

        current_row = list_widget.currentRow()
        if current_row == -1:
            return

        target_row = current_row - 1 if up else current_row + 1
        if 0 <= target_row < list_widget.count():
            current_item = list_widget.takeItem(current_row)
            list_widget.insertItem(target_row, current_item)
            list_widget.setCurrentRow(target_row)

    def cancel(self):
        """Emit a signal when the window is closed."""
        self.window_closed.emit()

    def confirm(self):
        """
        Emit a signal when the confirmation button is clicked.

        :raises ValueError: If the number of headings does not match the number of paths.
        """
        try:
            if self.heading_list.count() != self.dir_list.count():
                raise ValueError("The number of headings must match the number of paths.")
        except ValueError as ve:
            self._show_error_msg('Input Error', 'Data Validation Error', str(ve))
            return
        self.window_confirmed.emit()

    def add_heading(self):
        """Add a new heading from the input field to the heading list."""
        heading = self.heading_input.text().strip()
        if heading:
            self.heading_list.addItem(heading)
        self.heading_input.clear()

    def add_directory(self):
        """
        Add a directory from the input field to the directory list.
        Checks if the directory is valid and not already added.
        """
        directory = self.dir_input.text().strip()
        if not self._is_valid_directory(directory):
            self._show_error_msg(ERROR_TITLE, ERROR_DIR_NOT_FOUND, ERROR_DIR_NOT_FOUND_DESC)
            return
        if self._is_directory_duplicate(directory):
            self._show_error_msg(ERROR_TITLE, ERROR_DIR_ALREADY_EXISTS, ERROR_DIR_ALREADY_EXISTS_DESC)
            return
        self.dir_list.addItem(directory)
        self.dir_input.clear()

    @staticmethod
    def _is_valid_directory(directory: str) -> bool:
        """
        Validate if the provided path is a valid directory.

        :param str directory: The path to validate.
        :return: True if the path is a valid directory, False otherwise.
        :rtype: bool
        """
        return directory and Path(directory).is_dir()

    def _is_directory_duplicate(self, directory: str) -> bool:
        """
        Check if the provided directory is already in the list.

        :param str directory: The directory path to check.
        :return: True if the directory already exists in the list, False otherwise.
        :rtype: bool
        """
        return directory in (self.dir_list.item(i).text() for i in range(self.dir_list.count()))

    def _show_error_msg(self, title: str, text: str, desc: str):
        """
        Show an error message box with the specified title, text, and description.

        :param str title: The title of the message box.
        :param str text: The main text of the message box.
        :param str desc: Additional description for the error.
        """
        msg = QMessageBox(self)
        msg.setWindowIcon(self.icon)
        msg.setIcon(QMessageBox.Critical)
        msg.setInformativeText(desc)
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.exec_()

    def browse_dir(self):
        """Open a dialog to select a directory and set its path in the input field."""
        dir_path = QFileDialog.getExistingDirectory(self)
        if dir_path:
            self.dir_input.setText(dir_path)

    def remove_items(self):
        """Remove selected items from the heading and directory lists."""
        self._remove_selected_item(self.heading_list)
        self._remove_selected_item(self.dir_list)

    @staticmethod
    def _remove_selected_item(list_widget: QListWidget):
        """
        Remove the currently selected item from the given list widget.

        :param QListWidget list_widget: The list widget from which to remove the item.
        """
        if not list_widget.selectedItems():
            return

        current_row = list_widget.currentRow()
        if current_row != -1:
            list_widget.takeItem(current_row)

    def get_heading_path_mapping(self) -> dict[str, str]:
        """
        Get the mapping of headings to their corresponding directory paths.

        :return: A dictionary with headings as keys and paths as values.
        :rtype: dict[str, str]
        """
        return {
            self.heading_list.item(i).text().strip(): self.dir_list.item(i).text()
            for i in range(self.heading_list.count())
        }
