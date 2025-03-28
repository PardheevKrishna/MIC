import sys
import os
import glob
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QTreeWidget, QTreeWidgetItem, QLabel,
    QFileDialog, QSplitter, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from rapidfuzz import fuzz
import subprocess

# Set this to True to enable debug print statements.
DEBUG = True

# Worker thread to process Excel files asynchronously.
class SearchWorker(QThread):
    # Signal to emit when a row match is found.
    # It sends: file_path, sheet_name, row_index, row_data (tuple: (columns, values))
    resultFound = pyqtSignal(str, str, int, object)
    finished = pyqtSignal()
    
    def __init__(self, folder_path, search_term, parent=None):
        super().__init__(parent)
        self.folder_path = folder_path
        self.search_term = search_term.strip()
        self.search_term_lower = self.search_term.lower()  # Lowercase once for reuse

    def run(self):
        if DEBUG:
            print(f"DEBUG: Search started for term: {self.search_term}")

        # Get list of Excel files (.xlsx and .xls)
        file_list = glob.glob(os.path.join(self.folder_path, "*.xlsx")) + \
                    glob.glob(os.path.join(self.folder_path, "*.xls"))
        if DEBUG:
            print(f"DEBUG: Found {len(file_list)} files in folder: {self.folder_path}")
        
        for file_path in file_list:
            if DEBUG:
                print(f"DEBUG: Reading file: {file_path}")
            try:
                # Load all sheets as a dictionary {sheet_name: DataFrame}
                sheets = pd.read_excel(file_path, sheet_name=None)
            except Exception as e:
                print(f"ERROR: Could not read {file_path}: {e}")
                continue

            # Process each sheet
            for sheet_name, df in sheets.items():
                if DEBUG:
                    print(f"DEBUG: Processing sheet: {sheet_name} in file: {file_path}")
                # Iterate over each row of the DataFrame
                for idx, row in df.iterrows():
                    row_matched = False
                    # Check every cell in the row for a fuzzy match.
                    for col in df.columns:
                        cell_value = row[col]
                        if pd.isna(cell_value):
                            continue
                        cell_str = str(cell_value)
                        similarity = fuzz.ratio(self.search_term_lower, cell_str.lower())
                        if similarity >= 80:
                            row_matched = True
                            break  # Once one cell qualifies, mark the entire row as matched.
                    if row_matched:
                        # Save entire row data as a tuple: (columns, values)
                        row_data = (df.columns.tolist(), row.tolist())
                        if DEBUG:
                            print(f"DEBUG: Match found in file: {os.path.basename(file_path)}, sheet: {sheet_name}, row: {idx}")
                        self.resultFound.emit(file_path, sheet_name, idx, row_data)

        if DEBUG:
            print("DEBUG: Search completed.")
        self.finished.emit()


class ExcelSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Search Application")
        self.resize(1000, 600)
        self.folder_path = ""
        self.worker = None  # Placeholder for our worker thread

        # Dictionaries to store tree nodes for quick lookup.
        self.file_items = {}   # key: file path, value: QTreeWidgetItem for file
        self.sheet_items = {}  # key: (file_path, sheet_name), value: QTreeWidgetItem for sheet

        self.initUI()

    def initUI(self):
        # Main widget and layout.
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Top controls: folder selection and search input.
        controls_layout = QHBoxLayout()

        self.folder_label = QLabel("No folder selected")
        controls_layout.addWidget(self.folder_label)

        self.folder_button = QPushButton("Select Folder")
        self.folder_button.clicked.connect(self.select_folder)
        controls_layout.addWidget(self.folder_button)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter search term...")
        controls_layout.addWidget(self.search_input)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.start_search)
        controls_layout.addWidget(self.search_button)

        main_layout.addLayout(controls_layout)

        # Use a splitter to divide the tree view (left) and detail view (right)
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # Left side: hierarchical tree view
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["File", "Sheet", "Row"])
        self.result_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.result_tree.itemSelectionChanged.connect(self.on_item_selected)
        self.result_tree.itemDoubleClicked.connect(self.on_item_double_clicked)
        splitter.addWidget(self.result_tree)

        # Right side: detailed view of the row (as a table)
        self.detail_table = QTableWidget()
        self.detail_table.setEditTriggers(QTableWidget.NoEditTriggers)
        splitter.addWidget(self.detail_table)

        # Set initial splitter sizes
        splitter.setSizes([400, 600])

        # Modern, sleek styling with a refined dark palette and subtle hover effects.
        self.setStyleSheet("""
            /* Main window */
            QMainWindow {
                background-color: #1e1e2f;
                font-family: 'Segoe UI', sans-serif;
            }
            /* Labels */
            QLabel {
                font-size: 14px;
                color: #ffffff;
            }
            /* Line edits and buttons */
            QLineEdit, QPushButton {
                font-size: 16px;
                padding: 8px;
                border-radius: 8px;
                border: 1px solid #3c3c4e;
                background-color: #27293d;
                color: #ffffff;
            }
            QLineEdit:hover, QPushButton:hover {
                border: 1px solid #5c5c7e;
            }
            /* Tree and table views */
            QTreeWidget, QTableWidget {
                background-color: #27293d;
                color: #ffffff;
                border: none;
            }
            /* Header styling for tree and table */
            QHeaderView::section {
                background-color: #3c3c4e;
                color: #ffffff;
                padding: 8px;
                border: 1px solid #27293d;
            }
            /* Tree widget items */
            QTreeWidget::item {
                padding: 4px;
            }
            QTreeWidget::item:selected {
                background-color: #4a90e2;
            }
            /* Table widget items */
            QTableWidget::item:selected {
                background-color: #4a90e2;
            }
        """)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder with Excel Files")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"Folder: {folder}")
            if DEBUG:
                print(f"DEBUG: Folder selected: {folder}")

    def start_search(self):
        search_term = self.search_input.text().strip()
        if not search_term:
            if DEBUG:
                print("DEBUG: No search term entered.")
            return
        if not self.folder_path:
            if DEBUG:
                print("DEBUG: No folder selected.")
            return

        # Clear previous results
        self.result_tree.clear()
        self.detail_table.clear()
        self.file_items.clear()
        self.sheet_items.clear()

        # Disable search button while processing.
        self.search_button.setEnabled(False)
        if DEBUG:
            print("DEBUG: Starting search...")

        # Create and start the worker thread.
        self.worker = SearchWorker(self.folder_path, search_term)
        self.worker.resultFound.connect(self.handle_result)
        self.worker.finished.connect(self.search_finished)
        self.worker.start()

    def handle_result(self, file_path, sheet_name, row_index, row_data):
        # file_path: full file path
        # row_data: tuple (columns, values)
        print(f"DEBUG: Match found - File: {os.path.basename(file_path)}, Sheet: {sheet_name}, Row: {row_index}")

        # Create or get the file-level item.
        file_key = file_path
        if file_key not in self.file_items:
            file_item = QTreeWidgetItem([os.path.basename(file_path), "", ""])
            # Store the file path in the file item (for later use on double-click)
            file_item.setData(0, Qt.UserRole, file_path)
            self.file_items[file_key] = file_item
            self.result_tree.addTopLevelItem(file_item)
        else:
            file_item = self.file_items[file_key]

        # Create or get the sheet-level item.
        sheet_key = (file_path, sheet_name)
        if sheet_key not in self.sheet_items:
            sheet_item = QTreeWidgetItem(["", sheet_name, ""])
            self.sheet_items[sheet_key] = sheet_item
            file_item.addChild(sheet_item)
        else:
            sheet_item = self.sheet_items[sheet_key]

        # Create the row-level item. The text shows the row number.
        row_item = QTreeWidgetItem(["", "", f"Row {row_index}"])
        # Save the entire row data into the item (for detail view).
        row_item.setData(0, Qt.UserRole + 1, row_data)
        sheet_item.addChild(row_item)

    def search_finished(self):
        print("DEBUG: Search finished.")
        self.search_button.setEnabled(True)
        self.result_tree.expandAll()

    def on_item_selected(self):
        selected_items = self.result_tree.selectedItems()
        if not selected_items:
            return
        item = selected_items[0]
        # Check if this item has row data (stored under UserRole+1).
        row_data = item.data(0, Qt.UserRole + 1)
        if row_data:
            columns, values = row_data
            self.show_detail(columns, values)
        else:
            self.detail_table.clear()
            self.detail_table.setRowCount(0)
            self.detail_table.setColumnCount(0)

    def show_detail(self, columns, values):
        self.detail_table.clear()
        self.detail_table.setColumnCount(len(columns))
        self.detail_table.setRowCount(1)
        self.detail_table.setHorizontalHeaderLabels(columns)
        for col, val in enumerate(values):
            item = QTableWidgetItem(str(val))
            self.detail_table.setItem(0, col, item)
        self.detail_table.resizeColumnsToContents()

    def on_item_double_clicked(self, item, column):
        # If a file-level item is double-clicked, try to open the file.
        if item.parent() is None:
            file_path = item.data(0, Qt.UserRole)
            if file_path:
                self.open_file(file_path)

    def open_file(self, file_path):
        try:
            if sys.platform.startswith('win'):
                os.startfile(file_path)
            elif sys.platform.startswith('darwin'):
                subprocess.call(('open', file_path))
            else:
                subprocess.call(('xdg-open', file_path))
        except Exception as e:
            print(f"ERROR: Could not open file {file_path}: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelSearchApp()
    window.show()
    sys.exit(app.exec_())