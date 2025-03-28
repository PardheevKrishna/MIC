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

# Worker thread to process Excel files asynchronously.
class SearchWorker(QThread):
    # Signal emits: file_path, sheet_name, row_index, row_data (tuple: (columns, values))
    resultFound = pyqtSignal(str, str, int, object)
    finished = pyqtSignal()
    
    def __init__(self, folder_path, search_term, parent=None):
        super().__init__(parent)
        self.folder_path = folder_path
        self.search_term = search_term.strip()
        self.search_term_lower = self.search_term.lower()  # Pre-compute lowercase term

    def run(self):
        print(f"Search started for term: {self.search_term}")

        # Get list of Excel files (.xlsx and .xls)
        file_list = (glob.glob(os.path.join(self.folder_path, "*.xlsx")) + 
                     glob.glob(os.path.join(self.folder_path, "*.xls")))
        print(f"Found {len(file_list)} files.")
        
        for file_path in file_list:
            try:
                print(f"Reading file: {file_path}")
                # Load all sheets as a dict: {sheet_name: DataFrame}
                sheets = pd.read_excel(file_path, sheet_name=None)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

            for sheet_name, df in sheets.items():
                print(f"Processing sheet: {sheet_name} in file: {os.path.basename(file_path)}")
                # Iterate over each row in the DataFrame
                for idx, row in df.iterrows():
                    row_matched = False
                    for col in df.columns:
                        cell_value = row[col]
                        if pd.isna(cell_value):
                            continue
                        cell_str = str(cell_value)
                        similarity = fuzz.ratio(self.search_term_lower, cell_str.lower())
                        if similarity >= 80:
                            row_matched = True
                            break
                    if row_matched:
                        # Save the entire row as a tuple: (columns, values)
                        row_data = (df.columns.tolist(), row.tolist())
                        self.resultFound.emit(file_path, sheet_name, idx, row_data)
        print("Search completed.")
        self.finished.emit()


class ExcelSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Search Application")
        self.resize(1000, 600)
        self.folder_path = ""
        self.worker = None  # Will hold the worker thread

        # Dictionaries for tree lookup.
        self.file_items = {}   # Maps file_path to file-level QTreeWidgetItem.
        self.sheet_items = {}  # Maps (file_path, sheet_name) to sheet-level QTreeWidgetItem.

        self.initUI()

    def initUI(self):
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

        # Splitter divides tree view (left) and detail view (right)
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)

        # Left: hierarchical tree view (File -> Sheet -> Row)
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["File", "Sheet", "Row"])
        self.result_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.result_tree.itemSelectionChanged.connect(self.on_item_selected)
        self.result_tree.itemDoubleClicked.connect(self.on_item_double_clicked)
        splitter.addWidget(self.result_tree)

        # Right: detail view (aggregated view)
        self.detail_table = QTableWidget()
        self.detail_table.setEditTriggers(QTableWidget.NoEditTriggers)
        splitter.addWidget(self.detail_table)
        splitter.setSizes([400, 600])

        # Modern styling.
        self.setStyleSheet("""
            QMainWindow { background-color: #1e1e2f; font-family: 'Segoe UI', sans-serif; }
            QLabel { font-size: 14px; color: #ffffff; }
            QLineEdit, QPushButton {
                font-size: 16px; padding: 8px; border-radius: 8px;
                border: 1px solid #3c3c4e; background-color: #27293d; color: #ffffff;
            }
            QLineEdit:hover, QPushButton:hover { border: 1px solid #5c5c7e; }
            QTreeWidget, QTableWidget { background-color: #27293d; color: #ffffff; border: none; }
            QHeaderView::section {
                background-color: #3c3c4e; color: #ffffff; padding: 8px; border: 1px solid #27293d;
            }
            QTreeWidget::item { padding: 4px; }
            QTreeWidget::item:selected, QTableWidget::item:selected { background-color: #4a90e2; }
        """)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder with Excel Files")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"Folder: {folder}")

    def start_search(self):
        search_term = self.search_input.text().strip()
        if not search_term or not self.folder_path:
            return

        self.result_tree.clear()
        self.detail_table.clear()
        self.file_items.clear()
        self.sheet_items.clear()
        self.search_button.setEnabled(False)

        self.worker = SearchWorker(self.folder_path, search_term)
        self.worker.resultFound.connect(self.handle_result)
        self.worker.finished.connect(self.search_finished)
        self.worker.start()

    def handle_result(self, file_path, sheet_name, row_index, row_data):
        print(f"Row matched in file: {os.path.basename(file_path)}, sheet: {sheet_name}, row: {row_index}")

        file_key = file_path
        if file_key not in self.file_items:
            file_item = QTreeWidgetItem([os.path.basename(file_path), "", ""])
            file_item.setData(0, Qt.UserRole, file_path)
            self.file_items[file_key] = file_item
            self.result_tree.addTopLevelItem(file_item)
        else:
            file_item = self.file_items[file_key]

        sheet_key = (file_path, sheet_name)
        if sheet_key not in self.sheet_items:
            sheet_item = QTreeWidgetItem(["", sheet_name, ""])
            self.sheet_items[sheet_key] = sheet_item
            file_item.addChild(sheet_item)
        else:
            sheet_item = self.sheet_items[sheet_key]

        row_item = QTreeWidgetItem(["", "", f"Row {row_index}"])
        row_item.setData(0, Qt.UserRole + 1, row_data)
        sheet_item.addChild(row_item)

    def search_finished(self):
        print("Search finished.")
        self.search_button.setEnabled(True)
        self.result_tree.expandAll()

    def on_item_selected(self):
        selected_items = self.result_tree.selectedItems()
        if not selected_items:
            return
        item = selected_items[0]
        row_data = item.data(0, Qt.UserRole + 1)
        if row_data is not None:
            # Show individual row detail if a row-level node is selected.
            columns, values = row_data
            self.show_detail(columns, values)
        elif item.parent() is not None:
            # If a sheet-level node is selected, aggregate results grouped by file.
            sheet_name = item.text(1)
            print(f"Aggregating all rows for sheet: {sheet_name}")
            aggregated_records = []
            for (file_path, s_name), sheet_item in self.sheet_items.items():
                if s_name == sheet_name:
                    for i in range(sheet_item.childCount()):
                        row_item = sheet_item.child(i)
                        r_data = row_item.data(0, Qt.UserRole + 1)
                        if r_data is not None:
                            row_text = row_item.text(2)
                            # Append tuple: (source, row_text, columns, values)
                            aggregated_records.append((os.path.basename(file_path), row_text, r_data[0], r_data[1]))
            self.show_aggregated_detail(aggregated_records)
        else:
            self.detail_table.clear()
            self.detail_table.setRowCount(0)
            self.detail_table.setColumnCount(0)

    def show_detail(self, columns, values):
        # Show a single row's details in a multi-column table.
        self.detail_table.clear()
        self.detail_table.setColumnCount(len(columns))
        self.detail_table.setRowCount(1)
        self.detail_table.setHorizontalHeaderLabels(columns)
        for col, val in enumerate(values):
            self.detail_table.setItem(0, col, QTableWidgetItem(str(val)))
        self.detail_table.resizeColumnsToContents()

    def show_aggregated_detail(self, aggregated_records):
        """
        Display aggregated rows from matching sheets grouped by file.
        For each file (source) group:
          - First display a header row with the file name.
          - Next, display a row with that sheet's column names.
          - Then, display each matching row (with row number and values).
          - A blank row is inserted between file groups.
        The table is set to have one column; each row's text is built accordingly.
        """
        if not aggregated_records:
            self.detail_table.clear()
            self.detail_table.setRowCount(0)
            self.detail_table.setColumnCount(0)
            return

        # Group records by source.
        groups = {}
        for record in aggregated_records:
            source, row_text, cols, values = record
            groups.setdefault(source, []).append((row_text, cols, values))

        detail_lines = []
        # For each file group, add header and then the rows.
        for source in sorted(groups.keys()):
            group_rows = groups[source]
            # Assume the header (column names) from the first row.
            header_cols = group_rows[0][1]
            detail_lines.append(f"File: {source}")
            detail_lines.append("Columns: " + ", ".join(header_cols))
            for row_text, cols, values in group_rows:
                detail_lines.append(f"{row_text}: " + ", ".join(str(v) for v in values))
            detail_lines.append("")  # Blank line between groups

        # Remove trailing blank line if any.
        if detail_lines and detail_lines[-1] == "":
            detail_lines.pop()

        # Set the detail_table to have one column.
        self.detail_table.clear()
        self.detail_table.setColumnCount(1)
        self.detail_table.setRowCount(len(detail_lines))
        self.detail_table.setHorizontalHeaderLabels(["Details"])
        for i, line in enumerate(detail_lines):
            self.detail_table.setItem(i, 0, QTableWidgetItem(line))
        self.detail_table.resizeColumnsToContents()

    def on_item_double_clicked(self, item, column):
        # Double-clicking a file-level node opens the file.
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
            print(f"Error opening file {file_path}: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelSearchApp()
    window.show()
    sys.exit(app.exec_())