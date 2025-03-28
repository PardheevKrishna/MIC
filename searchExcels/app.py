import sys
import os
import glob
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QPushButton,
    QTreeWidget, QTreeWidgetItem, QLabel, QFileDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from rapidfuzz import fuzz

# Worker thread to process Excel files asynchronously.
class SearchWorker(QThread):
    # Signal emits: file_path, sheet_name, row_index, row_data (tuple: (columns, values))
    resultFound = pyqtSignal(str, str, int, object)
    finished = pyqtSignal()
    
    def __init__(self, folder_path, search_term, parent=None):
        super().__init__(parent)
        self.folder_path = folder_path
        self.search_term = search_term.strip()
    
    def run(self):
        # Collect all Excel files (.xlsx and .xls)
        file_list = glob.glob(os.path.join(self.folder_path, "*.xlsx")) + \
                    glob.glob(os.path.join(self.folder_path, "*.xls"))
        
        for file_path in file_list:
            try:
                # Load all sheets as a dictionary {sheet_name: DataFrame}
                sheets = pd.read_excel(file_path, sheet_name=None)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

            for sheet_name, df in sheets.items():
                # Iterate over each row in the DataFrame
                for idx, row in df.iterrows():
                    row_matched = False
                    # Check every cell in the row for an 80% fuzzy match
                    for col in df.columns:
                        cell_value = row[col]
                        if pd.isna(cell_value):
                            continue
                        cell_str = str(cell_value)
                        similarity = fuzz.ratio(self.search_term.lower(), cell_str.lower())
                        if similarity >= 80:
                            row_matched = True
                            break  # One matching cell is enough for the row to qualify
                    if row_matched:
                        # Capture the entire row (as tuple: (column names, values))
                        row_data = (df.columns.tolist(), row.tolist())
                        self.resultFound.emit(file_path, sheet_name, idx, row_data)
        self.finished.emit()


class ExcelSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Search Application")
        self.resize(800, 600)
        self.folder_path = ""
        self.worker = None

        # Dictionaries for quick lookup of file and sheet nodes in the tree
        self.file_items = {}   # key: file path, value: QTreeWidgetItem for file
        self.sheet_items = {}  # key: (file_path, sheet_name), value: QTreeWidgetItem for sheet

        self.initUI()

    def initUI(self):
        # Main widget and vertical layout.
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Top controls: folder selection and search input.
        self.folder_label = QLabel("No folder selected")
        layout.addWidget(self.folder_label)
        self.folder_button = QPushButton("Select Folder")
        self.folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.folder_button)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter search term...")
        layout.addWidget(self.search_input)
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.start_search)
        layout.addWidget(self.search_button)

        # Tree view for displaying search results.
        # Using a one-column tree view to show the entire hierarchical structure.
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["Search Results"])
        layout.addWidget(self.result_tree)

        # Modern, sleek styling (customize as needed)
        self.setStyleSheet("""
            QMainWindow { background-color: #2c3e50; color: #ecf0f1; }
            QLabel { font-size: 14px; }
            QLineEdit, QPushButton { font-size: 16px; padding: 6px; border-radius: 4px; }
            QLineEdit { background-color: #34495e; color: #ecf0f1; }
            QPushButton { background-color: #2980b9; color: #ecf0f1; }
            QTreeWidget { background-color: #34495e; color: #ecf0f1; }
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

        # Clear previous results.
        self.result_tree.clear()
        self.file_items.clear()
        self.sheet_items.clear()
        self.search_button.setEnabled(False)

        # Create and start the worker thread.
        self.worker = SearchWorker(self.folder_path, search_term)
        self.worker.resultFound.connect(self.handle_result)
        self.worker.finished.connect(self.search_finished)
        self.worker.start()

    def handle_result(self, file_path, sheet_name, row_index, row_data):
        # Create or retrieve the file-level node.
        file_key = file_path
        if file_key not in self.file_items:
            file_item = QTreeWidgetItem([f"File: {os.path.basename(file_path)}"])
            file_item.setData(0, Qt.UserRole, file_path)  # Optional: store full file path
            self.file_items[file_key] = file_item
            self.result_tree.addTopLevelItem(file_item)
        else:
            file_item = self.file_items[file_key]

        # Create or retrieve the sheet-level node.
        sheet_key = (file_path, sheet_name)
        if sheet_key not in self.sheet_items:
            sheet_item = QTreeWidgetItem([f"Sheet: {sheet_name}"])
            self.sheet_items[sheet_key] = sheet_item
            file_item.addChild(sheet_item)
        else:
            sheet_item = self.sheet_items[sheet_key]

        # Create the row-level node.
        row_item = QTreeWidgetItem([f"Row {row_index}"])
        sheet_item.addChild(row_item)

        # Create a child node under the row to display all row values with column names.
        columns, values = row_data
        detail_text = " | ".join([f"{col}: {val}" for col, val in zip(columns, values)])
        detail_item = QTreeWidgetItem([detail_text])
        row_item.addChild(detail_item)

    def search_finished(self):
        self.search_button.setEnabled(True)
        self.result_tree.expandAll()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelSearchApp()
    window.show()
    sys.exit(app.exec_())