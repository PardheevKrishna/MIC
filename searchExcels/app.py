import sys
import os
import glob
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLineEdit, QPushButton,
    QTreeWidget, QTreeWidgetItem, QLabel, QFileDialog
)
from PyQt5.QtCore import Qt
from rapidfuzz import fuzz

class ExcelSearchApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Search Application")
        self.resize(800, 600)
        self.folder_path = ""  # Will store the folder containing Excel files
        self.initUI()

    def initUI(self):
        # Main widget and layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Folder selection: display the current folder and add a button to select one.
        self.folder_label = QLabel("No folder selected")
        self.folder_button = QPushButton("Select Folder")
        self.folder_button.clicked.connect(self.select_folder)
        layout.addWidget(self.folder_label)
        layout.addWidget(self.folder_button)

        # Search input field
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Enter search term...")
        layout.addWidget(self.search_input)

        # Search button that triggers the search logic
        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_excel_files)
        layout.addWidget(self.search_button)

        # Results display using a hierarchical tree widget
        self.result_tree = QTreeWidget()
        self.result_tree.setHeaderLabels(["File", "Sheet", "Row", "Cell Value"])
        layout.addWidget(self.result_tree)

        # Optional: set a sleek, modern style (customize as desired)
        self.setStyleSheet("""
            QMainWindow { background-color: #2c3e50; color: #ecf0f1; }
            QLabel { font-size: 14px; }
            QLineEdit, QPushButton { font-size: 16px; padding: 8px; border-radius: 4px; }
            QLineEdit { background-color: #34495e; color: #ecf0f1; }
            QPushButton { background-color: #2980b9; color: #ecf0f1; }
            QTreeWidget { background-color: #34495e; color: #ecf0f1; }
        """)

    def select_folder(self):
        # Let the user select a folder containing Excel files
        folder = QFileDialog.getExistingDirectory(self, "Select Folder with Excel Files")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(f"Folder: {folder}")

    def search_excel_files(self):
        search_term = self.search_input.text().strip()
        if not search_term:
            return  # No search term entered

        if not self.folder_path:
            return  # No folder selected

        self.result_tree.clear()

        # Get all Excel files in the folder
        file_list = glob.glob(os.path.join(self.folder_path, "*.xlsx")) + \
                    glob.glob(os.path.join(self.folder_path, "*.xls"))
        
        # Iterate over each file
        for file_path in file_list:
            file_item = QTreeWidgetItem([os.path.basename(file_path), "", "", ""])
            try:
                # Load all sheets of the Excel file; returns a dict of {sheet_name: DataFrame}
                sheets = pd.read_excel(file_path, sheet_name=None)
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
                continue

            # Process each sheet in the file
            for sheet_name, df in sheets.items():
                sheet_item = QTreeWidgetItem(["", sheet_name, "", ""])
                found = False

                # Iterate over each row in the DataFrame
                for idx, row in df.iterrows():
                    # For each cell in the row, check for a fuzzy match
                    for col in df.columns:
                        cell_value = row[col]
                        # Skip NaN or empty values
                        if pd.isna(cell_value):
                            continue
                        cell_str = str(cell_value)
                        similarity = fuzz.ratio(search_term.lower(), cell_str.lower())
                        # If the similarity is at least 80%, record this match
                        if similarity >= 80:
                            # Each match is added as a child item under the sheet node
                            row_item = QTreeWidgetItem(["", "", str(idx), cell_str])
                            sheet_item.addChild(row_item)
                            found = True
                            # Uncomment the following break if you want only one match per row:
                            # break

                # Only add the sheet to the file node if any match was found
                if found:
                    file_item.addChild(sheet_item)

            # Only add the file to the main tree if it contains any matching sheets
            if file_item.childCount() > 0:
                self.result_tree.addTopLevelItem(file_item)

        # Expand all nodes to show the hierarchy
        self.result_tree.expandAll()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ExcelSearchApp()
    window.show()
    sys.exit(app.exec_())