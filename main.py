import sys
import os
import io
import pandas as pd
import code128
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QVBoxLayout, QWidget, QFileDialog, QMessageBox, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QLabel)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPixmap

class BarcodeApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # Window settings
        self.setWindowTitle("ToolsTrackr - Barcode Management App")
        self.setGeometry(100, 100, 500, 400)
        self.setStyleSheet("background-color: white;")

        # Main layout
        layout = QVBoxLayout()

        # Branding (Logo instead of text)
        self.logo_label = QLabel(self)

        # Get correct resource path for logo.png
        logo_path = self.get_resource_path("logo.png")

        # Load the logo image
        pixmap = QPixmap(logo_path)

        # Scale the logo to a more appropriate size for the app header
        self.logo_label.setPixmap(pixmap.scaled(300, 300, Qt.KeepAspectRatio))  # Width: 300px, Height: proportional
        self.logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.logo_label, alignment=Qt.AlignCenter)

        # Create Barcode button
        self.create_barcode_button = QPushButton("Generate a Database Barcodes")
        self.create_barcode_button.setFont(QFont("Arial", 12, QFont.Bold))
        self.create_barcode_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border: 2px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.create_barcode_button.setCursor(Qt.PointingHandCursor)  # Change cursor on hover
        self.create_barcode_button.clicked.connect(self.open_create_barcode_window)
        layout.addWidget(self.create_barcode_button)

        # Scan Barcode button
        self.scan_barcode_button = QPushButton("Scan Barcodes")
        self.scan_barcode_button.setFont(QFont("Arial", 12, QFont.Bold))
        self.scan_barcode_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border: 2px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.scan_barcode_button.setCursor(Qt.PointingHandCursor)  # Change cursor on hover
        self.scan_barcode_button.clicked.connect(self.scan_barcode)
        layout.addWidget(self.scan_barcode_button)

        # Central widget for the main window
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def get_resource_path(self, relative_path):
        """Returns the correct path for bundled or non-bundled resources."""
        if getattr(sys, 'frozen', False):  # If the app is bundled as a .exe
            return os.path.join(sys._MEIPASS, relative_path)  # Get resource from the temporary folder
        else:  # If the app is run as a script
            return os.path.join(os.path.dirname(__file__), relative_path)

    def open_create_barcode_window(self):
        self.create_window = QMainWindow(self)
        self.create_window.setWindowTitle("Create Barcodes")
        self.create_window.setGeometry(100, 100, 500, 400)
        self.create_window.setStyleSheet("background-color: white;")

        form_layout = QVBoxLayout()

        # Table to display the form-like data
        self.table = QTableWidget()
        self.table.setColumnCount(2)  # Item Code and Description
        self.table.setHorizontalHeaderLabels(['Item Code', 'Description'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setRowCount(1)  # Start with one row
        form_layout.addWidget(self.table)

        # Add Item button
        add_entry_button = QPushButton("Add Item")
        add_entry_button.setStyleSheet("""
            background-color: #ff9800;
            color: black;
            border: 2px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        add_entry_button.setCursor(Qt.PointingHandCursor)  # Change cursor on hover
        add_entry_button.clicked.connect(self.add_item)
        form_layout.addWidget(add_entry_button)

        # Save Barcodes button
        save_barcode_button = QPushButton("Save Barcodes in Excel")
        save_barcode_button.setStyleSheet("""
            background-color: #ff9800;
            color: black;
            border: 2px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        save_barcode_button.setCursor(Qt.PointingHandCursor)  # Change cursor on hover
        save_barcode_button.clicked.connect(self.save_barcodes)
        form_layout.addWidget(save_barcode_button)

        # Set layout for the window
        container = QWidget()
        container.setLayout(form_layout)
        self.create_window.setCentralWidget(container)
        self.create_window.show()

        # Connect the itemChanged signal to clear the placeholder text
        self.table.itemChanged.connect(self.clear_placeholder)

    def add_item(self):
        # Add a new row to the table
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        # Create an editable Item Code cell with a placeholder-like text
        item_code_item = QTableWidgetItem("Item Code")
        self.table.setItem(row_position, 0, item_code_item)

        # Create an editable Description cell with a placeholder-like text
        description_item = QTableWidgetItem("Description")
        self.table.setItem(row_position, 1, description_item)

        # Optionally, set focus to the first cell for easy editing
        self.table.setCurrentCell(row_position, 0)

    def clear_placeholder(self, item):
        # Check if the item has placeholder text and clear it
        if item.text() == "Item Code" or item.text() == "Description":
            item.setText("")  # Clear placeholder text when typing starts

    def save_barcodes(self):
        # Open dialog to select directory
        save_dir = QFileDialog.getExistingDirectory(self, "Select Directory to Save Excel")
        if not save_dir:
            QMessageBox.warning(self, "No Directory Selected", "Please select a directory.")
            return

        excel_file_path = os.path.join(save_dir, "database_with_barcodes.xlsx")

        # Create new workbook
        wb = Workbook()
        ws = wb.active
        ws.append(['Item Code', 'Description', 'Barcode'])

        # Loop through the table and collect item codes and descriptions
        for row in range(self.table.rowCount()):
            item_code = self.table.item(row, 0).text().strip()
            description = self.table.item(row, 1).text().strip()

            if item_code:
                ws.append([item_code, description])

                # Generate barcode
                barcode_image = code128.image(item_code)
                image_stream = io.BytesIO()
                barcode_image.save(image_stream, format='PNG')
                image_stream.seek(0)

                img = OpenpyxlImage(image_stream)
                cell_address = f'C{ws.max_row}'  # Place image in the third column
                img.anchor = cell_address
                ws.add_image(img)

                # Adjust row height for barcode
                ws.row_dimensions[ws.max_row].height = 90

        wb.save(excel_file_path)
        QMessageBox.information(self, "Success", f"Data saved to {excel_file_path}")

    def scan_barcode(self):
        # Get the path of scan.py correctly based on whether it's running as an .exe or a script
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        scan_script_path = os.path.join(base_path, 'scan.py')

        os.system(f'python "{scan_script_path}"')  # This calls the external script for scanning

# Application entry point
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())
