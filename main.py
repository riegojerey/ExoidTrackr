import sys
import os
import io
import pandas as pd
import code128
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, 
                             QVBoxLayout, QWidget, QFileDialog, QMessageBox, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QLabel, 
                             QLineEdit)
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
        self.scan_barcode_button.clicked.connect(self.open_scan_barcode_window)
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
        # Logic for creating barcode
        pass

    def open_scan_barcode_window(self):
        # This replaces the external scan.py functionality with internal logic
        self.scan_window = QMainWindow(self)
        self.scan_window.setWindowTitle("Scan Barcodes")
        self.scan_window.setGeometry(100, 100, 900, 600)

        form_layout = QVBoxLayout()

        # Logo and status label
        self.logo_label = QLabel(self)
        logo_path = self.get_resource_path("logo.png")
        self.logo_pixmap = QPixmap(logo_path).scaled(326, 250, Qt.KeepAspectRatio)
        self.logo_label.setPixmap(self.logo_pixmap)
        form_layout.addWidget(self.logo_label, alignment=Qt.AlignLeft)

        # Status label
        self.status_label = QLabel("Scan a barcode or enter item code", self)
        self.status_label.setFont(QFont("Arial", 16, QFont.Bold))
        self.status_label.setStyleSheet("border-radius: 15px; padding: 10px; background-color: #fff3e0;")
        form_layout.addWidget(self.status_label, alignment=Qt.AlignCenter)

        # Toggle button
        self.mode = "Check In"
        self.toggle_button = QPushButton("Switch to Check Out", self)
        self.toggle_button.setFont(QFont("Arial", 12, QFont.Bold))
        self.toggle_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.toggle_button.clicked.connect(self.toggle_mode)
        form_layout.addWidget(self.toggle_button)

        # Export to Excel button
        self.export_button = QPushButton("Export to Excel", self)
        self.export_button.setFont(QFont("Arial", 12, QFont.Bold))
        self.export_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.export_button.clicked.connect(self.export_to_excel)
        form_layout.addWidget(self.export_button)

        # Input field for item code
        self.item_code_entry = QLineEdit(self)
        self.item_code_entry.setFont(QFont("Arial", 12))
        self.item_code_entry.setStyleSheet("""
            color: black;
            background-color: white;
            border: 1px solid #ccc;
            border-radius: 10px;
            padding: 10px;
        """)
        self.item_code_entry.setPlaceholderText("Enter item code")
        self.item_code_entry.returnPressed.connect(self.process_item)
        form_layout.addWidget(self.item_code_entry)

        # Table widget for displaying items
        self.table = QTableWidget(self)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Item Code", "Description", "Status", "Quantity"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        form_layout.addWidget(self.table)

        # Set the layout for scan window
        scan_container = QWidget()
        scan_container.setLayout(form_layout)
        self.scan_window.setCentralWidget(scan_container)
        self.scan_window.show()

        # Data structures
        self.items = {}
        self.inventory_df = None

        # Load Excel data on startup
        self.load_excel_data()

    def toggle_mode(self):
        if self.mode == "Check In":
            self.mode = "Check Out"
            self.toggle_button.setText("Switch to Check In")
            self.status_label.setText("Scan a barcode or enter item code (Check Out mode)")
        else:
            self.mode = "Check In"
            self.toggle_button.setText("Switch to Check Out")
            self.status_label.setText("Scan a barcode or enter item code (Check In mode)")

    def load_excel_data(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "Open Excel File", "", "Excel Files (*.xlsx)")
        if not file_path:
            QMessageBox.critical(self, "Error", "No Excel file selected.")
            return

        try:
            self.inventory_df = pd.read_excel(file_path)
            if 'Item Code' not in self.inventory_df.columns or 'Description' not in self.inventory_df.columns:
                raise ValueError("Excel file must have 'Item Code' and 'Description' columns.")
            QMessageBox.information(self, "Success", "Excel file loaded successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")

    def process_item(self):
        item_code = self.item_code_entry.text().strip().lower()
        if not item_code:
            return

        match = self.inventory_df[self.inventory_df['Item Code'].str.lower() == item_code]
        if not match.empty:
            description = match['Description'].values[0]
            if self.mode == "Check In":
                if item_code in self.items:
                    if self.items[item_code]['Status'] == "Checked Out":
                        self.items[item_code]['Status'] = "Checked In"
                    else:
                        self.items[item_code]['Quantity'] += 1
                else:
                    self.items[item_code] = {'Description': description, 'Status': "Checked In", 'Quantity': 1}
            else:
                if item_code in self.items:
                    if self.items[item_code]['Status'] == "Checked In":
                        self.items[item_code]['Status'] = "Checked Out"
                    else:
                        self.items[item_code]['Quantity'] += 1
                else:
                    self.items[item_code] = {'Description': description, 'Status': "Checked Out", 'Quantity': 1}

            self.update_table()
        else:
            QMessageBox.critical(self, "Error", "Item code not found in the Excel file.")

    def update_table(self):
        self.table.setRowCount(0)  # Clear the table
        for i, (item_code, item_info) in enumerate(self.items.items()):
            self.table.insertRow(i)
            self.table.setItem(i, 0, QTableWidgetItem(item_code))
            self.table.setItem(i, 1, QTableWidgetItem(item_info['Description']))
            self.table.setItem(i, 2, QTableWidgetItem(item_info['Status']))
            self.table.setItem(i, 3, QTableWidgetItem(str(item_info['Quantity'])))

    def export_to_excel(self):
        if not self.items:
            QMessageBox.critical(self, "Error", "No items to export.")
            return

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(["Item Code", "Description", "Status", "Quantity"])

        for item_code, item_info in self.items.items():
            worksheet.append([item_code, item_info['Description'], item_info['Status'], item_info['Quantity']])

        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            workbook.save(file_path)
            QMessageBox.information(self, "Success", f"Data exported to {file_path} successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export to Excel: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BarcodeApp()
    window.show()
    sys.exit(app.exec_())
