import sys
import pandas as pd
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit, QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView

class BarcodeScannerApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Tooltrackr - Inventory Management System")
        self.setGeometry(100, 100, 900, 600)

        # Central widget and layout
        self.central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Branding with logo
        self.logo_label = QLabel(self)
        self.logo_pixmap = QtGui.QPixmap("logo.png").scaled(326, 51, QtCore.Qt.KeepAspectRatio)
        self.logo_label.setPixmap(self.logo_pixmap)
        self.layout.addWidget(self.logo_label, alignment=QtCore.Qt.AlignLeft)

        # Status label with rounded corners
        self.status_label = QLabel("Scan a barcode or enter item code", self)
        self.status_label.setFont(QtGui.QFont("Arial", 16, QtGui.QFont.Bold))
        self.status_label.setStyleSheet("""
            border-radius: 15px;
            padding: 10px;
            background-color: #fff3e0;
        """)
        self.layout.addWidget(self.status_label, alignment=QtCore.Qt.AlignCenter)

        # Toggle mode button
        self.mode = "Check In"
        self.toggle_button = QPushButton("Switch to Check Out", self)
        self.toggle_button.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Bold))
        self.toggle_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border: 1px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.toggle_button.clicked.connect(self.toggle_mode)
        self.layout.addWidget(self.toggle_button)

        # Export to Excel button
        self.export_button = QPushButton("Export to Excel", self)
        self.export_button.setFont(QtGui.QFont("Arial", 12, QtGui.QFont.Bold))
        self.export_button.setStyleSheet("""
            background-color: #ff9800;
            color: white;
            border: 1px solid #ff9800;
            border-radius: 15px;
            padding: 10px 20px;
        """)
        self.export_button.clicked.connect(self.export_to_excel)
        self.layout.addWidget(self.export_button)

        # Entry field for item code with black font
        self.item_code_entry = QLineEdit(self)
        self.item_code_entry.setFont(QtGui.QFont("Arial", 12))
        self.item_code_entry.setStyleSheet("""
            color: black;  /* Set the font color to black */
            background-color: white;
            border: 1px solid #ccc;
            border-radius: 10px;
            padding: 10px;
        """)
        self.item_code_entry.setPlaceholderText("Enter item code")
        self.layout.addWidget(self.item_code_entry)
        self.item_code_entry.returnPressed.connect(self.process_item)

        # Table widget for displaying items with rounded corners
        self.table = QTableWidget(self)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Item Code", "Description", "Status", "Quantity"])

        # Styling the table header (rounded corners, orange background, white font)
        self.table.horizontalHeader().setStyleSheet("""
            QHeaderView::section {
                background-color: #ff9800;
                color: white;
                border-radius: 10px;
                padding: 10px;
                font-weight: bold;
                margin-right: 10px;  /* Add margin to the right for spacing */
            }
        """)
        
        # Making the columns stretch
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        
        # Styling the table rows with black font for data
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ff9800;
                border-radius: 15px;
                background-color: #f9f9f9;
            }
            QTableWidget::item {
                padding: 10px;
                color: black;  /* Set the font color to black for the table data */
            }
        """)
        
        self.layout.addWidget(self.table)

        # Data structures
        self.items = {}
        self.inventory_df = None

        # Load Excel data at startup
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

        self.item_code_entry.clear()

    def update_table(self):
        self.table.setRowCount(0)
        for item_code, data in self.items.items():
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(item_code))
            self.table.setItem(row_position, 1, QTableWidgetItem(data['Description']))
            status_item = QTableWidgetItem(data['Status'])
            if data['Status'] == "Checked In":
                status_item.setBackground(QtGui.QColor(200, 255, 200))  # Soft Green
            else:
                status_item.setBackground(QtGui.QColor(255, 200, 200))  # Soft Red
            self.table.setItem(row_position, 2, status_item)
            self.table.setItem(row_position, 3, QTableWidgetItem(str(data['Quantity'])))

        # Make the Quantity column editable
        self.table.cellChanged.connect(self.on_cell_changed)

    def on_cell_changed(self, row, column):
        if column == 3:  # If the "Quantity" column is edited
            item_code = self.table.item(row, 0).text()
            new_quantity = self.table.item(row, 3).text()
            try:
                new_quantity = int(new_quantity)
                if new_quantity < 0:
                    self.table.item(row, 3).setText("0")  # Prevent negative quantities
                    new_quantity = 0
                if item_code in self.items:
                    self.items[item_code]['Quantity'] = new_quantity
            except ValueError:
                self.table.item(row, 3).setText(str(self.items[item_code]['Quantity']))

    def export_to_excel(self):
        if not self.items:
            QMessageBox.warning(self, "No Data", "No items to export.")
            return

        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(self, "Save Excel File", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        try:
            df = pd.DataFrame.from_dict(self.items, orient='index')
            df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Success", "File exported successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export to Excel: {e}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = BarcodeScannerApp()
    window.show()
    sys.exit(app.exec_())
