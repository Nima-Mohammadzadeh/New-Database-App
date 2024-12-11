import sys
import math
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
                               QLineEdit, QPushButton, QFileDialog, QCheckBox, QMessageBox,
                               QLabel)
from PySide6.QtGui import QFont
from openpyxl import Workbook
import os
import webbrowser

from logic import generate_epc  # Ensure logic.py is in the same directory and that generate_epc is properly defined there

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RFID Encoding Application")
        
        # Set a slightly larger default font for better readability
        font = QFont("Arial", 10)
        self.setFont(font)

        # Layouts
        main_layout = QVBoxLayout()
        form_layout = QFormLayout()
        form_layout.setHorizontalSpacing(20)
        form_layout.setVerticalSpacing(10)
        
        # Create input fields
        self.upc_input = QLineEdit()
        self.upc_input.setPlaceholderText("Enter 12-digit UPC")
        
        self.start_serial_input = QLineEdit()
        self.start_serial_input.setPlaceholderText("Enter Starting Serial Number")
        
        self.quantity_input = QLineEdit()
        self.quantity_input.setPlaceholderText("Enter Quantity")
        
        self.lpr_input = QLineEdit()
        self.lpr_input.setPlaceholderText("Enter LPR (Labels Per Roll)")
        
        self.qty_db_input = QLineEdit()
        self.qty_db_input.setPlaceholderText("Enter QTY/DB")
        
        self.save_location_label = QLabel("No directory selected")
        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_directory)
        
        # Checkboxes
        self.two_percent_check = QCheckBox("2%")
        self.seven_percent_check = QCheckBox("7%")
        
        # Add widgets to form layout
        form_layout.addRow("UPC:", self.upc_input)
        form_layout.addRow("Starting Serial #:", self.start_serial_input)
        form_layout.addRow("Quantity:", self.quantity_input)
        form_layout.addRow("LPR (Labels/Roll):", self.lpr_input)
        form_layout.addRow("QTY/DB:", self.qty_db_input)
        
        # Save location row
        hlayout_save = QHBoxLayout()
        hlayout_save.addWidget(self.save_location_label)
        hlayout_save.addWidget(self.browse_button)
        form_layout.addRow("Save Location:", hlayout_save)
        
        # Checkboxes
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(self.two_percent_check)
        checkbox_layout.addWidget(self.seven_percent_check)
        form_layout.addRow("Adjust Quantity:", checkbox_layout)
        
        main_layout.addLayout(form_layout)
        
        # Generate button
        self.generate_button = QPushButton("Generate Excel Files + Roll Tracker")
        self.generate_button.clicked.connect(self.generate_files)
        self.generate_button.setMinimumHeight(40)  # Make button a bit taller for visibility
        
        main_layout.addWidget(self.generate_button)
        
        self.setLayout(main_layout)
        
        self.output_directory = None  # will store chosen directory
    
    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Save Directory")
        if directory:
            self.output_directory = directory
            self.save_location_label.setText(directory)
        
    def generate_files(self):
        # Validate inputs
        upc = self.upc_input.text().strip()
        start_serial_str = self.start_serial_input.text().strip()
        quantity_str = self.quantity_input.text().strip()
        lpr_str = self.lpr_input.text().strip()
        qty_db_str = self.qty_db_input.text().strip()
        
        # Basic validations
        if len(upc) != 12 or not upc.isdigit():
            QMessageBox.warning(self, "Validation Error", "UPC must be exactly 12 digits.")
            return
        
        if not start_serial_str.isdigit():
            QMessageBox.warning(self, "Validation Error", "Starting Serial # must be an integer.")
            return
        start_serial = int(start_serial_str)
        
        if not quantity_str.isdigit():
            QMessageBox.warning(self, "Validation Error", "Quantity must be an integer.")
            return
        base_qty = int(quantity_str)
        
        if not lpr_str.isdigit():
            QMessageBox.warning(self, "Validation Error", "LPR must be an integer.")
            return
        lpr = int(lpr_str)
        if lpr <= 0:
            QMessageBox.warning(self, "Validation Error", "LPR must be greater than zero.")
            return
        
        if not qty_db_str.isdigit():
            QMessageBox.warning(self, "Validation Error", "QTY/DB must be an integer.")
            return
        qty_db = int(qty_db_str)
        if qty_db <= 0:
            QMessageBox.warning(self, "Validation Error", "QTY/DB must be greater than zero.")
            return
        
        if self.output_directory is None:
            QMessageBox.warning(self, "Validation Error", "Please select a save location.")
            return
        
        # Adjust Quantity based on checkboxes
        adjusted_qty = base_qty
        if self.two_percent_check.isChecked():
            adjusted_qty = int(round(adjusted_qty * 1.02))
        if self.seven_percent_check.isChecked():
            adjusted_qty = int(round(adjusted_qty * 1.07))
        
        if adjusted_qty <= 0:
            QMessageBox.warning(self, "Validation Error", "Adjusted Quantity must be greater than zero.")
            return
        
        num_files = math.ceil(adjusted_qty / qty_db)
        
        current_serial = start_serial
        
        try:
            # Generate DB files
            for file_index in range(1, num_files + 1):
                wb = Workbook()
                ws = wb.active
                ws.append(["UPC", "Serial #", "EPC"])
                
                rows_this_file = min(qty_db, adjusted_qty - (file_index - 1)*qty_db)
                for i in range(rows_this_file):
                    epc_hex = generate_epc(upc, current_serial)
                    ws.append([upc, current_serial, epc_hex])
                    current_serial += 1
                
                filename = os.path.join(self.output_directory, f"{upc}_DB{file_index}.xlsx")
                wb.save(filename)

            # Create "roll tracker" folder inside the chosen directory
            roll_tracker_dir = os.path.join(self.output_directory, "roll tracker")
            os.makedirs(roll_tracker_dir, exist_ok=True)

            # Build the roll tracker HTML inside the "roll tracker" folder
            roll_tracker_filename = os.path.join(roll_tracker_dir, f"roll_tracker_{upc}.html")

            def short_epc(e):
                return e[-5:] if len(e) >= 5 else e
            
            total_roll_num = 1

            with open(roll_tracker_filename, "w", encoding="utf-8") as f:
                f.write("<html><head><title>Roll Tracker</title>")
                f.write("""<style>
                body {
                    font-family: Arial, sans-serif; 
                    margin: 10px; 
                    color: #000;
                    text-align: center;
                }
                .container {
                    width: 95%; 
                    margin: 0 auto; 
                    text-align: center;
                }
                table {
                    border-collapse: collapse; 
                    margin: 0 auto; 
                    margin-bottom: 40px;
                    font-size: 13px;
                    width: 100%;
                    page-break-inside: avoid; 
                }
                tr, td, th {
                    page-break-inside: avoid;
                }
                th, td {
                    border: 1px solid #000; 
                    padding: 2px;
                    text-align: center;
                    vertical-align: middle;
                }
                th {
                    background: #f2f2f2;
                    font-weight: bold;
                }
                .db-header {
                    font-weight: bold;
                    background: #ddd;
                    page-break-inside: avoid;
                }
                .printer-row td {
                    text-align: left;
                    page-break-inside: avoid;
                }
                .sub-row td {
                    text-align: left;
                    font-size: 13px;
                    font-weight: bold;
                    padding-left: 20px;
                    page-break-inside: avoid;
                }
                .sub-row td p {
                    margin: 3px 0;
                }
                @page {
                    size: auto;
                    margin: 10mm;
                }
                @media print {
                    body {
                        margin: 0;
                    }
                    table {
                        page-break-inside: avoid;
                    }
                    tr {
                        page-break-inside: avoid;
                    }
                }
                </style>""")
                f.write("</head><body>")
                f.write("<div class='container'>")

                global_serial = start_serial
                # No title at the top.
                
                for db_index in range(1, num_files + 1):
                    db_label_count = min(qty_db, adjusted_qty - (db_index - 1)*qty_db)
                    if db_label_count <= 0:
                        break
                    db_start_label = 1
                    db_end_label = db_label_count
                    
                    f.write("<table>")
                    f.write(f"<tr class='db-header'><td colspan='5'>Database {db_index}</td></tr>")
                    f.write("<tr class='printer-row'><td colspan='5'>Printer: ______________________</td></tr>")

                    # Columns: ROLL #, LABEL RANGE, START, END
                    f.write("<tr>")
                    f.write("<th>ROLL #</th>")
                    f.write("<th>LABEL RANGE</th>")
                    f.write("<th>START</th><th>END</th>")
                    f.write("</tr>")

                    db_remaining = db_label_count
                    db_current_start_label = db_start_label

                    while db_remaining > 0:
                        roll_count = min(lpr, db_remaining)
                        roll_local_start = db_current_start_label
                        roll_local_end = roll_local_start + roll_count - 1
                        
                        first_third_end = roll_local_start + (roll_count // 3) - 1
                        second_third_end = roll_local_start + (2*(roll_count // 3)) - 1

                        epc_start_serial = global_serial
                        epc_start_val = generate_epc(upc, epc_start_serial)

                        epc_end_serial = global_serial + roll_count - 1
                        epc_end_val = generate_epc(upc, epc_end_serial)

                        # Use commas for readability
                        label_range_formatted = f"{roll_local_start:,}-{roll_local_end:,}"

                        f.write("<tr>")
                        f.write(f"<td>{total_roll_num}</td>")
                        f.write(f"<td>{label_range_formatted}</td>")
                        f.write(f"<td>{short_epc(epc_start_val)}</td>")
                        f.write(f"<td>{short_epc(epc_end_val)}</td>")
                        f.write("</tr>")

                        segment1_start = roll_local_start
                        segment1_end = first_third_end
                        segment1_start_epc = generate_epc(upc, global_serial + (segment1_start - roll_local_start))
                        segment1_end_epc = generate_epc(upc, global_serial + (segment1_end - roll_local_start))

                        segment2_start = first_third_end + 1
                        segment2_end = second_third_end
                        segment2_start_epc = generate_epc(upc, global_serial + (segment2_start - roll_local_start))
                        segment2_end_epc = generate_epc(upc, global_serial + (segment2_end - roll_local_start))

                        segment3_start = second_third_end + 1
                        segment3_end = roll_local_end
                        segment3_start_epc = generate_epc(upc, global_serial + (segment3_start - roll_local_start))
                        segment3_end_epc = generate_epc(upc, global_serial + (segment3_end - roll_local_start))

                        # Format these segments with commas
                        seg1_formatted = f"{segment1_start:,}-{segment1_end:,}"
                        seg2_formatted = f"{segment2_start:,}-{segment2_end:,}"
                        seg3_formatted = f"{segment3_start:,}-{segment3_end:,}"

                        f.write("<tr class='sub-row'>")
                        f.write("<td></td><td colspan='4'>")
                        f.write(f"<p><b>{seg1_formatted}:</b> Init ______ | Start: {short_epc(segment1_start_epc)}, End: {short_epc(segment1_end_epc)} | Notes: _________________________________</p>")
                        f.write(f"<p><b>{seg2_formatted}:</b> Init ______ | Start: {short_epc(segment2_start_epc)}, End: {short_epc(segment2_end_epc)} | Notes: _________________________________</p>")
                        f.write(f"<p><b>{seg3_formatted}:</b> Init ______ | Start: {short_epc(segment3_start_epc)}, End: {short_epc(segment3_end_epc)} | Notes: _________________________________</p>")
                        f.write("</td></tr>")

                        db_current_start_label = roll_local_end + 1
                        db_remaining -= roll_count
                        global_serial += roll_count
                        total_roll_num += 1

                    f.write("</table>")

                f.write("</div>")
                f.write("</body></html>")

            # After generation, auto open the roll tracker in the default browser
            # This will open the newly created HTML file for the user.
            webbrowser.open(f"file://{roll_tracker_filename}")

            QMessageBox.information(self, "Success", f"Successfully generated files and roll tracker:\n{roll_tracker_filename}")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {e}")

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
