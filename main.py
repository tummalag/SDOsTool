import sys
import pandas as pd
import PyQt5
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                           QLineEdit, QTableWidget, QTableWidgetItem, QLabel,
                           QHBoxLayout, QPushButton, QStatusBar, QMessageBox,
                           QInputDialog, QProgressDialog, QFrame)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QIcon
import subprocess
import os
from initRead import readData, VCM
from PyQt5.QtWidgets import QToolTip
from SDO_WriteTesting import writeToSDO, writeData  # Add this import
from constants import DOT_EXCEL_FILE, VCM, SAVE_NVM

class SearchTool(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Parameter Search Tool")
        self.setMinimumSize(800, 600)
        
        try:
            # Read all columns from Excel first
            self.df = pd.read_excel(DOT_EXCEL_FILE, sheet_name='DOT', header=1)
            print("Available columns:", self.df.columns.tolist())  # Debug print
            
            # Verify required columns
            required_cols = ['Object Id', 'Parameter Name', 'Data Type', 
                           'RD Access Level', 'Range', 'Offset', 'Scale', 'Unit']
            missing_cols = [col for col in required_cols if col not in self.df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

            # Clean and filter the data
            mask = (
                self.df['Parameter Name'].notna() &
                self.df['Object Id'].notna() &
                ~self.df['Parameter Name'].str.contains('reserved', case=False, na=False) &
                ~self.df['Parameter Name'].str.contains('placeholder', case=False, na=False)
            )
            
            # Apply the mask and drop duplicates
            self.df = self.df[mask].copy()
            self.df = self.df.drop_duplicates(subset=['Object Id'])
            self.df = self.df.drop_duplicates(subset=['Parameter Name'])
            
            # Replace underscores with spaces
            self.df['Parameter Name'] = self.df['Parameter Name'].str.replace('_', ' ')
            
            # Convert Offset to float, replacing NaN with 0
            self.df['Offset'] = pd.to_numeric(self.df['Offset'], errors='coerce').fillna(0)
            
            # Convert Scale to string, replacing NaN with '1'
            self.df['Scale'] = self.df['Scale'].fillna('1').astype(str)
            
            # Store Notes from column K explicitly
            self.df['Notes'] = self.df.iloc[:, 10]  # Column K is index 10 (0-based)
            print("Sample Notes:", self.df['Notes'].head())  # Debug print
            
            # Create parameter details
            self.all_data = list(zip(self.df['Object Id'], self.df['Parameter Name']))
            self.param_details = self.df.set_index('Object Id')
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")
            self.all_data = []
            self.param_details = pd.DataFrame()

        self.init_ui()

    def init_ui(self):
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)  # Changed to horizontal layout

        # Left panel for search and table
        left_panel = QWidget()
        layout = QVBoxLayout(left_panel)

        # Create search bar
        search_layout = QHBoxLayout()
        search_label = QLabel("Search:")
        search_label.setFont(QFont('Arial', 10))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Type to search parameters...")
        self.search_input.textChanged.connect(self.search_parameters)
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)  # Add search layout to main layout

        # Create table with tooltip support
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Object Id", "Parameter Name"])
        self.table.setMouseTracking(True)  # Enable mouse tracking
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, PyQt5.QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, PyQt5.QtWidgets.QHeaderView.Stretch)
        
        # Add hover event handling
        self.hover_timer = QTimer()
        self.hover_timer.setSingleShot(True)
        self.hover_timer.timeout.connect(self.show_tooltip)
        self.hover_item = None
        self.table.cellEntered.connect(self.handle_cell_hover)
        self.table.viewportEntered.connect(self.clear_tooltip)
        
        layout.addWidget(self.table)  # Add table to main layout

        # Modify button layout
        button_layout = QHBoxLayout()
        
        # Left-aligned buttons
        left_buttons = QHBoxLayout()
        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setFixedWidth(80)
        self.copy_btn.clicked.connect(self.copy_selected)
        left_buttons.addWidget(self.copy_btn)
        
        # Add stretch to push read/write buttons to the right
        button_layout.addLayout(left_buttons)
        button_layout.addStretch()
        
        # Right-aligned buttons and controls
        right_buttons = QHBoxLayout()
        
        # Add read display label
        self.read_value_label = QLabel("")
        self.read_value_label.setFrameStyle(QFrame.Panel | QFrame.Sunken)
        self.read_value_label.setMinimumWidth(100)
        self.read_value_label.setAlignment(Qt.AlignCenter)
        
        # Add write input field
        self.write_input = QLineEdit()
        self.write_input.setPlaceholderText("Enter value")
        self.write_input.setFixedWidth(100)
        
        # Create read/write buttons
        self.read_btn = QPushButton("Read")
        self.write_btn = QPushButton("Write")
        self.read_btn.setFixedWidth(60)
        self.write_btn.setFixedWidth(60)
        self.read_btn.clicked.connect(self.read_selected)
        self.write_btn.clicked.connect(self.write_selected)
        
        # Add widgets to right_buttons layout
        right_buttons.addWidget(self.read_btn)
        right_buttons.addWidget(self.read_value_label)
        right_buttons.addWidget(self.write_input)
        right_buttons.addWidget(self.write_btn)
        
        button_layout.addLayout(right_buttons)
        layout.addLayout(button_layout)
        
        # Add left panel to main layout
        main_layout.addWidget(left_panel, stretch=4)  # Takes 80% of width

        # Right panel for additional options
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # Add a separator line
        line = QFrame()
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)

        # Additional Options Section
        options_label = QLabel("Additional Options")
        options_label.setFont(QFont('Arial', 12, QFont.Bold))
        right_layout.addWidget(options_label)

        # Add InitRead button
        self.init_read_btn = QPushButton("Run InitRead")
        self.init_read_btn.clicked.connect(self.run_init_read)
        right_layout.addWidget(self.init_read_btn)

        # Add Auto Write Test button
        self.auto_write_btn = QPushButton("Auto Write Test")
        self.auto_write_btn.clicked.connect(self.run_auto_write)
        right_layout.addWidget(self.auto_write_btn)

        # Add Save NVM button
        self.save_nvm_btn = QPushButton("Save NVM")
        self.save_nvm_btn.clicked.connect(self.save_nvm)
        right_layout.addWidget(self.save_nvm_btn)
        
        # Add stretch to push everything to the top
        right_layout.addStretch()

        # Add right panel to main layout
        main_layout.addWidget(right_panel, stretch=1)  # Takes 20% of width

        # Add status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage(f"Loaded {len(self.all_data)} parameters")

        # Initial population of table
        self.populate_table(self.all_data)

    def run_init_read(self):
        try:
            script_path = os.path.join(os.path.dirname(__file__), 'initRead.py')
            self.status_bar.showMessage("Running InitRead script...")
            subprocess.Popen([sys.executable, script_path])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to run InitRead: {str(e)}")
            self.status_bar.showMessage("Failed to run InitRead script")

    def run_auto_write(self):
        """Run the Auto Write Test functionality"""
        try:
            self.status_bar.showMessage("Running Auto Write Test...")
            writeToSDO()  # Call the writeToSDO function directly
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to run Auto Write Test: {str(e)}")
            self.status_bar.showMessage("Failed to run Auto Write Test")

    def save_nvm(self):
        """Save parameters to non-volatile memory"""
        try:
            result = writeData(
                VCM,
                hex(SAVE_NVM),  # 0x2107
                'u8',           # 8-bit value
                'Technician',
                '0-1',         # Range
                0,             # No offset
                '1',           # No scaling
                '',            # No notes
                manual_value=1  # Write value 1
            )
            
            if result == "Success":
                self.status_bar.showMessage("Parameters saved to non-volatile memory", 3000)
            else:
                self.status_bar.showMessage("Failed to save parameters", 3000)
                
        except Exception as e:
            self.status_bar.showMessage(f"Error saving parameters: {str(e)}", 3000)

    def search_parameters(self):
        search_text = self.search_input.text().lower()
        filtered_data = [
            item for item in self.all_data 
            if search_text in str(item[0]).lower() or search_text in str(item[1]).lower()
        ]
        self.populate_table(filtered_data)
        self.status_bar.showMessage(f"Found {len(filtered_data)} matches")

    def populate_table(self, data):
        self.table.setRowCount(len(data))
        for i, (obj_id, param_name) in enumerate(data):
            obj_item = QTableWidgetItem(str(obj_id))
            param_item = QTableWidgetItem(str(param_name))
            
            # Set items non-editable
            obj_item.setFlags(obj_item.flags() & ~Qt.ItemIsEditable)
            param_item.setFlags(param_item.flags() & ~Qt.ItemIsEditable)
            
            self.table.setItem(i, 0, obj_item)
            self.table.setItem(i, 1, param_item)

    def handle_cell_hover(self, row, column):
        self.hover_timer.stop()
        self.hover_item = (row, column)
        self.hover_timer.start(1000)  # 1 second delay

    def clear_tooltip(self):
        self.hover_timer.stop()
        QToolTip.hideText()

    def show_tooltip(self):
        if self.hover_item is None:
            return
            
        row, column = self.hover_item
        obj_id = self.table.item(row, 0).text()
        
        try:
            notes = self.param_details.loc[obj_id, 'Notes']
            print(f"Tooltip for {obj_id}: {notes}")  # Debug print
            if pd.notna(notes) and str(notes).strip():  # Check if notes exist and not empty
                pos = self.table.viewport().mapToGlobal(
                    self.table.visualRect(self.table.model().index(row, column)).center())
                QToolTip.showText(pos, str(notes), self.table)
        except Exception as e:
            print(f"Tooltip error for {obj_id}: {e}")  # Debug print
            pass

    def copy_selected(self):
        selected_items = self.table.selectedItems()
        if selected_items:
            # Group items by row
            rows = {}
            for item in selected_items:
                row = item.row()
                if row not in rows:
                    rows[row] = []
                rows[row].append(item.text())
            
            # Create formatted text with both Object ID and Parameter Name
            copied_text = '\n'.join([' - '.join(row_items) for row_items in rows.values()])
            QApplication.clipboard().setText(copied_text)
            self.status_bar.showMessage("Copied to clipboard!", 2000)

    def read_selected(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("Please select a parameter to read", 2000)
            return

        row = selected_items[0].row()
        obj_id = self.table.item(row, 0).text()
        param_name = self.table.item(row, 1).text()

        try:
            params = self.param_details.loc[obj_id]
            result, unit = readData(
                VCM,
                obj_id,
                str(params['Data Type']),
                str(params['RD Access Level']),
                str(params['Range']),
                float(params['Offset']),
                str(params['Scale']),
                str(params['Unit'])
            )

            if result is None:
                self.read_value_label.setText("Error")
                self.status_bar.showMessage("Read operation failed", 2000)
                return

            # Format and display result
            result_str = str(result) if pd.isna(unit) else f"{result} {unit}"
            self.read_value_label.setText(result_str)
            self.status_bar.showMessage(f"Successfully read {param_name}", 3000)

        except Exception as e:
            self.read_value_label.setText("Error")
            self.status_bar.showMessage("Read operation failed", 2000)

    def write_selected(self):
        """Write value to selected parameter"""
        selected_items = self.table.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("Please select a parameter to write", 2000)
            return

        value = self.write_input.text().strip()
        if not value:
            self.status_bar.showMessage("Please enter a value to write", 2000)
            return

        try:
            float_value = float(value)
        except ValueError:
            self.status_bar.showMessage("Please enter a valid numeric value", 2000)
            return

        row = selected_items[0].row()
        obj_id = self.table.item(row, 0).text()
        param_name = self.table.item(row, 1).text()

        try:
            params = self.param_details.loc[obj_id]
            result = writeData(
                VCM,
                obj_id,
                str(params['Data Type']),
                str(params['WR Access level']),
                str(params['Range']),
                float(params['Offset']),
                str(params['Scale']),
                str(params.get('Notes', '')),
                manual_value=float_value
            )

            if result == "Success":
                self.write_input.clear()
                self.status_bar.showMessage(f"Successfully wrote to {param_name}", 3000)
            else:
                self.status_bar.showMessage(f"Write operation failed: {result}", 2000)

        except Exception as e:
            self.status_bar.showMessage(f"Write operation failed: {str(e)}", 2000)

    def closeEvent(self, event):
        event.accept()

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern style
    window = SearchTool()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
