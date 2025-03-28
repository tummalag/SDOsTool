"""
DOT Automation Script
This script automates the writing of Diagnostic Object Table (DOT) parameters
for vehicle ECUs using CAN communication. It reads parameters from an Excel file,
generates appropriate test values, and writes them to the ECU via CAN bus.
"""

import tkinter as tk
from tkinter import ttk, filedialog
import cankit as ck
import prettytable as pt
import pandas as pd
from typing import Union, Optional
import datetime
import glob
import os
import threading
from constants import DOT_EXCEL_FILE, VCM, VCM_RESP

class SDOGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SDO Write Operations")
        self.root.geometry("700x600")
        
        # Configure basic style
        style = ttk.Style()
        style.configure("Header.TLabel", font=('Segoe UI', 11, 'bold'))
        style.configure("Value.TLabel", font=('Segoe UI', 10))
        style.configure("Status.TLabel", font=('Segoe UI', 10))
        style.configure("Custom.TButton", font=('Segoe UI', 10), padding=5)
        style.configure("Custom.Horizontal.TProgressbar", thickness=12)
        
        # Main frame
        self.main_frame = ttk.Frame(self.root, padding="8")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=8, pady=8)
        
        # Header
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0,8))
        ttk.Label(header_frame, text="DOT Parameter Write Tool",
                 font=('Segoe UI', 14, 'bold')).pack()
        
        # Info frame with better grouping
        info_frame = ttk.LabelFrame(self.main_frame, 
                                  text="Current Operation",
                                  padding=5)
        info_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0,8))
        
        # Current SDO info with tighter spacing
        ttk.Label(info_frame, text="Current SDO:", style="Header.TLabel").grid(row=0, column=0, sticky=tk.W, padx=3)
        self.sdo_label = ttk.Label(info_frame, text="---", style="Value.TLabel")
        self.sdo_label.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(info_frame, text="Parameter:", style="Header.TLabel").grid(row=1, column=0, sticky=tk.W, padx=3)
        self.param_label = ttk.Label(info_frame, text="---", style="Value.TLabel")
        self.param_label.grid(row=1, column=1, sticky=tk.W)
        
        ttk.Label(info_frame, text="Test Value:", style="Header.TLabel").grid(row=2, column=0, sticky=tk.W, padx=3)
        self.value_label = ttk.Label(info_frame, text="---", style="Value.TLabel")
        self.value_label.grid(row=2, column=1, sticky=tk.W)
        
        # Status section with color coding
        status_frame = ttk.Frame(self.main_frame)
        status_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0,8))
        
        ttk.Label(status_frame, text="Status:", style="Header.TLabel").grid(row=0, column=0, sticky=tk.W)
        self.status_label = ttk.Label(status_frame, text="Ready", style="Status.TLabel")
        self.status_label.grid(row=0, column=1, sticky=tk.W)
        
        # Modern progress bar
        self.progress = ttk.Progressbar(status_frame, 
                                      length=350,
                                      mode='determinate',
                                      style="Custom.Horizontal.TProgressbar")
        self.progress.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Just keep the restore button
        self.restore_btn = ttk.Button(status_frame, text="Restore from File",
                                    command=self.restore_values, style="Custom.TButton",
                                    width=20)
        self.restore_btn.grid(row=1, column=2, padx=5)
        
        # Store restore file path
        self.restore_file = None
        
        # Store constant excel file path
        self.excel_file = DOT_EXCEL_FILE
        
        # Results area with improved visuals
        results_frame = ttk.LabelFrame(self.main_frame,
                                     text="Operation Results",
                                     padding=5)
        results_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Results text with simpler style
        self.results_text = tk.Text(results_frame,
                                  height=20,
                                  width=80,
                                  font=('Consolas', 9))
        self.results_text.grid(row=0, column=0, padx=3, pady=3)
        
        # Add modern scrollbar
        scrollbar = ttk.Scrollbar(results_frame,
                                orient='vertical',
                                command=self.results_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.results_text.configure(yscrollcommand=scrollbar.set)
        
        # Configure grid weights
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.is_running = True

    def update_sdo(self, sdo_id: str, param_name: str):
        self.sdo_label.config(text=sdo_id)
        self.param_label.config(text=param_name)
        self.root.update()

    def update_value(self, value: str):
        self.value_label.config(text=value)
        self.root.update()

    def update_status(self, status: str):
        """Update status message"""
        self.status_label.config(text=status)  # Fixed text() to text=
        self.root.update()

    def update_progress(self, current: int, total: int):
        progress = (current / total) * 100
        self.progress['value'] = progress
        self.root.update()

    def add_result(self, result: str):
        self.results_text.insert(tk.END, result + "\n")
        self.results_text.see(tk.END)
        self.root.update()

    def on_closing(self):
        self.is_running = False
        self.root.destroy()

    def restore_values(self):
        """Open file dialog to select restore file and start restore process"""
        filename = filedialog.askopenfilename(
            title="Select Values File to Restore",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.restore_file = filename
            thread = threading.Thread(target=restore_values, args=(self,))
            thread.start()

    def start(self):
        self.root.mainloop()

# Initialize CAN bus communication at 500kbps
MAINBUS = ck.MainBus('pcan', 'PCAN_USBBUS1', 500000)

def convert_scale(scale: Union[str, float, int]) -> float:
    """
    Convert scale value from various formats to float
    Args:
        scale: Can be string (e.g., "1/10"), float, or int
    Returns:
        Converted scale as float, defaults to 1.0 if invalid
    """
    if isinstance(scale, (float, int)):
        return float(scale)
    if isinstance(scale, str):
        if '/' in scale:
            num, denom = scale.split('/')
            return float(num) / float(denom)
        return float(scale)
    return 1.0

def parse_range(range_str: str) -> tuple[float, float]:
    """
    Parse range string from Excel into min and max values
    Args:
        range_str: String like "0 to 100" or "0-100"
    Returns:
        Tuple of (min_value, max_value) as floats
    """
    try:
        if isinstance(range_str, str):
            # Remove any whitespace and replace 'to' with '-'
            range_str = range_str.replace(' to ', '-').replace(' ', '')
            if '-' in range_str:
                min_val, max_val = range_str.split('-')
                return float(min_val), float(max_val)
    except:
        pass
    return 0, 100  # default range if parsing fails

def get_test_value(range_str: str) -> float:
    """Generate a test value within the given range"""
    min_val, max_val = parse_range(range_str)
    # Use the middle of the range as test value
    return (min_val + max_val) / 2

def get_precision_from_scale(scale: Union[str, float, int]) -> int:
    """
    Determine decimal precision needed based on scale value
    Args:
        scale: Scale value that determines precision (e.g., 0.1 needs 1 decimal)
    Returns:
        Number of decimal places needed
    """
    scale_val = convert_scale(scale)
    if scale_val >= 1:
        return 0
    # Convert scale to string and count decimals
    scale_str = str(scale_val).split('.')[-1]
    return len(scale_str)

def parse_enum_notes(notes: str) -> dict[int, str]:
    """Parse notes containing enum value descriptions"""
    try:
        enum_dict = {}
        if pd.isna(notes):
            return enum_dict
            
        notes = str(notes)
        # Match patterns like "0 = Text Display" or "0: Text Display"
        for line in notes.split(';'):
            line = line.strip()
            if '=' in line:
                val, desc = line.split('=', 1)
            elif ':' in line:
                val, desc = line.split(':', 1)
            else:
                continue
                
            try:
                enum_val = int(val.strip())
                enum_dict[enum_val] = desc.strip()
            except:
                continue
                
        return enum_dict
    except:
        return {}

def generate_test_values(range_str: str, offset: float, scale: Union[str, float, int], notes: str, datatype: str) -> tuple[list[float], dict[int, str]]:
    """Generate test values and their descriptions if available"""
    enum_desc = {}
    test_values = []
    
    try:
        if datatype.lower() == 'enum':
            enum_desc = parse_enum_notes(notes)
            if enum_desc:
                test_values = list(enum_desc.keys())
                return test_values, enum_desc
        
        if not range_str or str(range_str).upper() in ['N/A', 'NA']:
            return [], {}
        
        min_val, max_val = parse_range(range_str)
        scale_value = convert_scale(scale)
        precision = get_precision_from_scale(scale)
        
        # Check notes for specific test values
        if pd.notna(notes):
            notes_lower = str(notes).lower()
            
            # Boolean type values
            if "boolean" in notes_lower or "true/false" in notes_lower:
                return [0, 1], {}
            
            # Specific enumerated values
            if "values:" in notes_lower:
                values_str = notes_lower.split("values:")[-1].strip()
                try:
                    return [float(v.strip()) for v in values_str.split(',')], {}
                except:
                    pass
            
            # Special cases like "0 = disabled"
            if "disabled" in notes_lower:
                return [0, 1, max_val], {}
            
            # High/Low testing
            if any(x in notes_lower for x in ["high/low", "high-low", "min/max"]):
                return [min_val, max_val], {}

        # Default test strategy if no specific notes
        min_scaled = min_val * scale_value + offset
        max_scaled = max_val * scale_value + offset
        middle = (min_scaled + max_scaled) / 2
        
        test_points = [min_scaled, middle, max_scaled]
        return [round(x, precision) for x in test_points], {}
        
    except Exception as e:
        return [], {}

def create_payload_bytes(value: float, datatype: str) -> list[int]:
    """
    Convert test value to appropriate byte format for CAN message
    Args:
        value: Value to convert
        datatype: Target data type (u8, u16, u32)
    Returns:
        List of bytes in little-endian format
    """
    if datatype == 'u8':
        return list(int(value).to_bytes(1, 'little'))
    elif datatype == 'u16':
        return list(int(value).to_bytes(2, 'little'))
    elif datatype == 'u32':
        return list(int(value).to_bytes(4, 'little'))
    else:
        return list(int(value).to_bytes(4, 'little'))  # default to u32

def read_sdo_value(ecu: int, sdo: int) -> float:
    """Read value from an SDO"""
    sdo_low = sdo & 0xFF
    sdo_high = (sdo >> 8) & 0xFF
    can_msg = [0x40, sdo_low, sdo_high, 0x00, 0x00, 0x00, 0x00, 0x00]
    data = MAINBUS.fully_written_message_and_response(ecu, can_msg, VCM_RESP)
    if data == 'Timeout':
        return None
    # Extract value from response (assuming u32 format)
    value_bytes = [int(x, 16) for x in data[12:].split()]
    return int.from_bytes(value_bytes, 'little')

def writeData(ecu: int, sdo: str, datatype: str, write: str, 
            rng: str, offset: float, scale: Union[str, float, int], notes: str, 
            manual_value: float = None) -> str:
    """Write values to ECU. For manual writes, use manual_value parameter."""
    try:
        # Convert SDO from hex string to int, supporting multiple formats
        if isinstance(sdo, str):
            clean_sdo = sdo.replace('0x', '').strip()
            sdo_int = int(clean_sdo, 16)
        else:
            sdo_int = int(sdo)
        
        # Skip specific SDOs
        if sdo_int in [0x2600, 0x2705, 0x2733]:
            return None
            
        # Check write access first
        if write != 'Technician' and write != 'User':
            return None
            
        # Bit operations for CAN message
        sdo_low = sdo_int & 0xFF
        sdo_high = (sdo_int >> 8) & 0xFF
        
        # For manual writes, skip test value generation
        if manual_value is not None:
            # Create CAN message for single value
            payload_bytes = create_payload_bytes(manual_value, datatype)
            while len(payload_bytes) < 4:
                payload_bytes.append(0x00)
                
            can_msg = [0x23, sdo_low, sdo_high, 0x00, *payload_bytes]
            data = MAINBUS.fully_written_message_and_response(ecu, can_msg, VCM_RESP)
            return "Success" if data != 'Timeout' else "Timeout"

        # Auto write test logic (only for auto write test)
        # Generate test values and get descriptions if available
        test_values, value_descriptions = generate_test_values(rng, offset, scale, notes, datatype)
        if not test_values:
            print(f"Could not generate test values for SDO {sdo}. Skipping...")
            return "Test Value Error"
        
        results = []
        for test_value in test_values:
            description = value_descriptions.get(int(test_value), "")
            value_display = f"{test_value} ({description})" if description else str(test_value)
            print(f"Testing value: {value_display}")
            
            # No need to scale here since values are already scaled
            payload_bytes = create_payload_bytes(test_value, datatype)
            
            # Pad payload to 4 bytes
            while len(payload_bytes) < 4:
                payload_bytes.append(0x00)
            
            # Create 8-byte CAN message
            can_msg = [
                0x23,           # Write command
                sdo_low,        # SDO low byte
                sdo_high,       # SDO high byte
                0x00,           # Always 0
                *payload_bytes  # Payload bytes (will use 1-4 bytes based on datatype)
            ]
            
            # Write to ECU
            data = MAINBUS.fully_written_message_and_response(ecu, can_msg, VCM_RESP)
            
            if data == 'Timeout':
                results.append(f"Timeout at {value_display}")
                continue
            
            try:
                results.append(f"Success at {value_display}")
            except Exception as e:
                results.append(f"Error at {value_display}: {e}")
        
        return " | ".join(results) if results else "No results"
        
    except Exception as e:
        print(f"Error in writeData for SDO {sdo}: {e}")
        return "CAN_ERROR"

def get_output_filename() -> str:
    """Generate output filename based on date"""
    today = datetime.datetime.now().strftime('%d%m%y')
    base = f"WriteTests_{today}"
    existing = glob.glob(f"{base}_*.txt")
    count = len(existing) + 1
    return f"{base}_{count}.txt"

def read_original_values(filename: str) -> dict:
    """Read original SDO values from text file with format handling"""
    original_values = {}
    try:
        with open(filename, 'r') as f:
            current_sdo = None
            for line in f:
                line = line.strip()
                if '|' not in line:
                    continue
                    
                parts = [p.strip() for p in line.split('|')]
                if len(parts) < 3:
                    continue
                    
                # Check if this line has an SDO ID
                if parts[1].startswith('0x'):
                    try:
                        sdo = int(parts[1], 16)
                        value_str = parts[3].split()[0]  # Get first word of value
                        if value_str.replace('.', '').isdigit():  # Check if numeric
                            value = float(value_str)
                            original_values[sdo] = value
                    except:
                        continue
    except FileNotFoundError:
        print(f"Original values file {filename} not found")
    return original_values

def read_writable_sdos(filename: str, excel_file: str) -> dict:
    """Read original values but only for writable SDOs"""
    try:
        # Read Excel to get writable SDOs
        data = pd.read_excel(excel_file, header=1)
        writable_sdos = set()
        for _, row in data.iterrows():
            if row['WR Access level'] in ['Technician', 'User']:
                try:
                    sdo = int(str(row['Object Id']), 16)
                    writable_sdos.add(sdo)
                except:
                    continue

        # Read values but only keep writable ones
        original_values = {}
        with open(filename, 'r') as f:
            for line in f:
                if '|' not in line:
                    continue
                    
                parts = [p.strip() for p in line.split('|')]
                if len(parts) < 3 or not parts[1].startswith('0x'):
                    continue
                    
                try:
                    sdo = int(parts[1], 16)
                    if sdo not in writable_sdos:
                        continue
                        
                    value_str = parts[3].split()[0]
                    if value_str.replace('.', '').isdigit():
                        value = float(value_str)
                        original_values[sdo] = value
                except:
                    continue
                    
        return original_values
    except Exception as e:
        print(f"Error reading values: {e}")
        return {}

def restore_values(self):
    """Handle restore button click with selective restoration"""
    try:
        if not self.restore_file:
            self.update_status("Please select a file to restore from")
            return
            
        self.update_status("Reading original values...")
        original_values = read_writable_sdos(self.restore_file, self.excel_file)
        
        if not original_values:
            self.update_status("Error: No writable values found")
            return
            
        total = len(original_values)
        restored = 0
        
        self.update_status("Restoring writable values...")
        for sdo, value in original_values.items():
            if not self.is_running:
                break
                
            try:
                # Skip specific SDOs that shouldn't be written
                if sdo in [0x2600, 0x2705, 0x2733]:
                    continue
                    
                sdo_hex = f"0x{sdo:04X}"
                self.update_sdo(sdo_hex, "Restoring...")
                self.update_value(f"{value}")
                
                # Use writeData with manual_value parameter
                result = writeData(
                    VCM,
                    sdo_hex,
                    'u32',  # Default to u32 for restore
                    'Technician',
                    None,
                    0,
                    1,
                    '',
                    manual_value=value
                )
                
                restored += 1
                self.update_progress(restored, total)
                
                result_str = "Success" if result == "Success" else "Failed"
                self.add_result(f"Restored {sdo_hex} = {value} ({result_str})")
            except Exception as e:
                self.add_result(f"Error restoring {sdo_hex}: {str(e)}")
                continue
        
        self.update_status(f"Restore complete - {restored}/{total} values restored")
        
    except Exception as e:
        self.update_status(f"Error: {str(e)}")
        self.add_result(f"Restore failed: {str(e)}")

def restore_original_values(ecu: int, original_values: dict) -> None:
    """Restore original SDO values"""
    for sdo, value in original_values.items():
        sdo_low = sdo & 0xFF
        sdo_high = (sdo >> 8) & 0xFF
        payload_bytes = list(int(value).to_bytes(4, 'little'))
        
        can_msg = [
            0x23,           # Write command
            sdo_low,        # SDO low byte
            sdo_high,       # SDO high byte
            0x00,           # Always 0
            *payload_bytes  # Original value
        ]
        
        MAINBUS.fully_written_message_and_response(ecu, can_msg, VCM_RESP)

def writeToSDO() -> None:
    """Main function to process DOT parameters"""
    gui = SDOGUI()
    gui.excel_file = DOT_EXCEL_FILE  # Use constant file
    gui.update_status("Ready to process DOT parameters")
    
    def process_sdos():
        table = pt.PrettyTable()
        table.field_names = ['Object ID', 'Param Name', 'Range', 'Result']
        
        try:
            gui.update_status("Reading Excel file...")
            data = pd.read_excel(DOT_EXCEL_FILE, header=1)  # Use constant directly
            
            # Filter data for SDOs below 0x5100
            data = data[data['Object Id'].astype(str).apply(
                lambda x: int(x, 16) if x != 'Reserved' else 0) < 0x5100]

            data_subset = data[data['Object Id'] != 'Reserved']
            total_rows = len(data_subset)
            gui.update_status(f"Processing SDOs up to 0x5100 ({total_rows} entries)")
            
            written_count = 0
            for i in range(total_rows):
                if not gui.is_running:
                    break
                    
                current_row = data_subset.iloc[i]
                next_row = data_subset.iloc[i + 1] if i < total_rows - 1 else current_row
                
                # Update GUI with current SDO
                gui.update_sdo(current_row['Object Id'], current_row['Parameter Name'])
                gui.update_progress(i + 1, total_rows)
                
                if current_row['Object Id'] == 'Reserved':
                    continue
                    
                sdo_int = int(current_row['Object Id'], 16)
                if sdo_int in [0x2600, 0x2705, 0x2733]:
                    continue

                result = writeData(
                    VCM, 
                    current_row['Object Id'],
                    current_row['Data Type'],
                    current_row['WR Access level'],
                    current_row['Range'],
                    current_row['Offset'],
                    current_row['Scale'],
                    current_row.get('Notes', '')  # Handle missing Notes column
                )
                
                if result is not None:
                    written_count += 1
                    value_display = f"Testing SDO {current_row['Object Id']}"
                    gui.update_value(value_display)
                    gui.add_result(f"{value_display}: {result}")
                    
                    if result != "CAN_ERROR":
                        is_divider = next_row["Object Id"] == 'Reserved'
                        table.add_row([
                            current_row['Object Id'],
                            current_row['Parameter Name'],
                            current_row['Range'],
                            result
                        ], divider=is_divider)
            
            output_file = get_output_filename()
            with open(output_file, 'w') as f:
                f.write(table.get_string())
            
            gui.update_status(f"Complete - {written_count} SDOs processed")
            
        except Exception as e:
            gui.update_status(f"Error: {str(e)}")
            raise

    thread = threading.Thread(target=process_sdos)
    thread.start()
    gui.start()

def main():
    """Entry point for DOT automation script"""
    writeToSDO()

if __name__ == '__main__':
    main()
