"""
DOT Automation Script
This script automates the reading and writing of Diagnostic Object Table (DOT) parameters
for vehicle ECUs using CAN communication.

Author: [Gowtham Tummala]
Date: [2025-03-25]
Version: 0.1
"""

import cankit as ck
import prettytable as pt
import pandas as pd
from typing import Union, Optional
import tkinter as tk
from tkinter import ttk, filedialog
import datetime
import glob
import os
import argparse
from constants import (DOT_EXCEL_FILE, VCM, MCU, VCM_RESP, MCU_RESP,
                      SAVE_NVM, VCM_ver_num, MCU_ver_num, BMS_ver_num)

# CAN Bus Configuration
MAINBUS = ck.MainBus('pcan', 'PCAN_USBBUS1', 500000)

def convert_scale(scale: Union[str, float, int]) -> float:
    """Convert scale value from string or number to float"""
    if isinstance(scale, (float, int)):
        return float(scale)
    if isinstance(scale, str):
        if '/' in scale:
            num, denom = scale.split('/')
            return float(num) / float(denom)
        return float(scale)
    return 1.0

def show_error_dialog(sdo: str, param_name: str) -> bool:
    """Show error dialog with retry option"""
    error_window = tk.Tk()
    error_window.title("CAN Communication Error")
    error_window.geometry("400x200")
    
    # Center the window on screen
    window_width = 400
    window_height = 200
    screen_width = error_window.winfo_screenwidth()
    screen_height = error_window.winfo_screenheight()
    center_x = int(screen_width/2 - window_width/2)
    center_y = int(screen_height/2 - window_height/2)
    error_window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    # Message frame
    msg_frame = ttk.Frame(error_window, padding="20")
    msg_frame.pack(fill=tk.BOTH, expand=True)
    
    message = f"ERROR IN CAN Communication\nSDO: {sdo}\nParameter: {param_name}"
    ttk.Label(msg_frame, text=message, justify=tk.CENTER).pack(pady=20)
    
    # Button frame
    btn_frame = ttk.Frame(msg_frame)
    btn_frame.pack(fill=tk.X, pady=10)
    
    retry = tk.BooleanVar(value=False)
    
    def on_retry():
        retry.set(True)
        error_window.quit()
    
    def on_cancel():
        retry.set(False)
        error_window.quit()
    
    # Style configuration for buttons
    style = ttk.Style()
    style.configure('Retry.TButton', background='green')
    style.configure('Stop.TButton', background='red')
    
    # Center-aligned buttons with padding
    btn_frame.columnconfigure(0, weight=1)
    btn_frame.columnconfigure(3, weight=1)
    
    ttk.Button(btn_frame, text="Retry", style='Retry.TButton', command=on_retry).grid(row=0, column=1, padx=10)
    ttk.Button(btn_frame, text="Stop", style='Stop.TButton', command=on_cancel).grid(row=0, column=2, padx=10)
    
    error_window.mainloop()
    error_window.destroy()
    return retry.get()

def readData(ecu: int, sdo: str, datatype: str, read: str, 
            rng: str, offset: float, scale: Union[str, float, int], unit: str) -> Union[tuple[Union[str, float, int], str], tuple[Union[str, float, int], None]]:
    """
    Read data from specified ECU using SDO protocol.
    Returns tuple of (value, unit) where unit may be None
    """
    try:
        # Convert SDO from hex string to int
        sdo_int = int(sdo, 16)
        
        # Read raw data from ECU
        data = MAINBUS.objID_read_and_response_message(ecu, sdo_int, VCM_RESP)

        if data == 'Timeout':
            return (None, "CAN_ERROR")
    
        # Handle version number special cases - no units for these
        bytes_list = [int(x, 16) for x in data[12:].split()]
        if sdo_int in (VCM_ver_num, MCU_ver_num, BMS_ver_num):
            if sdo_int == VCM_ver_num:
                value = f"{(bytes_list[3] << 8) | bytes_list[2]}.{bytes_list[1]}.{bytes_list[0]}"
            else:
                value = f"{bytes_list[3]}.{bytes_list[2]}.{(bytes_list[1] << 8) | bytes_list[0]}"
            return (value, None)

        try:
            # Process ASCII data
            if datatype == 'ASCII' or unit == 'ASCII':
                chars = [chr(int(x, 16)) for x in data[12:].split()]
                # Filter out NUL characters and any non-printable characters
                value = ''.join(c for c in chars if c.isprintable() and c != '\x00').strip()
                return (value, None)  # Always return None for unit with ASCII data
                
            # Handle numeric types
            if datatype == 'u8':
                payload = int.from_bytes(bytes.fromhex(data[12:14]), 'little')
            elif datatype == 'u16':
                payload = int.from_bytes(bytes.fromhex(data[12:17]), 'little')
            elif datatype == 'u32':
                payload = int.from_bytes(bytes.fromhex(data[12:]), 'little')
            else:
                payload = int.from_bytes(bytes.fromhex(data[12:].replace(" ", "")), 'little')
            
            # Apply scaling if needed
            scale_value = convert_scale(scale)
            if scale_value != 1.0 or offset != 0:
                payload = payload * scale_value + offset
                payload = round(payload, 2)
            
            # Return with unit if it exists and isn't NA or ASCII
            return (payload, None if pd.isna(unit) or unit == 'ASCII' else unit)
        
        except ValueError as e:
            print(f"Error processing data for SDO {hex(sdo_int)}: {e}")
            return (None, None)
    except Exception as e:
        print(f"CAN communication error for SDO {sdo}: {e}")
        return (None, "CAN_ERROR")

def get_user_selection() -> tuple[bool, bool]:
    """
    Create GUI popup for user selection
    Returns tuple of (dtc_counts_selected, fault_logs_selected)
    """
    root = tk.Tk()
    root.title("DOT Reader Options")
    root.geometry("300x150")
    
    # Variables to store checkbox states
    dtc_counts = tk.BooleanVar()
    fault_logs = tk.BooleanVar()
    
    # Create and pack checkboxes
    ttk.Checkbutton(root, text="Until DTC Counts", variable=dtc_counts).pack(pady=10)
    ttk.Checkbutton(root, text="Only Fault Logs", variable=fault_logs).pack(pady=10)
    
    # Result variables
    result = [False, False]
    
    def on_submit():
        result[0] = dtc_counts.get()
        result[1] = fault_logs.get()
        root.quit()
        
    ttk.Button(root, text="Submit", command=on_submit).pack(pady=20)
    
    root.mainloop()
    root.destroy()
    return result[0], result[1]

def get_output_filename(dtc_counts: bool, fault_logs: bool) -> str:
    """Generate output filename based on selection and date"""
    today = datetime.datetime.now().strftime('%d%m%y')
    
    # Determine base filename based on selection
    if dtc_counts and not fault_logs:
        base = f"InitReadTillCounts_{today}"
    elif fault_logs and not dtc_counts:
        base = f"InitReadFaultLogs_{today}"
    else:
        base = f"InitRead_SDO_{today}"
    
    # Find existing files with same base name
    existing = glob.glob(f"{base}_*.txt")
    count = len(existing) + 1
    
    return f"{base}_{count}.txt"

def create_progress_window() -> tuple[tk.Tk, tk.Label]:
    """Create progress window with label"""
    prog_window = tk.Tk()
    prog_window.title("DOT Reading Progress")
    prog_window.geometry("400x100")
    
    progress_label = tk.Label(prog_window, text="Initializing...", pady=20)
    progress_label.pack()
    
    prog_window.update()
    return prog_window, progress_label

def select_excel_file() -> Optional[str]:
    """Open file dialog to select Excel file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select DOT Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")],
        initialdir=os.path.dirname(os.path.abspath(__file__))
    )
    
    root.destroy()
    return file_path if file_path else None

def initReadSDO(excel_file: str, dtc_counts: bool, fault_logs: bool) -> None:
    """Initialize and read all SDOs from the Excel file."""
    table = pt.PrettyTable()
    table.field_names = ['Object ID', 'Param Name', 'Result']
    try:
        print("\nStarting DOT reading process...")
        
        # Read DOT configuration from selected Excel file
        print(f"Reading DOT from Excel file: {excel_file}")
        data = pd.read_excel(excel_file, header=1)
        
        # Create progress window after selection
        prog_window, progress_label = create_progress_window()
        
        # Filter data based on user selection
        if dtc_counts and not fault_logs:
            data = data[data['Object Id'].astype(str).apply(
                lambda x: int(x, 16) if x != 'Reserved' else 0) <= 0x5013]
        elif fault_logs and not dtc_counts:
            data = data[data['Object Id'].astype(str).apply(
                lambda x: int(x, 16) if x != 'Reserved' else 0) >= 0x5100]

        output_file = get_output_filename(dtc_counts, fault_logs)

        # Process each DOT entry
        data_subset = data[data['Object Id'] != 'Reserved']  # Filter out Reserved entries
        total_rows = len(data_subset)
        print(f"Processing {total_rows} SDO entries based on selection")
        
        for i in range(total_rows):
            current_row = data_subset.iloc[i]
            next_row = data_subset.iloc[i + 1] if i < total_rows - 1 else current_row

            # Skip reserved object IDs
            if current_row['Object Id'] == 'Reserved':
                continue

            while True:  # Retry loop
                progress_text = f"Reading Object ID: {current_row['Object Id']}\n{current_row['Parameter Name']}"
                progress_label.config(text=progress_text)
                prog_window.update()

                print(f"\nReading SDO {i}: {current_row['Object Id']} - {current_row['Parameter Name']}")
                # Read and process SDO
                result, unit = readData(
                    VCM, 
                    current_row['Object Id'],
                    current_row['Data Type'],
                    current_row['RD Access Level'],
                    current_row['Range'],
                    current_row['Offset'],
                    current_row['Scale'],
                    current_row['Unit']
                )
                
                if unit == "CAN_ERROR":
                    prog_window.withdraw()  # Hide progress window
                    if show_error_dialog(current_row['Object Id'], current_row['Parameter Name']):
                        prog_window.deiconify()  # Show progress window
                        continue  # Retry
                    else:
                        prog_window.destroy()
                        print("\nOperation stopped by user after CAN error")
                        return
                
                break  # Success, exit retry loop
            
            # Format result string with unit if available
            result_str = str(result) if unit is None else f"{result} {unit}"
            print(f"Result: {result_str}")
            
            # Add to table with divider if next is Reserved
            is_divider = next_row["Object Id"] == 'Reserved'
            table.add_row([current_row['Object Id'], current_row['Parameter Name'], result_str], 
                         divider=is_divider)
        
        # Update progress for completion
        progress_label.config(text="SDO Reading Complete!")
        prog_window.update()
        prog_window.after(2000, prog_window.destroy)  # Close after 2 seconds
        
        # Save results to dynamically named file
        print("\n\nSaving results to file...")
        with open(output_file, 'w') as f:
            f.write(table.get_string())
        
        print(f"\nSDOs read and saved to: {output_file}")
        print(f"Total of {total_rows} SDOs read")
        
    except Exception as e:
        if 'prog_window' in locals():
            prog_window.destroy()
        print(f"Error during SDO reading: {e}")
        raise

def main():
    """Main execution function with command line argument support"""
    parser = argparse.ArgumentParser(description='DOT Reading Tool')
    parser.add_argument('--file', type=str, help='Path to DOT Excel file')
    parser.add_argument('--dtc', type=str, help='Read until DTC counts')
    parser.add_argument('--fault', type=str, help='Read only fault logs')
    args = parser.parse_args()

    if args.file:
        # Skip file selection if provided via arguments
        excel_file = args.file
        dtc_counts = args.dtc.lower() == 'true'
        fault_logs = args.fault.lower() == 'true'
    else:
        # Use interactive mode if no arguments
        excel_file = select_excel_file()
        if not excel_file:
            print("No file selected. Operation cancelled.")
            return
        dtc_counts, fault_logs = get_user_selection()

    initReadSDO(excel_file, dtc_counts, fault_logs)

if __name__ == '__main__':
    main()
