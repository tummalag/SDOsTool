"""Common constants for DOT Search Tool"""

# Excel file name
DOT_EXCEL_FILE = "DOT_22_DOT_ST.xlsx"

# ECU Node Addresses
VCM = 0x635
MCU = 0x626

# Response Addresses
VCM_RESP = 0x5B5
MCU_RESP = 0x5A6

# SDO Commands
SAVE_NVM = 0x2107     # Save Non-Volatile Memory
VCM_ver_num = 0x2008  # VCM version number SDO
MCU_ver_num = 0x2185  # MCU version number SDO
BMS_ver_num = 0x2205  # BMS version number SDO
