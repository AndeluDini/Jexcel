import jexcel
import sys
import os
import configparser
import pandas as pd
import subprocess
from jexcel import core


def pause_exit():
    input("Press Enter to exit...")
    sys.exit(1)

# Helper function to handle relative vs absolute paths
def make_absolute(path_str, base_dir):
    """
    If path_str is already absolute, return it as-is.
    Otherwise, join it with base_dir to form an absolute path.
    """
    if os.path.isabs(path_str):
        return path_str
    return os.path.abspath(os.path.join(base_dir, path_str))

# 1. Determine the base directory where the .exe is located
if getattr(sys, 'frozen', False):
    # Running in a bundle
    base_dir = os.path.dirname(sys.executable)
else:
    # Running from source: use the parent folder of the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.dirname(script_dir)
exe_base_dir = base_dir 
# exe_base_dir = os.path.dirname(sys.executable)

# 2. Read the config file (relative to exe_base_dir if needed)
config_path = 'JexcelConfig.ini'
config_path_abs = make_absolute(config_path, exe_base_dir)

if not os.path.exists(config_path_abs):
    print(f"Error: Configuration file '{config_path_abs}' not found.")
    pause_exit()

config = configparser.ConfigParser()
config.read(config_path_abs)

try:
    excel_file_path = config['Paths']['ExcelManager_ToJson']
except KeyError as e:
    print(f"Error: Missing configuration for {e} in {config_path_abs}.")
    pause_exit()

# Convert excel_file_path to an absolute path
excel_file_path_abs = make_absolute(excel_file_path, exe_base_dir)

if not os.path.exists(excel_file_path_abs):
    print(f"Error: Excel file not found at: {excel_file_path_abs}")
    pause_exit()

# # Make sure Python can find jexcel (adjust as needed)
# sys.path.insert(
#     0, 
#     os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'jexcel'))
# )
# Only add the external jexcel path when not frozen.
if not getattr(sys, 'frozen', False):
    sys.path.insert(
        0, 
        os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'jexcel'))
    )

# Load the Excel file into a DataFrame
try:
    df = pd.read_excel(excel_file_path_abs)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    pause_exit()

def run_jexcel_for_each_row(row):
    try:
        # Convert each path to absolute
        excel_file = make_absolute(row['Excel'], exe_base_dir)
        json_file = make_absolute(row['Json'], exe_base_dir)

        header_row = row['header_row']
        data_row = row['data_row']
        start_col = row['start_col']

        if not os.path.exists(excel_file):
            print(f"Warning: Excel file for row not found: {excel_file}")
            return
        
        result = core.excel_to_json(excel_file, header_row, data_row, start_col)

        # If an output file is specified, write the result; otherwise, print it.
        if json_file:
            with open(json_file, 'w', encoding='utf-8') as f:
                f.write(result)
            print(f"JSON has been written to {json_file}")
        else:
            print(result)

    except Exception as e:
        print(f"Error processing row: {e}")

# Iterate through the rows and run jexcel for each one
for index, row in df.iterrows():
    run_jexcel_for_each_row(row)

input("Processing complete. Press Enter to exit...")
