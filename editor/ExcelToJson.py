import sys
import os
import configparser
import pandas as pd
import subprocess

# Load configuration from Config.ini
config = configparser.ConfigParser()
config.read('Config.ini')

# Get the Excel file path from the Config.ini
excel_file_path = config['Paths']['ExcelManager_ToJson']  # Reads the path from the config file

# Add the parent folder to sys.path to ensure jexcel can be imported
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'jexcel')))

# Load the Excel file into a DataFrame
df = pd.read_excel(excel_file_path)

# Function to run jexcel for each row
def run_jexcel_for_each_row(row):
    excel_file = row['Excel']
    json_file = row['Json']
    header_row = row['header_row']
    data_row = row['data_row']
    start_col = row['start_col']
    
    # Run jexcel with the given parameters using subprocess
    subprocess.run([
        'python', '-m', 'jexcel', excel_file, 
        '-o', json_file,
        '-hr', str(header_row),
        '-dr', str(data_row),
        '-sc', str(start_col)
    ])

# Iterate through the rows and run jexcel for each one
for index, row in df.iterrows():
    run_jexcel_for_each_row(row)