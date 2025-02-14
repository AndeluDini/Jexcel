import pandas as pd
import json
import jexcel.header as hd
from jexcel.excel_collector import ExcelCollector

def excel_to_json(excel_file, header_row = 0, data_row = -1, start_col = 0):

    skiprows = []
    if data_row > header_row:
        skiprows = list(range(header_row + 1, data_row))

    df = pd.read_excel(excel_file, header=header_row, skiprows=skiprows, dtype=str)
    df = df.iloc[:, start_col:]
    df = df.where(pd.notnull(df), None)

    headers, root = hd.parse_headers(df)
    
    excel_collector = ExcelCollector(df, headers)
    result = excel_collector.parse()

    return json.dumps(result, indent=4, default=str)