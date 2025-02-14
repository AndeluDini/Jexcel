import argparse
import jexcel.core
from pathlib import Path

def main():
    
    parser = argparse.ArgumentParser(
        description="Convert an Excel file to JSON with customizable row and column settings."
    )

    # Positional argument: the input Excel file
    parser.add_argument(
        "excel_file",
        help="Path to the Excel file"
    )

    # Optional argument: output JSON file
    parser.add_argument(
        "-o", "--output",
        type=str,
        default=None,
        help="Path to the output JSON file"
    )

    # Optional argument for header_row
    parser.add_argument(
        "-hr", "--header_row",
        type=int,
        default=0,
        help="Specify the header row number (default: 0)"
    )

    # Optional argument for data_row
    parser.add_argument(
        "-dr", "--data_row",
        type=int,
        default=-1,
        help="Specify the first data row number (default: -1)"
    )

    # Optional argument for start_col
    parser.add_argument(
        "-sc", "--start_col",
        type=int,
        default=0,
        help="Specify the starting column number (default: 0)"
    )

    # Parse the arguments
    args = parser.parse_args()

    # Extract values
    excel_file = args.excel_file
    header_row = args.header_row
    data_row = args.data_row
    start_col = args.start_col
    out_path = args.output

    # implementation
    result = jexcel.core.excel_to_json(excel_file, header_row, data_row, start_col)

    # If an output path is provided, write the JSON to that file; otherwise, print to stdout
    if out_path:
        with open(Path(out_path), "w", encoding="utf-8") as f:
            f.write(result)
        print(f"JSON has been written to {out_path}")
    else:
        print(result)

if __name__ == "__main__":
    main()