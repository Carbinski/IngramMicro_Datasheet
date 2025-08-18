import pandas as pd
import numpy as np
from dotenv import load_dotenv
from my_calendar import My_Calendar
import time
import os
import re
import sys

# length of first and second descriptions
LEN_DESC_1 = 31
LEN_DESC_2 = 35
START_ROW = 14
MAX_SHEET_LENGTH = 1048576

cal = My_Calendar()

# Index, Name, Formula
FORMULAS = [
    (7, "WEEKS OH+OO", "=IF(M{row_num}=0,0,((E{row_num}+F{row_num})/M{row_num}))*4"),
    (9, "$ On Hand", "=G{row_num}*I{row_num}"),
    (10, "$ On Order", "=F{row_num}*I{row_num}"),
    (11, "$ MTD", "=O{row_num}*I{row_num}"),
    (12, "Trending 2 Month RR", "=(((O{row_num}/(ROUNDDOWN(($A$13-$C$13),0)/7))*$F$13)+P{row_num})/2"),
    (13, "Average Prev 2 Momth RR", "=(P{row_num}+Q{row_num})/2"),
    (21, "YTD Sales", "=U{row_num}*I{row_num}")
    ]

# max length of item ids
ID_LENGTH = 6

def is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def convert_accounting_number(num_str):
    """
    Converts accounting-style negative numbers (e.g., '1-') 
    into normal negative numbers (e.g., '-1').
    If not accounting format, returns as int/float unchanged.
    """
    # Remove whitespace
    num_str = num_str.strip()

    # Check if it ends with a minus sign
    if num_str.endswith('-') and num_str[:-1].replace('.', '', 1).isdigit():
        return -float(num_str[:-1]) if '.' in num_str else -int(num_str[:-1])
    
    # Otherwise, try normal number conversion
    try:
        return float(num_str) if '.' in num_str else int(num_str)
    except ValueError:
        return num_str


def add_images(worksheet):
    try:
        worksheet.insert_image('A1', 'Logos/UAG.png')
    except FileNotFoundError:
        print("Warning: Logos/UAG.png not found, skipping image insert.")

    try:
        worksheet.insert_image('A7', 'Logos/Ingram Micro.png')
    except FileNotFoundError:
        print("Warning: Logos/Ingram Micro.png not found, skipping image insert.")


def parse_line_by_format(line, format):
        idx = 0
        fields = []
        
        # deliniate the first 4 fields
        for part in format:
            if isinstance(part, int):
                fields.append(line[idx:idx + part].strip())
                idx += part
            elif isinstance(part, str):
                expected = part
                actual = line[idx:idx + len(expected)]
                # if actual != expected:
                #     print(f"Expected delimiter '{expected}' but got '{actual}' at position '{idx}'")
                idx += len(expected)
            else:
                raise TypeError("Format list must contain only integers and strings.")
            
        # split up until the next custom field
        remaining = line[idx:148]
        remaining = remaining.replace("*", "")

        extra_parts = re.split(r'\s+', remaining.strip())

        # Add remaining items until total fields == 13. This eliminates the extra un-used field
        while len(fields) < 13 and extra_parts:
            curr = extra_parts.pop(0)

            # Remove the un-used ['N'] field if it is in the item
            if curr == "N":
                curr = extra_parts.pop(0)
            elif re.fullmatch(r'[-+]?\d+(?:\.\d+)?-?', curr):
                # plain numeric or accounting-style negative
                curr = convert_accounting_number(curr)
            else:
                pass

            fields.append(curr)

        # add the the next custom field
        idx = 148
        next = line[idx:idx + LEN_DESC_2]
        fields.append(next.strip())
        idx += LEN_DESC_2

        # Get remaining items
        remaining = line[idx:]
        # clean up remaining data
        remaining = re.sub(r'[*%`]', '', remaining)
        extra_parts = re.split(r'\s+', remaining.strip())

        # Add remaining items
        while extra_parts:
            curr = extra_parts.pop(0)
            if re.fullmatch(r'[-+]?\d+(?:\.\d+)?-?', curr):
                # plain numeric or accounting-style negative
                curr = convert_accounting_number(curr)
            else:
                pass
            
            fields.append(curr)

        return fields


def define_formats(workbook, worksheet):
    # individual formats
    currency_format = workbook.add_format({'num_format': '$#,##0.00_-'})
    dash_format = workbook.add_format({'num_format': '#,##0;-#,##0;" - "', 'align': 'center', 'valign': 'center'})
    red_text = workbook.add_format({'font_color': 'red', 'align': 'center', 'valign': 'center'})

    centered_format = workbook.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })
    left_format = workbook.add_format({
        'align': 'left',
        'valign': 'vcenter',
        'num_format': '@' 
    })

    # combined formats
    currency_red = workbook.add_format({'num_format': '$#,##0.00', 'font_color': 'red'})
    currency_blue = workbook.add_format({'num_format': '$#,##0.00', 'font_color': 'blue'})

    dash_red = workbook.add_format({'num_format': '#,##0;-#,##0;" - "', 'font_color': 'red', 'align': 'center', 'valign': 'center'})
    dash_blue = workbook.add_format({'num_format': '#,##0;-#,##0;" - "', 'font_color': 'blue', 'align': 'center', 'valign': 'center'})

    # setting typography formatting
    worksheet.set_column("A:A", 9.17, centered_format)
    worksheet.set_column("B:B", 64.17, left_format)
    worksheet.set_column("C:C", 14.17, left_format)
    worksheet.set_column("D:D", 7.17, red_text)
    worksheet.set_column("E:F", 8.17, dash_blue)
    worksheet.set_column("G:G", 15.17, dash_blue)
    worksheet.set_column("H:H", 11.17, dash_red)
    worksheet.set_column("I:K", 15.17, currency_format)
    worksheet.set_column("L:L", 15.17, currency_red)
    worksheet.set_column("M:M", 15.17, dash_format)
    worksheet.set_column("N:N", 15.17, dash_blue)
    worksheet.set_column("O:T", 9.17, dash_format)
    worksheet.set_column("U:U", 11.17, dash_red)
    worksheet.set_column("V:V", 11.17, currency_blue)


def set_headers(workbook, worksheet):
    months = cal.get_relative_months()

    headers_row1 = [
        "", "", "MFG.", "Status", "Units", "Units on", "Balance", "WEEKS", "", "", "", "", "Trending 2", "Average Prev", "MTD", "", "", "", "", "", "YTD Unit", "YTD"
    ]
    headers_row2 = [
        "IM SKU#", "Product Description", "P/N", "Code", "Avail", "Order", "On Hand", "OH + OO", "Unit Cost",
        "$ On Hand", "$ On Order", "$ MTD", "Month RR", "2 Month RR", "Unit Sales", f'{months[-1]}', f'{months[-2]}', f'{months[-3]}', f'{months[-4]}', f'{months[-5]}', "Sales", "Sales"
    ]

    top_header_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1, 
        'border': 2,
        'border_color': 'black'
    })
    bottom_header_fmt = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1,
        'border': 2,
        'border_color': 'black'
    })
    top_header_fmt_blue = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1, 
        'border': 2,
        'border_color': 'black',
        'font_color': 'blue'
    })
    bottom_header_fmt_blue = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1,
        'border': 2,
        'border_color': 'black',
        'font_color': 'blue'
    })
    top_header_fmt_red = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1, 
        'border': 2,
        'border_color': 'black',
        'font_color': 'red'
    })
    bottom_header_fmt_red = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'underline': 1,
        'border': 2,
        'border_color': 'black',
        'font_color': 'red'
    })


    # Remove border between top and bottom row
    for fmt in [top_header_fmt, top_header_fmt_blue, top_header_fmt_red]:
        fmt.set_bottom(0)
    for fmt in [bottom_header_fmt, bottom_header_fmt_blue, bottom_header_fmt_red]:
        fmt.set_top(0)

    red_cols = [3, 7, 11, 20]
    blue_cols = [4, 5, 6, 13, 21]

    for i in range(len(headers_row1)):
        if i in red_cols:
            worksheet.write(START_ROW - 1, i, headers_row1[i], top_header_fmt_red)
            worksheet.write(START_ROW, i, headers_row2[i], bottom_header_fmt_red)  
        elif i in blue_cols:
            worksheet.write(START_ROW - 1, i, headers_row1[i], top_header_fmt_blue)
            worksheet.write(START_ROW, i, headers_row2[i], bottom_header_fmt_blue)
        else:
            worksheet.write(START_ROW - 1, i, headers_row1[i], top_header_fmt)  
            worksheet.write(START_ROW, i, headers_row2[i], bottom_header_fmt)  


def add_extra_info(workbook, worksheet):
    # Styling
    underline_fmt = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'underline': 1
    })
    right_align_fmt = workbook.add_format({
        'align': 'right',
        'valign': 'right',
        'font_color': 'black'
    })
    italics_blue_fmt = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'italic': True,
        'bold': True,
        'font_color': 'blue'
    })
    italics_green_fmt = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'italic': True,
        'bold': True,
        'font_color': '#32CD32'
    })
    italics_red_fmt = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'italic': True,
        'bold': True,
        'font_color': 'red'
    })
    italics_blue_currency_fmt = workbook.add_format({
        'align': 'center',
        'valign': 'center',
        'font_color': 'blue',
        'num_format': '$#,##0.00_-'
    })

    # Report Date
    date = cal.get_report_date_str()
    for i, info in enumerate(["Report", "Date", date]):
        worksheet.write(START_ROW - 4 + i, 0, info, underline_fmt)

    # First Day of This / Next Fiscal Month
    next_month = cal.get_next_fiscal_month()
    this_month = cal.get_this_fiscal_month()

    worksheet.write(START_ROW - 3, 1, "First Day of Next Fiscal Month:", right_align_fmt)
    worksheet.write(START_ROW - 2, 1, "First Day of Fiscal Month:", right_align_fmt)

    worksheet.write(START_ROW - 3, 2, next_month, right_align_fmt)
    worksheet.write(START_ROW - 2, 2, this_month, right_align_fmt)

    # Weeks in Month
    worksheet.write(START_ROW - 2, 4, "Weeks in Month:", right_align_fmt)
    worksheet.write(START_ROW - 2, 5, f"=(C{START_ROW - 2}-C{START_ROW - 1})/7", right_align_fmt)

    # Reporting Week
    worksheet.write(START_ROW - 2, 6, "Reporting Week:", right_align_fmt)
    worksheet.write(START_ROW - 2, 7, f"=(A{START_ROW - 1}-C{START_ROW - 1})/7", right_align_fmt)

    # On Hand
    worksheet.write(START_ROW - 9, 6, "$-On Hand", italics_blue_fmt)
    worksheet.write(START_ROW - 8, 6, f"=SUM(J{START_ROW + 2}:J{MAX_SHEET_LENGTH})", italics_blue_currency_fmt)

    # On Order
    worksheet.write(START_ROW - 9, 8, "$-On Order", italics_blue_fmt)
    worksheet.write(START_ROW - 8, 8, f"=SUM(K{START_ROW + 2}:K{MAX_SHEET_LENGTH})", italics_blue_currency_fmt)

    # OH + On Order
    worksheet.write(START_ROW - 9, 9, "$-OH + $-On Order", italics_blue_fmt)
    worksheet.write(START_ROW - 8, 9, f"=G{START_ROW - 7}+I{START_ROW - 7}", italics_blue_currency_fmt)

    # MTD
    worksheet.write(START_ROW - 9, 11, "$-MTD", italics_red_fmt)
    worksheet.write(START_ROW - 8, 11, f"=SUM(L{START_ROW + 2}:L{MAX_SHEET_LENGTH})", italics_blue_currency_fmt)

    # RUN RATE
    worksheet.write(START_ROW - 10, 12, "$-EST MONTHLY", italics_green_fmt)
    worksheet.write(START_ROW - 9, 12, "RUN RATE", italics_blue_fmt)
    worksheet.write(START_ROW - 8, 12, f"=(L{START_ROW - 7}/H{START_ROW - 1})*F{START_ROW - 1}", italics_blue_currency_fmt)


def write_equations(df):
    # creates the columns for the derived data we're going to add
    for index, name, _ in FORMULAS:
        df.insert(index, name, '')

    df.iloc[:, 2] = df.iloc[:, 2].astype(str)

    with pd.ExcelWriter(f"{os.getenv('NAME')} {cal.get_report_date_str().replace('/', '-')}.xlsx", engine="xlsxwriter") as writer:
        workbook  = writer.book
        df.to_excel(writer, index=False, sheet_name=os.getenv("NAME"), startrow = START_ROW)

        worksheet = writer.sheets[os.getenv("NAME")]


        left_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'num_format': '@' 
        })
        
        # Force ALL data rows in column C (index 2) to be written as text
        # Pandas writes the header at row START_ROW, and data begins at row START_ROW + 1 (0-indexed)
        for r, val in enumerate(df.iloc[:, 2], start=START_ROW + 1):
            worksheet.write_string(r, 2, str(val), left_format)

        for i in range(START_ROW, START_ROW + len(df)):
            row_num = i + 2  # Excel rows are 1-indexed and row 1 is the header

            for index, _, formula_template in FORMULAS:
                # turns formula into properly formated formula
                formula = formula_template.format(row_num=row_num)
                worksheet.write_formula(i + 1, index, formula)

        set_headers(workbook, worksheet)
        define_formats(workbook, worksheet)
        add_extra_info(workbook, worksheet)
        add_images(worksheet)


def clean_spreadsheet(df):

    # merge descriptions
    df[1] = df[1] + ' ' + df[13]

    months = cal.get_relative_months()

    rename_map = {
        0: 'IM SKU#',
        1: 'Product Description',
        15: 'MFG. P/N',
        11: 'Status Code',
        20: 'Units Avail',
        23: 'Units on Order',
        27: 'Balance On Hand',
        5: 'Unit Cost',
        30: 'MTD Unit Sales',
        32: f'{months[-1]}',
        33: f'{months[-2]}',
        34: f'{months[-3]}',
        35: f'{months[-4]}',
        36: f'{months[-5]}',
        31: 'YTD Unit Sales'
    }


    # Renaming relevant columns by label:
    df = df.rename(columns=rename_map)

    # Drop unnecessary columns based on labels
    df = df.drop(columns=[2,3,4,6,7,8,12,13,14,16,17,18,19,21,22,24,25,26,28,29])

    # re-ordering remaining columns
    desired_order = list(rename_map.values())
    df = df[desired_order]

    # making sure this is a string to prevent future issues
    df["MFG. P/N"] = df["MFG. P/N"].astype(str)

    return df

def run_script():
    # Get the directory where the executable resides (not the temp bundle directory)
    if getattr(sys, 'frozen', False):
        # If bundled by PyInstaller, get the executable's directory
        base_path = os.path.dirname(sys.executable)
    else:
        # If running as script, get the script's directory
        base_path = os.path.dirname(os.path.abspath(__file__))

    # Don't change working directory yet - first load the .env file
    env_path = os.path.join(base_path, '.env')

    # Load .env file from the executable's directory
    if os.path.exists(env_path):
        load_dotenv(env_path)
        print(f"Loaded .env from: {env_path}")
    else:
        print(f"Warning: .env file not found at {env_path}")
        # Optionally create a default .env file
        with open(env_path, 'w') as f:
            f.write('RAW_FILE=A764Y.TXT\nNAME=A764Y\n')
        print(f"Created default .env file at {env_path}")
        load_dotenv(env_path)

    # Now change to the executable's directory
    os.chdir(base_path)
    print("Current directory:", os.getcwd())

    print("Setting calendar")
    fiscal_periods_raw = os.getenv("FISCAL_PERIODS")
    if not fiscal_periods_raw:
        print ("FISCAL_PERIODS not set, using default")
    else:
        cal.set_calendar(fiscal_periods_raw)

    # Get the raw file path - it should be relative to the executable's directory
    raw_file = os.getenv("RAW_FILE")
    if not raw_file:
        raise ValueError("RAW_FILE not found in environment variables")

    # Check if the file exists
    raw_file_path = os.path.join(base_path, raw_file)
    if not os.path.exists(raw_file_path):
        raise FileNotFoundError(f"Data file not found: {raw_file_path}")

    print("Current directory:", os.getcwd())
    rows = []
    format = [ID_LENGTH, " "*8, LEN_DESC_1, "", 4, " ", 6]
    with open(os.getenv("RAW_FILE"), "r") as file:
        for line in file:
            if not line.strip():
                continue
            line = line.lstrip()
            # stip and split based on spaces
            row = parse_line_by_format(line, format)
            rows.append(row)
    df = pd.DataFrame(rows)
    df = clean_spreadsheet(df)
    write_equations(df)


def input_with_timeout(prompt, timeout=10, default="0"):
    """Return the user's input if entered within `timeout` seconds,
    otherwise return `default`. Works on both POSIX and Windows.
    """
    if os.name == "nt":  # Windows
        import msvcrt
        sys.stdout.write(prompt)
        sys.stdout.flush()
        buf = ""
        start = time.time()
        while True:
            if msvcrt.kbhit():
                ch = msvcrt.getwche()   # echo the character
                if ch in ("\r", "\n"):
                    return buf
                elif ch == "\x03":  # Ctrl-C
                    raise KeyboardInterrupt
                elif ch == "\x08":  # backspace
                    buf = buf[:-1]
                else:
                    buf += ch
            if (time.time() - start) > timeout:
                sys.stdout.write("\n")  # move to next line after timeout
                sys.stdout.flush()
                return default
            time.sleep(0.01)
    else:  # POSIX (Linux / macOS / etc.)
        import select
        sys.stdout.write(prompt)
        sys.stdout.flush()
        rlist, _, _ = select.select([sys.stdin], [], [], timeout)
        if rlist:
            return sys.stdin.readline().rstrip("\n")
        else:
            sys.stdout.write("\n")
            sys.stdout.flush()
            return default


def main():

    print(" Welcome! ") 
    print(" Choose one of the following commands: ")
    print(" 0. Run program for most recent Sunday")
    print(" 1. Set report date (If it is not on Sunday it may break)")
    outcome = input_with_timeout("Enter command (defaults to 0 after 10s): ", timeout=10, default="0")

    # normalize empty string just in case
    if outcome == "":
        outcome = "0"

    if outcome == "1":
        print("Enter report date (MM/DD/YYYY)")
        date = input("Set report date: ")
        cal.set_report_date(date)
        print("Sucessfully set custom report date. Now running script!")
    elif outcome != "0":
        print("Invalid input, please run program again and use valid input.")
        raise ValueError("Invalid input")
    
    run_script()


if __name__ == "__main__":
    main()