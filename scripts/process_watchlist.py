#!/usr/bin/env python3
"""Read concatenated messages from stdin and write them into an Excel file.
Usage: osascript extract_mail.scpt | python3 process_watchlist.py --out /path/to/watchlist.xlsx
"""
import sys
import argparse
import pandas as pd
import io
import re
import csv
from datetime import datetime
import os

SEPARATOR = '---MSG---'

def parse_message(msg_text):
    """Parse message to extract recap date and tabular data."""
    lines = [l.rstrip('\r') for l in msg_text.splitlines()]
    if not lines:
        return None, pd.DataFrame()
    
    # Extract the date from "End-of-Day Recap - Price quotes for Tue, January 27, 2026"
    recap_date = None
    data_start_idx = 0
    for i, line in enumerate(lines):
        if 'End-of-Day Recap' in line and 'Price quotes for' in line:
            # Extract date part: "Tue, January 27, 2026"
            match = re.search(r'Price quotes for (.+?)(?:\n|$)', line)
            if match:
                recap_date = match.group(1).strip()
            data_start_idx = i + 1
            break
    
    # Find the header row (contains "Symbol", "Latest", etc.)
    header_idx = None
    for i in range(data_start_idx, len(lines)):
        if lines[i].strip() == 'Symbol':
            header_idx = i
            break
    
    if header_idx is None:
        # No header found; try to parse as-is
        return recap_date, pd.DataFrame()
    
    # Collect header from consecutive lines after "Symbol"
    header = ['Symbol']
    data_start_line = header_idx + 1
    
    # Collect header columns (they appear as separate lines)
    i = header_idx + 1
    while i < len(lines) and lines[i].strip() and not lines[i][0].isdigit() and not lines[i][0] in ('^', '$', 'N'):
        col = lines[i].strip()
        if col and col not in header:
            header.append(col)
        else:
            break
        i += 1
    
    # Data rows start after header columns
    data_start_line = i
    
    # Collect data rows: each row is N consecutive lines where N = len(header)
    rows = []
    current_row = []
    for i in range(data_start_line, len(lines)):
        line = lines[i].strip()
        if not line:
            continue
        
        # If we have enough columns for a row, start a new one
        if len(current_row) == len(header):
            rows.append(current_row)
            current_row = []
        
        current_row.append(line)
        
        # Stop after ^USDCHF row is complete
        if len(current_row) == len(header) and current_row[0] == '^USDCHF':
            rows.append(current_row)
            break
    
    # Add the last row if it's complete and not already added
    if current_row and len(current_row) == len(header) and (not rows or current_row != rows[-1]):
        rows.append(current_row)
    
    # Create DataFrame with header
    if rows:
        try:
            print("Parsed rows:", file=sys.stderr)
            for r in rows:
                print(r, file=sys.stderr)
            df = pd.DataFrame(rows, columns=header)
            # Clean symbol names: remove $ and ^ characters
            if 'Symbol' in df.columns:
                df['Symbol'] = df['Symbol'].str.replace(r'[\$\^]', '', regex=True)
            # Convert numeric columns to float
            numeric_cols = ['Latest', 'Change', 'Open', 'High', 'Low', 'Volume']
            for col in numeric_cols:
                if col in df.columns:
                    print(f"Raw {col} values: ", df[col].tolist(), file=sys.stderr)
                    df[col] = pd.to_numeric(
                        df[col].astype(str).str.replace(',', '', regex=False).str.replace('s', '', regex=False).str.replace('+', '', regex=False), 
                        errors='coerce'
                    )
                    print(f"Converted {col} values: ", df[col].tolist(), file=sys.stderr)
            # Convert %Change: remove % prefix and convert to float (will be formatted as percentage in Excel)
            if '%Change' in df.columns:
                print("Raw %Change values: ", df['%Change'].tolist(), file=sys.stderr)
                df['%Change'] = pd.to_numeric(
                    df['%Change'].astype(str).str.replace('%', '', regex=False).str.replace(',', '', regex=False).str.replace('+', '', regex=False), 
                    errors='coerce'
                )
                print("Converted %Change values: ", df['%Change'].tolist(), file=sys.stderr)
        except Exception as e:
            print(f"Conversion error: {e}", file=sys.stderr)
            df = pd.DataFrame(rows)
    else:
        df = pd.DataFrame()
    
    return recap_date, df

def write_to_excel_by_month(dfs_with_dates, out_path):
    """Create separate Excel files per month, with worksheets per calendar week"""
    from openpyxl import Workbook
    
    # Get base directory from out_path
    base_dir = os.path.dirname(out_path)
    os.makedirs(base_dir, exist_ok=True)
    
    # Group dataframes by month and calendar week
    month_data = {}  # {month_name: {week_num: [(df, date_str, datetime), ...]}}
    
    for df, recap_date in dfs_with_dates:
        if not recap_date or df.empty:
            continue
        
        try:
            dt = datetime.strptime(recap_date, '%a, %B %d, %Y')
            month_name = dt.strftime('%B')  # e.g., 'January'
            day = dt.day
            # Calculate calendar week: days 1-7 = Week1, 8-14 = Week2, 15-21 = Week3, 22-28 = Week4, 29+ = Week5
            week_num = (day - 1) // 7 + 1
            
            if month_name not in month_data:
                month_data[month_name] = {}
            if week_num not in month_data[month_name]:
                month_data[month_name][week_num] = []
            month_data[month_name][week_num].append((df, recap_date, dt))
        except Exception as e:
            print(f"Date parse error for '{recap_date}': {e}", file=sys.stderr)
            continue
    
    # Create separate Excel file for each month
    for month_name in sorted(month_data.keys()):
        # Create filename: january.xlsx, february.xlsx, etc.
        month_file = os.path.join(base_dir, f'{month_name.lower()}.xlsx')
        
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        week_nums = sorted(month_data[month_name].keys())
        
        # Create one sheet per calendar week
        for week_num in week_nums:
            sheet_name = f'Week{week_num}'
            ws = wb.create_sheet(sheet_name)
            
            dfs_for_week = month_data[month_name][week_num]
            # Sort by date to ensure chronological order
            dfs_for_week_sorted = sorted(dfs_for_week, key=lambda x: x[2])
            
            current_col = 1
            
            # Write all data for this week in chronological order
            for df, recap_date, dt in dfs_for_week_sorted:
                if df.empty:
                    continue
                
                # Write header
                for col_num, col_name in enumerate(df.columns, start=0):
                    cell = ws.cell(row=1, column=current_col + col_num)
                    cell.value = col_name
                
                # Write data rows
                for row_num, row in enumerate(df.values, start=2):
                    for col_num, val in enumerate(row):
                        cell = ws.cell(row=row_num, column=current_col + col_num)
                        cell.value = val
                        
                        # Apply number formatting
                        col_name = df.columns[col_num]
                        if col_name == '%Change':
                            cell.number_format = '0.00%'
                        elif col_name == 'Time':
                            cell.number_format = 'h:mm AM/PM'
                        elif col_name in ['Latest', 'Change', 'Open', 'High', 'Low']:
                            cell.number_format = '0.00'
                        elif col_name == 'Volume':
                            cell.number_format = '0'
                
                # Move to next column block (add spacing between dates)
                current_col += len(df.columns) + 2
        
        wb.save(month_file)
        print(f'Created {month_file}', file=sys.stderr)

def write_to_excel(dfs_with_dates, out_path):
    # each dataframe in dfs is written to its own sheet named with the recap date
    from openpyxl.styles import numbers
    
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    mode = 'a' if os.path.exists(out_path) else 'w'
    try:
        with pd.ExcelWriter(out_path, engine='openpyxl', mode=mode) as writer:
            for df, sheet_name_base in dfs:
                # Use recap_date as sheet name, fallback to timestamp if not available
                if sheet_name_base:
                    sheet_name = sheet_name_base[:31]  # Excel sheet name max 31 chars
                else:
                    sheet_name = 'Watchlist_' + datetime.now().strftime('%Y%m%d_%H%M%S')
                    sheet_name = sheet_name[:31]
                
                if df.empty:
                    pd.DataFrame({'message': ['(empty)']}).to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Apply formatting to columns
                    ws = writer.sheets[sheet_name]
                    for col_num, col_name in enumerate(df.columns, start=1):
                        col_letter = ws.cell(row=1, column=col_num).column_letter
                        
                        if col_name == '%Change':
                            # Percentage format
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = '0.00%'
                        elif col_name == 'Time':
                            # Time format
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = 'h:mm AM/PM'
                        elif col_name in ['Latest', 'Change', 'Open', 'High', 'Low', 'Volume']:
                            # Number format (2 decimal places for most, 0 for Volume)
                            fmt = '0.00' if col_name != 'Volume' else '0'
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = fmt
    except Exception as e:
        # fallback: try writing single sheet if append mode fails
        with pd.ExcelWriter(out_path, engine='openpyxl', mode='w') as writer:
            for df, sheet_name_base in dfs:
                if sheet_name_base:
                    sheet_name = sheet_name_base[:31]
                else:
                    sheet_name = 'Watchlist_' + datetime.now().strftime('%Y%m%d_%H%M%S')
                    sheet_name = sheet_name[:31]
                
                if df.empty:
                    pd.DataFrame({'message': ['(empty)']}).to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Apply formatting to columns
                    ws = writer.sheets[sheet_name]
                    for col_num, col_name in enumerate(df.columns, start=1):
                        col_letter = ws.cell(row=1, column=col_num).column_letter
                        
                        if col_name == '%Change':
                            # Percentage format
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = '0.00%'
                        elif col_name == 'Time':
                            # Time format
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = 'h:mm AM/PM'
                        elif col_name in ['Latest', 'Change', 'Open', 'High', 'Low', 'Volume']:
                            # Number format (2 decimal places for most, 0 for Volume)
                            fmt = '0.00' if col_name != 'Volume' else '0'
                            for row_num in range(2, len(df) + 2):
                                ws[f'{col_letter}{row_num}'].number_format = fmt

def main():
    parser = argparse.ArgumentParser(description='Process watchlist messages from stdin and write to Excel')
    parser.add_argument('--out', '-o', default=os.path.expanduser('~/Documents/watchlist_summary.xlsx'), help='Output Excel path')
    args = parser.parse_args()

    data = sys.stdin.read()
    if not data.strip():
        print('No input received; no messages found.', file=sys.stderr)
        return

    parts = [p.strip() for p in data.split(SEPARATOR) if p.strip()]
    if not parts:
        print('No messages after splitting; exiting.', file=sys.stderr)
        return

    dfs = []
    for p in parts:
        recap_date, df = parse_message(p)
        dfs.append((df, recap_date))

    write_to_excel_by_month(dfs, args.out)
    print(f'Wrote {len(dfs)} message(s) to {args.out}', file=sys.stderr)

if __name__ == '__main__':
    main()
