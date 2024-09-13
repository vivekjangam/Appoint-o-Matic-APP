import datetime
import xlwings as xw

from .GeneralForBoth import *


def process_export_wb(export_path, source_ws, app):
    """Processes an export workbook by adding columns, calculating appointment dates, and filling missing dates with the most frequent date.

    This function takes an export workbook path, the source worksheet containing relevant data, and the xlwings App object.
    It performs the following actions:

    1. Opens the export workbook and gets the worksheet (with informative printing).
    2. Adds columns for "Location Code" and "Appointment Date".
    3. Gets column addresses for relevant data in the export worksheet (with printing).
    4. Extracts location code from "Ship to location" using formula and converts to values (with printing).
    5. Gets column addresses for "Amazon Code" and "Lead Time Date" in the source worksheet (with printing).
    6. Calculates the relative index number for VLOOKUP based on column positions (with printing).
    7. Uses VLOOKUP to get appointment dates based on location code and source data (with printing).
    8. Converts appointment date column to values (with printing).
    9. Adds a column for "Frequent Date".
    10. Calculates the most frequent date using MODE.SNGL and prints the result.
    11. Iterates through rows in the appointment date column:
        - Prints the existing value (for debugging).
        - Fills missing dates with the most frequent date (with printing).
    12. Converts frequent date column to values.
    13. Deletes the frequent date column.
    14. Formats the appointment date column to 'yyyy-mm-dd' format.
    15. Prints a completion message.

    Args:
        export_path (str): The path to the export workbook.
        source_ws (xw.Sheet): The worksheet from the source workbook containing Amazon code and lead time date information.
        app (xw.App): The xlwings app object.

    Returns:
        xw.Workbook: The processed export workbook object.
    """
    
    # Open the export workbook and get the worksheet
    wb_export = app.books.open(export_path)
    ws_export = wb_export.sheets["Sheet1"]
    last_row = ws_export.range(1, 1).end('down').row
    
    print(f"Opened export workbook: {export_path}")
    print(f"Last row in export worksheet: {last_row}")

    # Add columns for location code and appointment date
    add_column(ws_export, "Location Code")
    add_column(ws_export, "Appointment Date")

    # Get column addresses for relevant data in the export worksheet
    location_code_col = get_column_address(ws_export, "Location Code")
    ship_location_col = get_column_address(ws_export, "Ship to location")
    appointment_date_col = get_column_address(ws_export, "Appointment Date")
    
    print(f"Location code column address: {location_code_col}")
    print(f"Ship to location column address: {ship_location_col}")
    print(f"Appointment date column address: {appointment_date_col}")

    ws_export.range(f"{location_code_col}2:{location_code_col}{last_row}").formula = f'=LEFT({ship_location_col}2, FIND("-", {ship_location_col}2) - 2)'
    convert_to_values(ws_export,location_code_col,last_row)
    print(f"Location code column populated and converted to values.")
    
    amazon_code_col = get_column_address(source_ws, "Amazon Code")
    lead_time_date_col = get_column_address(source_ws, "Lead Time Date")

    lead_time_date_col_number = next(i + 1 for i, cell_value in enumerate(source_ws.range("1:1").value) if cell_value == "Lead Time Date")
    amazon_code_col_number = next(i + 1 for i, cell_value in enumerate(source_ws.range("1:1").value) if cell_value == "Amazon Code")
    relative_index_number = (lead_time_date_col_number - amazon_code_col_number) + 1

    print(f"Amazon code column address: {amazon_code_col}")
    print(f"Lead time date column address: {lead_time_date_col}")
    print(f"Relative index number for VLOOKUP: {relative_index_number}")
    
    ws_export.range(f"{appointment_date_col}2:{appointment_date_col}{last_row}").formula = f"=VLOOKUP({location_code_col}:{location_code_col},'[{source_ws.book.name}]{source_ws.name}'!${amazon_code_col}:${lead_time_date_col},{relative_index_number},0)"
    convert_to_values(ws_export,appointment_date_col,last_row)
    print(f"Appointment date column populated and converted to values.")
    
    # Process frequent date
    add_column(ws_export, "Frequent Date")
    frequent_date_col = get_column_address(ws_export, "Frequent Date")
    print(f"Frequent date column address: {frequent_date_col}")
    
    ws_export.range(f"{frequent_date_col}2:{frequent_date_col}{last_row}").formula = f"=MODE.SNGL({appointment_date_col}:{appointment_date_col})"
    mod_date = ws_export.range(f"{frequent_date_col}2").value
    print(f"Most frequent date (MODE.SNGL result): {mod_date}")
    
    for row in range(2, last_row + 1):
        print(ws_export.range(f"{appointment_date_col}{row}").value)
        if ws_export.range(f"{appointment_date_col}{row}").value == None:
            ws_export.range(f"{appointment_date_col}{row}").value = mod_date
            print("insidered", ws_export.range(f"{appointment_date_col}{row}").value )
    
    convert_to_values(ws_export,frequent_date_col,last_row)
    ws_export.range(f"{frequent_date_col}:{frequent_date_col}").api.Delete()
    ws_export.range(f"{appointment_date_col}2:{appointment_date_col}{last_row}").number_format = 'yyyy-mm-dd'

    print("Export workbook processing completed.")
    
    return wb_export