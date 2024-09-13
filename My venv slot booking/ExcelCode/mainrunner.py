import datetime
import xlwings as xw
import os
from .GeneralForBoth import *
from .HolidayWsSpecific import *
from .exportProcess import *
import time
def main(input_master_file_path,consignment_file_path,output_folder_path):
    """Main function that automates processing lead times for Amazon appointments.

    This function opens a specified workbook, performs calculations and updates on the "Lead Times" worksheet,
    processes data for export, saves the updated workbook and a processed export workbook, and closes the application.

    The function performs the following actions:
        - Adds columns for "Total Lead Time", "Lead Time Date", and "Today Date" to the "Lead Times" worksheet.
        - Calculates total lead time based on existing data and adds formulas.
        - Updates total lead time based on the time of day.
        - Calculates lead time date based on total lead time and today's date.
        - Fetches holiday data from another worksheet and updates total lead time accordingly.
        - Processes data for export to a separate workbook.
        - Formats worksheets in both workbooks.
        - Saves the updated workbook and the processed export workbook with timestamps in their names.
        - Closes the Excel application.
    """
    
    # wb_path = r"D:\Master Ler\Consolidation GIT and Inventory\Amazone Slot Booking\Amazon Appointment Automation Rule.xlsx"
    # export_path = r"export processed.xlsx"
    
    wb_path = input_master_file_path
    export_path = consignment_file_path
    output_folder = output_folder_path
    
    print(f"Opening workbook: {wb_path}")
    app, wb = open_workbook(wb_path)
    ws_lead_times = wb.sheets["Lead Times"]
    ws_holiday_calendar = wb.sheets["Holiday Calendar"]
    print("Workbooks opened successfully.")
    
    # Add columns
    total_lead_time_col = add_column(ws_lead_times, "Total Lead Time")
    lead_time_date_col = add_column(ws_lead_times, "Lead Time Date")
    today_date_col = add_column(ws_lead_times, "Today Date")
    print("Added columns for Total Lead Time, Lead Time Date, and Today Date.")
    
    # Add formulas and date values
    today = datetime.date.today()
    add_date_values(ws_lead_times, today_date_col, ws_lead_times.used_range.last_cell.row, today)
    print(f"Today's date ({today}) added to Today Date column.")
    
    lead_time_col = get_column_address(ws_lead_times, "Lead Time")
    traffic_consideration_col = get_column_address(ws_lead_times, "Traffice consideration")
    add_formulas(ws_lead_times, [(total_lead_time_col, f"={lead_time_col}2+{traffic_consideration_col}2")], ws_lead_times.used_range.last_cell.row)
    print("Formulas added for Total Lead Time.")
    
    now = datetime.datetime.now()
    if now.hour >= 12:
        update_total_lead_time(ws_lead_times, total_lead_time_col, ws_lead_times.used_range.last_cell.row, 0.5)
        print("Updated Total Lead Time for afternoon processing (added 0.5 days).")
    else:
        update_total_lead_time(ws_lead_times, total_lead_time_col, ws_lead_times.used_range.last_cell.row)
        print("Total Lead Time updated.")
    
    update_lead_time_date(ws_lead_times, total_lead_time_col, today_date_col, lead_time_date_col, ws_lead_times.used_range.last_cell.row)
    print("Lead Time Date updated based on Total Lead Time and Today Date.")

    fetch_holiday_data(ws_lead_times, ws_holiday_calendar, get_column_address(ws_lead_times, "Amazon Code"), lead_time_date_col, total_lead_time_col, ws_lead_times.used_range.last_cell.row)
    print("Holiday data fetched and total lead time updated accordingly.")
    
    
    ws_lead_times.range(f"{lead_time_date_col}2:{lead_time_date_col}{ws_lead_times.used_range.last_cell.row}").value = [[value] for value in ws_lead_times.range(f"{lead_time_date_col}2:{lead_time_date_col}{ws_lead_times.used_range.last_cell.row}").value]
    ws_lead_times.range(f"{today_date_col}:{today_date_col}").api.Delete()
    print("Deleted Today Date column.")
    
    print(f"Processing export workbook: {export_path}")
    wb_export = process_export_wb(export_path, ws_lead_times, app)
    format_worksheet(wb)
    format_worksheet(wb_export)
    print("Workbooks formatted successfully.")

    # save_and_close_workbooks([wb, wb_export], ["updated_26_june_" + wb.name, "Processed_POs_26_june.xlsx"])
    # print("Workbooks saved and closed successfully.")
    
    
    # Specify the full path to the output folder and filename
    # output_folder = 'C:/path/to/output_folder'
    output_filename = os.path.basename(wb_path)
    full_path = f"{output_folder}/{output_filename}"
    print(full_path)
    # time.sleep(100)
    # # Save the workbook
    wb.save(full_path)
    
    output_filename = os.path.basename(export_path)
    full_path = f"{output_folder}/{output_filename}"
    wb_export.save(full_path)
    
    print("Both Workbooks saved successfully")
    app.quit()
    print("Excel application closed.")