import datetime
import xlwings as xw

# is it needed
def add_date_values(ws, col_address, last_row, date_value):
    """Adds a date value to a specified range in the worksheet.

    Args:
        ws (xw.Sheet): The worksheet object.
        col_address (str): The column address for the date values.
        last_row (int): The last row of the data range.
        date_value (datetime.date): The date value to add.
    """
    ws.range(f"{col_address}2:{col_address}{last_row}").value = date_value

def update_total_lead_time(ws, col_address, last_row, adjustment=0):
    """Updates the total lead time in a column with an optional adjustment.

    Args:
        ws (xw.Sheet): The worksheet object.
        col_address (str): The column address for the total lead time.
        last_row (int): The last row of the data range.
        adjustment (int, optional): An optional adjustment value to add (default: 0).
    """
    for row in range(2, last_row + 1):
        cell_address = f"{col_address}{row}"
        ws.range(cell_address).value += adjustment
        ws.range(cell_address).value = round(ws.range(cell_address).value, 0)

def update_lead_time_date(ws, total_lead_time_col, today_date_col, lead_time_date_col, last_row):
    ws.range(f"{lead_time_date_col}2:{lead_time_date_col}{last_row}").formula = f"={today_date_col}2+(ROUNDUP({total_lead_time_col}2,0))"

def fetch_holiday_data(ws, holiday_ws, amazon_code_col, lead_time_date_col, total_lead_time_col, last_row):
    """Updates total lead time in a worksheet based on matching holidays in another worksheet.

    Iterates through each row in the main worksheet (`ws`) and checks if the lead time date matches any dates in the holiday worksheet (`holiday_ws`).
    If a match is found for the Amazon code (`amazon_code_col`), the corresponding holiday value from the holiday worksheet is added to the total lead time (`total_lead_time_col`) in the main worksheet.

    Args:
        ws (xw.Sheet): The worksheet containing the main data.
        holiday_ws (xw.Sheet): The worksheet containing holiday data.
        amazon_code_col (str): The column address for the Amazon code in the main worksheet.
        lead_time_date_col (str): The column address for the lead time date in the main worksheet.
        total_lead_time_col (str): The column address for the total lead time in the main worksheet.
        last_row (int): The last row of the data range in the main worksheet.

    Returns:
        None: The function modifies the worksheet in-place, no value is returned.
    """
    for row in range(2, last_row + 1):
        amazon_code = ws.range(f"{amazon_code_col}{row}").value
        lead_time_date = ws.range(f"{lead_time_date_col}{row}").value
        total_lead_time = ws.range(f"{total_lead_time_col}{row}").value
        
        for holiday_row in range(2, holiday_ws.used_range.last_cell.row + 1):
            sheet2_date = holiday_ws.range(f"C{holiday_row}").value
            
            if lead_time_date == sheet2_date:
                header_row_values = holiday_ws.range("1:1").value
                try:
                    index_of_amazon_code = header_row_values.index(amazon_code)
                    related_col_address = xw.utils.col_name(index_of_amazon_code + 1)
                    holiday_value = holiday_ws.range(f"{related_col_address}{holiday_row}").value or 0
                    ws.range(f"{total_lead_time_col}{row}").value += holiday_value
                    lead_time_date = ws.range(f"{lead_time_date_col}{row}").value
                    
                    # Print details about the matched holiday (one-liner after fetching)
                    print(f"Row {row}: Matched holiday for Amazon code '{amazon_code}' on {lead_time_date}, adding holiday value {holiday_value}.")

                except ValueError:
                    pass
