import datetime
import xlwings as xw

def open_workbook(wb_path, visible=True):
    """Opens an Excel workbook and returns the app and workbook objects.

    Args:
        wb_path (str): The path to the Excel workbook.
        visible (bool, optional): Whether to make the workbook visible (default: True).

    Returns:
        tuple: A tuple containing the xlwings app object and the opened workbook object.
    """
    app = xw.App(visible=visible)
    workbook = app.books.open(wb_path)
    return app, workbook



def get_column_address(sheet: xw.Sheet, header_name: str) -> str:
  """Gets the column address (e.g., "A", "B10") of a header name in the first row of an xlwings sheet.

  Args:
      sheet (xw.Sheet): The xlwings sheet object to work on.
      header_name (str): The header name you want to find the address for.

  Returns:
      str: The column address of the header name if found, otherwise an empty string.

  Raises:
      ValueError: If the sheet object is not provided or empty.
  """

  if not sheet:
    raise ValueError("Sheet object is not provided.")

  try:
    # Get values from the first row
    first_row_values = sheet.range("1:1").value

    # Use list comprehension for efficient iteration and address retrieval
    column_index = next((i for i, cell in enumerate(first_row_values) if cell == header_name), None)

    # Construct the address (assuming headers are in the first row)
    if column_index is not None:
      return chr(column_index + 65)  # Convert index to A-Z letter (0-based)
    else:
      return ""  # Return empty string if not found

  except Exception as e:
    raise ValueError(f"Error finding header '{header_name}': {e}")

def add_column(sheet, column_name):
    """Adds a new column with the specified name to the worksheet.

    Args:
        sheet (xw.Sheet): The worksheet object.
        column_name (str): The name of the new column.

    Returns:
        str: The column address of the newly added column.
    """
    last_col = sheet.used_range.last_cell.column
    new_col = last_col + 1
    sheet.cells(1, new_col).value = column_name
    return get_column_address(sheet, column_name)

def add_formulas(ws, formulas, last_row):
    """Adds formulas to a specified range in the worksheet.

    Args:
        ws (xw.Sheet): The worksheet object.
        formulas (list): A list of tuples containing the column address and formula string.
        last_row (int): The last row of the data range.
    """
    for formula in formulas:
        col_address, formula_str = formula
        ws.range(f"{col_address}2:{col_address}{last_row}").formula = formula_str

def convert_to_values(sheet, col_adress,last_row):
    """Converts a single column range in a worksheet to a list of single-element lists.

    Args:
        sheet (xw.Sheet): The worksheet object containing the data.
        col_address (str): The column address of the range to convert (e.g., "A", "B").
        last_row (int): The last row of the data range to convert.

    Returns:
        None: The function modifies the worksheet in-place, no value is returned.
    """
    values = sheet.range(f"{col_adress}2:{col_adress}{last_row}").value
    values = [[value] for value in values]
    sheet.range(f"{col_adress}2:{col_adress}{last_row}").value = values

def format_worksheet(wb):
    """Formats all worksheets in a workbook by bolding the first row and autofitting columns.

    Args:
        wb (xw.Workbook): The workbook object to format.

    Returns:
        None: The function modifies the workbook in-place, no value is returned.
    """
    for ws in wb.sheets:
        ws.range("1:1").api.Font.Bold = True  # Bold the first row
        ws.autofit()  # Autofit columns

def save_and_close_workbooks(workbooks, paths):
    """Saves and closes a list of workbooks to specified paths.

    Args:
        workbooks (list[xw.Workbook]): A list of workbook objects to save and close.
        paths (list[str]): A list of corresponding paths where to save the workbooks.

    Returns:
        None: The function performs the saving and closing actions, no value is returned.
    """
    for wb, path in zip(workbooks, paths):
        wb.save(path)
        print(f"Workbook saved: {path}")  # Print the path of the saved workbook
        wb.close()


