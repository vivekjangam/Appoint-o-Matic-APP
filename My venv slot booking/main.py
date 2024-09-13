# import xlwings as xw
import sys
from tkinter import messagebox

import sys

from ExcelCode import mainrunner

import os

def mains():
    if len(sys.argv) > 2:
        input_master_file_path = sys.argv[1]
        consignment_file_path = sys.argv[2]
        output_folder_path = sys.argv[3]
        # Process the paths as needed
        print(f"Input Master File Path: {input_master_file_path}")
        # messagebox.showinfo(message=input_master_file_path)
        
        print(f"Consignment File Path: {consignment_file_path}")
        # messagebox.showinfo(message=consignment_file_path)
        
        print(f"Output Master File Path: {output_folder_path}")
        # messagebox.showinfo(message=output_folder_path)
        
        
        output_filename = os.path.basename(consignment_file_path)
        full_path = f"{output_folder_path}/{output_filename}"
        
        # messagebox.showinfo(message=full_path)
        
        mainrunner.main(input_master_file_path,consignment_file_path,output_folder_path)
        # codes.main(full_path)
        
        
    # else:
        
    # you can paste default path
    #     # messagebox.showinfo(message=full_path)
        
    #     mainrunner.main(input_master_file_path,consignment_file_path,output_folder_path)

    #     #codes.main(full_path)
        
if __name__ == "__main__":
    mains()
    
