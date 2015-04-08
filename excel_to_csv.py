#!/usr/python

"""
This script takes crime data from the city of Houston in Excel format and puts
it into a more data science friendly CSV format.
"""

import sys
import re
from mmap import mmap, ACCESS_READ
from xlrd import open_workbook, xldate_as_tuple, XL_CELL_DATE
from datetime import date

# Parse the command line arguments to find Excel files
for argument in sys.argv:
    if argument[-4:] == ".xls":
        excel_filepath = argument
        """ For each Excel file passed as an argument, convert it to a CSV file """
        book = open_workbook(excel_filepath)
        excel_filename = re.split(r"/",excel_filepath)[-1]
        filepath = excel_filepath.replace(excel_filename,"")
        file_string = ""
        
        # Convert this Excel workbook to a string of comma separated values
        for sheet in book.sheets():
            """ Add the contents of each sheet in the workbook to the CSV file """
            file_string += "Workbook name: " + excel_filename + "\n"
            file_string += "Sheet name: " + sheet.name + "\n"
            file_string += "Number of rows: " + str(sheet.nrows) + "\n"
            file_string += "Number of columns: " + str(sheet.ncols) + "\n"
            for row in range(sheet.nrows):
                row_string = ""
                for col in range(sheet.ncols):
                    if col != 0:
                        """ no comma before the first entry in a row """
                        row_string += ", "
                    if sheet.cell(row, col).ctype == XL_CELL_DATE: 
                        """ if the cell contains an Excel date it must be
                        converted to a human readable string """
                        date_value = xldate_as_tuple(sheet.cell(row, col).value,
                                                      book.datemode)
                        row_string += str(date(*date_value[:3]))
                    else:
                        """ other cells are converted directly to strings """
                        row_string += str(sheet.cell(row, col).value)
                file_string += row_string + "\n"
            
        # Write the string of comma separated values to a file
        csv_filename = excel_filename.replace(".xls",".csv")
        csv_filehandle = open(csv_filename, "w")
        csv_filehandle.write(file_string)
        csv_filehandle.close()
        
