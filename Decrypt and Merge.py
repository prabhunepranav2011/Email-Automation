from xlrd import *
import win32com.client
import csv
import sys
import pandas as pd
import glob

def merger(file_list, password):
    files = []
    
    for file_name in file_list:
        xlwb = xlApp.Workbooks.Open(file_name, False, True , None, Password=password)
        xl_sh = xlwb.Worksheets(1)

        row_num = 0
        cell_val = ''
        while cell_val != None:
            row_num += 1
            cell_val = xl_sh.Cells(row_num, 1).Value
        # print(row_num, '|', cell_val, type(cell_val))
        last_row = row_num - 1
        print(last_row)

    # Get last_column
        col_num = 0
        cell_val = ''
        while cell_val != None:
            col_num += 1
            cell_val = xl_sh.Cells(1, col_num).Value
        # print(col_num, '|', cell_val, type(cell_val))
        last_col = col_num
        print(last_col)

        content = xl_sh.Range(xl_sh.Cells(2, 1), xl_sh.Cells(last_row, last_col)).Value
    # list(content)
        df = pd.DataFrame(list(content[1:]), columns=content[0])
        print(df)
        files.append(df)
    excl_merged = pd.concat(files, ignore_index=True)
    
    return excl_merged




file_list1 = glob.glob(r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles/*.xlsx") #Enter excel file location

#file_list2 = glob.glob(r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles/*.xlsx") #Enter excel file location
#file_list3 = glob.glob(r"C:\Users\PRABHUNEPRANAVKAILAS\Desktop\Udemy Python TCS\Excel_FIles/*.xlsx") #Enter excel file location

file_lists = [file_list1] #, file_list2, file_list3]
#files = []
xlApp = win32com.client.Dispatch("Excel.Application")
#print("Excel library version:", xlApp.Version)
password = 'password'      #Enter Password here

returned_excl_merged1 = merger(file_list1, password)
returned_excl_merged1.to_excel('Labeler1.xlsx', index=False)
"""
returned_excl_merged2 = merger(file_list2, password)
returned_excl_merged2.to_excel('Labeler2.xlsx', index=False)

returned_excl_merged3 = merger(file_list3, password)
returned_excl_merged3.to_excel('Labeler3.xlsx', index=False)
"""

