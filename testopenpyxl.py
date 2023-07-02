from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils import range_boundaries
from openpyxl.utils import column_index_from_string
#from openpyxl.utils import coordinate_from_string
from openpyxl.utils import coordinate_to_tuple
from openpyxl.worksheet.table import Table, TableStyleInfo
# from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import getcolumn
from array import *
import tkinter as tk
from tkinter import *
from ttkthemes import ThemedTk, THEMES
from PIL import Image
from PIL import ImageTk, Image
from tkinter.font import Font
from tkinter import filedialog
import gui_function as gui

import os
import win32com.client
from pathlib import Path  # core library

# wb_f = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test1.xlsx", data_only= False)
# wb_f.save(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test1_temp.xlsx")
# wb_d = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test1.xlsx",data_only=True)
# wb_d.save(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test1_temp2.xlsx")
# wb_d.close()
# excel_file = os.path.join(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py","Test1.xlsx")
# excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
# excel.DisplayAlerts = False # disabling prompts to overwrite existing file
# excel.Workbooks.Open(excel_file )

# excel.ActiveWorkbook.SaveAs("excel_file", FileFormat=51, ConflictResolution=2)
intpath = r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\INPUT"
outpath = r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\OUTPUT"
xl_file = "Test1.xlsx"
# excel.DisplayAlerts = True # enabling prompts
# excel.ActiveWorkbook.Close()
wb_convert = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx",data_only=True)
wb_convert.save("H137_UCIe_TC_Bump_coordination_valuecopy.xlsx")
wb_convert.close()
wb_d = load_workbook("H137_UCIe_TC_Bump_coordination_valuecopy.xlsx")
wb_f = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx")


ws_f = wb_f["Sheet1"]
ws_d = wb_d["Sheet1"]
# wstemp = wb["BALL"]
# ws2 = wb.create_sheet('BALL_temp')
# print(ws_f['D4'].value)
print(ws_d['D4'].value)
if(ws_d['D4'].value == 79):
    
    ws_f['E5'].value = f"=({str(ws_f['D4'].value).replace('=','')})+100"
    print(ws_f['E5'].value)

print(ws_d['D5'].value)
print(ws_d['E5'].value)
wb_f.save(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test1.xlsx")