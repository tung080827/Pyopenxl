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

# wb_convert = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx", data_only=True)
# wb_convert.save(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official_value.xlsx")
# wb_convert.close()
wb_d = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx", data_only=True)
ws1_d = wb_d["APD"]
wstemp_d = wb_d["BALL"]

wb = load_workbook(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx")
ws1 = wb["APD"]
wstemp = wb["BALL"]
# ws2 = wb.create_sheet('BALL_temp')
# table = ws1.tables.items()
# die1_col = input_params["Die_net_col"]
# die2_col = "N"
# row_start = 3

input_params = {
    "Die_X_col": "H", # The column contain Die X coordinate
    "Die_Y_col": "I", # The column contain Die Y coordinate
    "Die_net_col": "J", # The column contain Die net name
    "Die_maxrow": 1545, # Maximum row
    "Die_start": 4, # Begin row

    "mapping_die1": "A", # Mapping Die2Die first die column
    "mapping_die2": "B", # Mapping Die2Die second die column
    "mapping_start": 4,
    "mapping_end": 179,
 
    "Ball_net_col": "D", # Temporary Ball table
    "Ball_coord_col": "A",
    "Ball_end": 1297,
    "Ball_start": 3,
}
apd_table = {
    "Pin_number": "P",
    "Padstack": "Q",
    "X_coord": "R",
    "Y_coord": "S",
    "Pad_use": "T",
    "Net_PKG": "U",
    "Net_DIE": "V",
    "RefDes": "W",
    "Pin_BGA": "X",
    "tb_start": 4
    
}

last_vss_ball = "NULL"
last_vdd_ball = "NULL"
last_vccaon_ball = "NULL"
last_tcvddq_ball = "NULL"
last_vccio_ball = "NULL"
last_vaa_ball = "NULL"
last_vaa2_ball = "NULL"
ball_die_cmp = 0
# ball_maxrow = 1297
r = apd_table["tb_start"]
pin_number = 1

def matching(wsheet, swheet_temp, apd_table):     
    ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
    ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
    ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
    ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
    ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value
def common(val,r):
        for j in range(input_params["Ball_start"],input_params["Ball_end"] + 1):
            print(ws1[input_params["Die_net_col"] + str(i)].value)
            print(wstemp[input_params["Ball_net_col"]+ str(j)].value)              
            
            if (val == wstemp_d[input_params["Ball_net_col"]+ str(j)].value):

                ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value

                wstemp_d.delete_rows(idx=j, amount=1)
                wstemp.delete_rows(idx=j, amount=1)  
               
                r = r + 1
                input_params["Ball_end"] = input_params["Ball_end"] - 1
                break
            elif (j == input_params["Ball_end"]):
          
                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = "N/A"
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Pin_BGA"] + str(r)].value = "N/A"

                r = r + 1

for i in range(input_params["Die_start"],input_params["Die_maxrow"] + 1):
    ws1[apd_table["Pin_number"] + str(r)].value = pin_number
    
    if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("BP_") != -1 and str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("ATO") == -1 and str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DTO") == -1 and str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("ZN") == -1):
        ws1[apd_table["RefDes"] + str(r)].value = "BUMP"
        if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("BP_TX") != -1):
            ws1[apd_table["Pad_use"] + str(r)].value = "D2D"
        elif(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("BP_RX") != -1):
            ws1[apd_table["Pad_use"] + str(r)].value = "D2D"
        
        if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DIE3") != -1):
            for k in range(input_params["mapping_start"], input_params["mapping_end"]+1):
                if (ws1_d[input_params["Die_net_col"] + str(i)].value == ws1_d[input_params["mapping_die1"] + str(k)].value ):
                     ws1[apd_table["Net_PKG"] + str(r)].value = str(ws1_d[input_params["mapping_die1"] + str(k)].value).replace("BP_","").replace("[","_").replace("]","")
                     break
                # else:
                #     ws1[apd_table["Net_PKG"] + str(r)].value = "N/A"
        elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DIE7") != -1):
            for n in range(input_params["mapping_start"], input_params["mapping_end"]+1):
                if (ws1_d[input_params["Die_net_col"] + str(i)].value == ws1_d[input_params["mapping_die2"] + str(n)].value ):
                     ws1[apd_table["Net_PKG"] + str(r)].value = str(ws1[input_params["mapping_die1"] + str(n)].value).replace("BP_","").replace("[","_").replace("]","")
                     break
                    
        ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
        ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
        ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
        ws1[apd_table["Pin_BGA"] + str(r)].value = "-"
        r = r + 1
    elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("RDI_LP_CFG") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("RDI_CFG_CLK") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("RDI_MODE") != -1):
        ws1[apd_table["Pad_use"] + str(r)].value = "I"
        print(ws1_d[input_params["Die_net_col"] + str(i)].value)
        if (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DIE3") != -1 ):
            diecell_val = str(ws1_d[input_params["Die_net_col"] + str(i)].value).replace("DIE3","L")
            common(diecell_val,r)

        elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DIE7") != -1 ):
            diecell_val = str(ws1_d[input_params["Die_net_col"] + str(i)].value).replace("DIE7","R")
            common(diecell_val,r)

    elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VAA") != -1):
        ws1[apd_table["Pad_use"] + str(r)].value = "POWER"

        if (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VAA2") != -1):
            die_vaa2 = str(ws1_d[input_params["Die_net_col"] + str(i)].value).replace("VAA2","VDDA")
            for j in range(input_params["Ball_start"],input_params["Ball_end"] + 1):
                print(ws1[input_params["Die_net_col"] + str(i)].value)
                print(wstemp[input_params["Ball_net_col"]+ str(j)].value)              
                
                if (die_vaa2 == wstemp_d[input_params["Ball_net_col"]+ str(j)].value):
                    ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                    ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                    ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                    ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
                    ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value
                    # if(str(ws1[input_params["Die_net_col"] + str(i)].value).find("VDD") != -1):
                    # ball_die_cmp = 0
                    last_vaa2_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                    wstemp_d.delete_rows(idx=j, amount=1)
                    wstemp.delete_rows(idx=j, amount=1)                     
                    r = r + 1
                    input_params["Ball_end"] = input_params["Ball_end"] - 1
                    break
                elif (j == input_params["Ball_end"]):
                    ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                    ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                    ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                    ws1[apd_table["Net_PKG"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Pin_BGA"] + str(r)].value = last_vaa2_ball
        else:

            for j in range(input_params["Ball_start"],input_params["Ball_end"] + 1):

                if (ws1_d[input_params["Die_net_col"] + str(i)].value == wstemp_d[input_params["Ball_net_col"]+ str(j)].value):
                    ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                    ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                    ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                    ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
                    ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value
                    wstemp_d.delete_rows(idx=j, amount=1)
                    wstemp.delete_rows(idx=j, amount=1)  
                    last_vaa_ball = wstemp_d[input_params["Ball_coord_col"]+ str(j)].value
                    r = r + 1
                    input_params["Ball_end"] = input_params["Ball_end"] - 1
                    break
                elif (j == input_params["Ball_end"]):
                    ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                    ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                    ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                    ws1[apd_table["Net_PKG"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                    ws1[apd_table["Pin_BGA"] + str(r)].value = last_vaa_ball
                    r = r + 1
    elif ((str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VDD") != -1) or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VSS") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCIO") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCAON") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TC_VDDQ") != -1):

        if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VSS") != -1):
            ws1[apd_table["Pad_use"] + str(r)].value = "GROUND"
        else:
            ws1[apd_table["Pad_use"] + str(r)].value = "POWER"
        for j in range(input_params["Ball_start"],input_params["Ball_end"] + 1):
                
            if (ws1_d[input_params["Die_net_col"] + str(i)].value == wstemp_d[input_params["Ball_net_col"]+ str(j)].value):
                ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value
                if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VDD") != -1):
                    # ball_die_cmp = 0
                    last_vdd_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCIO") != -1):
                    last_vccio_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCAON") != -1):
                    last_vccaon_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TC_VDDQ") != -1):
                    last_tcvddq_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VSS") != -1):
                    last_vss_ball = wstemp[input_params["Ball_coord_col"]+ str(j)].value
                wstemp_d.delete_rows(idx=j, amount=1)
                wstemp.delete_rows(idx=j, amount=1)  
                r = r + 1
                input_params["Ball_end"] = input_params["Ball_end"] - 1
                break
            elif (j == input_params["Ball_end"]):
                ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value

                if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VDD") != -1):
                    # ball_die_cmp = 0
                    ws1[apd_table["Pin_BGA"] + str(r)].value = last_vdd_ball
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCIO") != -1):
                    ws1[apd_table["Pin_BGA"] + str(r)].value = last_vccio_ball
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VCCAON") != -1):
                    ws1[apd_table["Pin_BGA"] + str(r)].value = last_vccaon_ball
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TC_VDDQ") != -1):
                   ws1[apd_table["Pin_BGA"] + str(r)].value = last_tcvddq_ball
                elif (str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("VSS") != -1):
                   ws1[apd_table["Pin_BGA"] + str(r)].value = last_vss_ball 

                r = r + 1
                    
    else:
        if(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("RDI_PL_CFG") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TDO") != -1):
            temp = ws1[input_params["Die_net_col"] + str(i)].value
            ws1[apd_table["Pad_use"] + str(r)].value = "O"
        elif(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("CLK") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DBG") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("CHIP_RST") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TCK") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TRST") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TMS") != -1 or str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("TDI") != -1):
            temp = ws1[input_params["Die_net_col"] + str(i)].value
            ws1[apd_table["Pad_use"] + str(r)].value = "I"
        elif(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("ATO") != -1):
            temp = ws1[input_params["Die_net_col"] + str(i)].value
            ws1[apd_table["Pad_use"] + str(r)].value = "BI"
        elif(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("DTO") != -1):
            temp = ws1[input_params["Die_net_col"] + str(i)].value
            ws1[apd_table["Pad_use"] + str(r)].value = "O"
        elif(str(ws1_d[input_params["Die_net_col"] + str(i)].value).find("ZN") != -1):
            temp = ws1[input_params["Die_net_col"] + str(i)].value
            ws1[apd_table["Pad_use"] + str(r)].value = "O"
        for j in range(input_params["Ball_start"],input_params["Ball_end"] + 1):
             
    
            if (ws1_d[input_params["Die_net_col"] + str(i)].value == wstemp_d[input_params["Ball_net_col"]+ str(j)].value):
                ws1[apd_table["RefDes"] + str(r)].value = "BGA"
                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(j)].value
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["mapping_die1"]+ str(j)].value

                wstemp_d.delete_rows(idx=j, amount=1)
                wstemp.delete_rows(idx=j, amount=1)

                r = r + 1
                input_params["Ball_end"] = input_params["Ball_end"] - 1
                break
            elif (j == input_params["Ball_end"]):

                ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
                ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
                ws1[apd_table["Net_PKG"] + str(r)].value = "NA"
                ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
                ws1[apd_table["Pin_BGA"] + str(r)].value = "NA"

                r = r + 1
    pin_number += 1

for m in range (input_params["Ball_start"], input_params["Ball_end"] + 1):
    ws1[apd_table["RefDes"] + str(r)].value = "BGA"
    if(str(wstemp_d[input_params["Ball_net_col"]+ str(m)].value).find("VSS") != -1):
        ws1[apd_table["Pad_use"] + str(r)] = "GROUND"
    else:
        ws1[apd_table["Pad_use"] + str(r)] = "NC"
    # elif(str(wstemp[input_params["Ball_net_col"]+ str(m)].value).find("VDD") != -1 or str(wstemp[input_params["Ball_net_col"]+ str(m)].value).find("TC_VDDQ") != -1 or str(wstemp[input_params["Ball_net_col"]+ str(m)].value).find("VCCAON") != -1 or str(wstemp[input_params["Ball_net_col"]+ str(m)].value).find("VCCIO") != -1 or str(wstemp[input_params["Ball_net_col"]+ str(m)].value).find("VAA") != -1):
    #     ws1[apd_table["Pad_use"] + str(r)] = "POWER"
    # elif(str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("BP_") != -1):
    #     ws1[apd_table["Pad_use"] + str(r)].value = "D2D"
    # elif (str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("RDI_LP_CFG") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("RDI_CFG_CLK") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("RDI_MODE") != -1):
    #     ws1[apd_table["Pad_use"] + str(r)].value = "I"
    # elif(str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("RDI_PL_CFG") != 1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("TDO") != -1):
    #     ws1[apd_table["Pad_use"] + str(r)].value = "O"
    # elif(str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("CLK") != 1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("DBG") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("CHIP_RST") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("TCK") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("TRST") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("TMS") != -1 or str(wstemp[input_params["Ball_net_col"] + str(m)].value).find("TDI") != -1):
    #     ws1[apd_table["Pad_use"] + str(r)].value = "I"
    
    # ws1[apd_table["X_coord"] + str(r)].value = ws1[input_params["Die_X_col"] + str(i)].value
    # ws1[apd_table["Y_coord"] + str(r)].value = ws1[input_params["Die_Y_col"] + str(i)].value   
    ws1[apd_table["Net_PKG"] + str(r)].value = wstemp[input_params["Ball_net_col"]+ str(m)].value
    # ws1[apd_table["Net_DIE"] + str(r)].value = ws1[input_params["Die_net_col"] + str(i)].value
    ws1[apd_table["Pin_BGA"] + str(r)].value = wstemp[input_params["Ball_coord_col"]+ str(m)].value    
    r = r + 1
print(input_params["Ball_end"]) 
wb.save(r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\H137_UCIe_TC_Bump_coordination_official.xlsx")
# print(len(table))3446	4692.95	DIE7_BP_ATO

# print(table[0])
# print(table[0][0])

