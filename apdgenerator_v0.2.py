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
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection
import os
import win32com.client
from pathlib import Path  # core library


excel_file = r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Test_adp.xlsx"
# die_list={
#     diel_list = ['DIE3', 'DIE4']
#     diel_begin_list = ['T19', 'AB19']
#     diel_end_list = ['V791', 'AD791']
#     dier_list = ['DIE7', 'DIE8']
#     dier_begin_list = ['X19', 'AF19']
#     dier_end_list = ['Z791', 'AH791']
# }



def get_col_row_range(cell_begin, cell_end):
    row_begin = coordinate_to_tuple(cell_begin)[0]
    col_begin = coordinate_to_tuple(cell_begin)[1]
    row_end = coordinate_to_tuple(cell_end)[0]
    col_end = coordinate_to_tuple(cell_end)[1]
    return row_begin,row_end,col_begin,col_end
def get_config():

    # die_params = {
    #     "Die_sheet":"Package_substrate",
    #     "Die_L_begin_cell":"T19",
    #     "Die_L_end_cell": "V791",
    #     "Die_L_name": "DIE3",
    #     "Die_R_begin_cell":"X19",
    #     "Die_R_end_cell": "Z791",
    #     "Die_R_name": "DIE7",
    # }
    die_params= {
        "die_sheet" : "Package_substrate",
        "diel_list" : ['DIE3', 'DIE4'],
        "diel_begin_list" : ['T19', 'AB19'],
        "diel_end_list" : ['V791', 'AD791'],
        "dier_list" : ['DIE7', 'DIE8'],
        "dier_begin_list" : ['X19', 'AF19'],
        "dier_end_list" : ['Z791', 'AH791']
    }
    
     
    input_params = {
        "excel_file": excel_file,

        "mapping_sheet": "UCIe_Mapping_connection",
        "mapping_begin_cell": "F1",
        "mapping_end_cell":"G178",
        
        "ball_tb_sheet": "BGA",
        "ball_tb_begin_cell": "AQ2",
        "ball_tb_end_cell": "AT1297",
    

        "apd_sheet": "APD"
    }

    # adp table out put config 
    out_put = {
        "sheet": "APD",
        "tb_loc": "P4"
       
    }
    return die_params,input_params,out_put



def refresh_excel(excelfile):
    excel_file = os.path.join(excelfile)
    excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = True # disabling prompts to overwrite existing file
    excel.Workbooks.Open(excel_file )
    excel.ActiveWorkbook.Save()
    excel.DisplayAlerts = True # enabling prompts
    excel.ActiveWorkbook.Close()

def copy_table(cell):
    try:
        wb_tempsheet = load_workbook(cell['excel_file'])
    except:
        print("Ploc file does not exist or opening. Please close/recheck it")
        return

    # row_begin = coordinate_to_tuple(cell['ball_tb_begin_cell'])[0]
    # col_begin = coordinate_to_tuple(cell['ball_tb_begin_cell'])[1]
    # row_end = coordinate_to_tuple(cell['ball_tb_end_cell'])[0]
    # col_end = coordinate_to_tuple(cell['ball_tb_end_cell'])[1]
    row_begin = get_col_row_range(cell['ball_tb_begin_cell'],cell['ball_tb_end_cell'])[0]
    row_end =  get_col_row_range(cell['ball_tb_begin_cell'],cell['ball_tb_end_cell'])[1]
    col_begin = get_col_row_range(cell['ball_tb_begin_cell'],cell['ball_tb_end_cell'])[2]
    col_end =  get_col_row_range(cell['ball_tb_begin_cell'],cell['ball_tb_end_cell'])[3]
    sheet_ls = wb_tempsheet.sheetnames
    
    wstmp_create_name = str(cell['ball_tb_sheet']) + "TEMP_TABLE"

    if wstmp_create_name in sheet_ls:
        msg_ws = messagebox.askquestion('Create Sheet', 'The ' + wstmp_create_name + ' already exist. Do you want to overwrite it?',icon='question')
        # sh_index = sheet_ls.index(wstmp_create_name) 
        tmp_s = wb_tempsheet.get_sheet_by_name(wstmp_create_name)
        print("The " + wstmp_create_name + " exist.")
        if(msg_ws == 'yes'):
            wb_tempsheet.remove_sheet(tmp_s)
                       
        else:
            print("Stop the process")
            return
    wstmp_create = wb_tempsheet.create_sheet(wstmp_create_name) 
    cell_list = wb_tempsheet.sheetnames
    if cell['ball_tb_sheet'] in cell_list:
        source_sheet = wb_tempsheet[cell['ball_tb_sheet']]
        print(row_begin)
        print(col_begin)
        print(row_end)
        print(col_end)
        for i in range (row_begin, row_end + 1):
            for j in range (col_begin, col_end + 1):
                # if(str(source_sheet[get_column_letter(j) + str(i)].value).find("=") != -1):
                    # wstmp_create[get_column_letter(j) + str(i)].value = f"={cell['ball_tb_sheet']}!{str(source_sheet[get_column_letter(j) + str(i)].value).replace('=','')}"
                # else:
                wstmp_create[get_column_letter(j) + str(i)].value = f"={cell['ball_tb_sheet']}!{str(get_column_letter(j) + str(i))}"
        wb_tempsheet.save(excel_file)
        wb_tempsheet.close()

        return row_begin,col_begin,row_end,col_end,wstmp_create_name 
    else:
        print("The sheet does not exist. Please recheck")
        return

def gen(adp, die, mapping, mapping_prefix, ball, last_ball):
        for title in range(adp['begincol'], adp['begincol'] + 10):
            adp['sheet'][get_column_letter(title) + str(adp['beginrow'])].border = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
            adp['sheet'][get_column_letter(title) + str(adp['beginrow'])].fill = PatternFill(patternType='solid', fgColor='9e42f5')
        adp['sheet'][adp['Pin_number'] + str(adp['beginrow'])].value = "Pin_number"
        adp['sheet'][adp['Pad_stack'] + str(adp['beginrow'])].value = "Pad_stack"
        adp['sheet'][adp['X'] + str(adp['beginrow'])].value = "X_coord"
        adp['sheet'][adp['Y'] + str(adp['beginrow'])].value = "Y_coord"
        adp['sheet'][adp['rotation'] + str(adp['beginrow'])].value = "Rotation"
        adp['sheet'][adp['Pad_use'] + str(adp['beginrow'])].value = "Pad_use"
        adp['sheet'][adp['pkg_name'] + str(adp['beginrow'])].value = "Net_name(Package)"
        adp['sheet'][adp['die_name'] + str(adp['beginrow'])].value = "Net_name(Die)"
        adp['sheet'][adp['resdef'] + str(adp['beginrow'])].value = "ResDef"
        adp['sheet'][adp['bga_pin'] + str(adp['beginrow'])].value = "Package_pin"
        for i in range(die['row_min'] + 2, die['row_max'] + 1):
            adp['sheet'][adp['Pin_number'] + str(adp['r'])].value = adp['pin_num'] 
            #------------------------------Die2Die Connections--------------------------------------------
            if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("BP_") != -1 and str(die['sheet_d'][die['net_name'] + str(i)].value).find("ATO") == -1 and str(die['sheet_d'][die['net_name'] + str(i)].value).find("DTO") == -1 and str(die['sheet_d'][die['net_name'] + str(i)].value).find("ZN") == -1):
                adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BUMP"
                if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("BP_TX") != -1):
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "D2D"
                elif(str(die['sheet_d'][die['net_name'] + str(i)].value).find("BP_RX") != -1):
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "D2D"
                
                # if(str(die['sheet_d'][die['net_name'] + str(i)].value).find(die['name']) != -1):
                if(str(die['die_side']) == "L"):
                    for k in range(mapping['row_min'], mapping['row_max']+1):
                        if (str(die['sheet_d'][die['net_name'] + str(i)].value).find(str(mapping['sheet_d'][mapping['die_L'] + str(k)].value)) != -1):
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = mapping_prefix + str(mapping['sheet_d'][mapping['die_L'] + str(k)].value).replace("BP_","").replace("[","_").replace("]","")
                            break
                        # else:
                          
                        #     adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "NA"
                        #     print("The net name is not found. Please recheck mapping table")
                    # elif (str(ws_d[die['net_name'] + str(i)].value).find("DIE7") != -1):
                elif(str(die['die_side']) == "R"):
                    for n in range(mapping['row_min'], mapping['row_max']+1):
                        if (str(die['sheet_d'][die['net_name'] + str(i)].value).find(str(mapping['sheet_d'][mapping['die_R'] + str(n)].value)) != -1):
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = mapping_prefix + "_" + str(mapping['sheet_d'][mapping['die_L'] + str(n)].value).replace("BP_","").replace("[","_").replace("]","")
                            break
                        # else:                          
                        #     adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "NA"
                        #     print("The net name is not found. Please recheck mapping table")
                            
                adp['sheet'][adp['X'] + str(adp['r'])].value = f"={die['sheet_name']}!{die['X'] + str(i)}"
                adp['sheet'][adp['Y'] + str(adp['r'])].value = f"={die['sheet_name']}!{die['Y'] + str(i)}"   
                adp['sheet'][adp['die_name'] + str(adp['r'])].value = f"={die['sheet_name']}!{die['net_name'] + str(i)}"
                adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = "-"
                adp['r'] += 1
            #------------------------------Common RDI input Connections--------------------------------------------
            elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_LP_CFG") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_CFG_CLK") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_MODE") != -1):
                adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "I"
                print(adp['sheet'][die['net_name'] + str(i)].value)
                # if (str(die['sheet_d'][die['net_name'] + str(i)].value).find(die_L_name) != -1 ):
                if(die['die_side'] == "L"):
                    diecell_val = str(die['sheet_d'][die['net_name'] + str(i)].value).replace(die['name'],"L")
                    for j in range(ball['begin_row'], ball['end_row'] + 1):
            
                        if (diecell_val == ball['sheet_d'][ball['net_col']+ str(j)].value):

                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_f'][ball['net_col']+ str(j)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"

                            # ball['sheet_d'].delete_rows(idx=j, amount=1)
                            # ball['sheet_f'].delete_rows(idx=j, amount=1)  
                        
                            adp['r'] += 1
                            # ball['end_row'] = ball['end_row'] - 1
                            break
                        elif (j == ball['end_row']):
                    
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "N/A"
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = "N/A"
                            adp['r'] += 1

                elif (die['die_side'] == "R"):
                    diecell_val = str(die['sheet_d'][die['net_name'] + str(i)].value).replace(die['name'],"R")
                    for j in range(ball['begin_row'], ball['end_row'] + 1):
              
                        if (diecell_val == ball['sheet_d'][ball['net_col']+ str(j)].value):

                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_f'][ball['net_col']+ str(j)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"

                            # ball['sheet_d'].delete_rows(idx=j, amount=1)
                            # ball['sheet_f'].delete_rows(idx=j, amount=1)  
                        
                            adp['r'] += 1
                            # ball['end_row'] = ball['end_row'] - 1
                            break
                        elif (j == ball['end_row']):
                    
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "N/A"
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = "N/A"
                            adp['r'] += 1
                 

            elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VAA") != -1):
                adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "POWER"

                if (str(die['sheet_d'][die['net_name']  + str(i)].value).find("VAA2") != -1):
                    die_vaa2 = str(die['sheet_d'][die['net_name'] + str(i)].value).replace("VAA2","VDDA")
                    for j in range(ball['begin_row'],ball['end_row'] + 1):
                        # print(die['sheet_d'][die['net_name'] + str(i)].value)
                        # print(ball['sheet_d'][ball['net_col']+ str(j)].value)
                        if (die_vaa2 == ball['sheet_d'][ball['net_col']+ str(j)].value):
                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_d'][ball['net_col']+ str(j)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"
                            # if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("VDD") != -1):
                            # ball_die_cmp = 0
                            last_ball['vaa2'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                            ball['sheet_d'].delete_rows(idx=j, amount=1)
                            ball['sheet_f'].delete_rows(idx=j, amount=1)                     
                            adp['r'] += 1
                            ball['end_row'] = ball['end_row'] - 1
                            break
                        elif (j == ball['end_row']):
                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vaa2']
                else:

                    for j in range(ball['begin_row'],ball['end_row']  + 1):

                        if (die['sheet_d'][die['net_name'] + str(i)].value == ball['sheet_d'][ball['net_col']+ str(j)].value):
                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X']+ str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_d'][ball['net_col']+ str(j)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"


                            ball['sheet_d'].delete_rows(idx=j, amount=1)
                            ball['sheet_f'].delete_rows(idx=j, amount=1)  
                            last_ball['vaa'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                            adp['r'] += 1
                            ball['end_row'] = ball['end_row'] - 1
                            break
                        elif (j == ball['end_row'] ):
                            adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                            adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X']+ str(i)].value
                            adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value   
                            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vaa']
                            adp['r'] += 1
            elif ((str(die['sheet_d'][die['net_name'] + str(i)].value).find("VDD") != -1) or str(die['sheet_d'][die['net_name'] + str(i)].value).find("VSS") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCIO") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCAON") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TC_VDDQ") != -1):

                if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("VSS") != -1):
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "GROUND"
                else:
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "POWER"
                for j in range(ball['begin_row'],ball['end_row'] + 1):
                        
                    if (die['sheet_d'][die['net_name'] + str(i)].value == ball['sheet_d'][ball['net_col']+ str(j)].value):
                        adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                        adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                        adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value    
                        adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_d'][ball['net_col']+ str(j)].value
                        adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                        adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"
                        if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("VDD") != -1):
                            # ball_die_cmp = 0
                            last_ball['vdd'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCIO") != -1):
                            last_ball['vccio'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCAON") != -1):
                            last_ball['vccaon'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("TC_VDDQ") != -1):
                            last_ball['tc_vddq'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VSS") != -1):
                            last_ball['vss'] = adp['sheet'][adp['bga_pin'] + str(adp['r'])].value
                        ball['sheet_d'].delete_rows(idx=j, amount=1)
                        ball['sheet_f'].delete_rows(idx=j, amount=1)  
                        adp['r'] += 1
                        ball['end_row'] = ball['end_row'] - 1
                        break
                    elif (j == ball['end_row']):
                        adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                        adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                        adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value    
                        adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                        adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value

                        if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("VDD") != -1):
                            # ball_die_cmp = 0
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vdd']
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCIO") != -1):
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vccio']
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VCCAON") != -1):
                            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vccaon']
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("TC_VDDQ") != -1):
                           adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['tc_vddq']
                        elif (str(die['sheet_d'][die['net_name'] + str(i)].value).find("VSS") != -1):
                           adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = last_ball['vss'] 

                        adp['r'] += 1
                            
            else:
                if(str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_PL_CFG") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TDO") != -1):
                    temp = die['sheet_d'][die['net_name'] + str(i)].value
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "O"
                elif(str(die['sheet_d'][die['net_name'] + str(i)].value).find("CLK") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("DBG") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("CHIP_RST") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TCK") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TRST") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TMS") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("TDI") != -1):
                    temp = die['sheet_d'][die['net_name'] + str(i)].value
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "I"
                elif(str(die['sheet_d'][die['net_name'] + str(i)].value).find("ATO") != -1):
                    temp = die['sheet_d'][die['net_name'] + str(i)].value
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "BI"
                elif(str(die['sheet_d'][die['net_name'] + str(i)].value).find("DTO") != -1):
                    temp = die['sheet_d'][die['net_name'] + str(i)].value
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "O"
                elif(str(die['sheet_d'][die['net_name'] + str(i)].value).find("ZN") != -1):
                    temp = die['sheet_d'][die['net_name'] + str(i)].value
                    adp['sheet'][adp['Pad_use'] + str(adp['r'])].value = "O"
                for j in range(ball['begin_row'],ball['end_row'] + 1):
                    
            
                    if (die['sheet_d'][die['net_name'] + str(i)].value == ball['sheet_d'][ball['net_col']+ str(j)].value):
                        adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
                        adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                        adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value    
                        adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_d'][ball['net_col']+ str(j)].value
                        adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                        adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(j)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(j)].value).replace('=','')}"

                        ball['sheet_d'].delete_rows(idx=j, amount=1)
                        ball['sheet_f'].delete_rows(idx=j, amount=1)

                        adp['r'] += 1
                        ball['end_row'] = ball['end_row'] - 1
                        break
                    elif (j == ball['end_row']):

                        adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
                        adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value    
                        adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "NA"
                        adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
                        adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = "NA"

                        adp['r'] += 1
            adp['pin_num'] = int(adp['pin_num']) + 1

        return adp,mapping,ball,last_ball

def gen_apd():

    params = get_config()
    die_params = params[0]
    input_params = params[1]
    output_params = params[2]
    print("Creating the ball temple sheet...")
    ball_tmp_tb = copy_table(input_params)
    print("Refreshing the excelfile...")
    refresh_excel(input_params['excel_file'])



    last_ball ={
        "vss": "NULL",
        "vdd": "NULL",
        "vccio":"NULL",
        "vaa":"NULL",
        "vccaon":"NULL",
        "vaa2":"NULL",
        "tc_vddq":"NULL",
        "die_cmp": 0

    }

    # # ball_maxrow = 1297
    # r = apd_table["tb_start"]
    # pin_number = 1
    print("Loading workbook")
    wb_d = load_workbook(excel_file, data_only=True)
    wb_f = load_workbook(excel_file)
    apd_sheet = input_params["apd_sheet"]
    sheet_list = wb_f.sheetnames
    if apd_sheet in sheet_list:
        
        ws_adp_f = wb_f[apd_sheet]
        
    else:
        msg_ws = messagebox.askquestion('Create Sheet', 'The sheet' + apd_sheet + ' doesn\'t exist. Do you want to create it?', icon='question')
            
        if(msg_ws == 'yes'):
            ws_adp_f = wb_f.create_sheet(apd_sheet)
            # mynotif("")
            # mynotif('Creating the sheet...')
        else:
            # mynotif("")
            # progress_bar(0)
            # return
            pass

    # ws_d = wb_d[apd_sheet]
    # die_sheet_name = die_params['Die_sheet']
    # ws_die_d = wb_d[die_sheet_name]

    # ws_die_f = wb_f[die_sheet_name]
    wstemp_d = wb_d[ball_tmp_tb[4]]
    wstemp_f = wb_f[ball_tmp_tb[4]]

    # def matching(wsheet, swheet_temp, apd_table):     
    #     adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_d'][die['X'] + str(i)].value
    #     adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_d'][die['Y'] + str(i)].value    
    #     adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_d'][ball['net_col']+ str(j)].value
    #     adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_d'][die['net_name'] + str(i)].value
    #     adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = wstemp[input_params["mapping_die1"]+ str(j)].value
    pin_number = 1
    adp_tb_loc = output_params['tb_loc']
    adp_begincol = coordinate_to_tuple(adp_tb_loc)[1]
    adp_beginrow = coordinate_to_tuple(adp_tb_loc)[0]
    adp_Pin_number = get_column_letter(adp_begincol)
    adp_Pad_stack = get_column_letter(adp_begincol + 1)
    adp_Xcoor = get_column_letter(adp_begincol + 2)
    adp_Ycoor = get_column_letter(adp_begincol + 3)
    adp_rotation = get_column_letter(adp_begincol + 4)
    adp_Pad_use = get_column_letter(adp_begincol + 5)
    adp_pkg_name = get_column_letter(adp_begincol + 6)
    adp_die_name = get_column_letter(adp_begincol + 7)
    adp_resdef = get_column_letter(adp_begincol + 8)
    adp_bga_pin = get_column_letter(adp_begincol + 9)
    radp = adp_beginrow + 1

    adp ={
        "sheet":ws_adp_f,
        "pin_num":int(pin_number),
        "tb_loc": adp_tb_loc,
        "begincol": adp_begincol, 
        "beginrow" : adp_beginrow,
        "Pin_number": adp_Pin_number, 
        "Pad_stack": adp_Pad_stack, 
        "X": adp_Xcoor, 
        "Y": adp_Ycoor, 
        "rotation": adp_rotation, 
        "Pad_use": adp_Pad_use, 
        "pkg_name": adp_pkg_name, 
        "die_name": adp_die_name, 
        "resdef":  adp_resdef,
        "bga_pin":  adp_bga_pin,
        "r":radp
    }

   

    mapping_row_col = get_col_row_range(input_params['mapping_begin_cell'], input_params["mapping_end_cell"])
    mapping_row_min = mapping_row_col[0]
    mapping_row_max = mapping_row_col[1]
    mapping_col_min = mapping_row_col[2]
    mapping_col_max = mapping_row_col[3]
    mapping_die_L = get_column_letter(mapping_col_min)
    mapping_die_R = get_column_letter(mapping_col_min + 1)
    mapping_sheet_d = wb_d[input_params['mapping_sheet']]

    mapping={
        "sheet_d":mapping_sheet_d,
        "row_min":mapping_row_min,
        "row_max":mapping_row_max,
        "col_min":mapping_col_min,
        "col_max":mapping_col_max,
        "die_L":mapping_die_L,
        "die_R":mapping_die_R
        
    }

    ball_begin_row = ball_tmp_tb[0]
    ball_end_row = ball_tmp_tb[2]
    ball_begin_col = ball_tmp_tb[1]
    ball_end_col = ball_tmp_tb[3]
    ball_net_col = get_column_letter(ball_end_col)
    ball_X = get_column_letter(ball_end_col - 2)
    ball_Y = get_column_letter(ball_end_col - 1)

    ball={
        "sheet_d":wstemp_d,
        "sheet_f": wstemp_f,
        "begin_row":ball_begin_row,
        "end_row":ball_end_row,
        "begin_col":ball_begin_col,
        "end_col":ball_end_col,
        "net_col":ball_net_col,
        "X":ball_X,
        "Y":ball_Y
    }



    # "Die_sheet":"Package_substrate",
    #         "Die_L_begin_cell":"G20",
    #         "Die_L_end_cell": "I792",
    #         "Die_R_begin_cell":"K20",
    #         "Die_R_end_cell": "M792",
    
    # def common(val,die,ball_r_begin, ball_r_end,adp['r'], r_die):
    #         for j in range(ball_r_begin, ball_r_end + 1):
    #             # print(adp['sheet'][die_L_net_name + str(i)].value)
    #             # print(adp['sheet'][ball['net_col']+ str(j)].value)              
                
    #             if (val == ball['sheet_d'][ball['net_col']+ str(j)].value):

    #                 adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
    #                 adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(r_die)].value
    #                 adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(r_die)].value   
    #                 adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_f'][ball['net_col']+ str(j)].value
    #                 adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(r_die)].value
    #                 adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = ball['sheet_f'][ball['Y'] + str(j)].value + ball['sheet_f'][ball['X'] + str(j)].value

    #                 ball['sheet_d'].delete_rows(idx=j, amount=1)
    #                 ball['sheet_f'].delete_rows(idx=j, amount=1)  
                
    #                 adp['r'] += 1
    #                 ball_r_end = ball_r_end - 1
    #                 break
    #             elif (j == ball_r_end):
            
    #                 adp['sheet'][adp['X'] + str(adp['r'])].value = die['sheet_f'][die['X'] + str(r_die)].value
    #                 adp['sheet'][adp['Y'] + str(adp['r'])].value = die['sheet_f'][die['Y'] + str(r_die)].value   
    #                 adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = "N/A"
    #                 adp['sheet'][adp['die_name'] + str(adp['r'])].value = die['sheet_f'][die['net_name'] + str(r_die)].value
    #                 adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = "N/A"
    #                 adp['r'] += 1
    #         return adp['r'],ball_r_end   
    
    
    
    for die_cnt in range(0, len(die_params['diel_list'])):

        Die_L_row_col = get_col_row_range(die_params['diel_begin_list'][die_cnt], die_params['diel_end_list'][die_cnt])
        Die_L_name = die_params['diel_list'][die_cnt]

        die_L={
            "sheet_name": die_params['die_sheet'],
            "sheet_d": wb_d[die_params['die_sheet']],
            "sheet_f": wb_f[die_params['die_sheet']],
            "row_col" : Die_L_row_col,
            "row_min" : Die_L_row_col[0],
            "row_max" : Die_L_row_col[1],
            "col_min" : Die_L_row_col[2],
            "col_max" : Die_L_row_col[3],
            "X" : get_column_letter(Die_L_row_col[3] - 2),
            "Y" : get_column_letter(Die_L_row_col[3] - 1),
            "net_name" : get_column_letter(Die_L_row_col[3]),
            "name" : Die_L_name,
            "die_side":"L"
        }

        Die_R_row_col = get_col_row_range(die_params['dier_begin_list'][die_cnt], die_params['dier_end_list'][die_cnt])
        Die_R_name = die_params['dier_list'][die_cnt]
        die_R={
            "sheet_name": die_params['die_sheet'],
            "sheet_d": wb_d[die_params['die_sheet']],
            "sheet_f": wb_f[die_params['die_sheet']],
            "row_col" : Die_R_row_col,
            "row_min" : Die_R_row_col[0],
            "row_max" : Die_R_row_col[1],
            "col_min" : Die_R_row_col[2],
            "col_max" : Die_R_row_col[3],
            "X" : get_column_letter(Die_R_row_col[3] - 2),
            "Y" : get_column_letter(Die_R_row_col[3] - 1),
            "net_name" : get_column_letter(Die_R_row_col[3]),
            "name": Die_R_name,
            "die_side" : "R"
        }
        mapping_prefix = die_L['name']

        get_gen = gen(adp,die_L,mapping,mapping_prefix,ball,last_ball)
        adp = get_gen[0]
        mapping = get_gen[1]
        ball = get_gen[2]
        last_ball = get_gen[3]
        print(ball)
        print(last_ball)
        get_gen = gen(adp,die_R,mapping,mapping_prefix,ball,last_ball)
        adp = get_gen[0]
        mapping = get_gen[1]
        ball = get_gen[2]
        last_ball = get_gen[3]
    print(ball['begin_row'], ball['end_row'])
    for m in range (ball['begin_row'], ball['end_row'] + 1):
        adp['sheet'][adp['resdef'] + str(adp['r'])].value = "BGA"
        if(str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("VSS") != -1):
            adp['sheet'][adp['Pad_use'] + str(adp['r'])] = "GROUND"
        elif(str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("VDD") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("VCCIO") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("VAA") != -1 
             or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("TC_VDDQ") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("VCCAON") != -1):
            adp['sheet'][adp['Pad_use'] + str(adp['r'])] = "POWER"
        elif(str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_LP_CFG") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_CFG_CLK") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_MODE") != -1):
            pass
        #  (str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_LP_CFG") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_CFG_CLK") != -1 or str(die['sheet_d'][die['net_name'] + str(i)].value).find("RDI_MODE") != -1):
        else:
            adp['sheet'][adp['Pad_use'] + str(adp['r'])] = "NC"
        if(str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_LP_CFG") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_CFG_CLK") != -1 or str(ball['sheet_d'][ball['net_col']+ str(m)].value).find("RDI_MODE") != -1):
            pass
        else:
            adp['sheet'][adp['pkg_name'] + str(adp['r'])].value = ball['sheet_f'][ball['net_col'] + str(m)].value
            
            adp['sheet'][adp['bga_pin'] + str(adp['r'])].value = f"={str(ball['sheet_f'][ball['Y'] + str(m)].value).replace('=','')}&{str(ball['sheet_f'][ball['X'] + str(m)].value).replace('=','')}"   
            adp['r'] += 1
            # print(ball['end_row']) 

    wb_f.save(input_params['excel_file'])
  
   
        

gen_apd()


