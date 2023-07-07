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
from tkinter.font import Font as tkfont
from tkinter import filedialog
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection

import gui_function as gui
# adv = 1

# label.pack(padx=40,pady=40)
# Create a style
root = ThemedTk()
# my_canvas=tk.Canvas(root)
root.set_theme("scidpurple")

root.title("PLOC TABLE GENERATOR")
root.geometry("1000x1000+30+100")
root.resizable(width=False, height=False)
root.iconbitmap(r".\mylogo.ico")
root.option_add("*tearOff", False) # This is always a good idea

bg = ImageTk.PhotoImage(file=r".\bg3_1.png")
open_imag = PhotoImage(file = r".\open-folder.png")

# Define Canvas
my_canvas = tk.Canvas(root, width=1200, height=800, bd=0, highlightthickness=0)
my_canvas.pack(fill="both", expand=True)

# Put the image on the canvas
my_canvas.create_image(0,0, image=bg, anchor="nw")
# Make the app responsive
# root.columnconfigure(index=0, weight=1)
# root.columnconfigure(index=1, weight=1)
# root.columnconfigure(index=2, weight=1)
# root.columnconfigure(index=3, weight=1)

# root.grid_columnconfigure(0, weight=1)
# root.grid_rowconfigure(0, weight=1)
# root.rowconfigure(index=0, weight=1)
# root.rowconfigure(index=1, weight=1)
# root.rowconfigure(index=2, weight=1)
# root.rowconfigure(index=3, weight=1)
# root.rowconfigure(index=4, weight=1)
# root.rowconfigure(index=5, weight=1)
# root.rowconfigure(index=6, weight=1)
# root.rowconfigure(index=7, weight=1)




stfont= ("Franklin Gothic Medium", 10, 'underline', "italic")
# Create lists for the Comboboxes
theme_list = ["adapta", "aquativo", "arc", "black","blue", "breeze", "clearlooks", "elegance", "equilux", "itft1", "keramik", "keramik_alt", "kroc", "plastik", "radiance", "ubuntu", "scidblue", "scidgreen", "scidgrey", "scidmint", "scidpink", "scidpurple", "scidsand", "smog", "winxpblue", "yaru" ]
package_list = ["S-Organic", "A-CoWoS", "A-EMIB"]
foundry_list = ["TSMC-MapWSR", "TSMC-MapWoSR", "SS-MapWSR", "SS-MapWoSR", "GF-MapWSR", "GF-MapWSR"]
int_couple_number = ["2", "4", "6", "8", "10", "12", "14", "16"]

# Create control variables
a = tk.BooleanVar()
b = tk.BooleanVar(value=True)
c = tk.BooleanVar()
d = tk.IntVar(value=2)
# e = tk.StringVar(value=option_menu_list[1])
f = tk.BooleanVar()
g = tk.DoubleVar(value=75.0)
h = tk.BooleanVar()
tc_opt = tk.IntVar()
isIntp = tk.IntVar()

#Define a Function to enable the frame
def round_rectangle(x1, y1, x2, y2, radius=25, **kwargs):
        
    points = [x1+radius, y1,
              x1+radius, y1,
              x2-radius, y1,
              x2-radius, y1,
              x2, y1,
              x2, y1+radius,
              x2, y1+radius,
              x2, y2-radius,
              x2, y2-radius,
              x2, y2,
              x2-radius, y2,
              x2-radius, y2,
              x1+radius, y2,
              x1+radius, y2,
              x1, y2,
              x1, y2-radius,
              x1, y2-radius,
              x1, y1+radius,
              x1, y1+radius,
              x1, y1]

    return my_canvas.create_polygon(points, **kwargs, smooth=True)

def enable(children):
   for child in children:
      child.configure(state='enable')
def disable(children):  
    for child in children:
        child.configure(state='disable')
def entry_disable(*entries):
    for entry in entries:
        entry.config(state='disable')
def entry_enable(*entries):
    for entry in entries:
        entry.config(state='normal')
def entry_toggle():
    print("TUng day")
        # if(entry['state'] == 'disable'):
    if(tc_opt.get() == 1):
        foundry_combo.config(state='normal')
    elif(tc_opt.get() == 0):
        foundry_combo.config(state='disable')
def intp_toggle():
    print("interposer day")
        # if(entry['state'] == 'disable'):
    if(isIntp.get() == 1):
        entry_enable(xwidth_i, yheight_i, Die1_xoffset_i, Die1_yoffset_i, Die2_xoffset_i, Die2_yoffset_i, intp_sheet, Die1_name, Die2_name, int_tb_loc)
       
    elif(isIntp.get() == 0):
        entry_disable(xwidth_i, yheight_i, Die1_xoffset_i, Die1_yoffset_i, Die2_xoffset_i, Die2_yoffset_i, intp_sheet, Die1_name, Die2_name, int_tb_loc)
        
def progress_bar(value):
    progress['value'] = value
    root.update_idletasks()

def choosetheme(event):
    for theme in theme_list:
        if (theme_combo.get() == theme):
            root.set_theme(theme)
def choosemode(event):   
    if(package_combo.get() == "S-Organic"):
       
       entry_disable(cor1_x1y1, cor1_x2y2, cor1_Xget, cor1_Yget,
                     cor2_x1y1, cor2_x2y2, cor2_Xget, cor2_Yget,
                     cor3_x1y1, cor3_x2y2, cor3_Xget, cor3_Yget,
                     cor4_x1y1, cor4_x2y2, cor4_Xget, cor4_Yget)
      
       entry_enable(x1y1_i, x2y2_i, Xget_i, Yget_i)
       entry_enable(out_name_in, out_col_i)
       entry_disable(sheete_i, sheete_t)
       sheet_t['text']= "Bump sheet:"
        
       print(1)
    elif(package_combo.get() == "A-CoWoS"):
    #    enable(dmbump_frame.winfo_children())
        # cor1_x1y1.config(state='disable')
        entry_enable(cor1_x1y1, cor1_x2y2, cor1_Xget, cor1_Yget,
                     cor2_x1y1, cor2_x2y2, cor2_Xget, cor2_Yget,
                     cor3_x1y1, cor3_x2y2, cor3_Xget, cor3_Yget,
                     cor4_x1y1, cor4_x2y2, cor4_Xget, cor4_Yget)
        entry_enable(x1y1_i, x2y2_i, Xget_i, Yget_i)
        entry_disable(sheete_i, sheete_t)
        entry_enable(out_name_in, out_col_i)
        
        sheet_t['text']= "Bump sheet:"
        print(0)
    else:
        entry_disable(cor1_x1y1, cor1_x2y2, cor1_Xget, cor1_Yget,
                     cor2_x1y1, cor2_x2y2, cor2_Xget, cor2_Yget,
                     cor3_x1y1, cor3_x2y2, cor3_Xget, cor3_Yget,
                     cor4_x1y1, cor4_x2y2, cor4_Xget, cor4_Yget)
        entry_disable(x1y1_i, x2y2_i, Xget_i, Yget_i)
        entry_disable(out_name_in, out_col_i)
        
        entry_enable(sheete_i, sheete_t)
        sheet_t['text']= "uBump sheet:"

        popup("The EMIB package type have not developed yet, Please use S-Organic to gen 2 times (for C4 and uBump) instead!")
    
myLabel = ttk.Label(root,text="Info:")
myLabel_w =my_canvas.create_window(80,770,anchor="nw", window=myLabel)

frame = tk.Label(root, bg="#c9f2dc", font=("Courier New", 10), foreground="#f2a50a")
my_canvas.create_window(600, 80, window=frame, anchor="nw", width= 280, height=100)

def get_num_intdie(event):
    pass
def mynotif(content):
    if(content == ""):
        myLabel.configure(text="", anchor='w')
    else:
        myLabel.configure(text=content, anchor='w')
        # myLabel = ttk.Label(root,text=content)
        # myLabel_w =my_canvas.create_window(80,750,anchor="nw", window=myLabel)
        # myLabel.grid(row=5, column=0, columnspan=2, padx=(20, 10), pady=(20, 10), sticky="nsew")
def process_notify(content):    
        mynotif("")
        root.update_idletasks()
        mynotif(content)
        root.update_idletasks()
# ------------------------------------------------------------------------------------------------------------------------------------------------

def myguide(entries, content):
    if(content == ""):
        entries.configure(text="")
       
    else:
        entries.configure(text=content)

def handle_click(event):
   pass
    
def x1y1_guide(event):
     myguide(frame, "INFO:" + "Die window start cell\n\n - Example:   A0           ")
def un_guide(event):
     myguide(frame,"")

def x2y2_guide(event):
     myguide(frame, "INFO:" + "Die window end cell\n\n - Example:   CU100       ")
def Xget_guide(event):
     myguide(frame, "INFO:" + "Row contains X axis value     \n  which is X location of Bump.\n Must be interger           \n\n - Example: 8                       ")
def Yget_guide(event):
     myguide(frame, "INFO:" + "Row contains Y axis value  \nwhich is Y location of Bump.\n Must be Excel column format\n\n Example: CU ")
def out_name_in_guide(event):
    myguide(frame, "INFO:" + "This field to define the\n   output table name    ")
def out_name2_in_guide(event):
    myguide(frame, "INFO:" + "This field to define the \n  output table 2 name.\n Use for TC with 2 option\n with/without sealring ")
def out_col_in_guide(event):
    myguide(frame, "INFO:" + "This field to define the\n  first output table location. \n The next tables placed away \n2 column from previous table \n\n - Example: O64 ")
def out_col_wsr_i_guide(event):
    myguide(frame, "INFO:" + "This field to define the\n   output table 2 location.\n. Use for TC with 2 option\n with/without sealring\n\n - Example: U64 ")  
def dummystart_guide(event):
    myguide(frame, "INFO:" + "Dummy bump window start cell.\n\n - Example:   A0                   ")  
def dummyend_guide(event):
    myguide(frame, "INFO:" + "Dummy bump window end cell.\n\n - Example:   E3                   ") 
def dummy_Xget_guide(event):
     myguide(frame, "INFO:" + "Row contains X axis values              \n  which is X location of dummy Bump.\n Must be interger                 \n\n- Example: 8                                      ")
def dummy_Yget_guide(event):
     myguide(frame, "INFO:" + "Row contains Y axis value     \nwhich is Y location of dummy Bump. \n Must be Excel column format\n\n- Example: CU               ")
def xwidth_i_guide(event):
     myguide(frame, "INFO:" + "Width of Die/chip. \nThis param used for \n Flip, rotate die/chip to \nput on PKG ")
def yheight_i_guide(event):
     myguide(frame, "INFO:" + "Height of Die/chip. \nThis param used for    \n Flip, rotate die/chip to \nput on PKG            ")
def Die1_xoffset_i_guide(event):
     myguide(frame, "INFO:" + "X Offset of Die1/chip1 . \nThis param used for    \n Die/chip placement on PKG      ")  
def Die1_yoffset_i_guide(event):
     myguide(frame, "INFO:" + "Y Offset of Die1/chip1 . \nThis param used for    \n Die/chip placement on PKG      ")
def Die2_xoffset_i_guide(event):
     myguide(frame, "INFO:" + "X Offset of Die2/chip2 . \nThis param used for    \n Die/chip placement on PKG      ")
def Die2_yoffset_i_guide(event):
     myguide(frame, "INFO:" + "Y Offset of Die2/chip2 . \nThis param used for    \n Die/chip placement on PKG      ")
def intp_sheet_guide(event):
     myguide(frame, "INFO:" + "Name of interposer sheet.\n to put intterposer Die table \n ")
def Die1_name_guide(event):
     myguide(frame, "INFO:" + "Name of interposer Die1.\n Die Flipped + Rotate -90\n ")
def Die2_name_guide(event):
     myguide(frame, "INFO:" + "Name of interposer Die2.\n Die Flipped + Rotate +90\n ")
def int_tb_guide(event):
     myguide(frame, "INFO:" + "First cell to place the table.\n ")
        
xfont = ("System", 12, "bold", 'underline', 'italic')
theme_combo_t = ttk.Label(root,text="Choose theme:",border=20, font=xfont, background='#b434eb', borderwidth=3)
theme_combo_t_w = my_canvas.create_window(750, 15, window=theme_combo_t)

theme_combo = ttk.Combobox(root, state="readonly", values=theme_list, width=15)
theme_combo_w = my_canvas.create_window(870,15, window=theme_combo)
theme_combo.current(0)
theme_combo.bind('<<ComboboxSelected>>', choosetheme)

# -------------------------excelpath input--------------------------#
pfont= ("Rosewood Std Regular", 12, "bold", 'underline' )
excel_t = ttk.Label(root,text="PLOC path:",border=20,font=pfont, borderwidth=5)
excel_t_w = my_canvas.create_window(30,40, anchor="nw", window=excel_t)
excel_i = ttk.Entry(root, width=115)
excel_i.insert(0, r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Bump_CoWoS_S.xlsx")
excel_i_w = my_canvas.create_window(150,40, anchor="nw", window=excel_i)

# -------------------------excel sheet_name input--------------------------#
sheet_t = ttk.Label(root,text="Sheet name:",border=20,font=pfont, borderwidth=3)
sheet_t_w = my_canvas.create_window(30,80, anchor="nw", window=sheet_t)
sheet_i = ttk.Entry(root, background="#217346", width=20)
sheet_i.insert(0, "N3P_CoWoS")
sheet_i_w = my_canvas.create_window(150,80, anchor="nw", window=sheet_i)
# -------------------------excel sheet_name input--------------------------#
sheete_t = ttk.Label(root,text="C4 sheet:",border=20,font=pfont, borderwidth=3)
sheete_t_w = my_canvas.create_window(300,80, anchor="nw", window=sheete_t)
sheete_i = ttk.Entry(root, background="#217346", width=20)
sheete_i.insert(0, "C4 sheet")
sheete_i_w = my_canvas.create_window(400,80, anchor="nw", window=sheete_i)

# -------------------------pkg type input--------------------------#
pkg_t = ttk.Label(root,text="Package type:",border=20,font=pfont, borderwidth=3)
pkg_t_w = my_canvas.create_window(30,120, anchor="nw", window=pkg_t)
package_combo = ttk.Combobox(root, state="readonly", values=package_list, width=17)
package_combo_w = my_canvas.create_window(150,120, anchor="nw", window=package_combo)
package_combo.current(0)
package_combo.bind('<<ComboboxSelected>>', choosemode)
# -------------------------sealring option input--------------------------#
sr_opt = ttk.Checkbutton(root, text="For TC", variable=tc_opt,command= entry_toggle)
sr_opt_w =my_canvas.create_window(300, 120, anchor="nw", window=sr_opt)




# -------------------------foundary selection --------------------------#
foundry_combo = ttk.Combobox(root, state="readonly", values=foundry_list, width=17)
foundry_combo_w = my_canvas.create_window(400, 120, anchor="nw", window=foundry_combo)
foundry_combo.current(0)

# foundry_combo.bind('<<ComboboxSelected>>', choosemode)


# Separator
separator = ttk.Separator(root)
separator_w = my_canvas.create_window(30, 130, anchor="nw", window=separator)


my_canvas.create_text(30, 200, text="Die bump map config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")



# ------------------------Die bump visual parameters input --------------------------#
x1y1_i = ttk.Entry(root, width=20)
my_canvas.create_window(150, 230, anchor="nw", window=x1y1_i)
x1y1_i.insert(0, "C11")
x1y1_i.bind('<FocusIn>', x1y1_guide)
x1y1_i.bind('<FocusOut>', un_guide)


x2y2_i = ttk.Entry(root, width=20)
my_canvas.create_window(300, 230, anchor="nw", window=x2y2_i)
x2y2_i.insert(0, "CW103")
x2y2_i.bind('<FocusIn>', x2y2_guide)
x2y2_i.bind('<FocusOut>', un_guide)

Xget_i = ttk.Entry(root, width=20)
Xget_i_w = my_canvas.create_window(150, 270, anchor="nw", window=Xget_i)
Xget_i.insert(0, "8")
Xget_i.bind('<FocusIn>', Xget_guide)
Xget_i.bind('<FocusOut>', un_guide)

Yget_i = ttk.Entry(root, width=20)
Yget_i_w = my_canvas.create_window(300, 270, anchor="nw", window=Yget_i)
Yget_i.insert(0, "B")
Yget_i.bind('<FocusIn>', Yget_guide)
Yget_i.bind('<FocusOut>', un_guide)

# ------------------------Output table configure --------------------------#
my_canvas.create_text(500, 200, text="Die table config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
my_canvas.create_text(680, 200, text="Sheet:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")

out_tb_sheet = ttk.Entry(root)
out_tb_sheet_w = my_canvas.create_window(750, 195, anchor="nw", window=out_tb_sheet)
out_tb_sheet.insert(0, "Bump coordination")
out_tb_sheet.bind('<FocusIn>', xwidth_i_guide)
out_tb_sheet.bind('<FocusOut>', un_guide)

out_name_in = ttk.Entry(root, width=20)
out_name_in_w = my_canvas.create_window(600, 230, anchor="nw", window=out_name_in)
out_name_in.insert(0, "Name")
out_name_in.bind('<FocusIn>', out_name_in_guide)
out_name_in.bind('<FocusOut>', un_guide)
out_name2_in = ttk.Entry(root, width=20)
out_name2_in_w = my_canvas.create_window(750, 230, anchor="nw", window=out_name2_in)
out_name2_in.insert(0, "Name WO SR")
out_name2_in.bind('<FocusIn>', out_name2_in_guide)
out_name2_in.bind('<FocusOut>', un_guide)


out_col_i = ttk.Entry(root, width=20)
out_col_i_w = my_canvas.create_window(600, 270, anchor="nw", window=out_col_i)
out_col_i.insert(0, "S111")
out_col_i.bind('<FocusIn>', out_col_in_guide)
out_col_i.bind('<FocusOut>', un_guide)



out_col_wsr_i = ttk.Entry(root)
out_col_wsr_w = my_canvas.create_window(750, 270, anchor="nw", window=out_col_wsr_i)
out_col_wsr_i.insert(0, "T64")
out_col_wsr_i.bind('<FocusIn>', out_col_wsr_i_guide)
out_col_wsr_i.bind('<FocusOut>', un_guide)


#------------------------------------Dummybup at 4 corners for Advance package-----------------------------------------------------#

my_canvas.create_text(30, 310, text="Die Bummy bump config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")

my_canvas.create_text(245, 330, text="Corner 1 config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="black")
cor1_x1y1 = ttk.Entry(root, width=20)
cor1_x1y1_w = my_canvas.create_window(150, 350, anchor="nw", window=cor1_x1y1)
cor1_x1y1.insert(0, "C11")
cor1_x1y1.bind('<FocusIn>', dummystart_guide)
cor1_x1y1.bind('<FocusOut>', un_guide)

cor1_x2y2 = ttk.Entry(root, width=20)
cor1_x2y2_w = my_canvas.create_window(300, 350, anchor="nw", window=cor1_x2y2)
cor1_x2y2.insert(0, "E13")
cor1_x2y2.bind('<FocusIn>', dummyend_guide)
cor1_x2y2.bind('<FocusOut>', un_guide)

cor1_Xget = ttk.Entry(root,width=20)
cor1_Xget_w = my_canvas.create_window(150, 380, anchor="nw", window=cor1_Xget)
cor1_Xget.insert(0, "9")
cor1_Xget.bind('<FocusIn>', dummy_Xget_guide)
cor1_Xget.bind('<FocusOut>', un_guide)

cor1_Yget = ttk.Entry(root, width=20)
cor1_Yget_w = my_canvas.create_window(300, 380, anchor="nw", window=cor1_Yget)
cor1_Yget.insert(0, "B")
cor1_Yget.bind('<FocusIn>', dummy_Yget_guide)
cor1_Yget.bind('<FocusOut>', un_guide)
#---------------------------------------------------------------------------------------------------------#

my_canvas.create_text(670, 330, text="Corner 2 config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="black")

cor2_x1y1 = ttk.Entry(root, width=20)
cor2_x1y1_w = my_canvas.create_window(600, 350, anchor="nw", window=cor2_x1y1)
cor2_x1y1.insert(0, "CU11")
cor2_x1y1.bind('<FocusIn>', dummystart_guide)
cor2_x1y1.bind('<FocusOut>', un_guide)

cor2_x2y2 = ttk.Entry(root, width=20)
cor2_x2y2_w = my_canvas.create_window(750, 350, anchor="nw", window=cor2_x2y2)
cor2_x2y2.insert(0, "CW13")
cor2_x2y2.bind('<FocusIn>', dummyend_guide)
cor2_x2y2.bind('<FocusOut>', un_guide)

cor2_Xget = ttk.Entry(root, width=20)
cor2_Xget_w = my_canvas.create_window(600, 380, anchor="nw", window=cor2_Xget)
cor2_Xget.insert(0, "9")
cor2_Xget.bind('<FocusIn>', dummy_Xget_guide)
cor2_Xget.bind('<FocusOut>', un_guide)

cor2_Yget = ttk.Entry(root, width=20)
cor2_Yget_w = my_canvas.create_window(750, 380, anchor="nw", window=cor2_Yget)
cor2_Yget.insert(0, "B")
cor2_Yget.bind('<FocusIn>', dummy_Yget_guide)
cor2_Yget.bind('<FocusOut>', un_guide)

#--------------------------------------------------------------------------------------------------------#

my_canvas.create_text(245, 410, text="Corner 3 config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="black")
cor3_x1y1 = ttk.Entry(root, width=20)
cor3_x1y1_w = my_canvas.create_window(150, 430, anchor="nw", window=cor3_x1y1)
cor3_x1y1.insert(0, "C101")
cor3_x1y1.bind('<FocusIn>', dummystart_guide)
cor3_x1y1.bind('<FocusOut>', un_guide)

cor3_x2y2 = ttk.Entry(root, width=20)
cor3_x2y2_w = my_canvas.create_window(300, 430, anchor="nw", window=cor3_x2y2)
cor3_x2y2.insert(0, "E103")
cor3_x2y2.bind('<FocusIn>', dummyend_guide)
cor3_x2y2.bind('<FocusOut>', un_guide)

cor3_Xget = ttk.Entry(root)
cor3_Xget_w = my_canvas.create_window(150, 460, anchor="nw", window=cor3_Xget)
cor3_Xget.insert(0, "9")
cor3_Xget.bind('<FocusIn>', dummy_Xget_guide)
cor3_Xget.bind('<FocusOut>', un_guide)

cor3_Yget = ttk.Entry(root)
cor3_Yget_w = my_canvas.create_window(300, 460, anchor="nw", window=cor3_Yget)
cor3_Yget.insert(0, "B")
cor3_Yget.bind('<FocusIn>', dummy_Yget_guide)
cor3_Yget.bind('<FocusOut>', un_guide)

#--------------------------------------------------------------------------------------------------------#

my_canvas.create_text(670, 410, text="Corner 4 config", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="black")
cor4_x1y1 = ttk.Entry(root, width=20)
cor4_x1y1_w = my_canvas.create_window(600, 430, anchor="nw", window=cor4_x1y1)
cor4_x1y1.insert(0, "CU101")
cor4_x1y1.bind('<FocusIn>', dummystart_guide)
cor4_x1y1.bind('<FocusOut>', un_guide)

cor4_x2y2 = ttk.Entry(root, width=20)
cor4_x2y2_w = my_canvas.create_window(750, 430, anchor="nw", window=cor4_x2y2)
cor4_x2y2.insert(0, "CW103")
cor4_x2y2.bind('<FocusIn>', dummyend_guide)
cor4_x2y2.bind('<FocusOut>', un_guide)

cor4_Xget = ttk.Entry(root, width=20)
cor4_Xget_w = my_canvas.create_window(600, 460, anchor="nw", window=cor4_Xget)
cor4_Xget.insert(0, "9")
cor4_Xget.bind('<FocusIn>', dummy_Xget_guide)
cor4_Xget.bind('<FocusOut>', un_guide)

cor4_Yget = ttk.Entry(root, width=20)
cor4_Yget_w = my_canvas.create_window(750, 460, anchor="nw", window=cor4_Yget)
cor4_Yget.insert(0, "B")
cor4_Yget.bind('<FocusIn>', dummy_Yget_guide)
cor4_Yget.bind('<FocusOut>', un_guide)

# ---------------------------------------INTERPOSER DIE-------------------------------------------------
interopser = ttk.Checkbutton(root, text="Interposer Die generator", variable=isIntp, onvalue=1, offvalue=0,command= intp_toggle)
interopser_w =my_canvas.create_window(30, 500, anchor="nw", window=interopser)

my_canvas.create_text(30, 540, text="Die/Chip size input:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
xwidth_i = ttk.Entry(root)
xwidth_i_w = my_canvas.create_window(150, 560, anchor="nw", window=xwidth_i, width= 170)
xwidth_i.insert(0, "3938.352")
xwidth_i.bind('<FocusIn>', xwidth_i_guide)
xwidth_i.bind('<FocusOut>', un_guide)

yheight_i = ttk.Entry(root, width=20)
yheight_w = my_canvas.create_window(340, 560, anchor="nw", window=yheight_i, width=170)
yheight_i.insert(0, "2262.872")
yheight_i.bind('<FocusIn>', yheight_i_guide)
yheight_i.bind('<FocusOut>', un_guide)




my_canvas.create_text(500, 540, text="OUT DIE sheet/location:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")

intp_sheet = ttk.Entry(root)
intp_sheet_w = my_canvas.create_window(520, 560, anchor="nw", window=intp_sheet, width=170)
intp_sheet.insert(0, "Package_substrate")
intp_sheet.bind('<FocusIn>', intp_sheet_guide)
intp_sheet.bind('<FocusOut>', un_guide)

int_tb_loc = ttk.Entry(root)
int_tb_locc_w = my_canvas.create_window(710, 560, anchor="nw", window=int_tb_loc, width=170)
int_tb_loc.insert(0, "X111")
int_tb_loc.bind('<FocusIn>', int_tb_guide)
int_tb_loc.bind('<FocusOut>', un_guide)

my_canvas.create_text(30, 595, text="Die/Chip Offset:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
int_die_num = ttk.Label(root,text="Choose die num:",border=20, font=xfont, background='green', borderwidth=3)
int_die_num_w = my_canvas.create_window(300, 610, window=int_die_num)

int_die_num_combo = ttk.Combobox(root, state="readonly", values=int_couple_number, width=15)
int_die_num_combo_w = my_canvas.create_window(450,610, window=int_die_num_combo)
int_die_num_combo.current(0)
int_die_num_combo.bind('<<ComboboxSelected>>', get_num_intdie)

my_canvas.create_text(60, 635, text="Die name:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
Die1_name = ttk.Entry(root)
Die1_name_w = my_canvas.create_window(150, 630, anchor="nw", window=Die1_name, width= 360)
Die1_name.insert(0, "DIE3")
Die1_name.bind('<FocusIn>', Die1_name_guide)
Die1_name.bind('<FocusOut>', un_guide)



Die2_name = ttk.Entry(root)
Die2_name_w = my_canvas.create_window(520, 630, anchor="nw", window=Die2_name, width=360)
Die2_name.insert(0, "DIE7")
Die2_name.bind('<FocusIn>', Die2_name_guide)
Die2_name.bind('<FocusOut>', un_guide)



my_canvas.create_text(60, 675, text="X offset:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
Die1_xoffset_i = ttk.Entry(root)
Die1_xoffset_w = my_canvas.create_window(150, 670, anchor="nw", window=Die1_xoffset_i, width=360)
Die1_xoffset_i.insert(0, "-4350.8")
Die1_xoffset_i.bind('<FocusIn>', Die1_xoffset_i_guide)
Die1_xoffset_i.bind('<FocusOut>', un_guide)

my_canvas.create_text(60, 715, text="Y offset:", anchor="nw",font=("Helvetica", 10, 'italic', 'underline', 'bold'), fill="#b434eb")
Die1_yoffset_i = ttk.Entry(root, width=20)
Die1_yoffset_w = my_canvas.create_window(150, 710, anchor="nw", window=Die1_yoffset_i, width=360)
Die1_yoffset_i.insert(0, "16.2349999999999")
Die1_yoffset_i.bind('<FocusIn>', Die1_yoffset_i_guide)
Die1_yoffset_i.bind('<FocusOut>', un_guide)

Die2_xoffset_i = ttk.Entry(root)
Die2_xoffset_w = my_canvas.create_window(520, 670, anchor="nw", window=Die2_xoffset_i, width=360)
Die2_xoffset_i.insert(0, "1571.96")
Die2_xoffset_i.bind('<FocusIn>', Die2_xoffset_i_guide)
Die2_xoffset_i.bind('<FocusOut>', un_guide)

Die2_yoffset_i = ttk.Entry(root, width=20)
Die2_yoffset_w = my_canvas.create_window(520, 710, anchor="nw", window=Die2_yoffset_i, width=360)
Die2_yoffset_i.insert(0, "97.9849999999997")
Die2_yoffset_i.bind('<FocusIn>', Die2_yoffset_i_guide)
Die2_yoffset_i.bind('<FocusOut>', un_guide)







separator1 = ttk.Separator(root)

separator2 = ttk.Separator(root)




# ------------------------------
separator1 = ttk.Separator(root)

separator2 = ttk.Separator(root)








#--------------------------------------------------------------------------------------------------------#

my_canvas.create_text(880,980, text= "Internal contact: sytung@synopsys.com" ,font=("Helvetica", 8, 'underline'), fill="grey")

def open():
	# global my_image
    root.filename = filedialog.askopenfilename(initialdir="./", title="Select A File", filetypes=(("Excel files", "*.xlsx"),("all files", "*.*")))
    excel_i.delete(0,END)
    print(root.filename) 
    excel_i.insert(0, root.filename)
	# my_image = ImageTk.PhotoImage(Image.open(root.filename))
	# my_image_label = Label(image=my_image).pack()

#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def popup(notif):
    messagebox.showinfo("Notification", notif)

def show_error(error):
    messagebox.showerror("showerror", error)
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def getstring(string: str,c1: str, c2: str):
	cell = string
	idx1 = cell.find(c1)
	cell_2 = cell[:idx1] + cell[idx1+1 :]
	idx2 = cell_2.find(c2)
	if(idx2 != -1):
		
		cell_tmpx = cell[idx1+1:idx2+1]
		if(cell_tmpx.find("+") != -1 or cell_tmpx.find("-") != -1 or cell_tmpx.find("*") != -1  or cell_tmpx.find("/") != -1):
			return 1
		else: return 0
	else:
		return 1

def get_params_and_generate():
    # popup("Generating...")
    # button['state'] = tk.DISABLED

    mynotif("Processing the input parameter...")
    button['text']="Generating..."
    
    progress_bar(20)
    

    global excel_path 
 
    excel_path = excel_i.get()
    bump_visual_sheet = sheet_i.get()
    
    

    # bump_visual_params=[]
    # bump_visual_params.append(die_x1y1)
    # bump_visual_params.append(die_x2y2)
    # bump_visual_params.append(die_x_get)
    # bump_visual_params.append(die_y_get)

   
    die_table={
        "sheet": out_tb_sheet.get(),
        "name": out_name_in.get(),
        "location": out_col_i.get(),
        "name_wsr": out_name2_in.get(),
        "location_wsr": out_col_wsr_i.get(),
        
    }

    #---Bump map visual view parameter---#
    die_coor = {
        
        "window1": x1y1_i.get(), #Top Left of Bump map visual view
        "window2": x2y2_i.get(), #Bottom Right of Bump map visual view
        "xcoor": Xget_i.get(), #This define row where Xaxis value can be got
        "ycoor": Yget_i.get() #This define row where Yaxis value can be got
    }

    #---Dummy Bump visual view parameter---#
    dummybump={
        "corner_1":{
            "window1": cor1_x1y1.get(),
            "window2": cor1_x2y2.get(),
            "xcoor": cor1_Xget.get(),
            "ycoor": cor1_Yget.get()
            },
        "corner_2":{
          
            "window1": cor2_x1y1.get(),
            "window2": cor2_x2y2.get(),
            "xcoor": cor2_Xget.get(),
            "ycoor": cor2_Yget.get()
        },
        "corner_3":{
         
            "window1": cor3_x1y1.get(),
            "window2": cor3_x2y2.get(),
            "xcoor": cor3_Xget.get(),
            "ycoor": cor3_Yget.get()
        },
        "corner_4":{         
            "window1": cor4_x1y1.get(),
            "window2": cor4_x2y2.get(),
            "xcoor": cor4_Xget.get(),
            "ycoor": cor4_Yget.get()
        }

    }
# Die interposet prarams
    die_params={
        "chip_width": xwidth_i.get(),
        "chip_height": yheight_i.get(),
        "die1_xoffset": Die1_xoffset_i.get(),
        "die1_yoffset": Die1_yoffset_i.get(),
        "die2_xoffset": Die2_xoffset_i.get(),
        "die2_yoffset": Die2_yoffset_i.get(),
    }

    int_die_tb={
        "sheet": intp_sheet.get(),
        "Die1_name": Die1_name.get(),
        "int_tb_location": int_tb_loc.get(),
        "Die2_name": Die2_name.get(),
   
    }
    
    package = package_combo.get()
    print(package)
    if (package == "A-CoWoS"):
        print(1)
        package_type = 1
    elif(package == "S-Organic"):
        # dmbump_frame.state = tk.DISABLED
        package_type = 0
    else:
        package_type = 0

    isTC = tc_opt.get()
    int_gen = isIntp.get()

    if(isTC == 1 and foundry_combo.get() == "TSMC-MapWSR"):
        opt_sr = 1
    elif(isTC == 1 and foundry_combo.get() == "TSMC-MapWoSR"):
        opt_sr = 2
    elif(isTC == 3 and foundry_combo.get() == "SS-MapWSR"):
        opt_sr = 4
    elif(isTC == 1 and foundry_combo.get() == "SS-MapWoSR"):
        opt_sr = 5
    elif(isTC == 1 and foundry_combo.get() == "GF-MapWoSR"):
        opt_sr = 6
    elif(isTC == 1 and foundry_combo.get() == "GF-MapWoSR"):
        opt_sr = 7
    else:
        opt_sr = 0
    generate_bump_table(excel_path, bump_visual_sheet, package_type, die_table, die_coor, dummybump, opt_sr, die_params, int_die_tb, int_gen)
    button['text']="Generate"
    


def generate_bump_table(excel_path, bump_visual_sheet, package_type, die_table, die_coor, dummybump, opt_sr, die_params, int_die_tb, int_gen ):


    mynotif("")
    root.update_idletasks()
    mynotif("Loading the ploc file...")
    root.update_idletasks()
    try:
        # wb_d = load_workbook(excel_path, data_only=True)
        wb_f = load_workbook(excel_path)
        print(wb_f)   
    except:
        print("Wrong Ploc path or Ploc file is openning. Please recheck/close the PLOC file before generate :(")
        show_error("Wrong Ploc path or Ploc file is openning. Please recheck/close the PLOC file before generate :(")
        progress_bar(0)
        mynotif("Error")
        root.update_idletasks()
        return
    
    # ws = wb_f.create_sheet('Tung')
    try:
        sheet_list = wb_f.sheetnames
       #wsvisual_d = wb_d[bump_visual_sheet]  # use for further function
        wsvisual_f = wb_f[bump_visual_sheet]
      # wsdiebump_d = wb_d[die_table['sheet']] # use for further function
        # wsdiebump_f = wb_f[die_table['sheet']]
        # wsintbump_f = wb_f[int_die_tb['sheet']]
        if die_table['sheet'] in sheet_list:
            wsdiebump_f = wb_f[die_table['sheet']]
        else:
            msg_ws = messagebox.askquestion('Create Sheet', 'The ' + die_table['sheet'] + ' doesn\'t exist. Do you want to create it?',icon='question')
            mynotif("")
            mynotif("The " + die_table['sheet'] + " doesn't exist.")
            if(msg_ws == 'yes'):
                mynotif("")
                mynotif('Creating the sheet...')
                wsdiebump_f = wb_f.create_sheet(die_table['sheet'])
            else:
                mynotif("")
                progress_bar(0)
                return
        if(int_gen == 1):
            if int_die_tb['sheet'] in sheet_list:
                wsintbump_f = wb_f[int_die_tb['sheet']]
            else:
                mynotif("")
                mynotif("The " + int_die_tb['sheet'] + " doesn't exist.")
                msg_ws = messagebox.askquestion('Create Sheet', 'The ' + int_die_tb['sheet'] + ' doesn\'t exist. Do you want to create it?', icon='question')
            
                if(msg_ws == 'yes'):
                    wsintbump_f = wb_f.create_sheet(int_die_tb['sheet'])
                    mynotif("")
                    mynotif('Creating the sheet...')
                else:
                    mynotif("")
                    progress_bar(0)
                    return
            
      
        
       
       #wsintbump_d = wb_d[int_die_tb['sheet']] # use for further function
       
    except:
        print("Sheet " + bump_visual_sheet + " doesn't exist")
        show_error("Sheet " + bump_visual_sheet + " doesn't exist")
        progress_bar(0)
        mynotif("Error")
        root.update_idletasks()
        return
    
    

    #----- Create dummy bump at 4 corner 140x140u for advance package (CoWos)-----------#
    ymin = coordinate_to_tuple(die_coor['window1'])[0]
    xmin = coordinate_to_tuple(die_coor['window1'])[1]
    ymax = coordinate_to_tuple(die_coor['window2'])[0]
    xmax = coordinate_to_tuple(die_coor['window2'])[1]

    print(xmin,xmax)
    print(ymin,ymax)
    progress_bar(60)
    if(opt_sr == 0):
        try:
          #----- Create table from bump map-----------#
            die_tb_x = coordinate_to_tuple(die_table['location'])[1]
            die_tb_y = coordinate_to_tuple(die_table['location'])[0]
            int_tb_x = coordinate_to_tuple(int_die_tb['int_tb_location'])[1]
            int_tb_y = coordinate_to_tuple(int_die_tb['int_tb_location'])[0]
            

            r_die = die_tb_y + 2
            r_int = int_tb_y + 2

            title_bg_fill = PatternFill(patternType='solid', fgColor='9e42f5')
            subtil_bg_fill = PatternFill(patternType='solid',fgColor='0e7bf0')
            wsdiebump_f[die_table['location']].value = die_table['name']
           
            wsdiebump_f.merge_cells(die_table['location'] + ":" + get_column_letter(die_tb_x + 2) + str(die_tb_y))
            for c1 in range(0,3):
                wsdiebump_f[get_column_letter(die_tb_x + c1) + str(die_tb_y)].fill = title_bg_fill

            wsdiebump_f[get_column_letter(die_tb_x) + str(die_tb_y + 1)].value = "X"

            wsdiebump_f[get_column_letter(die_tb_x + 1) + str(die_tb_y + 1)].value = "Y"

            wsdiebump_f[get_column_letter(die_tb_x + 2)  + str(str(die_tb_y + 1))].value = "Bump name"
            for c2 in range(0,3):
                wsdiebump_f[get_column_letter(die_tb_x + c2) + str(die_tb_y + 1)].fill = subtil_bg_fill
            

            # xwidth = float (ws_f[get_column_letter(xmax) + die_coor["xcoor"]].value)
            # minxval = float (ws_f[get_column_letter(xmin) + die_coor["xcoor"]].value)
            # ywidth = float (ws_f[die_coor["ycoor"] + str(ymin)].value)
            # minyval = float (ws_f[die_coor["ycoor"] + str(ymax)].value)
            # xwidth = ws_f[get_column_letter(xmax) + die_coor["xcoor"]].value
            # minxval = ws_f[get_column_letter(xmin) + die_coor["xcoor"]].value
            # ywidth = ws_f[die_coor["ycoor"] + str(ymin)].value
            # minyval = ws_f[die_coor["ycoor"] + str(ymax)].value
            if (package_type == 1):
                print("Generate for Advance Package")
                dm_bump_coor= []
                dm_cnt=0
                mynotif("")
                root.update_idletasks()
                mynotif("Generating Dummy bump...")
                root.update_idletasks()
                for dm_bump in dummybump:
                    bump = list(dummybump[dm_bump].values())
                        
                    ymin_dm = coordinate_to_tuple(bump[0])[0]
                    xmin_dm = coordinate_to_tuple(bump[0])[1]
                    ymax_dm = coordinate_to_tuple(bump[1])[0]
                    xmax_dm = coordinate_to_tuple(bump[1])[1]
                    xcoor_dm = str(bump[2])
                    ycoor_dm = str(bump[3])

                    print(xmin_dm,xmax_dm)
                    print(ymin_dm,ymax_dm)
                    print(dummybump)

                    for dummycol1 in range(xmin_dm, xmax_dm + 1):
                        for dummyrow1 in range(ymin_dm, ymax_dm + 1):
                            col_dm = get_column_letter(dummycol1)
                            if (wsvisual_f[col_dm + str(dummyrow1)].value != None):
                                
                                # Den dummy bump table
                                wsdiebump_f[get_column_letter(die_tb_x)+str(r_die)].value = f"='{bump_visual_sheet}'!{col_dm + xcoor_dm}"
                                wsdiebump_f[get_column_letter(die_tb_x + 1)+str(r_die)].value = f"='{bump_visual_sheet}'!{ycoor_dm + str(dummyrow1)}"
                                wsdiebump_f[get_column_letter(die_tb_x + 2)+str(r_die)].value =  f"='{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}"

                                r_die += 1
                                coor = col_dm + str(dummyrow1)
                                dm_bump_coor.append(coor)
                                dm_cnt += 1

                                if(int_gen == 1):
                                # #----------------------------Flip bump map in y axis - Rotate -90 - Rotate +90---------------------------
                              
                                    wsintbump_f[get_column_letter(int_tb_x)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{str(col_dm + xcoor_dm)})" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 5)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{str(col_dm + xcoor_dm)})+({str(die_params['die1_yoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 9)+str(r_int)].value = f"=('{bump_visual_sheet}'!{str(col_dm + xcoor_dm)})+({str(die_params['die2_yoffset']).replace('=','')})" # Rotate +90
                            
                                    wsintbump_f[get_column_letter(int_tb_x + 1)+str(r_int)].value = f"='{bump_visual_sheet}'!{ycoor_dm + str(dummyrow1)}" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 4)+str(r_int)].value = f"=({str(die_params['chip_height']).replace('=','')})-('{bump_visual_sheet}'!{ycoor_dm + str(dummyrow1)})+({str(die_params['die1_xoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 8)+str(r_int)].value = f"=('{bump_visual_sheet}'!{ycoor_dm + str(dummyrow1)})+({str(die_params['die2_xoffset']).replace('=','')})" # Rotate +90

                                    wsintbump_f[get_column_letter(int_tb_x + 2)+str(r_int)].value =  f"='{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}" #Flip Y axis
                                    if(wsvisual_f[col_dm+ str(dummyrow1)].value == "VSS"):
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}" # Rotate +90
                                    else:
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"=\"{int_die_tb['Die1_name']}_\"&'{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"=\"{int_die_tb['Die2_name']}_\"&'{bump_visual_sheet}'!{col_dm+ str(dummyrow1)}" # Rotate +90

                                    r_int += 1
                               

                #---------Create Die bump exclued dummy bump at 4 corner-----------#

                match = 0
                mynotif("")
                root.update_idletasks()
                mynotif("Generating Die bump...")
                root.update_idletasks()
                for col in range(xmin, xmax + 1):
                    for row in range(ymin, ymax + 1):       
                        col_l = get_column_letter(col)
                        #print(col_l)
                        i = 0 
                        while(i < len(dm_bump_coor)):
                            xy = col_l + str(row)
                            if(xy ==  dm_bump_coor[i]):
                                match = 1
                            else:
                                match = 0
                            if(match == 1):
                                break
                            i += 1
                        if (match == 0 and wsvisual_f[col_l + str(row)].value != None):
                            #  get the X value from Visual bump sheet
                            if (wsvisual_f[col_l + str(row)].value != None):
                           
                                #  get the X value from Visual bump sheet
                            
                                wsdiebump_f[get_column_letter(die_tb_x)+str(r_die)].value = f"='{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])}"
                         
                                # #  get the Y value from Visual bump sheet
                         
                                wsdiebump_f[get_column_letter(die_tb_x + 1)+str(r_die)].value = f"='{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)}"
                                
                                #  get the Bump name from Visual bump sheet
                                wsdiebump_f[get_column_letter(die_tb_x + 2)+str(r_die)].value =  f"='{bump_visual_sheet}'!{col_l+ str(row)}"
                                r_die += 1
                                
                                if(int_gen == 1):
                                    wsintbump_f[get_column_letter(int_tb_x ) + str(int_tb_y)].value = "Die Flipped by Y axis"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 2) + str(int_tb_y))
                                    for c3 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + c3) + str(die_tb_y)].fill = title_bg_fill

                                    wsintbump_f[get_column_letter(int_tb_x) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 1) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 2)  + str(str(int_tb_y + 1))].value = "Bump name"
                                    for c4 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + c4) + str(die_tb_y + 1)].fill = subtil_bg_fill

                                    wsintbump_f[get_column_letter(int_tb_x + 4) + str(int_tb_y)].value =  str(int_die_tb['Die1_name']) + " = Die Flipped rotate -90 + Die1 offset"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x + 4) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 6) + str(int_tb_y))
                                    for c5 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + 4 + c5) + str(die_tb_y)].fill = title_bg_fill
                                    wsintbump_f[get_column_letter(int_tb_x + 4) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 5) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 6)  + str(str(int_tb_y + 1))].value = "Bump name"
                                    for c6 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + 4 + c6) + str(die_tb_y + 1)].fill = subtil_bg_fill

                                    wsintbump_f[get_column_letter(int_tb_x + 8) + str(int_tb_y)].value = str(int_die_tb['Die2_name']) + " = Die Flipped rotate +90 + Die2 offset"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x + 8) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 10) + str(int_tb_y))
                                    for c7 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + 8 + c7) + str(die_tb_y)].fill = title_bg_fill
                                    wsintbump_f[get_column_letter(int_tb_x + 8) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 9) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 10)  + str(str(int_tb_y + 1))].value = "Bump name"
                                    for c8 in range(0,3):
                                        wsintbump_f[get_column_letter(int_tb_x + 8 + c8) + str(die_tb_y + 1)].fill = subtil_bg_fill
                                     #----------------------------Flip bump map in y axis - Rotate -90 - Rotate +90---------------------------
                                    
                                    wsintbump_f[get_column_letter(int_tb_x)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 5)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})+({str(die_params['die1_yoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 9)+str(r_int)].value = f"=('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})+({str(die_params['die2_yoffset']).replace('=','')})" # Rotate +90
                                    
                                    wsintbump_f[get_column_letter(int_tb_x + 1)+str(r_int)].value = f"='{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)}" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 4)+str(r_int)].value = f"=({str(die_params['chip_height']).replace('=','')})-('{bump_visual_sheet}'!{die_coor['ycoor']+str(row)})+({str(die_params['die1_xoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 8)+str(r_int)].value = f"=('{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)})+({str(die_params['die2_xoffset']).replace('=','')})" # Rotate +90
                                    
                                    wsintbump_f[get_column_letter(int_tb_x + 2)+str(r_int)].value =  f"='{bump_visual_sheet}'!{col_l+ str(row)}" #Flip Y axis
                                    if(wsvisual_f[col_l+ str(row)].value == "VSS"):
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate +90
                                    else:
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"=\"{int_die_tb['Die1_name']}_\"&'{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"=\"{int_die_tb['Die2_name']}_\"&'{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate +90

                                    r_int += 1
                                    
            else:
                process_notify("Generating Die bump...")
                for col in range(xmin, xmax + 1):
                        for row in range(ymin , ymax + 1):       
                            col_l = get_column_letter(col)
                            #print(col_l)
                            if (wsvisual_f[col_l + str(row)].value != None):
                           
                                #  get the X value from Visual bump sheet
                            
                                wsdiebump_f[get_column_letter(die_tb_x)+str(r_die)].value = f"='{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])}"
                         
                                # #  get the Y value from Visual bump sheet
                         
                                wsdiebump_f[get_column_letter(die_tb_x + 1)+str(r_die)].value = f"='{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)}"
                                
                                #  get the Bump name from Visual bump sheet
                                wsdiebump_f[get_column_letter(die_tb_x + 2)+str(r_die)].value =  f"='{bump_visual_sheet}'!{col_l+ str(row)}"
                                r_die += 1
                                if(int_gen == 1):
                                    wsintbump_f[get_column_letter(int_tb_x ) + str(int_tb_y)].value = "Die Flipped by Y axis"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 2) + str(int_tb_y))
                                    wsintbump_f[get_column_letter(int_tb_x) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 1) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 2)  + str(str(int_tb_y + 1))].value = "Bump name"

                                    wsintbump_f[get_column_letter(int_tb_x + 4) + str(int_tb_y)].value =  str(int_die_tb['Die1_name']) + " = Die Flipped rotate -90 + Die1 offset"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x + 4) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 6) + str(int_tb_y))
                                    wsintbump_f[get_column_letter(int_tb_x + 4) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 5) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 6)  + str(str(int_tb_y + 1))].value = "Bump name"

                                    wsintbump_f[get_column_letter(int_tb_x + 8) + str(int_tb_y)].value = str(int_die_tb['Die2_name']) + " = Die Flipped rotate +90 + Die2 offset"
                                    wsintbump_f.merge_cells(get_column_letter(int_tb_x + 8) + str(int_tb_y) + ":" + get_column_letter(int_tb_x + 10) + str(int_tb_y))
                                    wsintbump_f[get_column_letter(int_tb_x + 8) + str(int_tb_y + 1)].value = "X"
                                    wsintbump_f[get_column_letter(int_tb_x + 9) + str(int_tb_y + 1)].value = "Y"
                                    wsintbump_f[get_column_letter(int_tb_x + 10)  + str(str(int_tb_y + 1))].value = "Bump name"
                                    
                                    #----------------------------Flip bump map in y axis - Rotate -90 - Rotate +90---------------------------
                                    
                                    wsintbump_f[get_column_letter(int_tb_x)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 5)+str(r_int)].value = f"=({str(die_params['chip_width']).replace('=','')})-('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})+({str(die_params['die1_yoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 9)+str(r_int)].value = f"=('{bump_visual_sheet}'!{col_l + str(die_coor['xcoor'])})+({str(die_params['die2_yoffset']).replace('=','')})" # Rotate +90
                                    
                                    wsintbump_f[get_column_letter(int_tb_x + 1)+str(r_int)].value = f"='{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)}" # Flip Y axis
                                    wsintbump_f[get_column_letter(int_tb_x + 4)+str(r_int)].value = f"=({str(die_params['chip_height']).replace('=','')})-('{bump_visual_sheet}'!{die_coor['ycoor']+str(row)})+({str(die_params['die1_xoffset'])})" # Rotate -90
                                    wsintbump_f[get_column_letter(int_tb_x + 8)+str(r_int)].value = f"=('{bump_visual_sheet}'!{die_coor['ycoor'] + str(row)})+({str(die_params['die2_xoffset']).replace('=','')})" # Rotate +90
                                    
                                    wsintbump_f[get_column_letter(int_tb_x + 2)+str(r_int)].value =  f"='{bump_visual_sheet}'!{col_l+ str(row)}" #Flip Y axis
                                    if(wsvisual_f[col_l+ str(row)].value == "VSS"):
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"='{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate +90
                                    else:
                                        wsintbump_f[get_column_letter(int_tb_x + 6)+str(r_int)].value = f"=\"{int_die_tb['Die1_name']}_\"&'{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate -90
                                        wsintbump_f[get_column_letter(int_tb_x + 10)+str(r_int)].value = f"=\"{int_die_tb['Die2_name']}_\"&'{bump_visual_sheet}'!{col_l+ str(row)}" # Rotate +90

                                    r_int += 1
                # tab = Table(displayName="Table1", ref="O65:Q500")
                # ws_f.add_table(tab)
            
            progress_bar(80)    
            wb_f.save(excel_path)
            progress_bar(100)
            mynotif("Generated")
            popup("PLOC generated successful!!!")
            mynotif("")
        except (ValueError):
            print ("Wrong input, Please check and regenerate")
            show_error("Wrong input, Please check and regenerate")
            progress_bar(0)
            mynotif("Error")
            root.update_idletasks()
            return
        except:
            print('Error!!!')
            
            show_error("There are an error in caculations, Please recheck and make sure the input is correct!")
            progress_bar(0)
            mynotif("Error")
            root.update_idletasks()
            return
            
            
    # elif(opt_sr == 1):

    #     tb_x = coordinate_to_tuple(die_table['location'])[1]
    #     tb_y = coordinate_to_tuple(die_table['location'])[0]
    #     tb_x2 = coordinate_to_tuple(die_table['location_wsr'])[1]
    #     tb_y2 = coordinate_to_tuple(die_table['location_wsr'])[0]
    #     r = tb_y + 2
    #     r2 = tb_y2 + 2

    #     ws_f[die_table['location']].value = die_table['name']
    #     ws_f[die_table['location_wsr']].value = die_table['name_wsr']
    #     # ws_f.merge_cells(die_table['xcol'] + str(die_table['begin']) + ":" + die_table['bumpcol'] + str(die_table['begin']))
    #     # print(die_table['xcol'] + str(die_table['begin']) + ":" + die_table['bumpcol'] + str(die_table['begin']))
    #     ws_f[get_column_letter(tb_x) + str(tb_y + 1)].value = "X"
    #     ws_f[get_column_letter(tb_x2) + str(tb_y2 + 1)].value = "X"
    #     ws_f[get_column_letter(tb_x + 1) + str(tb_y + 1)].value = "Y"
    #     ws_f[get_column_letter(tb_x2 + 1) + str(tb_y2 + 1)].value = "Y"
    #     ws_f[get_column_letter(tb_x + 2)  + str(str(tb_y + 1))].value = "Bump name"
    #     ws_f[get_column_letter(tb_x2 + 2)  + str(str(tb_y2 + 1))].value = "Bump name"
    #     if (package_type == 1):
    #         dm_bump_coor= []
    #         dm_cnt=0
    #         mynotif("")
    #         root.update_idletasks()
    #         mynotif("Generating Dummy bump...")
    #         root.update_idletasks()
    #         for dm_bump in dummybump:
    #             bump = list(dummybump[dm_bump].values())
                    
    #             ymin_dm = coordinate_to_tuple(bump[0])[0]
    #             xmin_dm = coordinate_to_tuple(bump[0])[1]
    #             ymax_dm = coordinate_to_tuple(bump[1])[0]
    #             xmax_dm = coordinate_to_tuple(bump[1])[1]
    #             xcoor_dm = str(bump[2])
    #             ycoor_dm = str(bump[3])

    #             print(xmin_dm,xmax_dm)
    #             print(ymin_dm,ymax_dm)

    #             for dummycol1 in range(xmin_dm, xmax_dm + 1):
    #                 for dummyrow1 in range(ymin_dm, ymax_dm + 1):
    #                     col_dm = get_column_letter(dummycol1)
    #                     if (ws_f[col_dm + str(dummyrow1)].value != None):
    #                         ws_f[get_column_letter(tb_x + 2)+str(r)].value =  ws_f[col_dm+ str(dummyrow1)].value
    #                         ws_f[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws_f[col_dm+ str(dummyrow1)].value 
                      
    #                         ws_f[get_column_letter(tb_x)+str(r)].value = ws_f[col_dm + xcoor_dm].value
                           
    #                         ws_f[get_column_letter(tb_x + 1)+str(r)].value = ws_f[ycoor_dm + str(dummyrow1)].value
                            
    #                         r = r + 1
    #                         r2 = r2 + 1
    #                         coor = col_dm + str(dummyrow1)
    #                         dm_bump_coor.append(coor)
    #                         dm_cnt += 1

    #         #---------Create Die bump exclued dummy bump at 4 corner-----------#

    #         match = 0
    #         mynotif("")
    #         root.update_idletasks()
    #         mynotif("Generating Die bump...")
    #         root.update_idletasks()
    #         for col in range(xmin, xmax + 1):
    #             for row in range(ymin, ymax + 1):       
    #                 col_l = get_column_letter(col)
    #                 #print(col_l)
    #                 i = 0 
    #                 while(i < len(dm_bump_coor)):
    #                     xy = col_l + str(row)
    #                     if(xy ==  dm_bump_coor[i]):
    #                         match = 1
    #                     else:
    #                         match = 0
    #                     if(match == 1):
    #                         break
    #                     i += 1
    #                 if (match == 0 and ws_f[col_l + str(row)].value != None):
    #                     ws_f[get_column_letter(tb_x + 2)+str(r)].value =  ws_f[col_l+ str(row)].value
    #                     ws_f[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws_f[col_l+ str(row)].value
    #                     print(col_l + " " + str(die_coor['xcoor']))
    #                     ws_f[get_column_letter(tb_x)+str(r)].value = ws_f[col_l + str(die_coor['xcoor'])].value
                       
    #                     print(die_coor['ycoor'] + " " + str(row)) 
    #                     ws_f[get_column_letter(tb_x + 1)+str(r)].value = ws_f[die_coor['ycoor'] + str(row)].value
                        
    #                     r = r + 1
    #                     r2 = r2 + 1
    #     else:
    #         mynotif("")
    #         root.update_idletasks()
    #         mynotif("Generating Die bump...")
    #         root.update_idletasks()
    #         for col in range(xmin, xmax + 1):
    #                 for row in range(ymin , ymax + 1):       
    #                     col_l = get_column_letter(col)
    #                     #print(col_l)
    #                     if (ws_f[col_l + str(row)].value != None):
    #                         ws_f[get_column_letter(tb_x + 2)+str(r)].value =  ws_f[col_l+ str(row)].value
    #                         ws_f[get_column_letter(tb_x2 + 2)+str(r2)].value =  ws_f[col_l+ str(row)].value
    #                         print(col_l + " " + str(die_coor['xcoor']))
    #                         ws_f[get_column_letter(tb_x)+str(r)].value = ws_f[col_l + str(die_coor['xcoor'])].value
                            
    #                         print(die_coor['ycoor'] + " " + str(row)) 
    #                         ws_f[get_column_letter(tb_x + 1)+str(r)].value = ws_f[die_coor['ycoor'] + str(row)].value
                           
    #                         r = r + 1
    #                         r2 = r2 + 1

# nofi = ttk.Entry(root,)

    
    # button['state'] = tk.NORMAL


                
# myButton = tk.Button(root,text="Button", command=get_path)
# myButton.pack()

entry_disable(cor1_x1y1, cor1_x2y2, cor1_Xget, cor1_Yget,
            cor2_x1y1, cor2_x2y2, cor2_Xget, cor2_Yget,
            cor3_x1y1, cor3_x2y2, cor3_Xget, cor3_Yget,
            cor4_x1y1, cor4_x2y2, cor4_Xget, cor4_Yget)

entry_disable(sheete_i, sheete_t)
entry_disable(sr_opt, foundry_combo, out_name2_in, out_col_wsr_i)
entry_disable(xwidth_i, yheight_i, Die1_xoffset_i, Die1_yoffset_i, Die2_xoffset_i, Die2_yoffset_i, intp_sheet, Die1_name, Die2_name, int_tb_loc)
sheet_t['text']= "Bump sheet:"
mynotif("")
treeScroll = ttk.Scrollbar(root, orient = 'vertical')



# root.configure(yscrollcomand)

progress = ttk.Progressbar(root, orient = 'horizontal',
              length = 100, mode = 'determinate')
progress_w = my_canvas.create_window(80,800, anchor="nw", window=progress, width= 800)


# Button
#Create style object
style = ttk.Style()

# #configure style
# style = ttk.Style()
# style.configure('TButton', font =
#                ('calibri', 20, 'bold'),
#                     borderwidth = '4',
#                     width = '80')
style.configure('TCheckbutton', font= ('System', 12, 'underline', 'bold'),
 foreground='black', border=50)
mediumFont = tkfont(
	family="System",
	size=16,
	weight="normal",
	slant="italic",
	underline=1,
	overstrike=0)
def hihi():
    button.configure(font=mediumFont, foreground='white', background='Green')
browse_btn = ttk.Button(root, text="Open File", image=open_imag, command=open)
browse_btn_w = my_canvas.create_window(865, 40, anchor="nw", window=browse_btn)
# button = tk.Button(root, text="Generate",font=("System", 14, 'underline', 'bold'), foreground='white', background='#9b34eb', command=get_path, width=40)
button = tk.Button(root, text="Generate",font = mediumFont, foreground='white', background='#9b34eb', command=get_params_and_generate, width=40)
# button = ttk.Button(root, text="Generate", command=get_path, width=80)

button_w = my_canvas.create_window(300, 860, anchor="nw", window=button)





root.mainloop()

