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
# from ttkthemes import ThemedTk, THEMES
from PIL import Image
from PIL import ImageTk, Image
from tkinter.font import Font
from tkinter import filedialog
import gui_function as gui
import subprocess
import ttkbootstrap as ttk
# from subprocess import call



# root = ThemedTk()
root = ttk.Window(themename='united')
# my_canvas=tk.Canvas(root)
# root.set_theme("scidpurple")

root.title("PLOC SCRIPT MENU")
root.geometry("600x800+30+100")
root.resizable(width=False, height=False)
root.iconbitmap(r".\mylogo.ico")
root.option_add("*tearOff", False) # This is always a good idea

# bg = ImageTk.PhotoImage(file=r".\bg3_1.png")
# bg1 = ImageTk.PhotoImage(file=r".\brain.png")
bg2 = PhotoImage(file=r".\brain.png")
# open_imag = PhotoImage(file = r".\sub\open-folder.png")
bg2 = bg2.subsample(2, 2)
# Define Canvas
my_canvas = tk.Canvas(root, width=600, height=800, bd=0, highlightthickness=0)
my_canvas.pack(fill="both", expand=True)

# Put the image on the canvas
my_canvas.create_image(0,0, image=bg2, anchor="nw")
# Make the app responsive


my_canvas.create_text(240, 100, text="MENU", anchor="nw",font=("Helvetica", 40, 'bold'), fill="black")

stfont= ("Franklin Gothic Medium", 10, 'underline', "italic")
def get_ploc_app():
    subprocess.call([r'.\ploctablegenerator\ploctablegenerator.exe'])
    # call(["python", 'ploc_v2.py'])
def get_ch_app():
    subprocess.call([r'.\datachanel_genv2\datachanel_genv2.exe'])
    #  call(["python", 'datachanel_gen.py'])
# def get_adp_app():
#     subprocess.call([r'.\apdgenerator\apdgenerator.exe'])
#     #  call(["python", 'apdgenerator_v0.2.py'])

# browse_btn = ttk.Button(root, text="Open File", image=open_imag, command=open)
# browse_btn_w = my_canvas.create_window(865, 40, anchor="nw", window=browse_btn)
# button = tk.Button(root, text="Generate",font=("System", 14, 'underline', 'bold'), foreground='white', background='#9b34eb', command=get_path, width=40)
button_ploc = tk.Button(root, text="GENERATE PLOC TABLE BASE ON VISUAL VIEW", foreground='white', background='#9b34eb',font=stfont, command=get_ploc_app, width=40)
# button = ttk.Button(root, text="Generate", command=get_path, width=80)

button_ploc_w = my_canvas.create_window(100, 200, anchor="nw", window=button_ploc, width=400)

button_datach_map = tk.Button(root, text="GENERATE DATA CHANNEL AND MAPPING TABLE", foreground='white', background='#9b34eb',font=stfont, command=get_ch_app, width=40)


my_canvas.create_window(100, 300, anchor="nw", window=button_datach_map, width=400)

# button_adp = tk.Button(root, text="GENERATE ADP NETLIST", foreground='white', background='#9b34eb',font=stfont, command=get_adp_app, width=40)
# button = ttk.Button(root, text="Generate", command=get_path, width=80)

# my_canvas.create_window(100, 400, anchor="nw", window=button_adp, width=400)





root.mainloop()