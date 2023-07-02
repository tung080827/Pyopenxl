from tkinter import messagebox
from tkinter import ttk
import tkinter as tk
from tkinter import *
from ttkthemes import ThemedTk, THEMES
from PIL import Image
from PIL import ImageTk, Image
from tkinter.font import Font
from tkinter import filedialog

def open(wd, excel):
	# global my_image
    wd.filename = filedialog.askopenfilename(initialdir="./", title="Select A File", filetypes=(("excel files", "*.xlsx"),("all files", "*.*")))
    excel.delete(0,END)
    print(wd.filename) 
    excel.insert(0, wd.filename)
    
    