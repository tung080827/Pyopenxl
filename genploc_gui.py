
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

def mynotif(content):
        
        myLabel = ttk.Label(root,text=content)
        myLabel.grid(row=4, column=0, columnspan=3,padx=(20, 10), pady=(20, 10), sticky="nsew")
        
        # excel_path = "r"+ e1.get()
        # e1.delete(0,END)
# Create a Frame for the Checkbuttons
# style.configure("TLabelframe", bordercolor="red")
# pst = ttk.Style()
# pst.configure("TLabelframe", font= ('Arial', 15),
# background="red")
ploc_frame = ttk.LabelFrame(root, text="Ploc input config", padding=(20, 10))
ploc_frame.grid(row=0,column=0, columnspan=2, padx=(20, 10), pady=(20, 10), sticky="nsew")

# Create a Frame for input widgets
# widgets_frame = ttk.Frame(check_frame, padding=(0, 0, 0, 10))
# widgets_frame.grid(row=0, column=1, padx=10, pady=(30, 10), sticky="nsew", rowspan=3)
# widgets_frame.columnconfigure(index=0, weight=1)

# Entry
pfont= ("Rosewood Std Regular", 12, "bold")
excel_t = ttk.Label(ploc_frame,text="PLOC path:",border=20,font=pfont, borderwidth=5)
excel_t.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")

excel_i = ttk.Entry(ploc_frame,width=130)
excel_i.insert(0, r"C:\Users\sytung\OneDrive - Synopsys, Inc\Desktop\py\Bump_map.xlsx")
# excel_t.place(x=200, y=20)
excel_i.grid(row=0, column=1, columnspan=4, padx=5, pady=(0, 10), sticky="ew",ipady=10)

sheet_t = ttk.Label(ploc_frame,text="Sheet name:",border=20,font=pfont, borderwidth=3)
sheet_t.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
sheet_i = ttk.Entry(ploc_frame, background="#217346")
sheet_i.insert(0, "EMIB_PUB_GPIO")
sheet_i.grid(row=1, column=1, padx=5, pady=(0, 10), sticky='ew',ipady=10)

sheete_t = ttk.Label(ploc_frame,text="C4 sheet:",border=20,font=pfont, borderwidth=3)
sheete_t.grid(row=1, column=2, padx=5, pady=(0, 10), sticky="e")
sheete_i = ttk.Entry(ploc_frame, background="#217346")
sheete_i.insert(0, "C4 sheet")
sheete_i.grid(row=1, column=3, padx=5, pady=(0, 10), sticky="e",ipady=10)

# Package type selection

pkg_t = ttk.Label(ploc_frame,text="Package type:",border=20,font=pfont, borderwidth=3)
pkg_t.grid(row=2, column=0, padx=5, pady=(0, 10), sticky="ew", ipady=10)
package_combo = ttk.Combobox(ploc_frame, state="readonly", values=package_list, )
package_combo.current(0)
package_combo.place(x=130, y=115, height= 40)
# package_combo.grid(row=2, column=1,padx=5, pady=(0, 10), sticky="ew", ipady=10)
package_combo.bind('<<ComboboxSelected>>', choosemode)

# Checkbuttons
sr_opt = ttk.Checkbutton(ploc_frame, text="With sealring (For TC)", variable=a)
sr_opt.place(x=400, y=115, height=50 )
# sr_opt.grid(row=0, column=0, padx=5, pady=10, sticky="nsew")
# foundry_t = ttk.Label(ploc_frame,text="Foundary:",border=20,font=pfont, borderwidth=3)
# foundry_t.grid(row=3, column=0, padx=5, pady=(0, 10), sticky="ew", ipady=10)
# foundry_combo = ttk.Combobox(ploc_frame, state="readonly", values=foundry_list, )
# foundry_combo.current(0)
# foundry_combo.grid(row=3, column=1,padx=5, pady=(0, 10), sticky="ew", ipady=10)
# foundry_combo.bind('<<ComboboxSelected>>', choosemode)
# package_combo.grid(row=2, column=1, padx=5, pady=(0, 10), sticky="ew")
# Separator
# separator = ttk.Separator(root)
# separator.grid(row=1, column=0, padx=(20, 10), pady=10, sticky="ew")

bumpvisual_frame = ttk.LabelFrame(root, text="Bump map config", padding=(20, 10))
bumpvisual_frame.grid(row=2, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew")

# x1y1_t = ttk.Label(bumpvisual_frame,text="Die start:",border=20, borderwidth=3)
# x1y1_t.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
x1y1_i = ttk.Entry(bumpvisual_frame)
x1y1_i.insert(0, "C11")

x1y1_i.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")

# x2y2_t = ttk.Label(bumpvisual_frame,text="Die end:",border=20, borderwidth=3)
# x2y2_t.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")
x2y2_i = ttk.Entry(bumpvisual_frame)
x2y2_i.insert(0, "AP55")
x2y2_i.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")

# Xget_t = ttk.Label(bumpvisual_frame,text="Row contains X:",border=20, borderwidth=3)
# Xget_t.grid(row=3, column=0, padx=5, pady=(0, 10), sticky="ew")
Xget_i = ttk.Entry(bumpvisual_frame)
Xget_i.insert(0, "9")
Xget_i.grid(row=2, column=0, padx=5, pady=(0, 10), sticky="ew")

# Yget_t = ttk.Label(bumpvisual_frame,text="Column contains Y:",border=20, borderwidth=3)
# Yget_t.grid(row=3, column=1, padx=5, pady=(0, 10), sticky="ew")
Yget_i = ttk.Entry(bumpvisual_frame)
Yget_i.insert(0, "A")
Yget_i.grid(row=2, column=1, padx=5, pady=(0, 10), sticky="ew")

out_table_frame = ttk.LabelFrame(root, text="Output Bump table config", padding=(20, 10))
out_table_frame.grid(row=2, column=1, padx=(20, 10), pady=(20, 10), sticky="nsew")
out_name = ttk.Label(out_table_frame,text="Bump table name:")
out_name.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
out_name_in = ttk.Entry(out_table_frame)
out_name_in.insert(0, "Name")
out_name_in.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

out_col_t = ttk.Label(out_table_frame,text="Out table location:")
out_col_t.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
out_col_i = ttk.Entry(out_table_frame)
out_col_i.insert(0, "O64")
out_col_i.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")

# ---------------------------------------gui for EMIB-------------------------------------------------
separator1 = ttk.Separator(bumpvisual_frame)
separator1.grid(row=3, column=0, padx=(0, 10), pady=10, sticky="ew")
separator2 = ttk.Separator(bumpvisual_frame)
separator2.grid(row=3, column=1, padx=(0, 10), pady=10, sticky="ew")

emib_t = ttk.Label(bumpvisual_frame,text="EMIB:",border=20, borderwidth=3)
emib_t.grid(row=4, column=0, padx=5, pady=(0, 10), sticky="ew")
c4_x1y1_i = ttk.Entry(bumpvisual_frame)
c4_x1y1_i.insert(0, "C4 window top-left")
c4_x1y1_i.grid(row=5, column=0, padx=5, pady=(0, 10), sticky="ew")

# x2y2_t = ttk.Label(bumpvisual_frame,text="Die end:",border=20, borderwidth=3)
# x2y2_t.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")
c4_x2y2_i = ttk.Entry(bumpvisual_frame)
c4_x2y2_i.insert(0, "C4 window bot-right")
c4_x2y2_i.grid(row=5, column=1, padx=5, pady=(0, 10), sticky="ew")

# Xget_t = ttk.Label(bumpvisual_frame,text="Row contains X:",border=20, borderwidth=3)
# Xget_t.grid(row=3, column=0, padx=5, pady=(0, 10), sticky="ew")
c4_Xget_i = ttk.Entry(bumpvisual_frame)
c4_Xget_i.insert(0, "Row contains C4 X value")
c4_Xget_i.grid(row=6, column=0, padx=5, pady=(0, 10), sticky="ew")

# Yget_t = ttk.Label(bumpvisual_frame,text="Column contains Y:",border=20, borderwidth=3)
# Yget_t.grid(row=3, column=1, padx=5, pady=(0, 10), sticky="ew")
c4_Yget_i = ttk.Entry(bumpvisual_frame)
c4_Yget_i.insert(0, "Column contains C4 Y value")
c4_Yget_i.grid(row=6, column=1, padx=5, pady=(0, 10), sticky="ew")
# ------------
u_x1y1_i = ttk.Entry(bumpvisual_frame)
u_x1y1_i.insert(0, "uBump window top-left")
u_x1y1_i.grid(row=7, column=0, padx=5, pady=(0, 10), sticky="ew", ipadx=20)

# x2y2_t = ttk.Label(bumpvisual_frame,text="Die end:",border=20, borderwidth=3)
# x2y2_t.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")
u_x2y2_i = ttk.Entry(bumpvisual_frame)
u_x2y2_i.insert(0, "uBump window bot-right")
u_x2y2_i.grid(row=7, column=1, padx=5, pady=(0, 10), sticky="ew", ipadx=20)

# Xget_t = ttk.Label(bumpvisual_frame,text="Row contains X:",border=20, borderwidth=3)
# Xget_t.grid(row=3, column=0, padx=5, pady=(0, 10), sticky="ew")
u_Xget_i = ttk.Entry(bumpvisual_frame)
u_Xget_i.insert(0, "Row contains uBump X value")
u_Xget_i.grid(row=8, column=0, padx=5, pady=(0, 10), sticky="ew")

# Yget_t = ttk.Label(bumpvisual_frame,text="Column contains Y:",border=20, borderwidth=3)
# Yget_t.grid(row=3, column=1, padx=5, pady=(0, 10), sticky="ew")
u_Yget_i = ttk.Entry(bumpvisual_frame)
u_Yget_i.insert(0, "Column contains uBump Y value")
u_Yget_i.grid(row=8, column=1, padx=5, pady=(0, 10), sticky="ew")
# ------------------------------
separator1 = ttk.Separator(out_table_frame)
separator1.grid(row=3, column=0, padx=(0, 10), pady=10, sticky="ew")
separator2 = ttk.Separator(out_table_frame)
separator2.grid(row=3, column=1, padx=(0, 10), pady=10, sticky="ew")

emib_tb_t = ttk.Label(out_table_frame,text="EMIB:")
emib_tb_t.grid(row=4, column=0, padx=5, pady=(0, 10), sticky="ew")
c4_tb_name = ttk.Entry(out_table_frame)
c4_tb_name.insert(0, "C4 Name")
c4_tb_name.grid(row=5, column=0, padx=5, pady=(0, 10), sticky="ew")

# out_col_t = ttk.Label(out_table_frame,text="Out table location:")
# out_col_t.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
c4_col = ttk.Entry(out_table_frame)
c4_col.insert(0, "C4 location")
c4_col.grid(row=5, column=1, padx=5, pady=(0, 10), sticky="ew")

u_tb_name = ttk.Entry(out_table_frame)
u_tb_name.insert(0, "uBump Name")
u_tb_name.grid(row=6, column=0, padx=5, pady=(0, 10), sticky="ew")

u_col = ttk.Entry(out_table_frame)
u_col.insert(0, "uBump location")
u_col.grid(row=6, column=1, padx=5, pady=(0, 10), sticky="ew")
# out_col.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
# out_row = ttk.Entry(bumpvisual_frame)
# out_row.insert(0, "X axis value get")
# out_row.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

#--------------------------------------------------------------------------------------------------------#
dmbump_frame = ttk.LabelFrame(root, text="Dummy bump config", padding=(20, 10))
dmbump_frame.grid(row=3, column=0, columnspan=2, padx=(20, 10), pady=(20, 10), sticky="nsew")

dmbump_cor1_frame = ttk.LabelFrame(dmbump_frame, text="Corner 1 config", padding=(20, 10))
dmbump_cor1_frame.grid(row=0, column=0,padx=(20, 10), pady=(20, 10), sticky="nsew")

cor1_x1y1 = ttk.Entry(dmbump_cor1_frame, width=33)
cor1_x1y1.insert(0, "window top-left")
cor1_x1y1.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
cor1_x2y2 = ttk.Entry(dmbump_cor1_frame, width=32)
cor1_x2y2.insert(0, "window bot-right")
cor1_x2y2.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

cor1_Xget = ttk.Entry(dmbump_cor1_frame)
cor1_Xget.insert(0, "Row contains X")
cor1_Xget.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
cor1_Yget = ttk.Entry(dmbump_cor1_frame)
cor1_Yget.insert(0, "Column contains Y")
cor1_Yget.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")
#---------------------------------------------------------------------------------------------------------#

dmbump_cor2_frame = ttk.LabelFrame(dmbump_frame, text="Corner 2 config", padding=(20, 10))
dmbump_cor2_frame.grid(row=0, column=1, padx=(20, 10), pady=(20, 10), sticky="nsew")

cor2_x1y1 = ttk.Entry(dmbump_cor2_frame, width=25)
cor2_x1y1.insert(0, "window top-left")
cor2_x1y1.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
cor2_x2y2 = ttk.Entry(dmbump_cor2_frame, width=25)
cor2_x2y2.insert(0, "window bot-right")
cor2_x2y2.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

cor2_Xget = ttk.Entry(dmbump_cor2_frame)
cor2_Xget.insert(0, "Row contains X")
cor2_Xget.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
cor2_Yget = ttk.Entry(dmbump_cor2_frame)
cor2_Yget.insert(0, "Column contains Y")
cor2_Yget.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")

#--------------------------------------------------------------------------------------------------------#
dmbump_cor3_frame = ttk.LabelFrame(dmbump_frame, text="Corner 3 config", padding=(20, 10))
dmbump_cor3_frame.grid(row=1, column=0, padx=(20, 10), pady=(20, 10), sticky="nsew")

cor3_x1y1 = ttk.Entry(dmbump_cor3_frame, width=33)
cor3_x1y1.insert(0, "window top-left")
cor3_x1y1.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
cor3_x2y2 = ttk.Entry(dmbump_cor3_frame, width=32)
cor3_x2y2.insert(0, "window bot-right")
cor3_x2y2.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

cor3_Xget = ttk.Entry(dmbump_cor3_frame)
cor3_Xget.insert(0, "Row contains X")
cor3_Xget.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
cor3_Yget = ttk.Entry(dmbump_cor3_frame)
cor3_Yget.insert(0, "Column contains Y")
cor3_Yget.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")
#--------------------------------------------------------------------------------------------------------#
dmbump_cor4_frame = ttk.LabelFrame(dmbump_frame, text="Corner 4 config", padding=(20, 10))
dmbump_cor4_frame.grid(row=1, column=1, padx=(20, 10), pady=(20, 10), sticky="nsew")

cor4_x1y1 = ttk.Entry(dmbump_cor4_frame, width=25)
cor4_x1y1.insert(0, "window top-left")
cor4_x1y1.grid(row=0, column=0, padx=5, pady=(0, 10), sticky="ew")
cor4_x2y2 = ttk.Entry(dmbump_cor4_frame, width=25)
cor4_x2y2.insert(0, "window bot-right")
cor4_x2y2.grid(row=0, column=1, padx=5, pady=(0, 10), sticky="ew")

cor4_Xget = ttk.Entry(dmbump_cor4_frame)
cor4_Xget.insert(0, "Row contains X")
cor4_Xget.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")
cor4_Yget = ttk.Entry(dmbump_cor4_frame)
cor4_Yget.insert(0, "Column contains Y")
cor4_Yget.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="ew")

#--------------------------------------------------------------------------------------------------------#