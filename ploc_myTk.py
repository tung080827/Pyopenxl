from tkinter import messagebox
# from tkinter import ttk
import getcolumn
from array import *
import tkinter as tk
from tkinter import *
# from ttkthemes import ThemedTk, THEMES
from PIL import Image
from PIL import ImageTk, Image
from tkinter.font import Font as tkfont
from tkinter import filedialog
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from time import perf_counter
import time
from ttkbootstrap import icons
from ttkbootstrap.style import StyleBuilderTK


def round_rectangle(canvas :tk.Canvas,x1:float, y1:float, x2:float, y2:float, radius=25, **kwargs):
        
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

    return canvas.create_polygon(points, **kwargs, smooth=True,outline='')

class TkCombobox():
    def __init__(self,
                 canvas : tk.Canvas,
                 x : float,
                 y : float,
                 values : list,
                 w = 140,
                 h=25,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 state = "readonly",
                 current = 0,
                 anchor = 'nw',
                 justify = 'center',
                 guidetext = "",
                 is_bind = 0,
                 ) -> None:
        self.canvas = canvas
        self.x = x
        self.y = y
        self.new_x = self.x
        self.new_y = self.y
        self.values = values
        self.w = w
        self.h = h
        self.new_w = self.w
        self.new_w = self.h
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.state = state
        self.current = current
        self.anchor = anchor
        self.justify = justify
        self.combobox = ttk.Combobox(self.canvas, state=self.state, values=self.values,justify=self.justify) 
        self.place = self.canvas.create_window(self.x,self.y, window=self.combobox, anchor=self.anchor,width= self.w,height=self.h)   
        self.combobox.current(0)
        if is_bind == 1:
            self.combobox.bind('<Enter>', lambda event:  self.guide(guidetext))        
            self.combobox.bind('<Leave>', lambda event: self.unguide())
    def moveto(self, w, h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
        self.new_x = (self.x/self.winx)*w
        self.new_y = (self.y/self.winy)*h
    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            # self.combobox.config(width=int((self.w/1000)*w))
            self.new_w = (self.w/self.winx)*w
        if(h!=0):
            # self.canvas.itemconfig(self.place,height=(self.h/1000)*h)
            # self.combobox.config(height=int((self.h/self.winy)*h))
            # self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.new_h = (self.h/self.winy)*h 
    def get_value(self):
        return self.combobox.get()
    def set_current(self, current):
        self.combobox.current(current)
    def guide(self,gui_list):
        self.myfont = tkfont(family="Courier", size="8" )
        if(self.new_y <= 120):
            y = self.new_y + self.new_h
        else:
            y = self.new_y - 100

        if(self.new_w < 200):
            w = 200
        else: 
            w = self.new_w
        self.guibox = TkTextbox(self.canvas, x=self.new_x, y=y, w=w, h=100)
        self.guibox.textbox.config(font=self.myfont,foreground='#A4A5A4')
        self.combobox.bind('<MouseWheel>', self.text_vertical)

        self.guibox.textbox.config(state='normal')
        self.guibox.textbox.delete("1.0","end")

        # enstate = self.entry["state"]
        # print(enstate)

        # if enstate == 'normal':
        for gui in gui_list:
            self.guibox.textbox.insert(tk.END, gui)
            
        # else:
        #     self.guibox.textbox.insert(tk.END, "This field is not needed for this Package type. Please do not care about it")
        self.guibox.textbox.config(state='disable')
    def unguide(self):
        self.guibox.textbox.destroy()
        self.guibox.scroll_y.destroy()
        self.guibox.scrFrame.destroy()
    def text_vertical(self,event):
        # self.TkTextbox.textbox.yview_scroll(-1 * event.delta, 'units')
        self.guibox.textbox.yview_scroll(-1 * event.delta, 'units')
class TkTextbox():
    def __init__(self,
                 canvas : tk.Canvas,
                 x : float,
                 y : float,
                 w,
                 h,
                 bd = 5,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 anchor = 'nw',
                 relief = 'sunken',
                 wrap = 'word',
                 font = ('arial',10),

                 ) -> None:
        self.canvas = canvas
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.bd = bd
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.anchor = anchor
        self.relief = relief
        self.wrap = wrap
        self.font = font
        # self.frame =tk.Frame(self.canvas)
        # self.place = self.canvas.create_window(x,y,anchor=self.anchor, window=self.frame, width=self.w, height=self.h)
        self.textbox = ttk.Text(self.canvas,foreground='green',relief=self.relief,wrap=self.wrap,font= self.font, highlightthickness=2, highlightbackground='green')
        
        # self.textbox.pack(side=LEFT,fill=BOTH,expand=True)
        self.place = self.canvas.create_window(x,y,anchor=self.anchor, window=self.textbox, width=self.w, height=self.h)
        self.scrFrame = ttk.Frame(self.canvas)
        self.placescr = self.canvas.create_window(self.x+self.w,y,anchor=self.anchor, window=self.scrFrame,height=self.h)
        self.scroll_y = ttk.Scrollbar(self.scrFrame)
        # self.placescr = self.canvas.create_window(self.x+self.w,y,anchor=self.anchor, window=scroll_y,height=self.h)
        self.textbox.config(yscrollcommand=self.scroll_y.set)
        self.scroll_y.pack(side=RIGHT, fill= BOTH,expand=True)
        self.scroll_y.config(command=self.textbox.yview)
    
    def moveto(self, w, h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
        self.canvas.moveto(self.placescr,((self.x+self.w)/self.winx)*w,(self.y/self.winy)*h)

    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            self.new_w = w
        if(h!=0):
            self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.canvas.itemconfig(self.placescr,height=(self.h/self.winy)*h)
            self.new_h = h 
    def add_new_text(self,text):
        self.textbox.config(state='normal')
        self.textbox.delete("1.0","end")
        for gui in text:
            self.textbox.insert(tk.END, gui)
        self.textbox.config(state='disable')
    def remove_text(self):
        self.textbox.config(state='normal')
        self.textbox.delete("1.0","end")
        self.textbox.config(state='disable')
    def add_text(self,text):
        self.textbox.config(state='normal')
        for gui in text:
            self.textbox.insert(tk.END, gui)
        self.textbox.config(state='disable')
    def show_at(self):
        pass
    def configure(self, **kw):
        self.textbox.config(kw)
class Tkentry():
    def __init__(
            self,
           
            canvas : tk.Canvas,
            # TkTextbox : TkTextbox,
            guide_text: list,
            x = 0,
            y = 0,
            w =100,
            h = 25,
            win_defaultx = 1000,
            win_defaulty = 1000,
            justify = 'center',
            fg = 'green') -> None:
        
        self.canvas = canvas
        self.TkTextbox = TkTextbox
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.justify =justify
        self.fg =fg
        self.font = tkfont(family="Bahnschrift",size= 10)
        self.entry = ttk.Entry(self.canvas,foreground=self.fg,justify=self.justify,font=self.font)

        self.bind = self.entry.bind('<Enter>', lambda event:  self.guide(guide_text))
      
        self.bind = self.entry.bind('<Leave>', lambda event: self.unguide())
        self.bind = self.entry.bind('<FocusIn>', lambda event: self.unguide())
        
       
        self.place = canvas.create_window(x, y, window=self.entry,anchor="nw", width=self.w, height =self.h)
        self.new_w = self.w
        self.new_h = self.h
        self.new_x = self.x
        self.new_y = self.y
        
        # self.enstate = self.entry["state"]
    def text_vertical(self,event):
        # self.TkTextbox.textbox.yview_scroll(-1 * event.delta, 'units')
        self.guibox.textbox.yview_scroll(-1 * event.delta, 'units')
    def guide(self,gui_list):
        self.myfont = tkfont(family="Courier", size="8" )
        if(self.new_y <= 120):
            y = self.new_y + self.new_h
        else:
            y = self.new_y - 100

        if(self.new_w < 200):
            w = 200
        else: 
            w = self.new_w
        self.guibox = TkTextbox(self.canvas, x=self.new_x, y=y, w=w, h=100)
        self.guibox.textbox.config(font=self.myfont,foreground='#A4A5A4')
        self.entry.bind('<MouseWheel>', self.text_vertical)

        self.guibox.textbox.config(state='normal')
        self.guibox.textbox.delete("1.0","end")

        enstate = self.entry["state"]
        # print(enstate)

        # if enstate == 'normal':
        for gui in gui_list:
            self.guibox.textbox.insert(tk.END, gui)
            
        # else:
        #     self.guibox.textbox.insert(tk.END, "This field is not needed for this Package type. Please do not care about it")
        self.guibox.textbox.config(state='disable')
    def unguide(self):
        self.guibox.textbox.destroy()
        self.guibox.scroll_y.destroy()
        self.guibox.scrFrame.destroy()

    def add_content(self,*ls):
        for content in ls:
            self.entry.insert(0, content)
    def del_content(self):
        self.entry.delete(0,END)
    def add_new_content(self, content):
        self.entry.delete(0,END)
        self.entry.insert(0, content)
    def destroy_and_createnew(self,x,y):
       self.delete_item()
       self.entry = ttk.Entry(self.canvas)
       self.canvas.create_window(self.x + x, self.y + y, window=self.entry,anchor="nw", width=self.w)
    def delete_item(self):
         self.entry.destroy()
    def get(self):
        return self.entry.get()
        # print(self.entry.get())
    def set_style():
        pass
    def moveto(self,w,h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
        self.new_x = (self.x/self.winx)*w
        self.new_y = (self.y/self.winy)*h
        # print(f"entry: x={(self.x/self.winx)*w} y = {(self.y/self.winy)*h}" )
    def get_width(self):
        return self.new_w
    def get_height(self):
        return self.new_h
    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            self.new_w = (self.w/self.winx)*w
        if(h!=0):
            self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.new_h = (self.h/self.winy)*h 
    def disable(self):
        self.entry.config(state='disable')
        
    def enable(self):
        self.entry.config(state='normal')
    def motion(event):
        x, y = event.x, event.y
        print('{}, {}'.format(x, y))
    def set_fg(self, color:str):
        self.entry.config(foreground=color)
    def change_textsize(self, w, h):
        if(500<=w<=600):
            self.font.config(size=7)
        elif(600<w<=700):
            self.font.config(size=8)
        else:
            self.font.config(size=10)
        self.entry.config(font=self.font)

       
        # return self.new_w
    

class CanvasText():
    def __init__(self,
                 canvas : tk.Canvas, 
                 x : float, 
                 y : float, 
                 text: str,
                 width = 800,
                 height = 20,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 isbg = False,
                 bgx = 120,
                 bgy=35,
                 bg_xo = 10,
                 bg_yo = 10,
                 bg_color = '#78C2AD',
                 anchor="nw", 
                 font = None,
                 fill="#b434eb") -> None:
        self.canvas = canvas
        self.x = x
        self.y = y
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.text = text
        self.anchor = anchor
        if font == None:
            self.font = tkfont(family="Helvetica", size=10, slant='italic', underline=True, weight='bold')
        else:
            self.font = font
        self.w = width
        self.fill = fill
        self.new_x = self.x
        self.new_y = self.y
        self.bgx = bgx
        self.bgy = bgy
        self.bg_xo = bg_xo
        self.bg_yo = bg_yo
        self.bg_color = bg_color
        self.bg : int
        self.isbg = isbg
        self.newxo, self.newyo, self.new__xo, self.new__yo = self.bgx, self.bgy, self.bg_xo, self.bg_yo
        if(self.isbg == True):
            self.bg = round_rectangle(self.canvas, x1=self.x-self.bg_xo, y1=self.y-self.bg_yo, x2=self.x+self.bgx, y2=self.y+self.bgy, fill=bg_color, radius=15)
        self.mytext = self.canvas.create_text(self.x, self.y, text=self.text, anchor=self.anchor,font=self.font, fill=self.fill, width=self.w)

    def change_color(self,color):
        self.fill = color
        self.canvas.itemconfig(self.mytext, fill=self.fill)
    def moveto(self,w,h):
        self.new_x = (self.x/self.winx)*w 
        self.new_y = (self.y/self.winy)*h
        self.canvas.moveto(self.mytext,self.new_x,self.new_y)
        if(self.isbg == True):
            self.canvas.moveto(self.bg, self.new_x -self.bg_xo, self.new_y-self.bg_yo)
    def change_font(self, font):
        self.font = font
        self.canvas.itemconfig(self.mytext, font=self.font)
    def change_text(self, text):
        self.text = text
        self.canvas.itemconfig(self.mytext, text=self.text)
    def set_size(self,w,h):
        self.new_x = (self.x/self.winx)*w
        self.new_y = (self.y/self.winy)*h
       
        if(500<=w<=600):
            self.font.config(size=7)
            if(self.isbg == True):
                xo = (self.bgx * 0.6)
                yo = (self.bgy * 0.47)
                __xo = self.bg_xo * 0.2
                __yo = self.bg_yo *0.2

                self.canvas.delete(self.bg)
                self.bg = round_rectangle(self.canvas,x1=self.new_x, y1=self.new_y,
                                          x2=self.new_x+xo, y2=self.new_y+yo,
                                          fill=self.bg_color, radius=8)
                self.canvas.moveto(self.bg, self.new_x -__xo,self.new_y - __yo)
                self.canvas.tag_raise(self.mytext)
                self.newxo, self.newyo, self.new__xo, self.new__yo = xo, yo, __xo, __yo
        elif(600<w<=700):
            self.font.config(size=8)
            if(self.isbg == True):
                xo = (self.bgx*0.77)
                yo = (self.bgy*0.68)
                __xo = self.bg_xo * 0.5
                __yo = self.bg_yo *0.5
                self.canvas.delete(self.bg)
                self.bg = round_rectangle(self.canvas,x1=self.new_x,y1=self.new_y,
                                          x2=self.new_x + xo, y2=self.new_y+yo,
                                          fill=self.bg_color, radius=10)
                self.canvas.moveto(self.bg, self.new_x - __xo,self.new_y -__yo)
                self.canvas.tag_raise(self.mytext)
                self.newxo, self.newyo, self.new__xo, self.new__yo = xo, yo, __xo, __yo
        else:
            self.font.config(size=10)
            if(self.isbg == True):
                self.canvas.delete(self.bg)
                if w > self.winx:
                    x2 = self.new_x + self.bgx
                else:
                    x2 = (self.new_x + self.bgx)+(w-self.winx)/20
                self.bg = round_rectangle(self.canvas,x1=self.new_x,y1=self.new_y,
                                          x2=x2, y2=self.new_y+self.bgy,
                                          fill=self.bg_color, radius=15) 
                self.canvas.moveto(self.bg, self.new_x -self.bg_xo,self.new_y -self.bg_yo)
                self.canvas.tag_raise(self.mytext)
                self.newxo, self.newyo, self.new__xo, self.new__yo = self.bgx, self.bgy, self.bg_xo, self.bg_yo
        self.change_font(self.font)
    def set_bg_color(self, color:str):
        # self.canvas.itemconfig(self.bg, fill= 'red')
        if(self.isbg == True):
            self.canvas.delete(self.bg)
            self.bg = round_rectangle(self.canvas,x1=self.new_x, y1=self.new_y, x2=self.new_x + self.newxo, y2=self.new_y + self.newyo,fill =color, radius=10) 
            self.canvas.moveto(self.bg, self.new_x -self.new__xo,self.new_y -self.new__yo)
            self.canvas.tag_raise(self.mytext)
            self.bg_color = color
    # def moveto(self,w,h):
    #     self.canvas.moveto(self.mytext,(self.new_x/1000)*w,(self.new_y/1000)*h)

class Tkbutton():
    def __init__(self,
                 canvas: tk.Canvas,
                 x,
                 y,
                 w = 300,
                 h=40,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 anchor = 'nw',
                 text = "Button",
                #  font = tkfont(family="System",	size=12,weight="normal",slant="italic",underline=1,overstrike=0),
                 fg = 'black',
                 bg = 'green'
                ) -> None:
        self.canvas = canvas
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.text = text
        self.anchor = anchor
        # self.font = font
        self.fg = fg
        self.bg = bg
        self.button = tk.Button(self.canvas, text=self.text)
        # self.button['font'] = tkfont(family='Impact')
        self.button.config(font= tkfont(family='Impact'))
        self.place = self.canvas.create_window(self.x, self.y,anchor=self.anchor,window=self.button, width =self.w, height=self.h)
    def moveto(self,w,h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            self.new_w = w
        if(h!=0):
            self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.new_h = h 
    def state(self, state: str):
        self.button.config(state=state)
    def set_text(self, text:str):
        self.button.config(text=text)

class Tkprogressbar():
    def __init__(self,
                 canvas : tk.Canvas,
                 x : int,
                 y : int,
                 w = 800,
                 h = 20,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 anchor = 'nw',
                 orient = 'horizontal',
                 len = 100,
                 mode = 'determinate',
                 ) -> None:
        self.canvas = canvas
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.anchor = anchor
        self.orient = orient
        self.len = len
        self.mode = mode
        self.progressbar = ttk.Progressbar(self.canvas, orient=self.orient, length= self.len,mode=self.mode)
        self.place = self.canvas.create_window(self.x, self.y, anchor=self.anchor,window=self.progressbar, width=self.w, height=self.h)
    def moveto(self,w,h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            self.new_w = w
        if(h!=0):
            self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.new_h = h 
    def update(self, value):
        self.progressbar['value'] = value

class TKcheckbtn():
    def __init__(self,
                 canvas : tk.Canvas,
                 win,
                 x : int,
                 y : int,
                 w = 140,
                 h = 30,
                 win_defaultx = 1000,
                 win_defaulty = 1000,
                 text = 'text',
                 anchor = 'nw',) -> None:
        self.s = ttk.Style()
        self.s.configure('TCheckbutton', foreground='maroon')
        self.canvas = canvas
        self.win = win
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.winx = win_defaultx
        self.winy = win_defaulty
        self.anchor = anchor
        self.var = tk.IntVar()
        self.text = text
        
        # self.checkbtn = Checkbutton(self.win,text=self.text,variable=self.var,onvalue=1, offvalue=0,width=20,height=20)
        self.checkbtn = ttk.Checkbutton(self.canvas, text=self.text, style='Roundtoggle.Toolbutton')
       
        # self.checkbtn['font'] = tkfont(family='Impact')
        self.place = self.canvas.create_window(self.x,self.y,window=self.checkbtn,anchor=self.anchor)
        self.checkbtn.config(variable=self.var, onvalue=1, offvalue=0)
    def moveto(self,w,h):
        self.canvas.moveto(self.place,(self.x/self.winx)*w,(self.y/self.winy)*h)
    def change_width_height(self,w = 0,h = 0):
        if(w != 0):
            self.canvas.itemconfig(self.place,width=(self.w/self.winx)*w)
            self.new_w = w
        if(h!=0):
            self.canvas.itemconfig(self.place,height=(self.h/self.winy)*h)
            self.new_h = h 
    def get_state(self):        
        return self.var.get()