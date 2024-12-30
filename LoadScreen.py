
#actualizacion 29/10/2022
from cProfile import label
from cgitb import text
from faulthandler import disable
from mimetypes import init
from multiprocessing.sharedctypes import Value
from pydoc import doc
from struct import pack
from tkinter import LEFT, RIGHT, SOLID, TOP, Button, Canvas, DoubleVar, Entry, Frame, Image, Label, LabelFrame, Listbox, OptionMenu, Scrollbar, StringVar, Tk, Toplevel, ttk
from tokenize import Double, String
from turtle import st, width
from docxtpl import DocxTemplate, InlineImage
from tkinter import filedialog 
from PIL import Image, ImageTk
from docx.shared import Mm
from tkinter import messagebox
from tkcalendar import *
from time import sleep
from tkinterdnd2 import DND_FILES,TkinterDnD
import tkinter as tk
import xlsxwriter

#pantalla de carga
class LoadingSplash:
    def __init__(self):
        # setting root window:
        self.secc = Tk()
        self.secc.config(bg="white")
        self.secc.title("Medidor CVM")
        self.secc.iconbitmap("recursos/logo.ico")
        self.secc.geometry("850x500")
        #self.root.attributes("-fullscreen",True)
        imagen_carga= ImageTk.PhotoImage(Image.open('recursos/pan_inicio.png').resize((800, 400)))
        Label(self.secc,image= imagen_carga).place(x=20,y=20)
        # loading text:
        Label(self.secc, text="Loading...", font="Bahnschrift 15",
            bg="white", fg="black").place(x=250, y=430)
        
        # loading blocks:
        for i in range(16):
            Label(self.secc, bg="#1F2732", width=2, height=1).place(x=(i+12)*22, y=460)
        
        # update root to see animation:
        self.secc.update()
        self.play_animation()
        # window in mainloop:
        #self.secc.mainloop()
    # loader animation:
    def play_animation(self):
        for i in range(4):
            for j in range(16):
                # make block yellow:
                Label(self.secc, bg="#FFBD09", width=2, height=1).place(x=(j+12)*22, y=460)
                sleep (0.06)
                self.secc.update_idletasks()
                # make block dark:
                Label(self.secc, bg="#1F2732", width=2, height=1).place(x=(j+12)*22, y=460)
        else:
            self.secc.destroy()
            #exit(0)