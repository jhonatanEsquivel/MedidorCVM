
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

from LoadScreen import LoadingSplash
from MainScreen import MainView

# import MainView from MainScreen

#pantalla de carga
class AppManager:
    def __init__(self):
        for i in range(2):
            print("iterado :",i)
            if i==0 :
                LoadingSplash()
                print("cargando")
            if i==1:

                MainView()

if __name__ == "__main__":
    app_manager = AppManager()

