
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
import os

from Functions import Funciones
from Singlenton import Singleton


#pantalla de carga
class MainView:        
     #interfaz
    def __init__(self):
        self.raiz=TkinterDnD.Tk()
        self.raiz.title("Medidor CVM")
        self.raiz.iconbitmap("recursos/logo.ico")

        #DEFINICION DE VARIABLES INICIALIZADAS  
        funciones_instance = Funciones(self,any,any)   
        self.icono_check= ImageTk.PhotoImage(Image.open('recursos/check.png').resize((20, 20)))
        singleton_instance = Singleton()
        

        self.mycanvas = Canvas(self.raiz,width=920,height=15)
        self.mycanvas.pack(side=LEFT,expand=1)

        self.yscrollbar=ttk.Scrollbar(self.raiz,orient="vertical", command=self.mycanvas.yview)
        self.yscrollbar.pack(side=RIGHT,fill="y",expand=1)

        self.mycanvas.configure(yscrollcommand=self.yscrollbar.set)

        self.mycanvas.bind('<Configure>',lambda e: self.mycanvas.configure(scrollregion = self.mycanvas.bbox('all')))

        #############Empieza la interfaz##################

        miframe= Frame(self.mycanvas,width=120,height=800)
        miframe.pack(side=TOP)
        miframe.config(bg="#ecefee")

        
        #-----------------FRAME TITULO------------------#
        titulo_frame= Frame(miframe,width=1200,height=100)
        titulo_frame.pack()
        titulo_frame.config(bg="#ecefee")

        titulo_label=Label(titulo_frame, text="MEDIDOR CVM",font=("Arial Black", 22))
        titulo_label.grid(row=0, column=0, pady=2,padx=5)
        titulo_label.config(bg="#ecefee")

        #-----------------FRAME PRINCIPAL-----------------#
        principal_frame= Frame(miframe,bd=1,relief="solid")
        principal_frame.pack()

        #-------------FRAME DATOS  GENERALES--------------#
        cabecera_frame= Frame(principal_frame,bd=1,relief="solid")
        cabecera_frame.grid(row=0,column=0,sticky='w')
        cabecera_frame.config(bg="#ecefee",pady=0,padx=0)
    
        titulo2_label=Label(cabecera_frame, text="DATOS  GENERALES",font=("Arial Black", 12))
        titulo2_label.grid(row=0, column=0, pady=5,padx=360,columnspan=6)

        #FRAME 1
        izquierda_cabecera_frame= Frame(cabecera_frame)
        izquierda_cabecera_frame.grid(row=1,column=0,padx=25)

        cliente_label=Label(izquierda_cabecera_frame, text="Cliente:",font=("Arial", 9))
        cliente_label.grid(row=0, column=0, sticky="e", pady=5, padx=0)
        cliente_cuadro=Entry(izquierda_cabecera_frame,textvariable=singleton_instance.get_variable('vista_cliente'),font=("Arial", 9))
        cliente_cuadro.grid(row=0, column=1, sticky="w", pady=5,padx=10)

        area_label=Label(izquierda_cabecera_frame, text="Area:",font=("Arial", 9))
        area_label.grid(row=1, column=0, sticky="e", pady=5, padx=0)
        area_cuadro=Entry(izquierda_cabecera_frame,textvariable=singleton_instance.get_variable('vista_area'),font=("Arial", 9))
        area_cuadro.grid(row=1, column=1, sticky="w", pady=5,padx=10)

        #FRAME 2
        centro_cabecera_frame= Frame(cabecera_frame)
        centro_cabecera_frame.grid(row=1,column=1,padx=25)

        fechalectura_label=Label(centro_cabecera_frame, text="Fecha de Lectura:",font=("Arial", 9))
        fechalectura_label.grid(row=1, column=2, sticky="e", pady=5,padx=0)
        fechalectura_cuadro=DateEntry(centro_cabecera_frame,textvariable=singleton_instance.get_variable('vista_lectura'),date_pattern="dd-mm-yyyy",locale="es",font=("Arial", 9))
        fechalectura_cuadro.grid(row=1, column=3, sticky="w", pady=5,padx=10)

        fechaemision_label=Label(centro_cabecera_frame, text="Fecha de Emision:",font=("Arial", 9))
        fechaemision_label.grid(row=2, column=2, sticky="e", pady=5,padx=0)
        fechaemision_cuadro=DateEntry(centro_cabecera_frame,textvariable=singleton_instance.get_variable('vista_emision'),date_pattern="dd-mm-yyyy",locale="es",font=("Arial", 9))
        fechaemision_cuadro.grid(row=2, column=3, sticky="w", pady=5,padx=10)

        #FRAME 3
        derecha_cabecera_frame= Frame(cabecera_frame)
        derecha_cabecera_frame.grid(row=1,column=2,padx=25)

        informetecnico_label=Label(derecha_cabecera_frame, text="Informe Técnico:",font=("Arial", 9))
        informetecnico_label.grid(row=1, column=4, pady=5,padx=0,sticky="e")
        informetecnico_cuadro=Entry(derecha_cabecera_frame,textvariable=singleton_instance.get_variable('vista_informe'),font=("Arial", 9))
        informetecnico_cuadro.grid(row=1, column=5, pady=5,padx=10,sticky="w")

        dias_label=Label(derecha_cabecera_frame, text="Dias del mes:",font=("Arial", 9))
        dias_label.grid(row=2, column=4, sticky="e", pady=5, padx=0)
        dias_cuadro=Entry(derecha_cabecera_frame,textvariable=singleton_instance.get_variable('vista_dias'),width=8,font=("Arial", 9))
        dias_cuadro.grid(row=2, column=5, sticky="w", pady=5,padx=10)

        #---------------FRAME MAXIMA DEMANDA---------------#
        maximademanda_frame= Frame(principal_frame,bd=1,relief="solid")
        maximademanda_frame.grid(row=1,column=0)
        maximademanda_frame.config(bg="#ecefee",padx=100,pady=4)

        titulo3_label=Label(maximademanda_frame, text="MAXIMA DEMANDA ",font=("Arial Black", 12))
        titulo3_label.grid(row=0, column=0, pady=5,padx=265,columnspan=6)

        #FRAME 1
        select_maximademanda_frame= Frame(maximademanda_frame,relief="solid")
        select_maximademanda_frame.grid(row=1,column=0,columnspan=6)
        select_maximademanda_frame.config(bg="#ecefee",padx=50,pady=5)

        boton_check = Button(select_maximademanda_frame,image=self.icono_check, command=(funciones_instance.change_mes))
        boton_check.grid(row=0,column=2)

        #FRAME 2
        izquierda_maximademanda_frame= Frame(maximademanda_frame)
        izquierda_maximademanda_frame.grid(row=2,column=0,padx=25)

        mes1_entry=Entry(izquierda_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_nombreMes1') ,width=11)
        mes1_entry.configure(state='disabled')
        mes1_entry.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        mes1_cuadro=Entry(izquierda_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_mes1'),width=8)
        mes1_cuadro.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(izquierda_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=0, column=2, sticky="w")
        mes4_entry=Entry(izquierda_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_nombreMes4') ,width=11)
        mes4_entry.configure(state='disabled')
        mes4_entry.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        mes4_cuadro=Entry(izquierda_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_mes4'),width=8)
        mes4_cuadro.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(izquierda_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=1, column=2, sticky="w")
        
        #FRAME 3
        centro_maximademanda_frame= Frame(maximademanda_frame)
        centro_maximademanda_frame.grid(row=2,column=1,padx=25)

        mes2_entry=Entry(centro_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_nombreMes2') ,width=11)
        mes2_entry.configure(state='disabled')
        mes2_entry.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        mes2_cuadro=Entry(centro_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_mes2'),width=8)
        mes2_cuadro.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(centro_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=0, column=2, sticky="w")
        mes5_entry=Entry(centro_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_nombreMes5') ,width=11)
        mes5_entry.configure(state='disabled')
        mes5_entry.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        mes5_cuadro=Entry(centro_maximademanda_frame ,textvariable=singleton_instance.get_variable('vista_mes5'),width=8)
        mes5_cuadro.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(centro_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=1, column=2, sticky="w")

        #FRAME 4
        derecha_maximademanda_frame= Frame(maximademanda_frame)
        derecha_maximademanda_frame.grid(row=2,column=2,padx=25)

        mes3_entry=Entry(derecha_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_nombreMes3'),width=11)
        mes3_entry.configure(state='disabled')
        mes3_entry.grid(row=0, column=0, sticky="e", pady=5,padx=10)
        mes3_cuadro=Entry(derecha_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_mes3'),width=8)
        mes3_cuadro.grid(row=0, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(derecha_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=0, column=2, sticky="w")
        entrymes6=Entry(derecha_maximademanda_frame ,textvariable=singleton_instance.get_variable('vista_nombreMes6'),width=11)
        entrymes6.configure(state='disabled')
        entrymes6.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        mes6_cuadro=Entry(derecha_maximademanda_frame,textvariable=singleton_instance.get_variable('vista_mes6'),width=8)
        mes6_cuadro.configure(state='disabled')
        mes6_cuadro.grid(row=1, column=1, sticky="w", pady=5,padx=0)
        kw_label=Label(derecha_maximademanda_frame, text="(kW)",font=("Arial", 9))
        kw_label.grid(row=1, column=2, sticky="w")

        #---------------FRAME TARIFARIO---------------#    

        frame_tarifario= Frame(miframe,width=1200,height=151,bd=2,relief="solid")
        frame_tarifario.pack()
        frame_tarifario.config(bg="#ecefee", padx=5)

        titulo4_label=Label(frame_tarifario, text="PLIEGO TARIFARIO",font=("Arial Black", 12))
        titulo4_label.grid(row=0, column=0, columnspan=8,pady=5,padx=0)

        tarifa1_label=Label(frame_tarifario, text="C.Fijo Mensual\n(S/./mes):",font=("Arial", 9))
        tarifa1_label.grid(row=2, column=0, sticky="e", pady=5)
        tarifa1_cuadro=Entry(frame_tarifario,textvariable=singleton_instance.get_variable('vista_c_fijoMensual'),width=10)
        tarifa1_cuadro.grid(row=2, column=1, sticky="w", pady=5,padx=12)

        tarifa2_label=Label(frame_tarifario, text="C.Energía Activa Punta\n(Ctm.S/./kW.h):",font=("Arial", 9))
        tarifa2_label.grid(row=2, column=2, sticky="e", pady=5)
        tarifa2_cuadro=Entry(frame_tarifario,textvariable=singleton_instance.get_variable('vista_c_energiaActivaPunta'),width=10)
        tarifa2_cuadro.grid(row=2, column=3, sticky="w", pady=5,padx=12)

        tarifa3_label=Label(frame_tarifario, text="C.Energía Activa Fuera\nPunta (Ctm.S/./kW.h):",font=("Arial", 9))
        tarifa3_label.grid(row=4, column=0, sticky="e", pady=5,padx=5)
        tarifa3_cuadro=Entry(frame_tarifario,textvariable=singleton_instance.get_variable('vista_c_energiaActivaFueraPunta') ,width=10)
        tarifa3_cuadro.grid(row=4, column=1, sticky="w", pady=5,padx=12)

        subtitulo_label=Label(frame_tarifario, text="Cargo por Potencia Activa de generación para Usuarios(S/./kW-mes)",font=("Arial", 9))
        subtitulo_label.grid(row=1, column=4, pady=5,padx=5,columnspan=4)
        subtitulo_label.config(bg="white")

        tarifa4_label=Label(frame_tarifario, text="Presente en Punta:",font=("Arial", 9))
        tarifa4_label.grid(row=2, column=4, sticky="e", pady=5)
        tarifa4_cuadro=Entry(frame_tarifario, textvariable=singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta') ,width=10)
        tarifa4_cuadro.grid(row=2, column=5, sticky="w", pady=5,padx=5)

        tarifa5_label=Label(frame_tarifario, text="Presente Fuera de Punta:",font=("Arial", 9))
        tarifa5_label.grid(row=2, column=6, sticky="e", pady=5)
        tarifa5_cuadro=Entry(frame_tarifario,textvariable=singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta') ,width=10)
        tarifa5_cuadro.grid(row=2, column=7, sticky="w", pady=5,padx=5)

        subtitulo_label=Label(frame_tarifario, text="Cargo por Potencia Activa de redes de distribución para Usuarios(S/./kW-mes)",font=("Arial", 9))
        subtitulo_label.grid(row=3, column=4, pady=5,padx=5,columnspan=4)
        subtitulo_label.config(bg="white")

        tarifa6_label=Label(frame_tarifario, text="Presente en Punta:",font=("Arial", 9))
        tarifa6_label.grid(row=4, column=4, sticky="e", pady=5)
        tarifa6_cuadro=Entry(frame_tarifario, textvariable=singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta') ,width=10)
        tarifa6_cuadro.grid(row=4, column=5, sticky="w", pady=5,padx=5)

        tarifa7_label=Label(frame_tarifario, text="Presente Fuera de Punta:",font=("Arial", 9))
        tarifa7_label.grid(row=4, column=6, sticky="e", pady=5)
        tarifa7_cuadro=Entry(frame_tarifario, textvariable=singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta') ,width=10)
        tarifa7_cuadro.grid(row=4, column=7, sticky="w", pady=5,padx=5)

        tarifa8_label=Label(frame_tarifario, text="C.Energía Reactiva exc.\n30% (Ctm.S/./kVar.h):",font=("Arial", 9))
        tarifa8_label.grid(row=4, column=2, sticky="e", pady=5)
        tarifa8_cuadro=Entry(frame_tarifario,textvariable=singleton_instance.get_variable('vista_c_energiaReactivaExc30') ,width=10)
        tarifa8_cuadro.grid(row=4, column=3, sticky="w", pady=5,padx=12)

        #---------------FRAME MEDICION---------------#  
    
        medicion_frame= Frame(miframe,width=1200,height=400,bd=1,relief="solid")
        medicion_frame.pack()
        medicion_frame.config(bg="#ecefee")

        #FRAME NOMBRE TABLERO
        tablerosss= Frame(medicion_frame,bd=1,relief="solid")
        tablerosss.config(bg="#ecefee")
        tablerosss.grid(row=0, column=0 ,columnspan=4,padx=0)

        titulo5label=Label(tablerosss, text="TABLEROS",font=("Arial Black", 12))
        titulo5label.grid(row=0, column=0, pady=5,padx=402 ,columnspan=2)

        titulo84label=Label(tablerosss, text="Nombre del tablero:",font=("Arial", 9))
        titulo84label.grid(row=1, column=0, sticky="e", pady=5,padx=10)
        cuadrota790=Entry(tablerosss,textvariable=singleton_instance.get_variable('vista_t_nombreTablero'))
        cuadrota790.grid(row=1, column=1, sticky="w", pady=5,padx=5)

        #FRAME SECTION A
        section_a_frame= Frame(medicion_frame,height=150,bd=1,relief="solid")
        section_a_frame.config(bg="#ecefee",padx=15,pady=5)
        section_a_frame.grid(row=1, column=0)

        titulo_label=Label(section_a_frame, text="Energía activa en hora fuera \nde punta, T1 (M/KWh)",font=("Arial", 9))
        titulo_label.grid(row=0, column=0, pady=5,padx=10, columnspan=2)
        titulo_label.config(bg="white")

        subtitulo_label=Label(section_a_frame, text="Mes actual:",font=("Arial", 9))
        subtitulo_label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadro1_entry=Entry(section_a_frame, textvariable=singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual'),width=14)
        cuadro1_entry.grid(row=1, column=1, sticky="w", pady=5,padx=2)

        subtitulo_label=Label(section_a_frame, text="Mes anterior:",font=("Arial", 9))
        subtitulo_label.grid(row=2, column=0, sticky="e", pady=5,padx=2)
        cuadro2_entry=Entry(section_a_frame,textvariable=singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior'),width=14)
        cuadro2_entry.grid(row=2, column=1, sticky="w", pady=5,padx=2)

        titulo_label=Label(section_a_frame, text="Energía activa en hora punta,\n T2 (M/KWh)",font=("Arial", 9))
        titulo_label.grid(row=0, column=2, pady=5,padx=10,columnspan=2)
        titulo_label.config(bg="white")

        subtitulo_label=Label(section_a_frame, text="Mes actual:",font=("Arial", 9))
        subtitulo_label.grid(row=1, column=2, sticky="e", pady=5,padx=2)
        cuadro3_entry=Entry(section_a_frame,textvariable=singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual'),width=14)
        cuadro3_entry.grid(row=1, column=3, sticky="w", pady=5,padx=2)

        subtitulo_label=Label(section_a_frame, text="Mes anterior:",font=("Arial", 9))
        subtitulo_label.grid(row=2, column=2, sticky="e", pady=5,padx=2)
        cuadro4_entry=Entry(section_a_frame,textvariable=singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior'),width=14)
        cuadro4_entry.grid(row=2, column=3, sticky="w", pady=5,padx=2)

        boton1_label = Label(section_a_frame, text="Evidencia",font=("Arial",9))
        boton1_label.grid(row=3,column=0,columnspan=4, pady=5)
        boton1_label.config(bg="white")

        self.evidencia1_canvas = tk.Canvas(section_a_frame, width=150, height=100, bg="white")
        self.evidencia1_canvas.grid(row=4, column=0, columnspan=4, pady=5)

        funciones_instance = Funciones(self,self.evidencia1_canvas,1)

        #Configuración para permitir el arrastre y la soltura
        self.evidencia1_canvas.drop_target_register(DND_FILES)
        self.evidencia1_canvas.dnd_bind('<<Drop>>', funciones_instance.load_image)
        
        #FRAME SECTION B
        section_b_frame= Frame(medicion_frame,height=150,bd=1,relief="solid")
        section_b_frame.config(bg="#ecefee",pady=12)
        section_b_frame.grid(row=1, column=1,rowspan=7)

        titulo_label=Label(section_b_frame, text="Máxima demanda (kW)",font=("Arial", 9))
        titulo_label.grid(row=0, column=0, sticky="n", pady=5,padx=5, columnspan=2)
        titulo_label.config(bg="white")

        subtitulo_label=Label(section_b_frame, text="Max:",font=("Arial", 9))
        subtitulo_label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadro5_entry=Entry(section_b_frame,textvariable=singleton_instance.get_variable('vista_t_maximaDemanda'),width=14)
        cuadro5_entry.grid(row=1, column=1, sticky="w", pady=5,padx=2)

        boton2_label = tk.Label(section_b_frame, text="Evidencia",font=("Arial",9))
        boton2_label.grid(row=2,column=0,columnspan=2, pady=(37, 5))
        boton2_label.config(bg="white")

        self.evidencia2_canvas = tk.Canvas(section_b_frame, width=150, height=100, bg="white")
        self.evidencia2_canvas.grid(row=3, column=0, columnspan=2, pady=5, padx=10)

        funciones_instance = Funciones(self,self.evidencia2_canvas,2)

        #Configuración para permitir el arrastre y la soltura
        self.evidencia2_canvas.drop_target_register(DND_FILES)
        self.evidencia2_canvas.dnd_bind('<<Drop>>', funciones_instance.load_image)

        #FRAME SECTION C
        section_c_frame= Frame(medicion_frame,height=150,bd=1,relief="solid")
        section_c_frame.config(bg="#ecefee",pady=5, padx=17)
        section_c_frame.grid(row=1, column=2,rowspan=7)

        titulo_label=Label(section_c_frame, text="Energía reactiva inductiva \ntotal (M/KvarLh)",font=("Arial", 9))
        titulo_label.grid(row=0, column=0, pady=5,padx=10, columnspan=2)
        titulo_label.config(bg="white")

        subtitulo_label=Label(section_c_frame, text="Mes actual:",font=("Arial",9))
        subtitulo_label.grid(row=1, column=0, sticky="e", pady=5,padx=2)
        cuadro6_entry=Entry(section_c_frame,textvariable=singleton_instance.get_variable('vista_t_energiaReactivaInductivaActual'),width=14)
        cuadro6_entry.grid(row=1, column=1, sticky="w", pady=5,padx=2)

        subtitulo_label=Label(section_c_frame, text="Mes anterior:",font=("Arial", 9))
        subtitulo_label.grid(row=2, column=0, sticky="e", pady=5,padx=2)
        cuadro7_entry=Entry(section_c_frame, textvariable=singleton_instance.get_variable('vista_t_energiaReactivaInductivaAnterior'),width=14)
        cuadro7_entry.grid(row=2, column=1, sticky="w", pady=5,padx=2)
            
        boton3_label = tk.Label(section_c_frame, text="Evidencia",font=("Arial",9))
        boton3_label.grid(row=3,column=0,columnspan=2,pady=5)
        boton3_label.config(bg="white")

        self.evidencia3_canvas = tk.Canvas(section_c_frame, width=150, height=100, bg="white")
        self.evidencia3_canvas.grid(row=4, column=0, columnspan=2, pady=5,padx=10)

        funciones_instance = Funciones(self,self.evidencia3_canvas,3)

        #Configuración para permitir el arrastre y la soltura
        self.evidencia3_canvas.drop_target_register(DND_FILES)
        self.evidencia3_canvas.dnd_bind('<<Drop>>', funciones_instance.load_image)

        #FRAME SECTION D
        section_d_frame= Frame(medicion_frame,height=150,bd=1,relief="solid")
        section_d_frame.config(bg="#ecefee",pady=46,padx=9)
        section_d_frame.grid(row=1, column=3, sticky="e",rowspan=7)

        titulo_label=Label(section_d_frame, text="Ingresados",font=("Arial", 12))
        titulo_label.grid(row=0, column=0, pady=2,padx=10)
        cuadro8_entry=Entry(section_d_frame, textvariable=any,font=("Arial", 50),width=2)
        cuadro8_entry.configure(state='disabled',textvariable=singleton_instance.get_variable('vista_cantidadMedidores'),justify='center')
        cuadro8_entry.grid(row=2, column=0, pady=9,padx=10)
            
        self.boton4_label=Button(section_d_frame,text="Agregar", command= (funciones_instance.add_medidor),font=("Arial", 10), relief="raised", borderwidth=4)
        self.boton4_label.grid(row=3, sticky="s",pady=7)

        #---------------FRAME MEDICION---------------#  
    
        medicion_frame= Frame(miframe,width=1200,height=400,bd=1,relief="solid")
        medicion_frame.pack()
        medicion_frame.config(bg="#ecefee")

        #FRAME NOMBRE TABLERO AGUA
        tablerosss= Frame(medicion_frame,bd=1,relief="solid")
        tablerosss.config(bg="#ecefee")
        tablerosss.grid(row=0, column=0 ,columnspan=4,padx=0)

        titulo5label=Label(tablerosss, text="MEDICIÓN DE AGUA",font=("Arial Black", 12))
        titulo5label.grid(row=0, column=0, pady=5, padx=(363, 364) ,columnspan=2)

        #FRAME AGUA
        section_c_frame= Frame(medicion_frame,height=170,bd=1,relief="solid")
        section_c_frame.config(bg="#ecefee")
        section_c_frame.grid(row=1, column=2,rowspan=7,pady=0)

        titulo_label=Label(section_c_frame, text="Consumo de agua (M3)",font=("Arial", 9))
        titulo_label.grid(row=1, column=2, pady=(12, 6),padx=15, columnspan=2)
        titulo_label.config(bg="white")

        subtitulo_label=Label(section_c_frame, text="Mes actual:",font=("Arial",9))
        subtitulo_label.grid(row=2, column=1, sticky="e", pady=5,padx=2)
        cuadro6_entry=Entry(section_c_frame,textvariable=singleton_instance.get_variable('vista_a_Actual'),width=14)
        cuadro6_entry.grid(row=2, column=2, sticky="w", pady=5,padx=(2, 30))

        subtitulo_label=Label(section_c_frame, text="Mes anterior:",font=("Arial", 9))
        subtitulo_label.grid(row=2, column=3, sticky="e", pady=5,padx=5)
        cuadro7_entry=Entry(section_c_frame, textvariable=singleton_instance.get_variable('vista_a_Anterior'),width=14)
        cuadro7_entry.grid(row=2, column=4, sticky="w", pady=5, padx=(2, 60))
            
        boton3_label = tk.Label(section_c_frame, text="Evidencia",font=("Arial",9))
        boton3_label.grid(row=1,column=0,pady=7,padx=(65, 80))
        boton3_label.config(bg="white")

        self.evidencia4_canvas = tk.Canvas(section_c_frame, width=150, height=100, bg="white")
        self.evidencia4_canvas.grid(row=2, column=0,  pady=7,padx=(80, 80))

        funciones_instance = Funciones(self,self.evidencia4_canvas,4)

        #Configuración para permitir el arrastre y la soltura
        self.evidencia4_canvas.drop_target_register(DND_FILES)
        self.evidencia4_canvas.dnd_bind('<<Drop>>', funciones_instance.load_image)

        #FRAME SECTION D
        section_d_frame= Frame(medicion_frame,height=150,bd=1,relief="solid")
        section_d_frame.config(bg="#ecefee",pady=4,padx=29)
        section_d_frame.grid(row=1, column=3, sticky="e",rowspan=7)

        titulo_label=Label(section_d_frame, text="Ingresados",font=("Arial", 12))
        titulo_label.grid(row=0, column=0, pady=2,padx=10)
        cuadro8_entry=Entry(section_d_frame, textvariable=any,font=("Arial", 35),width=2)
        cuadro8_entry.configure(state='disabled',textvariable=singleton_instance.get_variable('vista_a_cantidadMedidores'),justify='center')
        cuadro8_entry.grid(row=2, column=0, pady=9,padx=10)
            
        self.boton4_label=Button(section_d_frame,text="Agregar", command= (funciones_instance.add_a_medidor),font=("Arial", 10), relief="raised", borderwidth=4)
        self.boton4_label.grid(row=3, sticky="s",pady=7)


        # Opción de Firma
        firma_frame = Frame(miframe, bd=1, relief="solid")
        firma_frame.pack()
        firma_frame.config(bg="#ecefee", pady=5, padx=364)

        incluir_firma_check = ttk.Checkbutton(firma_frame, text="Incluir firma en el documento", variable=singleton_instance.get_variable('firma'))
        incluir_firma_check.grid(row=0, column=0, sticky="W", pady=5)

        #---------------FRAME FINAL---------------#  
        framefinal= Frame(miframe,width=1200,height=50)
        framefinal.pack( pady=(5, 10))
        framefinal.config(bg="#ecefee")

        btnfinal1=Button(framefinal,text="Guardar",command= (funciones_instance.save),  font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        btnfinal1.grid(row=0,column=1,padx=30,pady=2)
        btnfinal1.config(bg="#ecefee")

        self.btnfinal2=Button(framefinal,text="Exportar", command= (funciones_instance.exportar),font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        self.btnfinal2.grid(row=0,column=2,padx=30,pady=2)
        self.btnfinal2.config(bg="#ecefee")

        btnfinal3=Button(framefinal,text="Limpiar", command= (funciones_instance.reset),font=("Arial", 12),cursor="hand2", relief="raised", borderwidth=4)
        btnfinal3.grid(row=0,column=0,padx=30, pady=2)
        btnfinal3.config(bg="#ecefee")

        #MENSAJES EN PANTALLA
        # self.estadolabel=Label(miframe, text="Estado: ",font=("Arial", 9)).place(x=10,y=840)
        estadolabel=Label(miframe, text="Derechos Reservados INSTCAL SAC       v1.4",font=("Arial", 8)).place(x=680,y=1440)

        funciones_instance.reset()

        self.mycanvas.create_window((0,0), window=miframe, anchor="nw")

        self.mycanvas.pack(fill="both", expand="yes", padx=0, pady=0)

        self.raiz.geometry("945x880")
        self.raiz.resizable(False,True)

        self.raiz.mainloop()
    