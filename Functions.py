
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

from Singlenton import Singleton
#pantalla de carga
class Funciones:
    def __init__(self, app_instance, cuadrota500, number):
        self.app_instance = app_instance
        self.cuadrota500 = cuadrota500
        self.number = number
    def load_image(self, event):
        self.image_path = tk.StringVar()
        
        # Obtener la ruta de la imagen desde el evento de soltar
        image_path = event.data

        # Deshacerse de las llaves alrededor de la ruta (si las hay)
        image_path = image_path.strip('{}')

        # Actualizar la variable de la ruta de la imagen
        self.image_path.set(image_path)

        # Imprimir la ruta de la imagen en la consola
        print(f"Ruta de la imagen: {image_path}")

        # Cargar y mostrar la imagen en el lienzo
        self.show_image(image_path)

        singleton_instance = Singleton()
        if(self.number==1):
            singleton_instance.set_variable('vista_t_evidencia1',image_path)
        elif(self.number==2):
            singleton_instance.set_variable('vista_t_evidencia2',image_path)
        elif(self.number==3):
            singleton_instance.set_variable('vista_t_evidencia3',image_path)
        elif(self.number==4):
            singleton_instance.set_variable('vista_a_evidencia1',image_path)

    def show_image(self, image_path):
        try:
            # Abrir la imagen usando PIL
            image = Image.open(image_path)

            # Redimensionar la imagen a las dimensiones deseadas
            new_width = 150  # Ancho deseado
            new_height = 100  # Alto deseado
            image = image.resize((new_width, new_height), Image.LANCZOS)

            # Convertir la imagen a formato compatible con Tkinter
            tk_image = ImageTk.PhotoImage(image)

            # Mostrar la imagen en el lienzo
            self.cuadrota500.config(width=new_width, height=new_height)
            self.cuadrota500.create_image(0, 0, anchor='nw', image=tk_image)
            self.cuadrota500.image = tk_image

        except Exception as e:
            # Manejar posibles errores al abrir la imagen
            print(f"Error al abrir la imagen: {e}")
    
    def reset(self):
        singleton_instance = Singleton()
        
        singleton_instance.set_variable('vista_cliente','')
        singleton_instance.set_variable('vista_area','')
        singleton_instance.set_variable('vista_informe','')
        singleton_instance.set_variable('vista_dias','0')
        singleton_instance.set_variable('vista_mes1','0')
        singleton_instance.set_variable('vista_mes2','0')
        singleton_instance.set_variable('vista_mes3','0')
        singleton_instance.set_variable('vista_mes4','0')
        singleton_instance.set_variable('vista_mes5','0')
        singleton_instance.set_variable('vista_mes6','0')
        singleton_instance.set_variable('vista_nombreMes1','')
        singleton_instance.set_variable('vista_nombreMes2','')
        singleton_instance.set_variable('vista_nombreMes3','')
        singleton_instance.set_variable('vista_nombreMes4','')
        singleton_instance.set_variable('vista_nombreMes5','')
        singleton_instance.set_variable('vista_nombreMes6','')
        # singleton_instance.set_variable('vista_c_fijoMensual','0')
        # singleton_instance.set_variable('vista_c_energiaActivaPunta','0')
        # singleton_instance.set_variable('vista_c_energiaActivaFueraPunta','0')
        # singleton_instance.set_variable('vista_c_energiaReactivaExc30','0')
        # singleton_instance.set_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta','0')
        # singleton_instance.set_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta','0')
        # singleton_instance.set_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta','0')
        # singleton_instance.set_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta','0')
        singleton_instance.set_variable('vista_t_nombreTablero','')
        singleton_instance.set_variable('vista_t_energiaActivaHoraFueraPuntaActual','0')
        singleton_instance.set_variable('vista_t_energiaActivaHoraFueraPuntaAnterior','0')
        singleton_instance.set_variable('vista_t_energiaActivaHoraPuntaActual','0')
        singleton_instance.set_variable('vista_t_energiaActivaHoraPuntaAnterior','0')
        singleton_instance.set_variable('vista_t_evidencia1','')
        singleton_instance.set_variable('vista_t_maximaDemanda','0')
        singleton_instance.set_variable('vista_t_evidencia2','')
        singleton_instance.set_variable('vista_t_energiaReactivaInductivaActual','0')
        singleton_instance.set_variable('vista_t_energiaReactivaInductivaAnterior','0')
        singleton_instance.set_variable('vista_t_evidencia3','')
        singleton_instance.set_variable('vista_cantidadMedidores','0')
        singleton_instance.set_variable('data_sumaAB_t1','0')
        singleton_instance.set_variable('data_sumaAB_t2','0')
        singleton_instance.set_variable('data_sumaAB_total','0')
        singleton_instance.set_variable('data_sumaC_total','0')
        singleton_instance.set_variable('data_sumaD_total','0')
        singleton_instance.set_variable('data_promedio','0')
        singleton_instance.set_variable('data_horasPunta','0')
        singleton_instance.set_variable('data_calificacion','')
        singleton_instance.set_variable('data_calificacionTarifaria','0')

        singleton_instance.set_variable('vista_a_cantidadMedidores','0')
        singleton_instance.set_variable('vista_a_Actual','0')
        singleton_instance.set_variable('vista_a_Anterior','0')
        singleton_instance.set_variable('vista_a_evidencia1','')
        singleton_instance.set_variable('data_a_total','0')
        singleton_instance.set_variable('data_a_suma','0')
        singleton_instance.set_variable('array_a_Actual', [])
        singleton_instance.set_variable('array_a_Anterior', [])
        singleton_instance.set_variable('array_a_Total', [])
        singleton_instance.set_variable('array_a_evidencia1', [])

        singleton_instance.set_variable('vista_estado','falta guradar')

        singleton_instance.set_variable('array_t_nombreTablero', [])
        singleton_instance.set_variable('array_t_energiaActivaHoraFueraPuntaActual', [])
        singleton_instance.set_variable('array_t_energiaActivaHoraFueraPuntaAnterior', [])
        singleton_instance.set_variable('array_energiaActivaHoraFueraPunta', [])
        singleton_instance.set_variable('array_t_energiaActivaHoraPuntaActual', [])
        singleton_instance.set_variable('array_t_energiaActivaHoraPuntaAnterior', [])
        singleton_instance.set_variable('array_energiaActivaHoraPunta', [])
        singleton_instance.set_variable('array_energiaActivaActual', [])
        singleton_instance.set_variable('array_energiaActivaAnterior', [])
        singleton_instance.set_variable('array_energiaActivaTotal', [])
        singleton_instance.set_variable('array_t_evidencia1', [])
        singleton_instance.set_variable('array_t_maximaDemanda', [])
        singleton_instance.set_variable('array_t_evidencia2', [])
        singleton_instance.set_variable('array_t_energiaReactivaInductivaActual', [])
        singleton_instance.set_variable('array_t_energiaReactivaInductivaAnterior', [])
        singleton_instance.set_variable('array_energiaReactivaInductivaTotal', [])
        singleton_instance.set_variable('array_t_evidencia3', [])
        
        

        # reseat imagene
        ruta_imagen = 'recursos/subir_imagen.jpg'
        imagen = Image.open(ruta_imagen)
        imagen = imagen.resize((150, 100), Image.LANCZOS)
        tk_imagen = ImageTk.PhotoImage(imagen)

        self.app_instance.evidencia1_canvas.config(width=150, height=100)
        self.app_instance.evidencia1_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
        self.app_instance.evidencia1_canvas.image = tk_imagen
        self.app_instance.evidencia2_canvas.config(width=150, height=100)
        self.app_instance.evidencia2_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
        self.app_instance.evidencia2_canvas.image = tk_imagen
        self.app_instance.evidencia3_canvas.config(width=150, height=100)
        self.app_instance.evidencia3_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
        self.app_instance.evidencia3_canvas.image = tk_imagen
        self.app_instance.evidencia4_canvas.config(width=150, height=100)
        self.app_instance.evidencia4_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
        self.app_instance.evidencia4_canvas.image = tk_imagen
        self.app_instance.boton4_label.configure(state='normal')

        self.app_instance.btnfinal2.configure(state='disabled')

    def save(self):
        singleton_instance = Singleton()
        if (singleton_instance.get_variable('vista_cliente').get() =="" or 
            singleton_instance.get_variable('vista_area').get() =="" or
            singleton_instance.get_variable('vista_informe').get() =="" or
            singleton_instance.get_variable('vista_dias').get() =="0" or
            singleton_instance.get_variable('vista_dias').get() =="" or
            singleton_instance.get_variable('vista_mes1').get() =="" or
            singleton_instance.get_variable('vista_nombreMes1').get() =="" or
            singleton_instance.get_variable('vista_c_fijoMensual').get() =="" or
            singleton_instance.get_variable('vista_c_energiaActivaPunta').get() =="" or
            singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get() =="" or
            singleton_instance.get_variable('vista_c_energiaReactivaExc30').get() =="" or
            singleton_instance.get_variable('vista_c_energiaReactivaExc30').get() =="0" or
            singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get() =="" or
            singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get() =="" or
            singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get() =="" or
            singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get() =="" or
            singleton_instance.get_variable('vista_cantidadMedidores').get() =="0"   ):
            messagebox.showwarning('Mensaje de advertencia','Rellene todos los campos con datos reales e ingrese al menos 1 medidor')
        else:
            # self.app_instance.btnfinal2.configure(state='normal')

            meses=[]
            meses.append(round(float(singleton_instance.get_variable('vista_mes1').get()),2))
            meses.append(round(float(singleton_instance.get_variable('vista_mes2').get()),2))
            meses.append(round(float(singleton_instance.get_variable('vista_mes3').get()),2))
            meses.append(round(float(singleton_instance.get_variable('vista_mes4').get()),2))
            meses.append(round(float(singleton_instance.get_variable('vista_mes5').get()),2))
            meses.append(round(float(singleton_instance.get_variable('vista_mes6').get()),2))
            meses.sort()
            if meses[4] == 0:
                singleton_instance.set_variable('data_promedio',str(meses[5]))
            else:
                singleton_instance.set_variable('data_promedio',str(round(((meses[5]+meses[4])/2),2)))


            print (singleton_instance.get_variable('data_promedio').get())
            singleton_instance.set_variable('data_horasPunta',str(int(singleton_instance.get_variable('vista_dias').get())*5))
            print (singleton_instance.get_variable('data_horasPunta').get())
            singleton_instance.set_variable('data_calificacionTarifaria',str(round(float(singleton_instance.get_variable('data_sumaAB_t2').get())/(float(singleton_instance.get_variable('data_sumaC_total').get())*float(singleton_instance.get_variable('data_horasPunta').get())),2)))
            print (singleton_instance.get_variable('data_calificacionTarifaria').get())
            print("data_sumaAB_t2:", singleton_instance.get_variable('data_sumaAB_t2').get())
            print("data_sumaC_total:", singleton_instance.get_variable('data_sumaC_total').get())
            print("data_horasPunta:", singleton_instance.get_variable('data_horasPunta').get())
            
            if (float(singleton_instance.get_variable('data_calificacionTarifaria').get()) >= 0.50):
                singleton_instance.set_variable('data_calificacion',"0")
            else:
                singleton_instance.set_variable('data_calificacion',"1")
            
            self.app_instance.btnfinal2.configure(state='normal')

    def change_mes(self):
        singleton_instance = Singleton()
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        fecha_var = singleton_instance.get_variable('vista_lectura')
        fecha = fecha_var.get()  # Obtén el valor de la StringVar

        if fecha:
            partes_fecha = fecha.split('-')
            
            if len(partes_fecha) > 1:
                numero_mes = int(partes_fecha[1])
                
                if 1 <= numero_mes <= 12:
                    nombre_mes = meses[numero_mes - 1]
                    print('El mes actual es: ' + nombre_mes)
                else:
                    print('Número de mes no válido:', numero_mes)
            else:
                print('Formato de fecha no válido:', fecha)
        else:
            print('La variable de fecha está vacía.')

        
        singleton_instance.set_variable('vista_nombreMes1',meses[numero_mes - 6])
        singleton_instance.set_variable('vista_nombreMes2',meses[numero_mes - 5])
        singleton_instance.set_variable('vista_nombreMes3',meses[numero_mes - 4])
        singleton_instance.set_variable('vista_nombreMes4',meses[numero_mes - 3])
        singleton_instance.set_variable('vista_nombreMes5',meses[numero_mes - 2])
        singleton_instance.set_variable('vista_nombreMes6',meses[numero_mes - 1])
    
    def add_medidor(self):
        singleton_instance = Singleton()

        if (singleton_instance.get_variable('vista_t_nombreTablero').get() =="" or 
            singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get() =="" or
            singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get() =="" or
            singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get() =="" or
            singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get() =="" or
            singleton_instance.get_variable('vista_t_maximaDemanda').get() =="" or
            singleton_instance.get_variable('vista_t_energiaReactivaInductivaActual').get() =="" or
            singleton_instance.get_variable('vista_t_energiaReactivaInductivaAnterior').get() =="" or
            singleton_instance.get_variable('vista_t_evidencia1').get() =="" or
            singleton_instance.get_variable('vista_t_evidencia2').get() =="" or
            singleton_instance.get_variable('vista_t_evidencia3').get() ==""  ):
            messagebox.showwarning('Mensaje de advertencia','Rellene todos los campos del medidor')
        else:
            singleton_instance.get_variable('array_t_nombreTablero').append(singleton_instance.get_variable('vista_t_nombreTablero').get())
            singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaActual').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())))
            singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaAnterior').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())))
            singleton_instance.get_variable('array_energiaActivaHoraFueraPunta').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())))
            singleton_instance.get_variable('array_t_energiaActivaHoraPuntaActual').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())))
            singleton_instance.get_variable('array_t_energiaActivaHoraPuntaAnterior').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            singleton_instance.get_variable('array_energiaActivaHoraPunta').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            singleton_instance.get_variable('array_energiaActivaActual').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())))
            singleton_instance.get_variable('array_energiaActivaAnterior').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            singleton_instance.get_variable('array_energiaActivaTotal').append(str(float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            singleton_instance.get_variable('array_t_evidencia1').append(singleton_instance.get_variable('vista_t_evidencia1').get())
            singleton_instance.get_variable('array_t_maximaDemanda').append(str(float(singleton_instance.get_variable('vista_t_maximaDemanda').get())))
            singleton_instance.get_variable('array_t_evidencia2').append(singleton_instance.get_variable('vista_t_evidencia2').get())
            singleton_instance.get_variable('array_t_energiaReactivaInductivaActual').append(str(float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaActual').get())))
            singleton_instance.get_variable('array_t_energiaReactivaInductivaAnterior').append(str(float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaAnterior').get())))
            singleton_instance.get_variable('array_energiaReactivaInductivaTotal').append(str(float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaActual').get())-float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaAnterior').get())))
            singleton_instance.get_variable('array_t_evidencia3').append(singleton_instance.get_variable('vista_t_evidencia3').get())

            singleton_instance.set_variable('data_sumaAB_t1',str(float(singleton_instance.get_variable('data_sumaAB_t1').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())))
            singleton_instance.set_variable('data_sumaAB_t2',str(float(singleton_instance.get_variable('data_sumaAB_t2').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            singleton_instance.set_variable('data_sumaAB_total',str(float(singleton_instance.get_variable('data_sumaAB_total').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaActual').get())+float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaActual').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraFueraPuntaAnterior').get())-float(singleton_instance.get_variable('vista_t_energiaActivaHoraPuntaAnterior').get())))
            
            singleton_instance.set_variable('data_sumaC_total',str(float(singleton_instance.get_variable('data_sumaC_total').get())+float(singleton_instance.get_variable('vista_t_maximaDemanda').get())))

            singleton_instance.set_variable('data_sumaD_total',str(float(singleton_instance.get_variable('data_sumaD_total').get())+float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaActual').get())-float(singleton_instance.get_variable('vista_t_energiaReactivaInductivaAnterior').get())))

            singleton_instance.set_variable('vista_mes6',str(round(float(float(singleton_instance.get_variable('vista_mes6').get())+float(singleton_instance.get_variable('vista_t_maximaDemanda').get())),4)))
            print('array :')
            print(singleton_instance.get_variable('array_t_nombreTablero'))
            print(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaActual'))
            print(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaAnterior'))
            print(singleton_instance.get_variable('array_energiaActivaHoraFueraPunta'))
            print(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaActual'))
            print(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaAnterior'))
            print(singleton_instance.get_variable('array_energiaActivaHoraPunta'))
            print(singleton_instance.get_variable('array_energiaActivaActual'))
            print(singleton_instance.get_variable('array_energiaActivaAnterior'))
            print(singleton_instance.get_variable('array_energiaActivaTotal'))
            print(singleton_instance.get_variable('array_t_evidencia1'))
            print(singleton_instance.get_variable('array_t_maximaDemanda'))
            print(singleton_instance.get_variable('array_t_evidencia2'))
            print(singleton_instance.get_variable('array_t_energiaReactivaInductivaActual'))
            print(singleton_instance.get_variable('array_t_energiaReactivaInductivaAnterior'))
            print(singleton_instance.get_variable('array_energiaReactivaInductivaTotal'))
            print(singleton_instance.get_variable('array_t_evidencia3'))
            print('data_sumaAB_t1')
            print(singleton_instance.get_variable('data_sumaAB_t1').get())
            print('data_sumaAB_t2')
            print(singleton_instance.get_variable('data_sumaAB_t2').get())
            print('data_sumaAB_total')
            print(singleton_instance.get_variable('data_sumaAB_total').get())
        
            singleton_instance.set_variable('vista_t_nombreTablero','')
            singleton_instance.set_variable('vista_t_energiaActivaHoraFueraPuntaActual','0')
            singleton_instance.set_variable('vista_t_energiaActivaHoraFueraPuntaAnterior','0')
            singleton_instance.set_variable('vista_t_energiaActivaHoraPuntaActual','0')
            singleton_instance.set_variable('vista_t_energiaActivaHoraPuntaAnterior','0')
            singleton_instance.set_variable('vista_t_evidencia1','')
            singleton_instance.set_variable('vista_t_maximaDemanda','0')
            singleton_instance.set_variable('vista_t_evidencia2','')
            singleton_instance.set_variable('vista_t_energiaReactivaInductivaActual','0')
            singleton_instance.set_variable('vista_t_energiaReactivaInductivaAnterior','0')
            singleton_instance.set_variable('vista_t_evidencia3','')
            singleton_instance.set_variable('vista_cantidadMedidores',str(int(singleton_instance.get_variable('vista_cantidadMedidores').get())+1))

            # Restear imagen
            ruta_imagen = 'recursos/subir_imagen.jpg'
            imagen = Image.open(ruta_imagen)
            imagen = imagen.resize((150, 100), Image.LANCZOS)
            tk_imagen = ImageTk.PhotoImage(imagen)
            # Mostrar la imagen en el lienzo
            self.app_instance.evidencia1_canvas.config(width=150, height=100)
            self.app_instance.evidencia1_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
            self.app_instance.evidencia1_canvas.image = tk_imagen
            self.app_instance.evidencia2_canvas.config(width=150, height=100)
            self.app_instance.evidencia2_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
            self.app_instance.evidencia2_canvas.image = tk_imagen
            self.app_instance.evidencia3_canvas.config(width=150, height=100)
            self.app_instance.evidencia3_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
            self.app_instance.evidencia3_canvas.image = tk_imagen

            if singleton_instance.get_variable('vista_cantidadMedidores').get() =='14':
                self.app_instance.boton4_label.configure(state='disabled')

    def add_a_medidor(self):
        singleton_instance = Singleton()

        if (singleton_instance.get_variable('vista_a_Actual').get() =="" or
            singleton_instance.get_variable('vista_a_Anterior').get() =="" or
            singleton_instance.get_variable('vista_a_evidencia1').get() ==""  ):
            messagebox.showwarning('Mensaje de advertencia','Rellene todos los campos del medidor')
        else:
            singleton_instance.get_variable('array_a_Actual').append(str(float(singleton_instance.get_variable('vista_a_Actual').get())))
            singleton_instance.get_variable('array_a_Anterior').append(str(float(singleton_instance.get_variable('vista_a_Anterior').get())))
            singleton_instance.get_variable('array_a_Total').append(str(float(singleton_instance.get_variable('vista_a_Actual').get())-float(singleton_instance.get_variable('vista_a_Anterior').get())))
            singleton_instance.get_variable('array_a_evidencia1').append(singleton_instance.get_variable('vista_a_evidencia1').get())


            singleton_instance.set_variable('data_a_suma',str(float(singleton_instance.get_variable('data_a_suma').get())+(float(singleton_instance.get_variable('vista_a_Actual').get())-float(singleton_instance.get_variable('vista_a_Anterior').get()))))

            singleton_instance.set_variable('vista_a_cantidadMedidores',str(int(singleton_instance.get_variable('vista_a_cantidadMedidores').get())+1))

            singleton_instance.set_variable('vista_a_Actual','0')
            singleton_instance.set_variable('vista_a_Anterior','0')
            singleton_instance.set_variable('vista_a_evidencia1','')

            # Restear imagen
            ruta_imagen = 'recursos/subir_imagen.jpg'
            imagen = Image.open(ruta_imagen)
            imagen = imagen.resize((150, 100), Image.LANCZOS)
            tk_imagen = ImageTk.PhotoImage(imagen)
            # Mostrar la imagen en el lienzo
            self.app_instance.evidencia4_canvas.config(width=150, height=100)
            self.app_instance.evidencia4_canvas.create_image(0, 0, anchor='nw', image=tk_imagen)
            self.app_instance.evidencia4_canvas.image = tk_imagen
    
    def exportar(self):
        
        singleton_instance = Singleton()
        
        #DATOS DE LOS CARGOS DE CONSUMO 
        global cargo_eap
        global cargo_eafp
        global cargo_pagpp
        global cargo_pagfp
        global cargo_pardpp
        global cargo_pardfp
        global cargo_ere30
        global cargo_agua

        cargo_agua= float(singleton_instance.get_variable('data_a_suma').get())*15
        cargo_eap= float(singleton_instance.get_variable('data_sumaAB_t2').get())*float(singleton_instance.get_variable('vista_c_energiaActivaPunta').get())*0.01
        cargo_eafp= float(singleton_instance.get_variable('data_sumaAB_t1').get())*float(singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get())*0.01
        cargo_ere30= (float(singleton_instance.get_variable('data_sumaD_total').get())-(0.3*float(singleton_instance.get_variable('data_sumaAB_total').get())))*float(singleton_instance.get_variable('vista_c_energiaReactivaExc30').get())*0.01
        if (cargo_ere30<0):
            cargo_ere30=0
        
        
        #OTROS DATOS X QUE SOLO SIRVEN PARA RELLENAR EL DOCUMENTO WORD
        global dx_tableros
        global dx_fecha_emision
        global dx_fecha_lectura
        global dx_texto_1
        global dx_texto_2
        global dx_anio
        global dx_operacion1
        global dx_operacion2
        global dx_operacion3
        global dx_operacion4
        global detalle_1
        global detalle_2
        global detalle_3
        global detalle_4
        global detalle_5
        global detalle_6
        global detalle_7
        global x3

        nom_tablero_lista = singleton_instance.get_variable('array_t_nombreTablero')
        dx_tableros = ', '.join(map(str, nom_tablero_lista))
        dx_tableros = dx_tableros.upper()

        partes_fecha_lectura = singleton_instance.get_variable('vista_lectura').get().split('-')
        partes_fecha_emision = singleton_instance.get_variable('vista_emision').get().split('-')

        meses = {
            1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
            7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
        }

        mes_lectura = meses.get(int(partes_fecha_lectura[1]), "")
        mes_emision = meses.get(int(partes_fecha_emision[1]), "")

        dx_fecha_lectura = f"{partes_fecha_lectura[0]} de {mes_lectura} del {partes_fecha_lectura[2]}".upper()

        dx_fecha_emision = f"{partes_fecha_emision[0]} de {mes_emision} del {partes_fecha_emision[2]}"
        dx_anio = partes_fecha_emision[2]

        # Puedes imprimir o utilizar las variables como necesites
        print("Fecha de lectura:", dx_fecha_lectura)
        print("Fecha de emisión:", dx_fecha_emision)
        print("Año:", dx_anio)
        print("CALIFCACION DE USUARIO:", singleton_instance.get_variable('data_calificacion').get())
        def redondeo(self):
            number=round(float(self),2)
            return number
        def redondeo2(self):
            number=str(round(float(self),2))
            return number

        if (singleton_instance.get_variable('data_calificacion').get() == '0'):
            dx_texto_1 = "CLIENTE PRESENTE EN PUNTA"

            cargo_pagpp= float(singleton_instance.get_variable('data_sumaC_total').get())*float(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get())
            cargo_pagfp= 0
            cargo_pardpp= float(singleton_instance.get_variable('data_promedio').get())*float(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get())
            cargo_pardfp= 0

            dx_operacion1 = redondeo2(singleton_instance.get_variable('data_sumaC_total').get()) + ' Kw x ' + redondeo2(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get()) + ' S/./kW.h = '+str(redondeo2(cargo_pagpp))+' S/./kW.h'
            dx_operacion2 = ""
            dx_operacion3 = redondeo2(singleton_instance.get_variable('data_promedio').get()) + ' Kw x ' + redondeo2(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get()) + ' S/./kW.h = '+str(redondeo2(cargo_pardpp))+' S/./kW.h'
            dx_operacion4 = ""

            detalle_1= singleton_instance.get_variable('data_sumaC_total').get()
            detalle_2= "0"
            detalle_3= singleton_instance.get_variable('data_promedio').get()
            detalle_4= "0"

        elif(singleton_instance.get_variable('data_calificacion').get() == '1'):
            dx_texto_1 = "CLIENTE FUERA DE PUNTA"

            cargo_pagpp= 0
            cargo_pagfp= float(singleton_instance.get_variable('data_sumaC_total').get())*float(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get())
            cargo_pardpp= 0
            cargo_pardfp= float(singleton_instance.get_variable('data_promedio').get())*float(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get())

            dx_operacion1 = ""
            dx_operacion2 = redondeo2(singleton_instance.get_variable('data_sumaC_total').get()) + ' Kw x ' + redondeo2(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get()) + ' S/./kW.h = '+redondeo2(cargo_pagfp)+' S/./kW.h'
            dx_operacion3 = ""
            dx_operacion4 = redondeo2(singleton_instance.get_variable('data_promedio').get()) + ' Kw x ' + redondeo2(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get()) + ' S/./kW.h = '+redondeo2(cargo_pardfp)+' S/./kW.h'
        
            detalle_1= "0"
            detalle_2= singleton_instance.get_variable('data_sumaC_total').get()
            detalle_3= "0"
            detalle_4= singleton_instance.get_variable('data_promedio').get()
        
        detalle_5= round((float(singleton_instance.get_variable('vista_c_energiaActivaPunta').get())*0.01),4)
        detalle_6= round((float(singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get())*0.01),4)
        detalle_7= round((float(singleton_instance.get_variable('vista_c_energiaReactivaExc30').get())*0.01),4)

        if (float(singleton_instance.get_variable('data_sumaD_total').get())-0.3*float(singleton_instance.get_variable('data_sumaAB_total').get())) > 0:
            dx_texto_2= "Se factura cargos por energia reactiva, se cumple la condición inicial."
        else:
            dx_texto_2= "No se factura cargos por energia reactiva, no se cumple la condición inicial."
        
        subtotal= float(singleton_instance.get_variable('vista_c_fijoMensual').get())+ cargo_eap + cargo_eafp + cargo_ere30 + cargo_pagpp + cargo_pagfp + cargo_pardpp + cargo_pardfp
        print('DATA DE NOMBRE SUBTOTAL')
        print(singleton_instance.get_variable('vista_c_fijoMensual').get())   
        print(subtotal)  
        igv= float(subtotal) * 0.18
        total= subtotal + igv 

        if(float(singleton_instance.get_variable('data_sumaD_total').get())-0.3*float(singleton_instance.get_variable('data_sumaAB_total').get())<0):
            x3 = 0.0
        else:
            x3 = redondeo(float(singleton_instance.get_variable('data_sumaD_total').get())-0.3*float(singleton_instance.get_variable('data_sumaAB_total').get()))
        

        ##--------------------WORD--------------------##
        # Seleccionar template según el checkbox de firma
        incluir_firma = singleton_instance.get_variable('firma').get()
        cantidad_a_medidores = int(singleton_instance.get_variable('vista_a_cantidadMedidores').get())
        if cantidad_a_medidores==0:
            if incluir_firma:
                template_path = "recursos/templates/template_f.docx"
                print("firma")
            else:
                template_path = "recursos/templates/template.docx"
                print("sin firma")
        else:
            if incluir_firma:
                template_path = "recursos/templates/template_a_f.docx"
                print("firma")
            else:
                template_path = "recursos/templates/template_a.docx"
                print("sin firma")

        doc = DocxTemplate(template_path)
        medidores = []
        cantidad_medidores = int(singleton_instance.get_variable('vista_cantidadMedidores').get())
        for i in range(cantidad_medidores):
            medidores.append({
                'nombre_tablero': singleton_instance.get_variable('array_t_nombreTablero')[i],
                'energia_activa_hora_fuera_punta_actual': redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaActual')[i]),
                'energia_activa_hora_fuera_punta_anterior': redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaAnterior')[i]),
                'energia_activa_hora_fuera_punta': redondeo(singleton_instance.get_variable('array_energiaActivaHoraFueraPunta')[i]),
                'energia_activa_hora_punta_actual': redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaActual')[i]),
                'energia_activa_hora_punta_anterior': redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaAnterior')[i]),
                'energia_activa_hora_punta': redondeo(singleton_instance.get_variable('array_energiaActivaHoraPunta')[i]),
                'energia_activa_actual': redondeo(singleton_instance.get_variable('array_energiaActivaActual')[i]),
                'energia_activa_anterior': redondeo(singleton_instance.get_variable('array_energiaActivaAnterior')[i]),
                'energia_activa_total': redondeo(singleton_instance.get_variable('array_energiaActivaTotal')[i]),
                'evidencia1': InlineImage(doc, singleton_instance.get_variable('array_t_evidencia1')[i], height=Mm(35), width=Mm(50)),
                'maxima_demanda': redondeo(singleton_instance.get_variable('array_t_maximaDemanda')[i]),
                'evidencia2': InlineImage(doc, singleton_instance.get_variable('array_t_evidencia2')[i], height=Mm(35), width=Mm(50)),
                'energia_reactiva_inductiva_actual': redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaActual')[i]),
                'energia_reactiva_inductiva_anterior': redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaAnterior')[i]),
                'energia_reactiva_inductiva_total': redondeo(singleton_instance.get_variable('array_energiaReactivaInductivaTotal')[i]),
                'evidencia3': InlineImage(doc, singleton_instance.get_variable('array_t_evidencia3')[i], height=Mm(35), width=Mm(50))
            })
        medidores_a = []
        for i in range(cantidad_a_medidores):
            medidores_a.append({
                'id': i+1,
                'a_actual': redondeo(singleton_instance.get_variable('array_a_Actual')[i]),
                'a_anterior': redondeo(singleton_instance.get_variable('array_a_Anterior')[i]),
                'a_total': redondeo(singleton_instance.get_variable('array_a_Total')[i]),
                'a_evidencia': InlineImage(doc, singleton_instance.get_variable('array_a_evidencia1')[i], height=Mm(35), width=Mm(50))
            })
            
        context = {
        "informe_tecnico" : singleton_instance.get_variable('vista_informe').get(),
        "tablero" : dx_tableros,
        "fecha_lectura": dx_fecha_lectura,
        "fecha_emision": dx_fecha_emision,
        "cliente": singleton_instance.get_variable('vista_cliente').get(),
        "area": singleton_instance.get_variable('vista_area').get(),
        "anio": dx_anio,
        
        "ct_1": redondeo(singleton_instance.get_variable('vista_c_fijoMensual').get() ),
        "ct_2": redondeo(singleton_instance.get_variable('vista_c_energiaActivaPunta').get() ),
        "ct_3": redondeo(singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get()) ,
        "ct_4": redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get()) ,
        "ct_5": redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get() ),
        "ct_6": redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get() ),
        "ct_7": redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get()) ,
        "ct_8": redondeo(singleton_instance.get_variable('vista_c_energiaReactivaExc30').get()) ,
        
        "medidores":medidores,
        "medidores_a":medidores_a,
        
        "fp_mes_total": redondeo(singleton_instance.get_variable('data_sumaAB_t1').get() ) ,
        "hp_mes_total": redondeo(singleton_instance.get_variable('data_sumaAB_t2').get()) ,
        "ea_mes_total": redondeo(singleton_instance.get_variable('data_sumaAB_total').get()) ,
        "maxima_demanda_total": redondeo(singleton_instance.get_variable('data_sumaC_total').get()) ,
        "er_mes_total": redondeo(singleton_instance.get_variable('data_sumaD_total').get()) ,

        "dias": singleton_instance.get_variable('vista_dias').get(),
        "horas_punta": singleton_instance.get_variable('data_horasPunta').get() ,
        
        "texto_1": dx_texto_1,
        "calificacion_tarifaria": singleton_instance.get_variable('data_calificacionTarifaria').get(),

        "cargo_eapp": redondeo(cargo_eap),
        "cargo_eafp": redondeo(cargo_eafp),

        "operacion1": dx_operacion1 ,
        "operacion2": dx_operacion2 ,
        "operacion3": dx_operacion3 ,
        "operacion4": dx_operacion4 ,
    
        "x1": redondeo(0.3*float(singleton_instance.get_variable('data_sumaAB_total').get())),
        "x2": redondeo(float(singleton_instance.get_variable('data_sumaD_total').get())-0.3*float(singleton_instance.get_variable('data_sumaAB_total').get())),
        
        "texto_2": dx_texto_2,
        "x3": x3,
        "cargo_ere30": redondeo(cargo_ere30),

        "nombre_mes1": singleton_instance.get_variable('vista_nombreMes1').get(),
        "cantidad_mes1": redondeo(singleton_instance.get_variable('vista_mes1').get()),
        "nombre_mes2": singleton_instance.get_variable('vista_nombreMes2').get(),
        "cantidad_mes2": redondeo(singleton_instance.get_variable('vista_mes2').get()),
        "nombre_mes3": singleton_instance.get_variable('vista_nombreMes3').get(),
        "cantidad_mes3": redondeo(singleton_instance.get_variable('vista_mes3').get()),
        "nombre_mes4": singleton_instance.get_variable('vista_nombreMes4').get(),
        "cantidad_mes4": redondeo(singleton_instance.get_variable('vista_mes4').get()),
        "nombre_mes5": singleton_instance.get_variable('vista_nombreMes5').get(),
        "cantidad_mes5": redondeo(singleton_instance.get_variable('vista_mes5').get()),
        "nombre_mes6": singleton_instance.get_variable('vista_nombreMes6').get(),
        "cantidad_mes6": redondeo(singleton_instance.get_variable('vista_mes6').get()),

        "promedio":redondeo(singleton_instance.get_variable('data_promedio').get()), 

        
        "cargo_pagpp":  redondeo(cargo_pagpp),
        "cargo_pagfp":  redondeo(cargo_pagfp),
        "cargo_parpp":  redondeo(cargo_pardpp),
        "cargo_parfp":  redondeo(cargo_pardfp),

        "detalle_1":  redondeo(detalle_1),
        "detalle_2":  redondeo(detalle_2),
        "detalle_3":  redondeo(detalle_3),
        "detalle_4":  redondeo(detalle_4),
        "detalle_5":  detalle_5,
        "detalle_6":  detalle_6,
        "detalle_7":  detalle_7,

        "subtotal": redondeo(subtotal),
        "conigv": redondeo(igv),
        "total_final": redondeo(total) ,     

        "cargo_agua":  redondeo(cargo_agua),
        "data_a_suma":  redondeo(singleton_instance.get_variable('data_a_suma').get()),
        "a_conigv": redondeo(cargo_agua* 0.18),
        "a_total_final": redondeo(cargo_agua+cargo_agua* 0.18) ,    

        }
        doc.render(context)
        nombrefinal="Informes-Word\INFORME TECNICO N° "+singleton_instance.get_variable('vista_informe').get()+".docx"
        print(nombrefinal)
        doc.save(nombrefinal)
        print("template 1")

        messagebox.showinfo('Mensaje informativo',nombrefinal+' a sido creado exitosamente')

        ##--------------------EXCEL--------------------##

        if cantidad_a_medidores==0:
            workbook= xlsxwriter.Workbook('Informes-Excel\Excel_Informe-'+singleton_instance.get_variable('vista_informe').get()+'.xlsx')
            worksheet=workbook.add_worksheet()

            row=0
            col=0
            # Datos generales
            datos_generales = [
                ['DATOS GENERALES', '', '', '', ''],
                ['Informe Tecnico', singleton_instance.get_variable('vista_informe').get(), '', '', ''],
                ['Cliente', singleton_instance.get_variable('vista_cliente').get(), '', '', ''],
                ['Área', singleton_instance.get_variable('vista_area').get(), '', '', ''],
                ['Fecha de Lectura', singleton_instance.get_variable('vista_lectura').get(), '', '', ''],
                ['Fecha de Emision', singleton_instance.get_variable('vista_emision').get(), '', '', ''],
                ['Dias del mes', singleton_instance.get_variable('vista_dias').get(), '', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de máxima demanda
            maxima_demanda = [
                ['MAXIMA DEMANDA', '', '', '', ''],
                ['MES', 'CONSUMO', 'UNIDAD', '', ''],
                [singleton_instance.get_variable('vista_nombreMes1').get(), singleton_instance.get_variable('vista_mes1').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes2').get(), singleton_instance.get_variable('vista_mes2').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes3').get(), singleton_instance.get_variable('vista_mes3').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes4').get(), singleton_instance.get_variable('vista_mes4').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes5').get(), singleton_instance.get_variable('vista_mes5').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes6').get(), redondeo(singleton_instance.get_variable('vista_mes6').get()), 'Kw.', '', ''],
                ['Potencia Activa de redes de distribución', singleton_instance.get_variable('data_promedio').get(), 'Kw.', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de pliego tarifario
            pliego_tarifario = [
                ['PLIEGO TARIFARIO', '', '', '', ''],
                ['TARIFA BT3', 'TARIFA SIN IGV', 'UNIDAD', '', ''],
                ['Cargo fijo mensual', singleton_instance.get_variable('vista_c_fijoMensual').get(), '(S/./mes)', '', ''],
                ['Cargo de energia activa en punta', singleton_instance.get_variable('vista_c_energiaActivaPunta').get(), '(Ctm.S/./Kw.h)', '', ''],
                ['Cargo de energia activa fuera de punta', singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get(), '(Ctm.S/./Kw.h)', '', ''],
                ['Cargo por potencia activa de generación para ususarios presente en punta', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de generación para ususarios presente fuera de punta', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de redes de distribución para usuarios presente en punta', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de redes de distribución para usuarios presente fuera de punta', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por energía reactiva que exceda el 30% Del total del energía activa ', singleton_instance.get_variable('vista_c_energiaReactivaExc30').get(), '(Ctm.S/./kVar.h)', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de cada medidor
            datos_medidores = []
            for i in range(cantidad_medidores):
                datos_medidores.extend([
                    [f'Medidor {i+1} - {singleton_instance.get_variable("array_t_nombreTablero")[i]}', '', '', '', ''],
                    ['Energía activa en hora fuera de punta mes actual', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora fuera de punta mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa en hora fuera de punta del mes', redondeo(singleton_instance.get_variable('array_energiaActivaHoraFueraPunta')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora punta mes actual', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora punta mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa en hora punta del mes', redondeo(singleton_instance.get_variable('array_energiaActivaHoraPunta')[i]), 'Kw/h', '', ''],
                    ['Energía activa total de mes actual', redondeo(singleton_instance.get_variable('array_energiaActivaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa total de mes anterior', redondeo(singleton_instance.get_variable('array_energiaActivaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa total del mes', redondeo(singleton_instance.get_variable('array_energiaActivaTotal')[i]), 'Kw/h', '', ''],
                    ['Maxima demanda', redondeo(singleton_instance.get_variable('array_t_maximaDemanda')[i]), 'Kw', '', ''],
                    ['Energía reactiva inductiva total del mes actual', redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaActual')[i]), 'Kvar/Lh', '', ''],
                    ['Energía reactiva inductiva total del mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaAnterior')[i]), 'Kvar/Lh', '', ''],
                    ['Energía reactiva inductiva total del mes', redondeo(singleton_instance.get_variable('array_energiaReactivaInductivaTotal')[i]), 'Kvar/Lh', '', ''],
                    ['', '', '', '', '']
                ])

            # Datos de pliego tarifario
            Total_factura = [
                ['TOTAL A FACTURAR', '', '', '', ''],
                ['DETALLE', 'CONSUMO', 'PRECIO', 'UNIDAD', 'IMPORTE'],
                ['Cargo fijo mensual', '', singleton_instance.get_variable('vista_c_fijoMensual').get(), '(S/./mes)', redondeo(singleton_instance.get_variable('vista_c_fijoMensual').get() )],
                ['Cargo de energia activa en punta', '', singleton_instance.get_variable('vista_c_energiaActivaPunta').get(), '(Ctm.S/./Kw.h)', redondeo(singleton_instance.get_variable('vista_c_energiaActivaPunta').get() )],
                ['Cargo de energia activa fuera de punta', '', singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get(), '(Ctm.S/./Kw.h)', redondeo(singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get())],
                ['Cargo por potencia activa de generación para ususarios presente en punta', '', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get())],
                ['Cargo por potencia activa de generación para ususarios presente fuera de punta', '', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get() )],
                ['Cargo por potencia activa de redes de distribución para usuarios presente en punta', '', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get() )],
                ['Cargo por potencia activa de redes de distribución para usuarios presente fuera de punta', '', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get())],
                ['Cargo por energía reactiva que exceda el 30% Del total del energía activa ', '', singleton_instance.get_variable('vista_c_energiaReactivaExc30').get(), '(Ctm.S/./kVar.h)', redondeo(singleton_instance.get_variable('vista_c_energiaReactivaExc30').get())],
                ['', '', '', 'SUB', redondeo(subtotal)],
                ['', '', '', 'IGV', redondeo(igv)],
                ['', '', '', 'TOTAL', redondeo(total)],
                ['', '', '', '', '']
            ]

            # Unir todos los datos
            conjunto_listas = datos_generales + maxima_demanda + pliego_tarifario + datos_medidores +Total_factura

            # Escribir los datos en el worksheet
            for fila in conjunto_listas:
                worksheet.write_row(row, col, fila)
                row += 1

            # Cerrar el workbook
            workbook.close()
            print(f"Excel guardado en Informes-Excel")
        else:
            
            workbook= xlsxwriter.Workbook('Informes-Excel\Excel_Informe-'+singleton_instance.get_variable('vista_informe').get()+'.xlsx')
            worksheet=workbook.add_worksheet()

            row=0
            col=0
            # Datos generales
            datos_generales = [
                ['DATOS GENERALES', '', '', '', ''],
                ['Informe Tecnico', singleton_instance.get_variable('vista_informe').get(), '', '', ''],
                ['Cliente', singleton_instance.get_variable('vista_cliente').get(), '', '', ''],
                ['Área', singleton_instance.get_variable('vista_area').get(), '', '', ''],
                ['Fecha de Lectura', singleton_instance.get_variable('vista_lectura').get(), '', '', ''],
                ['Fecha de Emision', singleton_instance.get_variable('vista_emision').get(), '', '', ''],
                ['Dias del mes', singleton_instance.get_variable('vista_dias').get(), '', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de máxima demanda
            maxima_demanda = [
                ['MAXIMA DEMANDA', '', '', '', ''],
                ['MES', 'CONSUMO', 'UNIDAD', '', ''],
                [singleton_instance.get_variable('vista_nombreMes1').get(), singleton_instance.get_variable('vista_mes1').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes2').get(), singleton_instance.get_variable('vista_mes2').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes3').get(), singleton_instance.get_variable('vista_mes3').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes4').get(), singleton_instance.get_variable('vista_mes4').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes5').get(), singleton_instance.get_variable('vista_mes5').get(), 'Kw.', '', ''],
                [singleton_instance.get_variable('vista_nombreMes6').get(), redondeo(singleton_instance.get_variable('vista_mes6').get()), 'Kw.', '', ''],
                ['Potencia Activa de redes de distribución', singleton_instance.get_variable('data_promedio').get(), 'Kw.', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de pliego tarifario
            pliego_tarifario = [
                ['PLIEGO TARIFARIO', '', '', '', ''],
                ['TARIFA BT3', 'TARIFA SIN IGV', 'UNIDAD', '', ''],
                ['Cargo fijo mensual', singleton_instance.get_variable('vista_c_fijoMensual').get(), '(S/./mes)', '', ''],
                ['Cargo de energia activa en punta', singleton_instance.get_variable('vista_c_energiaActivaPunta').get(), '(Ctm.S/./Kw.h)', '', ''],
                ['Cargo de energia activa fuera de punta', singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get(), '(Ctm.S/./Kw.h)', '', ''],
                ['Cargo por potencia activa de generación para ususarios presente en punta', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de generación para ususarios presente fuera de punta', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de redes de distribución para usuarios presente en punta', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por potencia activa de redes de distribución para usuarios presente fuera de punta', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', '', ''],
                ['Cargo por energía reactiva que exceda el 30% Del total del energía activa ', singleton_instance.get_variable('vista_c_energiaReactivaExc30').get(), '(Ctm.S/./kVar.h)', '', ''],
                ['', '', '', '', '']
            ]

            # Datos de cada medidor
            datos_medidores = []
            for i in range(cantidad_medidores):
                datos_medidores.extend([
                    [f'Medidor {i+1} - {singleton_instance.get_variable("array_t_nombreTablero")[i]}', '', '', '', ''],
                    ['Energía activa en hora fuera de punta mes actual', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora fuera de punta mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraFueraPuntaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa en hora fuera de punta del mes', redondeo(singleton_instance.get_variable('array_energiaActivaHoraFueraPunta')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora punta mes actual', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa en hora punta mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaActivaHoraPuntaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa en hora punta del mes', redondeo(singleton_instance.get_variable('array_energiaActivaHoraPunta')[i]), 'Kw/h', '', ''],
                    ['Energía activa total de mes actual', redondeo(singleton_instance.get_variable('array_energiaActivaActual')[i]), 'Kw/h', '', ''],
                    ['Energía activa total de mes anterior', redondeo(singleton_instance.get_variable('array_energiaActivaAnterior')[i]), 'Kw/h', '', ''],
                    ['Energia activa total del mes', redondeo(singleton_instance.get_variable('array_energiaActivaTotal')[i]), 'Kw/h', '', ''],
                    ['Maxima demanda', redondeo(singleton_instance.get_variable('array_t_maximaDemanda')[i]), 'Kw', '', ''],
                    ['Energía reactiva inductiva total del mes actual', redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaActual')[i]), 'Kvar/Lh', '', ''],
                    ['Energía reactiva inductiva total del mes anterior', redondeo(singleton_instance.get_variable('array_t_energiaReactivaInductivaAnterior')[i]), 'Kvar/Lh', '', ''],
                    ['Energía reactiva inductiva total del mes', redondeo(singleton_instance.get_variable('array_energiaReactivaInductivaTotal')[i]), 'Kvar/Lh', '', ''],
                    ['', '', '', '', '']
                ])

            # Datos de cada medidor
            datos_a_medidores = []
            for i in range(cantidad_a_medidores):
                datos_a_medidores.extend([
                    [f'Medidor {i+1} ', '', '', '', ''],
                    ['Medicion Actual', redondeo(singleton_instance.get_variable('array_a_Actual')[i]), 'M3', '', ''],
                    ['Medicion Anterior', redondeo(singleton_instance.get_variable('array_a_Anterior')[i]), 'M3', '', ''],
                    ['Total', redondeo(singleton_instance.get_variable('array_a_Total')[i]), 'M3', '', ''],             
                    ['', '', '', '', '']
                ])

            Total_factura = [
                ['TOTAL A FACTURAR', '', '', '', ''],
                ['DETALLE', 'CONSUMO', 'PRECIO', 'UNIDAD', 'IMPORTE'],
                ['Cargo fijo mensual', '', singleton_instance.get_variable('vista_c_fijoMensual').get(), '(S/./mes)', redondeo(singleton_instance.get_variable('vista_c_fijoMensual').get() )],
                ['Cargo de energia activa en punta', '', singleton_instance.get_variable('vista_c_energiaActivaPunta').get(), '(Ctm.S/./Kw.h)', redondeo(singleton_instance.get_variable('vista_c_energiaActivaPunta').get() )],
                ['Cargo de energia activa fuera de punta', '', singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get(), '(Ctm.S/./Kw.h)', redondeo(singleton_instance.get_variable('vista_c_energiaActivaFueraPunta').get())],
                ['Cargo por potencia activa de generación para ususarios presente en punta', '', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresentePunta').get())],
                ['Cargo por potencia activa de generación para ususarios presente fuera de punta', '', singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaGeneracionUsuariosPresenteFueraPunta').get() )],
                ['Cargo por potencia activa de redes de distribución para usuarios presente en punta', '', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresentePunta').get() )],
                ['Cargo por potencia activa de redes de distribución para usuarios presente fuera de punta', '', singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get(), '(S/./Kw-mes)', redondeo(singleton_instance.get_variable('vista_c_potenciaActivaRedesDistribucionUsuariosPresenteFueraPunta').get())],
                ['Cargo por energía reactiva que exceda el 30% Del total del energía activa ', '', singleton_instance.get_variable('vista_c_energiaReactivaExc30').get(), '(Ctm.S/./kVar.h)', redondeo(singleton_instance.get_variable('vista_c_energiaReactivaExc30').get())],
                ['', '', '', 'SUB', redondeo(subtotal)],
                ['', '', '', 'IGV', redondeo(igv)],
                ['', '', '', 'TOTAL', redondeo(total)],
                ['', '', '', '', '']
            ]

            Total_a_facturar = [
                ['TOTAL A FACTURAR DE AGUA', '', '', '', ''],
                ['DETALLE', 'CONSUMO', 'PRECIO', 'UNIDAD', 'IMPORTE'],
                ['Cargo fijo mensual', '', singleton_instance.get_variable('data_a_suma').get(), '(M3/SOLES)', redondeo(cargo_agua)],
                ['', '', '', 'SUB', redondeo(cargo_agua)],
                ['', '', '', 'IGV', redondeo(cargo_agua* 0.18)],
                ['', '', '', 'TOTAL', redondeo(cargo_agua+cargo_agua* 0.18)],
                ['', '', '', '', '']
            ]

            # Unir todos los datos
            conjunto_listas = datos_generales + maxima_demanda + pliego_tarifario + datos_medidores+ datos_a_medidores + Total_factura+ Total_a_facturar

            # Escribir los datos en el worksheet
            for fila in conjunto_listas:
                worksheet.write_row(row, col, fila)
                row += 1

            # Cerrar el workbook
            workbook.close()
            print(f"Excel guardado en Informes-Excel")