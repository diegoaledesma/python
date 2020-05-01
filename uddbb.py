import os
import shutil
import pandas as pd
from tkinter import *
from tkinter import filedialog
from tkinter import Menu
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk

#
# Lista con el path de los archivos
#
archivos = []

#
# Funcion para unificar las bases de datos.
#
def unificar():

    #
    # Se verifica que 'Columna en común' no este vacía.
    #
    if len(columna_entry.get()) == 0:
        messagebox.showwarning('Atención!', "'Columna en común' no puede estar vacía!")
    else:
        try:
            #
            # Diccionario con información de la base de datos unificada
            #
            ddbb_unificada = dict() 

            # en este set almacenamos todas las columnas que se van progresivamente encontrando
            row = set()

            # iteramos todas las bases de datos 
            for i in archivos:

                #
                # Primero se lee la extensión del archivo y luego
                # se lee el contenido indexando por la columna comun
                #
                file_name, ext = os.path.splitext(i)

                if ext == ".csv":
                    archivo = pd.read_csv(i, index_col=columna_entry.get())
                else:
                    archivo = pd.read_excel(i, index_col=columna_entry.get())

                data = archivo.to_dict("index")
                
                #
                # Iteramos la columna comun de la base de datos
                # cc == columna_comun
                #
                for cc in data: 
                    
                    #
                    # Si la columna común es nueva para la base de datos unificada
                    # se inicializa como un diccionario vacio.
                    #
                    if cc not in ddbb_unificada: 
                        ddbb_unificada[cc] = dict() 
                    
                    #
                    # En fields se almacenan las columnas de la base de datos que estamos procesando
                    # tomando como key la columna común.
                    #
                    fields = data[cc]

                    #
                    # Se iteran todos los fields de la columna comun encontrada
                    # y se le asignan sus valores en la base de datos unificada.
                    #
                    for field in fields:

                        ddbb_unificada[cc][field] = str(data[cc][field])

                        #
                        # Se le agrega al set de fields (columas) el field encontrado
                        #
                        row.add(field)

            #
            # Si el archivo existe se genera con prefijo copia_n_
            #
            def save_file(nombre):

                if os.path.exists(nombre):

                    copia_numero = 1
                    
                    while os.path.exists("copia_{}_{}".format(str(copia_numero), nombre)):
                        copia_numero += 1

                    return "copia_{}_{}".format(str(copia_numero), nombre)
                else:
                    return nombre
            
            #
            # Se crea un diccionario con la base de datos unificada
            #
            df = pd.DataFrame.from_dict(ddbb_unificada, orient='index')

            #
            # Se reemplazan los valores vacíos por '-'
            #
            df = df.fillna('-')

            #
            # Guarda la base de datos unificada con el nombre generado
            # por la función save_file.
            #
            if salida_option.get() == '.csv':
                name_save_file = save_file('listas_unificadas.csv')
                df.to_csv(name_save_file, index_label=columna_entry.get())
            else:
                name_save_file = save_file('listas_unificadas.xlsx')
                df.to_excel(name_save_file, index_label=columna_entry.get())

            #
            # Muestra mensaje informando el nombre el archivo generado.
            #
            messagebox.showinfo('Operación exitosa!', "Listas correctamente unificadas.\nSe creó el archivo:\n - {}".format(name_save_file))

        except:
            #
            # Muestra mensaje informando que algo salió mal.
            #
            messagebox.showwarning('Atención!', 'Algo salió mal.\nVerifica que la información ingresada sea la correcta.')

#
# Funcion para seleccionar los archivos de bases de datos
# que van a ser unificados
#
def openfile():
    
    #
    # Cuadro de dialogo para abrir el archivo
    #
    nuevo_archivo = filedialog.askopenfilename(initialdir = "/", title = "Seleccionar archivo", filetypes = (("Excel 2007-2019","*.xlsx"),("Microsoft Excel 97-2003","*.xls"),("Texto CSV","*.csv")))
    
    #
    # Si se selecciono correctamente un archivo
    #
    if len(nuevo_archivo) > 0:
        archivos_cargados = archivos_label.cget("text") + "\n"

        #
        # Si es el primer archivo:
        #
        if len(archivos) == 0:
            archivos_cargados = ""

            #
            # Se habilita el Entry de texto y botón de unificar
            #
            columna_entry.configure(state="normal")
            unificar_button.configure(state="normal")

        #
        # Se agrega el nuevo archivo al label 'archivos_label'
        #
        nuevo_texto = archivos_cargados + nuevo_archivo
        archivos_label.configure(text=nuevo_texto)

        #
        # Se agrega el nuevo archivo a la lista
        #
        archivos.append(nuevo_archivo)

#
# Cerrar la app
#
def exit():
    window.destroy()

window = Tk()

#
# Título de ventana
#
window.title("Unificador de archivos de DDBB")

#
# Fuente y tamaño de los mensajes.
#
window.option_add('*Dialog.msg.font', 'Helvetica 11')

#
# Dimensiones de la ventana
#
windowWidth = 560
windowHeight = 400
window.minsize(height=windowHeight, width=windowWidth)

#
# Posición de la ventana en la pantalla.
# Se realiza el calculo para centrarla.
#
positionX = int(window.winfo_screenwidth()/2 - windowWidth/2)
positionY = int(window.winfo_screenheight()/2 - windowHeight/2)
window.geometry("+{}+{}".format(positionX, positionY))

#
# Menu
#
menu = Menu(window)
menu_item = Menu(menu, tearoff=0)
menu_item.add_command(label='Salir', command=exit)
menu.add_cascade(label='Menú', menu=menu_item)
window.config(menu=menu)

#
# Cabecera
#
image_open = Image.open("img.png").resize((185, 95))
image_png = ImageTk.PhotoImage(image_open)
Label(window, image=image_png).grid(column=0, row=0, columnspan=2)
title_label = Label(window, text="Unificador de archivos de DDBB", font=("Helvetica",16))
title_label.grid(column=0, row=1, columnspan=2, pady=5)

#
# Boton de selección de archivos
#
archivos_button = Button(window, text="Seleccionar archivos", command=openfile)
archivos_button.grid(column=0, row=2, padx=20, pady=10, sticky="E")

#
# Label con la lista de archivos
#
archivos_label = Label(window, text="No hay archivos seleccionados", borderwidth=1, relief="sunken", width=40)
archivos_label.grid(column=1, row=2, padx=10, pady=10, ipady=3, sticky="W")

#
# Entrada de texto del nombre de la columna en comun, comienza deshabilitado
# y cambia el estado luego de agregar un archivo.
#
columna_label = Label(window, text="Columna en común:")
columna_label.grid(column=0, row=3, padx=20, pady=10, sticky="E")
columna_entry = Entry(window, width=40, state='disabled')
columna_entry.grid(column=1, row=3, padx=10, pady=10, ipady=3, sticky="W")

#
# Seleccion de formato de salida.
#
salida_label = Label(window, text="Formato de salida:")
salida_label.grid(column=0, row=4, padx=20, pady=10, sticky="E")
salida_option = ttk.Combobox(window, values=[".xlsx",".csv"])
salida_option.current(0)
salida_option.grid(column=1, row=4, padx=10, pady=10, ipady=3, sticky="W")

#
# Boton de unificar, comienza deshabilitado y cambia el estado
# luego de agregar un archivo.
#
unificar_button = Button(window, text="Unificar listas", command=unificar, state='disabled')
unificar_button.grid(column=1, row=4, padx=10, pady=10, sticky="E")

#
# Footer
#
Label(window, height=3).grid(column=0, row=5, columnspan=2)
footer = Label(window, text="Desarrollado por Diego Ledesma | diego.a.ledesma@gmail.com", font=("Helvetica",9))
footer.grid(column=0, row=6, columnspan=2, ipady=20)

window.mainloop()
