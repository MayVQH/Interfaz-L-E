# -*- coding: utf-8 -*-
"""
Created on Mon Aug 19 11:09:47 2024

@author: Mayra Herrera
"""

import tkinter as tk
import pandas as pd
from collections import Counter


class VentanaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ventana de Usuario")
        self.geometry("500x600")
        
        #Importación del archivo excel para utilizar sus datos
        try:
            url = 'C:/Users/Mayra Herrera/OneDrive - Laboratorios Vanquish SA de CV/Documentos/PROCEDIMIENTOS.xlsx'
            self.dfs = pd.read_excel(url, sheet_name=None, header=None)
        except Exception as e:
            print(f'Error al cargar el archivo de {url} verifique la direccion: {str(e)}')
        #Arreglar el excel para su uso   
        self.df= self.dfs
        self.hojaProc = self.df['PROCEDIMIENTOS']
        self.hojaProc.columns=self.hojaProc.iloc[1]
        self.hojaProc = self.hojaProc[2:]
        self.hojaProc.reset_index(drop=True, inplace=True)
        
        
        #Etiqueta de Pagina de Inicio
        self.labelLicitacion = tk.Label(self,text='HOME')
        self.labelLicitacion.place(x=240,y=10)
        
        #Boton para ir a la siguiente ventana
        self.botonAbrirVentana = tk.Button(self, text="Continuar",command=self.abrirVentana2)
        self.botonAbrirVentana.place(x=200,y=550)
        
        #Variable para el tipo de usuario (publico/distribuidor)
        self.usuario = tk.StringVar(self,"desactivado")
        #Elección de tipo de usuario para que solo se eliga una sola opción
        #Usuario publico
        self.usuarioPublico = tk.Radiobutton(self,text="Público",variable=self.usuario,
                                             value="publico",command=self.verificar_estado)
        self.usuarioPublico.place(x=100,y=50)
        #Usuario distribuidor
        self.usuarioDist = tk.Radiobutton(self,text="Distribuidor",variable=self.usuario,
                                          value="distribuidor",command=self.verificar_estado)
        self.usuarioDist.place(x=290,y=50)
        
        #Etiqueta para el id sistema
        self.labelId = tk.Label(self,text='Id Sistema')
        self.labelId.place(x=20,y=130)
        #Variable para el id sistema
        self.idSist=tk.StringVar(self,'')
        #Entrada de texto para que el usuario ponga el procedimiento o licitacion
        self.entradaId = tk.Entry(self, width =40, textvariable=self.idSist,
                                          state=tk.DISABLED)
        self.entradaId.place(x=175,y=130)
        
        #Etiqueta para el procedimiento o licitacion
        self.labelLicitacion = tk.Label(self,text='Procedimiento/Licitacion')
        self.labelLicitacion.place(x=20,y=170)
        #Variable para el proc/licitacion
        self.procedimiento=tk.StringVar(self,'')
        #Entrada de texto para que el usuario ponga el procedimiento o licitacion
        self.entradaLicitacion = tk.Entry(self, width =40, textvariable=self.procedimiento,
                                          state=tk.DISABLED)
        self.entradaLicitacion.place(x=175,y=170)
        
        #Variable para la elección de busqueda
        self.eleccion = tk.StringVar(self,'desactivado')
        
        #Radiobutton para que el usuario publico eliga si id sistema o licitacion para ingresar
        #Eleccion Id sistema
        self.eleccionId = tk.Radiobutton(self,text="",variable=self.eleccion,
                                             value="id",command=self.verificar_eleccion)
        self.eleccionId.place(x=450,y=130)
        #Eleccion licitacion 
        self.eleccionLicitacion = tk.Radiobutton(self,text="",variable=self.eleccion,
                                             value="licitacion",command=self.verificar_eleccion)
        self.eleccionLicitacion.place(x=450,y=170)
        
        #Entrada de lista para que el usuario vea las claves de su procedimiento escrito
        #y eliga una de ellas
        self.listaClaves = tk.Listbox(self, width=30, height=4)
        self.listaClaves.place(x=175, y=210)
        # Variable para almacenar la clave seleccionada
        self.claveSeleccionada = tk.StringVar()
        
        # Botón para mostrar claves únicas
        self.MostrarClaves = tk.Button(self, text="Mostrar Claves", command=self.mostrarClaves)
        self.MostrarClaves.place(x=370, y=210)
          
        #Etiqueta para la cantidad
        self.labelCantidad = tk.Label(self,text='Cantidad')
        self.labelCantidad.place(x=20,y=290)
        #Variable para la cantidad de medicamento
        self.cantidad=tk.StringVar(self,'0')
        #Entrada de texto para que el usuario coloque la cantidad que desea
        self.entradaCantidad = tk.Entry(self, width =40, textvariable=self.cantidad,
                                          state=tk.DISABLED)
        self.entradaCantidad.place(x=175,y=290)
        
        #Etiqueta para el precio
        self.labelPrecio = tk.Label(self,text='Precio')
        self.labelPrecio.place(x=20,y=330)
        #Variable para el precio
        self.precio=tk.StringVar(self,'0.0')
        #Entrada de texto para que el usuario coloque el precio en caso de ser distribuidor
        self.entradaPrecio = tk.Entry(self, width =40, textvariable=self.precio,
                                          state=tk.DISABLED)
        self.entradaPrecio.place(x=175,y=330)
        
        #Etiqueta de Mes
        self.labelMes = tk.Label(self,text='Mes/Meses')
        self.labelMes.place(x=20,y=360)
        #Variable de mes
        self.mes=tk.StringVar(self,'ENERO')
        #Entrada de texto para que el usuario coloque el mes que desea en caso de ser único
        self.entradaMes = tk.Entry(self, width =40, textvariable=self.mes,
                                          state=tk.DISABLED)
        self.entradaMes.place(x=175,y=360)
        
        #Variable para elegir si es todo el año o no con opción a todo el año o solo un mes
        self.mesCompleto = tk.StringVar(self,"todos")
        #Todo el año
        self.mesesCompletos = tk.Radiobutton(self,text="Todo el año",variable=self.mesCompleto,
                                             value="completo",command=self.verificar_estado)
        self.mesesCompletos.place(x=200,y=400)
        #Meses únicos
        self.mesesUnico = tk.Radiobutton(self,text="Un solo mes",variable=self.mesCompleto,
                                             value="unico",command=self.verificar_estado)
        self.mesesUnico.place(x=320,y=400)
        
        # Botón para mostrar el resumen de la entrada de datos
        self.MostrarResumen = tk.Button(self, text="Mostrar Resumen", command=self.mostrarResumen)
        self.MostrarResumen.place(x=30, y=440)
        
        
        
    def verificar_estado(self):
        # Verificar el estado de los Checkbuttons
        if self.usuario.get()=='publico':
            print("El usuario es público")
            #Habilitar el entry para licitacion, cantidad y mes
            self.entradaId.config(state=tk.NORMAL)
            self.entradaLicitacion.config(state=tk.NORMAL)
            self.entradaCantidad.config(state=tk.NORMAL)
            self.entradaPrecio.config(state=tk.DISABLED)
            self.entradaMes.config(state=tk.NORMAL)
            
            #Verificar si se hizo la elección de todo el año
            if self.mesCompleto.get() == 'completo':
                print('Se entiende que todo el año')
                #Deshabilitar el entry para mes
                self.entradaMes.config(state=tk.DISABLED)
                self.mes = 'todos'
            else:
                print('Se entiende que un solo mes')
                self.entradaMes.config(state=tk.NORMAL)
                print(self.mes)
                
        elif self.usuario.get() == 'distribuidor':
            print("El usuario es distribuidor")
            #Habilitar el entry para licitacion, cantidad, precio y mes
            self.entradaId.config(state=tk.NORMAL)
            self.entradaLicitacion.config(state=tk.DISABLED)
            self.entradaCantidad.config(state=tk.NORMAL)
            self.entradaPrecio.config(state=tk.NORMAL)
            self.entradaMes.config(state=tk.NORMAL)

            #Verificar si se hizo la elección de todo el año
            if self.mesCompleto.get() == 'completo':
                print('Se entiende que todo el año')
                #Deshabilitar el entry para mes
                self.entradaMes.config(state=tk.DISABLED)
                self.mes = 'todos'
            else:
                print('Se entiende que un solo mes')
                self.entradaMes.config(state=tk.NORMAL)
                print(self.mes)
                
        else:
            print("El usuario no ha activado en ninguna categoría")
    
    def verificar_eleccion(self):
        # Verificar el estado de los Checkbuttons
        if self.eleccion.get()=='id':
            print("la busqueda es por id sistema")
            #Deshabilitar el entry para licitacion
            self.entradaLicitacion.config(state=tk.DISABLED)
            self.entradaId.config(state=tk.NORMAL)

        elif self.eleccion.get() == 'licitacion':
            print("la busqeuda es por licitacion")
            #Deshabilitar el entry para id
            self.entradaId.config(state=tk.DISABLED) 
            self.entradaLicitacion.config(state=tk.NORMAL)
        else:
            print("El usuario no ha activado en ninguna categoría")
    
    
    #Realizar la validación de datos ingresados por el usuario
    def verificarDatos(self):
        #Definimos los meses permitidos
        meses=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO',
               'SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
        
        #Obtenemos los valores de cada entry y los guardamos en variables locales
        cantidad_str = self.cantidad.get()
        precio_str = self.precio.get()
        self.proc_str = self.procedimiento.get()
        self.id_str = self.idSist.get()
        
        #Verificamos si la cantidad es un entero
        try:
            self.cantidadEntero = int(cantidad_str)
        except ValueError:
            tk.messagebox.showwarning(title=None, message='El tipo de dato en Cantidad es erróneo')
            return
        
        #Verificamos si el precio es un numero flotante o entero
        try:
            self.precioFlotante = float(precio_str)
        except ValueError:
            tk.messagebox.showwarning(title=None, message='El tipo de dato en Precio es erróneo')
            return
        
        #Verificamos si el mes se enuentra bien escrito y esta dentro de los valores definidos
        if self.mesCompleto.get()=='completo':
            self.mes_str = self.mes
        else:
            self.mes_str = self.mes.get().strip().upper()
            if self.mes_str not in meses:
                tk.messagebox.showwarning(title=None, message='Por favor escriba bien el mes')
                return
            
        # Obtener la clave seleccionada
        seleccion = self.listaClaves.curselection()
        #Verificar que si se ha seleccionado una clave
        if not seleccion:
            tk.messagebox.showwarning(title=None, message='Por favor seleccione una clave o licitacion.')
            return
        
        #Obtener la clave seleccionada 
        self.clave = self.listaClaves.get(seleccion[0])
        print(self.clave)
        

        #Obtener el precio en caso de usuario publico
        if self.usuario.get() == 'publico':
            if self.eleccion.get()=='id':
                self.precioFlotante= self.obtenerPrecio(self.clave)
                print(self.precioFlotante)
            else:
                self.precioFlotante= self.obtenerPrecio(self.clave)
                print(self.precioFlotante)
        
    
    #Función para mostrar claves del procedimiento ingresado
    def mostrarClaves(self):
        # Obtener el procedimiento ingresado
        procedimiento = self.procedimiento.get()
        idSistema = self.idSist.get()
        
        #Verificar que se utilizo para la busqueda si Id o licitacion
        if self.eleccion.get()=='id':
            if not idSistema:
                tk.messagebox.showwarning(title=None, message='Por favor ingrese una clave.')
                return
            licitacionEncontradas = self.obtenerLicitaciones(idSistema)
            self.listaClaves.delete(0, tk.END)
            
            # Agregar las claves a la lista
            for clave in licitacionEncontradas:
                self.listaClaves.insert(tk.END, clave)
        else:
            if not procedimiento:
                tk.messagebox.showwarning(title=None, message='Por favor ingrese un procedimiento.')
                return
        
            claves = self.obtenerClavesUnicas(procedimiento)
            
            self.listaClaves.delete(0, tk.END)
            
            # Agregar las claves a la lista
            for clave in claves:
                self.listaClaves.insert(tk.END, clave)
            
    def obtenerLicitaciones(self, idSistema):
        
        
        self.filasProc = self.hojaProc[self.hojaProc['Clave'] == idSistema]
        ProcClav = self.filasProc['N° Procedimiento'].tolist()
        ProcClav = set(ProcClav)

        return ProcClav
    
    def obtenerClavesUnicas(self, procedimiento):
        
        
        self.filasProc = self.hojaProc[self.hojaProc['N° Procedimiento'] == procedimiento]
        clavesProc = self.filasProc['Clave'].tolist()
        clavesProc = set(clavesProc)

        return clavesProc
    
    def obtenerPrecio(self, clave):
        
        if self.eleccion.get()=='id':
            procedimiento = self.idSist.get()
            self.filasClav = self.hojaProc[(self.hojaProc['Clave'] == procedimiento) & (self.hojaProc['N° Procedimiento']==clave)]
            precioClav = self.filasClav['P.U.'].tolist()
            contador = Counter(precioClav)
            precioClav,frecuencia = contador.most_common(1)[0]
        else:
            procedimiento = self.procedimiento.get()
            self.filasClav = self.hojaProc[(self.hojaProc['Clave'] == clave) & (self.hojaProc['N° Procedimiento']==procedimiento)]
            precioClav = self.filasClav['P.U.'].tolist()
            contador = Counter(precioClav)
            precioClav,frecuencia = contador.most_common(1)[0]

        return precioClav
    
    def mostrarResumen(self):
        
        self.verificarDatos()
        
        # Crear un Label para el resumen dependiendo la eleccion e busqueda
        if self.eleccion.get()=='id': 
            self.resumenVariables = tk.Label(self, 
                                             text=f"Resumen de Ingreso: \nProcedimiento:{self.clave} \n Clave:{self.id_str} \t Cantidad:{self.cantidadEntero} \n Precio:${self.precioFlotante} \t Mes:{self.mes_str}")
            self.resumenVariables.place(x=200,y=440)
        else:
            self.resumenVariables = tk.Label(self, 
                                             text=f"Resumen de Ingreso: \nProcedimiento:{self.proc_str} \n Clave:{self.clave} \t Cantidad:{self.cantidadEntero} \n Precio:${self.precioFlotante} \t Mes:{self.mes_str}")
            self.resumenVariables.place(x=200,y=440)
    
    #Función para abrir una segunda ventana y esconder la primera
    def abrirVentana2(self):
        self.verificarDatos()
        self.ventana2 = VentanaSecundaria(self)
        self.withdraw()
        
class VentanaSecundaria(tk.Toplevel):
    def __init__(self,VentanaPrincipal):
        super().__init__()
        self.title("Ventana Secundaria")
        self.geometry("500x600")
        self.Ventana_Principal = VentanaPrincipal
        
        self.botonVolver = tk.Button(self,text="Volver a inicio", command=self.volver)
        self.botonVolver.place(x=180,y=500)
        
        self.entradaLicitacion = tk.Entry()
        self.entradaLicitacion = tk.Entry(state=tk.DISABLED)
        self.entradaLicitacion.place(x=70,y=170)
        
        print(self.Ventana_Principal.cantidadEntero)
        print(self.Ventana_Principal.precioFlotante)
        print(self.Ventana_Principal.mes_str)
        print(self.Ventana_Principal.proc_str)
        print(self.Ventana_Principal.clave)
        print(self.Ventana_Principal.usuario.get())
        
        
    def volver(self):
        self.destroy()
        self.Ventana_Principal.deiconify()

        
if __name__ == "__main__":
    app = VentanaPrincipal()
    app.mainloop()