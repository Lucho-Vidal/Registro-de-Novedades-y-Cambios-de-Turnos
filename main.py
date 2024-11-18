import tkinter as tk
from tkinter import ttk
from tkinter import messagebox,scrolledtext  
from datetime import datetime
from ttkbootstrap import DateEntry
from ttkbootstrap import Style
import openpyxl
import os

class FormularioExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1110x620')  # Ajuste inicial de la ventana
        self.root.title("Registro de Novedades y Cambios de turnos TK")
        # root.state('zoomed')
        # Configuración del archivo de Excel
        try:
            with open("path_base", "r", encoding="utf-8") as file:
                self.excel_file = file.read().strip()
                self.excel_file = self.excel_file.replace("\\", "\\\\")  # Reemplazar '\' por '\\'
        except Exception as e:
            print(f"Error leyendo el archivo: {e}")
            self.excel_file = r'PLANILLA NOVEDADES PERSONAL ABORDO.xlsx'

        # Verificar si el archivo existe; si no, crearlo
        if not os.path.exists(self.excel_file):
            self.crear_archivo_excel()

        # Cargar el libro de Excel y las hojas
        self.wb = openpyxl.load_workbook(self.excel_file)
        self.sheet_base = self.wb["BASE"]
        self.sheet_novedades = self.wb["NOVEDADES"]
        self.sheet_tipo_novedad = self.wb["TipoNovedad"]
        self.sheet_cambio_turnos = self.wb["Cambio de Turnos"]
        
        # Cargar opciones de tipo de novedad
        self.tipo_novedades = []
        self.cargarTipoNovedades()
        
        # Aplicar el estilo ttkbootstrap
        style = Style()
        style.theme_use('superhero')
        
        # Variables del formulario
        self.legajo_var = tk.StringVar()
        self.apellidos_nombres_var = tk.StringVar()
        self.especialidad_var = tk.StringVar()
        self.dotacion_var = tk.StringVar()
        self.turnos_var = tk.StringVar()
        self.franco_var = tk.StringVar()
        self.novedad_var = tk.StringVar()
        self.fecha_inicio_novedad_var = tk.StringVar()
        self.fecha_fin_novedad_var = tk.StringVar()
        self.referencia_estacion_var = tk.StringVar()
        self.supervisor_var = tk.StringVar()
        self.observaciones_var = tk.StringVar()
        
        self.legajo_2_var = tk.StringVar()
        self.apellidos_nombres_2_var = tk.StringVar()
        self.especialidad_2_var = tk.StringVar()
        self.dotacion_2_var = tk.StringVar()
        self.turnos_2_var = tk.StringVar()
        self.franco_2_var = tk.StringVar()
        self.fecha_cambio_turno_var = tk.StringVar()

        # Marco principal
        main_frame = tk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # Crear marcos de formulario y tabla
        self.form_frame = tk.Frame(main_frame)
        self.form_cambios_frame = tk.Frame(main_frame)
        self.table_cambios_frame = tk.Frame(main_frame)
        self.table_frame = tk.Frame(main_frame)
        self.table_frame.grid(row=0, column=0, padx=10, pady=10)  # Muestra la tabla desde el inicio

        # Configurar el grid de la ventana principal para que el marco se expanda 
        self.root.grid_rowconfigure(0, weight=1) 
        self.root.grid_columnconfigure(0, weight=1)
        
        self.crear_tabla_novedades()

    def toggle_view(self, target_view=None):
        """
        Cambia a la vista especificada. Si no se proporciona 'target_view',
        alterna entre las vistas según el estado actual.
        """

        # Ocultar todas las vistas
        self.form_frame.grid_forget()
        self.table_frame.grid_forget()
        self.form_cambios_frame.grid_forget()
        self.table_cambios_frame.grid_forget()

        # Determinar la vista objetivo
        self.current_view = target_view  # Si se especifica un destino, cambiar a esa vista

        # Mostrar la vista correspondiente
        if self.current_view == "form":
            self.form_frame.grid(row=0, column=0, padx=10, pady=10)
            self.mostrar_formulario_novedades()
        elif self.current_view == "form_cambios":
            self.form_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            self.mostrar_formulario_cambios()
        elif self.current_view == "table_cambios":
            self.table_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            self.crear_tabla_cambios()
        else:
            self.table_frame.grid(row=0, column=0, padx=10, pady=10)
            self.crear_tabla_novedades()
    
    def cargarTipoNovedades(self):
        for row in self.sheet_tipo_novedad.iter_rows(min_row=2, values_only=True):
            tipo_novedad = row[0]  # Suponemos que los tipos de novedad están en la primera columna
            self.tipo_novedades.append(tipo_novedad)

    def crear_archivo_excel(self):
        """Crea un archivo Excel con encabezados si no existe"""
        wb = openpyxl.Workbook()
        # Crear la hoja "BASE" y agregar encabezados
        sheet_base = wb.create_sheet(title="BASE")  # Crea la hoja "BASE"
        encabezados_base = [
            "LEGAJO SAP", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", "TURNOS", "FRANCO"
        ]
        sheet_base.append(encabezados_base)  # Agregar los encabezados a la hoja "BASE"
        
        # Crear la hoja "NOVEDADES" y agregar encabezados
        sheet_novedades = wb.create_sheet(title="NOVEDADES")  # Crea la hoja "NOVEDADES"
        encabezados_novedades = [
            "LEGAJO SAP", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", "TURNOS", "FRANCO", 
            "NOVEDAD", "Fecha de Inicio Novedad","Fecha de Inicio Novedad", "REFERENCIA ESTACIÓN", "SUPERVISOR", "Observaciones"
        ]
        sheet_novedades.append(encabezados_novedades)  # Agregar los encabezados a la hoja "NOVEDADES"
        
        # Crear la hoja "NOVEDADES" y agregar encabezados
        sheet_novedades = wb.create_sheet(title="TipoNovedad")  # Crea la hoja "NOVEDADES"
        encabezados_novedades = ["Enfermo"]
        sheet_novedades.append(encabezados_novedades)  # Agregar los encabezados a la hoja "NOVEDADES"
        
        # Crear la hoja "Cambio de Turnos" y agregar encabezados
        sheet_Cambio_de_Turnos = wb.create_sheet(title="Cambio de Turnos")  # Crea la hoja "NOVEDADES"
        encabezados_novedades = [
            "ID","LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", "TURNOS", "FRANCO", 
            "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2", "TURNOS2", "FRANCO2", 
            "Fecha de Cambio de Turno", "REFERENCIA ESTACIÓN", "SUPERVISOR", "Observaciones"
        ]
        
        sheet_novedades.append(sheet_Cambio_de_Turnos)  # Agregar los encabezados a la hoja "NOVEDADES"
        
        # Eliminar la hoja predeterminada (por defecto, openpyxl crea una hoja vacía llamada "Sheet")
        del wb["Sheet"]
        
        # Guardar el archivo
        wb.save(self.excel_file)
        print(f"Archivo creado: {self.excel_file}")
        
    def crear_tabla_novedades(self):
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD",
            "DOTACION", "TURNOS", "FRANCO", "NOVEDAD", "Fecha de Inicio Novedad", "Fecha de Fin Novedad",
            "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones"
        ]
        # Crear un título en la vista del formulario
        ttk.Label(self.table_frame, text="Registro de novedades", font=("Helvetica", 20, "bold")).grid(row=0, column=0, pady=10, padx=10, sticky="nw")
        
        # Botón para ir al formulario
        ttk.Button(self.table_frame, text="Ver cambios de turno", command=lambda: self.toggle_view("table_cambios")).grid(row=0, column=2, pady=10, padx=10, sticky="e")
        ttk.Button(self.table_frame, text="Nueva novedad", command=lambda: self.toggle_view("form")).grid(row=0, column=3, pady=10, padx=10, sticky="w")
        ttk.Button(self.table_frame, text="Nuevo cambio de turno", command=lambda: self.toggle_view("form_cambios")).grid(row=0, column=4, pady=10, padx=10, sticky="w")


        # Crear el Treeview dentro de un marco contenedor para el scroll 
        self.tree_frame = ttk.Frame(self.table_frame, width=1100, height=550)
        self.tree_frame.grid(row=1, column=0, columnspan=5, sticky="nsew")
        self.tree_frame.grid_propagate(False)

        # Configurar el grid del contenedor (self.table_frame) para que permita la expansión
        self.table_frame.grid_rowconfigure(1, weight=1)  # Fila 1 (donde está el tree_frame) se expande
        self.table_frame.grid_columnconfigure(0, weight=1)  # Configurar la primera columna para que se expanda
        self.table_frame.grid_columnconfigure(1, weight=1)  # Configurar la segunda columna
        self.table_frame.grid_columnconfigure(2, weight=1)  # Configurar la tercera columna

        # Configurar el grid de self.table_frame para permitir expansión
        self.table_frame.grid_rowconfigure(1, weight=1)  # Fila de la tabla se expande
        self.table_frame.grid_columnconfigure(0, weight=1)  # Permitir expansión horizontal

        # Configurar el grid de tree_frame para que el Treeview ocupe todo el espacio
        self.tree_frame.grid_rowconfigure(0, weight=1)  # Fila 0 del tree_frame se expande
        self.tree_frame.grid_columnconfigure(0, weight=1)  # Columna 0 del tree_frame se expande

        # Crear el Treeview
        self.tabla_novedades = ttk.Treeview(self.tree_frame, columns=columnas, show="headings", height=40)
        self.tabla_novedades.grid(row=0, column=0, sticky="nsew")

        # Establecer encabezados y anchos de columna
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 120, 90, 120, 120, 120]
        for col, ancho in zip(columnas, anchuras):
            self.tabla_novedades.heading(col, text=col)
            self.tabla_novedades.column(col, width=ancho)

        # Crear y configurar la barra de scroll vertical
        scrollbar_vertical = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tabla_novedades.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")

        # Crear y configurar la barra de scroll horizontal
        scrollbar_horizontal = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tabla_novedades.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

        # Configurar el Treeview para usar las barras de scroll
        self.tabla_novedades.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

        # Cargar los datos de `self.sheet_novedades` en la tabla
        for fila in self.sheet_novedades.iter_rows(min_row=2, values_only=True):  # min_row=2 para omitir encabezados
            # Reemplazar valores None por "-"
            fila_procesada = ["-" if celda is None else celda for celda in fila]
            self.tabla_novedades.insert("", "end", values=fila_procesada)

        
    def crear_tabla_cambios(self):
        
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", 
            "TURNOS", "FRANCO", "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2", 
            "TURNOS2", "FRANCO2", "Fecha de Cambio de Turno", "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones"
        ]
        # Crear un título en la vista del formulario
        ttk.Label(self.table_cambios_frame, text="Registro de cambios de turnos", font=("Helvetica", 20, "bold")).grid(row=0, column=0, pady=10, padx=10, sticky="nw")
        
        # Botón para ir al formulario
        ttk.Button(self.table_cambios_frame, text="Ver novedades", command=lambda: self.toggle_view("table")).grid(row=0, column=2, pady=10,padx=1, sticky="e")
        ttk.Button(self.table_cambios_frame, text="Nueva novedad", command=lambda: self.toggle_view("form")).grid(row=0, column=3, pady=10, padx=1, sticky="e")
        ttk.Button(self.table_cambios_frame, text="Nuevo cambio de turno", command=lambda: self.toggle_view("form_cambios")).grid(row=0, column=4, pady=10,padx=1,  sticky="e")
        
        # Crear el Treeview dentro de un marco contenedor para el scroll 
        self.tree_frame = ttk.Frame(self.table_cambios_frame, width=1100, height=550)
        self.tree_frame.grid(row=1, column=0, columnspan=5, sticky="nsew")
        self.tree_frame.grid_propagate(False)

        # Configurar el grid del marco contenedor para que el Treeview se expanda
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.tabla_novedades = ttk.Treeview(self.tree_frame, columns=columnas, show="headings", height=38)
        self.tabla_novedades.grid(row=0, column=0, sticky="nsew")

        # Establecer encabezados y columnas con anchuras
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 60, 150, 150, 80, 60, 80, 120, 120, 120, 120]
        for col, ancho in zip(columnas, anchuras):
            self.tabla_novedades.heading(col, text=col)
            self.tabla_novedades.column(col, width=ancho, anchor='center')

        # Crear y configurar la barra de scroll vertical
        scrollbar_vertical = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tabla_novedades.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")

        # Crear y configurar la barra de scroll horizontal
        scrollbar_horizontal = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tabla_novedades.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

        # Configurar el Treeview para usar las barras de scroll
        self.tabla_novedades.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

        # Cargar los datos de `self.sheet_cambio_turnos` en la tabla
        for fila in self.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):  # min_row=2 para omitir encabezados
            fila_procesada = ["-" if celda is None else celda for celda in fila]
            self.tabla_novedades.insert("", "end", values=fila_procesada)

    

    def mostrar_formulario_novedades(self):
        
        ttk.Label(self.form_frame, text="Formulario de novedades", font=("Helvetica", 20, "bold")).grid(row=0, column=0,columnspan=3, pady=10, padx=10, sticky="nw")
        
        # Nivel 1: LEGAJO SAP
        ttk.Label(self.form_frame, text="   Legajo").grid(row=1, column=0, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.legajo_entry = ttk.Entry(self.form_frame, textvariable=self.legajo_var, width=10)
        self.legajo_entry.grid(row=2, column=0, sticky="ew", pady=5)
        self.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        # Botón para buscar y auto completar
        ttk.Button(self.form_frame, text="Buscar Personal", command=self.mostrar_modal).grid(row=2, column=1, pady=10, padx=10)

        # Nivel 2: APELLIDOS Y NOMBRES, ESPECIALIDAD, DOTACION, TURNOS, FRANCO

        # Apellidos y Nombres
        ttk.Label(self.form_frame, text="  Apellido y Nombre").grid(row=3, column=0, sticky="w", padx=5)
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=3, column=0, sticky="w")
        self.apellidos_nombres_entry = ttk.Entry(self.form_frame, textvariable=self.apellidos_nombres_var, state='disabled', width=40)
        self.apellidos_nombres_entry.grid(row=4, column=0, sticky="w", pady=5)

        # Especialidad
        ttk.Label(self.form_frame, text="Especialidad").grid(row=3, column=1, sticky="w", padx=5)
        self.especialidad_entry = ttk.Entry(self.form_frame, textvariable=self.especialidad_var, state='disabled', width=20)
        self.especialidad_entry.grid(row=4, column=1, sticky="w", pady=5, padx=8)

        # Dotación
        ttk.Label(self.form_frame, text="Dotación").grid(row=3, column=2, sticky="w", padx=5)
        self.dotacion_entry = ttk.Entry(self.form_frame, textvariable=self.dotacion_var, state='disabled', width=20)
        self.dotacion_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        # Turnos
        ttk.Label(self.form_frame, text="Turno").grid(row=3, column=3, sticky="w", padx=5)
        self.turnos_entry = ttk.Entry(self.form_frame, textvariable=self.turnos_var, state='disabled', width=40)
        self.turnos_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        # Franco
        ttk.Label(self.form_frame, text="Franco").grid(row=3, column=4, sticky="w", padx=5)
        self.franco_entry = ttk.Entry(self.form_frame, textvariable=self.franco_var, state='disabled', width=20)
        self.franco_entry.grid(row=4, column=4, sticky="w", pady=5, padx=8)

        # Nivel 3: NOVEDAD, Fecha de Inicio Novedad
        ttk.Label(self.form_frame, text="   Tipo Novedad").grid(row=5, column=0, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=0, sticky="w")
        self.novedad_entry = ttk.Combobox(self.form_frame, textvariable=self.novedad_var, values=self.tipo_novedades, width=38)
        self.novedad_entry.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.form_frame, text="   Fecha de Inicio Novedad").grid(row=5, column=3, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=3, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_inicio_novedad_entry = DateEntry(self.form_frame,dateformat='%d/%m/%Y')
        self.fecha_inicio_novedad_entry.grid(row=6, column=3, sticky="w", pady=5)       
        
        ttk.Label(self.form_frame, text="Fecha de Fin Novedad").grid(row=5, column=4, sticky="w")
        # ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=4, column=3, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_fin_novedad_entry = DateEntry(self.form_frame,dateformat='%d/%m/%Y')
        self.fecha_fin_novedad_entry.grid(row=6, column=4, sticky="w", pady=5) 
        
        # Nivel 4: REFERENCIA ESTACIÓN, SUPERVISOR
        ttk.Label(self.form_frame, text="   REFERENCIA ESTACIÓN").grid(row=7, column=0, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=0, sticky="w")
        self.referencia_estacion_entry = ttk.Entry(self.form_frame, textvariable=self.referencia_estacion_var, width=40)
        self.referencia_estacion_entry.grid(row=8, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.form_frame, text="   SUPERVISOR").grid(row=7, column=3, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=3, sticky="w")
        self.supervisor_entry = ttk.Entry(self.form_frame, textvariable=self.supervisor_var, width=40)
        self.supervisor_entry.grid(row=8, column=3, columnspan=2, sticky="w", pady=5)

        # Nivel 5: Observaciones
        ttk.Label(self.form_frame, text="Observaciones").grid(row=9, column=0, sticky="w")
        self.observaciones_text = scrolledtext.ScrolledText(self.form_frame, wrap=tk.WORD, height=12,width=180)
        self.observaciones_text.grid(row=10, column=0, columnspan=7, pady=5, sticky="w")

        # Botón de Guardar Datos en el formulario
        ttk.Button(self.form_frame, text="Guardar Novedad", command=self.guardar_datos_novedades).grid(row=11, column=3, columnspan=2, pady=10)

        # Botón para cerrar el formulario y regresar a la tabla
        ttk.Button(self.form_frame, text="Cerrar", command=self.toggle_view).grid(row=11, column=4, columnspan=2, pady=10)

        
    def mostrar_formulario_cambios(self):

        ttk.Label(self.form_cambios_frame, text="Formulario de cambios de turnos", font=("Helvetica", 20, "bold")).grid(row=0, column=0,columnspan=3, pady=10, padx=10, sticky="nw")

        # Nivel 1: LEGAJO SAP
        ttk.Label(self.form_cambios_frame, text="   Legajo").grid(row=1, column=0, sticky="w")# Ejemplo de cómo agregar un asterisco rojo a la etiqueta
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.legajo_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.legajo_var, width=10)
        self.legajo_entry.grid(row=2, column=0, sticky="w", pady=5)
        self.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        # Botón para buscar y auto completar
        ttk.Button(self.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(1)).grid(row=2, column=1, pady=10, padx=10)

        # Nivel 2: APELLIDOS Y NOMBRES, ESPECIALIDAD, DOTACION, TURNOS, FRANCO

        # Apellidos y Nombres
        ttk.Label(self.form_cambios_frame, text="  Apellido y Nombre").grid(row=3, column=0, sticky="w", padx=5)
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=3, column=0, sticky="w")
        self.apellidos_nombres_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.apellidos_nombres_var, state='disabled', width=40)
        self.apellidos_nombres_entry.grid(row=4, column=0,columnspan = 2, sticky="w", pady=5)

        # Especialidad
        ttk.Label(self.form_cambios_frame, text="Especialidad").grid(row=3, column=2, sticky="w", padx=5)
        self.especialidad_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.especialidad_var, state='disabled', width=20)
        self.especialidad_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        # Dotación
        ttk.Label(self.form_cambios_frame, text="Dotación").grid(row=3, column=3, sticky="w", padx=5)
        self.dotacion_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.dotacion_var, state='disabled', width=20)
        self.dotacion_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        # Turnos
        ttk.Label(self.form_cambios_frame, text="Turno").grid(row=3, column=4, sticky="w", padx=5)
        self.turnos_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.turnos_var, state='disabled', width=40)
        self.turnos_entry.grid(row=4, column=4, columnspan= 2 , sticky="w", pady=5, padx=8)

        # Franco
        ttk.Label(self.form_cambios_frame, text="Franco").grid(row=3, column=6, sticky="w", padx=5)
        self.franco_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.franco_var, state='disabled', width=20)
        self.franco_entry.grid(row=4, column=6, sticky="w", pady=5, padx=8)
        #--------------------------------------------------------------------------------
        # Nivel 3: LEGAJO 2
        ttk.Label(self.form_cambios_frame, text="   Legajo 2").grid(row=5, column=0, sticky="w")# Ejemplo de cómo agregar un asterisco rojo a la etiqueta
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=0, sticky="w")
        self.legajo_2_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.legajo_2_var, width=10)
        self.legajo_2_entry.grid(row=6, column=0, sticky="w", pady=5)
        self.legajo_2_entry.bind("<Return>", lambda event: self.buscar_legajo(2))

        # Botón para buscar y auto completar
        ttk.Button(self.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(2)).grid(row=6, column=1, pady=10, padx=10)

        # Nivel 4: APELLIDOS Y NOMBRES, ESPECIALIDAD, DOTACION, TURNOS, FRANCO
        # Apellidos y Nombres2
        ttk.Label(self.form_cambios_frame, text="  Apellido y Nombre 2").grid(row=7, column=0, sticky="w", padx=5)
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=0, sticky="w")
        self.apellidos_nombres_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.apellidos_nombres_2_var, state='disabled', width=40)
        self.apellidos_nombres_entry_2.grid(row=8, column=0,columnspan = 2, sticky="w", pady=5)

        # Especialidad2
        ttk.Label(self.form_cambios_frame, text="Especialidad 2").grid(row=7, column=2, sticky="w", padx=5)
        self.especialidad_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.especialidad_2_var, state='disabled', width=20)
        self.especialidad_entry_2.grid(row=8, column=2, sticky="w", pady=5, padx=8)

        # Dotación2
        ttk.Label(self.form_cambios_frame, text="Dotación 2").grid(row=7, column=3, sticky="w", padx=5)
        self.dotacion_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.dotacion_2_var, state='disabled', width=20)
        self.dotacion_entry_2.grid(row=8, column=3, sticky="w", pady=5, padx=8)

        # Turnos2
        ttk.Label(self.form_cambios_frame, text="Turno 2").grid(row=7, column=4, sticky="w", padx=5)
        self.turnos_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.turnos_2_var, state='disabled', width=40)
        self.turnos_entry_2.grid(row=8, column=4, columnspan= 2 , sticky="w", pady=5, padx=8)

        # Franco2
        ttk.Label(self.form_cambios_frame, text="Franco 2").grid(row=7, column=6, sticky="w", padx=5)
        self.franco_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.franco_2_var, state='disabled', width=20)
        self.franco_entry_2.grid(row=8, column=6, sticky="w", pady=5, padx=8)
        
        # Nivel 5: Fecha de cambio de turno
        ttk.Label(self.form_cambios_frame, text="   Fecha de cambio de turno").grid(row=9, column=0, sticky="w")
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=9, column=0, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_cambio_turno_entry = DateEntry(self.form_cambios_frame,dateformat='%d/%m/%Y')
        self.fecha_cambio_turno_entry.grid(row=10, column=0, columnspan=2, sticky="w", pady=5)       
        
        # Nivel 6: REFERENCIA ESTACIÓN, SUPERVISOR
        ttk.Label(self.form_cambios_frame, text="   REFERENCIA ESTACIÓN").grid(row=11, column=0, sticky="w")
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=11, column=0, sticky="w")
        self.referencia_estacion_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.referencia_estacion_var, width=40)
        self.referencia_estacion_entry.grid(row=12, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.form_cambios_frame, text="   SUPERVISOR").grid(row=11, column=3, sticky="w")
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=11, column=3, sticky="w")
        self.supervisor_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.supervisor_var, width=40)
        self.supervisor_entry.grid(row=12, column=3, columnspan=2, sticky="w", pady=5)

        # Nivel 7: Observaciones
        ttk.Label(self.form_cambios_frame, text="Observaciones").grid(row=13, column=0, sticky="w")
        self.observaciones_text = scrolledtext.ScrolledText(self.form_cambios_frame, wrap=tk.WORD, height=1,width=170)
        self.observaciones_text.grid(row=14, column=0, columnspan=7, pady=5, sticky="w")

        # Botón de Guardar Datos en el formulario
        ttk.Button(self.form_cambios_frame, text="Guardar cambio de turno", command=self.guardar_datos_cambios).grid(row=15, column=2, columnspan=2, pady=10)

        # Botón para cerrar el formulario y regresar a la tabla
        ttk.Button(self.form_cambios_frame, text="Cerrar", command=lambda: self.toggle_view("table_cambios")).grid(row=15, column=3, columnspan=2, pady=10)

        # self.form_cambios_frame.grid(row=1, column=0, padx=10, pady=10)  # Mostrar formulario
            
    def mostrar_modal(self,boton=1):
        """Mostrar el modal con el Treeview y filtro"""
        # Crear una ventana secundaria (modal)
        modal = tk.Toplevel(self.root)
        modal.title("Seleccionar Legajo")
        
        # Crear un campo de entrada para filtrar por apellido
        tk.Label(modal, text="Filtrar por Apellido:").grid(row=0, column=0, padx=10, pady=5)
        apellido_filter_var = tk.StringVar()
        apellido_filter = tk.Entry(modal, textvariable=apellido_filter_var)
        apellido_filter.grid(row=0, column=1, padx=10, pady=5)
        
        # Crear el Treeview para mostrar la tabla
        tree = ttk.Treeview(modal, columns=("legajo", "apellido_nombre", "especialidad", "dotacion", "turnos", "franco"), show="headings")
        tree.heading("legajo", text="LEGAJO SAP")
        tree.heading("apellido_nombre", text="APELLIDOS Y NOMBRES")
        tree.heading("especialidad", text="ESPECIALIDAD")
        tree.heading("dotacion", text="DOTACION")
        tree.heading("turnos", text="TURNOS")
        tree.heading("franco", text="FRANCO")
        tree.grid(row=1, column=0, columnspan=6, padx=20, pady=10)

        # Función para llenar la tabla con los datos de la hoja 'BASE'
        def cargar_tabla():
            for row in tree.get_children():
                tree.delete(row)
            
            apellido_filter = apellido_filter_var.get().lower()  # Obtener el filtro (en minúsculas)
            
            # Cargar las filas del Excel y mostrar en el Treeview
            for row in self.sheet_base.iter_rows(min_row=2, values_only=True):
                legajo = row[0]
                apellidos_nombres = row[1]
                turnos = row[2]
                franco = row[3]
                especialidad = row[4]
                dotacion = row[5]

                if apellido_filter in apellidos_nombres.lower():  # Comparar sin tener en cuenta mayúsculas
                    tree.insert("", "end", values=(legajo, apellidos_nombres, especialidad, dotacion, turnos, franco))
        # Llamar a la función para cargar la tabla al inicio
        cargar_tabla()
        # Asociar el evento de filtrar al campo de filtro por apellido
        apellido_filter_var.trace_add("write", lambda *args: cargar_tabla())

        def on_double_click(event):
            # Función para asignar el legajo al campo de entrada y cerrar el modal
            selected_item = tree.selection()[0]  # Obtener el item seleccionado
            legajo_sap = tree.item(selected_item, "values")[0]  # Obtener el legajo
            apellido_nombre = tree.item(selected_item, "values")[1]  
            especialidad = tree.item(selected_item, "values")[2]
            dotacion = tree.item(selected_item, "values")[3]
            turnos = tree.item(selected_item, "values")[4]
            franco = tree.item(selected_item, "values")[5]
            if boton == 1:
                self.legajo_var.set(legajo_sap)  # Asignar al campo "LEGAJO SAP"
                self.apellidos_nombres_var.set(apellido_nombre)  
                self.especialidad_var.set(especialidad)  
                self.dotacion_var.set(dotacion)  
                self.turnos_var.set(turnos)  
                self.franco_var.set(franco)  
            elif boton == 2:
                self.legajo_2_var.set(legajo_sap)  # Asignar al campo "LEGAJO SAP"
                self.apellidos_nombres_2_var.set(apellido_nombre)  
                self.especialidad_2_var.set(especialidad)  
                self.dotacion_2_var.set(dotacion)  
                self.turnos_2_var.set(turnos)  
                self.franco_2_var.set(franco) 
                
            modal.destroy()  # Cerrar el modal

        # Asociar el evento de doble clic
        tree.bind("<Double-1>", on_double_click)

        # Botón de cerrar el modal
        tk.Button(modal, text="Cerrar", command=modal.destroy).grid(row=2, column=0, columnspan=2, pady=10)

    def buscar_legajo(self,campo=1):
        """Buscar el legajo SAP en la hoja 'BASE' y auto completar los datos"""
        try:
            # Convertir el legajo ingresado a número (int)
            if campo == 1:
                legajo = int(self.legajo_var.get().strip())  # Eliminar espacios y convertir a número
            elif campo == 2:
                legajo = int(self.legajo_2_var.get().strip())  # Eliminar espacios y convertir a número
                
            print(f"Buscando legajo SAP: {legajo}")

            # Buscar en la hoja BASE (salta el encabezado en la primera fila)
            for row in self.sheet_base.iter_rows(min_row=2, values_only=True):
                # Convertir el legajo de la hoja a número (int)
                legajo_sap = int(row[0])  
                if legajo_sap == legajo:
                    # Si el legajo SAP coincide, auto completar los campos
                    if campo == 1:
                        self.apellidos_nombres_var.set(row[1])  # Nombre en la segunda columna
                        self.especialidad_var.set(row[2])  # Especialidad
                        self.dotacion_var.set(row[3])  # Dotación
                        self.turnos_var.set(row[4])  # Turnos
                        self.franco_var.set(row[5])  # Franco
                    elif campo == 2:
                        self.apellidos_nombres_2_var.set(row[1])  # Nombre en la segunda columna
                        self.especialidad_2_var.set(row[2])  # Especialidad
                        self.dotacion_2_var.set(row[3])  # Dotación
                        self.turnos_2_var.set(row[4])  # Turnos
                        self.franco_2_var.set(row[5])  # Franco
                    print(f"Legajo encontrado: {legajo_sap}")
                    # self.deshabilitar_campos()  # Deshabilitar los campos después de la búsqueda
                    break  # Salir del bucle una vez encontrado el legajo
            else:
                messagebox.showinfo("Legajo No Encontrado", f"El legajo {legajo} no fue encontrado.")
                print(f"Legajo SAP {legajo} no encontrado.")

        except ValueError:
            # Si el legajo ingresado no es un número, mostrar un mensaje de error
            messagebox.showerror("Error de entrada", "Por favor, ingrese un número de legajo válido.")

    def guardar_datos_novedades(self):
        """Guardar los datos del formulario en el archivo Excel"""
        self.fecha_inicio_novedad_var.set(self.fecha_inicio_novedad_entry.entry.get())
        self.fecha_fin_novedad_var.set(self.fecha_fin_novedad_entry.entry.get())
        self.observaciones_var.set(self.observaciones_text.get("1.0", "end-1c"))

        # Obtener el último ID en la columna 1
        # Obtener el último ID válido en la columna 1, omitiendo el encabezado
        last_id = None
        for row in self.sheet_novedades.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] and isinstance(row[0], int):  # Verifica que es un ID numérico
                last_id = row[0]
        
        # Incrementar el ID en 1
        new_id = int(last_id) + 1 if last_id else 1
        
        # Obtener la fecha y hora actual en el formato especificado
        current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        nuevo_dato = (
            new_id,
            current_datetime,
            self.legajo_var.get(),
            self.apellidos_nombres_var.get(),
            self.especialidad_var.get(),
            self.dotacion_var.get(),
            self.turnos_var.get(),
            self.franco_var.get(),
            self.novedad_var.get(),
            self.fecha_inicio_novedad_var.get(),
            self.fecha_fin_novedad_var.get(),
            self.referencia_estacion_var.get(),
            self.supervisor_var.get(),
            self.observaciones_var.get()
        )
        if self.validar_campos_requeridos_novedades():
            # Aquí guarda los datos si todo está bien
            # Puedes agregar el código para guardar en el archivo de Excel o donde necesites.
            # Añadir a la hoja de Excel y guardar
            try:
                # Añadir los nuevos datos a la hoja
                self.sheet_novedades.append(nuevo_dato)

                # Guardar el archivo Excel
                self.wb.save(self.excel_file)

                # Mensaje de confirmación
                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                # Limpiar el formulario y alternar a la vista de tabla
                self.limpiar_formulario_novedades()
                self.toggle_view()
                print("Datos guardados correctamente.")

            except Exception as e:
                # Mensaje de error si ocurre un problema
                messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)} Por favor, intente de nuevo y si el problema persiste avise al administrador")
                print(f"Error al guardar los datos: {e}")
            
        else:
            print("Algunos campos son obligatorios y están vacíos.")
            messagebox.showinfo("Faltan Datos requeridos", f"Algunos campos son obligatorios y están vacíos.")
    
    def guardar_datos_cambios(self):
        """Guardar los datos del formulario en el archivo Excel"""
        self.fecha_cambio_turno_var.set(self.fecha_cambio_turno_entry.entry.get())
        self.observaciones_var.set(self.observaciones_text.get("1.0", "end-1c"))
        
        # Obtener el último ID en la columna 1
        # Obtener el último ID válido en la columna 1, omitiendo el encabezado
        last_id = None
        for row in self.sheet_cambio_turnos.iter_rows(min_row=2, max_col=1, values_only=True):
            if row[0] and isinstance(row[0], int):  # Verifica que es un ID numérico
                last_id = row[0]
        
        # Incrementar el ID en 1
        new_id = int(last_id) + 1 if last_id else 1
        
        # Obtener la fecha y hora actual en el formato especificado
        current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
        
        nuevo_dato = (
            new_id,
            current_datetime,
            self.legajo_var.get(),
            self.apellidos_nombres_var.get(),
            self.especialidad_var.get(),
            self.dotacion_var.get(),
            self.turnos_var.get(),
            self.franco_var.get(),
            self.legajo_2_var.get(),
            self.apellidos_nombres_2_var.get(),
            self.especialidad_2_var.get(),
            self.dotacion_2_var.get(),
            self.turnos_2_var.get(),
            self.franco_2_var.get(),
            self.fecha_cambio_turno_var.get(),
            self.referencia_estacion_var.get(),
            self.supervisor_var.get(),
            self.observaciones_var.get()
        )
        if self.validar_campos_requeridos_cambios():
            # Aquí guarda los datos si todo está bien
            # Puedes agregar el código para guardar en el archivo de Excel o donde necesites.
            # Añadir a la hoja de Excel y guardar
            try:
                # Añadir los nuevos datos a la hoja
                self.sheet_cambio_turnos.append(nuevo_dato)

                # Guardar el archivo Excel
                self.wb.save(self.excel_file)

                # Mensaje de confirmación
                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                # Limpiar el formulario y alternar a la vista de tabla
                self.limpiar_formulario_cambios()
                self.toggle_view("table_cambios")
                print("Datos guardados correctamente.")

            except Exception as e:
                # Mensaje de error si ocurre un problema
                messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)} Por favor, intente de nuevo y si el problema persiste avise al administrador")
                print(f"Error al guardar los datos: {e}")
        else:
            print("Algunos campos son obligatorios y están vacíos.")
            messagebox.showinfo("Faltan Datos requeridos", f"Algunos campos son obligatorios y están vacíos.")
    def limpiar_formulario_novedades(self):
        self.legajo_var.set('')
        self.apellidos_nombres_var.set('')
        self.especialidad_var.set('')
        self.dotacion_var.set('')
        self.turnos_var.set('')
        self.franco_var.set('')
        self.novedad_var.set('')        
        self.fecha_inicio_novedad_var.set('')
        self.fecha_fin_novedad_var.set('')
        self.referencia_estacion_var.set('')
        self.supervisor_var.set('')
        self.observaciones_var.set('')
        self.observaciones_text.delete("1.0", "end")

    
    def limpiar_formulario_cambios(self):
        self.legajo_var.set('')
        self.apellidos_nombres_var.set('')
        self.especialidad_var.set('')
        self.dotacion_var.set('')
        self.turnos_var.set('')
        self.franco_var.set('')
        self.legajo_2_var.set('')
        self.apellidos_nombres_2_var.set('')
        self.especialidad_2_var.set('')
        self.dotacion_2_var.set('')
        self.turnos_2_var.set('')
        self.franco_2_var.set('')
        self.fecha_cambio_turno_var.set('')
        self.referencia_estacion_var.set('')
        self.supervisor_var.set('')
        self.observaciones_var.set('') 
        self.observaciones_text.delete("1.0", "end")
   
    def validar_campos_requeridos_novedades(self):
            # Verificar si los campos requeridos están vacíos
            if not self.legajo_var.get():
                self.mostrar_error_novedades("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_var.get():
                self.mostrar_error_novedades("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.novedad_var.get():
                self.mostrar_error_novedades("El campo 'Tipo de Novedad' es obligatorio.")
                return False
            if not self.fecha_inicio_novedad_var.get():
                self.mostrar_error_novedades("El campo 'Fecha de inicio novedad' es obligatorio.")
                return False
            if not self.referencia_estacion_var.get():
                self.mostrar_error_novedades("El campo 'Referencia Estacion' es obligatorio.")
                return False
            if not self.supervisor_var.get():
                self.mostrar_error_novedades("El campo 'Supervisor' es obligatorio.")
                return False
            # Verifica otros campos de la misma forma...

            return True
        
    def validar_campos_requeridos_cambios(self):
            # Verificar si los campos requeridos están vacíos
            if not self.legajo_var.get():
                self.mostrar_error_cambios("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_var.get():
                self.mostrar_error_cambios("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.legajo_2_var.get():
                self.mostrar_error_cambios("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_2_var.get():
                self.mostrar_error_cambios("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.fecha_cambio_turno_var.get():
                self.mostrar_error_cambios("El campo 'Fecha de cambio de turno' es obligatorio.")
                return False
            if not self.referencia_estacion_var.get():
                self.mostrar_error_cambios("El campo 'Referencia Estacion' es obligatorio.")
                return False
            if not self.supervisor_var.get():
                self.mostrar_error_cambios("El campo 'Supervisor' es obligatorio.")
                return False
            # Verifica otros campos de la misma forma...

            return True
    def mostrar_error_novedades(self, mensaje):
        """Mostrar mensaje de error."""
        error_label = ttk.Label(self.form_frame, text=mensaje, foreground="red")
        error_label.grid(row=11, column=0, columnspan=7, pady=5, sticky="w")
    
    def mostrar_error_cambios(self, mensaje):
        """Mostrar mensaje de error."""
        error_label = ttk.Label(self.form_cambios_frame, text=mensaje, foreground="red")
        error_label.grid(row=15, column=0, columnspan=7, pady=5, sticky="w")

# Crear la ventana principal de Tkinter
if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioExcelApp(root)
    root.mainloop()