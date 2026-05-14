import tkinter as tk
from tkinter import ttk
from tkinter import messagebox,scrolledtext 
from tkinter import filedialog 
from ttkbootstrap import DateEntry
from ttkbootstrap import Style
from datetime import datetime, timedelta
from tkinter.scrolledtext import ScrolledText
import os
import ctypes
import math
import unicodedata
from excel_store import (
    get_workbook_mtime,
    load_workbook_if_needed,
    build_base_cache,
    get_windows_user,
    ensure_user_column,
    get_last_id,
    create_default_workbook,
)

SHEET_BASE = "BASE"
SHEET_NOVEDADES = "NOVEDADES"
SHEET_TIPO_NOVEDAD = "TipoNovedad"
SHEET_CAMBIO_TURNOS = "Cambio de Turnos"
COL_USUARIO_WINDOWS = "USUARIO WINDOWS"
PLACEHOLDER_BUSCAR_NOMBRE = "Buscar por nombre"
DOTACIONES = ["Todas", "PC", "LLV", "TY", "LP", "OA", "K5", "RE", "CÑ", "AK"]

class FormularioExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry('1110x650')  # Ajuste inicial de la ventana
        root.state('zoomed')
        self.root.title("Registro de Novedades y Cambios de turnos TK")
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        self.WIDTH, self.HEIGHT  = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
        self.WIDTH = math.floor(self.WIDTH * 0.99)
        self.HEIGHT = math.floor(self.HEIGHT * 0.78)

        self.labelCarga = tk.Label(self.root, text="Cargando Excel...")
        self.labelCarga.grid(row=1, column=0, padx=10, pady=0, sticky="w")

        self.leer_archivo_base()
        self.theme_file = 'theme'
        self.theme = self.cargar_tema()
        self.excel_last_mtime = None
        self.base_rows = []
        self.base_index = {}
        self.filtro_after_novedades = None
        self.filtro_after_cambios = None
        self.error_novedades_label = None
        self.error_cambios_label = None

        # Verificar si el archivo existe; si no, crearlo
        if not os.path.exists(self.excel_file):
            self.crear_archivo_excel()
        
        self.current_view = 'table'
        self.cargar_excel()
        
        
        # Cargar opciones de tipo de novedad
        self.tipo_novedades = []
        self.cargarTipoNovedades()
        
        # Aplicar el estilo ttkbootstrap
        self.style = Style()
        self.style.theme_use(self.theme)  # Tema inicial
        self.configurar_estilos_formularios()
        # Obtener temas disponibles (puedes definir manualmente los temas si es necesario)
        if hasattr(self.style, "theme_names"):
            self.temas = self.style.theme_names()
        else:
            # Lista predeterminada de temas en ttkbootstrap
            self.temas = [
                "cosmo", "litera", "minty", "pulse", "quartz",
                "flatly", "journal", "solar", "cerculean",
                "darkly", "sandstone", "superhero", "morph"
            ]
        
        # Configurar el menú principal
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Menú Archivo
        self.archivo_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Archivo", menu=self.archivo_menu)

        # Opción para seleccionar archivo
        self.archivo_menu.add_command(label="Seleccionar archivo", command=self.seleccionar_archivo)
        # self.archivo_menu.add_command(label="Maximizar", command=self.cambiar_tamanio_ventana)
        
        # Menú "Opciones"
        self.opciones_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Opciones", menu=self.opciones_menu)
        
        # Submenú "Seleccionar Tema"
        self.temas_menu = tk.Menu(self.opciones_menu, tearoff=0)
        self.opciones_menu.add_cascade(label="Seleccionar Tema", menu=self.temas_menu)
        
        # Agregar temas al submenú
        for tema in self.temas:
            self.temas_menu.add_command(label=tema, command=lambda t=tema: self.cambiar_tema(t))
        
        # Etiqueta de ejemplo para visualizar el tema
        self.label = tk.Label(self.root, text=f"Tema actual: {self.theme}", font=("Arial", 8))
        self.label.grid(row=1, column=0, padx=10, pady=0, sticky="e")
        

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
        self.form_novedades_creado = False
        self.form_cambios_creado = False
        self.tabla_novedades_creada = False
        self.tabla_cambios_creada = False
        self.table_frame.grid(row=0, column=0, padx=10, pady=10)  # Muestra la tabla desde el inicio

        # Configurar el grid de la ventana principal para que el marco se expanda 
        self.root.grid_rowconfigure(0, weight=1) 
        self.root.grid_columnconfigure(0, weight=1)
        
        self.crear_tabla_novedades()
        self.root.after(60000, self.refrescar_excel_periodicamente)

    def obtener_mtime_excel(self):
        return get_workbook_mtime(self.excel_file)

    def actualizar_cache_base(self):
        self.base_rows, self.base_index = build_base_cache(self.sheet_base)

    def normalizar_texto(self, valor):
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        return " ".join(texto.lower().split())

    def actualizar_contador_resultados(self, label, cantidad, sufijo=""):
        if cantidad == 1:
            texto = "1 resultado"
        else:
            texto = f"{cantidad} resultados"
        label.config(text=f"{texto}{sufijo}")

    def obtener_usuario_windows(self):
        return get_windows_user()

    def asegurar_columna_usuario(self, sheet, titulo=COL_USUARIO_WINDOWS):
        return ensure_user_column(sheet, titulo)

    def cargar_excel(self, solo_si_cambio=False):
        try:
            wb, mtime_actual, changed = load_workbook_if_needed(
                self.excel_file,
                self.excel_last_mtime,
                only_if_changed=solo_si_cambio,
            )
            if not changed:
                return False

            print(f"Abriendo archivo Excel: {self.excel_file}")
            self.labelCarga.config(text="Actualizando....")
            self.wb = wb
            self.sheet_base = self.wb[SHEET_BASE]
            self.sheet_novedades = self.wb[SHEET_NOVEDADES]
            self.sheet_tipo_novedad = self.wb[SHEET_TIPO_NOVEDAD]
            self.sheet_cambio_turnos = self.wb[SHEET_CAMBIO_TURNOS]
            self.excel_last_mtime = mtime_actual
            self.actualizar_cache_base()

            if hasattr(self, "tabla_novedades") or hasattr(self, "table_cambios"):
                self.actualizar_tabla()
            if(self.current_view == 'table'):
                print("Novedades actualizadas correctamente.")
                self.labelCarga.config(text="Novedades actualizadas correctamente.")
            elif(self.current_view == 'table_cambios'):
                print("Cambios de turnos actualizados correctamente.")
                self.labelCarga.config(text="Cambios de turnos actualizados correctamente.")
            return True
        except FileNotFoundError:
            print("El archivo Excel no se encontró.")
            self.labelCarga.config(text="Archivo Excel no encontrado.")
            return False
        except Exception as e:
            print(f"Error abriendo o cargando el archivo Excel: {e}")
            self.labelCarga.config(text="Error al cargar el archivo Excel.")
            return False
    
    # Función para actualizar la tabla con datos (ejemplo usando Listbox)
    def actualizar_tabla(self):
        if self.sheet_novedades and self.current_view == 'table' and hasattr(self, "tabla_novedades"):
            if not hasattr(self, "apellido_filter_novedades_var") or not hasattr(self, "dotacion_filter_novedades_var"):
                return
            if self.apellido_filter_novedades_var.get() != PLACEHOLDER_BUSCAR_NOMBRE or self.dotacion_filter_novedades_var.get() != "Todas":
                return
            # Limpiar la tabla antes de actualizar
            self.tabla_novedades.delete(*self.tabla_novedades.get_children())
            total = 0
            # for item in self.tabla_novedades.get_children():  # Iterar sobre los ítems existentes
            #     self.tabla_novedades.delete(item)  # Eliminar cada ítem
            # Leer las primeras filas de la hoja de "BASE" y mostrarlas en la tabla
            for row in self.sheet_novedades.iter_rows(min_row=2, values_only=True):  # min_row=2 para omitir encabezados
                # Reemplazar valores None por "-"
                row_data = ["-" if celda is None else celda for celda in row]
                self.tabla_novedades.insert("", "end", values=row_data)  # Insertar los datos en la tabla
                total += 1
            if hasattr(self, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.resultados_novedades_label, total)
        
        if self.sheet_cambio_turnos and self.current_view == "table_cambios" and hasattr(self, "table_cambios"):
            if not hasattr(self, "apellido_filter_cambios_var") or not hasattr(self, "dotacion_filter_cambios_var"):
                return
            if self.apellido_filter_cambios_var.get() != PLACEHOLDER_BUSCAR_NOMBRE or self.dotacion_filter_cambios_var.get() != "Todas":
                return
            total = 0
            for item in self.table_cambios.get_children():  # Iterar sobre los ítems existentes
                self.table_cambios.delete(item)  # Eliminar cada ítem
            for row in self.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):  # min_row=2 para omitir encabezados
                # Reemplazar valores None por "-"
                row_data = ["-" if celda is None else celda for celda in row]
                self.table_cambios.insert("", "end", values=row_data)
                total += 1
            if hasattr(self, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.resultados_cambios_label, total)

    def refrescar_excel_periodicamente(self):
        self.cargar_excel(solo_si_cambio=True)
        self.root.after(60000, self.refrescar_excel_periodicamente)

    def leer_archivo_base(self):
        try:
            with open("path_base", "r", encoding="utf-8") as file:
                self.excel_file = file.read().strip()
                self.excel_file = self.excel_file.replace("\\", "\\\\")  # Reemplazar '\' por '\\'
        except Exception as e:
            print(f"Error leyendo el archivo: {e}")
            self.excel_file = r'PLANILLA NOVEDADES PERSONAL ABORDO.xlsx'

    def seleccionar_archivo(self):
        """Abre el explorador de archivos y guarda la ruta seleccionada en un archivo."""
        # Abrir el cuadro de diálogo para seleccionar archivo
        archivo_seleccionado = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Excel", "*.xlsx*")])

        if archivo_seleccionado:  # Si se seleccionó un archivo
            try:
                # Guardar la ruta en el archivo 'path_base'
                with open("path_base", "w", encoding="utf-8") as f:
                    f.write(archivo_seleccionado)
                print(f"Ruta guardada: {archivo_seleccionado}")
                self.excel_file = archivo_seleccionado
                self.cargar_excel()
                self.actualizar_tabla()
            except Exception as e:
                print(f"Error al guardar la ruta: {e}")

    def cambiar_tema(self, nuevo_tema):
        """Cambiar el tema y guardarlo en un archivo"""
        try:
            self.style.theme_use(nuevo_tema)
            self.configurar_estilos_formularios()
            self.theme = nuevo_tema
            with open(self.theme_file, "w", encoding="utf-8") as file:
                file.write(nuevo_tema)
            self.label.config(text=f"Tema actual: {self.theme}")
            print(f"Tema cambiado a: {nuevo_tema}")
        except Exception as e:
            print(f"Error al cambiar el tema: {e}")

    def configurar_estilos_formularios(self):
        temas_oscuros = {"darkly", "superhero", "solar"}
        tema_actual = self.style.theme_use()
        if tema_actual in temas_oscuros:
            fg = "#f4f4f4"
            bg = "#2b2f36"
        else:
            fg = "#111111"
            bg = "#ffffff"

        self.style.configure("Readonly.TEntry", foreground=fg, fieldbackground=bg)
        self.style.map(
            "Readonly.TEntry",
            foreground=[("readonly", fg)],
            fieldbackground=[("readonly", bg)]
        )

    def cargar_tema(self):
        """Cargar el tema almacenado en el archivo"""
        try:
            with open(self.theme_file, "r", encoding="utf-8") as file:
                return file.read().strip()  # Leer y limpiar espacios
        except FileNotFoundError:
            # Si no existe el archivo, devolver un tema por defecto
            return "flatly"
        except Exception as e:
            print(f"Error al cargar el tema: {e}")
            return "flatly"  # Tema por defecto en caso de error
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
        if target_view is None:
            target_view = "table"
        self.current_view = target_view

        # Mostrar la vista correspondiente
        if self.current_view == "form":
            self.form_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.form_novedades_creado:
                self.mostrar_formulario_novedades()
                self.form_novedades_creado = True
        elif self.current_view == "form_cambios":
            self.form_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.form_cambios_creado:
                self.mostrar_formulario_cambios()
                self.form_cambios_creado = True
        elif self.current_view == "table_cambios":
            self.table_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.tabla_cambios_creada:
                self.crear_tabla_cambios()
            else:
                self.cargar_datos_completos_cambios()
        else:
            self.table_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.tabla_novedades_creada:
                self.crear_tabla_novedades()
            else:
                self.cargar_datos_completos_novedades()
    
    def cargarTipoNovedades(self):
        for row in self.sheet_tipo_novedad.iter_rows(min_row=2, values_only=True):
            tipo_novedad = row[0]  # Suponemos que los tipos de novedad están en la primera columna
            self.tipo_novedades.append(tipo_novedad)

    def crear_archivo_excel(self):
        """Crea un archivo Excel con encabezados si no existe"""
        create_default_workbook(
            self.excel_file,
            SHEET_BASE,
            SHEET_NOVEDADES,
            SHEET_TIPO_NOVEDAD,
            SHEET_CAMBIO_TURNOS,
            COL_USUARIO_WINDOWS,
        )
        print(f"Archivo creado:{self.excel_file}")
        
    def cargar_datos_completos_novedades(self):
        try:
            # Limpiar la tabla antes de cargar datos
            self.tabla_novedades.delete(*self.tabla_novedades.get_children())
            total = 0
            if self.sheet_novedades:
                for fila in self.sheet_novedades.iter_rows(min_row=2, values_only=True):
                    # Reemplazar None por "-"
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    self.tabla_novedades.insert("", "end", values=fila_procesada)
                    total += 1
            else:
                print("La hoja 'NOVEDADES' está vacía o no se pudo cargar.")

            if hasattr(self, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.resultados_novedades_label, total)
        except Exception as e:
            print(f"Error al cargar los datos en el Treeview: {e}")
    def cargar_datos_completos_cambios(self):
        try:
            self.table_cambios.delete(*self.table_cambios.get_children())
            total = 0
            if self.sheet_cambio_turnos:
                for fila in self.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    self.table_cambios.insert("", "end", values=fila_procesada)
                    total += 1
            else:
                print("La hoja 'NOVEDADES' está vacía o no se pudo cargar.")

            if hasattr(self, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.resultados_cambios_label, total)
        except Exception as e:
            print(f"Error al cargar los datos en el Treeview: {e}")
        
    # Filtrar datos según el texto ingresado
    def filtrar_datos_novedades(self, nombre_filtro,dotacion_filtro):
        try:
            if dotacion_filtro == "Todas" and nombre_filtro == PLACEHOLDER_BUSCAR_NOMBRE:
                self.cargar_datos_completos_novedades()
                return

            nombre_filtro_norm = self.normalizar_texto(nombre_filtro)
            dotacion_filtro_norm = self.normalizar_texto(dotacion_filtro)

            # Limpiar la tabla antes de cargar datos filtrados
            self.tabla_novedades.delete(*self.tabla_novedades.get_children())
            total = 0
            if self.sheet_novedades:
                for fila in self.sheet_novedades.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    if len(fila_procesada) <= 5:
                        continue
                    nombre_fila_norm = self.normalizar_texto(fila_procesada[3])
                    dotacion_fila_norm = self.normalizar_texto(fila_procesada[5])
                    if dotacion_filtro == "Todas":
                        if nombre_filtro_norm in nombre_fila_norm:
                            self.tabla_novedades.insert("", "end", values=fila_procesada)
                            total += 1
                    elif nombre_filtro == PLACEHOLDER_BUSCAR_NOMBRE:
                        if dotacion_filtro_norm in dotacion_fila_norm:
                            self.tabla_novedades.insert("", "end", values=fila_procesada)
                            total += 1
                    else:
                        if nombre_filtro_norm in nombre_fila_norm and dotacion_filtro_norm in dotacion_fila_norm:
                            self.tabla_novedades.insert("", "end", values=fila_procesada)
                            total += 1

            if hasattr(self, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.resultados_novedades_label, total)
        except Exception as e:
            print(f"Error al filtrar los datos en el Treeview novedades: {e}")

    def programar_filtrado_novedades(self):
        if self.filtro_after_novedades:
            self.root.after_cancel(self.filtro_after_novedades)
        self.filtro_after_novedades = self.root.after(
            250,
            lambda: self.filtrar_datos_novedades(self.apellido_filter_novedades_var.get(), self.dotacion_filter_novedades_var.get())
        )

    def filtrar_datos_cambios(self, nombre_filtro,dotacion_filtro):
        try:
            if dotacion_filtro == "Todas" and nombre_filtro == PLACEHOLDER_BUSCAR_NOMBRE:
                self.cargar_datos_completos_cambios()
                return

            nombre_filtro_norm = self.normalizar_texto(nombre_filtro)
            dotacion_filtro_norm = self.normalizar_texto(dotacion_filtro)

            # Limpiar la tabla antes de cargar datos filtrados
            self.table_cambios.delete(*self.table_cambios.get_children())
            total = 0
            if self.sheet_cambio_turnos:
                for fila in self.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    if len(fila_procesada) <= 9:
                        continue
                    nombre_1_norm = self.normalizar_texto(fila_procesada[3])
                    nombre_2_norm = self.normalizar_texto(fila_procesada[9])
                    dotacion_norm = self.normalizar_texto(fila_procesada[5])
                    if dotacion_filtro == "Todas":
                        if nombre_filtro_norm in nombre_1_norm or nombre_filtro_norm in nombre_2_norm:
                            self.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1
                    elif nombre_filtro == PLACEHOLDER_BUSCAR_NOMBRE:
                        if dotacion_filtro_norm in dotacion_norm:
                            self.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1
                    else:
                        if (nombre_filtro_norm in nombre_1_norm or nombre_filtro_norm in nombre_2_norm) and dotacion_filtro_norm in dotacion_norm:
                            self.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1

            if hasattr(self, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.resultados_cambios_label, total)
        except Exception as e:
            print(f"Error al filtrar los datos en el Treeview cambios: {e}")        

    def programar_filtrado_cambios(self):
        if self.filtro_after_cambios:
            self.root.after_cancel(self.filtro_after_cambios)
        self.filtro_after_cambios = self.root.after(
            250,
            lambda: self.filtrar_datos_cambios(self.apellido_filter_cambios_var.get(), self.dotacion_filter_cambios_var.get())
        )

    def mostrar_modal_detalle(self,novedad,columnas, vista):
        """Mostrar el modal con el Treeview y filtro"""
        # Crear una ventana secundaria (modal)
        modal = tk.Toplevel(self.root)
        modal.title(f"Detalle de {vista}")
        if vista == "novedad":
            modal.geometry("560x500")
            rowButtonCerrar = 15
        else:
            modal.geometry("600x600")
            rowButtonCerrar = 20
        
        ttk.Label(modal, text=f"Detalle de {vista}", font=("Helvetica", 22, "bold")).grid(row=0, column=0, columnspan=2, pady=10, padx=10, sticky="we")

        for idx, (columna, valor) in enumerate(zip(columnas, novedad), start=1):
            ttk.Label(modal, text=(columna+":").capitalize()).grid(row=idx, column=0, pady=2, padx=10, sticky="w")
            if columna == "Observaciones":
                # Usar un ScrolledText para "Observaciones"
                text_area = ScrolledText(modal, wrap=tk.WORD, height=5, width=60)
                text_area.insert("1.0", valor)
                text_area.config(state="disabled")  # Deshabilitar edición
                text_area.grid(row=idx, column=1, pady=2, padx=10, sticky="w")
            else:
                ttk.Label(modal, text=valor).grid(row=idx, column=1, pady=2, padx=10, sticky="w")
        
        # Botón de cerrar el modal
        tk.Button(modal, text="Cerrar", command=modal.destroy).grid(row=rowButtonCerrar, column=0, columnspan=2, pady=10, padx=10)        
    
    def crear_tabla_novedades(self):
        def on_focus_in(event):
            if apellido_filter.get() == PLACEHOLDER_BUSCAR_NOMBRE:
                apellido_filter.delete(0, tk.END)
                apellido_filter.config(fg="black")

        def on_focus_out(event):
            if apellido_filter.get() == "":
                apellido_filter.insert(0, PLACEHOLDER_BUSCAR_NOMBRE)
                apellido_filter.config(fg="grey")   
        def on_double_click(event):
            # Obtener el Treeview que disparó el evento
            treeview = event.widget
            # Obtener el identificador del ítem seleccionado
            selected_item = treeview.selection()

            if selected_item:  # Verificar que haya una selección
                # Obtener los valores asociados al ítem seleccionado
                item_values = treeview.item(selected_item[0], "values")
                # print(item_values)  # Imprime los valores de la fila seleccionada
                self.mostrar_modal_detalle(item_values,columnas,"novedad")
            else:
                print("No se seleccionó ningún elemento.")
            
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD",
            "DOTACION", "TURNOS", "FRANCO", "NOVEDAD", "Fecha de Inicio Novedad", "Fecha de Fin Novedad",
            "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones", "USUARIO WINDOWS"
        ]
        self.apellido_filter_novedades_var = tk.StringVar()
        self.dotacion_filter_novedades_var = tk.StringVar()
        lst_dotaciones = DOTACIONES
        
        # Título y botones
        ttk.Label(self.table_frame, text="Registro de novedades", font=("Helvetica", 20, "bold")).grid(row=0, column=0, pady=10, padx=10, sticky="w")
        ttk.Button(self.table_frame, text="Ver cambios de turno", command=lambda: self.toggle_view("table_cambios")).grid(row=0, column=1, pady=10, padx=10, sticky="e")
        #Filter nombre o apellido
        apellido_filter = tk.Entry(self.table_frame, textvariable=self.apellido_filter_novedades_var, font=("Helvetica",10))
        apellido_filter.grid(row=0, column=2,sticky="e", pady=10, padx=10,ipady=5,ipadx=10)
        #Filter dotacion
        dotacion_filter = ttk.Combobox(self.table_frame, textvariable=self.dotacion_filter_novedades_var, values=lst_dotaciones, width=10)
        dotacion_filter.grid(row=0, column=3, sticky="e", pady=5)
        self.resultados_novedades_label = ttk.Label(self.table_frame, text="0 resultados", font=("Helvetica", 9))
        self.resultados_novedades_label.grid(row=0, column=4, sticky="w", padx=8)
        
        ttk.Button(self.table_frame, text="Nueva novedad", command=lambda: self.toggle_view("form")).grid(row=0, column=5, pady=10, padx=2, sticky="e")
        ttk.Button(self.table_frame, text="Nuevo cambio de turno", command=lambda: self.toggle_view("form_cambios")).grid(row=0, column=6, pady=10, padx=10, sticky="e")

        apellido_filter.insert(0, PLACEHOLDER_BUSCAR_NOMBRE)
        dotacion_filter.insert(0, "Todas")
        apellido_filter.config(fg="grey")
        apellido_filter.bind("<FocusIn>", on_focus_in)
        apellido_filter.bind("<FocusOut>", on_focus_out)
        
        # Detectar cambios en el input
        self.apellido_filter_novedades_var.trace_add("write", lambda *args: self.programar_filtrado_novedades())
        self.dotacion_filter_novedades_var.trace_add("write", lambda *args: self.programar_filtrado_novedades())
        
        # Contenedor del Treeview 
        self.tree_frame = ttk.Frame(self.table_frame, width=self.WIDTH, height=self.HEIGHT)
        # self.tree_frame = ttk.Frame(self.table_frame, width=1100, height=550)
        self.tree_frame.grid(row=1, column=0, columnspan=7, sticky="nsew")
        self.tree_frame.grid_propagate(False)

        # Configurar el grid de tree_frame para que el Treeview ocupe todo el espacio
        self.tree_frame.grid_rowconfigure(0, weight=1)  # Fila 0 del tree_frame se expande
        self.tree_frame.grid_columnconfigure(0, weight=1)  # Columna 0 del tree_frame se expande

        # Crear el Treeview
        self.tabla_novedades = ttk.Treeview(self.tree_frame, columns=columnas, show="headings", height=40)
        self.tabla_novedades.grid(row=0, column=0, sticky="nsew")

        self.tabla_novedades.bind("<Double-1>", on_double_click)
        # Establecer encabezados y anchos de columna
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 120, 90, 120, 120, 120, 140, 120]
        for col, ancho in zip(columnas, anchuras):
            self.tabla_novedades.heading(col, text=col.capitalize())
            self.tabla_novedades.column(col, width=ancho)

        # Scrollbar
        scrollbar_vertical = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tabla_novedades.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tabla_novedades.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        # Configurar el Treeview para usar las barras de scroll
        self.tabla_novedades.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

        # Cargar datos
        self.cargar_datos_completos_novedades()
        self.tabla_novedades_creada = True
        # Asociar el evento de doble clic

    def crear_tabla_cambios(self):
        def on_focus_in(event):
            if apellido_filter.get() == PLACEHOLDER_BUSCAR_NOMBRE:
                apellido_filter.delete(0, tk.END)
                apellido_filter.config(fg="black")

        def on_focus_out(event):
            if apellido_filter.get() == "":
                apellido_filter.insert(0, PLACEHOLDER_BUSCAR_NOMBRE)
                apellido_filter.config(fg="grey")   
        def on_double_click(event):
            # Obtener el Treeview que disparó el evento
            treeview = event.widget
            # Obtener el identificador del ítem seleccionado
            selected_item = treeview.selection()

            if selected_item:  # Verificar que haya una selección
                # Obtener los valores asociados al ítem seleccionado
                item_values = treeview.item(selected_item[0], "values")
                # print(item_values)  # Imprime los valores de la fila seleccionada
                self.mostrar_modal_detalle(item_values,columnas,"cambio de turno")
            else:
                print("No se seleccionó ningún elemento.")
        
        #variables
        self.apellido_filter_cambios_var = tk.StringVar()
        self.dotacion_filter_cambios_var = tk.StringVar()
        lst_dotaciones = DOTACIONES
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", 
            "TURNOS", "FRANCO", "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2", 
            "TURNOS2", "FRANCO2", "Fecha de Cambio de Turno", "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones", "USUARIO WINDOWS"
        ]
        
        # Detectar cambios en el input
        self.apellido_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        self.dotacion_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        
        # Título y botones
        ttk.Label(self.table_cambios_frame, text="Registro de cambios de turnos", font=("Helvetica", 20, "bold")).grid(row=0, column=0, pady=10, padx=10, sticky="w")
        ttk.Button(self.table_cambios_frame, text="Ver novedades", command=lambda: self.toggle_view("table")).grid(row=0, column=1, pady=10,padx=1, sticky="e")
        #Filter nombre o apellido
        apellido_filter = tk.Entry(self.table_cambios_frame, textvariable=self.apellido_filter_cambios_var, font=("Helvetica",10))
        apellido_filter.grid(row=0, column=2,sticky="e", pady=10, padx=10,ipady=5,ipadx=10)
        #Filter dotacion
        dotacion_filter = ttk.Combobox(self.table_cambios_frame, textvariable=self.dotacion_filter_cambios_var, values=lst_dotaciones, width=10)
        dotacion_filter.grid(row=0, column=3, sticky="e", pady=5)
        self.resultados_cambios_label = ttk.Label(self.table_cambios_frame, text="0 resultados", font=("Helvetica", 9))
        self.resultados_cambios_label.grid(row=0, column=4, sticky="w", padx=8)
        ttk.Button(self.table_cambios_frame, text="Nueva novedad", command=lambda: self.toggle_view("form")).grid(row=0, column=5, pady=10, padx=1, sticky="e")
        ttk.Button(self.table_cambios_frame, text="Nuevo cambio de turno", command=lambda: self.toggle_view("form_cambios")).grid(row=0, column=6, pady=10,padx=1,  sticky="e")

        dotacion_filter.insert(0, "Todas")
        apellido_filter.insert(0, PLACEHOLDER_BUSCAR_NOMBRE)
        apellido_filter.config(fg="grey")
        apellido_filter.bind("<FocusIn>", on_focus_in)
        apellido_filter.bind("<FocusOut>", on_focus_out)
        
        # Contenedor del Treeview  
        self.tree_frame = ttk.Frame(self.table_cambios_frame, width=self.WIDTH, height=self.HEIGHT)
        self.tree_frame.grid(row=1, column=0, columnspan=7, sticky="nsew")
        self.tree_frame.grid_propagate(False)

        # Configurar el grid del marco contenedor para que el Treeview se expanda
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.table_cambios = ttk.Treeview(self.tree_frame, columns=columnas, show="headings", height=38)
        self.table_cambios.grid(row=0, column=0, sticky="nsew")

        self.table_cambios.bind("<Double-1>", on_double_click)
        # Establecer encabezados y columnas con anchuras
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 60, 150, 150, 80, 60, 80, 120, 120, 120, 120, 140]
        for col, ancho in zip(columnas, anchuras):
            self.table_cambios.heading(col, text=col)
            self.table_cambios.column(col, width=ancho, anchor='center', stretch=True)

        # Scrollbar
        scrollbar_vertical = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.table_cambios.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.table_cambios.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")

        # Configurar el Treeview para usar las barras de scroll
        self.table_cambios.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
        # Cargar datos
        self.cargar_datos_completos_cambios()
        self.tabla_cambios_creada = True
    
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
        self.apellidos_nombres_entry = ttk.Entry(self.form_frame, textvariable=self.apellidos_nombres_var, state='readonly', style='Readonly.TEntry', width=40)
        self.apellidos_nombres_entry.grid(row=4, column=0, sticky="w", pady=5)

        # Especialidad
        ttk.Label(self.form_frame, text="Especialidad").grid(row=3, column=1, sticky="w", padx=5)
        self.especialidad_entry = ttk.Entry(self.form_frame, textvariable=self.especialidad_var, state='readonly', style='Readonly.TEntry', width=20)
        self.especialidad_entry.grid(row=4, column=1, sticky="w", pady=5, padx=8)

        # Dotación
        ttk.Label(self.form_frame, text="Dotación").grid(row=3, column=2, sticky="w", padx=5)
        self.dotacion_entry = ttk.Entry(self.form_frame, textvariable=self.dotacion_var, state='readonly', style='Readonly.TEntry', width=20)
        self.dotacion_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        # Turnos
        ttk.Label(self.form_frame, text="Turno").grid(row=3, column=3, sticky="w", padx=5)
        self.turnos_entry = ttk.Entry(self.form_frame, textvariable=self.turnos_var, state='readonly', style='Readonly.TEntry', width=40)
        self.turnos_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        # Franco
        ttk.Label(self.form_frame, text="Franco").grid(row=3, column=4, sticky="w", padx=5)
        self.franco_entry = ttk.Entry(self.form_frame, textvariable=self.franco_var, state='readonly', style='Readonly.TEntry', width=20)
        self.franco_entry.grid(row=4, column=4, sticky="w", pady=5, padx=8)

        # Nivel 3: NOVEDAD, Fecha de Inicio Novedad
        ttk.Label(self.form_frame, text="   Tipo Novedad").grid(row=5, column=0, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=0, sticky="w")
        self.novedad_entry = ttk.Combobox(self.form_frame, textvariable=self.novedad_var, values=self.tipo_novedades, width=38, state="readonly")
        self.novedad_entry.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.form_frame, text="   Fecha de Inicio Novedad").grid(row=5, column=3, sticky="w")
        ttk.Label(self.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=3, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_inicio_novedad_entry = DateEntry(self.form_frame,dateformat='%d/%m/%Y',bootstyle="danger")
        self.fecha_inicio_novedad_entry.grid(row=6, column=3, sticky="w", pady=5)       
        self.fecha_inicio_novedad_entry.entry.bind("<Key>", lambda event: "break")

        ttk.Label(self.form_frame, text="Fecha de Fin Novedad").grid(row=5, column=4, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_fin_novedad_entry = DateEntry(self.form_frame,dateformat='%d/%m/%Y',bootstyle="success")
        self.fecha_fin_novedad_entry.grid(row=6, column=4, sticky="w", pady=5) 
        self.fecha_fin_novedad_entry.entry.bind("<Key>", lambda event: "break")
        self.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        self.fecha_fin_novedad_var.set("")
        ttk.Button(self.form_frame, text="Limpiar", command=self.limpiar_fecha_fin_novedad).grid(row=6, column=5, sticky="w", padx=5)
        
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
        self.observaciones_novedades_text = scrolledtext.ScrolledText(self.form_frame, wrap=tk.WORD, height=12,width=180)
        self.observaciones_novedades_text.grid(row=10, column=0, columnspan=7, pady=5, sticky="w")

        # Botón de Guardar Datos en el formulario
        ttk.Button(self.form_frame, text="Guardar Novedad", command=self.guardar_datos_novedades).grid(row=11, column=3, columnspan=2, pady=10)

        # Botón para cerrar el formulario y regresar a la tabla
        ttk.Button(self.form_frame, text="Cerrar", command=lambda: (self.limpiar_formulario_novedades(), self.toggle_view())).grid(row=11, column=4, columnspan=2, pady=10)

        
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
        self.apellidos_nombres_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.apellidos_nombres_var, state='readonly', style='Readonly.TEntry', width=40)
        self.apellidos_nombres_entry.grid(row=4, column=0,columnspan = 2, sticky="w", pady=5)

        # Especialidad
        ttk.Label(self.form_cambios_frame, text="Especialidad").grid(row=3, column=2, sticky="w", padx=5)
        self.especialidad_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.especialidad_var, state='readonly', style='Readonly.TEntry', width=20)
        self.especialidad_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        # Dotación
        ttk.Label(self.form_cambios_frame, text="Dotación").grid(row=3, column=3, sticky="w", padx=5)
        self.dotacion_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.dotacion_var, state='readonly', style='Readonly.TEntry', width=20)
        self.dotacion_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        # Turnos
        ttk.Label(self.form_cambios_frame, text="Turno").grid(row=3, column=4, sticky="w", padx=5)
        self.turnos_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.turnos_var, state='readonly', style='Readonly.TEntry', width=40)
        self.turnos_entry.grid(row=4, column=4, columnspan= 2 , sticky="w", pady=5, padx=8)

        # Franco
        ttk.Label(self.form_cambios_frame, text="Franco").grid(row=3, column=6, sticky="w", padx=5)
        self.franco_entry = ttk.Entry(self.form_cambios_frame, textvariable=self.franco_var, state='readonly', style='Readonly.TEntry', width=20)
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
        self.apellidos_nombres_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.apellidos_nombres_2_var, state='readonly', style='Readonly.TEntry', width=40)
        self.apellidos_nombres_entry_2.grid(row=8, column=0,columnspan = 2, sticky="w", pady=5)

        # Especialidad2
        ttk.Label(self.form_cambios_frame, text="Especialidad 2").grid(row=7, column=2, sticky="w", padx=5)
        self.especialidad_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.especialidad_2_var, state='readonly', style='Readonly.TEntry', width=20)
        self.especialidad_entry_2.grid(row=8, column=2, sticky="w", pady=5, padx=8)

        # Dotación2
        ttk.Label(self.form_cambios_frame, text="Dotación 2").grid(row=7, column=3, sticky="w", padx=5)
        self.dotacion_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.dotacion_2_var, state='readonly', style='Readonly.TEntry', width=20)
        self.dotacion_entry_2.grid(row=8, column=3, sticky="w", pady=5, padx=8)

        # Turnos2
        ttk.Label(self.form_cambios_frame, text="Turno 2").grid(row=7, column=4, sticky="w", padx=5)
        self.turnos_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.turnos_2_var, state='readonly', style='Readonly.TEntry', width=40)
        self.turnos_entry_2.grid(row=8, column=4, columnspan= 2 , sticky="w", pady=5, padx=8)

        # Franco2
        ttk.Label(self.form_cambios_frame, text="Franco 2").grid(row=7, column=6, sticky="w", padx=5)
        self.franco_entry_2 = ttk.Entry(self.form_cambios_frame, textvariable=self.franco_2_var, state='readonly', style='Readonly.TEntry', width=20)
        self.franco_entry_2.grid(row=8, column=6, sticky="w", pady=5, padx=8)
        
        # Nivel 5: Fecha de cambio de turno
        ttk.Label(self.form_cambios_frame, text="   Fecha de cambio de turno").grid(row=9, column=0, sticky="w")
        ttk.Label(self.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=9, column=0, sticky="w")
        # Crear DateEntry con estilo ttkbootstrap
        self.fecha_cambio_turno_entry = DateEntry(self.form_cambios_frame,dateformat='%d/%m/%Y',startdate=datetime.today()+timedelta(days=+1))
        self.fecha_cambio_turno_entry.grid(row=10, column=0, columnspan=2, sticky="w", pady=5)
        self.fecha_cambio_turno_entry.entry.bind("<Key>", lambda event: "break")
        
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
        self.observaciones_cambios_text = scrolledtext.ScrolledText(self.form_cambios_frame, wrap=tk.WORD, height=1,width=170)
        self.observaciones_cambios_text.grid(row=14, column=0, columnspan=7, pady=5, sticky="w")

        # Botón de Guardar Datos en el formulario
        ttk.Button(self.form_cambios_frame, text="Guardar cambio de turno", command=self.guardar_datos_cambios).grid(row=15, column=2, columnspan=2, pady=10)

        # Botón para cerrar el formulario y regresar a la tabla
        ttk.Button(self.form_cambios_frame, text="Cerrar", command=lambda: (self.limpiar_formulario_cambios(), self.toggle_view("table_cambios"))).grid(row=15, column=3, columnspan=2, pady=10)
            
    def mostrar_modal(self,boton=1):
        """Mostrar el modal con el Treeview y filtro"""
        # Crear una ventana secundaria (modal)
        modal = tk.Toplevel(self.root)
        modal.title("Seleccionar Legajo")
        modal.geometry("1250x500")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        registros_busqueda = []
        for row in self.base_rows:
            if not row or len(row) < 6:
                continue
            legajo = row[0]
            apellidos_nombres = row[1]
            especialidad = row[2]
            dotacion = row[3]
            turnos = row[4]
            franco = row[5]

            registros_busqueda.append({
                "values": (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco),
                "nombre_norm": self.normalizar_texto(apellidos_nombres),
                "legajo_norm": self.normalizar_texto(legajo),
            })
        
        # Crear un campo de entrada para filtrar por apellido
        tk.Label(modal, text="Filtrar por apellido o nombre:").grid(row=0, column=0, padx=10, pady=5)
        # Crear el Entry con ttkbootstrap
        # apellido_filter = Entry(modal)
        apellido_filter_var = tk.StringVar()
        apellido_filter = tk.Entry(modal, textvariable=apellido_filter_var)
        apellido_filter.grid(row=0, column=1, padx=10, pady=5,ipady=5,ipadx=50)
        modal_resultados_label = ttk.Label(modal, text="0 resultados")
        modal_resultados_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        
        # Crear el Treeview para mostrar la tabla
        tree = ttk.Treeview(modal, columns=("legajo", "apellido_nombre", "especialidad", "dotacion", "turnos", "franco"), show="headings",height=25)
        tree.heading("legajo", text="LEGAJO SAP")
        tree.heading("apellido_nombre", text="APELLIDOS Y NOMBRES")
        tree.heading("especialidad", text="ESPECIALIDAD")
        tree.heading("dotacion", text="DOTACION")
        tree.heading("turnos", text="TURNOS")
        tree.heading("franco", text="FRANCO")
        tree.grid(row=1, column=0, columnspan=6, padx=20, pady=10)

        search_after_id = None

        def aplicar_filtro():
            for row in tree.get_children():
                tree.delete(row)

            filtro = self.normalizar_texto(apellido_filter_var.get())
            total = 0
            for registro in registros_busqueda:
                if not filtro or filtro in registro["nombre_norm"] or filtro in registro["legajo_norm"]:
                    tree.insert("", "end", values=registro["values"])
                    total += 1
            self.actualizar_contador_resultados(modal_resultados_label, total)

        # Función para llenar la tabla con debounce
        def cargar_tabla():
            nonlocal search_after_id
            if search_after_id:
                modal.after_cancel(search_after_id)
            search_after_id = modal.after(250, aplicar_filtro)

        # Llamar a la función para cargar la tabla al inicio
        aplicar_filtro()
        # Asociar el evento de filtrar al campo de filtro por apellido
        apellido_filter_var.trace_add("write", lambda *args: cargar_tabla())

        def on_double_click(event):
            # Función para asignar el legajo al campo de entrada y cerrar el modal
            selected_item = tree.selection()
            if not selected_item:
                return

            selected_values = tree.item(selected_item[0], "values")
            if not selected_values:
                return

            legajo_sap = selected_values[0]
            apellido_nombre = selected_values[1]
            especialidad = selected_values[2]
            dotacion = selected_values[3]
            turnos = selected_values[4]
            franco = selected_values[5]
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

            row = self.base_index.get(legajo)
            if row:
                # Si el legajo SAP coincide, auto completar los campos
                if campo == 1:
                    self.apellidos_nombres_var.set(row[1])
                    self.especialidad_var.set(row[2])
                    self.dotacion_var.set(row[3])
                    self.turnos_var.set(row[4])
                    self.franco_var.set(row[5])
                elif campo == 2:
                    self.apellidos_nombres_2_var.set(row[1])
                    self.especialidad_2_var.set(row[2])
                    self.dotacion_2_var.set(row[3])
                    self.turnos_2_var.set(row[4])
                    self.franco_2_var.set(row[5])
                print(f"Legajo encontrado: {legajo}")
            else:
                messagebox.showinfo("Legajo No Encontrado", f"El legajo {legajo} no fue encontrado.")
                print(f"Legajo SAP {legajo} no encontrado.")

        except ValueError:
            # Si el legajo ingresado no es un número, mostrar un mensaje de error
            messagebox.showerror("Error de entrada", "Por favor, ingrese un número de legajo válido.")

    def obtener_ultimo_id(self, sheet):
        return get_last_id(sheet)

    def obtener_nuevo_id_con_sincronizacion(self, nombre_hoja):
        hoja_local = self.wb[nombre_hoja]
        ultimo_id_local = self.obtener_ultimo_id(hoja_local)
        recarga_ok = self.cargar_excel(solo_si_cambio=True)
        if recarga_ok is False and self.obtener_mtime_excel() != self.excel_last_mtime:
            raise RuntimeError("No se pudo sincronizar el archivo antes de guardar. Intente nuevamente.")
        hoja_actualizada = self.wb[nombre_hoja]
        ultimo_id_actual = self.obtener_ultimo_id(hoja_actualizada)

        if ultimo_id_actual > ultimo_id_local:
            self.labelCarga.config(text="Se detectaron registros nuevos. Sincronizado antes de guardar.")
            print(f"Sincronización previa al guardado: {nombre_hoja} pasó de ID {ultimo_id_local} a {ultimo_id_actual}.")

        return ultimo_id_actual + 1

    def guardar_datos_novedades(self):
        """Guardar los datos del formulario en el archivo Excel"""
        # self.cargar_excel()
        self.fecha_inicio_novedad_var.set(self.fecha_inicio_novedad_entry.entry.get())
        self.fecha_fin_novedad_var.set(self.fecha_fin_novedad_entry.entry.get())
        self.observaciones_var.set(self.observaciones_novedades_text.get("1.0", "end-1c"))
        
        if self.validar_campos_requeridos_novedades():
            # Aquí guarda los datos si todo está bien
            # Puedes agregar el código para guardar en el archivo de Excel o donde necesites.
            # Añadir a la hoja de Excel y guardar
            try:
                new_id = self.obtener_nuevo_id_con_sincronizacion(SHEET_NOVEDADES)
                current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
                usuario_windows = self.obtener_usuario_windows()
                col_usuario = self.asegurar_columna_usuario(self.sheet_novedades)

                # Añadir los nuevos datos a la hoja
                self.sheet_novedades.insert_rows(2)
                self.sheet_novedades['A2'] = new_id
                self.sheet_novedades['B2'] = current_datetime
                self.sheet_novedades['C2'] = self.legajo_var.get()
                self.sheet_novedades['D2'] = self.apellidos_nombres_var.get()
                self.sheet_novedades['E2'] = self.especialidad_var.get()
                self.sheet_novedades['F2'] = self.dotacion_var.get()
                self.sheet_novedades['G2'] = self.turnos_var.get()
                self.sheet_novedades['H2'] = self.franco_var.get()
                self.sheet_novedades['I2'] = self.novedad_var.get()
                self.sheet_novedades['J2'] = self.fecha_inicio_novedad_var.get()
                self.sheet_novedades['K2'] = self.fecha_fin_novedad_var.get()
                self.sheet_novedades['L2'] = self.referencia_estacion_var.get()
                self.sheet_novedades['M2'] = self.supervisor_var.get()
                self.sheet_novedades['N2'] = self.observaciones_var.get()
                self.sheet_novedades.cell(row=2, column=col_usuario, value=usuario_windows)
                
                # Guardar el archivo Excel
                self.wb.save(self.excel_file)
                self.excel_last_mtime = self.obtener_mtime_excel()

                # Mensaje de confirmación
                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                # Limpiar el formulario y alternar a la vista de tabla
                self.limpiar_formulario_novedades()
                self.toggle_view()
                print("Datos guardados correctamente.")
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar el archivo porque está abierto en otro programa. Por favor, cierre el archivo y vuelva a intentarlo.")
            except Exception as e:
                # Mensaje de error si ocurre un problema
                messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)} Por favor, intente de nuevo y si el problema persiste avise al administrador")
                print(f"Error al guardar los datos: {e}")
            
        else:
            print("Algunos campos son obligatorios y están vacíos.")
            messagebox.showinfo("Faltan Datos requeridos", f"Algunos campos son obligatorios y están vacíos.")
    
    def guardar_datos_cambios(self):
        """Guardar los datos del formulario en el archivo Excel"""
        # self.cargar_excel()
        self.fecha_cambio_turno_var.set(self.fecha_cambio_turno_entry.entry.get())
        self.observaciones_var.set(self.observaciones_cambios_text.get("1.0", "end-1c"))
        
        if self.validar_campos_requeridos_cambios():
            # Aquí guarda los datos si todo está bien
            # Puedes agregar el código para guardar en el archivo de Excel o donde necesites.
            # Añadir a la hoja de Excel y guardar
            try:
                new_id = self.obtener_nuevo_id_con_sincronizacion(SHEET_CAMBIO_TURNOS)
                current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
                usuario_windows = self.obtener_usuario_windows()
                col_usuario = self.asegurar_columna_usuario(self.sheet_cambio_turnos)

                # Añadir los nuevos datos a la hoja
                self.sheet_cambio_turnos.insert_rows(2)
                self.sheet_cambio_turnos['A2'] = new_id
                self.sheet_cambio_turnos['B2'] = current_datetime
                self.sheet_cambio_turnos['C2'] = self.legajo_var.get()
                self.sheet_cambio_turnos['D2'] = self.apellidos_nombres_var.get()
                self.sheet_cambio_turnos['E2'] = self.especialidad_var.get()
                self.sheet_cambio_turnos['F2'] = self.dotacion_var.get()
                self.sheet_cambio_turnos['G2'] = self.turnos_var.get()
                self.sheet_cambio_turnos['H2'] = self.franco_var.get()
                self.sheet_cambio_turnos['I2'] = self.legajo_2_var.get()
                self.sheet_cambio_turnos['J2'] = self.apellidos_nombres_2_var.get()
                self.sheet_cambio_turnos['K2'] = self.especialidad_2_var.get()
                self.sheet_cambio_turnos['L2'] = self.dotacion_2_var.get()
                self.sheet_cambio_turnos['M2'] = self.turnos_2_var.get()
                self.sheet_cambio_turnos['N2'] = self.franco_2_var.get()
                self.sheet_cambio_turnos['O2'] = self.fecha_cambio_turno_var.get()
                self.sheet_cambio_turnos['P2'] = self.referencia_estacion_var.get()
                self.sheet_cambio_turnos['Q2'] = self.supervisor_var.get()
                self.sheet_cambio_turnos['R2'] = self.observaciones_var.get()
                self.sheet_cambio_turnos.cell(row=2, column=col_usuario, value=usuario_windows)

                # Guardar el archivo Excel
                self.wb.save(self.excel_file)
                self.excel_last_mtime = self.obtener_mtime_excel()

                # Mensaje de confirmación
                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                # Limpiar el formulario y alternar a la vista de tabla
                self.limpiar_formulario_cambios()
                self.toggle_view("table_cambios")
                print("Datos guardados correctamente.")
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar el archivo porque está abierto en otro programa. Por favor, cierre el archivo y vuelva a intentarlo.")
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
        if hasattr(self, "observaciones_novedades_text"):
            self.observaciones_novedades_text.delete("1.0", "end")
        if hasattr(self, "fecha_fin_novedad_entry"):
            self.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        if self.error_novedades_label is not None:
            self.error_novedades_label.config(text="")

    def limpiar_fecha_fin_novedad(self):
        if hasattr(self, "fecha_fin_novedad_entry"):
            self.fecha_fin_novedad_entry.entry.delete(0, tk.END)
            self.fecha_fin_novedad_var.set("")

    
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
        if hasattr(self, "observaciones_cambios_text"):
            self.observaciones_cambios_text.delete("1.0", "end")
        if self.error_cambios_label is not None:
            self.error_cambios_label.config(text="")

    def parsear_fecha(self, valor, nombre_campo, obligatorio=False):
        valor = str(valor or "").strip()
        if not valor:
            if obligatorio:
                raise ValueError(f"El campo '{nombre_campo}' es obligatorio.")
            return None
        try:
            return datetime.strptime(valor, "%d/%m/%Y")
        except ValueError:
            raise ValueError(f"El campo '{nombre_campo}' debe tener formato DD/MM/AAAA.")
   
    def validar_campos_requeridos_novedades(self):
            # Verificar si los campos requeridos están vacíos
            if not self.legajo_var.get().strip():
                self.mostrar_error_novedades("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_var.get().strip():
                self.mostrar_error_novedades("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.novedad_var.get().strip():
                self.mostrar_error_novedades("El campo 'Tipo de Novedad' es obligatorio.")
                return False
            if self.novedad_var.get().strip() not in self.tipo_novedades:
                self.mostrar_error_novedades("Seleccione un valor valido en 'Tipo de Novedad'.")
                return False
            if not self.referencia_estacion_var.get().strip():
                self.mostrar_error_novedades("El campo 'Referencia Estacion' es obligatorio.")
                return False
            if not self.supervisor_var.get().strip():
                self.mostrar_error_novedades("El campo 'Supervisor' es obligatorio.")
                return False

            try:
                fecha_inicio = self.parsear_fecha(self.fecha_inicio_novedad_entry.entry.get(), "Fecha de inicio novedad", obligatorio=True)
                fecha_fin = self.parsear_fecha(self.fecha_fin_novedad_entry.entry.get(), "Fecha de fin novedad", obligatorio=False)
            except ValueError as e:
                self.mostrar_error_novedades(str(e))
                return False

            if fecha_fin and fecha_fin < fecha_inicio:
                self.mostrar_error_novedades("'Fecha de Fin Novedad' no puede ser anterior a 'Fecha de Inicio Novedad'.")
                return False

            return True
        
    def validar_campos_requeridos_cambios(self):
            # Verificar si los campos requeridos están vacíos
            if not self.legajo_var.get().strip():
                self.mostrar_error_cambios("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_var.get().strip():
                self.mostrar_error_cambios("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.legajo_2_var.get().strip():
                self.mostrar_error_cambios("El campo 'Legajo' es obligatorio.")
                return False
            if not self.apellidos_nombres_2_var.get().strip():
                self.mostrar_error_cambios("El campo 'Apellido y Nombre' es obligatorio.")
                return False
            if not self.fecha_cambio_turno_var.get().strip():
                self.mostrar_error_cambios("El campo 'Fecha de cambio de turno' es obligatorio.")
                return False
            if not self.referencia_estacion_var.get().strip():
                self.mostrar_error_cambios("El campo 'Referencia Estacion' es obligatorio.")
                return False
            if not self.supervisor_var.get().strip():
                self.mostrar_error_cambios("El campo 'Supervisor' es obligatorio.")
                return False

            try:
                self.parsear_fecha(self.fecha_cambio_turno_entry.entry.get(), "Fecha de cambio de turno", obligatorio=True)
            except ValueError as e:
                self.mostrar_error_cambios(str(e))
                return False

            return True
    def mostrar_error_novedades(self, mensaje):
        """Mostrar mensaje de error."""
        if self.error_novedades_label is None:
            self.error_novedades_label = ttk.Label(self.form_frame, foreground="red")
            self.error_novedades_label.grid(row=11, column=0, columnspan=7, pady=5, sticky="w")
        self.error_novedades_label.config(text=mensaje)
    
    def mostrar_error_cambios(self, mensaje):
        """Mostrar mensaje de error."""
        if self.error_cambios_label is None:
            self.error_cambios_label = ttk.Label(self.form_cambios_frame, foreground="red")
            self.error_cambios_label.grid(row=15, column=0, columnspan=7, pady=5, sticky="w")
        self.error_cambios_label.config(text=mensaje)

# Crear la ventana principal de Tkinter
if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioExcelApp(root)

    # Iniciar el bucle principal de Tkinter
    root.mainloop()
