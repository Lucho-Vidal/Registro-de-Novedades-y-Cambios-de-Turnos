import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from ttkbootstrap import DateEntry, Style
from datetime import datetime, timedelta
import os
import ctypes
import math
import unicodedata

from config import (
    SHEET_BASE, SHEET_NOVEDADES, SHEET_TIPO_NOVEDAD, SHEET_CAMBIO_TURNOS,
    COL_USUARIO_WINDOWS, PLACEHOLDER_BUSCAR_NOMBRE, DOTACIONES
)
from excel_store import (
    get_workbook_mtime,
    load_workbook_if_needed,
    build_base_cache,
    get_windows_user,
    ensure_user_column,
    get_last_id,
    create_default_workbook,
)
from validators import validar_campos_requeridos_novedades, validar_campos_requeridos_cambios
from forms import FormsManager
from tables import TablesManager


class FormularioExcelApp:
    """Aplicación principal para registro de novedades y cambios de turnos."""
    
    def __init__(self, root):
        self.root = root
        self.root.geometry('1110x650')
        root.state('zoomed')
        self.root.title("Registro de Novedades y Cambios de turnos TK")
        
        # Configurar DPI awareness para Windows
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        self.WIDTH, self.HEIGHT = user32.GetSystemMetrics(0), user32.GetSystemMetrics(1)
        self.WIDTH = math.floor(self.WIDTH * 0.99)
        self.HEIGHT = math.floor(self.HEIGHT * 0.78)

        # Label de carga
        self.labelCarga = tk.Label(self.root, text="Cargando Excel...")
        self.labelCarga.grid(row=1, column=0, padx=10, pady=0, sticky="w")

        # Leer archivo base y cargar tema
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

        # Crear archivo Excel si no existe
        if not os.path.exists(self.excel_file):
            self.crear_archivo_excel()
        
        self.current_view = 'table'
        self.cargar_excel()
        
        # Cargar tipos de novedad
        self.tipo_novedades = []
        self.cargarTipoNovedades()
        
        # Aplicar estilo ttkbootstrap
        self.style = Style()
        self.style.theme_use(self.theme)
        self.configurar_estilos_formularios()
        
        # Temas disponibles
        if hasattr(self.style, "theme_names"):
            self.temas = self.style.theme_names()
        else:
            self.temas = [
                "cosmo", "litera", "minty", "pulse", "quartz",
                "flatly", "journal", "solar", "cerculean",
                "darkly", "sandstone", "superhero", "morph"
            ]
        
        # Menú principal
        self._crear_menu()
        
        # Etiqueta de tema
        self.label = tk.Label(self.root, text=f"Tema actual: {self.theme}", font=("Arial", 8))
        self.label.grid(row=1, column=0, padx=10, pady=0, sticky="e")
        
        # Variables de formularios
        self._inicializar_variables()
        
        # Marcos principales
        main_frame = tk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew")

        self.form_frame = tk.Frame(main_frame)
        self.form_cambios_frame = tk.Frame(main_frame)
        self.table_cambios_frame = tk.Frame(main_frame)
        self.table_frame = tk.Frame(main_frame)
        self.form_novedades_creado = False
        self.form_cambios_creado = False
        self.tabla_novedades_creada = False
        self.tabla_cambios_creada = False
        self.table_frame.grid(row=0, column=0, padx=10, pady=10)

        self.root.grid_rowconfigure(0, weight=1) 
        self.root.grid_columnconfigure(0, weight=1)
        
        # Inicializar managers
        self.forms_manager = FormsManager(self)
        self.tables_manager = TablesManager(self)
        
        # Crear tabla inicial
        self.tables_manager.crear_tabla_novedades()
        
        # Refresh periódico
        self.root.after(60000, self.refrescar_excel_periodicamente)

    def _crear_menu(self):
        """Crea el menú principal de la aplicación."""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Menú Archivo
        self.archivo_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Archivo", menu=self.archivo_menu)
        self.archivo_menu.add_command(label="Seleccionar archivo", command=self.seleccionar_archivo)
        
        # Menú Opciones
        self.opciones_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Opciones", menu=self.opciones_menu)
        
        # Submenú Seleccionar Tema
        self.temas_menu = tk.Menu(self.opciones_menu, tearoff=0)
        self.opciones_menu.add_cascade(label="Seleccionar Tema", menu=self.temas_menu)
        
        for tema in self.temas:
            self.temas_menu.add_command(label=tema, command=lambda t=tema: self.cambiar_tema(t))

    def _inicializar_variables(self):
        """Inicializa todas las variables StringVar de los formularios."""
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

    def obtener_mtime_excel(self):
        """Obtiene el mtime del archivo Excel."""
        return get_workbook_mtime(self.excel_file)

    def actualizar_cache_base(self):
        """Actualiza el caché de datos de la hoja BASE."""
        self.base_rows, self.base_index = build_base_cache(self.sheet_base)

    def normalizar_texto(self, valor):
        """Normaliza un texto para búsquedas."""
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        return " ".join(texto.lower().split())

    def obtener_usuario_windows(self):
        """Obtiene el nombre de usuario de Windows."""
        return get_windows_user()

    def asegurar_columna_usuario(self, sheet, titulo=COL_USUARIO_WINDOWS):
        """Asegura que la columna de usuario existe en la hoja."""
        return ensure_user_column(sheet, titulo)

    def cargar_excel(self, solo_si_cambio=False):
        """Carga el libro de Excel si ha cambiado."""
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
                self.tables_manager.actualizar_tabla()
            if self.current_view == 'table':
                print("Novedades actualizadas correctamente.")
                self.labelCarga.config(text="Novedades actualizadas correctamente.")
            elif self.current_view == 'table_cambios':
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

    def refrescar_excel_periodicamente(self):
        """Refresca el Excel cada 60 segundos."""
        self.cargar_excel(solo_si_cambio=True)
        self.root.after(60000, self.refrescar_excel_periodicamente)

    def leer_archivo_base(self):
        """Lee la ruta del archivo Excel desde path_base."""
        try:
            with open("path_base", "r", encoding="utf-8") as file:
                self.excel_file = file.read().strip()
                self.excel_file = self.excel_file.replace("\\", "\\\\")
        except Exception as e:
            print(f"Error leyendo el archivo: {e}")
            self.excel_file = r'PLANILLA NOVEDADES PERSONAL ABORDO.xlsx'

    def seleccionar_archivo(self):
        """Abre el diálogo para seleccionar un archivo Excel."""
        archivo_seleccionado = filedialog.askopenfilename(
            title="Seleccionar archivo",
            filetypes=[("Excel", "*.xlsx*")]
        )

        if archivo_seleccionado:
            try:
                with open("path_base", "w", encoding="utf-8") as f:
                    f.write(archivo_seleccionado)
                print(f"Ruta guardada: {archivo_seleccionado}")
                self.excel_file = archivo_seleccionado
                self.cargar_excel()
                self.tables_manager.actualizar_tabla()
            except Exception as e:
                print(f"Error al guardar la ruta: {e}")

    def cambiar_tema(self, nuevo_tema):
        """Cambia el tema de la aplicación."""
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
        """Configura los estilos de los campos readonly según el tema."""
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
        """Carga el tema almacenado en el archivo theme."""
        try:
            with open(self.theme_file, "r", encoding="utf-8") as file:
                return file.read().strip()
        except FileNotFoundError:
            return "flatly"
        except Exception as e:
            print(f"Error al cargar el tema: {e}")
            return "flatly"

    def toggle_view(self, target_view=None):
        """Alterna entre las vistas (tabla/formulario)."""
        self.form_frame.grid_forget()
        self.table_frame.grid_forget()
        self.form_cambios_frame.grid_forget()
        self.table_cambios_frame.grid_forget()

        if target_view is None:
            target_view = "table"
        self.current_view = target_view

        if self.current_view == "form":
            self.form_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.form_novedades_creado:
                self.forms_manager.mostrar_formulario_novedades()
                self.form_novedades_creado = True
        elif self.current_view == "form_cambios":
            self.form_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.form_cambios_creado:
                self.forms_manager.mostrar_formulario_cambios()
                self.form_cambios_creado = True
        elif self.current_view == "table_cambios":
            self.table_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.tabla_cambios_creada:
                self.tables_manager.crear_tabla_cambios()
            else:
                self.tables_manager.cargar_datos_completos_cambios()
        else:
            self.table_frame.grid(row=0, column=0, padx=10, pady=10)
            if not self.tabla_novedades_creada:
                self.tables_manager.crear_tabla_novedades()
            else:
                self.tables_manager.cargar_datos_completos_novedades()
    
    def cargarTipoNovedades(self):
        """Carga los tipos de novedad desde la hoja TipoNovedad."""
        for row in self.sheet_tipo_novedad.iter_rows(min_row=2, values_only=True):
            tipo_novedad = row[0]
            self.tipo_novedades.append(tipo_novedad)

    def crear_archivo_excel(self):
        """Crea un archivo Excel con la estructura por defecto."""
        create_default_workbook(
            self.excel_file,
            SHEET_BASE,
            SHEET_NOVEDADES,
            SHEET_TIPO_NOVEDAD,
            SHEET_CAMBIO_TURNOS,
            COL_USUARIO_WINDOWS,
        )
        print(f"Archivo creado: {self.excel_file}")

    def obtener_ultimo_id(self, sheet):
        """Obtiene el último ID de una hoja."""
        return get_last_id(sheet)

    def obtener_nuevo_id_con_sincronizacion(self, nombre_hoja):
        """Obtiene un nuevo ID sincronizando con el archivo en disco."""
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


# Constantes accesibles para la clase
FormularioExcelApp.SHEET_NOVEDADES = SHEET_NOVEDADES
FormularioExcelApp.SHEET_CAMBIO_TURNOS = SHEET_CAMBIO_TURNOS
FormularioExcelApp.PLACEHOLDER_BUSCAR_NOMBRE = PLACEHOLDER_BUSCAR_NOMBRE
FormularioExcelApp.DOTACIONES = DOTACIONES


if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioExcelApp(root)
    root.mainloop()
