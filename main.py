import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from ttkbootstrap import Style
import ctypes
import math
import time
import unicodedata
from pathlib import Path

from config import (
    PLACEHOLDER_BUSCAR_NOMBRE, DOTACIONES
)
from excel_store import get_windows_user
from forms import FormsManager
from tables import TablesManager
from sqlite_store import SQLiteSheetAdapter
from excel_migration import migrate_operational_sheet, migrate_empleados_sheet
from excel_exporter import export_database
from auth import AuthService
from admin_views import AdminViews
from database_bootstrap import open_configured_store
from login_view import LoginView
from records_service import RecordsService
from outlook_mailer import enviar_informe_outlook


class FormularioExcelApp:
    """Aplicación principal para registro de novedades y cambios de turnos."""
    
    def __init__(self, root, db_store=None, current_user=None):
        self.root = root
        self.current_user = current_user or {}
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
        self.labelCarga = tk.Label(self.root, text="Cargando base de datos...")
        self.labelCarga.grid(row=1, column=0, padx=10, pady=0, sticky="w")

        # Leer archivo base y cargar tema
        self.leer_archivo_base()
        self.theme_file = 'theme'
        self.theme = self.cargar_tema()
        self.PLACEHOLDER_BUSCAR_NOMBRE = PLACEHOLDER_BUSCAR_NOMBRE
        self.base_rows = []
        self.base_index = {}
        self.filtro_after_novedades = None
        self.filtro_after_cambios = None
        self.error_novedades_label = None
        self.error_cambios_label = None

        self.inicializar_base_datos(db_store)
        self.auth_service = AuthService(self.db_store)
        self.records_service = RecordsService(self.db_store)
        self.admin_views = AdminViews(self)
        self.current_view = 'table'
        self.cargar_excel()
        
        # Cargar tipos de novedad
        self.tipo_novedades = []
        self.cargarTipoNovedades()
        self.dotaciones = []
        self.cargarDotaciones()
        self.personal_estacion = []
        self.cargarPersonalEstacion()
        
        # Aplicar estilo ttkbootstrap
        self.style = Style()
        available_themes = set(self.style.theme_names())
        if self.theme not in available_themes:
            self.theme = {
                "nord-dark": "darkly",
                "nord-light": "cosmo",
                "bootstrap-dark": "darkly",
                "bootstrap-light": "flatly",
                "sandstone-dark": "superhero",
                "sandstone-light": "sandstone",
            }.get(self.theme, "flatly")
        if self.theme not in available_themes:
            self.theme = "flatly"
        self.style.theme_use(self.theme)
        self._aplicar_fondo_tema()
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
        self.label = tk.Label(self.root, text=f"Tema actual: {self.theme}", font=("Arial", 8),
                              bg=self.ui_background, fg=self.ui_foreground)
        self.label.grid(row=1, column=0, padx=10, pady=0, sticky="e")
        
        # Variables de formularios
        self._inicializar_variables()
        
        # Marcos principales
        main_frame = ttk.Frame(self.root)
        main_frame.grid(row=0, column=0, sticky="nsew")

        self.form_frame = ttk.Frame(main_frame)
        self.form_cambios_frame = ttk.Frame(main_frame)
        self.table_cambios_frame = ttk.Frame(main_frame)
        self.table_frame = ttk.Frame(main_frame)
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
        
        # Crear la primera vista disponible según los permisos de la sesión.
        if self.tiene_permiso("novedades.ver"):
            self.tables_manager.crear_tabla_novedades()
        elif self.tiene_permiso("cambios_turno.ver"):
            self.current_view = "table_cambios"
            self.table_frame.grid_forget()
            self.table_cambios_frame.grid(row=0, column=0, padx=10, pady=10)
            self.tables_manager.crear_tabla_cambios()
        else:
            self.table_frame.grid_forget()
            self.labelCarga.config(text="Su usuario no tiene módulos habilitados.")
        
        # Refresh periódico
        self.root.after(60000, self.refrescar_excel_periodicamente)
        self.session_active = True
        self.session_last_activity = time.monotonic()
        self.session_timeout_after = None
        self.iniciar_control_sesion()

    def tiene_permiso(self, permiso):
        user_id = self.current_user.get("id")
        return bool(user_id and self.auth_service.tiene_permiso(user_id, permiso))

    def _aplicar_fondo_tema(self):
        """Sincroniza el fondo de los widgets tk con el tema ttkbootstrap."""
        self.ui_background = str(self.style.colors.bg)
        self.ui_foreground = str(self.style.colors.fg)
        self.style.configure("TFrame", background=self.ui_background)
        self.style.configure("TLabel", background=self.ui_background, foreground=self.ui_foreground)
        self.root.configure(background=self.ui_background)
        if hasattr(self, "labelCarga"):
            self.labelCarga.configure(background=self.ui_background, foreground=self.ui_foreground)
        if hasattr(self, "label"):
            self.label.configure(background=self.ui_background, foreground=self.ui_foreground)

    def aplicar_tema_ventana(self, window):
        """Aplica el fondo del tema a una ventana y a sus widgets Tk clásicos.

        ttk toma el tema automáticamente, pero los widgets `tk.*` conservan
        sus colores por defecto. Este método evita fondos blancos aislados en
        ventanas secundarias, especialmente en temas oscuros.
        """
        background = self.ui_background
        foreground = self.ui_foreground
        window.configure(background=background)
        clases_tk = {"Frame", "Label", "Entry", "Text", "Button", "Checkbutton", "Listbox"}
        for child in window.winfo_children():
            if child.winfo_class() in clases_tk:
                opciones = {"background": background, "foreground": foreground}
                if child.winfo_class() in {"Entry", "Text"}:
                    opciones["insertbackground"] = foreground
                try:
                    child.configure(**opciones)
                except tk.TclError:
                    pass
            if child.winfo_children():
                self._aplicar_tema_a_descendientes(child, background, foreground, clases_tk)

    def _aplicar_tema_a_descendientes(self, parent, background, foreground, clases_tk):
        for child in parent.winfo_children():
            if child.winfo_class() in clases_tk:
                opciones = {"background": background, "foreground": foreground}
                if child.winfo_class() in {"Entry", "Text"}:
                    opciones["insertbackground"] = foreground
                try:
                    child.configure(**opciones)
                except tk.TclError:
                    pass
            if child.winfo_children():
                self._aplicar_tema_a_descendientes(child, background, foreground, clases_tk)

    def requerir_permiso(self, permiso):
        if self.tiene_permiso(permiso):
            return True
        messagebox.showwarning("Acceso denegado", "Su usuario no tiene permiso para esta sección.", parent=self.root)
        return False

    def iniciar_control_sesion(self):
        self.root.bind_all("<Any-KeyPress>", self._actividad_sesion, add="+")
        self.root.bind_all("<Any-Button>", self._actividad_sesion, add="+")
        self.renovar_sesion()

    def _actividad_sesion(self, _event=None):
        if self.session_active:
            self.renovar_sesion()

    def renovar_sesion(self):
        if not getattr(self, "session_active", False):
            return
        self.session_last_activity = time.monotonic()
        if self.session_timeout_after:
            self.root.after_cancel(self.session_timeout_after)
        self.session_timeout_after = self.root.after(30000, self.verificar_sesion)

    def verificar_sesion(self):
        if not self.session_active:
            return
        try:
            minutos = max(1, int(self.db_store.get_configuracion("sesion_minutos", 30)))
        except (TypeError, ValueError):
            minutos = 30
        if time.monotonic() - self.session_last_activity >= minutos * 60:
            self.cerrar_sesion_por_expiracion()
        else:
            self.renovar_sesion()

    def cerrar_sesion_por_expiracion(self):
        self.session_active = False
        if self.session_timeout_after:
            self.root.after_cancel(self.session_timeout_after)
        self.root.unbind_all("<Any-KeyPress>")
        self.root.unbind_all("<Any-Button>")
        messagebox.showwarning("Sesión expirada", "La sesión expiró por inactividad.", parent=self.root)
        self.root.config(menu="")
        for child in list(self.root.winfo_children()):
            child.destroy()
        self.root.withdraw()
        LoginView(self.root, self.db_store, lambda user, store: FormularioExcelApp(self.root, db_store=store, current_user=user))

    def enviar_informe_novedad(self, record_id):
        row = self.records_service.obtener_novedad(record_id)
        if not row:
            raise ValueError("No se encontró la novedad para enviar el informe.")
        destinatarios = [item[2] for item in self.records_service.listar_destinatarios_informe(False)]
        cuerpo = "\n".join([
            "Se registró una novedad tipo Informe.",
            f"ID: {row['id']}", f"Fecha de registro: {row['registrado_en']}",
            f"Legajo: {row['legajo']}", f"Apellidos y nombres: {row['apellidos_nombres']}",
            f"Especialidad: {row['especialidad']}", f"Dotación: {row['dotacion']}",
            f"Fecha inicio: {row['fecha_inicio']}", f"Fecha fin: {row['fecha_fin'] or '-'}",
            f"Referencia estación: {row['referencia_estacion']}", f"Supervisor: {row['supervisor']}",
            "", "Observaciones:", row['observaciones'] or "-",
        ])
        enviar_informe_outlook(destinatarios, f"Informe de novedad #{row['id']}", cuerpo)
        self.records_service.registrar_auditoria(
            "enviado", "informe_email", record_id, self.current_user.get("id"), self.obtener_usuario_windows(),
            after={"destinatarios": destinatarios},
        )

    def _crear_menu(self):
        """Crea el menú principal de la aplicación."""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Menú Archivo
        self.archivo_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Archivo", menu=self.archivo_menu)
        self.archivo_menu.add_command(label="Importar novedades", command=lambda: self.importar_excel_operativo("NOVEDADES"))
        self.archivo_menu.add_command(label="Importar cambios de turno", command=lambda: self.importar_excel_operativo("Cambio de Turnos"))
        if self.tiene_permiso("empleados.importar") or self.tiene_permiso("usuarios.administrar"):
            self.archivo_menu.add_command(label="Importar empleados", command=lambda: self.importar_excel_operativo("BASE"))
        if self.tiene_permiso("excel.exportar"):
            self.archivo_menu.add_command(label="Exportar a Excel", command=self.exportar_excel)
        
        # Menú Opciones
        self.opciones_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Opciones", menu=self.opciones_menu)
        
        # Submenú Seleccionar Tema
        self.temas_menu = tk.Menu(self.opciones_menu, tearoff=0)
        self.opciones_menu.add_cascade(label="Seleccionar Tema", menu=self.temas_menu)
        self.opciones_menu.add_separator()
        self.opciones_menu.add_command(label="Cambiar mi contraseña", command=self.cambiar_mi_password)
        
        for tema in self.temas:
            self.temas_menu.add_command(label=tema, command=lambda t=tema: self.cambiar_tema(t))

        if any(self.tiene_permiso(permission) for permission in (
            "usuarios.administrar", "roles.administrar", "novedades.editar", "dotaciones.administrar", "personalEstacion.ver", "destinatarios_informe.administrar", "sesion.configurar", "auditoria.ver", "registros.recuperar"
        )):
            self.administracion_menu = tk.Menu(self.menu_bar, tearoff=0)
            self.menu_bar.add_cascade(label="Administración", menu=self.administracion_menu)
            if self.tiene_permiso("usuarios.administrar"):
                        self.administracion_menu.add_command(label="Seleccionar base SQLite", command=self.seleccionar_base_sqlite)
            if self.tiene_permiso("usuarios.administrar"):
                self.administracion_menu.add_command(label="Usuarios", command=self.admin_views.mostrar_usuarios)
            if self.tiene_permiso("roles.administrar"):
                self.administracion_menu.add_command(label="Roles y permisos", command=self.admin_views.mostrar_roles)
            if self.tiene_permiso("novedades.editar"):
                self.administracion_menu.add_command(label="Tipos de novedad", command=self.admin_views.mostrar_tipos_novedad)
            if self.tiene_permiso("dotaciones.administrar"):
                self.administracion_menu.add_command(label="Dotaciones", command=self.admin_views.mostrar_dotaciones)
            if self.tiene_permiso("personalEstacion.ver"):
                self.administracion_menu.add_command(label="Personal de estación", command=self.admin_views.mostrar_personal_estacion)
            if self.tiene_permiso("destinatarios_informe.administrar"):
                self.administracion_menu.add_command(label="Destinatarios de informes", command=self.admin_views.mostrar_destinatarios_informe)
            if self.tiene_permiso("sesion.configurar"):
                self.administracion_menu.add_command(label="Tiempo de sesión", command=self.admin_views.mostrar_configuracion_sesion)
            if self.tiene_permiso("sesion.configurar"):
                self.administracion_menu.add_command(label="Tiempos de edición", command=self.admin_views.mostrar_configuracion_tiempos)
            if self.tiene_permiso("registros.recuperar"):
                self.administracion_menu.add_command(label="Registros eliminados", command=self.admin_views.mostrar_registros_eliminados)
            if self.tiene_permiso("auditoria.ver"):
                self.administracion_menu.add_command(label="Auditoría", command=self.admin_views.mostrar_auditoria)

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
        self.supervisor_var = tk.StringVar(value=self.current_user.get("nombre") or self.current_user.get("username", ""))
        self.observaciones_var = tk.StringVar()
        
        self.legajo_2_var = tk.StringVar()
        self.apellidos_nombres_2_var = tk.StringVar()
        self.especialidad_2_var = tk.StringVar()
        self.dotacion_2_var = tk.StringVar()
        self.turnos_2_var = tk.StringVar()
        self.franco_2_var = tk.StringVar()
        self.fecha_cambio_turno_var = tk.StringVar()

    def actualizar_cache_base(self):
        """Actualiza el caché de empleados desde SQLite."""
        self.base_rows = [tuple(row) for row in self.db_store.get_base_rows()]
        self.base_index = {int(row[0]): row for row in self.base_rows if row[0] is not None}

    def normalizar_texto(self, valor):
        """Normaliza un texto para búsquedas."""
        texto = unicodedata.normalize("NFKD", str(valor or ""))
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        return " ".join(texto.lower().split())

    def obtener_usuario_windows(self):
        """Obtiene el nombre de usuario de Windows."""
        return get_windows_user()

    def cargar_excel(self):
        """Actualiza las vistas leyendo SQLite."""
        try:
            self.labelCarga.config(text="Actualizando base de datos...")
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
        except Exception as e:
            print(f"Error cargando la base de datos: {e}")
            self.labelCarga.config(text="Error al cargar la base de datos.")
            return False

    def refrescar_excel_periodicamente(self):
        """Refresca las vistas desde SQLite cada 60 segundos."""
        self.cargar_excel()
        self.root.after(60000, self.refrescar_excel_periodicamente)

    def inicializar_base_datos(self, existing_store=None):
        """Crea la base compartida y migra el XLSX una sola vez."""
        if existing_store is None:
            excel_path, self.db_store = open_configured_store()
        else:
            excel_path = Path(self.excel_file)
            if excel_path.suffix.lower() == ".sqlite":
                excel_path = excel_path.with_suffix(".xlsx")
            self.db_store = existing_store
            self.db_store.initialize()
        self.excel_file = str(excel_path)
        self.database_file = str(self.db_store.database_path)
        self.sheet_novedades = SQLiteSheetAdapter(
            self.db_store,
            """SELECT id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion,
                      turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion,
                      supervisor, observaciones, usuario_windows
               FROM novedades WHERE activo=1 ORDER BY id DESC""",
        )
        self.sheet_cambio_turnos = SQLiteSheetAdapter(
            self.db_store,
            """SELECT id, registrado_en, legajo_1, apellidos_nombres_1, especialidad_1,
                      dotacion_1, turnos_1, franco_1, legajo_2, apellidos_nombres_2,
                      especialidad_2, dotacion_2, turnos_2, franco_2, fecha_cambio,
                      referencia_estacion, supervisor, observaciones, usuario_windows
               FROM cambios_turno WHERE activo=1 ORDER BY id DESC""",
        )

    def leer_archivo_base(self):
        """Lee la ruta de la base SQLite o del Excel inicial desde path_base."""
        try:
            with open("path_base", "r", encoding="utf-8") as file:
                self.excel_file = file.read().strip()
        except Exception as e:
            print(f"Error leyendo el archivo: {e}")
            self.excel_file = r'PLANILLA NOVEDADES PERSONAL ABORDO.xlsx'

    def seleccionar_base_sqlite(self):
        """Selecciona la base SQLite compartida para el próximo inicio."""
        if not self.requerir_permiso("usuarios.administrar"):
            return
        database_path = filedialog.askopenfilename(
            title="Seleccionar base SQLite",
            parent=self.root,
            filetypes=[("Base SQLite", "*.sqlite"), ("Todos los archivos", "*.*")],
        )
        if not database_path:
            return
        try:
            with open("path_base", "w", encoding="utf-8") as file:
                file.write(database_path)
            messagebox.showinfo(
                "Base SQLite",
                "La base fue configurada correctamente. Reinicie la aplicación para aplicar el cambio.",
                parent=self.root,
            )
        except OSError as error:
            messagebox.showerror("Base SQLite", f"No se pudo guardar la configuración: {error}", parent=self.root)

    def exportar_excel(self):
        """Abre el selector de filtros y exporta una copia consistente."""
        if not self.requerir_permiso("excel.exportar"):
            return
        window = tk.Toplevel(self.root)
        window.title("Exportar a Excel")
        window.geometry("500x380")
        self.aplicar_tema_ventana(window)
        form = ttk.Frame(window)
        form.pack(fill="x", padx=20, pady=15)
        tabla = tk.StringVar(value="NOVEDADES")
        fecha_desde = tk.StringVar()
        fecha_hasta = tk.StringVar()
        id_desde = tk.StringVar()
        id_hasta = tk.StringVar()
        tipo = tk.StringVar(value="Todos")
        ttk.Label(form, text="Tabla a exportar").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=tabla, values=["NOVEDADES", "Cambio de Turnos"], state="readonly", width=21).grid(row=0, column=1, sticky="w", pady=5)
        for row, label, variable in ((1, "Fecha desde (DD/MM/AAAA)", fecha_desde), (2, "Fecha hasta (DD/MM/AAAA)", fecha_hasta), (3, "ID desde", id_desde), (4, "ID hasta", id_hasta)):
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", pady=5)
            ttk.Entry(form, textvariable=variable, width=24).grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(form, text="Tipo de novedad").grid(row=5, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=tipo, values=["Todos"] + [row[1] for row in self.records_service.listar_tipos(False)], state="readonly", width=21).grid(row=5, column=1, sticky="w", pady=5)

        def run_export():
            try:
                parsed_id_desde = int(id_desde.get()) if id_desde.get().strip() else None
                parsed_id_hasta = int(id_hasta.get()) if id_hasta.get().strip() else None
                if fecha_desde.get() and len(fecha_desde.get().split("/")) != 3:
                    raise ValueError("La fecha desde debe usar DD/MM/AAAA.")
                if fecha_hasta.get() and len(fecha_hasta.get().split("/")) != 3:
                    raise ValueError("La fecha hasta debe usar DD/MM/AAAA.")
                destino = filedialog.asksaveasfilename(
                    title="Exportar registros a Excel", parent=window,
                    defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                    initialfile="Registro de Novedades.xlsx",
                )
                if not destino:
                    return
                export_database(
                    self.db_store, destino,
                    fecha_desde.get() or None, fecha_hasta.get() or None,
                    parsed_id_desde, parsed_id_hasta,
                    None if tipo.get() == "Todos" else tipo.get(),
                    tables=[tabla.get()],
                )
                window.destroy()
                messagebox.showinfo("Exportación", f"Archivo exportado correctamente:\n{destino}")
            except ValueError as error:
                messagebox.showerror("Exportación", str(error), parent=window)
            except PermissionError:
                messagebox.showerror("Exportación", "No se pudo reemplazar el Excel. Verifique que no esté abierto.", parent=window)
            except Exception as error:
                messagebox.showerror("Exportación", f"No se pudo exportar: {error}", parent=window)

        ttk.Button(window, text="Exportar", command=run_export).pack(pady=10)

    def importar_excel_operativo(self, sheet_name):
        """Importa solo una tabla operativa desde un Excel seleccionado."""
        if sheet_name == "BASE":
            allowed = self.tiene_permiso("empleados.importar") or self.tiene_permiso("usuarios.administrar")
        else:
            permission = "novedades.crear" if sheet_name == "NOVEDADES" else "cambios_turno.crear"
            allowed = self.tiene_permiso(permission)
        if not allowed:
            self.requerir_permiso("empleados.importar" if sheet_name == "BASE" else permission)
            return
        source = filedialog.askopenfilename(
            title=f"Importar {sheet_name}", parent=self.root,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not source:
            return
        try:
            clear_existing = messagebox.askyesno(
                "Limpiar antes de importar",
                f"¿Desea eliminar todos los registros actuales de {sheet_name} antes de importar?\n\n"
                "Esta acción no se puede deshacer. La otra tabla no será modificada.",
                parent=self.root,
            )
            if sheet_name == "BASE":
                count = migrate_empleados_sheet(source, self.db_store, clear_existing=clear_existing)
            else:
                count = migrate_operational_sheet(source, self.db_store, sheet_name, clear_existing=clear_existing)
            self.records_service.registrar_auditoria(
                "importar", sheet_name, None, self.current_user.get("id"), self.obtener_usuario_windows(),
                after={"archivo": source, "registros": count, "tabla_limpiada": clear_existing},
            )
            self.cargar_excel()
            modo = "reemplazaron" if clear_existing else "importaron"
            messagebox.showinfo("Importación", f"Se {modo} {count} registros de {sheet_name}.", parent=self.root)
        except Exception as error:
            messagebox.showerror("Importación", f"No se pudo importar {sheet_name}: {error}", parent=self.root)

    def cambiar_tema(self, nuevo_tema):
        """Cambia el tema de la aplicación."""
        try:
            self.style.theme_use(nuevo_tema)
            self._aplicar_fondo_tema()
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
        self.renovar_sesion()
        target_view = target_view or "table"
        required_permissions = {
            "table": "novedades.ver",
            "form": "novedades.crear",
            "table_cambios": "cambios_turno.ver",
            "form_cambios": "cambios_turno.crear",
        }
        if not self.requerir_permiso(required_permissions.get(target_view, "")):
            return
        self.form_frame.grid_forget()
        self.table_frame.grid_forget()
        self.form_cambios_frame.grid_forget()
        self.table_cambios_frame.grid_forget()

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
        """Carga los tipos de novedad desde SQLite."""
        self.tipo_novedades = self.db_store.get_tipo_novedades()
        combo = getattr(self, "tipo_filter_novedades", None)
        if combo is not None:
            combo.configure(values=["Todos", *self.tipo_novedades])

    def cargarDotaciones(self):
        """Carga las dotaciones activas y actualiza las tablas de filtros."""
        self.dotaciones = self.db_store.sincronizar_dotaciones()
        self.DOTACIONES = ["Todas", *self.dotaciones]
        for tree_name in ("dotacion_filter_novedades", "dotacion_filter_cambios"):
            tree = getattr(self, tree_name, None)
            if tree is not None:
                tree.configure(values=self.DOTACIONES)

    def cargarPersonalEstacion(self):
        self.personal_estacion = self.db_store.get_personal_estacion(False)
        valores = [row[1] for row in self.personal_estacion]
        for widget_name in ("referencia_estacion_novedades_entry", "referencia_estacion_cambios_entry"):
            widget = getattr(self, widget_name, None)
            if widget is not None:
                widget.configure(values=valores)

    def cambiar_mi_password(self):
        actual = simpledialog.askstring("Cambiar contraseña", "Contraseña actual:", show="*", parent=self.root)
        if actual is None:
            return
        nueva = simpledialog.askstring("Cambiar contraseña", "Nueva contraseña:", show="*", parent=self.root)
        if nueva is None:
            return
        confirmar = simpledialog.askstring("Cambiar contraseña", "Repita la nueva contraseña:", show="*", parent=self.root)
        if nueva != confirmar:
            messagebox.showerror("Contraseña", "Las contraseñas nuevas no coinciden.", parent=self.root)
            return
        try:
            self.auth_service.cambiar_mi_password(self.current_user["id"], actual, nueva)
            messagebox.showinfo("Contraseña", "La contraseña fue cambiada correctamente.", parent=self.root)
        except Exception as error:
            messagebox.showerror("Contraseña", str(error), parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    def iniciar_aplicacion(user, store):
        FormularioExcelApp(root, db_store=store, current_user=user)

    try:
        _excel_file, _store = open_configured_store()
        LoginView(root, _store, iniciar_aplicacion)
    except Exception as error:
        root.deiconify()
        messagebox.showerror("Inicio", f"No se pudo iniciar la aplicación: {error}")
        root.destroy()
        raise
    root.mainloop()
