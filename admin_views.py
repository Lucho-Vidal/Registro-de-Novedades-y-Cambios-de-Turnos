"""Vistas de administración de usuarios, roles y permisos."""

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from excel_exporter import export_auditoria, export_database
from excel_migration import migrate_dotaciones_sheet, migrate_personal_estacion_sheet, migrate_tipos_novedad_sheet
from backups import crear_backup, listar_backups, restaurar_backup
import os


PERMISOS_BASE = (
    "novedades.ver", "novedades.crear", "novedades.editar", "novedades.eliminar",
    "novedades.importar", "novedades.exportar",
    "cambios_turno.ver", "cambios_turno.crear", "cambios_turno.editar", "cambios_turno.eliminar",
    "cambios_turno.importar", "cambios_turno.exportar",
    "usuarios.administrar", "roles.administrar", "empleados.importar", "empleados.administrar",
    "auditoria.ver",
    "dotaciones.administrar",
    "personalEstacion.ver", "personalEstacion.crear", "personalEstacion.editar",
    "personalEstacion.importar", "personalEstacion.exportar",
    "destinatarios_informe.administrar", "sesion.configurar", "registros.recuperar",
    "backup.gestionar",
)


ESPECIALIDADES_EMPLEADO = (
    "Conductor Eléctrico", "Conductor Diesel", "Ayudante Habilitado", "Ayudante Conductor",
    "Guarda Eléctrico", "Guarda Diesel",
)

DIAS_SEMANA = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")

PERMISOS_GRUPOS = (
    ("novedades", "Novedades"),
    ("cambios_turno", "Cambios de turno"),
    ("empleados", "Empleados"),
    ("personalEstacion", "Personal de estación"),
    ("dotaciones", "Dotaciones"),
    ("destinatarios_informe", "Destinatarios de informes"),
    ("usuarios", "Usuarios"),
    ("roles", "Roles y permisos"),
    ("auditoria", "Auditoría"),
    ("sesion", "Sesión"),
    ("registros", "Registros"),
    ("backup", "Copias de seguridad"),
    ("excel", "Excel / Archivo"),
)

ACCIONES_PERMISO = {
    "ver": "Ver", "crear": "Crear", "editar": "Editar", "eliminar": "Eliminar",
    "importar": "Importar", "exportar": "Exportar", "administrar": "Administrar",
    "configurar": "Configurar", "recuperar": "Recuperar", "gestionar": "Gestionar",
}

_ORDEN_ACCION = {
    "ver": 0, "crear": 1, "editar": 2, "eliminar": 3, "importar": 4,
    "exportar": 5, "administrar": 6, "configurar": 7, "recuperar": 8, "gestionar": 9,
}


def _agrupar_permisos(codigos):
    """Agrupa códigos de permiso por módulo para el editor de roles."""
    grupos = [(titulo, []) for _, titulo in PERMISOS_GRUPOS]
    grupos.append(("Otros", []))
    indice = {prefijo: i for i, (prefijo, _) in enumerate(PERMISOS_GRUPOS)}
    for codigo in codigos:
        accion = ACCIONES_PERMISO.get(codigo.split(".")[-1], codigo.split(".")[-1])
        posicion = next((indice[prefijo] for prefijo in indice if codigo.startswith(prefijo)), len(PERMISOS_GRUPOS))
        grupos[posicion][1].append((codigo, accion))
    resultado = []
    for titulo, items in grupos:
        if not items:
            continue
        items.sort(key=lambda item: (_ORDEN_ACCION.get(item[0].split(".")[-1], 99), item[0]))
        resultado.append((titulo, items))
    return resultado


def _formatear_tamano(bytes_valor):
    if bytes_valor >= 1024 * 1024:
        return f"{bytes_valor / (1024 * 1024):.1f} MB"
    if bytes_valor >= 1024:
        return f"{bytes_valor / 1024:.1f} KB"
    return f"{bytes_valor} B"


class AdminViews:
    def __init__(self, app):
        self.app = app
        self.service = app.auth_service
        self.service.inicializar_permisos(PERMISOS_BASE)

    def _importar_desde_excel(self, window, titulo, permiso, entidad, migrador, despues=None):
        if not self.app.requerir_permiso(permiso):
            return
        source = filedialog.askopenfilename(
            title=f"Importar {titulo}", parent=window,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not source:
            return
        clear_existing = messagebox.askyesno(
            "Limpiar antes de importar",
            f"¿Desactivar todos los {titulo} actuales antes de importar?\n\n"
            "Los que no estén en el Excel quedarán desactivados.",
            parent=window,
        )
        try:
            count = migrador(source, self.app.db_store, clear_existing=clear_existing)
            self.app.records_service.registrar_auditoria(
                "importar", entidad, None, self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                after={"archivo": source, "registros": count, "tabla_limpiada": clear_existing},
            )
            if despues is not None:
                despues()
            messagebox.showinfo(titulo, f"Se importaron {count} registros.", parent=window)
        except Exception as error:
            messagebox.showerror(titulo, f"No se pudo importar: {error}", parent=window)

    def mostrar_usuarios(self):
        if not self.app.requerir_permiso("usuarios.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Administrar usuarios")
        window.geometry("760x430")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "usuario", "nombre", "legajo", "activo", "roles"), show="headings")
        for column, title, width in (("id", "ID", 50), ("usuario", "Usuario", 150), ("nombre", "Nombre", 190), ("legajo", "Legajo", 80), ("activo", "Activo", 70), ("roles", "Roles", 260)):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def selected_id():
            selection = tree.selection()
            return int(tree.item(selection[0], "values")[0]) if selection else None

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.service.listar_usuarios():
                tree.insert("", "end", values=(row[0], row[1], row[2], row[3] or "", "Sí" if row[4] else "No", row[5]))

        def _dialogo_usuario(valores_iniciales=None):
            editando = valores_iniciales is not None
            dialog = tk.Toplevel(window)
            dialog.title("Editar usuario" if editando else "Nuevo usuario")
            dialog.geometry("380x260")
            self.app.aplicar_tema_ventana(dialog)
            dialog.grab_set()
            username_var = tk.StringVar(value="" if not editando else valores_iniciales["username"])
            nombre_var = tk.StringVar(value="" if not editando else valores_iniciales["nombre"])
            legajo_var = tk.StringVar(value="" if not editando else str(valores_iniciales["legajo"] or ""))
            password_var = tk.StringVar()

            ttk.Label(dialog, text="Usuario").grid(row=0, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=username_var, width=30).grid(row=0, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Nombre completo").grid(row=1, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=nombre_var, width=30).grid(row=1, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Legajo").grid(row=2, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=legajo_var, width=30).grid(row=2, column=1, sticky="ew", padx=12, pady=4)
            if not editando:
                ttk.Label(dialog, text="Contraseña").grid(row=3, column=0, sticky="w", padx=12, pady=4)
                ttk.Entry(dialog, textvariable=password_var, show="*", width=30).grid(row=3, column=1, sticky="ew", padx=12, pady=4)

            def guardar():
                username = username_var.get().strip()
                nombre = nombre_var.get().strip()
                legajo_text = legajo_var.get().strip()
                if not username:
                    messagebox.showerror("Usuarios", "El usuario es obligatorio.", parent=dialog)
                    return
                if not editando and not password_var.get():
                    messagebox.showerror("Usuarios", "La contraseña es obligatoria.", parent=dialog)
                    return
                try:
                    legajo = int(legajo_text) if legajo_text else None
                except ValueError:
                    messagebox.showerror("Usuarios", "El legajo debe ser un número.", parent=dialog)
                    return
                if not editando and legajo is None:
                    messagebox.showerror("Usuarios", "El legajo es obligatorio para nuevos usuarios.", parent=dialog)
                    return
                try:
                    if not editando:
                        self.service.crear_usuario(username, password_var.get(), nombre, legajo)
                    else:
                        self.service.actualizar_usuario(int(valores_iniciales["id"]), username, nombre, legajo, valores_iniciales["activo"])
                except Exception as error:
                    messagebox.showerror("Usuarios", str(error), parent=dialog)
                    return
                dialog.destroy()
                refresh()

            botones = ttk.Frame(dialog)
            botones.grid(row=4 if not editando else 3, column=0, columnspan=2, pady=15)
            ttk.Button(botones, text="Guardar", command=guardar).pack(side="left", padx=6)
            ttk.Button(botones, text="Cancelar", command=dialog.destroy).pack(side="left", padx=6)

        def new_user():
            _dialogo_usuario()

        def change_password():
            user_id = selected_id()
            if not user_id:
                return
            password = simpledialog.askstring("Contraseña", "Nueva contraseña:", show="*", parent=window)
            if password:
                self.service.cambiar_password(user_id, password)

        def edit_user():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            _dialogo_usuario({
                "id": int(values[0]),
                "username": values[1],
                "nombre": values[2],
                "legajo": values[3] or None,
                "activo": values[4] == "Sí",
            })

        def toggle_user():
            user_id = selected_id()
            if user_id:
                self.service.cambiar_estado_usuario(user_id)
                refresh()

        def assign_roles():
            user_id = selected_id()
            if not user_id:
                return
            roles = self.service.listar_roles()
            selected = set(self.service.roles_de_usuario(user_id))
            dialog = tk.Toplevel(window)
            dialog.title("Roles del usuario")
            dialog.configure(background=self.app.ui_background)
            variables = []
            for row in roles:
                variable = tk.BooleanVar(value=row[1] in selected)
                ttk.Checkbutton(dialog, text=row[1], variable=variable).pack(anchor="w", padx=12)
                variables.append((row[1], variable))
            def save():
                self.service.establecer_roles_usuario(user_id, [name for name, var in variables if var.get()])
                dialog.destroy()
                refresh()
            ttk.Button(dialog, text="Guardar", command=save).pack(pady=10)
            self.app.aplicar_tema_ventana(dialog)

        buttons = ttk.Frame(window)
        buttons.pack(side="bottom", fill="x", padx=10, pady=8)
        ttk.Button(buttons, text="Nuevo usuario", command=new_user).pack(side="left", padx=3)
        ttk.Button(buttons, text="Cambiar contraseña", command=change_password).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar usuario", command=edit_user).pack(side="left", padx=3)
        ttk.Button(buttons, text="Asignar roles", command=assign_roles).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=toggle_user).pack(side="left", padx=3)
        refresh()

    def mostrar_roles(self):
        if not self.app.requerir_permiso("roles.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Administrar roles y permisos")
        window.geometry("1080x430")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "rol", "permisos"), show="headings")
        for column, title, width in (("id", "ID", 35), ("rol", "Rol", 180), ("permisos", "Permisos", 820)):
            tree.heading(column, text=title)
            tree.column(column, width=width, stretch=column != "id")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.service.listar_roles():
                tree.insert("", "end", values=(row[0], row[1], row[2]))

        def new_role():
            name = simpledialog.askstring("Nuevo rol", "Nombre del rol:", parent=window)
            if name:
                try:
                    self.service.crear_rol(name)
                    refresh()
                except Exception as error:
                    messagebox.showerror("Roles", str(error), parent=window)

        def edit_permissions():
            selection = tree.selection()
            if not selection:
                return
            role_id = int(tree.item(selection[0], "values")[0])
            current = set(self.service.permisos_de_rol(role_id))
            dialog = tk.Toplevel(window)
            dialog.title("Permisos del rol")
            dialog.geometry("640x560")
            dialog.configure(background=self.app.ui_background)
            self.app.aplicar_tema_ventana(dialog)
            dialog.transient(window)
            dialog.grab_set()

            filtro_var = tk.StringVar()
            top = ttk.Frame(dialog)
            top.pack(fill="x", padx=12, pady=8)
            ttk.Label(top, text="Buscar:").pack(side="left")
            ttk.Entry(top, textvariable=filtro_var, width=38).pack(side="left", padx=6)

            canvas = tk.Canvas(dialog, highlightthickness=0, background=self.app.ui_background)
            scrollbar = ttk.Scrollbar(dialog, orient="vertical", command=canvas.yview)
            listado = ttk.Frame(canvas)
            listado.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=listado, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side="left", fill="both", expand=True, padx=(12, 0), pady=(0, 8))
            scrollbar.pack(side="right", fill="y", pady=(0, 8))

            def _rueda(event):
                canvas.yview_scroll(int(-event.delta / 120), "units")

            canvas.bind("<MouseWheel>", _rueda)
            listado.bind("<MouseWheel>", _rueda)

            variables = {}

            def rebuild():
                for child in listado.winfo_children():
                    child.destroy()
                variables.clear()
                consulta = filtro_var.get().strip().lower()
                for titulo, items in _agrupar_permisos(self.service.listar_permisos()):
                    visibles = [
                        item for item in items
                        if not consulta or consulta in item[0].lower() or consulta in item[1].lower()
                    ]
                    if not visibles:
                        continue
                    header = ttk.Frame(listado)
                    header.pack(fill="x", padx=8, pady=(10, 0))
                    ttk.Label(header, text=titulo, font=("", 10, "bold")).pack(side="left")
                    ttk.Button(header, text="Ninguno", command=lambda items=visibles: [variables[i].set(False) for i, _ in items]).pack(side="right", padx=2)
                    ttk.Button(header, text="Todos", command=lambda items=visibles: [variables[i].set(True) for i, _ in items]).pack(side="right", padx=2)
                    for codigo, accion in visibles:
                        variable = tk.BooleanVar(value=codigo in current)
                        variables[codigo] = variable
                        ttk.Checkbutton(listado, text=f"{accion} ({codigo})", variable=variable).pack(anchor="w", padx=20)

            filtro_var.trace_add("write", lambda *_args: rebuild())
            rebuild()

            def guardar():
                self.service.establecer_permisos_rol(role_id, [codigo for codigo, variable in variables.items() if variable.get()])
                dialog.destroy()
                refresh()

            bottom = ttk.Frame(dialog)
            bottom.pack(fill="x", padx=12, pady=8)
            ttk.Button(bottom, text="Guardar", command=guardar).pack(side="left", padx=3)
            ttk.Button(bottom, text="Cancelar", command=dialog.destroy).pack(side="left", padx=3)

        buttons = ttk.Frame(window)
        buttons.pack(side="bottom", fill="x", padx=10, pady=8)
        ttk.Button(buttons, text="Nuevo rol", command=new_role).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar permisos", command=edit_permissions).pack(side="left", padx=3)
        refresh()

    def mostrar_tipos_novedad(self):
        if not self.app.requerir_permiso("novedades.editar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Tipos de novedad")
        window.geometry("640x400")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "nombre", "activo"), show="headings")
        for column, title, width in (("id", "ID", 60), ("nombre", "Nombre", 300), ("activo", "Activo", 80)):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_tipos():
                tree.insert("", "end", values=(row[0], row[1], "Sí" if row[2] else "No"))

        def add_type():
            name = simpledialog.askstring("Tipo de novedad", "Nombre:", parent=window)
            if name:
                try:
                    self.app.records_service.crear_tipo(name, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.app.cargarTipoNovedades()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Tipos de novedad", str(error), parent=window)

        def edit_type():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            name = simpledialog.askstring("Tipo de novedad", "Nombre:", initialvalue=values[1], parent=window)
            if name:
                try:
                    self.app.records_service.actualizar_tipo(
                        int(values[0]), name, values[2] == "Sí",
                        self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                    )
                    self.app.cargarTipoNovedades()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Tipos de novedad", str(error), parent=window)

        def toggle_type():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            self.app.records_service.actualizar_tipo(
                int(values[0]), values[1], values[2] != "Sí",
                self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
            )
            self.app.cargarTipoNovedades()
            refresh()

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=add_type).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar", command=edit_type).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=toggle_type).pack(side="left", padx=3)
        ttk.Button(buttons, text="Importar desde Excel", command=lambda: self._importar_desde_excel(
            window, "Tipos de novedad", "novedades.editar", "tipos_novedad",
            migrate_tipos_novedad_sheet,
            despues=lambda: (self.app.cargarTipoNovedades(), refresh()),
        )).pack(side="left", padx=3)
        refresh()

    def mostrar_dotaciones(self):
        if not self.app.requerir_permiso("dotaciones.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Dotaciones")
        window.geometry("640x400")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "nombre", "activo"), show="headings")
        for column, title, width in (("id", "ID", 60), ("nombre", "Nombre", 300), ("activo", "Activo", 80)):
            tree.heading(column, text=title)
            tree.column(column, width=width, stretch=column != "id")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_dotaciones():
                tree.insert("", "end", values=(row[0], row[1], "Sí" if row[2] else "No"))

        def add():
            name = simpledialog.askstring("Dotación", "Nombre:", parent=window)
            if name:
                try:
                    self.app.records_service.crear_dotacion(name, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.app.cargarDotaciones()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Dotaciones", str(error), parent=window)

        def edit():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            name = simpledialog.askstring("Dotación", "Nombre:", initialvalue=values[1], parent=window)
            if name:
                try:
                    self.app.records_service.actualizar_dotacion(int(values[0]), name, values[2] == "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.app.cargarDotaciones()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Dotaciones", str(error), parent=window)

        def toggle():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            try:
                self.app.records_service.actualizar_dotacion(int(values[0]), values[1], values[2] != "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                self.app.cargarDotaciones()
                refresh()
            except Exception as error:
                messagebox.showerror("Dotaciones", str(error), parent=window)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=add).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar", command=edit).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=toggle).pack(side="left", padx=3)
        ttk.Button(buttons, text="Importar desde Excel", command=lambda: self._importar_desde_excel(
            window, "Dotaciones", "dotaciones.administrar", "dotaciones",
            migrate_dotaciones_sheet,
            despues=lambda: (self.app.cargarDotaciones(), refresh()),
        )).pack(side="left", padx=3)
        refresh()

    def mostrar_empleados(self):
        if not self.app.requerir_permiso("empleados.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Empleados")
        window.geometry("980x460")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "legajo", "apellidos", "especialidad", "dotacion", "turnos", "franco", "activo"), show="headings")
        for column, title, width in (
            ("id", "ID", 50), ("legajo", "Legajo", 70), ("apellidos", "Apellidos y nombres", 260),
            ("especialidad", "Especialidad", 150), ("dotacion", "Dotación", 80), ("turnos", "Turnos", 160),
            ("franco", "Franco", 100), ("activo", "Activo", 70),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width, stretch=column in {"apellidos", "especialidad", "turnos"})
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_empleados():
                tree.insert("", "end", values=(row[0], row[1], row[2], row[3] or "-", row[4] or "-", row[5] or "-", row[6] or "-", "Sí" if row[7] else "No"))

        def _dialogo(valores_iniciales=None):
            self.app.cargarDotaciones()
            dialog = tk.Toplevel(window)
            dialog.title("Nuevo empleado" if valores_iniciales is None else "Editar empleado")
            dialog.geometry("430x330")
            self.app.aplicar_tema_ventana(dialog)
            dialog.transient(window)
            dialog.grab_set()

            def opciones(lista, valor):
                opciones_lista = list(lista)
                if valor and valor not in opciones_lista:
                    opciones_lista.append(valor)
                return opciones_lista

            legajo_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("legajo") or ""))
            apellidos_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("apellidos_nombres") or ""))
            especialidad_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("especialidad") or ""))
            dotacion_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("dotacion") or ""))
            turnos_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("turnos") or ""))
            franco_var = tk.StringVar(value="" if valores_iniciales is None else str(valores_iniciales.get("franco") or ""))

            ttk.Label(dialog, text="Legajo").grid(row=0, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=legajo_var, width=30).grid(row=0, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Apellidos y nombres").grid(row=1, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=apellidos_var, width=30).grid(row=1, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Especialidad").grid(row=2, column=0, sticky="w", padx=12, pady=4)
            ttk.Combobox(dialog, textvariable=especialidad_var, values=opciones(ESPECIALIDADES_EMPLEADO, especialidad_var.get()), state="readonly", width=28).grid(row=2, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Dotación").grid(row=3, column=0, sticky="w", padx=12, pady=4)
            ttk.Combobox(dialog, textvariable=dotacion_var, values=opciones(self.app.dotaciones, dotacion_var.get()), state="readonly", width=28).grid(row=3, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Turnos").grid(row=4, column=0, sticky="w", padx=12, pady=4)
            ttk.Entry(dialog, textvariable=turnos_var, width=30).grid(row=4, column=1, sticky="ew", padx=12, pady=4)
            ttk.Label(dialog, text="Franco").grid(row=5, column=0, sticky="w", padx=12, pady=4)
            ttk.Combobox(dialog, textvariable=franco_var, values=opciones(DIAS_SEMANA, franco_var.get()), state="readonly", width=28).grid(row=5, column=1, sticky="ew", padx=12, pady=4)

            def guardar():
                data = {
                    "legajo": legajo_var.get(),
                    "apellidos_nombres": apellidos_var.get(),
                    "especialidad": especialidad_var.get(),
                    "dotacion": dotacion_var.get(),
                    "turnos": turnos_var.get(),
                    "franco": franco_var.get(),
                }
                try:
                    int(data["legajo"])
                except ValueError:
                    messagebox.showerror("Empleados", "El legajo debe ser un número.", parent=dialog)
                    return
                if not data["apellidos_nombres"].strip():
                    messagebox.showerror("Empleados", "Los apellidos y nombres son obligatorios.", parent=dialog)
                    return
                for campo, etiqueta in (("especialidad", "Especialidad"), ("dotacion", "Dotación"), ("franco", "Franco")):
                    if not data[campo]:
                        messagebox.showerror("Empleados", f"Debe elegir {etiqueta}.", parent=dialog)
                        return
                try:
                    if valores_iniciales is None:
                        self.app.records_service.crear_empleado(data, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    else:
                        self.app.records_service.actualizar_empleado(int(valores_iniciales["id"]), data, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                except Exception as error:
                    messagebox.showerror("Empleados", str(error), parent=dialog)
                    return
                self.app.actualizar_cache_base()
                self.app.cargarDotaciones()
                refresh()
                dialog.destroy()

            ttk.Button(dialog, text="Guardar", command=guardar).grid(row=6, column=0, columnspan=2, pady=15)

        def nuevo():
            _dialogo()

        def editar():
            selection = tree.selection()
            if not selection:
                return
            fila = self.app.records_service.obtener_empleado(int(tree.item(selection[0], "values")[0]))
            if not fila:
                messagebox.showerror("Empleados", "El empleado ya no existe.", parent=window)
                return
            _dialogo(dict(fila))

        def alternar():
            selection = tree.selection()
            if not selection:
                return
            empleado_id = int(tree.item(selection[0], "values")[0])
            try:
                self.app.records_service.cambiar_estado_empleado(empleado_id, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
            except Exception as error:
                messagebox.showerror("Empleados", str(error), parent=window)
                return
            self.app.actualizar_cache_base()
            refresh()

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=nuevo).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar", command=editar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=alternar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Cerrar", command=window.destroy).pack(side="right", padx=3)
        refresh()

    def mostrar_personal_estacion(self):
        if not self.app.requerir_permiso("personalEstacion.ver"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Personal de estación")
        window.geometry("720x440")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "nombre", "activo"), show="headings")
        for column, title, width in (("id", "ID", 55), ("nombre", "Nombre", 350), ("activo", "Activo", 80)):
            tree.heading(column, text=title)
            tree.column(column, width=width, stretch=column != "id")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_personal_estacion(True):
                tree.insert("", "end", values=(row[0], row[1], "Sí" if row[2] else "No"))

        def add():
            if not self.app.requerir_permiso("personalEstacion.crear"):
                return
            name = simpledialog.askstring("Personal de estación", "Nombre:", parent=window)
            if name:
                try:
                    self.app.records_service.crear_personal_estacion(name, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.app.cargarPersonalEstacion()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Personal de estación", str(error), parent=window)

        def toggle():
            if not self.app.requerir_permiso("personalEstacion.editar"):
                return
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            try:
                self.app.records_service.actualizar_personal_estacion(int(values[0]), values[1], values[2] != "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                self.app.cargarPersonalEstacion()
                refresh()
            except Exception as error:
                messagebox.showerror("Personal de estación", str(error), parent=window)

        def edit():
            if not self.app.requerir_permiso("personalEstacion.editar"):
                return
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            name = simpledialog.askstring("Personal de estación", "Nombre:", initialvalue=values[1], parent=window)
            if name:
                try:
                    self.app.records_service.actualizar_personal_estacion(int(values[0]), name, values[2] == "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.app.cargarPersonalEstacion()
                    refresh()
                except Exception as error:
                    messagebox.showerror("Personal de estación", str(error), parent=window)

        def importar():
            self._importar_desde_excel(
                window, "Personal de estación", "personalEstacion.importar", "personal_estacion",
                migrate_personal_estacion_sheet,
                despues=lambda: (self.app.cargarPersonalEstacion(), refresh()),
            )

        def exportar():
            if not self.app.requerir_permiso("personalEstacion.exportar"):
                return
            destino = filedialog.asksaveasfilename(
                title="Exportar personal de estación a Excel", parent=window,
                defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                initialfile="Personal de estación.xlsx",
            )
            if not destino:
                return
            try:
                export_database(self.app.db_store, destino, tables=["PersonalEstacion"])
                self.app.records_service.registrar_auditoria(
                    "exportar", "personal_estacion", None,
                    self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                    after={"archivo": destino},
                )
                messagebox.showinfo("Personal de estación", f"Archivo exportado correctamente:\n{destino}", parent=window)
            except PermissionError:
                messagebox.showerror("Personal de estación", "No se pudo reemplazar el Excel. Verifique que no esté abierto.", parent=window)
            except Exception as error:
                messagebox.showerror("Personal de estación", f"No se pudo exportar: {error}", parent=window)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=add).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar", command=edit).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=toggle).pack(side="left", padx=3)
        ttk.Button(buttons, text="Importar desde Excel", command=importar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Exportar a Excel", command=exportar).pack(side="left", padx=3)
        refresh()

    def mostrar_destinatarios_informe(self):
        if not self.app.requerir_permiso("destinatarios_informe.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Destinatarios de informes")
        window.geometry("720x380")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("id", "nombre", "email", "activo"), show="headings")
        for column, title, width in (("id", "ID", 55), ("nombre", "Nombre", 220), ("email", "Correo", 300), ("activo", "Activo", 80)):
            tree.heading(column, text=title)
            tree.column(column, width=width, stretch=column not in {"id", "activo"})
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_destinatarios_informe(True):
                tree.insert("", "end", values=(row[0], row[1], row[2], "Sí" if row[3] else "No"))

        def add():
            nombre = simpledialog.askstring("Destinatario", "Nombre:", parent=window) or ""
            email = simpledialog.askstring("Destinatario", "Correo:", parent=window)
            if email:
                try:
                    self.app.records_service.crear_destinatario_informe(nombre, email, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    refresh()
                except Exception as error:
                    messagebox.showerror("Destinatarios", str(error), parent=window)

        def edit():
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            nombre = simpledialog.askstring("Destinatario", "Nombre:", initialvalue=values[1], parent=window)
            email = simpledialog.askstring("Destinatario", "Correo:", initialvalue=values[2], parent=window)
            if nombre is not None and email:
                try:
                    self.app.records_service.actualizar_destinatario_informe(int(values[0]), nombre, email, values[3] == "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    refresh()
                except Exception as error:
                    messagebox.showerror("Destinatarios", str(error), parent=window)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=add).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar / activar", command=edit).pack(side="left", padx=3)
        refresh()

    def mostrar_configuracion_sesion(self):
        if not self.app.requerir_permiso("sesion.configurar"):
            return
        actual = self.app.db_store.get_configuracion("sesion_minutos", "30")
        minutos = simpledialog.askinteger(
            "Tiempo de sesión", "Minutos de inactividad antes de cerrar la sesión:",
            initialvalue=int(actual), minvalue=1, maxvalue=1440, parent=self.app.root,
        )
        if minutos is None:
            return
        self.app.db_store.set_configuracion("sesion_minutos", minutos)
        self.app.renovar_sesion()
        messagebox.showinfo("Tiempo de sesión", f"La sesión expirará luego de {minutos} minutos sin actividad.", parent=self.app.root)

    def mostrar_configuracion_tiempos(self):
        if not self.app.requerir_permiso("sesion.configurar"):
            return
        editar_horas = simpledialog.askinteger(
            "Tiempos de edición", "Horas permitidas para editar un registro desde su creación:",
            initialvalue=int(self.app.db_store.get_configuracion("editar_horas", "24")),
            minvalue=1, maxvalue=8760, parent=self.app.root,
        )
        if editar_horas is None:
            return
        eliminar_horas = simpledialog.askinteger(
            "Tiempos de edición", "Horas permitidas para eliminar un registro desde su creación:",
            initialvalue=int(self.app.db_store.get_configuracion("eliminar_horas", "72")),
            minvalue=1, maxvalue=8760, parent=self.app.root,
        )
        if eliminar_horas is None:
            return
        self.app.db_store.set_configuracion("editar_horas", editar_horas)
        self.app.db_store.set_configuracion("eliminar_horas", eliminar_horas)
        messagebox.showinfo(
            "Tiempos de edición",
            f"Se podrá editar dentro de {editar_horas} horas y eliminar dentro de {eliminar_horas} horas.",
            parent=self.app.root,
        )

    def mostrar_configuracion_notificaciones(self):
        if not self.app.requerir_permiso("sesion.configurar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Notificaciones")
        window.geometry("480x250")
        window.resizable(False, False)
        self.app.aplicar_tema_ventana(window)
        window.transient(self.app.root)
        window.grab_set()
        activo = tk.BooleanVar(
            value=str(self.app.db_store.get_configuracion("notificaciones_activo", "1")) == "1"
        )
        try:
            duracion = int(self.app.db_store.get_configuracion("toast_duracion", "6"))
        except (TypeError, ValueError):
            duracion = 6
        duracion_var = tk.IntVar(value=max(1, duracion))
        form = ttk.Frame(window)
        form.pack(fill="x", padx=20, pady=18)
        ttk.Label(
            form,
            text=("Notificar registros (novedades y cambios de turno)\n"
                  "cargados por otros usuarios, desde su último ingreso."),
            justify="left",
        ).pack(anchor="w")
        ttk.Checkbutton(form, text="Activar notificaciones", variable=activo).pack(anchor="w", pady=(12, 4))
        ttk.Label(form, text="Duración del toast (segundos)").pack(anchor="w", pady=(10, 2))
        ttk.Spinbox(form, from_=1, to=60, textvariable=duracion_var, width=8).pack(anchor="w")

        def guardar():
            try:
                segundos = duracion_var.get()
            except (TypeError, ValueError):
                segundos = 6
            if segundos < 1:
                segundos = 6
            self.app.db_store.set_configuracion("notificaciones_activo", 1 if activo.get() else 0)
            self.app.db_store.set_configuracion("toast_duracion", segundos)
            window.destroy()
            estado = "activadas" if activo.get() else "desactivadas"
            messagebox.showinfo(
                "Notificaciones",
                f"Notificaciones {estado}. El toast se mostrará durante {segundos} segundos.",
                parent=self.app.root,
            )

        bottom = ttk.Frame(window)
        bottom.pack(fill="x", padx=20, pady=(0, 16))
        ttk.Button(bottom, text="Guardar", command=guardar).pack(side="left", padx=3)
        ttk.Button(bottom, text="Cancelar", command=window.destroy).pack(side="left", padx=3)

    def mostrar_registros_eliminados(self):
        if not self.app.requerir_permiso("registros.recuperar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Registros eliminados")
        window.geometry("1100x480")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("tipo", "id", "fecha", "legajo", "empleado", "dotacion", "detalle", "observaciones"), show="headings")
        for column, title, width in (
            ("tipo", "Tipo", 80), ("id", "ID", 50), ("fecha", "Fecha de registro", 140),
            ("legajo", "Legajo", 70), ("empleado", "Empleado", 220), ("dotacion", "Dotación", 80),
            ("detalle", "Detalle", 180), ("observaciones", "Observaciones", 240),
        ):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_eliminados():
                tree.insert("", "end", values=tuple("-" if value is None else value for value in row))

        def selected():
            selection = tree.selection()
            return (tree.item(selection[0], "values")[0], int(tree.item(selection[0], "values")[1])) if selection else None

        def recuperar():
            if not selected():
                return
            tipo, record_id = selected()
            try:
                self.app.records_service.recuperar_registro(tipo, record_id, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                self.app.cargar_excel()
                refresh()
            except Exception as error:
                messagebox.showerror("Recuperar", str(error), parent=window)

        def borrar_definitivo():
            if not selected():
                return
            tipo, record_id = selected()
            if not messagebox.askyesno(
                "Borrado definitivo",
                f"¿Eliminar definitivamente el registro #{record_id}?\n\nEsta acción no se puede deshacer.",
                parent=window,
            ):
                return
            try:
                self.app.records_service.borrar_definitivo(tipo, record_id, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                refresh()
            except Exception as error:
                messagebox.showerror("Borrado definitivo", str(error), parent=window)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=8)
        ttk.Button(buttons, text="Recuperar", command=recuperar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Borrar definitivamente", command=borrar_definitivo).pack(side="left", padx=3)
        ttk.Button(buttons, text="Cerrar", command=window.destroy).pack(side="right", padx=3)
        refresh()

    def mostrar_backups(self):
        if not self.app.requerir_permiso("backup.gestionar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Copias de seguridad")
        window.geometry("720x420")
        self.app.aplicar_tema_ventana(window)
        tree = ttk.Treeview(window, columns=("nombre", "tamano", "fecha"), show="headings")
        for column, title, width in (("nombre", "Archivo", 300), ("tamano", "Tamaño", 110), ("fecha", "Fecha", 160)):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())
            for nombre, tamano in listar_backups(self.app.db_store):
                fecha = nombre.replace("backup_", "").replace(".sqlite", "").replace("_", " ")
                tree.insert("", "end", values=(nombre, _formatear_tamano(tamano), fecha))

        def crear_copia():
            try:
                destino = crear_backup(self.app.db_store)
                self.app.records_service.registrar_auditoria(
                    "crear", "backup", None, self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                    after={"archivo": destino},
                )
                refresh()
                messagebox.showinfo("Copias de seguridad", "Copia de seguridad creada correctamente.", parent=window)
            except Exception as error:
                messagebox.showerror("Copias de seguridad", f"No se pudo crear la copia: {error}", parent=window)

        def restaurar():
            selection = tree.selection()
            if not selection:
                return
            nombre = tree.item(selection[0], "values")[0]
            if not messagebox.askyesno(
                "Restaurar copia",
                f"¿Restaurar la base de datos desde {nombre}?\n\n"
                "Los registros actuales serán reemplazados por los de la copia.",
                parent=window,
            ):
                return
            try:
                restaurar_backup(self.app.db_store, nombre)
                self.app.records_service.registrar_auditoria(
                    "restaurar", "backup", None, self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                    after={"archivo": nombre},
                )
                messagebox.showinfo(
                    "Restaurar copia",
                    "La base fue restaurada. Cierre y vuelva a abrir la aplicación para aplicar los cambios.",
                    parent=window,
                )
                window.destroy()
            except Exception as error:
                messagebox.showerror("Restaurar copia", f"No se pudo restaurar: {error}", parent=window)

        def borrar_copia():
            selection = tree.selection()
            if not selection:
                return
            nombre = tree.item(selection[0], "values")[0]
            if not messagebox.askyesno(
                "Borrar copia", f"¿Eliminar la copia {nombre}?", parent=window
            ):
                return
            try:
                ruta = os.path.join(os.path.dirname(self.app.db_store.database_path), "backups", nombre)
                os.remove(ruta)
                refresh()
            except OSError as error:
                messagebox.showerror("Borrar copia", f"No se pudo borrar la copia: {error}", parent=window)

        def configurar():
            try:
                retencion = simpledialog.askinteger(
                    "Copias de seguridad",
                    "Cantidad de copias a conservar:",
                    initialvalue=int(self.app.db_store.get_configuracion("backup_retencion", "10")),
                    minvalue=1, maxvalue=365, parent=window,
                )
                if retencion is None:
                    return
                self.app.db_store.set_configuracion("backup_retencion", retencion)
                activo = messagebox.askyesno(
                    "Copias de seguridad",
                    "¿Crear una copia automática diaria al iniciar la aplicación?",
                    parent=window,
                )
                self.app.db_store.set_configuracion("backup_activo", 1 if activo else 0)
                messagebox.showinfo(
                    "Copias de seguridad",
                    f"Se conservarán {retencion} copias y el backup automático está "
                    f"{'activado' if activo else 'desactivado'}.",
                    parent=window,
                )
            except Exception as error:
                messagebox.showerror("Copias de seguridad", str(error), parent=window)

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=8)
        ttk.Button(buttons, text="Hacer copia ahora", command=crear_copia).pack(side="left", padx=3)
        ttk.Button(buttons, text="Restaurar", command=restaurar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Borrar copia", command=borrar_copia).pack(side="left", padx=3)
        ttk.Button(buttons, text="Configurar", command=configurar).pack(side="left", padx=3)
        ttk.Button(buttons, text="Cerrar", command=window.destroy).pack(side="right", padx=3)
        refresh()

    def mostrar_auditoria(self):
        if not self.app.requerir_permiso("auditoria.ver"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Auditoría")
        window.geometry("1050x500")
        self.app.aplicar_tema_ventana(window)
        top = ttk.Frame(window)
        top.pack(fill="x", padx=10, pady=8)
        filter_var = tk.StringVar()
        ttk.Label(top, text="Buscar usuario, acción o entidad:").pack(side="left")
        ttk.Entry(top, textvariable=filter_var, width=35).pack(side="left", padx=8)
        tree = ttk.Treeview(window, columns=("id", "fecha", "usuario", "accion", "entidad", "registro", "antes", "despues"), show="headings")
        for column, title, width in (("id", "ID", 50), ("fecha", "Fecha", 145), ("usuario", "Usuario", 120), ("accion", "Acción", 90), ("entidad", "Entidad", 120), ("registro", "Registro", 70), ("antes", "Antes", 210), ("despues", "Después", 210)):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=5)

        def refresh(*_args):
            tree.delete(*tree.get_children())
            for row in self.app.records_service.listar_auditoria(filter_var.get()):
                tree.insert("", "end", values=tuple("-" if value is None else value for value in row))

        def export():
            destination = filedialog.asksaveasfilename(
                title="Exportar auditoría", parent=window, defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")], initialfile="Auditoria.xlsx",
            )
            if not destination:
                return
            try:
                export_auditoria(self.app.db_store, destination, filter_var.get())
                messagebox.showinfo("Auditoría", "La auditoría se exportó correctamente.", parent=window)
            except PermissionError:
                messagebox.showerror("Auditoría", "No se pudo guardar el archivo. Verifique que no esté abierto.", parent=window)
            except Exception as error:
                messagebox.showerror("Auditoría", str(error), parent=window)

        filter_var.trace_add("write", refresh)
        ttk.Button(top, text="Exportar", command=export).pack(side="left", padx=8)
        refresh()
