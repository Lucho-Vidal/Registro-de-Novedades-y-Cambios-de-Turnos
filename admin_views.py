"""Vistas de administración de usuarios, roles y permisos."""

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from excel_exporter import export_auditoria


PERMISOS_BASE = (
    "novedades.ver", "novedades.crear", "novedades.editar", "novedades.eliminar",
    "cambios_turno.ver", "cambios_turno.crear", "cambios_turno.editar", "cambios_turno.eliminar",
    "excel.exportar",
    "usuarios.administrar", "roles.administrar", "empleados.importar", "auditoria.ver",
    "dotaciones.administrar",
    "personalEstacion.ver", "personalEstacion.crear", "personalEstacion.editar",
    "destinatarios_informe.administrar", "sesion.configurar", "registros.recuperar",
)


class AdminViews:
    def __init__(self, app):
        self.app = app
        self.service = app.auth_service
        self.service.inicializar_permisos(PERMISOS_BASE)

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

        def new_user():
            username = simpledialog.askstring("Nuevo usuario", "Usuario:", parent=window)
            if not username:
                return
            nombre = simpledialog.askstring("Nuevo usuario", "Nombre completo:", parent=window) or ""
            legajo = simpledialog.askstring("Nuevo usuario", "Legajo:", parent=window) or ""
            password = simpledialog.askstring("Nuevo usuario", "Contraseña:", show="*", parent=window)
            if not password:
                return
            try:
                self.service.crear_usuario(username, password, nombre, int(legajo) if legajo else None)
                refresh()
            except Exception as error:
                messagebox.showerror("Usuarios", str(error), parent=window)

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
            username = simpledialog.askstring("Usuario", "Usuario:", initialvalue=values[1], parent=window)
            nombre = simpledialog.askstring("Usuario", "Nombre completo:", initialvalue=values[2], parent=window)
            legajo = simpledialog.askstring("Usuario", "Legajo:", initialvalue=values[3], parent=window)
            if username and nombre is not None and legajo is not None:
                try:
                    self.service.actualizar_usuario(int(values[0]), username, nombre, int(legajo) if legajo.strip() else None, values[4] == "Sí")
                    refresh()
                except Exception as error:
                    messagebox.showerror("Usuarios", str(error), parent=window)

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
            dialog.configure(background=self.app.ui_background)
            variables = []
            for permission in self.service.listar_permisos():
                variable = tk.BooleanVar(value=permission in current)
                ttk.Checkbutton(dialog, text=permission, variable=variable).pack(anchor="w", padx=12)
                variables.append((permission, variable))
            def save():
                self.service.establecer_permisos_rol(role_id, [name for name, var in variables if var.get()])
                dialog.destroy()
                refresh()
            ttk.Button(dialog, text="Guardar", command=save).pack(pady=10)
            self.app.aplicar_tema_ventana(dialog)

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
        window.geometry("520x360")
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
        refresh()

    def mostrar_dotaciones(self):
        if not self.app.requerir_permiso("dotaciones.administrar"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Dotaciones")
        window.geometry("520x360")
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
        refresh()

    def mostrar_personal_estacion(self):
        if not self.app.requerir_permiso("personalEstacion.ver"):
            return
        window = tk.Toplevel(self.app.root)
        window.title("Personal de estación")
        window.geometry("560x360")
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

        buttons = ttk.Frame(window)
        buttons.pack(fill="x", padx=10, pady=5)
        ttk.Button(buttons, text="Nuevo", command=add).pack(side="left", padx=3)
        ttk.Button(buttons, text="Editar", command=edit).pack(side="left", padx=3)
        ttk.Button(buttons, text="Activar / desactivar", command=toggle).pack(side="left", padx=3)
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
