"""Vistas de administración de usuarios, roles y permisos."""

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from excel_exporter import export_auditoria


PERMISOS_BASE = (
    "novedades.ver", "novedades.crear", "novedades.editar", "novedades.eliminar",
    "cambios_turno.ver", "cambios_turno.crear", "cambios_turno.editar", "excel.exportar",
    "usuarios.administrar", "roles.administrar", "empleados.importar", "auditoria.ver",
    "dotaciones.administrar",
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
        tree = ttk.Treeview(window, columns=("id", "usuario", "nombre", "activo", "roles"), show="headings")
        for column, title, width in (("id", "ID", 50), ("usuario", "Usuario", 150), ("nombre", "Nombre", 190), ("activo", "Activo", 70), ("roles", "Roles", 260)):
            tree.heading(column, text=title)
            tree.column(column, width=width)
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def selected_id():
            selection = tree.selection()
            return int(tree.item(selection[0], "values")[0]) if selection else None

        def refresh():
            tree.delete(*tree.get_children())
            for row in self.service.listar_usuarios():
                tree.insert("", "end", values=(row[0], row[1], row[2], "Sí" if row[3] else "No", row[4]))

        def new_user():
            username = simpledialog.askstring("Nuevo usuario", "Usuario:", parent=window)
            if not username:
                return
            nombre = simpledialog.askstring("Nuevo usuario", "Nombre completo:", parent=window) or ""
            password = simpledialog.askstring("Nuevo usuario", "Contraseña:", show="*", parent=window)
            if not password:
                return
            try:
                self.service.crear_usuario(username, password, nombre)
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
            for row in self.service.listar_dotaciones():
                tree.insert("", "end", values=(row[0], row[1], "Sí" if row[2] else "No"))

        def add():
            name = simpledialog.askstring("Dotación", "Nombre:", parent=window)
            if name:
                try:
                    self.service.crear_dotacion(name, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
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
                    self.service.actualizar_dotacion(int(values[0]), name, values[2] == "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
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
                self.service.actualizar_dotacion(int(values[0]), values[1], values[2] != "Sí", self.app.current_user.get("id"), self.app.obtener_usuario_windows())
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
