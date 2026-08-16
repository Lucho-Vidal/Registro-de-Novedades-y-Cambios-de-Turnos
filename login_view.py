"""Pantalla de inicio de sesión y alta del administrador inicial."""

import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from ttkbootstrap import Style

from auth import AuthService


class LoginView:
    def __init__(self, root, store, on_authenticated):
        self.root = root
        self.store = store
        self.auth = AuthService(store)
        self.on_authenticated = on_authenticated
        self.style = self._apply_saved_theme()
        self.window = tk.Toplevel(root)
        self.window.configure(background=str(self.style.colors.bg))
        self.window.title("Inicio de sesión")
        self.window.geometry("420x375")
        self.window.minsize(420, 375)
        self.window.resizable(False, False)
        self.window.protocol("WM_DELETE_WINDOW", root.destroy)
        self.window.grab_set()

        # Centra el contenido verticalmente; la altura fija mantiene el
        # espacio reservado para el campo de legajo aunque no se muestre.
        self.window.grid_rowconfigure(0, weight=1)
        self.window.grid_rowconfigure(2, weight=1)
        self.window.grid_columnconfigure(0, weight=1)
        content = ttk.Frame(self.window)
        content.grid(row=1, column=0)

        ttk.Label(content, text="Registro de novedades", font=("Helvetica", 18, "bold")).pack(pady=(0, 14))
        form = ttk.Frame(content)
        form.pack(fill="x", padx=35)
        ttk.Label(form, text="Usuario").grid(row=0, column=0, sticky="w", pady=5)
        self.username = ttk.Entry(form, width=30)
        self.username.grid(row=1, column=0, pady=2)
        ttk.Label(form, text="Contraseña").grid(row=2, column=0, sticky="w", pady=(8, 5))
        self.password = ttk.Entry(form, show="*", width=30)
        self.password.grid(row=3, column=0, pady=2)

        self.legajo_label = None
        self.legajo = None
        self.admin_button = None

        ttk.Button(content, text="Ingresar", command=self.login).pack(pady=(14, 5))
        self.username.focus_set()
        self.password.bind("<Return>", lambda _event: self.login())

        if not self._hay_usuarios():
            self.legajo_label = ttk.Label(form, text="Legajo (solo para crear administrador)")
            self.legajo_label.grid(row=4, column=0, sticky="w", pady=(8, 5))
            self.legajo = ttk.Entry(form, width=30)
            self.legajo.grid(row=5, column=0, pady=2)
            self.admin_button = ttk.Button(
                content, text="Crear administrador inicial",
                command=self.crear_administrador_inicial,
            )
            self.admin_button.pack(pady=(0, 5))

    def _apply_saved_theme(self):
        try:
            theme = open("theme", "r", encoding="utf-8").read().strip()
        except OSError:
            theme = "flatly"
        style = Style()
        if theme not in set(style.theme_names()):
            theme = {
                "nord-dark": "darkly",
                "nord-light": "cosmo",
                "bootstrap-dark": "darkly",
                "bootstrap-light": "flatly",
                "sandstone-dark": "superhero",
                "sandstone-light": "sandstone",
            }.get(theme, "flatly")
        if theme not in set(style.theme_names()):
            theme = "flatly"
        style.theme_use(theme)
        background = str(style.colors.bg)
        foreground = str(style.colors.fg)
        style.configure("TFrame", background=background)
        style.configure("TLabel", background=background, foreground=foreground)
        self.root.configure(background=background)
        return style

    def _hay_usuarios(self):
        with self.store.read_connection() as connection:
            return connection.execute("SELECT 1 FROM usuarios LIMIT 1").fetchone() is not None

    def login(self):
        user = self.auth.autenticar(self.username.get(), self.password.get())
        if not user:
            messagebox.showerror("Inicio de sesión", "Usuario o contraseña incorrectos.", parent=self.window)
            return
        self.window.grab_release()
        self.window.destroy()
        self.root.deiconify()
        self.on_authenticated(user, self.store)

    def crear_administrador_inicial(self):
        username = self.username.get().strip()
        password = self.password.get()
        legajo = self.legajo.get().strip()
        if not username or not password or not legajo:
            messagebox.showwarning("Administrador", "Complete usuario, contraseña y legajo.", parent=self.window)
            return
        try:
            user_id = self.auth.crear_administrador_inicial(username, password, int(legajo))
            messagebox.showinfo("Administrador", "Administrador creado. Ya puede ingresar.", parent=self.window)
            if self.legajo_label is not None:
                self.legajo_label.destroy()
                self.legajo_label = None
            if self.legajo is not None:
                self.legajo.destroy()
                self.legajo = None
            if self.admin_button is not None:
                self.admin_button.destroy()
                self.admin_button = None
            self.username.delete(0, tk.END)
            self.password.delete(0, tk.END)
            self.username.insert(0, username)
            self.username.focus_set()
            self.window.update_idletasks()
        except Exception as error:
            messagebox.showerror("Administrador", str(error), parent=self.window)