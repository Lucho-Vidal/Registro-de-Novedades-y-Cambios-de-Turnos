"""Lógica de tablas, filtrado y vistas de datos."""

import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText


class TablesManager:
    """Gestor de tablas, filtrado y modales de detalle.
    
    Maneja la creación de tablas Treeview para novedades y cambios de turnos,
    filtrado con debounce (250ms) por nombre/dotación, carga de datos desde Excel,
    modales de detalle (vista de registros), y actualización de contadores.
    
    Attributes:
        app: Referencia a la aplicación FormularioExcelApp con acceso a
             sheet_novedades, sheet_cambio_turnos, etc.
    
    Methods:
        crear_tabla_novedades(): Crea Treeview para hoja NOVEDADES.
        crear_tabla_cambios(): Crea Treeview para hoja Cambio de Turnos.
        cargar_datos_completos_novedades(): Carga datos sin filtros.
        cargar_datos_completos_cambios(): Carga datos sin filtros.
        filtrar_datos_novedades(nombre_filtro, dotacion_filtro): Filtra por criterios.
        filtrar_datos_cambios(nombre_filtro, dotacion_filtro): Filtra por criterios.
        mostrar_modal_detalle(novedad, columnas, vista): Muestra detalles de un registro.
        actualizar_tabla(): Recarga datos (llamado cada 60s).
    """
    
    def __init__(self, app):
        """Inicializa el gestor de tablas.
        
        Args:
            app: Instancia de FormularioExcelApp con acceso a workbook, sheets, etc.
        """
        self.app = app
    
    def actualizar_contador_resultados(self, label, cantidad, sufijo=""):
        """Actualiza la etiqueta de cantidad de resultados.
        
        Modifica el texto de un ttk.Label para mostrar cantidad de resultados
        con pluralización correcta (ej. "1 resultado" vs "3 resultados").
        
        Args:
            label: ttk.Label a actualizar con el texto del contador.
            cantidad: Número de resultados a mostrar.
            sufijo: Texto adicional a agregar al final (default="").
        
        Returns:
            None (Modifica label.config()).
        
        Example:
            >>> actualizar_contador_resultados(label, 5)
            # Label ahora muestra "5 resultados"
            >>> actualizar_contador_resultados(label, 1, " - Filtrado")
            # Label ahora muestra "1 resultado - Filtrado"
        """
        if cantidad == 1:
            texto = "1 resultado"
        else:
            texto = f"{cantidad} resultados"
        label.config(text=f"{texto}{sufijo}")

    def ajustar_ancho_columnas(self, tree):
        """Ajusta cada columna al contenido visible sin repartirlas por igual.

        Se aplica después de cargar o filtrar una tabla. Se mantienen límites
        razonables para que un texto extenso no vuelva inutilizable la vista;
        las columnas se pueden recorrer con la barra horizontal.
        """
        columnas = list(tree["columns"])
        for indice, columna in enumerate(columnas):
            encabezado = str(tree.heading(columna, "text") or columna)
            maximo = len(encabezado)
            for item in tree.get_children(""):
                valores = tree.item(item, "values")
                if indice < len(valores):
                    maximo = max(maximo, len(str(valores[indice] or "")))

            # Aproximación de píxeles basada en el ancho medio de los caracteres.
            ancho = max(45, min(320, maximo * 7 + 24))
            if columna.strip().upper() == "ID":
                ancho = min(ancho, 65)
            tree.column(columna, width=ancho, minwidth=35, stretch=False)
    
    def mostrar_modal_detalle(self, novedad, columnas, vista):
        """Muestra un modal con los detalles de un registro.
        
        Crea un modal (Toplevel) que presenta todos los campos de una novedad
        o cambio de turno. El campo "Observaciones" se muestra como ScrolledText
        de lectura (read-only), otros campos como Labels.
        
        Args:
            novedad: Tupla o lista con valores del registro (en orden de columnas).
            columnas: Lista de nombres de columnas. Ejemplo: ["ID", "Fecha", ..., "Observaciones"].
            vista: Tipo de registro ("novedad" o "cambio" - usado para título y tamaño modal).
        
        Returns:
            None (Crea y muestra modal).
        
        Example:
            >>> mostrar_modal_detalle(
            ...     ("1", "15/06/2025", "...", "Licencia aprobada"),
            ...     ["ID", "Fecha", ..., "Observaciones"],
            ...     "novedad"
            ... )
        """
        modal = tk.Toplevel(self.app.root)
        modal.title(f"Detalle de {vista}")
        modal.configure(background=self.app.ui_background)
        if vista == "novedad":
            modal.geometry("620x650")
        else:
            modal.geometry("700x780")
        row_button = len(columnas) + 1
        
        ttk.Label(modal, text=f"Detalle de {vista}", font=("Helvetica", 22, "bold")).grid(
            row=0, column=0, columnspan=2, pady=10, padx=10, sticky="we"
        )

        for idx, (columna, valor) in enumerate(zip(columnas, novedad), start=1):
            ttk.Label(modal, text=(columna+":").capitalize()).grid(
                row=idx, column=0, pady=2, padx=10, sticky="w"
            )
            if columna == "Observaciones":
                text_area = ScrolledText(modal, wrap=tk.WORD, height=5, width=60)
                text_area.insert("1.0", valor)
                text_area.config(state="disabled")
                text_area.configure(
                    background=self.app.ui_background,
                    foreground=self.app.ui_foreground,
                    insertbackground=self.app.ui_foreground,
                )
                text_area.grid(row=idx, column=1, pady=2, padx=10, sticky="w")
            else:
                ttk.Label(modal, text=valor).grid(row=idx, column=1, pady=2, padx=10, sticky="w")
        
        ttk.Button(modal, text="Cerrar", command=modal.destroy).grid(
            row=row_button, column=0, pady=10, padx=10
        )
        if vista == "novedad" and self.app.tiene_permiso("novedades.editar"):
            ttk.Button(modal, text="Editar", command=lambda: (modal.destroy(), self.mostrar_editor(int(novedad[0]), "novedad"))).grid(
                row=row_button, column=1, pady=10, padx=10
            )
        elif vista != "novedad" and self.app.tiene_permiso("cambios_turno.editar"):
            ttk.Button(modal, text="Editar", command=lambda: (modal.destroy(), self.mostrar_editor(int(novedad[0]), "cambio"))).grid(
                row=row_button, column=1, pady=10, padx=10
            )

    def mostrar_editor(self, record_id, record_type):
        """Edita un registro y registra el antes/después en auditoría."""
        if record_type == "novedad":
            row = self.app.records_service.obtener_novedad(record_id)
            fields = [
                ("legajo", "Legajo"), ("apellidos_nombres", "Apellidos y nombres"),
                ("especialidad", "Especialidad"), ("dotacion", "Dotación"), ("turnos", "Turnos"),
                ("franco", "Franco"), ("novedad", "Tipo de novedad"), ("fecha_inicio", "Fecha inicio"),
                ("fecha_fin", "Fecha fin"), ("referencia_estacion", "Referencia estación"),
                ("supervisor", "Supervisor"), ("observaciones", "Observaciones"),
            ]
        else:
            row = self.app.records_service.obtener_cambio(record_id)
            fields = [
                ("legajo_1", "Legajo 1"), ("apellidos_nombres_1", "Apellidos y nombres 1"),
                ("especialidad_1", "Especialidad 1"), ("dotacion_1", "Dotación 1"),
                ("turnos_1", "Turnos 1"), ("franco_1", "Franco 1"), ("legajo_2", "Legajo 2"),
                ("apellidos_nombres_2", "Apellidos y nombres 2"), ("especialidad_2", "Especialidad 2"),
                ("dotacion_2", "Dotación 2"), ("turnos_2", "Turnos 2"), ("franco_2", "Franco 2"),
                ("fecha_cambio", "Fecha cambio"), ("referencia_estacion", "Referencia estación"),
                ("supervisor", "Supervisor"), ("observaciones", "Observaciones"),
            ]
        if not row:
            messagebox.showerror("Edición", "El registro ya no existe.", parent=self.app.root)
            return
        window = tk.Toplevel(self.app.root)
        window.title("Editar registro")
        window.geometry("720x680")
        self.app.aplicar_tema_ventana(window)
        variables = {}
        widgets = {}
        for index, (field, label) in enumerate(fields):
            ttk.Label(window, text=label).grid(row=index, column=0, sticky="w", padx=12, pady=4)
            value = "" if row[field] is None else str(row[field])
            variable = tk.StringVar(value=value)
            variables[field] = variable
            if field == "novedad":
                widget = ttk.Combobox(window, textvariable=variable, values=self.app.tipo_novedades, width=42)
            elif field == "observaciones":
                widget = tk.Text(window, height=4, width=48)
                widget.insert("1.0", value)
            else:
                widget = ttk.Entry(window, textvariable=variable, width=45)
            widget.grid(row=index, column=1, sticky="ew", padx=12, pady=4)
            widgets[field] = widget

        def save():
            data = {}
            for field, _label in fields:
                if field == "observaciones":
                    data[field] = widgets[field].get("1.0", "end-1c")
                else:
                    data[field] = variables[field].get().strip()
            for field in ("legajo", "legajo_1", "legajo_2"):
                if field in data:
                    try:
                        data[field] = int(data[field])
                    except ValueError:
                        messagebox.showerror("Edición", f"{field} debe ser numérico.", parent=window)
                        return
            try:
                if record_type == "novedad":
                    self.app.records_service.actualizar_novedad(record_id, data, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.cargar_datos_completos_novedades()
                else:
                    self.app.records_service.actualizar_cambio(record_id, data, self.app.current_user.get("id"), self.app.obtener_usuario_windows())
                    self.cargar_datos_completos_cambios()
                window.destroy()
            except Exception as error:
                messagebox.showerror("Edición", str(error), parent=window)

        ttk.Button(window, text="Guardar cambios", command=save).grid(row=len(fields), column=0, columnspan=2, pady=15)
        self.app.aplicar_tema_ventana(window)
    
    def cargar_datos_completos_novedades(self):
        """Carga todos los datos de novedades en la tabla sin filtros.
        
        Lee todas las filas (desde row 2 en adelante) de la hoja NOVEDADES,
        normaliza valores None a "-", inserta en Treeview, y actualiza contador.
        
        Returns:
            None (Modifica tabla_novedades y resultados_novedades_label).
        
        Side Effects:
            - Limpia tabla_novedades
            - Inserta filas desde sheet_novedades
            - Actualiza contador de resultados
        """
        try:
            self.app.tabla_novedades.delete(*self.app.tabla_novedades.get_children())
            total = 0
            if self.app.sheet_novedades:
                for fila in self.app.sheet_novedades.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    self.app.tabla_novedades.insert("", "end", values=fila_procesada)
                    total += 1
            else:
                print("La hoja 'NOVEDADES' está vacía o no se pudo cargar.")

            self.ajustar_ancho_columnas(self.app.tabla_novedades)
            if hasattr(self.app, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.app.resultados_novedades_label, total)
        except Exception as e:
            print(f"Error al cargar los datos en el Treeview: {e}")

    def cargar_datos_completos_cambios(self):
        """Carga todos los datos de cambios de turnos en la tabla sin filtros.
        
        Lee todas las filas (desde row 2 en adelante) de la hoja Cambio de Turnos,
        normaliza valores None a "-", inserta en Treeview, y actualiza contador.
        
        Returns:
            None (Modifica table_cambios y resultados_cambios_label).
        
        Side Effects:
            - Limpia table_cambios
            - Inserta filas desde sheet_cambio_turnos
            - Actualiza contador de resultados
        """
        try:
            self.app.table_cambios.delete(*self.app.table_cambios.get_children())
            total = 0
            if self.app.sheet_cambio_turnos:
                for fila in self.app.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    self.app.table_cambios.insert("", "end", values=fila_procesada)
                    total += 1
            else:
                print("La hoja 'Cambio de Turnos' está vacía o no se pudo cargar.")

            self.ajustar_ancho_columnas(self.app.table_cambios)
            if hasattr(self.app, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.app.resultados_cambios_label, total)
        except Exception as e:
            print(f"Error al cargar los datos en el Treeview: {e}")
    
    def filtrar_datos_novedades(self, nombre_filtro, dotacion_filtro, tipo_filtro):
        """Filtra los datos de novedades según nombre y dotación.
        
        Busca coincidencias en la columna Apellidos y Nombres (índice 3) y
        Dotación (índice 5). Utiliza búsqueda case-insensitive y sin acentos
        (normalización unicode). Si ambos filtros están en valores por defecto,
        carga datos completos.
        
        Args:
            nombre_filtro: Cadena de búsqueda por apellido/nombre.
            dotacion_filtro: Dotación a filtrar ("Todas" = sin filtro).
        
        Returns:
            None (Modifica tabla_novedades).
        
        Side Effects:
            - Limpia tabla_novedades
            - Inserta filas filtradas desde sheet_novedades
            - Actualiza contador de resultados
        
        Example:
            >>> filtrar_datos_novedades("García", "A")
            # Muestra solo novedades de empleados con "García" y dotación "A"
        """
        try:
            if dotacion_filtro == "Todas" and tipo_filtro == "Todos" and nombre_filtro == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                self.cargar_datos_completos_novedades()
                return

            nombre_filtro_norm = self.app.normalizar_texto(nombre_filtro)
            dotacion_filtro_norm = self.app.normalizar_texto(dotacion_filtro)
            tipo_filtro_norm = self.app.normalizar_texto(tipo_filtro)

            self.app.tabla_novedades.delete(*self.app.tabla_novedades.get_children())
            total = 0
            if self.app.sheet_novedades:
                for fila in self.app.sheet_novedades.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    if len(fila_procesada) <= 5:
                        continue
                    nombre_fila_norm = self.app.normalizar_texto(fila_procesada[3])
                    dotacion_fila_norm = self.app.normalizar_texto(fila_procesada[5])
                    tipo_fila_norm = self.app.normalizar_texto(fila_procesada[8])
                    coincide_nombre = nombre_filtro == self.app.PLACEHOLDER_BUSCAR_NOMBRE or nombre_filtro_norm in nombre_fila_norm
                    coincide_dotacion = dotacion_filtro == "Todas" or dotacion_filtro_norm in dotacion_fila_norm
                    coincide_tipo = tipo_filtro == "Todos" or tipo_filtro_norm == tipo_fila_norm
                    if coincide_nombre and coincide_dotacion and coincide_tipo:
                        self.app.tabla_novedades.insert("", "end", values=fila_procesada)
                        total += 1

            self.ajustar_ancho_columnas(self.app.tabla_novedades)
            if hasattr(self.app, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.app.resultados_novedades_label, total)
        except Exception as e:
            print(f"Error al filtrar los datos en el Treeview novedades: {e}")

    def programar_filtrado_novedades(self):
        """Programa el filtrado de novedades con debounce de 250ms.
        
        Cancela cualquier filtrado pendiente y agenda uno nuevo con delay 250ms.
        Esto evita filtrar a cada pulsación de tecla, mejorando rendimiento.
        
        Returns:
            None (Modifica app.filtro_after_novedades).
        
        Side Effects:
            - Cancela task anterior si existe
            - Programa nueva tarea en root.after()
        """
        if self.app.filtro_after_novedades:
            self.app.root.after_cancel(self.app.filtro_after_novedades)
        self.app.filtro_after_novedades = self.app.root.after(
            250,
            lambda: self.filtrar_datos_novedades(
                self.app.apellido_filter_novedades_var.get(),
                self.app.dotacion_filter_novedades_var.get(),
                self.app.tipo_filter_novedades_var.get()
            )
        )

    def filtrar_datos_cambios(self, nombre_filtro, dotacion_filtro):
        """Filtra los datos de cambios de turnos según nombre y dotación.
        
        Busca coincidencias en cualquiera de los dos empleados (apellidos índices 3 y 9)
        y en cualquiera de las dos dotaciones (índices 5 y 11).
        Utiliza búsqueda case-insensitive y sin acentos (normalización unicode).
        Si ambos filtros están en valores por defecto, carga datos completos.
        
        Args:
            nombre_filtro: Cadena de búsqueda por apellido/nombre de empleados.
            dotacion_filtro: Dotación a filtrar ("Todas" = sin filtro).
        
        Returns:
            None (Modifica table_cambios).
        
        Side Effects:
            - Limpia table_cambios
            - Inserta filas filtradas desde sheet_cambio_turnos
            - Actualiza contador de resultados
        
        Example:
            >>> filtrar_datos_cambios("García", "A")
            # Muestra solo cambios donde García está en cualquier empleado y dotación es "A"
        """
        try:
            if dotacion_filtro == "Todas" and nombre_filtro == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                self.cargar_datos_completos_cambios()
                return

            nombre_filtro_norm = self.app.normalizar_texto(nombre_filtro)
            dotacion_filtro_norm = self.app.normalizar_texto(dotacion_filtro)

            self.app.table_cambios.delete(*self.app.table_cambios.get_children())
            total = 0
            if self.app.sheet_cambio_turnos:
                for fila in self.app.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    if len(fila_procesada) <= 11:
                        continue
                    nombre_1_norm = self.app.normalizar_texto(fila_procesada[3])
                    nombre_2_norm = self.app.normalizar_texto(fila_procesada[9])
                    dotacion_1_norm = self.app.normalizar_texto(fila_procesada[5])
                    dotacion_2_norm = self.app.normalizar_texto(fila_procesada[11])
                    coincide_nombre = nombre_filtro_norm in nombre_1_norm or nombre_filtro_norm in nombre_2_norm
                    coincide_dotacion = dotacion_filtro_norm in dotacion_1_norm or dotacion_filtro_norm in dotacion_2_norm
                    if dotacion_filtro == "Todas":
                        if coincide_nombre:
                            self.app.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1
                    elif nombre_filtro == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                        if coincide_dotacion:
                            self.app.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1
                    else:
                        if coincide_nombre and coincide_dotacion:
                            self.app.table_cambios.insert("", "end", values=fila_procesada)
                            total += 1

            self.ajustar_ancho_columnas(self.app.table_cambios)
            if hasattr(self.app, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.app.resultados_cambios_label, total)
        except Exception as e:
            print(f"Error al filtrar los datos en el Treeview cambios: {e}")

    def programar_filtrado_cambios(self):
        """Programa el filtrado de cambios con debounce de 250ms."""
        if self.app.filtro_after_cambios:
            self.app.root.after_cancel(self.app.filtro_after_cambios)
        self.app.filtro_after_cambios = self.app.root.after(
            250,
            lambda: self.filtrar_datos_cambios(
                self.app.apellido_filter_cambios_var.get(),
                self.app.dotacion_filter_cambios_var.get()
            )
        )

    def actualizar_tabla(self):
        """Actualiza las tablas si es necesario (llamado en cada refresh de Excel)."""
        if self.app.sheet_novedades and self.app.current_view == 'table' and hasattr(self.app, "tabla_novedades"):
            if not hasattr(self.app, "apellido_filter_novedades_var") or not hasattr(self.app, "dotacion_filter_novedades_var") or not hasattr(self.app, "tipo_filter_novedades_var"):
                return
            if self.app.apellido_filter_novedades_var.get() != self.app.PLACEHOLDER_BUSCAR_NOMBRE or self.app.dotacion_filter_novedades_var.get() != "Todas" or self.app.tipo_filter_novedades_var.get() != "Todos":
                return
            self.app.tabla_novedades.delete(*self.app.tabla_novedades.get_children())
            total = 0
            for row in self.app.sheet_novedades.iter_rows(min_row=2, values_only=True):
                row_data = ["-" if celda is None else celda for celda in row]
                self.app.tabla_novedades.insert("", "end", values=row_data)
                total += 1
            self.ajustar_ancho_columnas(self.app.tabla_novedades)
            if hasattr(self.app, "resultados_novedades_label"):
                self.actualizar_contador_resultados(self.app.resultados_novedades_label, total)
        
        if self.app.sheet_cambio_turnos and self.app.current_view == "table_cambios" and hasattr(self.app, "table_cambios"):
            if not hasattr(self.app, "apellido_filter_cambios_var") or not hasattr(self.app, "dotacion_filter_cambios_var"):
                return
            if self.app.apellido_filter_cambios_var.get() != self.app.PLACEHOLDER_BUSCAR_NOMBRE or self.app.dotacion_filter_cambios_var.get() != "Todas":
                return
            total = 0
            for item in self.app.table_cambios.get_children():
                self.app.table_cambios.delete(item)
            for row in self.app.sheet_cambio_turnos.iter_rows(min_row=2, values_only=True):
                row_data = ["-" if celda is None else celda for celda in row]
                self.app.table_cambios.insert("", "end", values=row_data)
                total += 1
            self.ajustar_ancho_columnas(self.app.table_cambios)
            if hasattr(self.app, "resultados_cambios_label"):
                self.actualizar_contador_resultados(self.app.resultados_cambios_label, total)
    
    def crear_tabla_novedades(self):
        """Crea la interfaz de tabla de novedades con filtros y botones."""
        def on_focus_in(event):
            if apellido_filter.get() == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                apellido_filter.delete(0, tk.END)
                apellido_filter.config(fg="black")

        def on_focus_out(event):
            if apellido_filter.get() == "":
                apellido_filter.insert(0, self.app.PLACEHOLDER_BUSCAR_NOMBRE)
                apellido_filter.config(fg="grey")
        
        def on_double_click(event):
            treeview = event.widget
            selected_item = treeview.selection()
            if selected_item:
                item_values = treeview.item(selected_item[0], "values")
                self.mostrar_modal_detalle(item_values, columnas, "novedad")
            else:
                print("No se seleccionó ningún elemento.")
            
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD",
            "DOTACION", "TURNOS", "FRANCO", "NOVEDAD", "Fecha de Inicio Novedad", "Fecha de Fin Novedad",
            "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones", "USUARIO WINDOWS"
        ]
        self.app.apellido_filter_novedades_var = tk.StringVar()
        self.app.dotacion_filter_novedades_var = tk.StringVar()
        self.app.tipo_filter_novedades_var = tk.StringVar()
        
        # Título y botones
        ttk.Label(self.app.table_frame, text="Registro de novedades", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, pady=10, padx=10, sticky="w"
        )
        if self.app.tiene_permiso("cambios_turno.ver"):
            ttk.Button(
                self.app.table_frame, text="Ver cambios de turno",
                command=lambda: self.app.toggle_view("table_cambios")
            ).grid(row=0, column=1, pady=10, padx=10, sticky="e")
        
        # Filtro nombre
        apellido_filter = tk.Entry(self.app.table_frame, textvariable=self.app.apellido_filter_novedades_var, font=("Helvetica", 10))
        apellido_filter.grid(row=0, column=2, sticky="e", pady=10, padx=10, ipady=5, ipadx=10)
        
        # Filtro dotación
        self.app.dotacion_filter_novedades = ttk.Combobox(
            self.app.table_frame, textvariable=self.app.dotacion_filter_novedades_var,
            values=self.app.DOTACIONES, width=10, state="normal"
        )
        dotacion_filter = self.app.dotacion_filter_novedades
        dotacion_filter.grid(row=0, column=3, sticky="e", pady=5)

        self.app.tipo_filter_novedades = ttk.Combobox(
            self.app.table_frame, textvariable=self.app.tipo_filter_novedades_var,
            values=["Todos", *self.app.tipo_novedades], width=18, state="readonly"
        )
        self.app.tipo_filter_novedades.grid(row=0, column=4, sticky="e", pady=5, padx=6)
        
        self.app.resultados_novedades_label = ttk.Label(self.app.table_frame, text="0 resultados", font=("Helvetica", 9))
        self.app.resultados_novedades_label.grid(row=0, column=5, sticky="w", padx=8)
        
        if self.app.tiene_permiso("novedades.crear"):
            ttk.Button(self.app.table_frame, text="Nueva novedad", command=lambda: self.app.toggle_view("form")).grid(
                row=0, column=6, pady=10, padx=2, sticky="e"
            )
        if self.app.tiene_permiso("cambios_turno.crear"):
            ttk.Button(self.app.table_frame, text="Nuevo cambio de turno", command=lambda: self.app.toggle_view("form_cambios")).grid(
                row=0, column=7, pady=10, padx=10, sticky="e"
            )

        apellido_filter.insert(0, self.app.PLACEHOLDER_BUSCAR_NOMBRE)
        dotacion_filter.insert(0, "Todas")
        self.app.tipo_filter_novedades.set("Todos")
        apellido_filter.config(fg="grey")
        apellido_filter.bind("<FocusIn>", on_focus_in)
        apellido_filter.bind("<FocusOut>", on_focus_out)
        
        # Detectar cambios en filtros
        self.app.apellido_filter_novedades_var.trace_add("write", lambda *args: self.programar_filtrado_novedades())
        self.app.dotacion_filter_novedades_var.trace_add("write", lambda *args: self.programar_filtrado_novedades())
        self.app.tipo_filter_novedades_var.trace_add("write", lambda *args: self.programar_filtrado_novedades())
        
        # Contenedor del Treeview
        self.app.tree_frame = ttk.Frame(self.app.table_frame, width=self.app.WIDTH, height=self.app.HEIGHT)
        self.app.tree_frame.grid(row=1, column=0, columnspan=8, sticky="nsew")
        self.app.tree_frame.grid_propagate(False)

        self.app.tree_frame.grid_rowconfigure(0, weight=1)
        self.app.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.app.tabla_novedades = ttk.Treeview(self.app.tree_frame, columns=columnas, show="headings", height=40)
        self.app.tabla_novedades.grid(row=0, column=0, sticky="nsew")
        self.app.tabla_novedades.bind("<Double-1>", on_double_click)
        
        # Encabezados y anchos
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 120, 90, 120, 120, 120, 140, 120]
        for col, ancho in zip(columnas, anchuras):
            self.app.tabla_novedades.heading(col, text=col.capitalize())
            self.app.tabla_novedades.column(col, width=ancho, stretch=False)

        # Scrollbars
        scrollbar_vertical = ttk.Scrollbar(self.app.tree_frame, orient="vertical", command=self.app.tabla_novedades.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal = ttk.Scrollbar(self.app.tree_frame, orient="horizontal", command=self.app.tabla_novedades.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.app.tabla_novedades.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)

        # Cargar datos
        self.cargar_datos_completos_novedades()
        self.app.tabla_novedades_creada = True

    def crear_tabla_cambios(self):
        """Crea la interfaz de tabla de cambios de turnos con filtros y botones."""
        def on_focus_in(event):
            if apellido_filter.get() == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                apellido_filter.delete(0, tk.END)
                apellido_filter.config(fg="black")

        def on_focus_out(event):
            if apellido_filter.get() == "":
                apellido_filter.insert(0, self.app.PLACEHOLDER_BUSCAR_NOMBRE)
                apellido_filter.config(fg="grey")
        
        def on_double_click(event):
            treeview = event.widget
            selected_item = treeview.selection()
            if selected_item:
                item_values = treeview.item(selected_item[0], "values")
                self.mostrar_modal_detalle(item_values, columnas, "cambio de turno")
            else:
                print("No se seleccionó ningún elemento.")
        
        self.app.apellido_filter_cambios_var = tk.StringVar()
        self.app.dotacion_filter_cambios_var = tk.StringVar()
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", 
            "TURNOS", "FRANCO", "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2", 
            "TURNOS2", "FRANCO2", "Fecha de Cambio de Turno", "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones", "USUARIO WINDOWS"
        ]
        
        # Detectar cambios en filtros
        self.app.apellido_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        self.app.dotacion_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        
        # Título y botones
        ttk.Label(self.app.table_cambios_frame, text="Registro de cambios de turnos", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, pady=10, padx=10, sticky="w"
        )
        if self.app.tiene_permiso("novedades.ver"):
            ttk.Button(
                self.app.table_cambios_frame, text="Ver novedades",
                command=lambda: self.app.toggle_view("table")
            ).grid(row=0, column=1, pady=10, padx=1, sticky="e")
        
        # Filtro nombre
        apellido_filter = tk.Entry(self.app.table_cambios_frame, textvariable=self.app.apellido_filter_cambios_var, font=("Helvetica", 10))
        apellido_filter.grid(row=0, column=2, sticky="e", pady=10, padx=10, ipady=5, ipadx=10)
        
        # Filtro dotación
        self.app.dotacion_filter_cambios = ttk.Combobox(
            self.app.table_cambios_frame, textvariable=self.app.dotacion_filter_cambios_var,
            values=self.app.DOTACIONES, width=10, state="normal"
        )
        dotacion_filter = self.app.dotacion_filter_cambios
        dotacion_filter.grid(row=0, column=3, sticky="e", pady=5)
        
        self.app.resultados_cambios_label = ttk.Label(self.app.table_cambios_frame, text="0 resultados", font=("Helvetica", 9))
        self.app.resultados_cambios_label.grid(row=0, column=4, sticky="w", padx=8)
        
        if self.app.tiene_permiso("novedades.crear"):
            ttk.Button(self.app.table_cambios_frame, text="Nueva novedad", command=lambda: self.app.toggle_view("form")).grid(
                row=0, column=5, pady=10, padx=1, sticky="e"
            )
        if self.app.tiene_permiso("cambios_turno.crear"):
            ttk.Button(self.app.table_cambios_frame, text="Nuevo cambio de turno", command=lambda: self.app.toggle_view("form_cambios")).grid(
                row=0, column=6, pady=10, padx=1, sticky="e"
            )

        dotacion_filter.insert(0, "Todas")
        apellido_filter.insert(0, self.app.PLACEHOLDER_BUSCAR_NOMBRE)
        apellido_filter.config(fg="grey")
        apellido_filter.bind("<FocusIn>", on_focus_in)
        apellido_filter.bind("<FocusOut>", on_focus_out)
        
        # Contenedor del Treeview
        self.app.tree_frame = ttk.Frame(self.app.table_cambios_frame, width=self.app.WIDTH, height=self.app.HEIGHT)
        self.app.tree_frame.grid(row=1, column=0, columnspan=7, sticky="nsew")
        self.app.tree_frame.grid_propagate(False)

        self.app.tree_frame.grid_rowconfigure(0, weight=1)
        self.app.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.app.table_cambios = ttk.Treeview(self.app.tree_frame, columns=columnas, show="headings", height=38)
        self.app.table_cambios.grid(row=0, column=0, sticky="nsew")
        self.app.table_cambios.bind("<Double-1>", on_double_click)
        
        # Encabezados y anchos
        anchuras = [30, 100, 60, 150, 150, 80, 60, 80, 60, 150, 150, 80, 60, 80, 120, 120, 120, 120, 140]
        for col, ancho in zip(columnas, anchuras):
            self.app.table_cambios.heading(col, text=col)
            self.app.table_cambios.column(col, width=ancho, anchor='center', stretch=False)

        # Scrollbars
        scrollbar_vertical = ttk.Scrollbar(self.app.tree_frame, orient="vertical", command=self.app.table_cambios.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal = ttk.Scrollbar(self.app.tree_frame, orient="horizontal", command=self.app.table_cambios.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.app.table_cambios.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
        
        # Cargar datos
        self.cargar_datos_completos_cambios()
        self.app.tabla_cambios_creada = True
