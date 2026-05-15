"""Lógica de formularios de novedades y cambios de turnos."""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from ttkbootstrap import DateEntry
from datetime import datetime, timedelta
import validators


class FormsManager:
    """Gestor de formularios, guardado de datos y búsqueda de legajos.
    
    Maneja la presentación de formularios de novedades y cambios de turnos,
    búsqueda de personal por legajo (con modal filtrable), validación,
    guardado en Excel y manejo de errores.
    
    Attributes:
        app: Referencia a la aplicación FormularioExcelApp que contiene
             ui, excel_store, base_rows, base_index, etc.
    
    Methods:
        mostrar_modal(boton): Muestra modal para seleccionar legajo.
        buscar_legajo(campo): Abre modal de búsqueda en campo específico.
        mostrar_formulario_novedades(): Crea y muestra formulario de novedades.
        mostrar_formulario_cambios(): Crea y muestra formulario de cambios de turnos.
        guardar_datos_novedades(): Valida y guarda novedad en Excel.
        guardar_datos_cambios(): Valida y guarda cambio de turno en Excel.
        limpiar_formulario_novedades(): Limpia campos del formulario de novedades.
        limpiar_formulario_cambios(): Limpia campos del formulario de cambios.
        mostrar_error_novedades(mensaje): Muestra messagebox de error.
        mostrar_error_cambios(mensaje): Muestra messagebox de error.
    """
    
    def __init__(self, app):
        """Inicializa el gestor de formularios.
        
        Args:
            app: Instancia de FormularioExcelApp con acceso a root, base_rows, etc.
        """
        self.app = app
    
    def mostrar_modal(self, boton=1):
        """Muestra un modal para seleccionar personal por legajo.
        
        Presenta un árbol filtrable con todos los empleados de la base (legajo,
        apellidos, especialidad, dotación, turnos, franco). Permite búsqueda por
        apellido/nombre/legajo con filtro debounced (250ms). Al hacer doble clic,
        rellena los campos del formulario activo (novedades o cambios).
        
        Args:
            boton: Identificador del campo de destino (1=Legajo 1, 2=Legajo 2).
        
        Returns:
            None (Modifica campos de la aplicación al seleccionar).
        
        Raises:
            (No levanta excepciones; errores de datos manejados silenciosamente).
        
        Example:
            >>> self.mostrar_modal(boton=1)
            # Muestra modal y llena Legajo 1 al seleccionar
        """
        modal = tk.Toplevel(self.app.root)
        modal.title("Seleccionar Legajo")
        modal.geometry("1250x500")
        modal.transient(self.app.root)
        modal.grab_set()
        modal.focus_set()

        registros_busqueda = []
        for row in self.app.base_rows:
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
                "nombre_norm": self.app.normalizar_texto(apellidos_nombres),
                "legajo_norm": self.app.normalizar_texto(str(legajo)),
            })
        
        tk.Label(modal, text="Filtrar por apellido o nombre:").grid(row=0, column=0, padx=10, pady=5)
        apellido_filter_var = tk.StringVar()
        apellido_filter = tk.Entry(modal, textvariable=apellido_filter_var)
        apellido_filter.grid(row=0, column=1, padx=10, pady=5, ipady=5, ipadx=50)
        modal_resultados_label = ttk.Label(modal, text="0 resultados")
        modal_resultados_label.grid(row=0, column=2, padx=10, pady=5, sticky="w")
        
        tree = ttk.Treeview(modal, columns=("legajo", "apellido_nombre", "especialidad", "dotacion", "turnos", "franco"), 
                           show="headings", height=25)
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

            filtro = self.app.normalizar_texto(apellido_filter_var.get())
            total = 0
            for registro in registros_busqueda:
                if not filtro or filtro in registro["nombre_norm"] or filtro in registro["legajo_norm"]:
                    tree.insert("", "end", values=registro["values"])
                    total += 1
            self.app.tables_manager.actualizar_contador_resultados(modal_resultados_label, total)

        def cargar_tabla():
            nonlocal search_after_id
            if search_after_id:
                modal.after_cancel(search_after_id)
            search_after_id = modal.after(250, aplicar_filtro)

        aplicar_filtro()
        apellido_filter_var.trace_add("write", lambda *args: cargar_tabla())

        def on_double_click(event):
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
                self.app.legajo_var.set(legajo_sap)
                self.app.apellidos_nombres_var.set(apellido_nombre)
                self.app.especialidad_var.set(especialidad)
                self.app.dotacion_var.set(dotacion)
                self.app.turnos_var.set(turnos)
                self.app.franco_var.set(franco)
            elif boton == 2:
                self.app.legajo_2_var.set(legajo_sap)
                self.app.apellidos_nombres_2_var.set(apellido_nombre)
                self.app.especialidad_2_var.set(especialidad)
                self.app.dotacion_2_var.set(dotacion)
                self.app.turnos_2_var.set(turnos)
                self.app.franco_2_var.set(franco)
                
            modal.destroy()

        tree.bind("<Double-1>", on_double_click)
        tk.Button(modal, text="Cerrar", command=modal.destroy).grid(row=2, column=0, columnspan=2, pady=10)

    def buscar_legajo(self, campo=1):
        """Busca un legajo en BASE y autocompleta los campos del formulario.
        
        Busca el legajo en el índice en memoria (base_index), y si existe,
        rellena automáticamente apellido, especialidad, dotación, turnos y franco.
        
        Args:
            campo: Campo a completar (1=Legajo 1, 2=Legajo 2 en cambios de turno).
        
        Returns:
            None (Modifica StringVars de la aplicación).
        
        Raises:
            ValueError: Si el legajo no es un número válido (mostrado en messagebox).
        """
        try:
            if campo == 1:
                legajo = int(self.app.legajo_var.get().strip())
            elif campo == 2:
                legajo = int(self.app.legajo_2_var.get().strip())
                
            print(f"Buscando legajo SAP: {legajo}")

            row = self.app.base_index.get(legajo)
            if row:
                if campo == 1:
                    self.app.apellidos_nombres_var.set(row[1])
                    self.app.especialidad_var.set(row[2])
                    self.app.dotacion_var.set(row[3])
                    self.app.turnos_var.set(row[4])
                    self.app.franco_var.set(row[5])
                elif campo == 2:
                    self.app.apellidos_nombres_2_var.set(row[1])
                    self.app.especialidad_2_var.set(row[2])
                    self.app.dotacion_2_var.set(row[3])
                    self.app.turnos_2_var.set(row[4])
                    self.app.franco_2_var.set(row[5])
                print(f"Legajo encontrado: {legajo}")
            else:
                messagebox.showinfo("Legajo No Encontrado", f"El legajo {legajo} no fue encontrado.")
                print(f"Legajo SAP {legajo} no encontrado.")

        except ValueError:
            messagebox.showerror("Error de entrada", "Por favor, ingrese un número de legajo válido.")

    def limpiar_formulario_novedades(self):
        """Limpia todos los campos del formulario de novedades."""
        self.app.legajo_var.set('')
        self.app.apellidos_nombres_var.set('')
        self.app.especialidad_var.set('')
        self.app.dotacion_var.set('')
        self.app.turnos_var.set('')
        self.app.franco_var.set('')
        self.app.novedad_var.set('')        
        self.app.fecha_inicio_novedad_var.set('')
        self.app.fecha_fin_novedad_var.set('')
        self.app.referencia_estacion_var.set('')
        self.app.supervisor_var.set('')
        self.app.observaciones_var.set('')
        if hasattr(self.app, "observaciones_novedades_text"):
            self.app.observaciones_novedades_text.delete("1.0", "end")
        if hasattr(self.app, "fecha_fin_novedad_entry"):
            self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        if self.app.error_novedades_label is not None:
            self.app.error_novedades_label.config(text="")

    def limpiar_fecha_fin_novedad(self):
        """Limpia el campo de fecha de fin de novedad."""
        if hasattr(self.app, "fecha_fin_novedad_entry"):
            self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
            self.app.fecha_fin_novedad_var.set("")

    def limpiar_formulario_cambios(self):
        """Limpia todos los campos del formulario de cambios de turnos."""
        self.app.legajo_var.set('')
        self.app.apellidos_nombres_var.set('')
        self.app.especialidad_var.set('')
        self.app.dotacion_var.set('')
        self.app.turnos_var.set('')
        self.app.franco_var.set('')
        self.app.legajo_2_var.set('')
        self.app.apellidos_nombres_2_var.set('')
        self.app.especialidad_2_var.set('')
        self.app.dotacion_2_var.set('')
        self.app.turnos_2_var.set('')
        self.app.franco_2_var.set('')
        self.app.fecha_cambio_turno_var.set('')
        self.app.referencia_estacion_var.set('')
        self.app.supervisor_var.set('')
        self.app.observaciones_var.set('') 
        if hasattr(self.app, "observaciones_cambios_text"):
            self.app.observaciones_cambios_text.delete("1.0", "end")
        if self.app.error_cambios_label is not None:
            self.app.error_cambios_label.config(text="")

    def guardar_datos_novedades(self):
        """Valida y guarda una novedad en el Excel.
        
        Ejecuta validación mediante validators.validar_campos_requeridos_novedades(),
        genera un ID único, captura timestamp y usuario Windows, inserta fila en
        NOVEDADES (posición 2, latest-first), y guarda workbook. Si hay error
        de permisos (archivo abierto), muestra messagebox de error.
        
        Returns:
            None (Modifica Excel, limpia formulario, muestra messagebox).
        
        Raises:
            PermissionError: (Capturado internamente si Excel está abierto).
        
        Side Effects:
            - Inserta fila en sheet_novedades en posición 2
            - Guarda workbook en disco
            - Limpia formulario y recarga vistas de datos
            - Muestra messagebox de éxito o error
        """
        self.app.fecha_inicio_novedad_var.set(self.app.fecha_inicio_novedad_entry.entry.get())
        self.app.fecha_fin_novedad_var.set(self.app.fecha_fin_novedad_entry.entry.get())
        self.app.observaciones_var.set(self.app.observaciones_novedades_text.get("1.0", "end-1c"))
        
        es_valido, mensaje_error = validators.validar_campos_requeridos_novedades(
            self.app.legajo_var, self.app.apellidos_nombres_var, self.app.novedad_var,
            self.app.referencia_estacion_var, self.app.supervisor_var,
            self.app.fecha_inicio_novedad_entry, self.app.fecha_fin_novedad_entry,
            self.app.tipo_novedades
        )
        
        if es_valido:
            try:
                new_id = self.app.obtener_nuevo_id_con_sincronizacion(self.app.SHEET_NOVEDADES)
                current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
                usuario_windows = self.app.obtener_usuario_windows()
                col_usuario = self.app.asegurar_columna_usuario(self.app.sheet_novedades)

                self.app.sheet_novedades.insert_rows(2)
                self.app.sheet_novedades['A2'] = new_id
                self.app.sheet_novedades['B2'] = current_datetime
                self.app.sheet_novedades['C2'] = self.app.legajo_var.get()
                self.app.sheet_novedades['D2'] = self.app.apellidos_nombres_var.get()
                self.app.sheet_novedades['E2'] = self.app.especialidad_var.get()
                self.app.sheet_novedades['F2'] = self.app.dotacion_var.get()
                self.app.sheet_novedades['G2'] = self.app.turnos_var.get()
                self.app.sheet_novedades['H2'] = self.app.franco_var.get()
                self.app.sheet_novedades['I2'] = self.app.novedad_var.get()
                self.app.sheet_novedades['J2'] = self.app.fecha_inicio_novedad_var.get()
                self.app.sheet_novedades['K2'] = self.app.fecha_fin_novedad_var.get()
                self.app.sheet_novedades['L2'] = self.app.referencia_estacion_var.get()
                self.app.sheet_novedades['M2'] = self.app.supervisor_var.get()
                self.app.sheet_novedades['N2'] = self.app.observaciones_var.get()
                self.app.sheet_novedades['O2'] = usuario_windows
                # self.app.sheet_novedades.cell(row=2, column=col_usuario, value=usuario_windows)
                
                self.app.wb.save(self.app.excel_file)
                self.app.excel_last_mtime = self.app.obtener_mtime_excel()

                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                self.limpiar_formulario_novedades()
                self.app.toggle_view()
                print("Datos guardados correctamente.")
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar el archivo porque está abierto en otro programa. Por favor, cierre el archivo y vuelva a intentarlo.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)} Por favor, intente de nuevo y si el problema persiste avise al administrador")
                print(f"Error al guardar los datos: {e}")
        else:
            print("Algunos campos son obligatorios y están vacíos.")
            self.mostrar_error_novedades(mensaje_error)

    def guardar_datos_cambios(self):
        """Valida y guarda un cambio de turno en el Excel.
        
        Ejecuta validación mediante validators.validar_campos_requeridos_cambios(),
        genera un ID único, captura timestamp y usuario Windows, inserta fila en
        Cambio de Turnos (posición 2, latest-first), y guarda workbook. Si hay error
        de permisos (archivo abierto), muestra messagebox de error.
        
        Returns:
            None (Modifica Excel, limpia formulario, muestra messagebox).
        
        Raises:
            PermissionError: (Capturado internamente si Excel está abierto).
        
        Side Effects:
            - Inserta fila en sheet_cambio_turnos en posición 2
            - Guarda workbook en disco
            - Limpia formulario y recarga vistas de datos
            - Muestra messagebox de éxito o error
        """
        self.app.fecha_cambio_turno_var.set(self.app.fecha_cambio_turno_entry.entry.get())
        self.app.observaciones_var.set(self.app.observaciones_cambios_text.get("1.0", "end-1c"))
        
        es_valido, mensaje_error = validators.validar_campos_requeridos_cambios(
            self.app.legajo_var, self.app.apellidos_nombres_var, self.app.legajo_2_var,
            self.app.apellidos_nombres_2_var, self.app.fecha_cambio_turno_entry,
            self.app.referencia_estacion_var, self.app.supervisor_var
        )
        
        if es_valido:
            try:
                new_id = self.app.obtener_nuevo_id_con_sincronizacion(self.app.SHEET_CAMBIO_TURNOS)
                current_datetime = datetime.now().strftime("%d/%m/%Y %H:%M")
                usuario_windows = self.app.obtener_usuario_windows()
                col_usuario = self.app.asegurar_columna_usuario(self.app.sheet_cambio_turnos)

                self.app.sheet_cambio_turnos.insert_rows(2)
                self.app.sheet_cambio_turnos['A2'] = new_id
                self.app.sheet_cambio_turnos['B2'] = current_datetime
                self.app.sheet_cambio_turnos['C2'] = self.app.legajo_var.get()
                self.app.sheet_cambio_turnos['D2'] = self.app.apellidos_nombres_var.get()
                self.app.sheet_cambio_turnos['E2'] = self.app.especialidad_var.get()
                self.app.sheet_cambio_turnos['F2'] = self.app.dotacion_var.get()
                self.app.sheet_cambio_turnos['G2'] = self.app.turnos_var.get()
                self.app.sheet_cambio_turnos['H2'] = self.app.franco_var.get()
                self.app.sheet_cambio_turnos['I2'] = self.app.legajo_2_var.get()
                self.app.sheet_cambio_turnos['J2'] = self.app.apellidos_nombres_2_var.get()
                self.app.sheet_cambio_turnos['K2'] = self.app.especialidad_2_var.get()
                self.app.sheet_cambio_turnos['L2'] = self.app.dotacion_2_var.get()
                self.app.sheet_cambio_turnos['M2'] = self.app.turnos_2_var.get()
                self.app.sheet_cambio_turnos['N2'] = self.app.franco_2_var.get()
                self.app.sheet_cambio_turnos['O2'] = self.app.fecha_cambio_turno_var.get()
                self.app.sheet_cambio_turnos['P2'] = self.app.referencia_estacion_var.get()
                self.app.sheet_cambio_turnos['Q2'] = self.app.supervisor_var.get()
                self.app.sheet_cambio_turnos['R2'] = self.app.observaciones_var.get()
                self.app.sheet_cambio_turnos['S2'] = usuario_windows

                self.app.wb.save(self.app.excel_file)
                self.app.excel_last_mtime = self.app.obtener_mtime_excel()

                messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")

                self.limpiar_formulario_cambios()
                self.app.toggle_view("table_cambios")
                print("Datos guardados correctamente.")
            except PermissionError:
                messagebox.showerror("Error", "No se pudo guardar el archivo porque está abierto en otro programa. Por favor, cierre el archivo y vuelva a intentarlo.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudieron guardar los datos: {str(e)} Por favor, intente de nuevo y si el problema persiste avise al administrador")
                print(f"Error al guardar los datos: {e}")
        else:
            print("Algunos campos son obligatorios y están vacíos.")
            self.mostrar_error_cambios(mensaje_error)

    def mostrar_error_novedades(self, mensaje):
        """Muestra un mensaje de error en el formulario de novedades."""
        if self.app.error_novedades_label is None:
            self.app.error_novedades_label = ttk.Label(self.app.form_frame, foreground="red")
            self.app.error_novedades_label.grid(row=11, column=0, columnspan=7, pady=5, sticky="w")
        self.app.error_novedades_label.config(text=mensaje)
    
    def mostrar_error_cambios(self, mensaje):
        """Muestra un mensaje de error en el formulario de cambios."""
        if self.app.error_cambios_label is None:
            self.app.error_cambios_label = ttk.Label(self.app.form_cambios_frame, foreground="red")
            self.app.error_cambios_label.grid(row=15, column=0, columnspan=7, pady=5, sticky="w")
        self.app.error_cambios_label.config(text=mensaje)

    def mostrar_formulario_novedades(self):
        """Crea la interfaz del formulario de novedades."""
        ttk.Label(self.app.form_frame, text="Formulario de novedades", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, columnspan=3, pady=10, padx=10, sticky="nw"
        )
        
        # Legajo
        ttk.Label(self.app.form_frame, text="   Legajo").grid(row=1, column=0, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.app.legajo_entry = ttk.Entry(self.app.form_frame, textvariable=self.app.legajo_var, width=10)
        self.app.legajo_entry.grid(row=2, column=0, sticky="ew", pady=5)
        self.app.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        ttk.Button(self.app.form_frame, text="Buscar Personal", command=self.mostrar_modal).grid(row=2, column=1, pady=10, padx=10)

        # Apellidos y Nombres
        ttk.Label(self.app.form_frame, text="  Apellido y Nombre").grid(row=3, column=0, sticky="w", padx=5)
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=3, column=0, sticky="w")
        self.app.apellidos_nombres_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.apellidos_nombres_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.apellidos_nombres_entry.grid(row=4, column=0, sticky="w", pady=5)

        # Especialidad
        ttk.Label(self.app.form_frame, text="Especialidad").grid(row=3, column=1, sticky="w", padx=5)
        self.app.especialidad_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.especialidad_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.especialidad_entry.grid(row=4, column=1, sticky="w", pady=5, padx=8)

        # Dotación
        ttk.Label(self.app.form_frame, text="Dotación").grid(row=3, column=2, sticky="w", padx=5)
        self.app.dotacion_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.dotacion_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.dotacion_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        # Turnos
        ttk.Label(self.app.form_frame, text="Turno").grid(row=3, column=3, sticky="w", padx=5)
        self.app.turnos_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.turnos_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.turnos_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        # Franco
        ttk.Label(self.app.form_frame, text="Franco").grid(row=3, column=4, sticky="w", padx=5)
        self.app.franco_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.franco_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.franco_entry.grid(row=4, column=4, sticky="w", pady=5, padx=8)

        # Tipo Novedad
        ttk.Label(self.app.form_frame, text="   Tipo Novedad").grid(row=5, column=0, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=0, sticky="w")
        self.app.novedad_entry = ttk.Combobox(
            self.app.form_frame, textvariable=self.app.novedad_var,
            values=self.app.tipo_novedades, width=38, state="readonly"
        )
        self.app.novedad_entry.grid(row=6, column=0, columnspan=2, sticky="w", pady=5)

        # Fecha Inicio
        ttk.Label(self.app.form_frame, text="   Fecha de Inicio Novedad").grid(row=5, column=3, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=3, sticky="w")
        self.app.fecha_inicio_novedad_entry = DateEntry(
            self.app.form_frame, dateformat='%d/%m/%Y', bootstyle="danger"
        )
        self.app.fecha_inicio_novedad_entry.grid(row=6, column=3, sticky="w", pady=5)
        self.app.fecha_inicio_novedad_entry.entry.bind("<Key>", lambda event: "break")

        # Fecha Fin
        ttk.Label(self.app.form_frame, text="Fecha de Fin Novedad").grid(row=5, column=4, sticky="w")
        self.app.fecha_fin_novedad_entry = DateEntry(
            self.app.form_frame, dateformat='%d/%m/%Y', bootstyle="success"
        )
        self.app.fecha_fin_novedad_entry.grid(row=6, column=4, sticky="w", pady=5)
        self.app.fecha_fin_novedad_entry.entry.bind("<Key>", lambda event: "break")
        self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        self.app.fecha_fin_novedad_var.set("")
        ttk.Button(self.app.form_frame, text="Limpiar", command=self.limpiar_fecha_fin_novedad).grid(row=6, column=5, sticky="w", padx=5)
        
        # Referencia Estación
        ttk.Label(self.app.form_frame, text="   REFERENCIA ESTACIÓN").grid(row=7, column=0, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=0, sticky="w")
        self.app.referencia_estacion_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.referencia_estacion_var, width=40
        )
        self.app.referencia_estacion_entry.grid(row=8, column=0, columnspan=2, sticky="w", pady=5)

        # Supervisor
        ttk.Label(self.app.form_frame, text="   SUPERVISOR").grid(row=7, column=3, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=3, sticky="w")
        self.app.supervisor_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.supervisor_var, width=40
        )
        self.app.supervisor_entry.grid(row=8, column=3, columnspan=2, sticky="w", pady=5)

        # Observaciones
        ttk.Label(self.app.form_frame, text="Observaciones").grid(row=9, column=0, sticky="w")
        self.app.observaciones_novedades_text = scrolledtext.ScrolledText(
            self.app.form_frame, wrap=tk.WORD, height=12, width=180
        )
        self.app.observaciones_novedades_text.grid(row=10, column=0, columnspan=7, pady=5, sticky="w")

        # Botones
        ttk.Button(self.app.form_frame, text="Guardar Novedad", command=self.guardar_datos_novedades).grid(
            row=11, column=3, columnspan=2, pady=10
        )
        ttk.Button(
            self.app.form_frame, text="Cerrar",
            command=lambda: (self.limpiar_formulario_novedades(), self.app.toggle_view())
        ).grid(row=11, column=4, columnspan=2, pady=10)

    def mostrar_formulario_cambios(self):
        """Crea la interfaz del formulario de cambios de turnos."""
        ttk.Label(self.app.form_cambios_frame, text="Formulario de cambios de turnos", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, columnspan=3, pady=10, padx=10, sticky="nw"
        )

        # Legajo 1
        ttk.Label(self.app.form_cambios_frame, text="   Legajo").grid(row=1, column=0, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.app.legajo_entry = ttk.Entry(self.app.form_cambios_frame, textvariable=self.app.legajo_var, width=10)
        self.app.legajo_entry.grid(row=2, column=0, sticky="w", pady=5)
        self.app.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        ttk.Button(self.app.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(1)).grid(
            row=2, column=1, pady=10, padx=10
        )

        # Persona 1 - Datos
        ttk.Label(self.app.form_cambios_frame, text="  Apellido y Nombre").grid(row=3, column=0, sticky="w", padx=5)
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=3, column=0, sticky="w")
        self.app.apellidos_nombres_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.apellidos_nombres_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.apellidos_nombres_entry.grid(row=4, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.app.form_cambios_frame, text="Especialidad").grid(row=3, column=2, sticky="w", padx=5)
        self.app.especialidad_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.especialidad_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.especialidad_entry.grid(row=4, column=2, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Dotación").grid(row=3, column=3, sticky="w", padx=5)
        self.app.dotacion_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.dotacion_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.dotacion_entry.grid(row=4, column=3, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Turno").grid(row=3, column=4, sticky="w", padx=5)
        self.app.turnos_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.turnos_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.turnos_entry.grid(row=4, column=4, columnspan=2, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Franco").grid(row=3, column=6, sticky="w", padx=5)
        self.app.franco_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.franco_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.franco_entry.grid(row=4, column=6, sticky="w", pady=5, padx=8)

        # Legajo 2
        ttk.Label(self.app.form_cambios_frame, text="   Legajo 2").grid(row=5, column=0, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=0, sticky="w")
        self.app.legajo_2_entry = ttk.Entry(self.app.form_cambios_frame, textvariable=self.app.legajo_2_var, width=10)
        self.app.legajo_2_entry.grid(row=6, column=0, sticky="w", pady=5)
        self.app.legajo_2_entry.bind("<Return>", lambda event: self.buscar_legajo(2))

        ttk.Button(self.app.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(2)).grid(
            row=6, column=1, pady=10, padx=10
        )

        # Persona 2 - Datos
        ttk.Label(self.app.form_cambios_frame, text="  Apellido y Nombre 2").grid(row=7, column=0, sticky="w", padx=5)
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=0, sticky="w")
        self.app.apellidos_nombres_entry_2 = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.apellidos_nombres_2_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.apellidos_nombres_entry_2.grid(row=8, column=0, columnspan=2, sticky="w", pady=5)

        ttk.Label(self.app.form_cambios_frame, text="Especialidad 2").grid(row=7, column=2, sticky="w", padx=5)
        self.app.especialidad_entry_2 = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.especialidad_2_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.especialidad_entry_2.grid(row=8, column=2, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Dotación 2").grid(row=7, column=3, sticky="w", padx=5)
        self.app.dotacion_entry_2 = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.dotacion_2_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.dotacion_entry_2.grid(row=8, column=3, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Turno 2").grid(row=7, column=4, sticky="w", padx=5)
        self.app.turnos_entry_2 = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.turnos_2_var,
            state='readonly', style='Readonly.TEntry', width=40
        )
        self.app.turnos_entry_2.grid(row=8, column=4, columnspan=2, sticky="w", pady=5, padx=8)

        ttk.Label(self.app.form_cambios_frame, text="Franco 2").grid(row=7, column=6, sticky="w", padx=5)
        self.app.franco_entry_2 = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.franco_2_var,
            state='readonly', style='Readonly.TEntry', width=20
        )
        self.app.franco_entry_2.grid(row=8, column=6, sticky="w", pady=5, padx=8)
        
        # Fecha cambio turno
        ttk.Label(self.app.form_cambios_frame, text="   Fecha de cambio de turno").grid(row=9, column=0, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=9, column=0, sticky="w")
        self.app.fecha_cambio_turno_entry = DateEntry(
            self.app.form_cambios_frame, dateformat='%d/%m/%Y',
            startdate=datetime.today() + timedelta(days=1)
        )
        self.app.fecha_cambio_turno_entry.grid(row=10, column=0, columnspan=2, sticky="w", pady=5)
        self.app.fecha_cambio_turno_entry.entry.bind("<Key>", lambda event: "break")
        
        # Referencia Estación
        ttk.Label(self.app.form_cambios_frame, text="   REFERENCIA ESTACIÓN").grid(row=11, column=0, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=11, column=0, sticky="w")
        self.app.referencia_estacion_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.referencia_estacion_var, width=40
        )
        self.app.referencia_estacion_entry.grid(row=12, column=0, columnspan=2, sticky="w", pady=5)

        # Supervisor
        ttk.Label(self.app.form_cambios_frame, text="   SUPERVISOR").grid(row=11, column=3, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=11, column=3, sticky="w")
        self.app.supervisor_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.supervisor_var, width=40
        )
        self.app.supervisor_entry.grid(row=12, column=3, columnspan=2, sticky="w", pady=5)

        # Observaciones
        ttk.Label(self.app.form_cambios_frame, text="Observaciones").grid(row=13, column=0, sticky="w")
        self.app.observaciones_cambios_text = scrolledtext.ScrolledText(
            self.app.form_cambios_frame, wrap=tk.WORD, height=1, width=170
        )
        self.app.observaciones_cambios_text.grid(row=14, column=0, columnspan=7, pady=5, sticky="w")

        # Botones
        ttk.Button(self.app.form_cambios_frame, text="Guardar cambio de turno", command=self.guardar_datos_cambios).grid(
            row=15, column=2, columnspan=2, pady=10
        )
        ttk.Button(
            self.app.form_cambios_frame, text="Cerrar",
            command=lambda: (self.limpiar_formulario_cambios(), self.app.toggle_view("table_cambios"))
        ).grid(row=15, column=3, columnspan=2, pady=10)
