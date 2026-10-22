"""Lógica de tablas, filtrado y vistas de datos."""

from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from ttkbootstrap import DateEntry
from validators import (
    es_turno_ciclo,
    validar_datos_edicion_novedad,
    validar_datos_edicion_cambio,
)


ANCHOS_NOVEDADES = {
    "ID": 55, "Fecha de registro": 110, "LEGAJO": 70, "APELLIDOS Y NOMBRES": 230,
    "ESPECIALIDAD": 125, "DOTACION": 75, "TURNOS": 65, "FRANCO": 75, "NOVEDAD": 145,
    "Fecha de Inicio Novedad": 110, "Fecha de Fin Novedad": 110,
    "CUBIERTO": 170, "REFERENCIA ESTACION": 130, "SUPERVISOR": 150,
}

ANCHOS_CAMBIOS = {
    "ID": 55, "Fecha de registro": 110, "LEGAJO": 70, "APELLIDOS Y NOMBRES": 240,
    "ESPECIALIDAD": 130, "DOTACION": 75, "TURNOS": 60, "FRANCO": 75,
    "LEGAJO2": 70, "APELLIDOS Y NOMBRES2": 240, "ESPECIALIDAD2": 130, "DOTACION2": 75,
    "TURNOS2": 60, "FRANCO2": 75, "Fecha de Cambio de Turno": 125,
    "REFERENCIA ESTACION": 130, "SUPERVISOR": 160,
}

# Índices de corte: se quitan Observaciones y USUARIO WINDOWS del listado.
COLUMNAS_NOVEDADES_VISIBLES = 14
COLUMNAS_CAMBIOS_VISIBLES = 17

COLUMNAS_DETALLE_NOVEDAD = [
    "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD",
    "DOTACION", "TURNOS", "FRANCO", "NOVEDAD", "Fecha de Inicio Novedad", "Fecha de Fin Novedad",
    "CUBIERTO", "REEMPLAZOS", "REFERENCIA ESTACION", "SUPERVISOR", "Observaciones", "USUARIO WINDOWS"
]

COLUMNAS_DETALLE_CAMBIO = [
    "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION",
    "TURNOS", "FRANCO", "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2",
    "TURNOS2", "FRANCO2", "Fecha de Cambio de Turno", "REFERENCIA ESTACION", "SUPERVISOR",
    "Observaciones", "USUARIO WINDOWS"
]


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

    def _valores_novedad(self, fila):
        estado_cob = "No"
        detalle_reemp = "-"
        if hasattr(self.app, "records_service") and self.app.records_service is not None:
            try:
                estado_cob, detalle = self.app.records_service.calcular_estado_cobertura(
                    fila["id"], fila["fecha_inicio"], fila["fecha_fin"], fila["turnos"]
                )
                detalle_reemp = detalle or "-"
            except Exception:
                pass
        return (
            fila["id"], fila["registrado_en"], fila["legajo"], fila["apellidos_nombres"],
            fila["especialidad"], fila["dotacion"], fila["turnos"], fila["franco"],
            fila["novedad"], fila["fecha_inicio"], fila["fecha_fin"],
            estado_cob, detalle_reemp,
            fila["referencia_estacion"], fila["supervisor"], fila["observaciones"],
            fila["usuario_windows"],
        )

    def _construir_fila_tabla_novedad(self, fila_procesada, mapa_cobertura):
        """Inserta la columna CUBIERTO en la posición 11 (luego de Fecha de Fin Novedad)."""
        try:
            nid = int(fila_procesada[0])
        except (TypeError, ValueError):
            nid = None
        info = mapa_cobertura.get(nid) if (mapa_cobertura and nid is not None) else None
        if info:
            cubierto_str = info["resumen"]
        else:
            turno_val = fila_procesada[6] if len(fila_procesada) > 6 else ""
            cubierto_str = "N/A (Ciclo)" if es_turno_ciclo(turno_val) else "No"
        return fila_procesada[:11] + [cubierto_str] + fila_procesada[11:13]

    def _valores_cambio(self, fila):
        return (
            fila["id"], fila["registrado_en"], fila["legajo_1"], fila["apellidos_nombres_1"],
            fila["especialidad_1"], fila["dotacion_1"], fila["turnos_1"], fila["franco_1"],
            fila["legajo_2"], fila["apellidos_nombres_2"], fila["especialidad_2"], fila["dotacion_2"],
            fila["turnos_2"], fila["franco_2"], fila["fecha_cambio"],
            fila["referencia_estacion"], fila["supervisor"], fila["observaciones"],
            fila["usuario_windows"],
        )
    
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
            modal.geometry("680x700")
        else:
            modal.geometry("700x780")
        try:
            modal.state("zoomed")
        except Exception:
            pass
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
                ttk.Label(modal, text=valor, wraplength=480, justify="left").grid(row=idx, column=1, pady=2, padx=10, sticky="w")
        
        botones = ttk.Frame(modal)
        botones.grid(row=row_button, column=0, columnspan=2, pady=10, padx=10)
        ttk.Button(botones, text="Cerrar", command=modal.destroy).pack(side="left", padx=5)

        registrado_en = novedad[1] if len(novedad) > 1 else None
        puede_modificar_todo = self.app.tiene_permiso("registros.modificar")
        if vista == "novedad":
            # Si le falta fecha de fin o el usuario puede asignar reemplazos, se permite abrir edición fuera de la ventana
            fecha_fin_valor = novedad[10] if len(novedad) > 10 else None
            turno_valor = novedad[6] if len(novedad) > 6 else ""
            falta_fecha_fin = not fecha_fin_valor or str(fecha_fin_valor).strip() in ("", "-")
            puede_editar_nov = (
                self.app.tiene_permiso("novedades.editar") or puede_modificar_todo
            )
            puede_asignar_reemp = (
                self.app.tiene_permiso("reemplazos.asignar") or puede_modificar_todo
            ) and not es_turno_ciclo(turno_valor)
            dentro_ventana_edicion = puede_modificar_todo or self.app.records_service.dentro_de_ventana(registrado_en, "editar_horas")
            puede_editar = (
                (puede_editar_nov and (dentro_ventana_edicion or falta_fecha_fin))
                or puede_asignar_reemp
            )
            puede_eliminar = (
                self.app.tiene_permiso("novedades.eliminar") or puede_modificar_todo
            ) and (
                puede_modificar_todo or self.app.records_service.dentro_de_ventana(registrado_en, "eliminar_horas")
            )
        else:
            puede_editar = (
                self.app.tiene_permiso("cambios_turno.editar") or puede_modificar_todo
            ) and (
                puede_modificar_todo or self.app.records_service.dentro_de_ventana(registrado_en, "editar_horas")
            )
            puede_eliminar = (
                self.app.tiene_permiso("cambios_turno.eliminar") or puede_modificar_todo
            ) and (
                puede_modificar_todo or self.app.records_service.dentro_de_ventana(registrado_en, "eliminar_horas")
            )

        if puede_editar:
            def editar_registro():
                modal.destroy()
                if vista == "novedad":
                    edicion_completa = puede_editar_nov and dentro_ventana_edicion
                    modo = "completa" if edicion_completa else "fecha_fin"
                    self.app.forms_manager.iniciar_edicion_novedad(int(novedad[0]), modo=modo)
                else:
                    self.app.forms_manager.iniciar_edicion_cambio(int(novedad[0]))

            ttk.Button(botones, text="Editar", command=editar_registro).pack(side="left", padx=5)

        if puede_eliminar:
            def eliminar_registro():
                if not messagebox.askyesno(
                    "Eliminar",
                    f"¿Eliminar el registro #{novedad[0]}?\n\n"
                    "El registro dejará de listarse pero podrá recuperarse desde Administración.",
                    parent=modal,
                ):
                    return
                try:
                    if vista == "novedad":
                        self.app.records_service.eliminar_novedad(int(novedad[0]), self.app.current_user.get("id"), self.app.obtener_usuario_windows(), bypass_window=puede_modificar_todo)
                        self.cargar_datos_completos_novedades()
                    else:
                        self.app.records_service.eliminar_cambio(int(novedad[0]), self.app.current_user.get("id"), self.app.obtener_usuario_windows(), bypass_window=puede_modificar_todo)
                        self.cargar_datos_completos_cambios()
                    modal.destroy()
                except Exception as error:
                    messagebox.showerror("Eliminar", str(error), parent=modal)
            ttk.Button(botones, text="Eliminar", command=eliminar_registro).pack(side="left", padx=5)
    
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
            mapa_cobertura = {}
            if hasattr(self.app, "records_service") and self.app.records_service is not None:
                try:
                    mapa_cobertura = self.app.records_service.obtener_mapa_cobertura_novedades()
                except Exception:
                    mapa_cobertura = {}
            if self.app.sheet_novedades:
                for fila in self.app.sheet_novedades.iter_rows(min_row=2, values_only=True):
                    fila_procesada = ["-" if celda is None else celda for celda in fila]
                    valores_tabla = self._construir_fila_tabla_novedad(fila_procesada, mapa_cobertura)
                    self.app.tabla_novedades.insert("", "end", values=valores_tabla[:COLUMNAS_NOVEDADES_VISIBLES])
                    total += 1
            else:
                print("La hoja 'NOVEDADES' está vacía o no se pudo cargar.")

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
                    self.app.table_cambios.insert("", "end", values=fila_procesada[:COLUMNAS_CAMBIOS_VISIBLES])
                    total += 1
            else:
                print("La hoja 'Cambio de Turnos' está vacía o no se pudo cargar.")

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
            mapa_cobertura = {}
            if hasattr(self.app, "records_service") and self.app.records_service is not None:
                try:
                    mapa_cobertura = self.app.records_service.obtener_mapa_cobertura_novedades()
                except Exception:
                    mapa_cobertura = {}
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
                        valores_tabla = self._construir_fila_tabla_novedad(fila_procesada, mapa_cobertura)
                        self.app.tabla_novedades.insert("", "end", values=valores_tabla[:COLUMNAS_NOVEDADES_VISIBLES])
                        total += 1

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
                            self.app.table_cambios.insert("", "end", values=fila_procesada[:COLUMNAS_CAMBIOS_VISIBLES])
                            total += 1
                    elif nombre_filtro == self.app.PLACEHOLDER_BUSCAR_NOMBRE:
                        if coincide_dotacion:
                            self.app.table_cambios.insert("", "end", values=fila_procesada[:COLUMNAS_CAMBIOS_VISIBLES])
                            total += 1
                    else:
                        if coincide_nombre and coincide_dotacion:
                            self.app.table_cambios.insert("", "end", values=fila_procesada[:COLUMNAS_CAMBIOS_VISIBLES])
                            total += 1

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
        """Actualiza las tablas manteniendo los filtros activos (llamado en cada refresh periódico)."""
        if self.app.sheet_novedades and self.app.current_view == 'table' and hasattr(self.app, "tabla_novedades"):
            if hasattr(self.app, "apellido_filter_novedades_var") and hasattr(self.app, "dotacion_filter_novedades_var") and hasattr(self.app, "tipo_filter_novedades_var"):
                self.filtrar_datos_novedades(
                    self.app.apellido_filter_novedades_var.get(),
                    self.app.dotacion_filter_novedades_var.get(),
                    self.app.tipo_filter_novedades_var.get()
                )
        
        if self.app.sheet_cambio_turnos and self.app.current_view == "table_cambios" and hasattr(self.app, "table_cambios"):
            if hasattr(self.app, "apellido_filter_cambios_var") and hasattr(self.app, "dotacion_filter_cambios_var"):
                self.filtrar_datos_cambios(
                    self.app.apellido_filter_cambios_var.get(),
                    self.app.dotacion_filter_cambios_var.get()
                )

        if self.app.current_view == "ciclo" and hasattr(self, "tabla_ciclo_libres"):
            self.actualizar_vista_personal_ciclo()
    
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
                if item_values:
                    fila = self.app.records_service.obtener_novedad(int(item_values[0]))
                    if fila:
                        self.mostrar_modal_detalle(self._valores_novedad(fila), COLUMNAS_DETALLE_NOVEDAD, "novedad")
            else:
                print("No se seleccionó ningún elemento.")
            
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD",
            "DOTACION", "TURNOS", "FRANCO", "NOVEDAD", "Fecha de Inicio Novedad", "Fecha de Fin Novedad",
            "CUBIERTO", "REFERENCIA ESTACION", "SUPERVISOR"
        ]
        self.app.apellido_filter_novedades_var = tk.StringVar()
        self.app.dotacion_filter_novedades_var = tk.StringVar()
        self.app.tipo_filter_novedades_var = tk.StringVar()
        
        # Título y botones
        ttk.Label(self.app.table_frame, text="Registro de novedades", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, pady=10, padx=10, sticky="w"
        )
        ttk.Button(
            self.app.table_frame, text="Panel de control",
            command=lambda: self.app.toggle_view("dashboard")
        ).grid(row=0, column=1, pady=10, padx=6, sticky="e")
        if self.app.tiene_permiso("cambios_turno.ver"):
            ttk.Button(
                self.app.table_frame, text="Ver cambios de turno",
                command=lambda: self.app.toggle_view("table_cambios")
            ).grid(row=0, column=2, pady=10, padx=6, sticky="e")
        
        # Filtro nombre
        apellido_filter = tk.Entry(self.app.table_frame, textvariable=self.app.apellido_filter_novedades_var, font=("Helvetica", 10))
        apellido_filter.grid(row=0, column=3, sticky="e", pady=10, padx=6, ipady=5, ipadx=10)
        
        # Filtro dotación
        self.app.dotacion_filter_novedades = ttk.Combobox(
            self.app.table_frame, textvariable=self.app.dotacion_filter_novedades_var,
            values=self.app.DOTACIONES, width=10, state="normal"
        )
        dotacion_filter = self.app.dotacion_filter_novedades
        dotacion_filter.grid(row=0, column=4, sticky="e", pady=5)

        self.app.tipo_filter_novedades = ttk.Combobox(
            self.app.table_frame, textvariable=self.app.tipo_filter_novedades_var,
            values=["Todos", *self.app.tipo_novedades], width=18, state="readonly"
        )
        self.app.tipo_filter_novedades.grid(row=0, column=5, sticky="e", pady=5, padx=6)
        
        self.app.resultados_novedades_label = ttk.Label(self.app.table_frame, text="0 resultados", font=("Helvetica", 9))
        self.app.resultados_novedades_label.grid(row=0, column=6, sticky="w", padx=6)
        
        if self.app.tiene_permiso("novedades.crear"):
            ttk.Button(self.app.table_frame, text="Nueva novedad", command=lambda: self.app.toggle_view("form")).grid(
                row=0, column=7, pady=10, padx=2, sticky="e"
            )
        if self.app.tiene_permiso("cambios_turno.crear"):
            ttk.Button(self.app.table_frame, text="Nuevo cambio de turno", command=lambda: self.app.toggle_view("form_cambios")).grid(
                row=0, column=8, pady=10, padx=6, sticky="e"
            )
        if self.app.tiene_permiso("novedades.exportar"):
            ttk.Button(self.app.table_frame, text="Turnos sin cubrir", command=self.app.exportar_turnos_sin_cubrir).grid(
                row=0, column=9, pady=10, padx=4, sticky="e"
            )
        if self.app.tiene_permiso("reemplazos.ver"):
            ttk.Button(self.app.table_frame, text="Personal de Ciclo", command=lambda: self.app.toggle_view("ciclo")).grid(
                row=0, column=10, pady=10, padx=4, sticky="e"
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
        self.app.tree_frame.grid(row=1, column=0, columnspan=11, sticky="nsew")
        self.app.tree_frame.grid_propagate(False)

        self.app.tree_frame.grid_rowconfigure(0, weight=1)
        self.app.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.app.tabla_novedades = ttk.Treeview(self.app.tree_frame, columns=columnas, show="headings", height=40)
        self.app.tabla_novedades.grid(row=0, column=0, sticky="nsew")
        self.app.tabla_novedades.bind("<Double-1>", on_double_click)
        
        # Encabezados y anchos por columna
        for col in columnas:
            self.app.tabla_novedades.heading(col, text=col.capitalize())
            self.app.tabla_novedades.column(col, width=ANCHOS_NOVEDADES[col], stretch=False)

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
                if item_values:
                    fila = self.app.records_service.obtener_cambio(int(item_values[0]))
                    if fila:
                        self.mostrar_modal_detalle(self._valores_cambio(fila), COLUMNAS_DETALLE_CAMBIO, "cambio de turno")
            else:
                print("No se seleccionó ningún elemento.")
        
        self.app.apellido_filter_cambios_var = tk.StringVar()
        self.app.dotacion_filter_cambios_var = tk.StringVar()
        columnas = [
            "ID", "Fecha de registro", "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", 
            "TURNOS", "FRANCO", "LEGAJO2", "APELLIDOS Y NOMBRES2", "ESPECIALIDAD2", "DOTACION2", 
            "TURNOS2", "FRANCO2", "Fecha de Cambio de Turno", "REFERENCIA ESTACION", "SUPERVISOR"
        ]
        
        # Detectar cambios en filtros
        self.app.apellido_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        self.app.dotacion_filter_cambios_var.trace_add("write", lambda *args: self.programar_filtrado_cambios())
        
        # Título y botones
        ttk.Label(self.app.table_cambios_frame, text="Registro de cambios de turnos", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, pady=10, padx=10, sticky="w"
        )
        ttk.Button(
            self.app.table_cambios_frame, text="Panel de control",
            command=lambda: self.app.toggle_view("dashboard")
        ).grid(row=0, column=1, pady=10, padx=10, sticky="e")
        if self.app.tiene_permiso("novedades.ver"):
            ttk.Button(
                self.app.table_cambios_frame, text="Ver novedades",
                command=lambda: self.app.toggle_view("table")
            ).grid(row=0, column=2, pady=10, padx=1, sticky="e")
        
        # Filtro nombre
        apellido_filter = tk.Entry(self.app.table_cambios_frame, textvariable=self.app.apellido_filter_cambios_var, font=("Helvetica", 10))
        apellido_filter.grid(row=0, column=3, sticky="e", pady=10, padx=10, ipady=5, ipadx=10)
        
        # Filtro dotación
        self.app.dotacion_filter_cambios = ttk.Combobox(
            self.app.table_cambios_frame, textvariable=self.app.dotacion_filter_cambios_var,
            values=self.app.DOTACIONES, width=10, state="normal"
        )
        dotacion_filter = self.app.dotacion_filter_cambios
        dotacion_filter.grid(row=0, column=4, sticky="e", pady=5)
        
        self.app.resultados_cambios_label = ttk.Label(self.app.table_cambios_frame, text="0 resultados", font=("Helvetica", 9))
        self.app.resultados_cambios_label.grid(row=0, column=5, sticky="w", padx=8)
        
        if self.app.tiene_permiso("novedades.crear"):
            ttk.Button(self.app.table_cambios_frame, text="Nueva novedad", command=lambda: self.app.toggle_view("form")).grid(
                row=0, column=6, pady=10, padx=1, sticky="e"
            )
        if self.app.tiene_permiso("cambios_turno.crear"):
            ttk.Button(self.app.table_cambios_frame, text="Nuevo cambio de turno", command=lambda: self.app.toggle_view("form_cambios")).grid(
                row=0, column=7, pady=10, padx=1, sticky="e"
            )

        dotacion_filter.insert(0, "Todas")
        apellido_filter.insert(0, self.app.PLACEHOLDER_BUSCAR_NOMBRE)
        apellido_filter.config(fg="grey")
        apellido_filter.bind("<FocusIn>", on_focus_in)
        apellido_filter.bind("<FocusOut>", on_focus_out)
        
        # Contenedor del Treeview
        self.app.tree_frame = ttk.Frame(self.app.table_cambios_frame, width=self.app.WIDTH, height=self.app.HEIGHT)
        self.app.tree_frame.grid(row=1, column=0, columnspan=8, sticky="nsew")
        self.app.tree_frame.grid_propagate(False)

        self.app.tree_frame.grid_rowconfigure(0, weight=1)
        self.app.tree_frame.grid_columnconfigure(0, weight=1)

        # Crear el Treeview
        self.app.table_cambios = ttk.Treeview(self.app.tree_frame, columns=columnas, show="headings", height=38)
        self.app.table_cambios.grid(row=0, column=0, sticky="nsew")
        self.app.table_cambios.bind("<Double-1>", on_double_click)
        
        # Encabezados y anchos por columna
        for col in columnas:
            self.app.table_cambios.heading(col, text=col)
            self.app.table_cambios.column(col, width=ANCHOS_CAMBIOS[col], anchor='center', stretch=False)

        # Scrollbars
        scrollbar_vertical = ttk.Scrollbar(self.app.tree_frame, orient="vertical", command=self.app.table_cambios.yview)
        scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        scrollbar_horizontal = ttk.Scrollbar(self.app.tree_frame, orient="horizontal", command=self.app.table_cambios.xview)
        scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.app.table_cambios.configure(yscrollcommand=scrollbar_vertical.set, xscrollcommand=scrollbar_horizontal.set)
        
        # Cargar datos
        self.cargar_datos_completos_cambios()
        self.app.tabla_cambios_creada = True

    def crear_vista_personal_ciclo(self):
        """Crea la vista de Personal de Ciclo dividida en dos tablas:
        - Arriba: personal de CICLO sin relevo asignado en la fecha consultada.
        - Abajo: personal de CICLO con relevo asignado en la fecha consultada.
        Al hacer doble clic en cualquier fila se abre el calendario mensual del empleado.
        """
        header = ttk.Frame(self.app.ciclo_frame)
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=8)

        ttk.Label(header, text="Personal de Ciclo", font=("Helvetica", 20, "bold")).grid(
            row=0, column=0, padx=(0, 15), pady=4, sticky="w"
        )
        ttk.Button(
            header, text="Panel de control",
            command=lambda: self.app.toggle_view("dashboard"),
        ).grid(row=0, column=1, padx=4, pady=4)
        if self.app.tiene_permiso("novedades.ver"):
            ttk.Button(
                header, text="Ver novedades",
                command=lambda: self.app.toggle_view("table"),
            ).grid(row=0, column=2, padx=4, pady=4)
        if self.app.tiene_permiso("cambios_turno.ver"):
            ttk.Button(
                header, text="Ver cambios de turno",
                command=lambda: self.app.toggle_view("table_cambios"),
            ).grid(row=0, column=3, padx=4, pady=4)

        ttk.Label(header, text="Dotación:").grid(row=0, column=4, padx=(18, 4), pady=4, sticky="e")
        self.app.dotacion_filter_ciclo_var = tk.StringVar(value="Todas")
        self.app.dotacion_filter_ciclo = ttk.Combobox(
            header,
            textvariable=self.app.dotacion_filter_ciclo_var,
            values=self.app.DOTACIONES,
            width=12,
            state="readonly",
        )
        self.app.dotacion_filter_ciclo.grid(row=0, column=5, padx=4, pady=4)
        self.app.dotacion_filter_ciclo.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_personal_ciclo())

        ttk.Label(header, text="Especialidad:").grid(row=0, column=6, padx=(12, 4), pady=4, sticky="e")
        self.app.especialidad_filter_ciclo_var = tk.StringVar(value="Todas")
        especialidades = ["Todas"] + sorted({str(r[2]).strip() for r in self.app.base_rows if r and len(r) > 2 and str(r[2]).strip()})
        self.app.especialidad_filter_ciclo = ttk.Combobox(
            header,
            textvariable=self.app.especialidad_filter_ciclo_var,
            values=especialidades,
            width=16,
            state="readonly",
        )
        self.app.especialidad_filter_ciclo.grid(row=0, column=7, padx=4, pady=4)
        self.app.especialidad_filter_ciclo.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_personal_ciclo())

        ttk.Label(header, text="Fecha de consulta:").grid(row=0, column=8, padx=(12, 4), pady=4, sticky="e")
        self.fecha_ciclo_entry = DateEntry(header, dateformat='%d/%m/%Y', bootstyle="primary", width=12)
        self.fecha_ciclo_entry.grid(row=0, column=9, padx=4, pady=4)
        self.fecha_ciclo_entry.entry.bind("<Key>", lambda _e: "break")
        self.fecha_ciclo_entry.entry.bind("<FocusOut>", lambda _e: self.actualizar_vista_personal_ciclo())
        try:
            self.fecha_ciclo_entry.bind("<<DateEntrySelected>>", lambda _e: self.actualizar_vista_personal_ciclo())
        except Exception:
            pass

        ttk.Button(
            header, text="Actualizar", bootstyle="info-outline",
            command=self.actualizar_vista_personal_ciclo,
        ).grid(row=0, column=10, padx=6, pady=4)

        ttk.Button(
            header, text="Ver calendario mensual", bootstyle="primary-outline",
            command=self._abrir_calendario_seleccionado,
        ).grid(row=0, column=11, padx=6, pady=4)

        ttk.Label(
            header,
            text="(Doble clic sobre un empleado para abrir su calendario mensual)",
            font=("Helvetica", 9, "italic"),
        ).grid(row=0, column=12, padx=10, pady=4, sticky="w")

        # Contenedor principal de las dos tablas
        contenedor = ttk.Frame(self.app.ciclo_frame, width=self.app.WIDTH, height=self.app.HEIGHT)
        contenedor.grid(row=1, column=0, padx=10, pady=4, sticky="nsew")
        contenedor.grid_columnconfigure(0, weight=1)
        contenedor.grid_rowconfigure(1, weight=1)
        contenedor.grid_rowconfigure(3, weight=1)

        # --- TABLA SUPERIOR: SIN RELEVO ASIGNADO ---
        top_bar = ttk.Frame(contenedor)
        top_bar.grid(row=0, column=0, sticky="ew", pady=(4, 2))
        ttk.Label(
            top_bar,
            text="Personal de CICLO sin relevo asignado en la fecha",
            font=("Helvetica", 12, "bold"),
        ).pack(side="left", padx=4)
        self.ciclo_libres_count_label = ttk.Label(top_bar, text="0 resultados", font=("Helvetica", 10))
        self.ciclo_libres_count_label.pack(side="left", padx=12)

        frame_libres = ttk.Frame(contenedor)
        frame_libres.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        frame_libres.grid_rowconfigure(0, weight=1)
        frame_libres.grid_columnconfigure(0, weight=1)

        cols_libres = ("LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION", "TURNOS", "FRANCO", "ESTADO EN FECHA")
        self.tabla_ciclo_libres = ttk.Treeview(
            frame_libres, columns=cols_libres, show="headings", height=14
        )
        anchos_libres = {
            "LEGAJO": 90,
            "APELLIDOS Y NOMBRES": 300,
            "ESPECIALIDAD": 180,
            "DOTACION": 110,
            "TURNOS": 140,
            "FRANCO": 120,
            "ESTADO EN FECHA": 240,
        }
        for col in cols_libres:
            self.tabla_ciclo_libres.heading(col, text=col)
            self.tabla_ciclo_libres.column(
                col, width=anchos_libres[col],
                anchor="w" if col in ("APELLIDOS Y NOMBRES", "ESTADO EN FECHA") else "center",
                stretch=False,
            )
        self.tabla_ciclo_libres.grid(row=0, column=0, sticky="nsew")
        sb_libres_v = ttk.Scrollbar(frame_libres, orient="vertical", command=self.tabla_ciclo_libres.yview)
        sb_libres_v.grid(row=0, column=1, sticky="ns")
        sb_libres_h = ttk.Scrollbar(frame_libres, orient="horizontal", command=self.tabla_ciclo_libres.xview)
        sb_libres_h.grid(row=1, column=0, sticky="ew")
        self.tabla_ciclo_libres.configure(yscrollcommand=sb_libres_v.set, xscrollcommand=sb_libres_h.set)
        self.tabla_ciclo_libres.bind("<Double-1>", lambda e: self._on_double_click_ciclo(self.tabla_ciclo_libres))

        # --- TABLA INFERIOR: CON RELEVO ASIGNADO ---
        bot_bar = ttk.Frame(contenedor)
        bot_bar.grid(row=2, column=0, sticky="ew", pady=(8, 2))
        ttk.Label(
            bot_bar,
            text="Personal de CICLO con relevo asignado en la fecha",
            font=("Helvetica", 12, "bold"),
        ).pack(side="left", padx=4)
        self.ciclo_ocupados_count_label = ttk.Label(bot_bar, text="0 resultados", font=("Helvetica", 10))
        self.ciclo_ocupados_count_label.pack(side="left", padx=12)

        frame_ocupados = ttk.Frame(contenedor)
        frame_ocupados.grid(row=3, column=0, sticky="nsew", pady=(0, 4))
        frame_ocupados.grid_rowconfigure(0, weight=1)
        frame_ocupados.grid_columnconfigure(0, weight=1)

        cols_ocupados = (
            "LEGAJO", "APELLIDOS Y NOMBRES", "ESPECIALIDAD", "DOTACION",
            "TURNOS", "FRANCO", "REEMPLAZA A", "NOVEDAD TITULAR", "DESDE", "HASTA",
        )
        self.tabla_ciclo_ocupados = ttk.Treeview(
            frame_ocupados, columns=cols_ocupados, show="headings", height=14
        )
        anchos_ocupados = {
            "LEGAJO": 85,
            "APELLIDOS Y NOMBRES": 250,
            "ESPECIALIDAD": 150,
            "DOTACION": 95,
            "TURNOS": 120,
            "FRANCO": 100,
            "REEMPLAZA A": 270,
            "NOVEDAD TITULAR": 160,
            "DESDE": 105,
            "HASTA": 105,
        }
        for col in cols_ocupados:
            self.tabla_ciclo_ocupados.heading(col, text=col)
            self.tabla_ciclo_ocupados.column(
                col, width=anchos_ocupados[col],
                anchor="w" if col in ("APELLIDOS Y NOMBRES", "REEMPLAZA A", "NOVEDAD TITULAR") else "center",
                stretch=False,
            )
        self.tabla_ciclo_ocupados.grid(row=0, column=0, sticky="nsew")
        sb_ocup_v = ttk.Scrollbar(frame_ocupados, orient="vertical", command=self.tabla_ciclo_ocupados.yview)
        sb_ocup_v.grid(row=0, column=1, sticky="ns")
        sb_ocup_h = ttk.Scrollbar(frame_ocupados, orient="horizontal", command=self.tabla_ciclo_ocupados.xview)
        sb_ocup_h.grid(row=1, column=0, sticky="ew")
        self.tabla_ciclo_ocupados.configure(yscrollcommand=sb_ocup_v.set, xscrollcommand=sb_ocup_h.set)
        self.tabla_ciclo_ocupados.bind("<Double-1>", lambda e: self._on_double_click_ciclo(self.tabla_ciclo_ocupados))

        self.actualizar_vista_personal_ciclo()
        self.app.vista_ciclo_creada = True

    def _obtener_fecha_consulta_ciclo(self):
        if hasattr(self, "fecha_ciclo_entry"):
            val = self.fecha_ciclo_entry.entry.get().strip()
            if val:
                return val
        return datetime.now().strftime("%d/%m/%Y")

    def actualizar_vista_personal_ciclo(self):
        """Recarga las tablas de personal de ciclo sin relevo (arriba) y con relevo (abajo)."""
        if not hasattr(self, "tabla_ciclo_libres") or not hasattr(self, "tabla_ciclo_ocupados"):
            return
        fecha_str = self._obtener_fecha_consulta_ciclo()
        dotacion_sel = (
            self.app.dotacion_filter_ciclo_var.get().strip()
            if hasattr(self.app, "dotacion_filter_ciclo_var")
            else "Todas"
        )
        esp_sel = (
            self.app.especialidad_filter_ciclo_var.get().strip()
            if hasattr(self.app, "especialidad_filter_ciclo_var")
            else "Todas"
        )
        try:
            estado = self.app.records_service.obtener_estado_personal_ciclo(
                fecha_str,
                dotacion=None if dotacion_sel in ("", "Todas") else dotacion_sel,
                especialidad=None if esp_sel in ("", "Todas") else esp_sel,
            )
        except Exception as e:
            messagebox.showerror("Personal de Ciclo", f"Error al consultar personal de ciclo: {e}", parent=self.app.root)
            return

        self.tabla_ciclo_libres.delete(*self.tabla_ciclo_libres.get_children())
        for emp in estado.get("libres", []):
            self.tabla_ciclo_libres.insert(
                "", "end",
                values=(
                    emp["legajo"],
                    emp["apellidos_nombres"],
                    emp["especialidad"],
                    emp["dotacion"],
                    emp["turnos"],
                    emp["franco"],
                    emp.get("estado_dia") or "Disponible",
                ),
            )
        self.actualizar_contador_resultados(self.ciclo_libres_count_label, len(estado.get("libres", [])))

        self.tabla_ciclo_ocupados.delete(*self.tabla_ciclo_ocupados.get_children())
        for emp in estado.get("ocupados", []):
            titular_str = f"{emp.get('reemplazando_legajo') or ''} - {emp.get('reemplazando_nombre') or ''}".strip(" -")
            self.tabla_ciclo_ocupados.insert(
                "", "end",
                values=(
                    emp["legajo"],
                    emp["apellidos_nombres"],
                    emp["especialidad"],
                    emp["dotacion"],
                    emp["turnos"],
                    emp["franco"],
                    titular_str or "-",
                    emp.get("novedad_titular") or "-",
                    emp.get("reemplazo_desde") or "-",
                    emp.get("reemplazo_hasta") or "Sin fecha fin",
                ),
            )
        self.actualizar_contador_resultados(self.ciclo_ocupados_count_label, len(estado.get("ocupados", [])))

    def _abrir_calendario_seleccionado(self):
        for tree in (getattr(self, "tabla_ciclo_ocupados", None), getattr(self, "tabla_ciclo_libres", None)):
            if tree is not None and tree.selection():
                self._on_double_click_ciclo(tree)
                return
        messagebox.showinfo(
            "Calendario mensual",
            "Seleccione un empleado en cualquiera de las dos tablas (o haga doble clic sobre su fila).",
            parent=self.app.root,
        )

    def _on_double_click_ciclo(self, treeview):
        sel = treeview.selection()
        if not sel:
            return
        valores = treeview.item(sel[0], "values")
        if not valores:
            return
        try:
            legajo_ciclo = int(valores[0])
        except (TypeError, ValueError):
            return
        self.mostrar_calendario_ciclo(legajo_ciclo, self._obtener_fecha_consulta_ciclo())

    def mostrar_calendario_ciclo(self, legajo_ciclo, fecha_ref_str=None):
        """Abre un modal con el calendario mensual (una fila por día del mes seleccionado)
        mostrando a quién reemplaza el empleado de CICLO en cada día.
        """
        hoy = datetime.now()
        mes_inicial = hoy.month
        anio_inicial = hoy.year
        if fecha_ref_str:
            try:
                dt_ref = datetime.strptime(str(fecha_ref_str).strip(), "%d/%m/%Y")
                mes_inicial = dt_ref.month
                anio_inicial = dt_ref.year
            except ValueError:
                pass

        emp_info = self.app.base_index.get(int(legajo_ciclo))
        nombre_emp = emp_info[1] if emp_info else f"Legajo {legajo_ciclo}"
        esp_emp = emp_info[2] if emp_info else "-"
        dot_emp = emp_info[3] if emp_info else "-"
        tur_emp = emp_info[4] if emp_info else "-"
        fra_emp = emp_info[5] if emp_info else "-"
        fra_hor = emp_info[6] if emp_info and len(emp_info) > 6 and emp_info[6] else ""
        franco_mostrar = f"{fra_emp} (desde {fra_hor})" if fra_hor else fra_emp

        modal = tk.Toplevel(self.app.root)
        modal.title(f"Calendario Mensual de Reemplazos - {legajo_ciclo} {nombre_emp}")
        modal.geometry("980x680")
        try:
            modal.state("zoomed")
        except Exception:
            pass
        self.app.aplicar_tema_ventana(modal)

        top = ttk.Frame(modal, padding=10)
        top.pack(fill="x")

        ttk.Label(
            top,
            text=f"{legajo_ciclo} — {nombre_emp}",
            font=("Helvetica", 15, "bold"),
        ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 4))

        ttk.Label(
            top,
            text=f"Dotación: {dot_emp}   |   Especialidad: {esp_emp}   |   Turno: {tur_emp}   |   Franco: {franco_mostrar}",
            font=("Helvetica", 10),
        ).grid(row=1, column=0, columnspan=6, sticky="w", pady=(0, 8))

        meses_nombres = [
            "01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril",
            "05 - Mayo", "06 - Junio", "07 - Julio", "08 - Agosto",
            "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre",
        ]
        mes_var = tk.StringVar(value=meses_nombres[mes_inicial - 1])
        anio_var = tk.StringVar(value=str(anio_inicial))

        ctrl_bar = ttk.Frame(top)
        ctrl_bar.grid(row=2, column=0, columnspan=6, sticky="w")

        ttk.Label(ctrl_bar, text="Mes:").pack(side="left", padx=(0, 4))
        mes_combo = ttk.Combobox(ctrl_bar, textvariable=mes_var, values=meses_nombres, width=16, state="readonly")
        mes_combo.pack(side="left", padx=(0, 12))

        ttk.Label(ctrl_bar, text="Año:").pack(side="left", padx=(0, 4))
        anios_lista = [str(a) for a in range(anio_inicial - 3, anio_inicial + 5)]
        anio_combo = ttk.Combobox(ctrl_bar, textvariable=anio_var, values=anios_lista, width=8, state="readonly")
        anio_combo.pack(side="left", padx=(0, 12))

        resumen_mes_label = ttk.Label(ctrl_bar, text="", font=("Helvetica", 10, "bold"))

        tree_frame = ttk.Frame(modal, padding=(10, 4))
        tree_frame.pack(fill="both", expand=True)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        cols_cal = (
            "FECHA", "DIA", "ESTADO", "ID NOVEDAD", "LEGAJO TITULAR",
            "REEMPLAZA A (APELLIDO Y NOMBRE)", "NOVEDAD TITULAR", "TRAMO REEMPLAZO",
        )
        cal_tree = ttk.Treeview(tree_frame, columns=cols_cal, show="headings", height=22)
        anchos_cal = {
            "FECHA": 100,
            "DIA": 100,
            "ESTADO": 130,
            "ID NOVEDAD": 90,
            "LEGAJO TITULAR": 110,
            "REEMPLAZA A (APELLIDO Y NOMBRE)": 240,
            "NOVEDAD TITULAR": 140,
            "TRAMO REEMPLAZO": 130,
        }
        for col in cols_cal:
            cal_tree.heading(col, text=col)
            cal_tree.column(
                col, width=anchos_cal[col],
                anchor="w" if col in ("REEMPLAZA A (APELLIDO Y NOMBRE)", "NOVEDAD TITULAR", "ESTADO") else "center",
                stretch=False,
            )

        hoy_str = datetime.now().strftime("%d/%m/%Y")
        es_oscuro = False
        if hasattr(self.app, "style") and hasattr(self.app.style, "theme"):
            es_oscuro = getattr(self.app.style.theme, "type", "") == "dark"
        if es_oscuro:
            cal_tree.tag_configure("hoy", background="#4a3b00", foreground="#ffffff")
        else:
            cal_tree.tag_configure("hoy", background="#fff3cd", foreground="#000000")

        cal_tree.grid(row=0, column=0, sticky="nsew")
        sb_v = ttk.Scrollbar(tree_frame, orient="vertical", command=cal_tree.yview)
        sb_v.grid(row=0, column=1, sticky="ns")
        cal_tree.configure(yscrollcommand=sb_v.set)

        def recargar_calendario(*_args):
            try:
                mes_num = int(mes_var.get().split("-")[0].strip())
                anio_num = int(anio_var.get().strip())
            except (ValueError, IndexError):
                return
            cal_tree.delete(*cal_tree.get_children())
            dias = self.app.records_service.obtener_calendario_mensual_ciclo(legajo_ciclo, anio_num, mes_num)
            dias_reemplazo = 0
            for d in dias:
                if d.get("reemplazando_legajo") is not None:
                    dias_reemplazo += 1
                    tramo_txt = f"{d.get('tramo_inicio') or ''} al {d.get('tramo_fin') or 'S/F'}"
                else:
                    tramo_txt = "-"
                tags = ("hoy",) if d["fecha"] == hoy_str else ()
                cal_tree.insert(
                    "", "end",
                    values=(
                        d["fecha"],
                        d["dia_semana"],
                        d["estado"],
                        d.get("novedad_id") or "-",
                        d.get("reemplazando_legajo") or "-",
                        d.get("reemplazando_nombre") or "-",
                        d.get("novedad_titular") or "-",
                        tramo_txt,
                    ),
                    tags=tags,
                )
            resumen_mes_label.config(text=f"Días con reemplazo en el mes: {dias_reemplazo} / {len(dias)}")

        def on_double_click_dia(event=None):
            sel = cal_tree.selection()
            if not sel:
                return
            valores = cal_tree.item(sel[0], "values")
            if not valores:
                return
            nov_id_str = str(valores[3]).strip()
            if not nov_id_str or nov_id_str == "-":
                return
            try:
                nov_id = int(nov_id_str)
            except (ValueError, TypeError):
                return
            fila = self.app.records_service.obtener_novedad(nov_id)
            if not fila:
                messagebox.showerror("Novedad", f"No se encontró la novedad #{nov_id}.", parent=modal)
                return
            self.mostrar_modal_detalle(self._valores_novedad(fila), COLUMNAS_DETALLE_NOVEDAD, "novedad")

        cal_tree.bind("<Double-1>", on_double_click_dia)

        def cambiar_mes(delta):
            try:
                m = int(mes_var.get().split("-")[0].strip()) + delta
                a = int(anio_var.get().strip())
                if m < 1:
                    m = 12
                    a -= 1
                elif m > 12:
                    m = 1
                    a += 1
                mes_var.set(meses_nombres[m - 1])
                anio_var.set(str(a))
                recargar_calendario()
            except Exception:
                pass

        ttk.Button(ctrl_bar, text="◀ Mes anterior", command=lambda: cambiar_mes(-1)).pack(side="left", padx=4)
        ttk.Button(ctrl_bar, text="Mes siguiente ▶", command=lambda: cambiar_mes(1)).pack(side="left", padx=4)
        resumen_mes_label.pack(side="left", padx=16)

        mes_combo.bind("<<ComboboxSelected>>", recargar_calendario)
        anio_combo.bind("<<ComboboxSelected>>", recargar_calendario)

        ttk.Button(modal, text="Cerrar", command=modal.destroy).pack(pady=8)
        recargar_calendario()
