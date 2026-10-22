"""Lógica de formularios de novedades y cambios de turnos."""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from ttkbootstrap import DateEntry
from datetime import datetime, timedelta
import validators
from sqlite_store import DatabaseBusyError


class FormsManager:
    """Gestor de formularios, guardado de datos y búsqueda de legajos.
    
    Maneja la presentación de formularios de novedades y cambios de turnos,
    búsqueda de personal por legajo (con modal filtrable), validación,
    guardado en la base de datos y manejo de errores.
    
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
        self.edicion = None
        self._personal_estacion_lista_original = self.app.nombres_personal_estacion() if hasattr(self.app, "nombres_personal_estacion") else []
        self.tramos_reemplazo = []
        self._candidatos_ciclo_map = {}
        self._candidatos_ciclo_lista = []

    def _puede_asignar_reemplazo(self):
        if not hasattr(self.app, "tiene_permiso"):
            return True
        return bool(self.app.tiene_permiso("reemplazos.asignar") or self.app.tiene_permiso("registros.modificar"))

    def _formatear_candidato_ciclo(self, c):
        leg = c.get("legajo")
        nom = c.get("apellidos_nombres") or ""
        tur = c.get("turnos") or ""
        fra = c.get("franco") or ""
        return f"{leg} - {nom} ({tur} - Franco: {fra})"

    def _sincronizar_fechas_reemplazo_default(self):
        """Sincroniza las fechas de inicio/fin del selector de reemplazo cuando aún no hay tramos agregados."""
        if not hasattr(self, "reemplazo_desde_entry") or not hasattr(self, "reemplazo_hasta_entry"):
            return
        if self.tramos_reemplazo:
            return
        try:
            f_ini = self.app.fecha_inicio_novedad_entry.entry.get().strip() if hasattr(self.app, "fecha_inicio_novedad_entry") else ""
            f_fin = self.app.fecha_fin_novedad_entry.entry.get().strip() if hasattr(self.app, "fecha_fin_novedad_entry") else ""
            if f_ini:
                self.reemplazo_desde_entry.entry.delete(0, tk.END)
                self.reemplazo_desde_entry.entry.insert(0, f_ini)
            self.reemplazo_hasta_entry.entry.delete(0, tk.END)
            if f_fin and f_fin != "-":
                self.reemplazo_hasta_entry.entry.insert(0, f_fin)
        except Exception:
            pass

    def _limpiar_fecha_fin_reemplazo(self):
        if hasattr(self, "reemplazo_hasta_entry"):
            self.reemplazo_hasta_entry.entry.delete(0, tk.END)

    def _renderizar_tramos_reemplazo(self):
        if not hasattr(self, "reemplazos_tree"):
            return
        for item in self.reemplazos_tree.get_children():
            self.reemplazos_tree.delete(item)
        for idx, t in enumerate(self.tramos_reemplazo):
            self.reemplazos_tree.insert(
                "", "end", iid=str(idx),
                values=(
                    t.get("legajo_reemplazo"),
                    t.get("apellidos_nombres_reemplazo"),
                    t.get("turnos_reemplazo") or "-",
                    t.get("franco_reemplazo") or "-",
                    t.get("fecha_inicio") or "-",
                    t.get("fecha_fin") or "Sin fecha fin",
                ),
            )

    def _configurar_controles_reemplazo(self, habilitar):
        if not hasattr(self, "reemplazo_combo"):
            return
        estado_combo = "normal" if habilitar else "disabled"
        estado_btn = "normal" if habilitar else "disabled"
        self.reemplazo_combo.config(state=estado_combo)
        self.reemplazo_desde_entry.entry.config(state=estado_btn)
        if hasattr(self.reemplazo_desde_entry, "button"):
            self.reemplazo_desde_entry.button.config(state=estado_btn)
        self.reemplazo_hasta_entry.entry.config(state=estado_btn)
        if hasattr(self.reemplazo_hasta_entry, "button"):
            self.reemplazo_hasta_entry.button.config(state=estado_btn)
        self.limpiar_reemplazo_hasta_btn.config(state=estado_btn)
        self.agregar_reemplazo_btn.config(state=estado_btn)
        if hasattr(self, "modificar_reemplazo_btn"):
            self.modificar_reemplazo_btn.config(state=estado_btn)
        self.quitar_reemplazo_btn.config(state=estado_btn)

    def _actualizar_candidatos_reemplazo(self, limpiar_tramos=True):
        """Actualiza la lista del combobox de reemplazos según dotación, especialidad y turno del titular."""
        if limpiar_tramos:
            self.tramos_reemplazo = []
            self._renderizar_tramos_reemplazo()
        if hasattr(self, "reemplazo_combo_var"):
            self.reemplazo_combo_var.set("")

        turno_titular = self.app.turnos_var.get().strip() if hasattr(self.app, "turnos_var") else ""
        dotacion = self.app.dotacion_var.get().strip() if hasattr(self.app, "dotacion_var") else ""
        especialidad = self.app.especialidad_var.get().strip() if hasattr(self.app, "especialidad_var") else ""
        legajo_tit = self.app.legajo_var.get().strip() if hasattr(self.app, "legajo_var") else None

        if validators.es_turno_ciclo(turno_titular):
            self._candidatos_ciclo_map = {}
            self._candidatos_ciclo_lista = []
            if hasattr(self, "reemplazo_combo"):
                self.reemplazo_combo["values"] = []
                self._configurar_controles_reemplazo(False)
                self.reemplazo_estado_label.config(
                    text="El personal con turno CICLO no puede tener reemplazo.",
                    foreground="#d9534f",
                )
            return

        candidatos = []
        if hasattr(self.app, "records_service") and self.app.records_service is not None and dotacion and especialidad:
            try:
                candidatos = self.app.records_service.obtener_candidatos_ciclo(
                    dotacion, especialidad, excluir_legajo=legajo_tit
                )
            except Exception:
                candidatos = []

        self._candidatos_ciclo_map = {}
        self._candidatos_ciclo_lista = []
        for c in candidatos:
            etiqueta = self._formatear_candidato_ciclo(c)
            self._candidatos_ciclo_map[etiqueta] = c
            self._candidatos_ciclo_lista.append(etiqueta)

        if hasattr(self, "reemplazo_combo"):
            self.reemplazo_combo["values"] = self._candidatos_ciclo_lista
            puede = self._puede_asignar_reemplazo()
            self._configurar_controles_reemplazo(puede)
            if not puede:
                self.reemplazo_estado_label.config(
                    text="Su rol no posee el permiso 'reemplazos.asignar' (solo lectura).",
                    foreground="#6c757d",
                )
            elif not dotacion or not especialidad:
                self.reemplazo_estado_label.config(
                    text="Seleccione un empleado titular para filtrar el personal de CICLO disponible.",
                    foreground="#6c757d",
                )
            elif not self._candidatos_ciclo_lista:
                self.reemplazo_estado_label.config(
                    text=f"No hay personal de CICLO activo para Dotación '{dotacion}' y Especialidad '{especialidad}'.",
                    foreground="#d9534f",
                )
            else:
                self.reemplazo_estado_label.config(
                    text=f"{len(self._candidatos_ciclo_lista)} candidato(s) de CICLO para {dotacion} - {especialidad}.",
                    foreground="#198754",
                )
            self._sincronizar_fechas_reemplazo_default()

    def _filtrar_combobox_reemplazo(self, event=None):
        if event is not None and event.keysym in ("Down", "Up", "Return", "Escape", "Tab", "Left", "Right", "Home", "End"):
            return
        if not hasattr(self, "reemplazo_combo"):
            return
        texto = self.reemplazo_combo.get().strip().lower()
        if not texto:
            self.reemplazo_combo["values"] = self._candidatos_ciclo_lista
        else:
            self.reemplazo_combo["values"] = [v for v in self._candidatos_ciclo_lista if texto in v.lower()]

    def _resolver_candidato_seleccionado(self, texto):
        texto = (texto or "").strip()
        if not texto:
            return None
        if texto in self._candidatos_ciclo_map:
            return self._candidatos_ciclo_map[texto]
        for k, v in self._candidatos_ciclo_map.items():
            if k.casefold() == texto.casefold():
                return v
        partes = texto.split("-", 1)
        if partes[0].strip().isdigit():
            leg_buscado = int(partes[0].strip())
            for v in self._candidatos_ciclo_map.values():
                if int(v["legajo"]) == leg_buscado:
                    return v
        return None

    def agregar_tramo_reemplazo(self):
        """Valida y agrega un tramo de reemplazo a la lista temporal del formulario."""
        if not self._puede_asignar_reemplazo():
            self.mostrar_error_novedades("No tiene permisos para asignar reemplazos.")
            return
        turno_titular = self.app.turnos_var.get().strip()
        if validators.es_turno_ciclo(turno_titular):
            self.mostrar_error_novedades("El personal cuyo turno comienza con CICLO no puede tener reemplazo.")
            return
        seleccion = self.reemplazo_combo_var.get().strip()
        cand = self._resolver_candidato_seleccionado(seleccion)
        if not cand:
            self.mostrar_error_novedades("Seleccione un personal de CICLO válido del listado de reemplazos.")
            return

        f_ini_nov = self.app.fecha_inicio_novedad_entry.entry.get().strip()
        f_fin_nov = self.app.fecha_fin_novedad_entry.entry.get().strip()
        f_desde = self.reemplazo_desde_entry.entry.get().strip() or f_ini_nov
        f_hasta = self.reemplazo_hasta_entry.entry.get().strip() or (f_fin_nov if f_fin_nov and f_fin_nov != "-" else "")

        nuevo_tramo = {
            "legajo_reemplazo": int(cand["legajo"]),
            "apellidos_nombres_reemplazo": cand["apellidos_nombres"],
            "turnos_reemplazo": cand.get("turnos") or "",
            "franco_reemplazo": cand.get("franco") or "",
            "fecha_inicio": f_desde,
            "fecha_fin": f_hasta or None,
        }
        tramos_prueba = list(self.tramos_reemplazo) + [nuevo_tramo]
        es_valido, err = validators.validar_tramos_reemplazo(
            f_ini_nov, f_fin_nov or None, turno_titular, tramos_prueba
        )
        if not es_valido:
            self.mostrar_error_novedades(err)
            return

        excluir_id = self.edicion["id"] if (self.edicion and self.edicion.get("tipo") == "novedad") else None
        if hasattr(self.app, "records_service") and self.app.records_service is not None:
            try:
                self.app.records_service.validar_disponibilidad_ciclo(
                    nuevo_tramo["legajo_reemplazo"],
                    nuevo_tramo["fecha_inicio"],
                    nuevo_tramo["fecha_fin"],
                    excluir_novedad_id=excluir_id,
                )
            except ValueError as e:
                self.mostrar_error_novedades(str(e))
                return

        self.tramos_reemplazo.append(nuevo_tramo)
        self._renderizar_tramos_reemplazo()
        self.reemplazo_combo_var.set("")
        if self.app.error_novedades_label is not None:
            self.app.error_novedades_label.config(text="")
        # Si el tramo agregado tiene fecha_fin, sugerir el día siguiente para un eventual próximo tramo
        if nuevo_tramo["fecha_fin"]:
            try:
                sig_dia = datetime.strptime(nuevo_tramo["fecha_fin"], "%d/%m/%Y") + timedelta(days=1)
                self.reemplazo_desde_entry.entry.delete(0, tk.END)
                self.reemplazo_desde_entry.entry.insert(0, sig_dia.strftime("%d/%m/%Y"))
                self.reemplazo_hasta_entry.entry.delete(0, tk.END)
                if f_fin_nov and f_fin_nov != "-":
                    self.reemplazo_hasta_entry.entry.insert(0, f_fin_nov)
            except Exception:
                pass

    def quitar_tramo_reemplazo(self):
        """Elimina el tramo de reemplazo seleccionado en la tabla del formulario."""
        if not hasattr(self, "reemplazos_tree"):
            return
        sel = self.reemplazos_tree.selection()
        if not sel:
            return
        indices = sorted([int(i) for i in sel if str(i).isdigit()], reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.tramos_reemplazo):
                self.tramos_reemplazo.pop(idx)
        self._renderizar_tramos_reemplazo()
        self._sincronizar_fechas_reemplazo_default()

    def _on_reemplazo_tree_select(self, event=None):
        """Carga en los controles los datos del tramo de reemplazo seleccionado para facilitar su edición."""
        if not hasattr(self, "reemplazos_tree"):
            return
        sel = self.reemplazos_tree.selection()
        if not sel:
            return
        try:
            idx = int(sel[0])
            if 0 <= idx < len(self.tramos_reemplazo):
                tramo = self.tramos_reemplazo[idx]
                legajo = tramo.get("legajo_reemplazo")
                nombre = tramo.get("apellidos_nombres_reemplazo")

                # Seleccionar en combobox si existe el candidato
                etiqueta = None
                for k, v in self._candidatos_ciclo_map.items():
                    if int(v.get("legajo", -1)) == int(legajo):
                        etiqueta = k
                        break
                if etiqueta:
                    self.reemplazo_combo_var.set(etiqueta)
                else:
                    self.reemplazo_combo_var.set(f"{legajo} - {nombre}")

                f_ini = tramo.get("fecha_inicio") or ""
                f_fin = tramo.get("fecha_fin") or ""
                self.reemplazo_desde_entry.entry.delete(0, tk.END)
                if f_ini:
                    self.reemplazo_desde_entry.entry.insert(0, f_ini)
                self.reemplazo_hasta_entry.entry.delete(0, tk.END)
                if f_fin and f_fin != "-":
                    self.reemplazo_hasta_entry.entry.insert(0, f_fin)
        except Exception:
            pass

    def modificar_tramo_reemplazo(self):
        """Modifica las fechas o personal del tramo de reemplazo seleccionado."""
        if not self._puede_asignar_reemplazo():
            self.mostrar_error_novedades("No tiene permisos para modificar reemplazos.")
            return
        if not hasattr(self, "reemplazos_tree"):
            return
        sel = self.reemplazos_tree.selection()
        if not sel:
            self.mostrar_error_novedades("Seleccione un tramo de la tabla de reemplazos para modificar.")
            return
        try:
            idx = int(sel[0])
            if not (0 <= idx < len(self.tramos_reemplazo)):
                return
        except Exception:
            return

        tramo_actual = self.tramos_reemplazo[idx]
        turno_titular = self.app.turnos_var.get().strip()
        if validators.es_turno_ciclo(turno_titular):
            self.mostrar_error_novedades("El personal cuyo turno comienza con CICLO no puede tener reemplazo.")
            return

        seleccion = self.reemplazo_combo_var.get().strip()
        cand = self._resolver_candidato_seleccionado(seleccion)

        f_ini_nov = self.app.fecha_inicio_novedad_entry.entry.get().strip()
        f_fin_nov = self.app.fecha_fin_novedad_entry.entry.get().strip()
        f_desde = self.reemplazo_desde_entry.entry.get().strip() or f_ini_nov
        f_hasta = self.reemplazo_hasta_entry.entry.get().strip() or (f_fin_nov if f_fin_nov and f_fin_nov != "-" else "")

        tramo_modificado = dict(tramo_actual)
        if cand:
            tramo_modificado["legajo_reemplazo"] = int(cand["legajo"])
            tramo_modificado["apellidos_nombres_reemplazo"] = cand["apellidos_nombres"]
            tramo_modificado["turnos_reemplazo"] = cand.get("turnos") or ""
            tramo_modificado["franco_reemplazo"] = cand.get("franco") or ""
        tramo_modificado["fecha_inicio"] = f_desde
        tramo_modificado["fecha_fin"] = f_hasta or None

        # Validar solapamiento con los demás tramos
        tramos_prueba = [dict(t) for i, t in enumerate(self.tramos_reemplazo) if i != idx] + [tramo_modificado]
        es_valido, err = validators.validar_tramos_reemplazo(
            f_ini_nov, f_fin_nov or None, turno_titular, tramos_prueba
        )
        if not es_valido:
            self.mostrar_error_novedades(err)
            return

        excluir_id = self.edicion["id"] if (self.edicion and self.edicion.get("tipo") == "novedad") else None
        if hasattr(self.app, "records_service") and self.app.records_service is not None:
            try:
                self.app.records_service.validar_disponibilidad_ciclo(
                    tramo_modificado["legajo_reemplazo"],
                    tramo_modificado["fecha_inicio"],
                    tramo_modificado["fecha_fin"],
                    excluir_novedad_id=excluir_id,
                )
            except ValueError as e:
                self.mostrar_error_novedades(str(e))
                return

        self.tramos_reemplazo[idx] = tramo_modificado
        self._renderizar_tramos_reemplazo()
        if self.app.error_novedades_label is not None:
            self.app.error_novedades_label.config(text="")

    def _obtener_tramos_para_guardar(self, fecha_inicio_nov, fecha_fin_nov, turno_titular, excluir_novedad_id=None):
        """Prepara y valida la lista de tramos de reemplazo antes de guardar la novedad.

        Si el usuario seleccionó un candidato en el combobox pero no pulsó '+ Agregar reemplazo',
        lo incorpora automáticamente como tramo. También ajusta tramos abiertos si la novedad
        define ahora una fecha de fin.
        """
        if validators.es_turno_ciclo(turno_titular):
            return [], None

        tramos = [dict(t) for t in self.tramos_reemplazo]
        seleccion = self.reemplazo_combo_var.get().strip() if hasattr(self, "reemplazo_combo_var") else ""
        if seleccion:
            cand = self._resolver_candidato_seleccionado(seleccion)
            if not cand:
                return None, "Seleccione un personal de CICLO válido del listado de reemplazos o limpie el campo."
            f_desde = self.reemplazo_desde_entry.entry.get().strip() if hasattr(self, "reemplazo_desde_entry") else ""
            f_hasta = self.reemplazo_hasta_entry.entry.get().strip() if hasattr(self, "reemplazo_hasta_entry") else ""
            f_desde = f_desde or (fecha_inicio_nov or "")
            f_hasta = f_hasta or (fecha_fin_nov if fecha_fin_nov and fecha_fin_nov != "-" else "")
            tramos.append({
                "legajo_reemplazo": int(cand["legajo"]),
                "apellidos_nombres_reemplazo": cand["apellidos_nombres"],
                "turnos_reemplazo": cand.get("turnos") or "",
                "franco_reemplazo": cand.get("franco") or "",
                "fecha_inicio": f_desde,
                "fecha_fin": f_hasta or None,
            })

        # Si la novedad tiene fecha_fin y algún tramo previo estaba abierto (sin fecha_fin), acotarlo a fecha_fin_nov
        if fecha_fin_nov and fecha_fin_nov != "-":
            for t in tramos:
                if not t.get("fecha_fin") or t.get("fecha_fin") == "-":
                    t["fecha_fin"] = fecha_fin_nov

        if not tramos:
            return [], None

        es_valido, err = validators.validar_tramos_reemplazo(
            fecha_inicio_nov, fecha_fin_nov, turno_titular, tramos
        )
        if not es_valido:
            return None, err

        if hasattr(self.app, "records_service") and self.app.records_service is not None:
            for t in tramos:
                try:
                    self.app.records_service.validar_disponibilidad_ciclo(
                        t["legajo_reemplazo"],
                        t["fecha_inicio"],
                        t.get("fecha_fin"),
                        excluir_novedad_id=excluir_novedad_id,
                    )
                except ValueError as e:
                    return None, str(e)

        return tramos, None
    
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
        try:
            modal.state("zoomed")
        except Exception:
            pass
        self.app.aplicar_tema_ventana(modal)
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
        
        ttk.Label(modal, text="Filtrar por apellido o nombre:").grid(row=0, column=0, padx=10, pady=5)
        apellido_filter_var = tk.StringVar()
        apellido_filter = ttk.Entry(modal, textvariable=apellido_filter_var)
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
                self._actualizar_candidatos_reemplazo(limpiar_tramos=True)
            elif boton == 2:
                self.app.legajo_2_var.set(legajo_sap)
                self.app.apellidos_nombres_2_var.set(apellido_nombre)
                self.app.especialidad_2_var.set(especialidad)
                self.app.dotacion_2_var.set(dotacion)
                self.app.turnos_2_var.set(turnos)
                self.app.franco_2_var.set(franco)
                
            modal.destroy()

        tree.bind("<Double-1>", on_double_click)
        ttk.Button(modal, text="Cerrar", command=modal.destroy).grid(row=2, column=0, columnspan=2, pady=10)

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

            row = self.app.base_index.get(legajo)
            if row:
                if campo == 1:
                    self.app.apellidos_nombres_var.set(row[1])
                    self.app.especialidad_var.set(row[2])
                    self.app.dotacion_var.set(row[3])
                    self.app.turnos_var.set(row[4])
                    self.app.franco_var.set(row[5])
                    self._actualizar_candidatos_reemplazo(limpiar_tramos=True)
                elif campo == 2:
                    self.app.apellidos_nombres_2_var.set(row[1])
                    self.app.especialidad_2_var.set(row[2])
                    self.app.dotacion_2_var.set(row[3])
                    self.app.turnos_2_var.set(row[4])
                    self.app.franco_2_var.set(row[5])
            else:
                messagebox.showinfo("Legajo No Encontrado", f"El legajo {legajo} no fue encontrado.")

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
        self.app.supervisor_var.set(self.app.current_user.get("nombre") or self.app.current_user.get("username", ""))
        self.app.observaciones_var.set('')
        if hasattr(self.app, "observaciones_novedades_text"):
            self.app.observaciones_novedades_text.delete("1.0", "end")
        if hasattr(self.app, "fecha_fin_novedad_entry"):
            self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        if self.app.error_novedades_label is not None:
            self.app.error_novedades_label.config(text="")
        self._restaurar_combobox_personal_estacion()
        self._actualizar_candidatos_reemplazo(limpiar_tramos=True)

    def limpiar_fecha_fin_novedad(self):
        """Limpia el campo de fecha de fin de novedad."""
        if hasattr(self.app, "fecha_fin_novedad_entry"):
            self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
            self.app.fecha_fin_novedad_var.set("")
        self._sincronizar_fechas_reemplazo_default()

    def _calcular_y_asignar_fecha_fin(self, origen="tipo"):
        """Autocompleta la fecha fin según la cantidad de días configurada en el tipo de novedad,
        o la limpia si se cambia a un tipo sin días definidos (dias <= 0)."""
        if self.edicion and self.edicion.get("modo") == "fecha_fin":
            return
        tipo_nombre = self.app.novedad_var.get().strip()
        info = getattr(self.app, "tipo_novedades_dict", {}).get(tipo_nombre)
        if not info:
            self._sincronizar_fechas_reemplazo_default()
            return
        dias = info.get("dias", 0)
        if dias <= 0:
            if origen == "tipo":
                self.limpiar_fecha_fin_novedad()
            else:
                self._sincronizar_fechas_reemplazo_default()
            return

        fecha_inicio_str = self.app.fecha_inicio_novedad_entry.entry.get().strip()
        if not fecha_inicio_str:
            return
        try:
            f_ini = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
            f_fin = f_ini + timedelta(days=dias - 1)
            fecha_fin_str = f_fin.strftime("%d/%m/%Y")
            self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
            self.app.fecha_fin_novedad_entry.entry.insert(0, fecha_fin_str)
            self.app.fecha_fin_novedad_var.set(fecha_fin_str)
            self._sincronizar_fechas_reemplazo_default()
        except Exception:
            pass

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
        self.app.supervisor_var.set(self.app.current_user.get("nombre") or self.app.current_user.get("username", ""))
        self.app.observaciones_var.set('') 
        if hasattr(self.app, "observaciones_cambios_text"):
            self.app.observaciones_cambios_text.delete("1.0", "end")
        if self.app.error_cambios_label is not None:
            self.app.error_cambios_label.config(text="")
        self._restaurar_combobox_personal_estacion()

    def _filtrar_combobox_personal_estacion(self, combobox, event=None):
        """Filtra los valores del combobox de personal de estación según el texto ingresado."""
        if event is not None and event.keysym in ("Down", "Up", "Return", "Escape", "Tab", "Left", "Right", "Home", "End"):
            return
        if not hasattr(self, "_personal_estacion_lista_original") or not self._personal_estacion_lista_original:
            return
        texto = combobox.get().strip().lower()
        if not texto:
            combobox["values"] = self._personal_estacion_lista_original
        else:
            combobox["values"] = [v for v in self._personal_estacion_lista_original if texto in v.lower()]

    def iniciar_edicion_novedad(self, record_id, modo="completa"):
        """Inicia el modo de edición de una novedad reutilizando el formulario de alta."""
        row = self.app.records_service.obtener_novedad(record_id)
        if not row:
            messagebox.showerror("Edición", "El registro ya no existe.", parent=self.app.root)
            return

        # Cambiar a la vista de formulario en modo edición
        if not self.app.toggle_view("form", edicion=True):
            return

        self.edicion = {"tipo": "novedad", "id": record_id, "modo": modo, "original": dict(row)}

        # Precargar datos en variables
        self.app.legajo_var.set(str(row["legajo"] or ""))
        self.app.apellidos_nombres_var.set(row["apellidos_nombres"] or "")
        self.app.especialidad_var.set(row["especialidad"] or "")
        self.app.dotacion_var.set(row["dotacion"] or "")
        self.app.turnos_var.set(row["turnos"] or "")
        self.app.franco_var.set(row["franco"] or "")

        # Tipo de novedad (asegurar que esté en la lista si es un valor histórico)
        novedad_val = row["novedad"] or ""
        if novedad_val and novedad_val not in self.app.tipo_novedades:
            self.app.novedad_entry["values"] = list(self.app.tipo_novedades) + [novedad_val]
        self.app.novedad_var.set(novedad_val)

        # Fechas
        fecha_inicio_val = row["fecha_inicio"] or ""
        self.app.fecha_inicio_novedad_entry.entry.delete(0, tk.END)
        self.app.fecha_inicio_novedad_entry.entry.insert(0, fecha_inicio_val)
        self.app.fecha_inicio_novedad_var.set(fecha_inicio_val)

        fecha_fin_val = "" if row["fecha_fin"] in (None, "-") else str(row["fecha_fin"])
        self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        if fecha_fin_val:
            self.app.fecha_fin_novedad_entry.entry.insert(0, fecha_fin_val)
        self.app.fecha_fin_novedad_var.set(fecha_fin_val)

        # Referencia estación
        ref_val = row["referencia_estacion"] or ""
        if ref_val and ref_val not in self._personal_estacion_lista_original:
            self.app.referencia_estacion_novedades_entry["values"] = list(self._personal_estacion_lista_original) + [ref_val]
        self.app.referencia_estacion_var.set(ref_val)

        # Supervisor original
        self.app.supervisor_var.set(row["supervisor"] or "")

        # Observaciones
        if hasattr(self.app, "observaciones_novedades_text"):
            self.app.observaciones_novedades_text.config(state="normal")
            self.app.observaciones_novedades_text.delete("1.0", "end")
            self.app.observaciones_novedades_text.insert("1.0", row["observaciones"] or "")

        if self.app.error_novedades_label is not None:
            self.app.error_novedades_label.config(text="")

        # Cargar candidatos de ciclo y tramos de reemplazo existentes
        self._actualizar_candidatos_reemplazo(limpiar_tramos=True)
        try:
            self.tramos_reemplazo = self.app.records_service.listar_reemplazos_por_novedad(record_id)
        except Exception:
            self.tramos_reemplazo = []
        self._renderizar_tramos_reemplazo()

        if modo == "fecha_fin":
            falta_fecha_fin = not fecha_fin_val
            puede_editar_nov = self.app.tiene_permiso("novedades.editar") or self.app.tiene_permiso("registros.modificar")
            puede_reemp = self._puede_asignar_reemplazo() and not validators.es_turno_ciclo(row["turnos"])
            habilitar_fecha_fin = falta_fecha_fin and puede_editar_nov

            if hasattr(self, "titulo_novedades_label"):
                if habilitar_fecha_fin and puede_reemp:
                    self.titulo_novedades_label.config(text=f"Completar Fecha Fin / Reemplazos - Novedad #{record_id}")
                elif puede_reemp and not habilitar_fecha_fin:
                    self.titulo_novedades_label.config(text=f"Gestionar Reemplazos - Novedad #{record_id}")
                else:
                    self.titulo_novedades_label.config(text=f"Completar Fecha Fin - Novedad #{record_id}")
            if hasattr(self, "aviso_edicion_novedades_label"):
                self.aviso_edicion_novedades_label.config(
                    text="Fuera del tiempo de edición: solo pueden actualizarse Fecha de Fin, observaciones adicionales y/o reemplazos."
                )
            if hasattr(self, "guardar_novedades_btn"):
                self.guardar_novedades_btn.config(text="Guardar actualización")

            # Deshabilitar edición de otros campos
            self.app.legajo_entry.config(state="disabled")
            if hasattr(self, "buscar_personal_novedades_btn"):
                self.buscar_personal_novedades_btn.config(state="disabled")
            self.app.novedad_entry.config(state="disabled")
            self.app.fecha_inicio_novedad_entry.entry.config(state="disabled")
            if hasattr(self.app.fecha_inicio_novedad_entry, "button"):
                self.app.fecha_inicio_novedad_entry.button.config(state="disabled")
            self.app.referencia_estacion_novedades_entry.config(state="disabled")
            self.app.observaciones_novedades_text.config(state="disabled")

            # Habilitar o deshabilitar fecha fin según corresponda
            estado_ff = "normal" if habilitar_fecha_fin else "disabled"
            self.app.fecha_fin_novedad_entry.entry.config(state=estado_ff)
            if hasattr(self.app.fecha_fin_novedad_entry, "button"):
                self.app.fecha_fin_novedad_entry.button.config(state=estado_ff)
            if hasattr(self, "limpiar_fecha_fin_novedades_btn"):
                self.limpiar_fecha_fin_novedades_btn.config(state=estado_ff)

            self._configurar_controles_reemplazo(puede_reemp)

            if hasattr(self, "label_nueva_obs") and hasattr(self, "nueva_observacion_text"):
                if habilitar_fecha_fin:
                    self.label_nueva_obs.grid(row=12, column=0, sticky="w")
                    self.nueva_observacion_text.grid(row=13, column=0, columnspan=7, pady=5, sticky="w")
                    self.nueva_observacion_text.delete("1.0", "end")
                else:
                    self.label_nueva_obs.grid_remove()
                    self.nueva_observacion_text.grid_remove()
        else:
            if hasattr(self, "titulo_novedades_label"):
                self.titulo_novedades_label.config(text=f"Editar Novedad #{record_id}")
            if hasattr(self, "aviso_edicion_novedades_label"):
                self.aviso_edicion_novedades_label.config(text="")
            if hasattr(self, "guardar_novedades_btn"):
                self.guardar_novedades_btn.config(text="Guardar cambios")

            # Habilitar campos editables en alta
            self.app.legajo_entry.config(state="normal")
            if hasattr(self, "buscar_personal_novedades_btn"):
                self.buscar_personal_novedades_btn.config(state="normal")
            self.app.novedad_entry.config(state="readonly")
            self.app.fecha_inicio_novedad_entry.entry.config(state="normal")
            if hasattr(self.app.fecha_inicio_novedad_entry, "button"):
                self.app.fecha_inicio_novedad_entry.button.config(state="normal")
            self.app.fecha_fin_novedad_entry.entry.config(state="normal")
            if hasattr(self.app.fecha_fin_novedad_entry, "button"):
                self.app.fecha_fin_novedad_entry.button.config(state="normal")
            if hasattr(self, "limpiar_fecha_fin_novedades_btn"):
                self.limpiar_fecha_fin_novedades_btn.config(state="normal")
            self.app.referencia_estacion_novedades_entry.config(state="normal")
            self.app.observaciones_novedades_text.config(state="normal")

            if hasattr(self, "label_nueva_obs") and hasattr(self, "nueva_observacion_text"):
                self.label_nueva_obs.grid_remove()
                self.nueva_observacion_text.grid_remove()

    def iniciar_edicion_cambio(self, record_id):
        """Inicia el modo de edición de un cambio de turno reutilizando el formulario de alta."""
        row = self.app.records_service.obtener_cambio(record_id)
        if not row:
            messagebox.showerror("Edición", "El registro ya no existe.", parent=self.app.root)
            return

        if not self.app.toggle_view("form_cambios", edicion=True):
            return

        self.edicion = {"tipo": "cambio", "id": record_id, "modo": "completa", "original": dict(row)}

        # Precargar datos Persona 1
        self.app.legajo_var.set(str(row["legajo_1"] or ""))
        self.app.apellidos_nombres_var.set(row["apellidos_nombres_1"] or "")
        self.app.especialidad_var.set(row["especialidad_1"] or "")
        self.app.dotacion_var.set(row["dotacion_1"] or "")
        self.app.turnos_var.set(row["turnos_1"] or "")
        self.app.franco_var.set(row["franco_1"] or "")

        # Precargar datos Persona 2
        self.app.legajo_2_var.set(str(row["legajo_2"] or ""))
        self.app.apellidos_nombres_2_var.set(row["apellidos_nombres_2"] or "")
        self.app.especialidad_2_var.set(row["especialidad_2"] or "")
        self.app.dotacion_2_var.set(row["dotacion_2"] or "")
        self.app.turnos_2_var.set(row["turnos_2"] or "")
        self.app.franco_2_var.set(row["franco_2"] or "")

        # Fecha cambio de turno
        fecha_cambio_val = row["fecha_cambio"] or ""
        self.app.fecha_cambio_turno_entry.entry.delete(0, tk.END)
        self.app.fecha_cambio_turno_entry.entry.insert(0, fecha_cambio_val)
        self.app.fecha_cambio_turno_var.set(fecha_cambio_val)

        # Referencia estación
        ref_val = row["referencia_estacion"] or ""
        if ref_val and ref_val not in self._personal_estacion_lista_original:
            self.app.referencia_estacion_cambios_entry["values"] = list(self._personal_estacion_lista_original) + [ref_val]
        self.app.referencia_estacion_var.set(ref_val)

        # Supervisor original
        self.app.supervisor_var.set(row["supervisor"] or "")

        # Observaciones
        if hasattr(self.app, "observaciones_cambios_text"):
            self.app.observaciones_cambios_text.delete("1.0", "end")
            self.app.observaciones_cambios_text.insert("1.0", row["observaciones"] or "")

        if self.app.error_cambios_label is not None:
            self.app.error_cambios_label.config(text="")

        if hasattr(self, "titulo_cambios_label"):
            self.titulo_cambios_label.config(text=f"Editar Cambio de Turno #{record_id}")
        if hasattr(self, "guardar_cambios_btn"):
            self.guardar_cambios_btn.config(text="Guardar cambios")

    def salir_edicion(self):
        """Restaura los formularios al modo de alta estándar."""
        self.edicion = None
        if hasattr(self, "titulo_novedades_label"):
            self.titulo_novedades_label.config(text="Formulario de novedades")
        if hasattr(self, "aviso_edicion_novedades_label"):
            self.aviso_edicion_novedades_label.config(text="")
        if hasattr(self, "guardar_novedades_btn"):
            self.guardar_novedades_btn.config(text="Guardar Novedad")
        if hasattr(self, "label_nueva_obs"):
            self.label_nueva_obs.grid_remove()
        if hasattr(self, "nueva_observacion_text"):
            self.nueva_observacion_text.grid_remove()
            self.nueva_observacion_text.delete("1.0", "end")

        if hasattr(self.app, "legajo_entry"):
            self.app.legajo_entry.config(state="normal")
        if hasattr(self, "buscar_personal_novedades_btn"):
            self.buscar_personal_novedades_btn.config(state="normal")
        if hasattr(self.app, "novedad_entry"):
            self.app.novedad_entry.config(state="readonly")
            self.app.novedad_entry["values"] = self.app.tipo_novedades
        if hasattr(self.app, "fecha_inicio_novedad_entry"):
            self.app.fecha_inicio_novedad_entry.entry.config(state="normal")
            if hasattr(self.app.fecha_inicio_novedad_entry, "button"):
                self.app.fecha_inicio_novedad_entry.button.config(state="normal")
        if hasattr(self.app, "fecha_fin_novedad_entry"):
            self.app.fecha_fin_novedad_entry.entry.config(state="normal")
            if hasattr(self.app.fecha_fin_novedad_entry, "button"):
                self.app.fecha_fin_novedad_entry.button.config(state="normal")
        if hasattr(self, "limpiar_fecha_fin_novedades_btn"):
            self.limpiar_fecha_fin_novedades_btn.config(state="normal")
        if hasattr(self.app, "referencia_estacion_novedades_entry"):
            self.app.referencia_estacion_novedades_entry.config(state="normal")
        if hasattr(self.app, "observaciones_novedades_text"):
            self.app.observaciones_novedades_text.config(state="normal")

        if hasattr(self, "titulo_cambios_label"):
            self.titulo_cambios_label.config(text="Formulario de cambios de turnos")
        if hasattr(self, "guardar_cambios_btn"):
            self.guardar_cambios_btn.config(text="Guardar cambio de turno")

    def _restaurar_combobox_personal_estacion(self):
        """Restaura la lista completa en los combobox de personal de estación."""
        self._personal_estacion_lista_original = self.app.nombres_personal_estacion()
        for attr in ("referencia_estacion_novedades_entry", "referencia_estacion_cambios_entry"):
            combobox = getattr(self.app, attr, None)
            if combobox:
                combobox["values"] = self._personal_estacion_lista_original

    def guardar_datos_novedades(self):
        # Modo edición restringida (fecha_fin y/o reemplazos fuera de ventana de edición)
        if self.edicion and self.edicion.get("tipo") == "novedad" and self.edicion.get("modo") == "fecha_fin":
            orig = self.edicion.get("original", {})
            orig_ff = "" if orig.get("fecha_fin") in (None, "-") else str(orig.get("fecha_fin")).strip()
            fecha_fin = self.app.fecha_fin_novedad_entry.entry.get().strip()
            nueva_obs = self.nueva_observacion_text.get("1.0", "end-1c").strip() if hasattr(self, "nueva_observacion_text") else ""
            puede_editar_nov = self.app.tiene_permiso("novedades.editar") or self.app.tiene_permiso("registros.modificar")
            puede_reemp = self._puede_asignar_reemplazo() and not validators.es_turno_ciclo(orig.get("turnos"))
            habilitar_ff = (not orig_ff) and puede_editar_nov

            if not habilitar_ff and not puede_reemp:
                self.mostrar_error_novedades("No tiene permisos para actualizar este registro.")
                return
            if habilitar_ff and not fecha_fin and not nueva_obs and not puede_reemp:
                self.mostrar_error_novedades("Debe seleccionar una fecha de fin o ingresar una observación.")
                return

            fecha_fin_efectiva = fecha_fin or orig_ff or None
            tramos = None
            if puede_reemp:
                tramos, err_tramos = self._obtener_tramos_para_guardar(
                    orig.get("fecha_inicio"),
                    fecha_fin_efectiva,
                    orig.get("turnos"),
                    excluir_novedad_id=self.edicion["id"],
                )
                if err_tramos:
                    self.mostrar_error_novedades(err_tramos)
                    return

            try:
                if habilitar_ff and (fecha_fin or nueva_obs):
                    self.app.records_service.actualizar_fecha_fin_novedad(
                        self.edicion["id"], fecha_fin, nueva_observacion=nueva_obs,
                        user_id=self.app.current_user.get("id"),
                        windows_user=self.app.obtener_usuario_windows(),
                    )
                if puede_reemp and tramos is not None:
                    self.app.records_service.guardar_reemplazos_novedad(
                        self.edicion["id"],
                        tramos,
                        user_id=self.app.current_user.get("id"),
                        windows_user=self.app.obtener_usuario_windows(),
                    )
                messagebox.showinfo("Guardado", "La novedad ha sido actualizada correctamente.")
                self.salir_edicion()
                self.limpiar_formulario_novedades()
                self.app.toggle_view("table")
                return
            except Exception as error:
                self.mostrar_error_novedades(str(error))
                return

        # Modo alta o edición completa
        self.app.fecha_inicio_novedad_var.set(self.app.fecha_inicio_novedad_entry.entry.get())
        self.app.fecha_fin_novedad_var.set(self.app.fecha_fin_novedad_entry.entry.get())
        self.app.observaciones_var.set(self.app.observaciones_novedades_text.get("1.0", "end-1c"))
        es_valido, mensaje_error = validators.validar_campos_requeridos_novedades(
            self.app.legajo_var, self.app.apellidos_nombres_var, self.app.novedad_var,
            self.app.referencia_estacion_var, self.app.supervisor_var,
            self.app.fecha_inicio_novedad_entry, self.app.fecha_fin_novedad_entry,
            self.app.tipo_novedades if not self.edicion else self.app.novedad_entry["values"]
        )
        if not es_valido:
            self.mostrar_error_novedades(mensaje_error)
            return

        # Validar que la referencia estación sea válida (permite mantener el valor original si se está editando)
        ref = self.app.referencia_estacion_var.get().strip()
        orig_ref = (self.edicion.get("original", {}).get("referencia_estacion") or "").strip() if self.edicion else ""
        if ref and ref != orig_ref:
            if hasattr(self, "_personal_estacion_lista_original") and self._personal_estacion_lista_original:
                if not any(ref.casefold() == v.casefold() for v in self._personal_estacion_lista_original):
                    self.mostrar_error_novedades("La estación de referencia debe ser un valor de la lista.")
                    if hasattr(self.app, "referencia_estacion_novedades_entry"):
                        self.app.referencia_estacion_novedades_entry.focus_set()
                    return

        excluir_id = self.edicion["id"] if (self.edicion and self.edicion.get("tipo") == "novedad") else None
        tramos_guardar, err_tramos = self._obtener_tramos_para_guardar(
            self.app.fecha_inicio_novedad_var.get().strip(),
            self.app.fecha_fin_novedad_var.get().strip() or None,
            self.app.turnos_var.get().strip(),
            excluir_novedad_id=excluir_id,
        )
        if err_tramos:
            self.mostrar_error_novedades(err_tramos)
            return

        # Si estamos en edición completa de novedad
        if self.edicion and self.edicion.get("tipo") == "novedad":
            data = {
                "legajo": int(self.app.legajo_var.get().strip()),
                "apellidos_nombres": self.app.apellidos_nombres_var.get(),
                "especialidad": self.app.especialidad_var.get(),
                "dotacion": self.app.dotacion_var.get(),
                "turnos": self.app.turnos_var.get(),
                "franco": self.app.franco_var.get(),
                "novedad": self.app.novedad_var.get(),
                "fecha_inicio": self.app.fecha_inicio_novedad_var.get(),
                "fecha_fin": self.app.fecha_fin_novedad_var.get() or None,
                "referencia_estacion": self.app.referencia_estacion_var.get(),
                "supervisor": self.app.supervisor_var.get(),
                "observaciones": self.app.observaciones_var.get(),
            }
            try:
                bypass = self.app.tiene_permiso("registros.modificar")
                self.app.records_service.actualizar_novedad(
                    self.edicion["id"], data,
                    user_id=self.app.current_user.get("id"),
                    windows_user=self.app.obtener_usuario_windows(),
                    bypass_window=bypass,
                )
                if self._puede_asignar_reemplazo() or validators.es_turno_ciclo(self.app.turnos_var.get()):
                    self.app.records_service.guardar_reemplazos_novedad(
                        self.edicion["id"],
                        tramos_guardar,
                        user_id=self.app.current_user.get("id"),
                        windows_user=self.app.obtener_usuario_windows(),
                    )
                messagebox.showinfo("Guardado", "Novedad actualizada correctamente.")
                self.salir_edicion()
                self.limpiar_formulario_novedades()
                self.app.toggle_view("table")
                return
            except Exception as error:
                self.mostrar_error_novedades(str(error))
                return

        # Alta estándar de novedad
        try:
            record_id = self.app.db_store.insert_novedad((
                None, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), int(self.app.legajo_var.get().strip()),
                self.app.apellidos_nombres_var.get(), self.app.especialidad_var.get(),
                self.app.dotacion_var.get(), self.app.turnos_var.get(), self.app.franco_var.get(),
                self.app.novedad_var.get(), self.app.fecha_inicio_novedad_var.get(),
                self.app.fecha_fin_novedad_var.get() or None, self.app.referencia_estacion_var.get(),
                self.app.supervisor_var.get(), self.app.observaciones_var.get(),
                self.app.obtener_usuario_windows(), self.app.current_user.get("id"),
            ))
            if tramos_guardar and self._puede_asignar_reemplazo():
                self.app.records_service.guardar_reemplazos_novedad(
                    record_id,
                    tramos_guardar,
                    user_id=self.app.current_user.get("id"),
                    windows_user=self.app.obtener_usuario_windows(),
                )
            self.app.records_service.registrar_auditoria(
                "crear", "novedad", record_id, self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                after={"novedad": self.app.novedad_var.get(), "legajo": self.app.legajo_var.get()},
            )
            if self.app.novedad_var.get().strip().casefold() == "informe":
                try:
                    self.app.enviar_informe_novedad(record_id)
                except Exception as error:
                    self.app.records_service.registrar_auditoria(
                        "error_envio", "informe_email", record_id, self.app.current_user.get("id"),
                        self.app.obtener_usuario_windows(), after={"error": str(error)},
                    )
                    messagebox.showwarning("Informe", f"La novedad se guardó, pero no se pudo enviar el correo:\n{error}")
            messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")
            self.limpiar_formulario_novedades()
            self.app.toggle_view()
        except DatabaseBusyError as error:
            messagebox.showerror("Base ocupada", str(error))
        except Exception as error:
            messagebox.showerror("Error", f"No se pudieron guardar los datos: {error}")
        return

    def guardar_datos_cambios(self):
        self.app.fecha_cambio_turno_var.set(self.app.fecha_cambio_turno_entry.entry.get())
        self.app.observaciones_var.set(self.app.observaciones_cambios_text.get("1.0", "end-1c"))
        es_valido, mensaje_error = validators.validar_campos_requeridos_cambios(
            self.app.legajo_var, self.app.apellidos_nombres_var, self.app.legajo_2_var,
            self.app.apellidos_nombres_2_var, self.app.fecha_cambio_turno_entry,
            self.app.referencia_estacion_var, self.app.supervisor_var
        )
        if not es_valido:
            self.mostrar_error_cambios(mensaje_error)
            return

        # Validar que la referencia estación sea válida (permite mantener el valor original si se está editando)
        ref = self.app.referencia_estacion_var.get().strip()
        orig_ref = (self.edicion.get("original", {}).get("referencia_estacion") or "").strip() if self.edicion else ""
        if ref and ref != orig_ref:
            if hasattr(self, "_personal_estacion_lista_original") and self._personal_estacion_lista_original:
                if not any(ref.casefold() == v.casefold() for v in self._personal_estacion_lista_original):
                    self.mostrar_error_cambios("La estación de referencia debe ser un valor de la lista.")
                    if hasattr(self.app, "referencia_estacion_cambios_entry"):
                        self.app.referencia_estacion_cambios_entry.focus_set()
                    return

        # Si estamos en edición de cambio de turno
        if self.edicion and self.edicion.get("tipo") == "cambio":
            data = {
                "legajo_1": int(self.app.legajo_var.get().strip()),
                "apellidos_nombres_1": self.app.apellidos_nombres_var.get(),
                "especialidad_1": self.app.especialidad_var.get(),
                "dotacion_1": self.app.dotacion_var.get(),
                "turnos_1": self.app.turnos_var.get(),
                "franco_1": self.app.franco_var.get(),
                "legajo_2": int(self.app.legajo_2_var.get().strip()),
                "apellidos_nombres_2": self.app.apellidos_nombres_2_var.get(),
                "especialidad_2": self.app.especialidad_2_var.get(),
                "dotacion_2": self.app.dotacion_2_var.get(),
                "turnos_2": self.app.turnos_2_var.get(),
                "franco_2": self.app.franco_2_var.get(),
                "fecha_cambio": self.app.fecha_cambio_turno_var.get(),
                "referencia_estacion": self.app.referencia_estacion_var.get(),
                "supervisor": self.app.supervisor_var.get(),
                "observaciones": self.app.observaciones_var.get(),
            }
            try:
                bypass = self.app.tiene_permiso("registros.modificar")
                self.app.records_service.actualizar_cambio(
                    self.edicion["id"], data,
                    user_id=self.app.current_user.get("id"),
                    windows_user=self.app.obtener_usuario_windows(),
                    bypass_window=bypass,
                )
                messagebox.showinfo("Guardado", "Cambio de turno actualizado correctamente.")
                self.salir_edicion()
                self.limpiar_formulario_cambios()
                self.app.toggle_view("table_cambios")
                return
            except Exception as error:
                self.mostrar_error_cambios(str(error))
                return

        # Alta estándar de cambio de turno
        try:
            record_id = self.app.db_store.insert_cambio_turno((
                None, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), int(self.app.legajo_var.get().strip()),
                self.app.apellidos_nombres_var.get(), self.app.especialidad_var.get(),
                self.app.dotacion_var.get(), self.app.turnos_var.get(), self.app.franco_var.get(),
                int(self.app.legajo_2_var.get().strip()), self.app.apellidos_nombres_2_var.get(),
                self.app.especialidad_2_var.get(), self.app.dotacion_2_var.get(),
                self.app.turnos_2_var.get(), self.app.franco_2_var.get(),
                self.app.fecha_cambio_turno_var.get(), self.app.referencia_estacion_var.get(),
                self.app.supervisor_var.get(), self.app.observaciones_var.get(),
                self.app.obtener_usuario_windows(), self.app.current_user.get("id"),
            ))
            self.app.records_service.registrar_auditoria(
                "crear", "cambio_turno", record_id, self.app.current_user.get("id"), self.app.obtener_usuario_windows(),
                after={"legajo_1": self.app.legajo_var.get(), "legajo_2": self.app.legajo_2_var.get()},
            )
            messagebox.showinfo("Guardado", "Los datos han sido guardados correctamente.")
            self.limpiar_formulario_cambios()
            self.app.toggle_view("table_cambios")
        except DatabaseBusyError as error:
            messagebox.showerror("Base ocupada", str(error))
        except Exception as error:
            messagebox.showerror("Error", f"No se pudieron guardar los datos: {error}")
        return

    def mostrar_error_novedades(self, mensaje):
        """Muestra un mensaje de error en el formulario de novedades."""
        if self.app.error_novedades_label is None:
            self.app.error_novedades_label = ttk.Label(self.app.form_frame, foreground="red")
            self.app.error_novedades_label.grid(row=14, column=0, columnspan=7, pady=5, sticky="w")
        self.app.error_novedades_label.config(text=mensaje)
    
    def mostrar_error_cambios(self, mensaje):
        """Muestra un mensaje de error en el formulario de cambios."""
        if self.app.error_cambios_label is None:
            self.app.error_cambios_label = ttk.Label(self.app.form_cambios_frame, foreground="red")
            self.app.error_cambios_label.grid(row=16, column=0, columnspan=7, pady=5, sticky="w")
        self.app.error_cambios_label.config(text=mensaje)

    def mostrar_formulario_novedades(self):
        """Crea la interfaz del formulario de novedades."""
        self.titulo_novedades_label = ttk.Label(self.app.form_frame, text="Formulario de novedades", font=("Helvetica", 20, "bold"))
        self.titulo_novedades_label.grid(
            row=0, column=0, columnspan=3, pady=10, padx=10, sticky="nw"
        )
        self.aviso_edicion_novedades_label = ttk.Label(self.app.form_frame, text="", foreground="#d9534f", font=("Helvetica", 10, "italic"))
        self.aviso_edicion_novedades_label.grid(row=0, column=3, columnspan=4, pady=10, padx=10, sticky="w")
        
        # Legajo
        ttk.Label(self.app.form_frame, text="   Legajo").grid(row=1, column=0, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.app.legajo_entry = ttk.Entry(self.app.form_frame, textvariable=self.app.legajo_var, width=10)
        self.app.legajo_entry.grid(row=2, column=0, sticky="ew", pady=5)
        self.app.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        self.buscar_personal_novedades_btn = ttk.Button(self.app.form_frame, text="Buscar Personal", command=self.mostrar_modal)
        self.buscar_personal_novedades_btn.grid(row=2, column=1, pady=10, padx=10)

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
        self.app.novedad_entry.bind("<<ComboboxSelected>>", lambda e: self._calcular_y_asignar_fecha_fin(origen="tipo"))

        # Fecha Inicio
        ttk.Label(self.app.form_frame, text="   Fecha de Inicio Novedad").grid(row=5, column=3, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=5, column=3, sticky="w")
        self.app.fecha_inicio_novedad_entry = DateEntry(
            self.app.form_frame, dateformat='%d/%m/%Y', bootstyle="danger"
        )
        self.app.fecha_inicio_novedad_entry.grid(row=6, column=3, sticky="w", pady=5)
        self.app.fecha_inicio_novedad_entry.entry.bind("<Key>", lambda event: "break")
        self.app.fecha_inicio_novedad_entry.entry.bind("<FocusOut>", lambda e: self._calcular_y_asignar_fecha_fin(origen="fecha_inicio"))
        try:
            self.app.fecha_inicio_novedad_entry.bind("<<DateEntrySelected>>", lambda e: self._calcular_y_asignar_fecha_fin(origen="fecha_inicio"))
        except Exception:
            pass

        # Fecha Fin
        ttk.Label(self.app.form_frame, text="Fecha de Fin Novedad").grid(row=5, column=4, sticky="w")
        self.app.fecha_fin_novedad_entry = DateEntry(
            self.app.form_frame, dateformat='%d/%m/%Y', bootstyle="success"
        )
        self.app.fecha_fin_novedad_entry.grid(row=6, column=4, sticky="w", pady=5)
        self.app.fecha_fin_novedad_entry.entry.bind("<Key>", lambda event: "break")
        self.app.fecha_fin_novedad_entry.entry.bind("<FocusOut>", lambda e: self._sincronizar_fechas_reemplazo_default())
        try:
            self.app.fecha_fin_novedad_entry.bind("<<DateEntrySelected>>", lambda e: self._sincronizar_fechas_reemplazo_default())
        except Exception:
            pass
        self.app.fecha_fin_novedad_entry.entry.delete(0, tk.END)
        self.app.fecha_fin_novedad_var.set("")
        self.limpiar_fecha_fin_novedades_btn = ttk.Button(self.app.form_frame, text="Limpiar", command=self.limpiar_fecha_fin_novedad)
        self.limpiar_fecha_fin_novedades_btn.grid(row=6, column=5, sticky="w", padx=5)
        
        # Referencia Estación
        ttk.Label(self.app.form_frame, text="   REFERENCIA ESTACIÓN").grid(row=7, column=0, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=0, sticky="w")
        self._personal_estacion_lista_original = self.app.nombres_personal_estacion()
        self.app.referencia_estacion_novedades_entry = ttk.Combobox(
            self.app.form_frame, textvariable=self.app.referencia_estacion_var,
            values=self._personal_estacion_lista_original, width=40,
            state="normal",
        )
        self.app.referencia_estacion_novedades_entry.bind(
            "<KeyRelease>", lambda e: self._filtrar_combobox_personal_estacion(self.app.referencia_estacion_novedades_entry, e)
        )
        self.app.referencia_estacion_entry = self.app.referencia_estacion_novedades_entry
        self.app.referencia_estacion_entry.grid(row=8, column=0, columnspan=2, sticky="w", pady=5)

        # Supervisor
        ttk.Label(self.app.form_frame, text="   SUPERVISOR").grid(row=7, column=3, sticky="w")
        ttk.Label(self.app.form_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=7, column=3, sticky="w")
        self.app.supervisor_entry = ttk.Entry(
            self.app.form_frame, textvariable=self.app.supervisor_var, width=40, state="readonly", style="Readonly.TEntry"
        )
        self.app.supervisor_entry.grid(row=8, column=3, columnspan=2, sticky="w", pady=5)

        # Reemplazos / Relevos (Personal de CICLO)
        self.reemplazos_frame = ttk.LabelFrame(
            self.app.form_frame, text="Reemplazos / Relevos (Personal de CICLO)", padding=8
        )
        self.reemplazos_frame.grid(row=9, column=0, columnspan=7, sticky="ew", padx=5, pady=6)

        self.reemplazo_estado_label = ttk.Label(
            self.reemplazos_frame,
            text="Seleccione un empleado titular para filtrar el personal de CICLO disponible.",
            foreground="#6c757d",
        )
        self.reemplazo_estado_label.grid(row=0, column=0, columnspan=9, sticky="w", padx=4, pady=(0, 4))

        ttk.Label(self.reemplazos_frame, text="Personal CICLO:").grid(row=1, column=0, sticky="w", padx=4)
        self.reemplazo_combo_var = tk.StringVar()
        self.reemplazo_combo = ttk.Combobox(
            self.reemplazos_frame, textvariable=self.reemplazo_combo_var,
            values=[], width=46, state="normal",
        )
        self.reemplazo_combo.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        self.reemplazo_combo.bind("<KeyRelease>", self._filtrar_combobox_reemplazo)

        ttk.Label(self.reemplazos_frame, text="Desde:").grid(row=1, column=2, sticky="w", padx=(8, 2))
        self.reemplazo_desde_entry = DateEntry(
            self.reemplazos_frame, dateformat='%d/%m/%Y', bootstyle="info", width=12
        )
        self.reemplazo_desde_entry.grid(row=1, column=3, sticky="w", padx=2, pady=2)
        self.reemplazo_desde_entry.entry.bind("<Key>", lambda event: "break")

        ttk.Label(self.reemplazos_frame, text="Hasta:").grid(row=1, column=4, sticky="w", padx=(8, 2))
        self.reemplazo_hasta_entry = DateEntry(
            self.reemplazos_frame, dateformat='%d/%m/%Y', bootstyle="info", width=12
        )
        self.reemplazo_hasta_entry.grid(row=1, column=5, sticky="w", padx=2, pady=2)
        self.reemplazo_hasta_entry.entry.bind("<Key>", lambda event: "break")
        self.reemplazo_hasta_entry.entry.delete(0, tk.END)

        self.limpiar_reemplazo_hasta_btn = ttk.Button(
            self.reemplazos_frame, text="Limpiar Hasta", command=self._limpiar_fecha_fin_reemplazo
        )
        self.limpiar_reemplazo_hasta_btn.grid(row=1, column=6, sticky="w", padx=4, pady=2)

        self.agregar_reemplazo_btn = ttk.Button(
            self.reemplazos_frame, text="+ Agregar tramo", bootstyle="success-outline", command=self.agregar_tramo_reemplazo
        )
        self.agregar_reemplazo_btn.grid(row=1, column=7, sticky="w", padx=4, pady=2)

        self.modificar_reemplazo_btn = ttk.Button(
            self.reemplazos_frame, text="Modificar tramo", bootstyle="info-outline", command=self.modificar_tramo_reemplazo
        )
        self.modificar_reemplazo_btn.grid(row=1, column=8, sticky="w", padx=4, pady=2)

        self.quitar_reemplazo_btn = ttk.Button(
            self.reemplazos_frame, text="Quitar seleccionado", bootstyle="danger-outline", command=self.quitar_tramo_reemplazo
        )
        self.quitar_reemplazo_btn.grid(row=1, column=9, sticky="w", padx=4, pady=2)

        cols_reemp = ("legajo", "nombre", "turno", "franco", "desde", "hasta")
        self.reemplazos_tree = ttk.Treeview(
            self.reemplazos_frame, columns=cols_reemp, show="headings", height=3
        )
        for col_id, titulo, ancho in (
            ("legajo", "Legajo", 90),
            ("nombre", "Apellido y Nombre (Reemplazo)", 260),
            ("turno", "Turno", 140),
            ("franco", "Franco", 110),
            ("desde", "Desde", 110),
            ("hasta", "Hasta", 110),
        ):
            self.reemplazos_tree.heading(col_id, text=titulo)
            self.reemplazos_tree.column(col_id, width=ancho, anchor="center" if col_id != "nombre" else "w")
        self.reemplazos_tree.grid(row=2, column=0, columnspan=10, sticky="ew", padx=4, pady=4)
        self.reemplazos_tree.bind("<<TreeviewSelect>>", self._on_reemplazo_tree_select)
        self._actualizar_candidatos_reemplazo(limpiar_tramos=True)

        # Observaciones
        ttk.Label(self.app.form_frame, text="Observaciones").grid(row=10, column=0, sticky="w")
        self.app.observaciones_novedades_text = scrolledtext.ScrolledText(
            self.app.form_frame, wrap=tk.WORD, height=5, width=180
        )
        self.app.observaciones_novedades_text.grid(row=11, column=0, columnspan=7, pady=5, sticky="w")

        # Campo dinámico para agregar observaciones en modo fecha_fin
        self.label_nueva_obs = ttk.Label(self.app.form_frame, text="Agregar a observaciones:")
        self.nueva_observacion_text = scrolledtext.ScrolledText(
            self.app.form_frame, wrap=tk.WORD, height=3, width=180
        )

        # Botones
        self.guardar_novedades_btn = ttk.Button(self.app.form_frame, text="Guardar Novedad", command=self.guardar_datos_novedades)
        self.guardar_novedades_btn.grid(
            row=15, column=3, columnspan=2, pady=10
        )
        self.cerrar_novedades_btn = ttk.Button(
            self.app.form_frame, text="Cerrar",
            command=lambda: (self.salir_edicion(), self.limpiar_formulario_novedades(), self.app.toggle_view())
        )
        self.cerrar_novedades_btn.grid(row=15, column=4, columnspan=2, pady=10)

    def mostrar_formulario_cambios(self):
        """Crea la interfaz del formulario de cambios de turnos."""
        self.titulo_cambios_label = ttk.Label(self.app.form_cambios_frame, text="Formulario de cambios de turnos", font=("Helvetica", 20, "bold"))
        self.titulo_cambios_label.grid(
            row=0, column=0, columnspan=3, pady=10, padx=10, sticky="nw"
        )

        # Legajo 1
        ttk.Label(self.app.form_cambios_frame, text="   Legajo").grid(row=1, column=0, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=1, column=0, sticky="w")
        self.app.legajo_entry = ttk.Entry(self.app.form_cambios_frame, textvariable=self.app.legajo_var, width=10)
        self.app.legajo_entry.grid(row=2, column=0, sticky="w", pady=5)
        self.app.legajo_entry.bind("<Return>", lambda event: self.buscar_legajo())

        self.buscar_personal_cambios_1_btn = ttk.Button(self.app.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(1))
        self.buscar_personal_cambios_1_btn.grid(
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

        self.buscar_personal_cambios_2_btn = ttk.Button(self.app.form_cambios_frame, text="Buscar Personal", command=lambda: self.mostrar_modal(2))
        self.buscar_personal_cambios_2_btn.grid(
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
        self._personal_estacion_lista_original = self.app.nombres_personal_estacion()
        self.app.referencia_estacion_cambios_entry = ttk.Combobox(
            self.app.form_cambios_frame, textvariable=self.app.referencia_estacion_var,
            values=self._personal_estacion_lista_original, width=40,
            state="normal",
        )
        self.app.referencia_estacion_cambios_entry.bind(
            "<KeyRelease>", lambda e: self._filtrar_combobox_personal_estacion(self.app.referencia_estacion_cambios_entry, e)
        )
        self.app.referencia_estacion_entry = self.app.referencia_estacion_cambios_entry
        self.app.referencia_estacion_entry.grid(row=12, column=0, columnspan=2, sticky="w", pady=5)

        # Supervisor
        ttk.Label(self.app.form_cambios_frame, text="   SUPERVISOR").grid(row=11, column=3, sticky="w")
        ttk.Label(self.app.form_cambios_frame, text="*", foreground="red", font=('Helvetica', 12, 'bold')).grid(row=11, column=3, sticky="w")
        self.app.supervisor_entry = ttk.Entry(
            self.app.form_cambios_frame, textvariable=self.app.supervisor_var, width=40, state="readonly", style="Readonly.TEntry"
        )
        self.app.supervisor_entry.grid(row=12, column=3, columnspan=2, sticky="w", pady=5)

        # Observaciones
        ttk.Label(self.app.form_cambios_frame, text="Observaciones").grid(row=13, column=0, sticky="w")
        self.app.observaciones_cambios_text = scrolledtext.ScrolledText(
            self.app.form_cambios_frame, wrap=tk.WORD, height=4, width=170
        )
        self.app.observaciones_cambios_text.grid(row=14, column=0, columnspan=7, pady=5, sticky="w")

        # Botones
        self.guardar_cambios_btn = ttk.Button(self.app.form_cambios_frame, text="Guardar cambio de turno", command=self.guardar_datos_cambios)
        self.guardar_cambios_btn.grid(
            row=17, column=2, columnspan=2, pady=10
        )
        self.cerrar_cambios_btn = ttk.Button(
            self.app.form_cambios_frame, text="Cerrar",
            command=lambda: (self.salir_edicion(), self.limpiar_formulario_cambios(), self.app.toggle_view("table_cambios"))
        )
        self.cerrar_cambios_btn.grid(row=17, column=3, columnspan=2, pady=10)
