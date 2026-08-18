"""Panel de control (dashboard) con resumen de registros y gráficos."""

import tkinter as tk
from tkinter import ttk


class DashboardManager:
    """Gestor del panel de control.

    Arma el dashboard dentro de `app.dashboard_frame`: tarjetas de resumen,
    navegación rápida y gráficos de barras (tk.Canvas) para novedades por tipo,
    registros por dotación y tendencia diaria. Los datos se recalculan en cada
    `actualizar_dashboard()` (al entrar a la vista y en el refresh de 60s).
    """

    ANCHO_CANVAS = 340
    ALTO_CANVAS = 190

    def __init__(self, app):
        self.app = app
        self.tendencia_dias = 30
        self._creado = False

    # ---------- Construcción ----------

    def crear_dashboard(self):
        frame = self.app.dashboard_frame
        for child in frame.winfo_children():
            child.destroy()

        ttk.Label(frame, text="Panel de control", font=("Helvetica", 22, "bold")).grid(
            row=0, column=0, pady=10, padx=10, sticky="w"
        )
        self.lbl_actualizado = ttk.Label(frame, text="", font=("Helvetica", 9))
        self.lbl_actualizado.grid(row=0, column=1, pady=10, padx=10, sticky="e")

        barra = ttk.Frame(frame)
        barra.grid(row=1, column=0, columnspan=2, pady=5, padx=10, sticky="w")
        if self.app.tiene_permiso("novedades.ver"):
            ttk.Button(barra, text="Novedades", command=lambda: self.app.toggle_view("table")).pack(side="left", padx=3)
        if self.app.tiene_permiso("cambios_turno.ver"):
            ttk.Button(barra, text="Cambios de turno", command=lambda: self.app.toggle_view("table_cambios")).pack(side="left", padx=3)
        if self.app.tiene_permiso("novedades.crear"):
            ttk.Button(barra, text="Nueva novedad", command=lambda: self.app.toggle_view("form")).pack(side="left", padx=3)
        if self.app.tiene_permiso("cambios_turno.crear"):
            ttk.Button(barra, text="Nuevo cambio de turno", command=lambda: self.app.toggle_view("form_cambios")).pack(side="left", padx=3)
        if self.app.tiene_permiso("novedades.exportar") or self.app.tiene_permiso("cambios_turno.exportar"):
            ttk.Button(barra, text="Exportar a Excel", command=self.app.exportar_excel).pack(side="left", padx=3)
        ttk.Label(barra, text="Tendencia:").pack(side="right", padx=(16, 4))
        self.tendencia_dias_var = tk.StringVar(value=str(self.tendencia_dias))
        combo_periodo = ttk.Combobox(
            barra, textvariable=self.tendencia_dias_var,
            values=["7", "30", "90"], width=6, state="readonly",
        )
        combo_periodo.pack(side="right")
        combo_periodo.bind("<<ComboboxSelected>>", lambda _event: self._cambiar_tendencia())

        tarjetas = ttk.Frame(frame)
        tarjetas.grid(row=2, column=0, columnspan=2, pady=5, padx=6, sticky="ew")
        self.tarjeta_novedades_hoy = self._crear_tarjeta(tarjetas, "Novedades hoy")
        if self.app.tiene_permiso("cambios_turno.ver"):
            self.tarjeta_cambios_hoy = self._crear_tarjeta(tarjetas, "Cambios hoy")
        else:
            self.tarjeta_cambios_hoy = None
        self.tarjeta_total_mes = self._crear_tarjeta(tarjetas, "Total del mes")
        self.tarjeta_sin_revisar = self._crear_tarjeta(tarjetas, "Sin revisar")

        graficos = ttk.Frame(frame)
        graficos.grid(row=3, column=0, columnspan=2, pady=8, padx=6, sticky="nsew")

        self._crear_grafico_caja(graficos, 0, "Novedades por tipo", self.ANCHO_CANVAS, self.ALTO_CANVAS)
        self._canvas_tipo = getattr(self, "_canvas_tipo")

        self._crear_grafico_caja(graficos, 1, "Registros por dotación (top 10)", self.ANCHO_CANVAS, self.ALTO_CANVAS)
        self._canvas_dotacion = getattr(self, "_canvas_dotacion")

        caja_tendencia = ttk.Frame(graficos)
        caja_tendencia.grid(row=1, column=2, padx=6, pady=4, sticky="n")
        ttk.Label(caja_tendencia, text="Tendencia", font=("Helvetica", 12, "bold")).pack(anchor="w")
        self._canvas_tendencia = tk.Canvas(
            caja_tendencia, width=self.ANCHO_CANVAS, height=self.ALTO_CANVAS,
            background=self.app.ui_background, highlightthickness=1,
            highlightbackground=self.app.ui_foreground,
        )
        self._canvas_tendencia.pack()

        self._creado = True
        self.actualizar_dashboard()

    def _crear_grafico_caja(self, parent, columna, titulo, ancho, alto):
        caja = ttk.Frame(parent)
        caja.grid(row=1, column=columna, padx=6, pady=4, sticky="n")
        ttk.Label(caja, text=titulo, font=("Helvetica", 12, "bold")).pack(anchor="w")
        canvas = tk.Canvas(
            caja, width=ancho, height=alto,
            background=self.app.ui_background, highlightthickness=1,
            highlightbackground=self.app.ui_foreground,
        )
        canvas.pack()
        if columna == 0:
            self._canvas_tipo = canvas
        elif columna == 1:
            self._canvas_dotacion = canvas

    def _crear_tarjeta(self, parent, titulo):
        marco = tk.Frame(
            parent, background=self.app.ui_background,
            highlightbackground=self.app.ui_foreground, highlightthickness=1,
        )
        marco.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        valor = tk.Label(
            marco, text="0", font=("Helvetica", 26, "bold"),
            background=self.app.ui_background, foreground=self.app.ui_foreground,
        )
        valor.pack(pady=(10, 0))
        tk.Label(
            marco, text=titulo, font=("Helvetica", 10),
            background=self.app.ui_background, foreground=self.app.ui_foreground,
        ).pack(pady=(0, 10))
        return valor

    # ---------- Colores ----------

    def _color(self, nombre, fallback="#4682b4"):
        try:
            return str(getattr(self.app.style.colors, nombre))
        except Exception:
            return fallback

    # ---------- Datos y dibujado ----------

    def actualizar_dashboard(self):
        if not self._creado:
            return
        resumen = self.app.records_service.dashboard_resumen()
        self.tarjeta_novedades_hoy.config(text=str(resumen["novedades_hoy"]))
        if self.tarjeta_cambios_hoy is not None:
            self.tarjeta_cambios_hoy.config(text=str(resumen["cambios_hoy"]))
        self.tarjeta_total_mes.config(text=str(resumen["total_mes"]))
        self.tarjeta_sin_revisar.config(text=str(getattr(self.app, "_no_revisados", 0)))
        self._dibujar_por_tipo()
        self._dibujar_por_dotacion()
        self._dibujar_tendencia()
        from datetime import datetime
        self.lbl_actualizado.config(text=f"Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    def _cambiar_tendencia(self):
        try:
            self.tendencia_dias = int(self.tendencia_dias_var.get())
        except (TypeError, ValueError):
            self.tendencia_dias = 30
        self._dibujar_tendencia()

    def _dibujar_por_tipo(self):
        datos = self.app.records_service.dashboard_por_tipo(self.tendencia_dias)
        self._dibujar_barras_h(self._canvas_tipo, datos, self._color("primary", "#4682b4"))

    def _dibujar_por_dotacion(self):
        datos = self.app.records_service.dashboard_por_dotacion(self.tendencia_dias)
        self._dibujar_barras_h(self._canvas_dotacion, datos, self._color("success", "#3cb371"))

    def _dibujar_tendencia(self):
        datos = self.app.records_service.dashboard_tendencia(self.tendencia_dias)
        self._dibujar_barras_v(
            self._canvas_tendencia, datos,
            self._color("primary", "#4682b4"), self._color("info", "#5bc0de"),
        )

    def _dibujar_barras_h(self, canvas, datos, color):
        canvas.delete("all")
        fg = self.app.ui_foreground
        ancho, alto = self.ANCHO_CANVAS, self.ALTO_CANVAS
        if not datos:
            canvas.create_text(ancho // 2, alto // 2, text="Sin datos", fill=fg)
            return
        maximo = max(v for _label, v in datos) or 1
        filas = len(datos)
        altura_fila = alto / filas
        margen_izq = 92
        area_ancho = ancho - margen_izq - 48
        for indice, (label, valor) in enumerate(datos):
            y = indice * altura_fila + altura_fila / 2
            canvas.create_text(
                margen_izq - 6, y, text=str(label)[:14], anchor="e",
                fill=fg, font=("Calibri", 9),
            )
            largo = max(2, area_ancho * (valor / maximo))
            canvas.create_rectangle(
                margen_izq, y - altura_fila * 0.3,
                margen_izq + largo, y + altura_fila * 0.3,
                fill=color, outline=color,
            )
            canvas.create_text(
                margen_izq + largo + 4, y, text=str(valor), anchor="w",
                fill=fg, font=("Calibri", 9, "bold"),
            )

    def _dibujar_barras_v(self, canvas, datos, color_nov, color_cam):
        canvas.delete("all")
        fg = self.app.ui_foreground
        ancho, alto = self.ANCHO_CANVAS, self.ALTO_CANVAS
        if not datos:
            canvas.create_text(ancho // 2, alto // 2, text="Sin datos", fill=fg)
            return
        maximo = max((n + c) for _fecha, n, c in datos) or 1
        margen_izq, margen_inf, margen_sup = 36, 20, 28
        area_ancho = ancho - margen_izq - 6
        area_alto = alto - margen_inf - margen_sup
        cantidad = len(datos)
        paso = area_ancho / cantidad
        barra = max(1, paso * 0.6)
        etiquetas = max(1, cantidad // 7)
        for indice, (fecha, n_, c_) in enumerate(datos):
            x = margen_izq + indice * paso + (paso - barra) / 2
            y_base = margen_sup + area_alto
            alto_n = area_alto * (n_ / maximo)
            alto_c = area_alto * (c_ / maximo)
            if n_:
                canvas.create_rectangle(x, y_base - alto_n, x + barra, y_base, fill=color_nov, outline=color_nov)
            if c_:
                canvas.create_rectangle(
                    x, y_base - alto_n - alto_c, x + barra, y_base - alto_n,
                    fill=color_cam, outline=color_cam,
                )
            if indice % etiquetas == 0:
                canvas.create_text(
                    margen_izq + indice * paso + paso / 2, alto - 8,
                    text=fecha[8:10] + "/" + fecha[5:7], fill=fg, font=("Calibri", 7),
                )
        canvas.create_rectangle(6, 4, 12, 10, fill=color_nov, outline=color_nov)
        canvas.create_text(16, 7, text="Novedades", anchor="w", fill=fg, font=("Calibri", 8))
        canvas.create_rectangle(6, 18, 12, 24, fill=color_cam, outline=color_cam)
        canvas.create_text(16, 21, text="Cambios", anchor="w", fill=fg, font=("Calibri", 8))