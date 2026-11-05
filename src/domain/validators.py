"""Validaciones y reglas de negocio puras para RENO v2.0.0.

0% dependencias externas (solo Python standard library: datetime, unicodedata, re).
"""

from __future__ import annotations

import unicodedata
from datetime import date, datetime, time, timedelta
from typing import Any, Sequence

from .enums import EstadoCobertura

_DIAS_SEMANA = ("Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo")
DIAS_SEMANA = _DIAS_SEMANA
DIAS_SEMANA_NORM = {"lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"}


def normalizar_dia_semana(valor: Any) -> str:
    """Normaliza un texto eliminando acentos y mayúsculas para comparar días de la semana."""
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    return "".join(c for c in texto if not unicodedata.combining(c)).lower().strip()



def parsear_fecha(valor: Any, nombre_campo: str = "Fecha", obligatorio: bool = False) -> datetime | None:
    """Parsea una cadena en formato DD/MM/YYYY o YYYY-MM-DD a objeto datetime."""
    if isinstance(valor, datetime):
        return valor
    if isinstance(valor, date):
        return datetime.combine(valor, time.min)

    val = str(valor or "").strip()
    if not val or val in ("-", "—", "None", "null", "nan"):
        if obligatorio:
            raise ValueError(f"El campo '{nombre_campo}' es obligatorio.")
        return None

    # Vía rápida ISO: YYYY-MM-DD
    if len(val) >= 10 and val[4] == "-" and val[7] == "-":
        try:
            return datetime.fromisoformat(val[:10])
        except ValueError:
            pass

    # Vía rápida estándar: DD/MM/YYYY
    if len(val) >= 10 and val[2] == "/" and val[5] == "/":
        try:
            return datetime.strptime(val[:10], "%d/%m/%Y")
        except ValueError:
            pass

    for fmt in (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%d/%m/%y",
        "%d-%m-%y",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M",
    ):
        try:
            return datetime.strptime(val, fmt)
        except ValueError:
            continue

    if "T" in val:
        try:
            return datetime.fromisoformat(val)
        except ValueError:
            pass

    # Limpieza de caracteres espurios / no imprimibles accidentales (ej: 1¡4/08/2025 -> 14/08/2025)
    import re
    val_clean = re.sub(r"[^\d/\- :T\.]", "", val).strip()
    if val_clean and val_clean != val:
        for fmt in (
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%Y/%m/%d",
            "%d/%m/%y",
            "%d-%m-%y",
            "%d/%m/%Y %H:%M:%S",
            "%Y-%m-%d %H:%M:%S",
        ):
            try:
                return datetime.strptime(val_clean, fmt)
            except ValueError:
                continue

    raise ValueError(f"El campo '{nombre_campo}' debe tener formato DD/MM/AAAA.")


def parsear_registrado_en(valor: Any) -> datetime | None:
    """Parsea timestamp de registro ISO, DD/MM/YYYY o serial float."""
    if valor is None:
        return None
    texto = str(valor).strip()
    if not texto:
        return None
    if len(texto) >= 10 and texto[4] == "-" and texto[7] == "-":
        try:
            return datetime.fromisoformat(texto)
        except ValueError:
            pass
    elif len(texto) >= 10 and texto[2] == "/" and texto[5] == "/":
        if len(texto) == 10:
            try:
                return datetime.strptime(texto, "%d/%m/%Y")
            except ValueError:
                pass
        elif len(texto) == 16:
            try:
                return datetime.strptime(texto, "%d/%m/%Y %H:%M")
            except ValueError:
                pass
        elif len(texto) >= 19:
            try:
                return datetime.strptime(texto[:19], "%d/%m/%Y %H:%M:%S")
            except ValueError:
                pass
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    try:
        return datetime(1899, 12, 30) + timedelta(days=float(texto))
    except (ValueError, OverflowError):
        return None


def es_turno_ciclo(turno_str: Any) -> bool:
    """Indica si el turno corresponde a personal de CICLO (comienza con 'CICLO')."""
    return str(turno_str or "").strip().upper().startswith("CICLO")


def validar_horario_hm(valor: str | None, nombre_campo: str = "Horario") -> tuple[int, int] | None:
    """Valida formato HH:MM (00:00 a 23:59)."""
    val = (valor or "").strip()
    if not val:
        return None
    partes = val.split(":")
    if len(partes) != 2 or not partes[0].isdigit() or not partes[1].isdigit():
        raise ValueError(f"El campo '{nombre_campo}' debe tener formato HH:MM (ej. 06:00).")
    h, m = int(partes[0]), int(partes[1])
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError(f"El campo '{nombre_campo}' contiene horas o minutos fuera de rango (00:00-23:59).")
    return h, m


def dentro_de_ventana_horaria(registrado_en: Any, horas_limite: int, ahora: datetime | None = None) -> bool:
    """Indica si el registro está dentro de la ventana horaria permitida para edición/eliminación."""
    fecha = parsear_registrado_en(registrado_en)
    if fecha is None:
        return False
    ref = ahora or datetime.now()
    return (ref - fecha).total_seconds() <= max(1, horas_limite) * 3600


def validar_datos_novedad(
    data: dict[str, Any],
    tipo_novedades: Sequence[str] | None = None,
    personal_estacion: Sequence[str] | None = None,
) -> tuple[bool, str | None]:
    """Valida los datos requeridos para alta o edición de novedad."""
    legajo = str(data.get("legajo", "")).strip()
    if not legajo:
        return False, "El campo 'Legajo' es obligatorio."
    try:
        int(legajo)
    except ValueError:
        return False, "El legajo debe ser un número entero."

    if not str(data.get("apellidos_nombres", "")).strip():
        return False, "El campo 'Apellido y Nombre' es obligatorio."

    novedad = str(data.get("novedad", "")).strip()
    if not novedad:
        return False, "El campo 'Tipo de Novedad' es obligatorio."
    if tipo_novedades and novedad not in tipo_novedades:
        return False, "Seleccione un valor válido en 'Tipo de Novedad'."

    ref = str(data.get("referencia_estacion", "")).strip()
    if not ref:
        return False, "El campo 'Referencia Estación' es obligatorio."
    if personal_estacion and not any(ref.casefold() == p.casefold() for p in personal_estacion):
        return False, "La estación de referencia debe ser un valor de la lista."

    if not str(data.get("supervisor", "")).strip():
        return False, "El campo 'Supervisor' es obligatorio."

    try:
        f_ini = parsear_fecha(data.get("fecha_inicio"), "Fecha de inicio novedad", obligatorio=True)
        f_fin = parsear_fecha(data.get("fecha_fin"), "Fecha de fin novedad", obligatorio=False)
    except ValueError as e:
        return False, str(e)

    if f_ini and f_fin and f_fin < f_ini:
        return False, "'Fecha de Fin Novedad' no puede ser anterior a 'Fecha de Inicio Novedad'."

    return True, None


def validar_datos_edicion_novedad(
    data: dict[str, Any],
    tipo_novedades: Sequence[str] | None = None,
    personal_estacion: Sequence[str] | None = None,
) -> tuple[bool, str | None]:
    """Alias compatible para validar_datos_novedad."""
    return validar_datos_novedad(data, tipo_novedades, personal_estacion)


def validar_datos_cambio(
    data: dict[str, Any],
    personal_estacion: Sequence[str] | None = None,
) -> tuple[bool, str | None]:
    """Valida los datos requeridos para alta o edición de cambio de turno."""
    for campo_leg, label in (("legajo_1", "Legajo 1"), ("legajo_2", "Legajo 2")):
        val = str(data.get(campo_leg, "")).strip()
        if not val:
            return False, f"El campo '{label}' es obligatorio."
        try:
            int(val)
        except ValueError:
            return False, f"{label} debe ser un número entero."

    if not str(data.get("apellidos_nombres_1", "")).strip():
        return False, "El campo 'Apellido y Nombre 1' es obligatorio."
    if not str(data.get("apellidos_nombres_2", "")).strip():
        return False, "El campo 'Apellido y Nombre 2' es obligatorio."

    ref = str(data.get("referencia_estacion", "")).strip()
    if not ref:
        return False, "El campo 'Referencia Estación' es obligatorio."
    if personal_estacion and not any(ref.casefold() == p.casefold() for p in personal_estacion):
        return False, "La estación de referencia debe ser un valor de la lista."

    if not str(data.get("supervisor", "")).strip():
        return False, "El campo 'Supervisor' es obligatorio."

    try:
        parsear_fecha(data.get("fecha_cambio"), "Fecha de cambio de turno", obligatorio=True)
    except ValueError as e:
        return False, str(e)

    return True, None


def validar_datos_edicion_cambio(
    data: dict[str, Any],
    personal_estacion: Sequence[str] | None = None,
) -> tuple[bool, str | None]:
    """Alias compatible para validar_datos_cambio."""
    return validar_datos_cambio(data, personal_estacion)


def validar_tramos_reemplazo(
    fecha_inicio_nov: str,
    fecha_fin_nov: str | None,
    turno_titular: str,
    tramos: Sequence[dict[str, Any]],
) -> tuple[bool, str | None]:
    """Valida la coherencia y no solapamiento interno de tramos de reemplazo."""
    if not tramos:
        return True, None

    if es_turno_ciclo(turno_titular):
        return False, "El personal cuyo turno comienza con CICLO no puede tener reemplazo."

    try:
        dt_ini_nov = parsear_fecha(fecha_inicio_nov, "Fecha de inicio novedad", obligatorio=True)
        dt_fin_nov = parsear_fecha(fecha_fin_nov, "Fecha de fin novedad", obligatorio=False) if fecha_fin_nov else None
    except ValueError as e:
        return False, str(e)

    assert dt_ini_nov is not None
    f_ini_nov = dt_ini_nov.date()
    f_fin_nov = dt_fin_nov.date() if dt_fin_nov else None

    intervalos: list[tuple[date, date, int, int]] = []
    for idx, tramo in enumerate(tramos, start=1):
        leg = str(tramo.get("legajo_reemplazo") or tramo.get("legajo") or "").strip()
        if not leg:
            return False, f"El reemplazo #{idx} debe tener un legajo válido."
        try:
            leg_int = int(leg)
        except ValueError:
            return False, f"El legajo del reemplazo #{idx} debe ser un número entero."

        t_turno = tramo.get("turnos_reemplazo") or tramo.get("turnos") or ""
        if t_turno and not es_turno_ciclo(t_turno):
            return False, f"El reemplazo #{idx} (Legajo {leg}) no pertenece a un turno CICLO."

        f_ini_t = tramo.get("fecha_inicio")
        f_fin_t = tramo.get("fecha_fin")
        try:
            dt_ini_t = parsear_fecha(f_ini_t, f"Fecha inicio reemplazo #{idx}", obligatorio=True)
            dt_fin_t = parsear_fecha(f_fin_t, f"Fecha fin reemplazo #{idx}", obligatorio=False) if f_fin_t else None
        except ValueError as e:
            return False, str(e)

        assert dt_ini_t is not None
        d_ini = dt_ini_t.date()
        d_fin = dt_fin_t.date() if dt_fin_t else (f_fin_nov or date.max)

        if d_fin < d_ini:
            return False, f"El tramo #{idx} tiene fecha de fin anterior a la de inicio."

        if d_ini < f_ini_nov:
            return False, f"El tramo #{idx} inicia antes de la fecha de inicio de la novedad ({dt_ini_nov.strftime('%d/%m/%Y')})."

        if f_fin_nov and d_fin > f_fin_nov:
            return False, f"El tramo #{idx} supera la fecha de fin de la novedad ({f_fin_nov.strftime('%d/%m/%Y')})."

        intervalos.append((d_ini, d_fin, leg_int, idx))

    intervalos.sort(key=lambda x: (x[0], x[1]))
    for i in range(len(intervalos) - 1):
        actual_ini, actual_fin, _, idx_a = intervalos[i]
        sig_ini, _, _, idx_b = intervalos[i + 1]
        if actual_fin >= sig_ini:
            return False, f"Existe solapamiento de fechas entre los tramos #{idx_a} y #{idx_b}."

    return True, None


def calcular_estado_cobertura_puro(
    fecha_inicio_nov: str,
    fecha_fin_nov: str | None,
    turno_titular: str,
    tramos: Sequence[dict[str, Any]],
) -> dict[str, str]:
    """Calcula el estado ('Sí', 'Parcial', 'No', 'N/A (Ciclo)'), resumen y detalle."""
    if es_turno_ciclo(turno_titular):
        return {
            "estado": EstadoCobertura.NA_CICLO.value,
            "resumen": "N/A (Ciclo)",
            "detalle": "Personal de CICLO (sin reemplazo)",
        }

    if not tramos:
        return {
            "estado": EstadoCobertura.NO.value,
            "resumen": "No",
            "detalle": "Sin reemplazo asignado",
        }

    try:
        dt_ini_nov = parsear_fecha(fecha_inicio_nov, obligatorio=False)
        f_ini_nov = dt_ini_nov.date() if dt_ini_nov else None
    except Exception:
        f_ini_nov = None

    raw_fin_nov = (fecha_fin_nov or "").strip()
    try:
        dt_fin_nov = parsear_fecha(raw_fin_nov, obligatorio=False) if raw_fin_nov and raw_fin_nov != "-" else None
        f_fin_nov = dt_fin_nov.date() if dt_fin_nov else None
    except Exception:
        f_fin_nov = None

    intervalos: list[tuple[date, date | None]] = []
    nombres_vistos: list[str] = []
    partes_detalle: list[str] = []

    for t in tramos:
        leg = t.get("legajo_reemplazo") or t.get("legajo")
        nom = t.get("apellidos_nombres_reemplazo") or t.get("apellidos_nombres") or ""
        etiqueta = f"{nom} ({leg})" if nom else f"Leg. {leg}"
        if etiqueta not in nombres_vistos:
            nombres_vistos.append(etiqueta)

        f_ini_t = t.get("fecha_inicio")
        f_fin_t = (t.get("fecha_fin") or "").strip()
        hasta_txt = f_fin_t if f_fin_t and f_fin_t != "-" else "abierto"
        partes_detalle.append(f"{leg} - {nom} [{f_ini_t} a {hasta_txt}]")

        try:
            dt_ini = parsear_fecha(f_ini_t, obligatorio=False)
            d_ini = dt_ini.date() if dt_ini else None
        except Exception:
            d_ini = None

        d_fin = None
        if f_fin_t and f_fin_t != "-":
            try:
                dt_fin = parsear_fecha(f_fin_t, obligatorio=False)
                d_fin = dt_fin.date() if dt_fin else None
            except Exception:
                d_fin = None
        elif f_fin_nov:
            d_fin = f_fin_nov

        if d_ini:
            intervalos.append((d_ini, d_fin))

    intervalos.sort(key=lambda x: (x[0], x[1] or date.max))
    cubierto_total = False
    if f_ini_nov and intervalos:
        primero_ini, _ = intervalos[0]
        if primero_ini <= f_ini_nov:
            cursor = intervalos[0][1]
            sin_huecos = True
            for sig_ini, sig_fin in intervalos[1:]:
                if cursor is None:
                    break
                if sig_ini > cursor + timedelta(days=1):
                    sin_huecos = False
                    break
                if sig_fin is None or sig_fin > cursor:
                    cursor = sig_fin
            if sin_huecos:
                if f_fin_nov is not None:
                    if cursor is None or cursor >= f_fin_nov:
                        cubierto_total = True
                else:
                    if cursor is None:
                        cubierto_total = True

    estado = EstadoCobertura.SI.value if cubierto_total else EstadoCobertura.PARCIAL.value
    resumen = f"{estado}: {', '.join(nombres_vistos)}"
    detalle = " | ".join(partes_detalle)
    return {
        "estado": estado,
        "resumen": resumen,
        "detalle": detalle,
    }


def calcular_descansos_y_horas_ciclo(
    legajo: int,
    empleado_data: dict[str, Any],
    tramos_asignados: Sequence[dict[str, Any]],
    jornadas_map: dict[str, dict[str, Any]],  # YYYY-MM-DD -> dict
    anio: int,
    mes: int,
) -> list[dict[str, Any]]:
    """Calcula la serie mensual día por día con descansos CCT y horas semanales."""
    import calendar

    _, num_dias = calendar.monthrange(anio, mes)
    jornada_hs = 7 if int(empleado_data.get("jornada_horas", 6) or 6) == 7 else 6
    descanso_entre_jornada = 18 if jornada_hs == 6 else 17
    descanso_post_franco = 42 if jornada_hs == 6 else 41

    franco_emp = str(empleado_data.get("franco", "") or "").strip()
    franco_horario = str(empleado_data.get("franco_horario", "") or "").strip()
    franco_norm = unicodedata.normalize("NFKD", franco_emp)
    franco_norm = "".join(c for c in franco_norm if not unicodedata.combining(c)).lower().strip()

    tramos_parsed: list[tuple[date, date | None, dict[str, Any]]] = []
    for r in tramos_asignados:
        try:
            d_ini_dt = parsear_fecha(r.get("fecha_inicio"))
            d_ini = d_ini_dt.date() if d_ini_dt else None
        except Exception:
            d_ini = None
        if not d_ini:
            continue
        d_fin = None
        raw_fin = (r.get("fecha_fin") or r.get("fecha_fin_novedad") or "").strip()
        if raw_fin and raw_fin != "-":
            try:
                d_fin_dt = parsear_fecha(raw_fin)
                d_fin = d_fin_dt.date() if d_fin_dt else None
            except Exception:
                d_fin = None
        tramos_parsed.append((d_ini, d_fin, dict(r)))

    def _parsear_hm(val_str: str | None) -> tuple[int, int] | None:
        if not val_str:
            return None
        try:
            p = val_str.strip().split(":")
            return int(p[0]), int(p[1])
        except Exception:
            return None

    def _calc_hs_trabajadas(t_str: str | None, d_str: str | None) -> float:
        hm_t = _parsear_hm(t_str)
        hm_d = _parsear_hm(d_str)
        if not hm_t or not hm_d:
            return 0.0
        min_t = hm_t[0] * 60 + hm_t[1]
        min_d = hm_d[0] * 60 + hm_d[1]
        if min_d < min_t:
            min_d += 24 * 60
        return round((min_d - min_t) / 60.0, 2)

    def _calc_deja_dt(dia_date: date, t_str: str | None, d_str: str | None) -> datetime | None:
        hm_d = _parsear_hm(d_str)
        if not hm_d:
            return None
        hm_t = _parsear_hm(t_str)
        if hm_t:
            min_t = hm_t[0] * 60 + hm_t[1]
            min_d = hm_d[0] * 60 + hm_d[1]
            if min_d < min_t:
                return datetime.combine(dia_date + timedelta(days=1), time(hm_d[0], hm_d[1]))
        return datetime.combine(dia_date, time(hm_d[0], hm_d[1]))

    fecha_primer_dia = date(anio, mes, 1)
    fecha_inicio_calc = fecha_primer_dia - timedelta(days=21)
    fecha_fin_calc = date(anio, mes, num_dias)

    dias_historia: dict[date, dict[str, Any]] = {}
    acum_semana = 0.0
    d_step = fecha_inicio_calc

    while d_step <= fecha_fin_calc:
        f_iso = d_step.strftime("%Y-%m-%d")
        dia_semana_str = _DIAS_SEMANA[d_step.weekday()]
        dia_sem_norm = unicodedata.normalize("NFKD", dia_semana_str)
        dia_sem_norm = "".join(c for c in dia_sem_norm if not unicodedata.combining(c)).lower().strip()
        es_franco_reg = bool(franco_norm and franco_norm in dia_sem_norm)

        relevo_dia: dict[str, Any] | None = None
        for r_ini, r_fin_opt, r_data in tramos_parsed:
            if r_ini <= d_step and (r_fin_opt is None or d_step <= r_fin_opt):
                relevo_dia = r_data
                break

        jornada_db = jornadas_map.get(f_iso)
        if relevo_dia:
            estado = "Reemplazando"
            es_franco = False
            novedad_id = relevo_dia["novedad_id"]
            reemplaza_a = f"{relevo_dia.get('legajo_titular', '')} - {relevo_dia.get('apellidos_titular', '')}".strip(" -")
            turno_cubierto = relevo_dia.get("turno_titular") or "-"
            motivo = f"#{relevo_dia['novedad_id']} ({relevo_dia.get('tipo_novedad', '')})"
            periodo = f"{relevo_dia.get('fecha_inicio', '')} a {relevo_dia.get('fecha_fin') or relevo_dia.get('fecha_fin_novedad') or 'abierto'}"
            reemplazando_legajo = relevo_dia.get("legajo_titular")
            reemplazando_nombre = relevo_dia.get("apellidos_titular")
            novedad_titular = relevo_dia.get("tipo_novedad")
            tramo_inicio = relevo_dia.get("fecha_inicio")
            tramo_fin = relevo_dia.get("fecha_fin") or relevo_dia.get("fecha_fin_novedad")
            toma = jornada_db.get("toma") or "" if jornada_db else ""
            deja = jornada_db.get("deja") or "" if jornada_db else ""
        else:
            novedad_id = ""
            reemplaza_a = "-"
            turno_cubierto = "-"
            motivo = "-"
            periodo = "-"
            reemplazando_legajo = None
            reemplazando_nombre = None
            novedad_titular = None
            tramo_inicio = None
            tramo_fin = None
            if jornada_db and jornada_db.get("franco_anticipado"):
                es_franco = True
                estado = "Franco Anticipado"
                toma = ""
                deja = ""
            elif es_franco_reg:
                es_franco = True
                estado = f"Franco (desde {franco_horario})" if franco_horario else "Franco"
                toma = ""
                deja = ""
            else:
                es_franco = False
                estado = "Disponible"
                toma = jornada_db.get("toma") or "" if jornada_db else ""
                deja = jornada_db.get("deja") or "" if jornada_db else ""

        hs_trab = _calc_hs_trabajadas(toma, deja) if (toma and deja) else 0.0
        deja_dt = _calc_deja_dt(d_step, toma, deja) if deja else None

        if es_franco:
            acum_semana = 0.0
            hs_semana_str = "-"
        else:
            acum_semana += hs_trab
            hs_semana_str = f"{acum_semana:.1f} hs" if acum_semana > 0 else "-"

        if es_franco:
            hora_disp_str = "-"
        else:
            d_prev = d_step - timedelta(days=1)
            disp_dt: datetime | None = None
            if d_prev in dias_historia and not dias_historia[d_prev]["es_franco"]:
                prev_deja = dias_historia[d_prev]["deja_dt"]
                if prev_deja:
                    disp_dt = prev_deja + timedelta(hours=descanso_entre_jornada)
            else:
                d_busca = d_prev
                while d_busca in dias_historia and dias_historia[d_busca]["es_franco"]:
                    d_busca -= timedelta(days=1)
                if d_busca in dias_historia and dias_historia[d_busca]["deja_dt"]:
                    prev_deja = dias_historia[d_busca]["deja_dt"]
                    disp_dt = prev_deja + timedelta(hours=descanso_post_franco)

            if disp_dt:
                if disp_dt.date() == d_step:
                    hora_disp_str = disp_dt.strftime("%H:%M")
                elif disp_dt.date() < d_step:
                    hora_disp_str = "00:00"
                else:
                    hora_disp_str = disp_dt.strftime("%d/%m %H:%M")
            else:
                hora_disp_str = "-"

        dias_historia[d_step] = {
            "d": d_step,
            "es_franco": es_franco,
            "estado": estado,
            "toma": toma,
            "deja": deja,
            "deja_dt": deja_dt,
            "hs_trabajadas": hs_trab,
            "hs_trabajadas_str": f"{hs_trab:.1f} hs" if hs_trab > 0 else "-",
            "hs_semana_str": hs_semana_str,
            "hora_disponible": hora_disp_str,
            "franco_anticipado": bool(jornada_db and jornada_db.get("franco_anticipado")),
            "observaciones": jornada_db.get("observaciones") or "" if jornada_db else "",
            "novedad_id": novedad_id,
            "reemplaza_a": reemplaza_a,
            "turno_cubierto": turno_cubierto,
            "motivo": motivo,
            "periodo": periodo,
            "reemplazando_legajo": reemplazando_legajo,
            "reemplazando_nombre": reemplazando_nombre,
            "novedad_titular": novedad_titular,
            "tramo_inicio": tramo_inicio,
            "tramo_fin": tramo_fin,
        }
        d_step += timedelta(days=1)

    resultado = []
    for dia_num in range(1, num_dias + 1):
        f_actual = date(anio, mes, dia_num)
        info = dias_historia[f_actual]
        resultado.append({
            "dia": dia_num,
            "fecha": f_actual.strftime("%d/%m/%Y"),
            "fecha_iso": f_actual.strftime("%Y-%m-%d"),
            "dia_semana": _DIAS_SEMANA[f_actual.weekday()],
            "estado": info["estado"],
            "hora_disponible": info["hora_disponible"],
            "toma": info["toma"] or "-",
            "deja": info["deja"] or "-",
            "horas_trabajadas": info["hs_trabajadas_str"],
            "hs_semana": info["hs_semana_str"],
            "novedad_id": info["novedad_id"],
            "reemplaza_a": info["reemplaza_a"],
            "turno_cubierto": info["turno_cubierto"],
            "motivo": info["motivo"],
            "periodo": info["periodo"],
            "reemplazando_legajo": info["reemplazando_legajo"],
            "reemplazando_nombre": info["reemplazando_nombre"],
            "novedad_titular": info["novedad_titular"],
            "tramo_inicio": info["tramo_inicio"],
            "tramo_fin": info["tramo_fin"],
            "franco_anticipado": info["franco_anticipado"],
            "observaciones": info["observaciones"],
        })

    return resultado


calcular_descansos_ciclo = calcular_descansos_y_horas_ciclo


def resolver_subindice_diagrama_6d(dia_semana_nombre: str, franco_agente_nombre: str) -> tuple[int | None, bool]:
    """Calcula el subíndice (1 a 6) de un diagrama de 6 días según el día objetivo y el franco del agente.

    Devuelve (subindice, es_franco). Si el día de la semana coincide con el franco, devuelve (None, True).
    """
    dia_norm = normalizar_dia_semana(dia_semana_nombre)
    franco_norm = normalizar_dia_semana(franco_agente_nombre)

    if dia_norm == franco_norm:
        return None, True

    # Los 7 días en orden semanal empezando por Lunes (0) hasta Domingo (6)
    dias_orden = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "domingo"]
    if dia_norm not in dias_orden:
        return None, False

    # Para Franco Domingo habitual:
    # Lunes=1, Martes=2, Miércoles=3, Jueves=4, Viernes=5, Sábado=6
    # Si el franco es otro día, el ciclo rota a partir del día siguiente al franco:
    dia_idx = dias_orden.index(dia_norm)
    if franco_norm in dias_orden:
        franco_idx = dias_orden.index(franco_norm)
        # Distancia desde el día posterior al franco
        dist = (dia_idx - (franco_idx + 1)) % 7
        subindice = dist + 1
        return subindice, False
    else:
        # Si no hay franco específico o no coincide, mapear Lunes=1..Sábado=6
        if dia_idx == 6:  # Domingo
            return None, True
        return dia_idx + 1, False


def resolver_jornada_itinerario(
    turno_base: str,
    fecha_target: date | datetime | str,
    franco_agente: str,
    turnos_dict: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Resuelve la jornada exacta (toma, deje, duración, código de variante y estado)

    para una fecha dada y un agente, cruzando el tipo de turno y su franco CCT.
    """
    f_obj = parsear_fecha(fecha_target)
    if not f_obj:
        raise ValueError("Fecha objetivo no válida.")

    dia_sem_nombre = _DIAS_SEMANA[f_obj.weekday()]
    base_clean = str(turno_base or "").strip().replace(",", ".")

    # 1. Verificar si ya es una clave de variante explícita (ej. '401.3')
    if base_clean in turnos_dict:
        t_info = turnos_dict[base_clean]
        return {
            "codigo_turno": base_clean,
            "es_franco": False,
            "toma": t_info.get("toma_servicio") or "",
            "deje": t_info.get("deje_servicio") or "",
            "duracion": t_info.get("duracion_jornada") or "",
            "tipo_turno": t_info.get("tipo_turno") or "DESCONOCIDO",
            "deposito": t_info.get("deposito_base") or "",
            "observaciones": t_info.get("observaciones_turno") or "",
        }

    # 2. Evaluar si es un Diagrama de 6 Días (ej: '401', '501', '301')
    variante_1 = f"{base_clean}.1"
    if variante_1 in turnos_dict:
        subidx, es_franco = resolver_subindice_diagrama_6d(dia_sem_nombre, franco_agente)
        if es_franco or subidx is None:
            return {
                "codigo_turno": f"{base_clean} (FRANCO)",
                "es_franco": True,
                "toma": "",
                "deje": "",
                "duracion": "",
                "tipo_turno": "DIAGRAMA_6D",
                "deposito": turnos_dict[variante_1].get("deposito_base", ""),
                "observaciones": f"Franco CCT ({dia_sem_nombre})",
            }

        cod_variante = f"{base_clean}.{subidx}"
        info_var = turnos_dict.get(cod_variante, {})
        return {
            "codigo_turno": cod_variante,
            "es_franco": False,
            "toma": info_var.get("toma_servicio") or "",
            "deje": info_var.get("deje_servicio") or "",
            "duracion": info_var.get("duracion_jornada") or "",
            "tipo_turno": "DIAGRAMA_6D",
            "deposito": info_var.get("deposito_base") or "",
            "observaciones": info_var.get("observaciones_turno") or "",
        }

    # 3. Evaluar Unipersonales / Programas / Franqueros
    dia_norm = normalizar_dia_semana(dia_sem_nombre)
    franco_norm = normalizar_dia_semana(franco_agente)

    if dia_norm == franco_norm:
        return {
            "codigo_turno": f"{base_clean} (FRANCO)",
            "es_franco": True,
            "toma": "",
            "deje": "",
            "duracion": "",
            "tipo_turno": "FRANQUERO" if "ORD" in base_clean else "UNIPERSONAL",
            "deposito": "",
            "observaciones": f"Franco CCT ({dia_sem_nombre})",
        }

    if base_clean in turnos_dict:
        t_info = turnos_dict[base_clean]
        return {
            "codigo_turno": base_clean,
            "es_franco": False,
            "toma": t_info.get("toma_servicio") or "",
            "deje": t_info.get("deje_servicio") or "",
            "duracion": t_info.get("duracion_jornada") or "",
            "tipo_turno": t_info.get("tipo_turno") or "UNIPERSONAL",
            "deposito": t_info.get("deposito_base") or "",
            "observaciones": t_info.get("observaciones_turno") or "",
        }

    return {
        "codigo_turno": base_clean,
        "es_franco": False,
        "toma": "",
        "deje": "",
        "duracion": "",
        "tipo_turno": "DESCONOCIDO",
        "deposito": "",
        "observaciones": "",
    }


def validar_datos_ordenamiento(
    datos: dict[str, Any],
    tramos: Sequence[dict[str, Any]],
) -> tuple[bool, str | None]:
    """Valida los campos obligatorios de un Ordenamiento Operativo y su lista de tramos."""
    legajo_raw = str(datos.get("legajo_personal", "")).strip()
    if not legajo_raw:
        return False, "El campo 'Legajo' del personal es obligatorio."
    try:
        int(legajo_raw)
    except ValueError:
        return False, "El legajo del personal debe ser un número entero."

    if not str(datos.get("nombre_personal", "")).strip():
        return False, "El campo 'Nombre y Apellido' del personal es obligatorio."

    if not str(datos.get("deposito", "")).strip():
        return False, "El campo 'Dotación / Depósito' es obligatorio."

    if not str(datos.get("motivo_orden", "")).strip():
        return False, "El campo 'Motivo de la Orden' es obligatorio."

    if not str(datos.get("ref_pct", "")).strip():
        return False, "El campo 'Referencia de Estación (REF)' es obligatorio."

    try:
        parsear_fecha(datos.get("fecha_emision"), "Fecha de emisión", obligatorio=True)
    except ValueError as e:
        return False, str(e)

    if not tramos:
        return False, "Debe agregar al menos un tramo u observación en la tabla del ordenamiento."

    return True, None



