"""Validación de datos para formularios de novedades y cambios de turnos."""

from datetime import datetime


def parsear_fecha(valor, nombre_campo, obligatorio=False):
    """Parsea una cadena de fecha en formato DD/MM/YYYY a objeto datetime.
    
    Args:
        valor: Cadena de fecha o None. Ejemplo: "31/12/2025".
        nombre_campo: Nombre del campo para mensajes de error. Ejemplo: "Fecha de inicio".
        obligatorio: Si True, lanza ValueError si el campo está vacío.
    
    Returns:
        datetime: Objeto datetime parseado si valor es válido.
        None: Si no es obligatorio y está vacío.
    
    Raises:
        ValueError: Si el formato es inválido (no es DD/MM/YYYY) o el campo es obligatorio pero vacío.
    
    Example:
        >>> parsear_fecha("15/06/2025", "Fecha inicio", obligatorio=True)
        datetime.datetime(2025, 6, 15, 0, 0)
        >>> parsear_fecha("", "Fecha fin", obligatorio=False)
        None
        >>> parsear_fecha("31-12-2025", "Fecha", obligatorio=True)  # Lanza ValueError
    """
    valor = str(valor or "").strip()
    if not valor:
        if obligatorio:
            raise ValueError(f"El campo '{nombre_campo}' es obligatorio.")
        return None
    try:
        return datetime.strptime(valor, "%d/%m/%Y")
    except ValueError:
        raise ValueError(f"El campo '{nombre_campo}' debe tener formato DD/MM/AAAA.")


def validar_campos_requeridos_novedades(legajo_var, apellidos_nombres_var, novedad_var, 
                                        referencia_estacion_var, supervisor_var,
                                        fecha_inicio_entry, fecha_fin_entry, tipo_novedades):
    """Valida los campos obligatorios del formulario de novedades.
    
    Verifica que todos los campos requeridos tengan valores válidos, incluyendo fechas
    en formato correcto y coherencia entre fechas de inicio y fin.
    
    Args:
        legajo_var: StringVar con el legajo del empleado.
        apellidos_nombres_var: StringVar con apellido y nombre.
        novedad_var: StringVar con el tipo de novedad seleccionado.
        referencia_estacion_var: StringVar con referencia de estación.
        supervisor_var: StringVar con nombre del supervisor.
        fecha_inicio_entry: DateEntry widget con fecha de inicio de novedad.
        fecha_fin_entry: DateEntry widget con fecha de fin de novedad (opcional).
        tipo_novedades: Lista de tipos de novedad válidos. Ejemplo: ["Licencia", "Enfermedad"].
    
    Returns:
        Tupla (es_valido, mensaje_error):
        - (True, None) si todas las validaciones pasan.
        - (False, "mensaje de error") si alguna validación falla.
    
    Raises:
        ValueError: (Nunca lanzado directamente; errores de fecha convertidos a tupla).
    
    Example:
        >>> validar_campos_requeridos_novedades(
        ...     legajo_var, apellidos_nombres_var, novedad_var,
        ...     referencia_estacion_var, supervisor_var,
        ...     fecha_inicio_entry, fecha_fin_entry,
        ...     ["Licencia", "Enfermedad"]
        ... )
        (True, None)  # Si todas las validaciones pasan
        (False, "El campo 'Legajo' es obligatorio.")  # Si legajo está vacío
    """
    if not legajo_var.get().strip():
        return False, "El campo 'Legajo' es obligatorio."
    if not apellidos_nombres_var.get().strip():
        return False, "El campo 'Apellido y Nombre' es obligatorio."
    if not novedad_var.get().strip():
        return False, "El campo 'Tipo de Novedad' es obligatorio."
    if novedad_var.get().strip() not in tipo_novedades:
        return False, "Seleccione un valor valido en 'Tipo de Novedad'."
    if not referencia_estacion_var.get().strip():
        return False, "El campo 'Referencia Estacion' es obligatorio."
    if not supervisor_var.get().strip():
        return False, "El campo 'Supervisor' es obligatorio."

    try:
        fecha_inicio = parsear_fecha(fecha_inicio_entry.entry.get(), "Fecha de inicio novedad", obligatorio=True)
        fecha_fin = parsear_fecha(fecha_fin_entry.entry.get(), "Fecha de fin novedad", obligatorio=False)
    except ValueError as e:
        return False, str(e)

    if fecha_fin and fecha_fin < fecha_inicio:
        return False, "'Fecha de Fin Novedad' no puede ser anterior a 'Fecha de Inicio Novedad'."

    return True, None


def validar_campos_requeridos_cambios(legajo_var, apellidos_nombres_var, legajo_2_var, 
                                      apellidos_nombres_2_var, fecha_cambio_turno_entry,
                                      referencia_estacion_var, supervisor_var):
    """Valida los campos obligatorios del formulario de cambios de turnos.
    
    Verifica que ambos empleados, fechas y referencias sean proporcionados y tengan
    formatos válidos.
    
    Args:
        legajo_var: StringVar con el legajo del primer empleado.
        apellidos_nombres_var: StringVar con apellido y nombre del primer empleado.
        legajo_2_var: StringVar con el legajo del segundo empleado.
        apellidos_nombres_2_var: StringVar con apellido y nombre del segundo empleado.
        fecha_cambio_turno_entry: DateEntry widget con fecha del cambio de turno.
        referencia_estacion_var: StringVar con referencia de estación.
        supervisor_var: StringVar con nombre del supervisor.
    
    Returns:
        Tupla (es_valido, mensaje_error):
        - (True, None) si todas las validaciones pasan.
        - (False, "mensaje de error") si alguna validación falla.
    
    Raises:
        ValueError: (Nunca lanzado directamente; errores de fecha convertidos a tupla).
    
    Example:
        >>> validar_campos_requeridos_cambios(
        ...     legajo_var, apellidos_nombres_var, legajo_2_var, apellidos_nombres_2_var,
        ...     fecha_cambio_turno_entry, referencia_estacion_var, supervisor_var
        ... )
        (True, None)  # Si todas las validaciones pasan
        (False, "El campo 'Legajo' es obligatorio.")  # Si legajo está vacío
    """
    if not legajo_var.get().strip():
        return False, "El campo 'Legajo' es obligatorio."
    if not apellidos_nombres_var.get().strip():
        return False, "El campo 'Apellido y Nombre' es obligatorio."
    if not legajo_2_var.get().strip():
        return False, "El campo 'Legajo 2' es obligatorio."
    if not apellidos_nombres_2_var.get().strip():
        return False, "El campo 'Apellido y Nombre 2' es obligatorio."
    if not fecha_cambio_turno_entry.entry.get().strip():
        return False, "El campo 'Fecha de cambio de turno' es obligatorio."
    if not referencia_estacion_var.get().strip():
        return False, "El campo 'Referencia Estacion' es obligatorio."
    if not supervisor_var.get().strip():
        return False, "El campo 'Supervisor' es obligatorio."

    try:
        parsear_fecha(fecha_cambio_turno_entry.entry.get(), "Fecha de cambio de turno", obligatorio=True)
    except ValueError as e:
        return False, str(e)

    return True, None


def validar_datos_edicion_novedad(data, tipo_novedades=None, personal_estacion=None):
    """Valida los datos de edición de una novedad."""
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
        fecha_ini = parsear_fecha(data.get("fecha_inicio"), "Fecha de inicio novedad", obligatorio=True)
        fecha_fin = parsear_fecha(data.get("fecha_fin"), "Fecha de fin novedad", obligatorio=False)
    except ValueError as e:
        return False, str(e)

    if fecha_fin and fecha_fin < fecha_ini:
        return False, "'Fecha de Fin Novedad' no puede ser anterior a 'Fecha de Inicio Novedad'."

    return True, None


def validar_datos_edicion_cambio(data, personal_estacion=None):
    """Valida los datos de edición de un cambio de turno."""
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


def es_turno_ciclo(turno_str):
    """Indica si el turno corresponde a personal de CICLO (comienza con 'CICLO')."""
    return str(turno_str or "").strip().upper().startswith("CICLO")


def validar_tramos_reemplazo(fecha_inicio_nov, fecha_fin_nov, turno_titular, tramos):
    """Valida la lista de tramos de reemplazo para una novedad.

    Verifica:
    - Que el titular no sea de turno CICLO si se asignan reemplazos.
    - Que cada reemplazo tenga turno que comience con CICLO (si se informa el turno).
    - Que las fechas de cada tramo sean válidas y estén dentro del rango de la novedad.
    - Que los tramos no se solapen entre sí.
    """
    if not tramos:
        return True, None

    if es_turno_ciclo(turno_titular):
        return False, "El personal cuyo turno comienza con CICLO no puede tener reemplazo."

    try:
        f_ini_nov = parsear_fecha(fecha_inicio_nov, "Fecha de inicio novedad", obligatorio=True)
        f_fin_nov = parsear_fecha(fecha_fin_nov, "Fecha de fin novedad", obligatorio=False)
    except ValueError as e:
        return False, str(e)

    intervalos = []
    for idx, tramo in enumerate(tramos, start=1):
        leg = str(tramo.get("legajo_reemplazo") or tramo.get("legajo") or "").strip()
        if not leg:
            return False, f"El reemplazo #{idx} debe tener un legajo válido."
        try:
            int(leg)
        except ValueError:
            return False, f"El legajo del reemplazo #{idx} debe ser numérico."

        turno_remp = tramo.get("turnos_reemplazo") or tramo.get("turnos")
        if turno_remp is not None and str(turno_remp).strip() != "" and not es_turno_ciclo(turno_remp):
            return False, f"El reemplazo #{idx} debe pertenecer a un turno que comience con CICLO."

        try:
            t_ini = parsear_fecha(tramo.get("fecha_inicio"), f"Fecha de inicio del reemplazo #{idx}", obligatorio=True)
            t_fin = parsear_fecha(tramo.get("fecha_fin"), f"Fecha de fin del reemplazo #{idx}", obligatorio=False)
        except ValueError as e:
            return False, str(e)

        if t_fin and t_fin < t_ini:
            return False, f"En el reemplazo #{idx}, la fecha de fin no puede ser anterior a la fecha de inicio."

        if t_ini < f_ini_nov:
            return False, f"En el reemplazo #{idx}, la fecha de inicio ({tramo.get('fecha_inicio')}) no puede ser anterior al inicio de la novedad ({fecha_inicio_nov})."

        if f_fin_nov:
            if t_fin is None:
                return False, f"El reemplazo #{idx} debe tener fecha de fin porque la novedad finaliza el {fecha_fin_nov}."
            if t_fin > f_fin_nov:
                return False, f"En el reemplazo #{idx}, la fecha de fin ({tramo.get('fecha_fin')}) no puede superar el fin de la novedad ({fecha_fin_nov})."

        intervalos.append((t_ini, t_fin or datetime.max, idx))

    intervalos.sort(key=lambda item: (item[0], item[1]))
    for i in range(len(intervalos) - 1):
        actual_ini, actual_fin, idx_a = intervalos[i]
        sig_ini, sig_fin, idx_b = intervalos[i + 1]
        if actual_fin >= sig_ini:
            return False, f"Los tramos de reemplazo #{idx_a} y #{idx_b} se solapan en las fechas seleccionadas."

    return True, None


