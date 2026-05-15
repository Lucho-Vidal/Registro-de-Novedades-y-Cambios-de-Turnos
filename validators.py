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
