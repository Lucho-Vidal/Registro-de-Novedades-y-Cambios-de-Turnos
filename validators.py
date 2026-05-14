"""Validación de datos para formularios de novedades y cambios de turnos."""

from datetime import datetime


def parsear_fecha(valor, nombre_campo, obligatorio=False):
    """Parsea una cadena de fecha en formato DD/MM/YYYY a objeto datetime.
    
    Retorna None si no es obligatorio y está vacío.
    Lanza ValueError si el formato es inválido o el campo es obligatorio pero vacío.
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
    
    Retorna (es_valido, mensaje_error).
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
    
    Retorna (es_valido, mensaje_error).
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
