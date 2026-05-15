"""Unit tests for validators.py - Validación de datos para formularios."""

import pytest
from datetime import datetime
import tkinter as tk
from ttkbootstrap import DateEntry
from validators import (
    parsear_fecha,
    validar_campos_requeridos_novedades,
    validar_campos_requeridos_cambios,
)


@pytest.fixture(scope="session")
def tk_root():
    root = tk.Tk()
    root.withdraw()
    yield root
    root.destroy()


class TestParsearFecha:
    """Pruebas para la función parsear_fecha."""

    def test_parsear_fecha_valida(self):
        """Debe parsear una fecha válida en formato DD/MM/YYYY."""
        resultado = parsear_fecha("31/12/2025", "Fecha", obligatorio=False)
        assert resultado == datetime(2025, 12, 31)

    def test_parsear_fecha_valida_obligatorio(self):
        """Debe parsear una fecha válida cuando es obligatorio."""
        resultado = parsear_fecha("01/01/2026", "Fecha", obligatorio=True)
        assert resultado == datetime(2026, 1, 1)

    def test_parsear_fecha_vacia_no_obligatorio(self):
        """Debe retornar None si la fecha está vacía y no es obligatorio."""
        resultado = parsear_fecha("", "Fecha", obligatorio=False)
        assert resultado is None

    def test_parsear_fecha_vacia_con_espacios_no_obligatorio(self):
        """Debe retornar None si la fecha es solo espacios y no es obligatorio."""
        resultado = parsear_fecha("   ", "Fecha", obligatorio=False)
        assert resultado is None

    def test_parsear_fecha_vacia_obligatorio(self):
        """Debe lanzar ValueError si la fecha está vacía y es obligatorio."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha("", "Fecha inicio", obligatorio=True)
        assert "El campo 'Fecha inicio' es obligatorio" in str(excinfo.value)

    def test_parsear_fecha_formato_invalido(self):
        """Debe lanzar ValueError si el formato no es DD/MM/YYYY."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha("2025-12-31", "Fecha", obligatorio=True)
        assert "formato DD/MM/AAAA" in str(excinfo.value)

    def test_parsear_fecha_formato_invalido_mm_dd(self):
        """Debe lanzar ValueError si el formato es MM/DD/YYYY."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha("12/31/2025", "Fecha", obligatorio=True)
        assert "formato DD/MM/AAAA" in str(excinfo.value)

    def test_parsear_fecha_con_espacios_trimmed(self):
        """Debe trimmar espacios antes de parsear."""
        resultado = parsear_fecha("  15/06/2025  ", "Fecha", obligatorio=True)
        assert resultado == datetime(2025, 6, 15)

    def test_parsear_fecha_dia_invalido(self):
        """Debe lanzar ValueError si el día es inválido (ej. 32/01/2025)."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha("32/01/2025", "Fecha", obligatorio=True)
        assert "formato DD/MM/AAAA" in str(excinfo.value)

    def test_parsear_fecha_mes_invalido(self):
        """Debe lanzar ValueError si el mes es inválido (ej. 15/13/2025)."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha("15/13/2025", "Fecha", obligatorio=True)
        assert "formato DD/MM/AAAA" in str(excinfo.value)

    def test_parsear_fecha_valor_none(self):
        """Debe manejar None como valor vacío."""
        resultado = parsear_fecha(None, "Fecha", obligatorio=False)
        assert resultado is None

    def test_parsear_fecha_valor_entero(self):
        """Debe lanzar ValueError si recibe un número."""
        with pytest.raises(ValueError) as excinfo:
            parsear_fecha(15062025, "Fecha", obligatorio=True)
        assert "formato DD/MM/AAAA" in str(excinfo.value)


class MockDateEntry:
    def __init__(self, value="", master=None):
        self.entry = tk.Entry(master=master)
        self.entry.insert(0, value)

    def get(self):
        return self.entry.get()


class TestValidarCamposRequeridosNovedades:
    """Pruebas para validar_campos_requeridos_novedades."""

    @pytest.fixture
    def campos_validos(self, tk_root):
        """Fixture con campos válidos para una novedad."""
        legajo = tk.StringVar(master=tk_root, value="12345")
        apellidos_nombres = tk.StringVar(master=tk_root, value="García, Juan")
        novedad = tk.StringVar(master=tk_root, value="Licencia")
        referencia_estacion = tk.StringVar(master=tk_root, value="EST001")
        supervisor = tk.StringVar(master=tk_root, value="López, Carlos")
        fecha_inicio = MockDateEntry("15/06/2025", master=tk_root)
        fecha_fin = MockDateEntry("20/06/2025", master=tk_root)
        tipo_novedades = ["Licencia", "Enfermedad", "Falta"]

        return {
            "legajo": legajo,
            "apellidos_nombres": apellidos_nombres,
            "novedad": novedad,
            "referencia_estacion": referencia_estacion,
            "supervisor": supervisor,
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin,
            "tipo_novedades": tipo_novedades,
        }

    def test_validacion_exitosa(self, campos_validos):
        """Debe retornar (True, None) si todos los campos son válidos."""
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is True
        assert msg is None

    def test_legajo_vacio(self, campos_validos):
        """Debe retornar error si Legajo está vacío."""
        campos_validos["legajo"].set("")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Legajo" in msg

    def test_apellidos_nombres_vacio(self, campos_validos):
        """Debe retornar error si Apellido y Nombre está vacío."""
        campos_validos["apellidos_nombres"].set("")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Apellido y Nombre" in msg

    def test_novedad_vacia(self, campos_validos):
        """Debe retornar error si Tipo de Novedad está vacío."""
        campos_validos["novedad"].set("")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Tipo de Novedad" in msg

    def test_novedad_invalida(self, campos_validos):
        """Debe retornar error si Tipo de Novedad no está en la lista válida."""
        campos_validos["novedad"].set("Tipo Inválido")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "valido" in msg or "Tipo de Novedad" in msg

    def test_referencia_estacion_vacia(self, campos_validos):
        """Debe retornar error si Referencia Estación está vacía."""
        campos_validos["referencia_estacion"].set("")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Referencia Estacion" in msg

    def test_supervisor_vacio(self, campos_validos):
        """Debe retornar error si Supervisor está vacío."""
        campos_validos["supervisor"].set("")
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Supervisor" in msg

    def test_fecha_inicio_invalida(self, campos_validos):
        """Debe retornar error si Fecha de Inicio tiene formato inválido."""
        campos_validos["fecha_inicio"].entry.delete(0, tk.END)
        campos_validos["fecha_inicio"].entry.insert(0, "2025-06-15")  # Formato incorrecto
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "Fecha de inicio novedad" in msg or "formato" in msg

    def test_fecha_fin_anterior_a_inicio(self, campos_validos):
        """Debe retornar error si Fecha Fin es anterior a Fecha Inicio."""
        campos_validos["fecha_fin"].entry.delete(0, tk.END)
        campos_validos["fecha_fin"].entry.insert(0, "10/06/2025")  # Antes que inicio (15/06)
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is False
        assert "no puede ser anterior" in msg

    def test_fecha_fin_vacia_es_valida(self, campos_validos):
        """Debe permitir que Fecha Fin esté vacía (es opcional)."""
        campos_validos["fecha_fin"].entry.delete(0, tk.END)
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is True
        assert msg is None

    def test_fecha_fin_igual_a_inicio(self, campos_validos):
        """Debe permitir que Fecha Fin sea igual a Fecha Inicio."""
        campos_validos["fecha_fin"].entry.delete(0, tk.END)
        campos_validos["fecha_fin"].entry.insert(0, "15/06/2025")  # Igual que inicio
        es_valido, msg = validar_campos_requeridos_novedades(
            campos_validos["legajo"],
            campos_validos["apellidos_nombres"],
            campos_validos["novedad"],
            campos_validos["referencia_estacion"],
            campos_validos["supervisor"],
            campos_validos["fecha_inicio"],
            campos_validos["fecha_fin"],
            campos_validos["tipo_novedades"],
        )
        assert es_valido is True
        assert msg is None


class TestValidarCamposRequeridosCambios:
    """Pruebas para validar_campos_requeridos_cambios."""

    @pytest.fixture
    def campos_validos_cambios(self, tk_root):
        """Fixture con campos válidos para un cambio de turno."""
        legajo = tk.StringVar(master=tk_root, value="12345")
        apellidos_nombres = tk.StringVar(master=tk_root, value="García, Juan")
        legajo_2 = tk.StringVar(master=tk_root, value="67890")
        apellidos_nombres_2 = tk.StringVar(master=tk_root, value="López, Maria")
        fecha_cambio = MockDateEntry("15/06/2025", master=tk_root)
        referencia_estacion = tk.StringVar(master=tk_root, value="EST001")
        supervisor = tk.StringVar(master=tk_root, value="Pérez, Carlos")

        return {
            "legajo": legajo,
            "apellidos_nombres": apellidos_nombres,
            "legajo_2": legajo_2,
            "apellidos_nombres_2": apellidos_nombres_2,
            "fecha_cambio": fecha_cambio,
            "referencia_estacion": referencia_estacion,
            "supervisor": supervisor,
        }

    def test_validacion_exitosa(self, campos_validos_cambios):
        """Debe retornar (True, None) si todos los campos son válidos."""
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is True
        assert msg is None

    def test_legajo_vacio(self, campos_validos_cambios):
        """Debe retornar error si Legajo 1 está vacío."""
        campos_validos_cambios["legajo"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Legajo" in msg

    def test_apellidos_nombres_vacio(self, campos_validos_cambios):
        """Debe retornar error si Apellido y Nombre 1 está vacío."""
        campos_validos_cambios["apellidos_nombres"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Apellido y Nombre" in msg

    def test_legajo_2_vacio(self, campos_validos_cambios):
        """Debe retornar error si Legajo 2 está vacío."""
        campos_validos_cambios["legajo_2"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Legajo 2" in msg

    def test_apellidos_nombres_2_vacio(self, campos_validos_cambios):
        """Debe retornar error si Apellido y Nombre 2 está vacío."""
        campos_validos_cambios["apellidos_nombres_2"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Apellido y Nombre 2" in msg

    def test_fecha_cambio_vacia(self, campos_validos_cambios):
        """Debe retornar error si Fecha de cambio de turno está vacía."""
        campos_validos_cambios["fecha_cambio"].entry.delete(0, tk.END)
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Fecha de cambio de turno" in msg

    def test_fecha_cambio_formato_invalido(self, campos_validos_cambios):
        """Debe retornar error si Fecha de cambio tiene formato inválido."""
        campos_validos_cambios["fecha_cambio"].entry.delete(0, tk.END)
        campos_validos_cambios["fecha_cambio"].entry.insert(0, "2025-06-15")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False

    def test_referencia_estacion_vacia(self, campos_validos_cambios):
        """Debe retornar error si Referencia Estación está vacía."""
        campos_validos_cambios["referencia_estacion"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Referencia Estacion" in msg

    def test_supervisor_vacio(self, campos_validos_cambios):
        """Debe retornar error si Supervisor está vacío."""
        campos_validos_cambios["supervisor"].set("")
        es_valido, msg = validar_campos_requeridos_cambios(
            campos_validos_cambios["legajo"],
            campos_validos_cambios["apellidos_nombres"],
            campos_validos_cambios["legajo_2"],
            campos_validos_cambios["apellidos_nombres_2"],
            campos_validos_cambios["fecha_cambio"],
            campos_validos_cambios["referencia_estacion"],
            campos_validos_cambios["supervisor"],
        )
        assert es_valido is False
        assert "Supervisor" in msg
