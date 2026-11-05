"""Pruebas unitarias de Drag & Drop y métodos rápidos de asignación/liberación en Personal de Ciclo."""

from pathlib import Path
from PySide6.QtWidgets import QApplication

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.services.records_service import RecordsService
from src.services.excel_service import ExcelService
from src.services.pdf_service import PDFService
from src.ui.views.personal_ciclo_view import PersonalCicloView


def test_personal_ciclo_drag_drop_actions(tmp_path: Path):
    app = QApplication.instance() or QApplication([])

    db_file = tmp_path / "test_ciclo_dd.sqlite"
    db = DatabaseManager(db_file, journal_mode="WAL")
    init_database_schema(db)

    svc = RecordsService(db)
    excel_svc = ExcelService(db)
    pdf_svc = PDFService()

    # 1. Crear titular y agente de ciclo
    with db.write_transaction() as conn:
        conn.execute(
            """INSERT INTO empleados (legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, activo, actualizado_en)
               VALUES (20189, 'FERNANDEZ Juan Ariel', 'Conductor Electrico', 'PC', '401', 'Lunes', 1, '2026-08-01 00:00:00'),
                      (12479, 'ALEGRE Maximiliano Nicolas', 'Conductor Electrico', 'PC', 'CICLO 11', 'Domingo', 1, '2026-08-01 00:00:00');"""
        )
        conn.execute(
            """INSERT INTO novedades (id, registrado_en, legajo, apellidos_nombres, especialidad, dotacion, turnos, franco, novedad, fecha_inicio, fecha_fin, referencia_estacion, supervisor, observaciones, activo)
               VALUES (5753, '15/08/2026 10:00', 20189, 'FERNANDEZ Juan Ariel', 'Conductor Electrico', 'PC', '401', 'Lunes', 'Enfermo', '15/08/2026', '-', 'CARDOZO', 'TORRES', '', 1);"""
        )

    # 2. Asignar reemplazo rápido (simula drop en ocupados)
    r_id = svc.asignar_reemplazo_rapido(
        novedad_id=5753,
        legajo_reemplazo=12479,
        fecha_inicio="15/08/2026",
        fecha_fin=None,
    )
    assert r_id > 0

    # Verificar estado de personal de ciclo: debe figurar ocupado
    estado1 = svc.obtener_estado_personal_ciclo("03/11/2026", dotacion="PC", especialidad="CONDUCTOR")
    ocupados = [o for o in estado1["ocupados"] if o["legajo"] == 12479]
    assert len(ocupados) == 1
    assert ocupados[0]["reemplazo_id"] == r_id

    # 3. Desactivar reemplazo (simula drop en libres)
    svc.desactivar_reemplazo(r_id)

    # Verificar estado de personal de ciclo: ahora debe figurar libre
    estado2 = svc.obtener_estado_personal_ciclo("03/11/2026", dotacion="PC", especialidad="CONDUCTOR")
    libres = [l for l in estado2["libres"] if l["legajo"] == 12479]
    assert len(libres) == 1

    # 4. Instanciar vista y verificar componentes
    view = PersonalCicloView(svc, pdf_svc, excel_svc)
    assert view.table_libres.rol == "LIBRES"
    assert view.table_ocupados.rol == "OCUPADOS"
    assert view.table_libres.acceptDrops() is True
    assert view.table_ocupados.acceptDrops() is True

    view.close()
