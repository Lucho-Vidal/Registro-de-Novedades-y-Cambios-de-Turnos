from pathlib import Path
from unittest.mock import patch
import pytest

from src.data.database import DatabaseManager
from src.data.migrations import init_database_schema
from src.data.sync.supabase_sync import SupabaseSyncEngine, normalizar_valor_fecha
from src.services.records_service import RecordsService


def test_normalizar_valor_fecha():
    assert normalizar_valor_fecha(None) is None
    assert normalizar_valor_fecha("-") is None
    assert normalizar_valor_fecha("") is None
    assert normalizar_valor_fecha("None") is None
    assert normalizar_valor_fecha("22/08/2026") == "2026-08-22"
    assert normalizar_valor_fecha("5/11/2026") == "2026-11-05"
    assert normalizar_valor_fecha("2026-11-05") == "2026-11-05"


def test_sync_outbox_sanitizacion_y_resiliencia(tmp_path: Path):
    db_path = tmp_path / "test_sync.sqlite"
    db = DatabaseManager(db_path)
    init_database_schema(db)
    svc = RecordsService(db)
    engine = SupabaseSyncEngine(db)

    # Encolar un payload malformado de prueba (reemplazos legado vacíos)
    with db.write_transaction() as conn:
        conn.execute(
            """INSERT INTO sync_outbox (tabla, registro_pk, operacion, payload_json, creado_en)
               VALUES ('novedad_reemplazos', '9999', 'UPDATE', '{"reemplazos": []}', '2026-11-05 12:00:00');"""
        )
        conn.execute(
            """INSERT INTO sync_outbox (tabla, registro_pk, operacion, payload_json, creado_en)
               VALUES ('novedades', '9999', 'UPDATE', '{"id": 9999, "fecha_inicio": "05/11/2026", "fecha_fin": "-", "tramos_count": 3}', '2026-11-05 12:00:00');"""
        )

    assert engine.outbox.contar_pendientes() == 2

    # Mock HTTP response de Supabase
    with patch("src.data.sync.supabase_sync.ejecutar_request_http") as mock_http:
        mock_http.return_value = (201, b'[{"id": 9999}]', {})
        enviados, recibidos = engine._ejecutar_sync(solo_push=True)

    # El registro huérfano '9999' de novedad_reemplazos debe descartarse sin error
    # y el registro de novedades debe enviarse sanitizado
    assert engine.outbox.contar_pendientes() == 0
    assert mock_http.called
