"""Pruebas unitarias completas para el servicio de actualizaciones automáticas (Auto-Updater OTA)."""

import hashlib
import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from src.services.updater_service import (
    InfoActualizacion,
    calcular_sha256,
    comparar_versiones,
    dividir_archivo,
    es_actualizacion_obligatoria,
    parsear_version,
    resolver_url_manifiesto,
)


class TestUpdaterService(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_parsear_version(self):
        self.assertEqual(parsear_version("2.0.0"), (2, 0, 0))
        self.assertEqual(parsear_version("v2.1.3"), (2, 1, 3))
        self.assertEqual(parsear_version("V1.7"), (1, 7, 0))
        self.assertEqual(parsear_version("3"), (3, 0, 0))
        self.assertEqual(parsear_version(""), (0, 0, 0))
        self.assertEqual(parsear_version("2.0.0-beta.1"), (2, 0, 0))

    def test_comparar_versiones_major(self):
        es_mayor, es_menor, es_parche, es_sup = comparar_versiones("2.0.0", "3.0.0")
        self.assertTrue(es_sup)
        self.assertTrue(es_mayor)
        self.assertFalse(es_menor)
        self.assertFalse(es_parche)

    def test_comparar_versiones_minor(self):
        es_mayor, es_menor, es_parche, es_sup = comparar_versiones("2.0.0", "2.1.0")
        self.assertTrue(es_sup)
        self.assertFalse(es_mayor)
        self.assertTrue(es_menor)
        self.assertFalse(es_parche)

    def test_comparar_versiones_patch(self):
        es_mayor, es_menor, es_parche, es_sup = comparar_versiones("2.0.0", "2.0.1")
        self.assertTrue(es_sup)
        self.assertFalse(es_mayor)
        self.assertFalse(es_menor)
        self.assertTrue(es_parche)

    def test_comparar_versiones_iguales_y_menores(self):
        # Misma versión
        self.assertEqual(comparar_versiones("2.0.0", "2.0.0"), (False, False, False, False))
        # Versión remota inferior
        self.assertEqual(comparar_versiones("2.1.0", "2.0.9"), (False, False, False, False))
        self.assertEqual(comparar_versiones("3.0.0", "2.9.9"), (False, False, False, False))

    def test_criterio_obligatoriedad_semver(self):
        # Versiones Mayores (X) y Menores (Y) son OBLIGATORIAS
        self.assertTrue(es_actualizacion_obligatoria("2.0.0", "3.0.0"))
        self.assertTrue(es_actualizacion_obligatoria("2.0.0", "2.1.0"))

        # Versión Parche (Z) es OPCIONAL
        self.assertFalse(es_actualizacion_obligatoria("2.0.0", "2.0.1"))
        self.assertFalse(es_actualizacion_obligatoria("2.0.1", "2.0.2"))

        # Si el flag de manifiesto es True, cualquier versión es OBLIGATORIA
        self.assertTrue(es_actualizacion_obligatoria("2.0.0", "2.0.1", flag_manifiesto=True))

        # Si local es menor a min_version_permitida, es OBLIGATORIA
        self.assertTrue(
            es_actualizacion_obligatoria(
                v_local="1.9.0",
                v_remota="2.0.1",
                min_version_permitida="2.0.0",
            )
        )
        self.assertFalse(
            es_actualizacion_obligatoria(
                v_local="2.0.0",
                v_remota="2.0.1",
                min_version_permitida="2.0.0",
            )
        )

    def test_info_actualizacion_dataclass(self):
        data = {
            "version": "2.0.1",
            "fecha": "12/11/2026",
            "titulo": "Actualización de Rendimiento",
            "novedades": ["[Operación] Mejora en carga de novedades", "[Seguridad] Corrección en auth"],
            "instalador_nombre": "RENO_Setup_v2.0.1.exe",
            "instalador_url": "https://supabase.co/storage/v1/object/public/updates/RENO_Setup_v2.0.1.exe",
            "sha256": "abcdef1234567890",
            "tamano_bytes": 45000000,
            "obligatoria": False,
            "min_version_permitida": "2.0.0",
            "partes": [
                "https://supabase.co/storage/v1/object/public/updates/RENO_Setup_v2.0.1.exe.part1",
                "https://supabase.co/storage/v1/object/public/updates/RENO_Setup_v2.0.1.exe.part2",
            ],
        }
        info = InfoActualizacion.from_dict(data)
        self.assertEqual(info.version, "2.0.1")
        self.assertEqual(info.fecha, "12/11/2026")
        self.assertEqual(len(info.novedades), 2)
        self.assertEqual(info.tamano_bytes, 45000000)
        self.assertFalse(info.obligatoria)
        self.assertEqual(len(info.partes), 2)

        dict_export = info.to_dict()
        self.assertEqual(dict_export["version"], "2.0.1")
        self.assertEqual(len(dict_export["partes"]), 2)

    def test_calcular_sha256(self):
        archivo = Path(self.temp_dir) / "prueba.bin"
        contenido = b"RENO Actualizacion Test Content 12345"
        with open(archivo, "wb") as f:
            f.write(contenido)

        esperado = hashlib.sha256(contenido).hexdigest().lower()
        calculado = calcular_sha256(archivo)
        self.assertEqual(calculado, esperado)

    def test_dividir_archivo_en_partes_y_reconstruir(self):
        # Crear un archivo simulado de 6 MB y dividirlo en partes de 2 MB
        archivo_original = Path(self.temp_dir) / "instalador_grande.bin"
        tamano_total = 6 * 1024 * 1024  # 6 MB
        contenido = os.urandom(tamano_total)

        with open(archivo_original, "wb") as f:
            f.write(contenido)

        hash_original = calcular_sha256(archivo_original)

        carpeta_partes = Path(self.temp_dir) / "partes"
        partes = dividir_archivo(archivo_original, tamano_max_mb=2, carpeta_salida=carpeta_partes)

        # Deben ser 3 partes de 2 MB cada una
        self.assertEqual(len(partes), 3)
        self.assertTrue(partes[0].name.endswith(".part1"))
        self.assertTrue(partes[1].name.endswith(".part2"))
        self.assertTrue(partes[2].name.endswith(".part3"))

        # Reconstruir concatenando las partes (tal como hace DescargarActualizacionWorker)
        archivo_reconstruido = Path(self.temp_dir) / "instalador_reconstruido.bin"
        with open(archivo_reconstruido, "wb") as f_out:
            for p in partes:
                with open(p, "rb") as f_in:
                    f_out.write(f_in.read())

        self.assertEqual(archivo_reconstruido.stat().st_size, tamano_total)
        hash_reconstruido = calcular_sha256(archivo_reconstruido)
        self.assertEqual(hash_reconstruido, hash_original)

    def test_dividir_archivo_menor_al_limite_no_se_divide(self):
        archivo_chico = Path(self.temp_dir) / "archivo_chico.bin"
        with open(archivo_chico, "wb") as f:
            f.write(b"Contenido menor a 25 MB")

        partes = dividir_archivo(archivo_chico, tamano_max_mb=25)
        self.assertEqual(len(partes), 1)
        self.assertEqual(partes[0], archivo_chico)

    def test_resolver_url_manifiesto(self):
        config = {
            "updater_url": "https://mi-servidor.com/version.json",
            "supabase_url": "https://xyz.supabase.co",
        }
        urls = resolver_url_manifiesto(config)
        self.assertIn("https://mi-servidor.com/version.json", urls)
        self.assertIn("https://xyz.supabase.co/storage/v1/object/public/updates/version.json", urls)
        self.assertIn("https://lucho-vidal.github.io/crono-pct-web/updates/reno_version.json", urls)


if __name__ == "__main__":
    unittest.main()
