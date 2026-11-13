#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script CLI para empaquetado, cálculo de checksums, división multi-parte y publicación de actualizaciones para RENO.

Adopta el estándar de CRONO-PCT:
1. Detecta la versión actual de la aplicación.
2. Localiza el instalador generado por Inno Setup (o el ejecutable compilado).
3. Calcula el hash criptográfico SHA-256 y el tamaño exacto en bytes.
4. Si el archivo excede los 48 MB (o flag --split), lo divide automáticamente en partes
   de ~25 MB (.part1, .part2) para eludir el límite estricto de 50 MB de Supabase Storage (Free Tier).
5. Genera el manifiesto version.json (y reno_version.json).
6. Opcionalmente sube los binarios y el manifiesto a Supabase Storage (bucket 'updates').
"""

import argparse
import datetime
import hashlib
import json
import os
import sys
from pathlib import Path

# Agregar raíz al path para importar módulos de RENO
ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))

from src.core.config import APP_NAME, APP_VERSION, cargar_configuracion_app
from src.services.updater_service import calcular_sha256, dividir_archivo


def parse_args():
    parser = argparse.ArgumentParser(description="Publicador y empaquetador de actualizaciones para RENO")
    parser.add_argument(
        "--version-app",
        default=APP_VERSION,
        help=f"Versión de la aplicación a publicar (por defecto: {APP_VERSION})",
    )
    parser.add_argument(
        "--instalador",
        default="",
        help="Ruta al ejecutable o instalador generado (ej. Instalador/RENO_Setup_v2.0.0.exe)",
    )
    parser.add_argument(
        "--titulo",
        default="",
        help="Título descriptivo del release (ej. 'Mejoras en Módulo de Reemplazos y Sincronización')",
    )
    parser.add_argument(
        "--obligatoria",
        action="store_true",
        help="Marca la actualización como de instalación obligatoria en el cliente",
    )
    parser.add_argument(
        "--min-version",
        default="",
        help="Versión mínima requerida para forzar actualización obligatoria",
    )
    parser.add_argument(
        "--split",
        action="store_true",
        help="Fuerza la división del instalador en partes de 25 MB aunque pese menos de 48 MB",
    )
    parser.add_argument(
        "--tamano-parte-mb",
        type=int,
        default=25,
        help="Tamaño de cada parte al dividir (por defecto 25 MB para asegurar entrada en Supabase)",
    )
    parser.add_argument(
        "--base-url",
        default="",
        help="URL base pública donde se alojarán los archivos (ej. https://.../storage/v1/object/public/updates)",
    )
    parser.add_argument(
        "--salida-dir",
        default=str(ROOT_DIR / "Instalador"),
        help="Carpeta donde se guardará el version.json y los fragmentos divididos",
    )
    parser.add_argument(
        "--subir",
        action="store_true",
        help="Intenta subir automáticamente los archivos a Supabase Storage bucket 'updates'",
    )
    parser.add_argument(
        "--supabase-url",
        default=os.environ.get("SUPABASE_URL", ""),
        help="URL base de Supabase (por defecto lee SUPABASE_URL de .env)",
    )
    parser.add_argument(
        "--service-key",
        default=os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "") or os.environ.get("SUPABASE_SERVICE_KEY", ""),
        help="Clave de servicio (service_role) para subida desatendida a Supabase Storage (por defecto lee SUPABASE_SERVICE_ROLE_KEY de .env)",
    )
    parser.add_argument(
        "--bucket",
        default=os.environ.get("SUPABASE_STORAGE_BUCKET", "updates"),
        help="Bucket de Supabase Storage donde se publicarán los archivos (por defecto: 'updates')",
    )
    return parser.parse_args()


def buscar_instalador_candidato(version: str, carpeta_salida: Path) -> Path | None:
    """Busca automáticamente el instalador más probable según convenciones."""
    candidatos = [
        carpeta_salida / f"RENO_Setup_v{version}.exe",
        carpeta_salida / "RENO.exe",
        carpeta_salida / f"RENO_Setup.exe",
        ROOT_DIR / "dist" / "RENO.exe",
    ]
    for c in candidatos:
        if c.exists():
            return c
    return None


def extraer_novedades_desde_changelog(version: str) -> list[str]:
    """Intenta extraer las viñetas del changelog si existe CHANGELOG.md."""
    changelog_path = ROOT_DIR / "CHANGELOG.md"
    if not changelog_path.exists():
        return [
            f"[Actualización] Nueva versión {version} de RENO con mejoras de rendimiento y estabilidad.",
            "[Operación] Sincronización híbrida de base de datos optimizada y corrección de interfaz.",
            "[Seguridad] Actualización de dependencias y validaciones criptográficas de integridad.",
        ]

    try:
        with open(changelog_path, "r", encoding="utf-8") as f:
            lineas = f.readlines()

        novedades = []
        en_version = False
        for linea in lineas:
            linea_s = linea.strip()
            if linea_s.startswith("## ") or linea_s.startswith("### "):
                if version in linea_s:
                    en_version = True
                    continue
                elif en_version:
                    break
            if en_version and (linea_s.startswith("- ") or linea_s.startswith("* ")):
                item = linea_s.lstrip("-* ").strip()
                if item:
                    novedades.append(item)
        if novedades:
            return novedades
    except Exception:
        pass

    return [
        f"[Actualización] Nueva versión {version} de RENO con mejoras de estabilidad.",
        "[Operación] Refresco de datos en red y optimización de memoria.",
    ]


def subir_archivo_a_supabase(
    supabase_url: str,
    bucket: str,
    api_key: str,
    archivo_local: Path,
) -> bool:
    """Sube un archivo a Supabase Storage con soporte de sobrescritura (x-upsert: true)."""
    import mimetypes

    nombre_archivo = archivo_local.name
    endpoint = f"{supabase_url.rstrip('/')}/storage/v1/object/{bucket}/{nombre_archivo}"
    content_type, _ = mimetypes.guess_type(str(archivo_local))
    if not content_type:
        content_type = "application/json" if archivo_local.suffix == ".json" else "application/octet-stream"

    headers = {
        "apikey": api_key,
        "Authorization": f"Bearer {api_key}",
        "x-upsert": "true",
        "Content-Type": content_type,
    }

    tamano_bytes = archivo_local.stat().st_size
    tamano_mb = tamano_bytes / (1024 * 1024)
    print(f"  -> Subiendo {nombre_archivo} ({tamano_mb:.2f} MB)...", end=" ", flush=True)

    try:
        try:
            import httpx

            with open(archivo_local, "rb") as f:
                contenido = f.read()
            with httpx.Client(timeout=300.0) as client:
                resp = client.post(endpoint, headers=headers, content=contenido)
                if resp.status_code in (200, 201):
                    print("✅ [OK]")
                    return True
                else:
                    print(f"❌ Error HTTP {resp.status_code}: {resp.text}")
                    return False
        except ImportError:
            import urllib.request

            with open(archivo_local, "rb") as f:
                contenido = f.read()
            req = urllib.request.Request(endpoint, data=contenido, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=300) as resp:
                if resp.status in (200, 201):
                    print("✅ [OK]")
                    return True
                else:
                    print(f"❌ Error HTTP {resp.status}")
                    return False
    except Exception as e:
        print(f"❌ Falló subida: {e}")
        return False


def main():
    args = parse_args()
    version = args.version_app.strip().lstrip("vV")
    salida_dir = Path(args.salida_dir).resolve()
    salida_dir.mkdir(parents=True, exist_ok=True)

    print(f"=== Publicador de Actualizaciones RENO v{version} ===")

    # 1. Resolver ruta del instalador
    if args.instalador:
        ruta_instalador = Path(args.instalador).resolve()
    else:
        ruta_instalador = buscar_instalador_candidato(version, salida_dir)

    tamano_bytes = 0
    sha256_hash = ""
    partes_generadas = []

    if ruta_instalador and ruta_instalador.exists():
        tamano_bytes = ruta_instalador.stat().st_size
        tamano_mb = tamano_bytes / (1024 * 1024)
        print(f"Instalador encontrado: {ruta_instalador} ({tamano_mb:.2f} MB)")
        print("Calculando hash SHA-256...")
        sha256_hash = calcular_sha256(ruta_instalador)
        print(f"SHA-256: {sha256_hash}")

        # Comprobar límite de 50 MB de Supabase
        if tamano_mb > 48.0 or args.split:
            print(f"⚠️ El archivo supera los 48 MB ({tamano_mb:.1f} MB) o se solicitó --split.")
            print(f"Dividiendo en partes de ~{args.tamano_parte_mb} MB para cumplir con Supabase Storage (Free Tier)...")
            partes_paths = dividir_archivo(
                ruta_instalador,
                tamano_max_mb=args.tamano_parte_mb,
                carpeta_salida=salida_dir,
            )
            partes_generadas = [p.name for p in partes_paths]
            print(f"Se generaron {len(partes_generadas)} partes: {', '.join(partes_generadas)}")
    else:
        print("⚠️ Advertencia: No se encontró un binario de instalador. Se generará un manifiesto preliminar.")

    # 2. Resolver URLs
    cfg = cargar_configuracion_app()
    sb_url = (args.supabase_url or str(cfg.get("supabase_url", ""))).strip().rstrip("/")
    bucket = (args.bucket or str(cfg.get("supabase_storage_bucket", "updates"))).strip()

    base_url = args.base_url.rstrip("/")
    if not base_url:
        if sb_url:
            base_url = f"{sb_url}/storage/v1/object/public/{bucket}"
        else:
            base_url = "https://lucho-vidal.github.io/crono-pct-web/updates"

    nombre_exe = ruta_instalador.name if ruta_instalador else f"RENO_Setup_v{version}.exe"
    instalador_url = f"{base_url}/{nombre_exe}"

    partes_urls = [f"{base_url}/{p}" for p in partes_generadas] if partes_generadas else []

    # 3. Novedades y metadatos
    titulo = args.titulo or f"Actualización General de Rendimiento y Cobertura v{version}"
    novedades = extraer_novedades_desde_changelog(version)
    fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")

    manifiesto = {
        "version": version,
        "fecha": fecha_hoy,
        "titulo": titulo,
        "novedades": novedades,
        "instalador_nombre": nombre_exe,
        "instalador_url": instalador_url,
        "sha256": sha256_hash,
        "tamano_bytes": tamano_bytes,
        "obligatoria": args.obligatoria,
        "min_version_permitida": args.min_version,
    }

    if partes_urls:
        manifiesto["partes"] = partes_urls

    # 4. Guardar archivos version.json y reno_version.json
    ruta_version_json = salida_dir / "version.json"
    ruta_reno_version_json = salida_dir / "reno_version.json"

    with open(ruta_version_json, "w", encoding="utf-8") as f:
        json.dump(manifiesto, f, indent=2, ensure_ascii=False)
    with open(ruta_reno_version_json, "w", encoding="utf-8") as f:
        json.dump(manifiesto, f, indent=2, ensure_ascii=False)

    print(f"\n✅ Manifiesto generado en: {ruta_version_json}")
    print(f"✅ Manifiesto alternativo en: {ruta_reno_version_json}")
    print(json.dumps(manifiesto, indent=2, ensure_ascii=False))

    # 5. Subida automática a Supabase Storage (opcional)
    if args.subir:
        if not sb_url:
            print("❌ Error: No se especificó SUPABASE_URL ni en argumentos ni en .env.")
            return

        api_key = (
            args.service_key
            or str(cfg.get("supabase_service_role_key", ""))
            or str(cfg.get("supabase_anon_key", ""))
        )
        if not api_key:
            print("❌ Error: No se configuró SUPABASE_SERVICE_ROLE_KEY ni en .env ni como argumento.")
            return

        print(f"\nSubiendo archivos a Supabase Storage ({sb_url}/storage/v1/object/{bucket})...")

        archivos_a_subir = [ruta_version_json, ruta_reno_version_json]
        if partes_generadas:
            archivos_a_subir.extend([salida_dir / p for p in partes_generadas])
        elif ruta_instalador and ruta_instalador.exists():
            archivos_a_subir.append(ruta_instalador)

        exitos = 0
        for arch in archivos_a_subir:
            if subir_archivo_a_supabase(sb_url, bucket, api_key, arch):
                exitos += 1

        if exitos == len(archivos_a_subir):
            print(f"\n✅ Todos los archivos ({exitos}/{len(archivos_a_subir)}) se subieron exitosamente.")
        else:
            print(f"\n⚠️ Se subieron {exitos} de {len(archivos_a_subir)} archivos.")

    print("\n🎉 ¡Proceso de publicación completado con éxito!")



if __name__ == "__main__":
    main()
