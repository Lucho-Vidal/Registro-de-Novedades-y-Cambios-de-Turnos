#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Script CLI para empaquetado, cálculo de checksums, división multi-parte y publicación de actualizaciones para RENO.

Adopta el estándar de CRONO-PCT y Horos:
1. Carga automáticamente variables desde .env en la raíz (priorizando .env sobre sesión).
2. Localiza el instalador generado por Inno Setup en Instalador/ (o parámetro --instalador).
3. Calcula el hash criptográfico SHA-256 y el tamaño exacto en bytes.
4. Si el archivo excede los 48 MB (o flag --split), lo divide automáticamente en partes
   de ~25 MB (.part1, .part2, ...) para cumplir con el límite de 50 MB de Supabase Storage (Free Tier).
5. Genera los manifiestos version.json y reno_version.json.
6. Sube los binarios/fragmentos y manifiestos a Supabase Storage (bucket 'updates') soportando
   tanto claves modernas (sb_secret_* / sb_publishable_*) sin Bearer como JWT legacy (eyJ...).
"""

import argparse
import datetime
import json
import os
import sys
from pathlib import Path

# Agregar raíz al path para importar módulos de RENO
ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))


def _cargar_dotenv() -> dict[str, str]:
    """Carga variables desde .env en la raíz del proyecto priorizando sobre variables de sesión."""
    env_path = ROOT_DIR / ".env"
    cargadas: dict[str, str] = {}
    if not env_path.exists():
        return cargadas
    try:
        for raw_line in env_path.read_text(encoding="utf-8-sig").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            clave, valor = line.split("=", 1)
            clave = clave.strip()
            valor = valor.strip().strip('"').strip("'")
            if clave:
                cargadas[clave] = valor
                os.environ[clave] = valor
    except Exception:
        pass
    return cargadas


ENV_DOTENV = _cargar_dotenv()

from src.core.config import APP_VERSION, cargar_configuracion_app
from src.services.updater_service import calcular_sha256, dividir_archivo


def parse_args():
    parser = argparse.ArgumentParser(description="Publicador y empaquetador de actualizaciones OTA para RENO")
    parser.add_argument(
        "--version-app",
        "--version",
        dest="version_app",
        default=APP_VERSION,
        help=f"Versión de la aplicación a publicar (por defecto: {APP_VERSION})",
    )
    parser.add_argument(
        "--instalador",
        "--installer",
        dest="instalador",
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
        "--url",
        dest="base_url",
        default="",
        help="URL base pública donde se alojarán los archivos (también lee RENO_UPDATES_URL de .env)",
    )
    parser.add_argument(
        "--salida-dir",
        default=str(ROOT_DIR / "Instalador"),
        help="Carpeta donde se guardará el version.json y los fragmentos divididos",
    )
    parser.add_argument(
        "--subir",
        "--supabase",
        dest="subir",
        action="store_true",
        help="Fuerza la subida automática a Supabase Storage bucket 'updates' (también automático si .env tiene credenciales válidas)",
    )
    parser.add_argument(
        "--no-subir",
        "--solo-manifiesto",
        dest="no_subir",
        action="store_true",
        help="Solo genera los archivos locales (version.json y partes) sin subir a Supabase",
    )
    parser.add_argument(
        "--supabase-url",
        default=None,
        help="URL base de Supabase (por defecto lee SUPABASE_URL de .env)",
    )
    parser.add_argument(
        "--service-key",
        default=None,
        help="Secret Key (sb_secret_...) o Service Role JWT de Supabase Storage (por defecto lee SUPABASE_SERVICE_KEY / SUPABASE_SERVICE_ROLE_KEY de .env)",
    )
    parser.add_argument(
        "--bucket",
        default=None,
        help="Bucket de Supabase Storage donde se publicarán los archivos (por defecto: 'updates')",
    )
    return parser.parse_args()


def buscar_instalador_candidato(version: str, carpeta_salida: Path) -> Path | None:
    """Busca automáticamente el instalador más probable según convenciones."""
    candidatos = [
        carpeta_salida / f"RENO_Setup_v{version}.exe",
        carpeta_salida / "RENO_Setup.exe",
    ]
    for c in candidatos:
        if c.exists():
            return c

    if carpeta_salida.exists():
        setups = sorted(
            carpeta_salida.glob("RENO_Setup*.exe"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if setups:
            return setups[0]

    for c in (carpeta_salida / "RENO.exe", ROOT_DIR / "dist" / "RENO.exe"):
        if c.exists():
            return c

    if carpeta_salida.exists():
        exes = sorted(carpeta_salida.glob("*.exe"), key=lambda p: p.stat().st_mtime, reverse=True)
        if exes:
            return exes[0]

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


def construir_headers_storage(api_key: str, content_type: str) -> dict[str, str]:
    """Construye cabeceras compatibles con sb_secret_*/sb_publishable_* (sin Bearer) y JWT legacy."""
    key = (api_key or "").strip().strip('"').strip("'")
    headers = {
        "apikey": key,
        "x-upsert": "true",
        "Content-Type": content_type,
    }
    # Las claves modernas (sb_secret_... / sb_publishable_...) NO deben llevar Authorization: Bearer
    # porque Supabase Storage las rechaza con HTTP 400/401 ('Invalid Compact JWS').
    if key and not (key.startswith("sb_secret_") or key.startswith("sb_publishable_")):
        headers["Authorization"] = f"Bearer {key}"
    return headers


def _asegurar_bucket_publico(supabase_url: str, bucket: str, api_key: str) -> None:
    """Intenta crear el bucket público en Supabase Storage si no existiera."""
    import httpx

    print(f"     -> Bucket '{bucket}' no encontrado; intentando crearlo público...")
    endpoint = f"{supabase_url.rstrip('/')}/storage/v1/bucket"
    headers = construir_headers_storage(api_key, "application/json")
    headers.pop("x-upsert", None)
    payload = {"id": bucket, "name": bucket, "public": True}
    try:
        with httpx.Client(timeout=30.0) as client:
            client.post(endpoint, headers=headers, json=payload)
    except Exception:
        pass


def _es_url_supabase_valida(url: str) -> bool:
    """Determina si una URL de Supabase es real y no un valor de ejemplo."""
    u = (url or "").strip().lower()
    if not u:
        return False
    if any(p in u for p in ("tu-proyecto", "ejemplo.supabase", "xxxx.supabase")):
        return False
    return u.startswith(("http://", "https://"))


def _es_credencial_real(url: str, key: str) -> bool:
    """Determina si URL y Key corresponden a credenciales reales y no a valores de ejemplo."""
    if not _es_url_supabase_valida(url):
        return False
    k = (key or "").strip().lower()
    if not k:
        return False
    if k.endswith("...") or any(
        p in k
        for p in (
            "tu_clave",
            "tu_service_role",
            "pega_tu_clave",
            "clave_secreta_aqui",
            "key_aqui",
            "placeholder",
        )
    ):
        return False
    return len(k) > 15


def _imprimir_tip_error_upload(status_code: int, detalle: str) -> None:
    if "Invalid Compact JWS" in detalle:
        print(
            "     [TIP] Verificá que la clave en SUPABASE_SERVICE_KEY sea 'sb_secret_...' "
            "o el JWT legacy 'service_role' (eyJhbGci...)."
        )
    elif status_code == 413:
        print(
            "     [TIP] El archivo supera el límite de tamaño de Supabase Storage (50 MB en Free Tier). "
            "Usá --split o reducí --tamano-parte-mb."
        )


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

    headers = construir_headers_storage(api_key, content_type)
    tamano_bytes = archivo_local.stat().st_size
    tamano_mb = tamano_bytes / (1024 * 1024)
    print(f"  -> Subiendo {nombre_archivo} ({tamano_mb:.2f} MB)...", end=" ", flush=True)

    try:
        try:
            import httpx

            with open(archivo_local, "rb") as f:
                contenido = f.read()
            with httpx.Client(timeout=300.0, follow_redirects=True) as client:
                resp = client.post(endpoint, headers=headers, content=contenido)
                if resp.status_code in (400, 404) and "bucket not found" in resp.text.lower():
                    _asegurar_bucket_publico(supabase_url, bucket, api_key)
                    resp = client.post(endpoint, headers=headers, content=contenido)

                if resp.status_code in (200, 201):
                    print("✅ [OK]")
                    return True
                print(f"❌ Error HTTP {resp.status_code}: {resp.text}")
                _imprimir_tip_error_upload(resp.status_code, resp.text)
                return False
        except ImportError:
            import urllib.error
            import urllib.request

            with open(archivo_local, "rb") as f:
                contenido = f.read()
            req = urllib.request.Request(endpoint, data=contenido, headers=headers, method="POST")
            try:
                with urllib.request.urlopen(req, timeout=300) as resp:
                    if resp.status in (200, 201):
                        print("✅ [OK]")
                        return True
                    print(f"❌ Error HTTP {resp.status}")
                    return False
            except urllib.error.HTTPError as e:
                body_txt = e.read().decode("utf-8", errors="ignore")
                print(f"❌ Error HTTP {e.code}: {body_txt}")
                _imprimir_tip_error_upload(e.code, body_txt)
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

    # 2. Resolver URLs y credenciales (prioridad: CLI > .env > reno_config.json)
    env_dotenv = _cargar_dotenv() or ENV_DOTENV
    cfg = cargar_configuracion_app()
    sb_url = (
        args.supabase_url
        or os.environ.get("SUPABASE_URL")
        or env_dotenv.get("SUPABASE_URL")
        or str(cfg.get("supabase_url", ""))
    ).strip().strip('"').strip("'").rstrip("/")

    bucket = (
        args.bucket
        or os.environ.get("SUPABASE_STORAGE_BUCKET")
        or env_dotenv.get("SUPABASE_STORAGE_BUCKET")
        or str(cfg.get("supabase_storage_bucket", "updates"))
    ).strip() or "updates"

    api_key = (
        args.service_key
        or os.environ.get("SUPABASE_SERVICE_KEY")
        or os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        or env_dotenv.get("SUPABASE_SERVICE_KEY")
        or env_dotenv.get("SUPABASE_SERVICE_ROLE_KEY")
        or str(cfg.get("supabase_service_role_key", ""))
        or str(cfg.get("supabase_anon_key", ""))
    ).strip().strip('"').strip("'")

    base_url = (
        args.base_url
        or os.environ.get("RENO_UPDATES_URL")
        or env_dotenv.get("RENO_UPDATES_URL")
        or ""
    ).strip().rstrip("/")
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

    # 5. Subida automática a Supabase Storage
    creds_cli = bool(args.supabase_url or args.service_key)
    service_env_presente = bool(
        os.environ.get("SUPABASE_SERVICE_KEY")
        or os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        or env_dotenv.get("SUPABASE_SERVICE_KEY")
        or env_dotenv.get("SUPABASE_SERVICE_ROLE_KEY")
    )
    creds_env_reales = _es_credencial_real(sb_url, api_key) and service_env_presente
    debe_subir = (args.subir or creds_cli or creds_env_reales) and not args.no_subir

    if debe_subir:
        if not _es_url_supabase_valida(sb_url):
            print("\n❌ Error: No se especificó una SUPABASE_URL válida ni en .env ni por --supabase-url.")
            sys.exit(1)

        if not _es_credencial_real(sb_url, api_key):
            print("\n❌ Error: Falta configurar SUPABASE_SERVICE_KEY (sb_secret_... o JWT service_role) en .env o --service-key.")
            sys.exit(1)

        prefijo = f"{api_key[:12]}..." if len(api_key) > 12 else "***"
        print(f"\n[*] Conectando con Supabase Storage ({sb_url}, bucket: {bucket}, key: {prefijo})...")

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
            base_public = f"{sb_url}/storage/v1/object/public/{bucket}"
            print(f"\n✅ Todos los archivos ({exitos}/{len(archivos_a_subir)}) se subieron exitosamente.")
            print(f"   - URL pública JSON: {base_public}/version.json")
            if partes_generadas:
                print(f"   - Partes subidas:   {len(partes_generadas)} fragmentos en {base_public}/")
            elif ruta_instalador and ruta_instalador.exists():
                print(f"   - URL pública EXE:  {base_public}/{ruta_instalador.name}")
        else:
            print(f"\n❌ Falló la subida completa ({exitos}/{len(archivos_a_subir)} subidos).")
            sys.exit(1)
    elif not args.no_subir:
        print(
            "\nℹ️  Nota: No se subió a Supabase porque no se detectaron credenciales reales en .env "
            "(SUPABASE_URL y SUPABASE_SERVICE_KEY)."
        )

    print("\n🎉 ¡Proceso de publicación completado con éxito!")


if __name__ == "__main__":
    main()
