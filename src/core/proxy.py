"""Módulo de gestión de proxy corporativo, autodetección y transporte HTTP multi-tier para RENO."""

from __future__ import annotations

import base64
import json
import logging
import os
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

logger = logging.getLogger("reno.proxy")

MODO_DIRECTO = "direct"
MODO_WINDOWS_SSO = "windows_sso"
MODO_MANUAL = "manual"

AUTH_MANUAL_FIJO = "manual_fijo"
AUTH_SESION_RENO = "sesion_reno"
AUTH_WINDOWS_SSPI = "windows_sspi"

OPCIONES_MODO = [
    (MODO_DIRECTO, "Sin Proxy (Conexión Directa a Internet)"),
    (MODO_MANUAL, "Proxy Configurado (IP / Host y Puerto)"),
    (MODO_WINDOWS_SSO, "Detección Automática de Windows (SSO / NTLM / WPAD)"),
]

OPCIONES_AUTH_PROXY = [
    (AUTH_MANUAL_FIJO, "Usar credenciales manuales configuradas (Usuario y Contraseña fijos)"),
    (AUTH_SESION_RENO, "Usar automáticamente el usuario y contraseña de la sesión actual de RENO"),
    (AUTH_WINDOWS_SSPI, "Intentar usar credencial integrada de Windows / Outlook (win32cred / SSPI)"),
]


class ProxyAuthError(ConnectionError):
    """Excepción específica cuando el proxy corporativo responde HTTP 407."""
    pass


class OfflineError(ConnectionError):
    """Excepción específica cuando no hay conexión de red o DNS."""
    pass


def detectar_proxy_sistema(config: dict | None = None) -> str:
    """Detecta la configuración de proxy del sistema en Windows o desde variables de entorno."""
    if config and config.get("proxy_mode") == MODO_WINDOWS_SSO and config.get("proxy_sso_server"):
        return f"{config['proxy_sso_server'].strip()} (Configurado en SSO)"

    # 1. Registro de Windows (Internet Settings)
    if sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
            )
            try:
                enabled, _ = winreg.QueryValueEx(key, "ProxyEnable")
                server, _ = winreg.QueryValueEx(key, "ProxyServer")
            except FileNotFoundError:
                enabled, server = 0, ""

            try:
                autoconfig_url, _ = winreg.QueryValueEx(key, "AutoConfigURL")
            except FileNotFoundError:
                autoconfig_url = ""
            winreg.CloseKey(key)

            if autoconfig_url:
                return f"PAC: {autoconfig_url}"
            if enabled and server:
                return f"{server} (Registro de Windows)"
        except Exception as e:
            logger.debug("Error leyendo registro de Windows: %s", e)

    # 2. urllib.request.getproxies()
    proxies = urllib.request.getproxies()
    if proxies:
        return ", ".join(f"{k}://{v}" for k, v in proxies.items())

    return "Ninguno (Conexión Directa)"


def aplicar_configuracion_proxy(config: dict) -> None:
    """Aplica la configuración de proxy al entorno del proceso y a QtNetwork."""
    modo = config.get("proxy_mode", MODO_DIRECTO)
    claves_env = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")

    try:
        from PySide6.QtNetwork import QNetworkProxy, QNetworkProxyFactory
        has_qt_network = True
    except ImportError:
        has_qt_network = False

    if modo == MODO_DIRECTO:
        for k in claves_env:
            os.environ.pop(k, None)
        urllib.request.install_opener(urllib.request.build_opener(urllib.request.ProxyHandler({})))
        if has_qt_network:
            QNetworkProxyFactory.setUseSystemConfiguration(False)
            QNetworkProxy.setApplicationProxy(QNetworkProxy(QNetworkProxy.ProxyType.NoProxy))
        return

    if modo == MODO_WINDOWS_SSO:
        sso_server = (config.get("proxy_sso_server") or "").strip()
        if has_qt_network:
            if sso_server:
                partes = sso_server.split(":")
                host = partes[0]
                port = int(partes[1]) if len(partes) > 1 and partes[1].isdigit() else 8080
                proxy = QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port)
                QNetworkProxy.setApplicationProxy(proxy)
                QNetworkProxyFactory.setUseSystemConfiguration(False)
            else:
                QNetworkProxyFactory.setUseSystemConfiguration(True)
        return

    if modo == MODO_MANUAL:
        host = config.get("proxy_host", "").strip()
        port = int(config.get("proxy_port", 8080) or 8080)
        user = config.get("proxy_user", "").strip()
        pwd = config.get("proxy_pass", "")

        user_enc = urllib.parse.quote(user, safe="")
        pwd_enc = urllib.parse.quote(pwd, safe="")

        if user_enc:
            proxy_url = f"http://{user_enc}:{pwd_enc}@{host}:{port}"
        else:
            proxy_url = f"http://{host}:{port}"

        for k in claves_env:
            os.environ[k] = proxy_url

        proxy_handler = urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
        opener = urllib.request.build_opener(proxy_handler)
        urllib.request.install_opener(opener)

        if has_qt_network and host:
            proxy = QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port, user, pwd)
            QNetworkProxy.setApplicationProxy(proxy)
            QNetworkProxyFactory.setUseSystemConfiguration(False)


def _ejecutar_via_qnam(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición HTTP mediante QNetworkAccessManager (usa SSPI/WinHTTP en Windows)."""
    from PySide6.QtCore import QByteArray, QEventLoop, QTimer, QUrl
    from PySide6.QtNetwork import QNetworkAccessManager, QNetworkRequest

    loop = QEventLoop()
    timer = QTimer()
    timer.setSingleShot(True)

    nam = QNetworkAccessManager()
    qurl = QUrl(url)
    req = QNetworkRequest(qurl)

    for k, v in headers.items():
        req.setRawHeader(QByteArray(k.encode("utf-8")), QByteArray(v.encode("utf-8")))

    reply = None
    if method == "GET":
        reply = nam.get(req)
    elif method == "POST":
        reply = nam.post(req, QByteArray(data if data is not None else b""))
    elif method == "PATCH":
        reply = nam.sendCustomRequest(req, QByteArray(b"PATCH"), QByteArray(data if data is not None else b""))
    elif method == "DELETE":
        reply = nam.deleteResource(req)
    else:
        reply = nam.sendCustomRequest(req, QByteArray(method.encode("utf-8")), QByteArray(data if data is not None else b""))

    timed_out = False

    def on_timeout():
        nonlocal timed_out
        timed_out = True
        if reply:
            reply.abort()
        loop.quit()

    timer.timeout.connect(on_timeout)
    reply.finished.connect(loop.quit)
    timer.start(timeout_sec * 1000)

    loop.exec()

    if timed_out:
        raise TimeoutError(f"Tiempo de espera agotado ({timeout_sec}s) al conectar con Supabase.")

    status_code = reply.attribute(QNetworkRequest.Attribute.HttpStatusCodeAttribute) or 0
    body = bytes(reply.readAll().data())

    resp_headers = {}
    for hname in reply.rawHeaderList():
        resp_headers[bytes(hname.data()).decode("utf-8", "ignore").lower()] = bytes(
            reply.rawHeader(hname).data()
        ).decode("utf-8", "ignore")

    reply.deleteLater()
    return status_code, body, resp_headers


def _ejecutar_via_winhttp_com(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    proxy_server: str = "",
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición usando el objeto COM WinHttp.WinHttpRequest.5.1 de Windows."""
    import win32com.client  # type: ignore

    winhttp = win32com.client.Dispatch("WinHttp.WinHttpRequest.5.1")
    ms = timeout_sec * 1000
    winhttp.SetTimeouts(ms, ms, ms, ms)

    winhttp.Open(method, url, False)

    # 0 = HTTPREQUEST_PROXYSETTING_DEFAULT, 2 = HTTPREQUEST_PROXYSETTING_PROXY
    if proxy_server:
        winhttp.SetProxy(2, proxy_server)

    # AutoLogonPolicy: 0 = AutoLogonPolicy_Always (envía credenciales NTLM/Kerberos de Windows automáticamente)
    try:
        winhttp.SetAutoLogonPolicy(0)
    except Exception:
        pass

    for k, v in headers.items():
        winhttp.SetRequestHeader(k, v)

    if data is not None:
        winhttp.Send(data)
    else:
        winhttp.Send()

    status = int(winhttp.Status)
    try:
        raw_body = winhttp.ResponseBody
        body = bytes(raw_body) if raw_body is not None else b""
    except Exception:
        body = winhttp.ResponseText.encode("utf-8", "ignore")

    raw_headers = winhttp.GetAllResponseHeaders()
    resp_headers = {}
    for line in raw_headers.splitlines():
        if ":" in line:
            hk, hv = line.split(":", 1)
            resp_headers[hk.strip().lower()] = hv.strip()

    return status, body, resp_headers


def _ejecutar_via_powershell(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    proxy_server: str = "",
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Fallback mediante PowerShell System.Net.Http.HttpClient usando credenciales de sesión integradas."""
    headers_json = json.dumps(headers)
    body_b64 = base64.b64encode(data).decode("ascii") if data is not None else ""

    ps_script = f"""
$ErrorActionPreference = 'Stop'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
$handler = New-Object System.Net.Http.HttpClientHandler
$handler.UseDefaultCredentials = $true
$handler.DefaultProxyCredentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
if ('{proxy_server}') {{
    $handler.Proxy = New-Object System.Net.WebProxy('{proxy_server}')
}}
$client = New-Object System.Net.Http.HttpClient($handler)
$client.Timeout = [TimeSpan]::FromSeconds({timeout_sec})

$req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::new('{method}'), '{url}')
$hdr = ConvertFrom-Json @'
{headers_json}
'@
foreach ($prop in $hdr.PSObject.Properties) {{
    if ($prop.Name -ieq 'Content-Type') {{
        continue
    }}
    $req.Headers.TryAddWithoutValidation($prop.Name, [string]$prop.Value) | Out-Null
}}

if ('{body_b64}') {{
    $rawBytes = [System.Convert]::FromBase64String('{body_b64}')
    $req.Content = New-Object System.Net.Http.ByteArrayContent($rawBytes, 0, $rawBytes.Length)
    if ($hdr.'Content-Type') {{
        $req.Content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse($hdr.'Content-Type')
    }}
}}

$resp = $client.SendAsync($req).GetAwaiter().GetResult()
$code = [int]$resp.StatusCode
$respBytes = $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
$respB64 = [System.Convert]::ToBase64String($respBytes)

[PSCustomObject]@{{
    StatusCode = $code
    BodyB64 = $respB64
}} | ConvertTo-Json -Compress
"""
    proc = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
        capture_output=True,
        text=True,
        timeout=timeout_sec + 5,
    )
    if proc.returncode != 0:
        raise ConnectionError(f"Fallo en PowerShell HttpClient: {proc.stderr.strip()}")

    res = json.loads(proc.stdout.strip())
    status = int(res["StatusCode"])
    body = base64.b64decode(res["BodyB64"]) if res.get("BodyB64") else b""
    return status, body, {}


def _ejecutar_via_urllib(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    timeout_sec: int = 15,
    proxy_config: dict | None = None,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta petición HTTP usando urllib con soporte explícito de proxy y credenciales."""
    req = urllib.request.Request(url, data=data, method=method)
    for k, v in headers.items():
        req.add_header(k, v)

    opener = None
    if proxy_config and proxy_config.get("proxy_mode") == MODO_MANUAL:
        host = proxy_config.get("proxy_host", "").strip()
        port = int(proxy_config.get("proxy_port", 8080) or 8080)
        user = proxy_config.get("proxy_user", "").strip()
        pwd = proxy_config.get("proxy_pass", "")
        if host:
            user_enc = urllib.parse.quote(user, safe="")
            pwd_enc = urllib.parse.quote(pwd, safe="")
            proxy_url = f"http://{user_enc}:{pwd_enc}@{host}:{port}" if user_enc else f"http://{host}:{port}"
            proxy_handler = urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
            handlers = [proxy_handler]
            if user:
                auth_str = f"{user}:{pwd}"
                b64_auth = base64.b64encode(auth_str.encode("utf-8")).decode("ascii")
                req.add_header("Proxy-Authorization", f"Basic {b64_auth}")
            opener = urllib.request.build_opener(*handlers)
    elif proxy_config and proxy_config.get("proxy_mode") == MODO_DIRECTO:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    try:
        if opener:
            resp_cm = opener.open(req, timeout=timeout_sec)
        else:
            resp_cm = urllib.request.urlopen(req, timeout=timeout_sec)

        with resp_cm as resp:
            code = resp.status
            body = resp.read()
            resp_headers = {k.lower(): v for k, v in resp.headers.items()}
            return code, body, resp_headers
    except urllib.error.HTTPError as e:
        body = e.read()
        resp_headers = {k.lower(): v for k, v in e.headers.items()} if e.headers else {}
        return e.code, body, resp_headers
    except urllib.error.URLError as e:
        if "407" in str(e):
            raise ProxyAuthError("407 Proxy Authentication Required")
        raise OfflineError(f"Error de red o sin conexión: {e.reason}")


def ejecutar_request_supabase(
    config: dict,
    url: str,
    method: str = "GET",
    data: bytes | None = None,
    headers: dict[str, str] | None = None,
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición contra Supabase según el modo de proxy configurado.
    
    Aplica auto-fallback en caso de que Modo 2 o 3 requieran negociación de protocolo alternativa.
    """
    modo = config.get("proxy_mode", MODO_DIRECTO)
    hdrs = dict(headers or {})

    # Modo 1: Directo
    if modo == MODO_DIRECTO:
        code, body, resp_headers = _ejecutar_via_urllib(url, method, data, hdrs, timeout_sec, config)
        if code == 407:
            raise ProxyAuthError("HTTP 407: La red requiere autenticación en el proxy corporativo.")
        return code, body, resp_headers

    # Modo 2: Sesión Integrada de Windows (SSO / NTLM)
    if modo == MODO_WINDOWS_SSO:
        proxy_sso = config.get("proxy_sso_server", "").strip()

        # 1. Intentar con QNetworkAccessManager si hay QApplication
        try:
            from PySide6.QtWidgets import QApplication
            if QApplication.instance() is not None:
                aplicar_configuracion_proxy(config)
                code, body, resp_headers = _ejecutar_via_qnam(url, method, data, hdrs, timeout_sec)
                if code != 407 and code != 0:
                    return code, body, resp_headers
                logger.warning("QNAM devolvió código %s en Modo SSO; probando fallback nativo WinHTTP...", code)
        except Exception as e:
            logger.debug("Fallo en QNAM para SSO (%s), probando fallback nativo...", e)

        # 2. Fallback COM WinHttp
        if sys.platform == "win32":
            try:
                code, body, resp_headers = _ejecutar_via_winhttp_com(url, method, data, hdrs, proxy_sso, timeout_sec)
                if code != 407:
                    return code, body, resp_headers
                logger.warning("WinHTTP COM devolvió 407 en Modo SSO; probando fallback PowerShell...")
            except Exception as e:
                logger.debug("WinHTTP COM falló (%s); probando PowerShell...", e)

            # 3. Fallback PowerShell HttpClient
            try:
                code, body, resp_headers = _ejecutar_via_powershell(url, method, data, hdrs, proxy_sso, timeout_sec)
                if code != 407:
                    return code, body, resp_headers
            except Exception as e:
                logger.error("Todos los canales de SSO fallaron: %s", e)

        raise ProxyAuthError("HTTP 407: El proxy corporativo rechazó las credenciales de sesión de Windows.")

    # Modo 3: Proxy Manual
    if modo == MODO_MANUAL:
        try:
            code, body, resp_headers = _ejecutar_via_urllib(url, method, data, hdrs, timeout_sec, config)
        except ProxyAuthError:
            code = 407
            body = b""
            resp_headers = {}

        if code == 407:
            host = config.get("proxy_host", "").strip()
            port = int(config.get("proxy_port", 8080) or 8080)
            user = config.get("proxy_user", "").strip()
            pwd = config.get("proxy_pass", "")
            proxy_server = f"{host}:{port}" if host else ""

            if sys.platform == "win32" and host:
                try:
                    import win32com.client
                    winhttp = win32com.client.Dispatch("WinHttp.WinHttpRequest.5.1")
                    ms = timeout_sec * 1000
                    winhttp.SetTimeouts(ms, ms, ms, ms)
                    winhttp.Open(method, url, False)
                    winhttp.SetProxy(2, proxy_server)
                    if user:
                        winhttp.SetCredentials(user, pwd, 1)
                    for k, v in hdrs.items():
                        winhttp.SetRequestHeader(k, v)
                    if data is not None:
                        winhttp.Send(data)
                    else:
                        winhttp.Send()
                    res_status = int(winhttp.Status)
                    if res_status != 407:
                        return res_status, bytes(winhttp.ResponseBody), {}
                except Exception as e:
                    logger.debug("Reintento WinHTTP manual falló: %s", e)

            raise ProxyAuthError("HTTP 407: Autenticación de proxy requerida. Verifique usuario y contraseña.")

        return code, body, resp_headers

    raise ValueError(f"Modo de proxy desconocido: {modo}")


def construir_headers_supabase(api_key: str, extra: dict[str, str] | None = None) -> dict[str, str]:
    """Construye cabeceras HTTP compatibles con Publishable keys (sb_publishable_*) y JWT legacy."""
    key = (api_key or "").strip()
    headers: dict[str, str] = {"apikey": key}
    # Las Publishable keys de Supabase no deben enviarse como Authorization: Bearer
    # porque el API Gateway las rechaza con HTTP 401 ('Secret API key required').
    if key and not key.startswith("sb_publishable_"):
        headers["Authorization"] = f"Bearer {key}"
    if extra:
        headers.update(extra)
    return headers


def probar_conexion_supabase(config: dict) -> dict[str, Any]:
    """Prueba en vivo la conexión contra el endpoint de Supabase usando la configuración suministrada."""
    supabase_url = (config.get("supabase_url") or "").strip().rstrip("/")
    api_key = (config.get("supabase_anon_key") or "").strip()
    modo = config.get("proxy_mode", MODO_DIRECTO)
    proxy_det = detectar_proxy_sistema(config)

    modo_nombre = {
        MODO_DIRECTO: "Directo",
        MODO_WINDOWS_SSO: "Sesión Windows NTLM/WinHTTP (SSO)",
        MODO_MANUAL: "Proxy Manual",
    }.get(modo, modo)

    if not supabase_url or not api_key:
        return {
            "ok": False,
            "modo_usado": modo_nombre,
            "proxy_detectado": proxy_det,
            "status_code": 0,
            "detalle": "Falta configurar la URL de Supabase y/o la Publishable Key.",
        }

    from ..data.sync.supabase_sync import sanitizar_supabase_url
    clean_url = sanitizar_supabase_url(supabase_url)
    # Consultar una tabla real en lugar de /rest/v1/ (el esquema raíz OpenAPI requiere secret key
    # cuando se usa sb_publishable_*, mientras que las tablas funcionan con Publishable Key).
    test_url = f"{clean_url}/rest/v1/reno_dotaciones?select=id&limit=1"
    headers = construir_headers_supabase(api_key, {"Accept": "application/json"})

    try:
        status, body, _ = ejecutar_request_supabase(
            config=config,
            url=test_url,
            method="GET",
            headers=headers,
            timeout_sec=12,
        )

        if 200 <= status < 300:
            return {
                "ok": True,
                "modo_usado": modo_nombre,
                "proxy_detectado": proxy_det,
                "status_code": status,
                "detalle": f"HTTP {status} OK — Conexión exitosa con Supabase.",
            }
        elif status == 404 or b"PGRST205" in body:
            return {
                "ok": True,
                "modo_usado": modo_nombre,
                "proxy_detectado": proxy_det,
                "status_code": status,
                "detalle": (
                    f"Conexión autenticada con Supabase (HTTP {status}), pero falta ejecutar "
                    "el script 'supabase_reno_schema.sql' en el SQL Editor para crear las tablas reno_*."
                ),
            }
        elif status == 407:
            return {
                "ok": False,
                "modo_usado": modo_nombre,
                "proxy_detectado": proxy_det,
                "status_code": 407,
                "detalle": "HTTP 407 Proxy Authentication Required — El proxy corporativo rechazó la autenticación.",
            }
        else:
            txt_err = body.decode("utf-8", "ignore")[:200]
            return {
                "ok": False,
                "modo_usado": modo_nombre,
                "proxy_detectado": proxy_det,
                "status_code": status,
                "detalle": f"HTTP {status} recibido desde Supabase: {txt_err}",
            }
    except ProxyAuthError as e:
        return {
            "ok": False,
            "modo_usado": modo_nombre,
            "proxy_detectado": proxy_det,
            "status_code": 407,
            "detalle": f"HTTP 407 Proxy Authentication Required: {e}",
        }
    except OfflineError as e:
        return {
            "ok": False,
            "modo_usado": modo_nombre,
            "proxy_detectado": proxy_det,
            "status_code": 0,
            "detalle": f"Error de red / Servidor no alcanzable: {e}",
        }
    except Exception as e:
        return {
            "ok": False,
            "modo_usado": modo_nombre,
            "proxy_detectado": proxy_det,
            "status_code": 0,
            "detalle": f"Error inesperado al probar conexión: {e}",
        }
