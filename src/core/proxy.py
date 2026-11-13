"""Módulo de gestión de proxy corporativo, autodetección y transporte HTTP multi-tier para RENO."""

from __future__ import annotations

import base64
import json
import logging
import os
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Callable

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


def _limpiar_servidor_proxy(servidor: str) -> str:
    """Normaliza un string 'host:puerto' removiendo esquema y barra final."""
    s = (servidor or "").strip().rstrip("/")
    if "://" in s:
        s = s.split("://", 1)[1].rstrip("/")
    return s


def _normalizar_proxy_manual(config: dict | None) -> tuple[str, int, str, str, str]:
    """Retorna (host, port, user, pwd, auth_mode) normalizados desde config."""
    cfg = config or {}
    raw_host = _limpiar_servidor_proxy(str(cfg.get("proxy_host") or ""))
    port_cfg = int(cfg.get("proxy_port") or 8080)
    host = raw_host
    port = port_cfg

    if ":" in raw_host and not raw_host.startswith("["):
        h_part, p_part = raw_host.rsplit(":", 1)
        if p_part.isdigit():
            host = h_part.strip()
            if not cfg.get("proxy_port") or int(cfg.get("proxy_port") or 0) == 8080:
                port = int(p_part)

    user = str(cfg.get("proxy_user") or "").strip()
    pwd = str(cfg.get("proxy_pass") or "")
    auth_mode = str(cfg.get("proxy_auth") or AUTH_MANUAL_FIJO)
    return host, port, user, pwd, auth_mode


def construir_proxy_url_manual(config: dict | None, incluir_credenciales: bool = True) -> str:
    """Construye la URL HTTP de proxy manual si hay proxy_host configurado."""
    host, port, user, pwd, _ = _normalizar_proxy_manual(config)
    if not host:
        return ""
    if incluir_credenciales and user:
        user_enc = urllib.parse.quote(user, safe="")
        pwd_enc = urllib.parse.quote(pwd, safe="")
        return f"http://{user_enc}:{pwd_enc}@{host}:{port}"
    return f"http://{host}:{port}"


def _obtener_proxy_wininet() -> str:
    """Lee ProxyServer desde el registro de Windows si ProxyEnable está activo (sin esquema)."""
    if sys.platform != "win32":
        return ""
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        ) as key:
            try:
                enabled, _ = winreg.QueryValueEx(key, "ProxyEnable")
                server, _ = winreg.QueryValueEx(key, "ProxyServer")
            except FileNotFoundError:
                return ""
            if not enabled or not server:
                return ""
            srv_str = str(server).strip()
            if "=" in srv_str:
                mapa: dict[str, str] = {}
                for tramo in srv_str.split(";"):
                    if "=" in tramo:
                        proto, addr = tramo.split("=", 1)
                        mapa[proto.strip().lower()] = _limpiar_servidor_proxy(addr)
                return mapa.get("https") or mapa.get("http") or next(iter(mapa.values()), "")
            return _limpiar_servidor_proxy(srv_str)
    except Exception:
        return ""


def _configurar_proxy_qnam(nam: Any, config: dict | None) -> None:
    """Aplica la configuración de proxy directamente sobre una instancia QNetworkAccessManager."""
    from PySide6.QtNetwork import QNetworkProxy, QNetworkProxyFactory

    cfg = config or {}
    modo = cfg.get("proxy_mode", MODO_DIRECTO)
    if modo == MODO_DIRECTO:
        nam.setProxy(QNetworkProxy(QNetworkProxy.ProxyType.NoProxy))
        return

    if modo == MODO_MANUAL:
        host, port, user, pwd, auth_mode = _normalizar_proxy_manual(cfg)
        if host:
            if auth_mode == AUTH_WINDOWS_SSPI and not user:
                nam.setProxy(QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port))
            else:
                nam.setProxy(QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port, user, pwd))
        return

    if modo == MODO_WINDOWS_SSO:
        sso_server = _limpiar_servidor_proxy(str(cfg.get("proxy_sso_server") or ""))
        if sso_server:
            partes = sso_server.rsplit(":", 1)
            host = partes[0]
            port = int(partes[1]) if len(partes) > 1 and partes[1].isdigit() else 8080
            nam.setProxy(QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port))
        else:
            QNetworkProxyFactory.setUseSystemConfiguration(True)
            nam.setProxy(QNetworkProxy(QNetworkProxy.ProxyType.DefaultProxy))


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


_ULTIMO_PROXY_ENV_RENO: str | None = None


def aplicar_configuracion_proxy(config: dict) -> None:
    """Aplica la configuración de proxy al entorno del proceso y a QtNetwork."""
    global _ULTIMO_PROXY_ENV_RENO
    cfg = config or {}
    modo = cfg.get("proxy_mode", MODO_DIRECTO)
    claves_env = ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy")

    try:
        from PySide6.QtNetwork import QNetworkProxy, QNetworkProxyFactory
        has_qt_network = True
    except ImportError:
        has_qt_network = False

    if modo == MODO_DIRECTO:
        for k in claves_env:
            os.environ.pop(k, None)
        _ULTIMO_PROXY_ENV_RENO = None
        urllib.request.install_opener(urllib.request.build_opener(urllib.request.ProxyHandler({})))
        if has_qt_network:
            try:
                QNetworkProxyFactory.setUseSystemConfiguration(False)
                QNetworkProxy.setApplicationProxy(QNetworkProxy(QNetworkProxy.ProxyType.NoProxy))
            except Exception:
                pass
        return

    if modo == MODO_WINDOWS_SSO:
        sso_server = _limpiar_servidor_proxy(str(cfg.get("proxy_sso_server") or ""))
        if sso_server:
            sso_url = f"http://{sso_server}"
            for k in claves_env:
                os.environ[k] = sso_url
            _ULTIMO_PROXY_ENV_RENO = sso_url
            urllib.request.install_opener(
                urllib.request.build_opener(urllib.request.ProxyHandler({"http": sso_url, "https": sso_url}))
            )
        else:
            win_srv = _obtener_proxy_wininet()
            if win_srv:
                srv_url = f"http://{win_srv}"
                for k in claves_env:
                    os.environ[k] = srv_url
                _ULTIMO_PROXY_ENV_RENO = srv_url
                urllib.request.install_opener(
                    urllib.request.build_opener(urllib.request.ProxyHandler({"http": srv_url, "https": srv_url}))
                )
            else:
                if _ULTIMO_PROXY_ENV_RENO:
                    for k in claves_env:
                        if os.environ.get(k) == _ULTIMO_PROXY_ENV_RENO:
                            os.environ.pop(k, None)
                    _ULTIMO_PROXY_ENV_RENO = None
                urllib.request.install_opener(urllib.request.build_opener())

        if has_qt_network:
            try:
                if sso_server:
                    partes = sso_server.rsplit(":", 1)
                    host = partes[0]
                    port = int(partes[1]) if len(partes) > 1 and partes[1].isdigit() else 8080
                    proxy = QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port)
                    QNetworkProxy.setApplicationProxy(proxy)
                    QNetworkProxyFactory.setUseSystemConfiguration(False)
                else:
                    QNetworkProxyFactory.setUseSystemConfiguration(True)
            except Exception:
                pass
        return

    if modo == MODO_MANUAL:
        host, port, user, pwd, auth_mode = _normalizar_proxy_manual(cfg)
        if not host:
            return
        incluir_creds = bool(user) and auth_mode != AUTH_WINDOWS_SSPI
        proxy_url = construir_proxy_url_manual(cfg, incluir_credenciales=incluir_creds)

        for k in claves_env:
            os.environ[k] = proxy_url
        _ULTIMO_PROXY_ENV_RENO = proxy_url

        proxy_handler = urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
        opener = urllib.request.build_opener(proxy_handler)
        urllib.request.install_opener(opener)

        if has_qt_network:
            try:
                if auth_mode == AUTH_WINDOWS_SSPI and not user:
                    proxy = QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port)
                else:
                    proxy = QNetworkProxy(QNetworkProxy.ProxyType.HttpProxy, host, port, user, pwd)
                QNetworkProxy.setApplicationProxy(proxy)
                QNetworkProxyFactory.setUseSystemConfiguration(False)
            except Exception:
                pass



def _ejecutar_via_qnam(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    timeout_sec: int = 15,
    config: dict | None = None,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición HTTP mediante QNetworkAccessManager (usa SSPI/WinHTTP en Windows)."""
    from PySide6.QtCore import QByteArray, QEventLoop, QTimer, QUrl
    from PySide6.QtNetwork import QNetworkAccessManager, QNetworkRequest

    loop = QEventLoop()
    timer = QTimer()
    timer.setSingleShot(True)

    nam = QNetworkAccessManager()
    if config is not None:
        _configurar_proxy_qnam(nam, config)
    qurl = QUrl(url)
    req = QNetworkRequest(qurl)
    req.setAttribute(
        QNetworkRequest.Attribute.RedirectPolicyAttribute,
        QNetworkRequest.RedirectPolicy.NoLessSafeRedirectPolicy,
    )

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

    if not reply.isFinished():
        loop.exec()
    timer.stop()

    if timed_out:
        reply.deleteLater()
        raise TimeoutError(f"Tiempo de espera agotado ({timeout_sec}s) al conectar.")

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
    proxy_user: str = "",
    proxy_pass: str = "",
    timeout_sec: int = 15,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta una petición usando el objeto COM WinHttp.WinHttpRequest.5.1 de Windows."""
    import win32com.client  # type: ignore

    winhttp = win32com.client.Dispatch("WinHttp.WinHttpRequest.5.1")
    ms = timeout_sec * 1000
    winhttp.SetTimeouts(ms, ms, ms, ms)
    try:
        # WinHttpRequestOption_EnableRedirects = 6
        winhttp.Option[6] = True
    except Exception:
        pass

    winhttp.Open(method, url, False)

    # 0 = HTTPREQUEST_PROXYSETTING_DEFAULT, 2 = HTTPREQUEST_PROXYSETTING_PROXY
    proxy_clean = _limpiar_servidor_proxy(proxy_server)
    if proxy_clean:
        winhttp.SetProxy(2, proxy_clean)

    # AutoLogonPolicy: 0 = AutoLogonPolicy_Always (envía credenciales NTLM/Kerberos de Windows automáticamente)
    try:
        winhttp.SetAutoLogonPolicy(0)
    except Exception:
        pass

    if proxy_user:
        try:
            # 1 = HTTPREQUEST_SETCREDENTIALS_FOR_PROXY
            winhttp.SetCredentials(proxy_user, proxy_pass or "", 1)
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
    proxy_user: str = "",
    proxy_pass: str = "",
    timeout_sec: int = 15,
    output_file: str = "",
) -> tuple[int, bytes, dict[str, str]]:
    """Fallback mediante PowerShell System.Net.Http.HttpClient con SSO o credenciales explícitas."""
    proxy_clean = _limpiar_servidor_proxy(proxy_server)
    proxy_url_ps = f"http://{proxy_clean}" if proxy_clean else ""

    cfg_tmp_fd, cfg_tmp_path = tempfile.mkstemp(prefix="reno_ps_cfg_", suffix=".json")
    os.close(cfg_tmp_fd)
    in_tmp_path = ""

    try:
        if data is not None:
            in_fd, in_tmp_path = tempfile.mkstemp(prefix="reno_ps_in_", suffix=".bin")
            os.close(in_fd)
            Path(in_tmp_path).write_bytes(data)

        cfg_payload = {
            "url": url,
            "method": method,
            "headers": headers,
            "input_file": in_tmp_path,
            "proxy_url": proxy_url_ps,
            "proxy_user": proxy_user or "",
            "proxy_pass": proxy_pass or "",
            "timeout_sec": int(timeout_sec),
            "output_file": output_file or "",
        }
        Path(cfg_tmp_path).write_text(json.dumps(cfg_payload), encoding="utf-8")
        cfg_path_b64 = base64.b64encode(cfg_tmp_path.encode("utf-8")).decode("ascii")

        ps_script = f"""
$ErrorActionPreference = 'Stop'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls13
$cfgPath = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String('{cfg_path_b64}'))
$cfg = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
$handler = New-Object System.Net.Http.HttpClientHandler
$handler.AllowAutoRedirect = $true
$handler.UseDefaultCredentials = $true
$handler.DefaultProxyCredentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
if ($cfg.proxy_url) {{
    $p = New-Object System.Net.WebProxy([string]$cfg.proxy_url)
    if ($cfg.proxy_user) {{
        $p.Credentials = New-Object System.Net.NetworkCredential([string]$cfg.proxy_user, [string]$cfg.proxy_pass)
    }} else {{
        $p.UseDefaultCredentials = $true
    }}
    $handler.Proxy = $p
    $handler.UseProxy = $true
}} elseif ($cfg.proxy_user) {{
    $handler.DefaultProxyCredentials = New-Object System.Net.NetworkCredential([string]$cfg.proxy_user, [string]$cfg.proxy_pass)
}}
$client = New-Object System.Net.Http.HttpClient($handler)
$client.Timeout = [TimeSpan]::FromSeconds([int]$cfg.timeout_sec)

$req = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::new([string]$cfg.method), [string]$cfg.url)
foreach ($prop in $cfg.headers.PSObject.Properties) {{
    if ($prop.Name -ieq 'Content-Type') {{ continue }}
    $req.Headers.TryAddWithoutValidation($prop.Name, [string]$prop.Value) | Out-Null
}}

$inStreamReq = $null
try {{
    if ($cfg.input_file) {{
        $inStreamReq = [System.IO.File]::OpenRead([string]$cfg.input_file)
        $req.Content = New-Object System.Net.Http.StreamContent($inStreamReq)
        if ($cfg.headers.'Content-Type') {{
            $req.Content.Headers.ContentType = [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse([string]$cfg.headers.'Content-Type')
        }}
    }}

    $resp = $client.SendAsync($req, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    $code = [int]$resp.StatusCode
    if ($cfg.output_file) {{
        if ($code -ge 200 -and $code -lt 300) {{
            $inStream = $resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
            $outStream = [System.IO.File]::Create([string]$cfg.output_file)
            try {{
                $inStream.CopyTo($outStream)
            }} finally {{
                $outStream.Dispose()
                $inStream.Dispose()
            }}
        }}
        [PSCustomObject]@{{ StatusCode = $code; BodyB64 = '' }} | ConvertTo-Json -Compress
    }} else {{
        $respBytes = $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
        $respB64 = [System.Convert]::ToBase64String($respBytes)
        [PSCustomObject]@{{ StatusCode = $code; BodyB64 = $respB64 }} | ConvertTo-Json -Compress
    }}
}} finally {{
    if ($inStreamReq -ne $null) {{ $inStreamReq.Dispose() }}
}}
"""
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=timeout_sec + 10,
        )
        if proc.returncode != 0:
            raise ConnectionError(f"Fallo en PowerShell HttpClient: {proc.stderr.strip()}")

        stdout_txt = proc.stdout.strip()
        res = json.loads(stdout_txt.splitlines()[-1] if stdout_txt else "{}")
        status = int(res.get("StatusCode", 0))
        body = base64.b64decode(res["BodyB64"]) if res.get("BodyB64") else b""
        return status, body, {}
    finally:
        for p_tmp in (cfg_tmp_path, in_tmp_path):
            if p_tmp:
                try:
                    Path(p_tmp).unlink(missing_ok=True)
                except Exception:
                    pass


def _crear_opener_urllib_y_request(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    proxy_config: dict | None = None,
) -> tuple[urllib.request.OpenerDirector | None, urllib.request.Request]:
    """Construye Request y Opener de urllib según el modo de proxy."""
    req = urllib.request.Request(url, data=data, method=method)
    for k, v in headers.items():
        req.add_header(k, v)

    opener = None
    modo = (proxy_config or {}).get("proxy_mode", MODO_DIRECTO)
    if modo == MODO_MANUAL:
        host, port, user, pwd, auth_mode = _normalizar_proxy_manual(proxy_config)
        if host:
            incluir_creds = bool(user) and auth_mode != AUTH_WINDOWS_SSPI
            proxy_url = construir_proxy_url_manual(proxy_config, incluir_credenciales=incluir_creds)
            proxy_handler = urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
            if incluir_creds:
                auth_str = f"{user}:{pwd}"
                b64_auth = base64.b64encode(auth_str.encode("utf-8")).decode("ascii")
                req.add_header("Proxy-Authorization", f"Basic {b64_auth}")
            opener = urllib.request.build_opener(proxy_handler)
    elif modo == MODO_WINDOWS_SSO:
        sso_server = _limpiar_servidor_proxy(str((proxy_config or {}).get("proxy_sso_server") or ""))
        if not sso_server:
            sso_server = _obtener_proxy_wininet()
        if sso_server:
            sso_url = f"http://{sso_server}"
            opener = urllib.request.build_opener(
                urllib.request.ProxyHandler({"http": sso_url, "https": sso_url})
            )
    elif modo == MODO_DIRECTO:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    return opener, req


def _ejecutar_via_urllib(
    url: str,
    method: str,
    data: bytes | None,
    headers: dict[str, str],
    timeout_sec: int = 15,
    proxy_config: dict | None = None,
) -> tuple[int, bytes, dict[str, str]]:
    """Ejecuta petición HTTP usando urllib con soporte explícito de proxy y credenciales."""
    opener, req = _crear_opener_urllib_y_request(url, method, data, headers, proxy_config)

    try:
        resp_cm = opener.open(req, timeout=timeout_sec) if opener else urllib.request.urlopen(req, timeout=timeout_sec)
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
    """Ejecuta una petición HTTP/HTTPS respetando el modo de proxy configurado (multi-tier)."""
    cfg = dict(config or {})
    modo = cfg.get("proxy_mode", MODO_DIRECTO)
    hdrs = dict(headers or {})

    # Modo 1: Directo
    if modo == MODO_DIRECTO:
        code, body, resp_headers = _ejecutar_via_urllib(url, method, data, hdrs, timeout_sec, cfg)
        if code == 407:
            raise ProxyAuthError("HTTP 407: La red requiere autenticación en el proxy corporativo.")
        return code, body, resp_headers

    # Modo 2: Sesión Integrada de Windows (SSO / NTLM)
    if modo == MODO_WINDOWS_SSO:
        proxy_sso = _limpiar_servidor_proxy(str(cfg.get("proxy_sso_server") or ""))
        wininet_srv = _obtener_proxy_wininet() if not proxy_sso else ""
        proxy_efectivo = proxy_sso or wininet_srv

        # 1. Intentar con QNetworkAccessManager si hay QApplication
        try:
            from PySide6.QtWidgets import QApplication
            if QApplication.instance() is not None:
                code, body, resp_headers = _ejecutar_via_qnam(url, method, data, hdrs, timeout_sec, config=cfg)
                if code != 407 and code != 0:
                    return code, body, resp_headers
                logger.warning("QNAM devolvió código %s en Modo SSO; probando fallback nativo WinHTTP...", code)
        except Exception as e:
            logger.debug("Fallo en QNAM para SSO (%s), probando fallback nativo...", e)

        # 2. Fallback COM WinHttp + PowerShell HttpClient en Windows
        if sys.platform == "win32":
            for cand_srv in (proxy_efectivo, ""):
                try:
                    code, body, resp_headers = _ejecutar_via_winhttp_com(
                        url, method, data, hdrs, proxy_server=cand_srv, timeout_sec=timeout_sec
                    )
                    if code != 407 and code != 0:
                        return code, body, resp_headers
                except Exception as e:
                    logger.debug("WinHTTP COM (%s) falló: %s", cand_srv or "default", e)
                if not cand_srv:
                    break

            for cand_srv in (proxy_efectivo, ""):
                try:
                    code, body, resp_headers = _ejecutar_via_powershell(
                        url, method, data, hdrs, proxy_server=cand_srv, timeout_sec=timeout_sec
                    )
                    if code != 407 and code != 0:
                        return code, body, resp_headers
                except Exception as e:
                    logger.debug("PowerShell HttpClient (%s) falló: %s", cand_srv or "default", e)
                if not cand_srv:
                    break

        # 3. Fallback urllib con proxy de entorno / registro (útiles para proxies sin NTLM o en tests/CLI)
        try:
            code, body, resp_headers = _ejecutar_via_urllib(url, method, data, hdrs, timeout_sec, cfg)
            if code != 407:
                return code, body, resp_headers
        except ProxyAuthError:
            pass
        except OfflineError:
            if sys.platform != "win32":
                raise

        raise ProxyAuthError("HTTP 407: El proxy corporativo rechazó las credenciales de sesión de Windows.")

    # Modo 3: Proxy Manual
    if modo == MODO_MANUAL:
        host, port, user, pwd, auth_mode = _normalizar_proxy_manual(cfg)
        proxy_server = f"{host}:{port}" if host else ""
        err_offline: Exception | None = None
        code = 0
        body = b""
        resp_headers: dict[str, str] = {}

        if not (auth_mode == AUTH_WINDOWS_SSPI and not user):
            try:
                code, body, resp_headers = _ejecutar_via_urllib(url, method, data, hdrs, timeout_sec, cfg)
                if code != 407 and code != 0:
                    return code, body, resp_headers
            except ProxyAuthError as e:
                code = 407
                err_offline = e
            except OfflineError as e:
                err_offline = e
        else:
            code = 407

        # Intentar QNAM (soporta NTLM/Negotiate integrado de Windows con host/port manual)
        try:
            from PySide6.QtWidgets import QApplication
            if QApplication.instance() is not None and host:
                q_code, q_body, q_hdrs = _ejecutar_via_qnam(url, method, data, hdrs, timeout_sec, config=cfg)
                if q_code != 407 and q_code != 0:
                    return q_code, q_body, q_hdrs
        except Exception as e:
            logger.debug("QNAM en modo manual falló: %s", e)

        if sys.platform == "win32" and host:
            try:
                w_code, w_body, w_hdrs = _ejecutar_via_winhttp_com(
                    url,
                    method,
                    data,
                    hdrs,
                    proxy_server=proxy_server,
                    proxy_user="" if auth_mode == AUTH_WINDOWS_SSPI else user,
                    proxy_pass="" if auth_mode == AUTH_WINDOWS_SSPI else pwd,
                    timeout_sec=timeout_sec,
                )
                if w_code != 407 and w_code != 0:
                    return w_code, w_body, w_hdrs
            except Exception as e:
                logger.debug("Reintento WinHTTP manual falló: %s", e)

            try:
                p_code, p_body, p_hdrs = _ejecutar_via_powershell(
                    url,
                    method,
                    data,
                    hdrs,
                    proxy_server=proxy_server,
                    proxy_user="" if auth_mode == AUTH_WINDOWS_SSPI else user,
                    proxy_pass="" if auth_mode == AUTH_WINDOWS_SSPI else pwd,
                    timeout_sec=timeout_sec,
                )
                if p_code != 407 and p_code != 0:
                    return p_code, p_body, p_hdrs
            except Exception as e:
                logger.debug("Reintento PowerShell manual falló: %s", e)

        if code == 407 or isinstance(err_offline, ProxyAuthError):
            raise ProxyAuthError("HTTP 407: Autenticación de proxy requerida. Verifique usuario y contraseña.")
        if err_offline:
            raise err_offline
        return code, body, resp_headers

    raise ValueError(f"Modo de proxy desconocido: {modo}")


def _descargar_stream_via_qnam(
    url: str,
    archivo_out: Any,
    headers: dict[str, str],
    config: dict,
    on_chunk: Callable[[int, int], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
    timeout_sec: int = 60,
) -> tuple[int, int]:
    """Descarga un archivo en streaming mediante QNetworkAccessManager respetando proxy."""
    from PySide6.QtCore import QByteArray, QEventLoop, QTimer, QUrl
    from PySide6.QtNetwork import QNetworkAccessManager, QNetworkRequest

    loop = QEventLoop()
    timeout_timer = QTimer()
    timeout_timer.setSingleShot(True)
    cancel_timer = QTimer()
    cancel_timer.setInterval(150)

    nam = QNetworkAccessManager()
    _configurar_proxy_qnam(nam, config)
    req = QNetworkRequest(QUrl(url))
    req.setAttribute(
        QNetworkRequest.Attribute.RedirectPolicyAttribute,
        QNetworkRequest.RedirectPolicy.NoLessSafeRedirectPolicy,
    )
    for k, v in headers.items():
        req.setRawHeader(QByteArray(k.encode("utf-8")), QByteArray(v.encode("utf-8")))

    reply = nam.get(req)
    bytes_recibidos = 0
    total_esperado = 0
    timed_out = False
    cancelado = False

    def _on_meta() -> None:
        nonlocal total_esperado
        val = reply.header(QNetworkRequest.KnownHeaders.ContentLengthHeader)
        if val:
            try:
                total_esperado = int(val)
            except Exception:
                pass

    def _on_ready_read() -> None:
        nonlocal bytes_recibidos
        timeout_timer.start(timeout_sec * 1000)
        if should_cancel and should_cancel():
            _abort_cancel()
            return
        status_attr = reply.attribute(QNetworkRequest.Attribute.HttpStatusCodeAttribute)
        if status_attr is not None and int(status_attr) >= 400:
            return
        datos = bytes(reply.readAll().data())
        if datos:
            archivo_out.write(datos)
            bytes_recibidos += len(datos)
            if on_chunk:
                on_chunk(bytes_recibidos, total_esperado)

    def _abort_timeout() -> None:
        nonlocal timed_out
        timed_out = True
        reply.abort()
        loop.quit()

    def _abort_cancel() -> None:
        nonlocal cancelado
        cancelado = True
        reply.abort()
        loop.quit()

    def _check_cancel() -> None:
        if should_cancel and should_cancel():
            _abort_cancel()

    def _on_finished() -> None:
        nonlocal bytes_recibidos
        status_attr = reply.attribute(QNetworkRequest.Attribute.HttpStatusCodeAttribute)
        if status_attr is None or int(status_attr) < 400:
            rem = bytes(reply.readAll().data())
            if rem:
                archivo_out.write(rem)
                bytes_recibidos += len(rem)
                if on_chunk:
                    on_chunk(bytes_recibidos, total_esperado)
        loop.quit()

    reply.metaDataChanged.connect(_on_meta)
    reply.readyRead.connect(_on_ready_read)
    reply.finished.connect(_on_finished)
    timeout_timer.timeout.connect(_abort_timeout)
    cancel_timer.timeout.connect(_check_cancel)

    timeout_timer.start(timeout_sec * 1000)
    if should_cancel:
        cancel_timer.start()

    if not reply.isFinished():
        loop.exec()

    timeout_timer.stop()
    cancel_timer.stop()

    status_code = int(reply.attribute(QNetworkRequest.Attribute.HttpStatusCodeAttribute) or 0)
    reply.deleteLater()

    if cancelado:
        raise InterruptedError("Descarga cancelada por el usuario.")
    if timed_out:
        raise TimeoutError(f"Tiempo de espera agotado ({timeout_sec}s) durante la descarga.")
    return status_code, bytes_recibidos


def descargar_archivo_con_proxy(
    config: dict | None,
    url: str,
    archivo_out: Any,
    headers: dict[str, str] | None = None,
    on_chunk: Callable[[int, int], None] | None = None,
    should_cancel: Callable[[], bool] | None = None,
    timeout_sec: int = 60,
) -> int:
    """Descarga por streaming un binario respetando proxy directo, manual o Windows SSO.

    Retorna la cantidad de bytes escritos en ``archivo_out`` para este segmento.
    """
    cfg = dict(config or {})
    modo = cfg.get("proxy_mode", MODO_DIRECTO)
    hdrs = dict(headers or {})
    start_pos = archivo_out.tell()

    def _reset_stream() -> None:
        archivo_out.seek(start_pos)
        archivo_out.truncate()

    def _stream_urllib() -> tuple[int, int]:
        opener, req = _crear_opener_urllib_y_request(url, "GET", None, hdrs, cfg)
        try:
            resp_cm = opener.open(req, timeout=timeout_sec) if opener else urllib.request.urlopen(req, timeout=timeout_sec)
            with resp_cm as resp:
                code = int(resp.status)
                total = int(resp.headers.get("Content-Length") or 0)
                escrito = 0
                while True:
                    if should_cancel and should_cancel():
                        raise InterruptedError("Descarga cancelada por el usuario.")
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    archivo_out.write(chunk)
                    escrito += len(chunk)
                    if on_chunk:
                        on_chunk(escrito, total)
                return code, escrito
        except urllib.error.HTTPError as e:
            return int(e.code), 0
        except urllib.error.URLError as e:
            if "407" in str(e):
                raise ProxyAuthError("407 Proxy Authentication Required")
            raise OfflineError(f"Error de red o sin conexión: {e.reason}")

    def _stream_powershell(proxy_srv: str = "", p_user: str = "", p_pass: str = "") -> tuple[int, int]:
        tmp_fd, tmp_path = tempfile.mkstemp(prefix="reno_dl_", suffix=".part")
        os.close(tmp_fd)
        try:
            code, _, _ = _ejecutar_via_powershell(
                url=url,
                method="GET",
                data=None,
                headers=hdrs,
                proxy_server=proxy_srv,
                proxy_user=p_user,
                proxy_pass=p_pass,
                timeout_sec=max(timeout_sec, 120),
                output_file=tmp_path,
            )
            if not (200 <= code < 300):
                return code, 0
            total = Path(tmp_path).stat().st_size
            escrito = 0
            with open(tmp_path, "rb") as f_in:
                while True:
                    if should_cancel and should_cancel():
                        raise InterruptedError("Descarga cancelada por el usuario.")
                    chunk = f_in.read(65536)
                    if not chunk:
                        break
                    archivo_out.write(chunk)
                    escrito += len(chunk)
                    if on_chunk:
                        on_chunk(escrito, total)
            return code, escrito
        finally:
            try:
                Path(tmp_path).unlink(missing_ok=True)
            except Exception:
                pass

    # Modo 1: Directo
    if modo == MODO_DIRECTO:
        code, escrito = _stream_urllib()
        if code == 407:
            raise ProxyAuthError("HTTP 407: La red requiere autenticación en el proxy corporativo.")
        if not (200 <= code < 300):
            raise ConnectionError(f"El servidor respondió HTTP {code} al descargar {url}")
        return escrito

    # Modo 2: Windows SSO
    if modo == MODO_WINDOWS_SSO:
        proxy_sso = _limpiar_servidor_proxy(str(cfg.get("proxy_sso_server") or ""))
        proxy_efectivo = proxy_sso or (_obtener_proxy_wininet() if not proxy_sso else "")

        try:
            from PySide6.QtWidgets import QApplication
            if QApplication.instance() is not None:
                _reset_stream()
                code, escrito = _descargar_stream_via_qnam(
                    url, archivo_out, hdrs, cfg, on_chunk=on_chunk, should_cancel=should_cancel, timeout_sec=timeout_sec
                )
                if 200 <= code < 300:
                    return escrito
                if code not in (0, 407):
                    raise ConnectionError(f"El servidor respondió HTTP {code} al descargar {url}")
        except InterruptedError:
            raise
        except ConnectionError:
            raise
        except Exception as e:
            logger.debug("Descarga QNAM en modo SSO falló (%s); probando fallback...", e)

        if sys.platform == "win32":
            for cand_srv in (proxy_efectivo, ""):
                try:
                    _reset_stream()
                    code, body, _ = _ejecutar_via_winhttp_com(
                        url, "GET", None, hdrs, proxy_server=cand_srv, timeout_sec=max(timeout_sec, 120)
                    )
                    if 200 <= code < 300:
                        archivo_out.write(body)
                        if on_chunk:
                            on_chunk(len(body), len(body))
                        return len(body)
                except Exception as e:
                    logger.debug("Descarga WinHTTP (%s) falló: %s", cand_srv or "default", e)
                if not cand_srv:
                    break

            for cand_srv in (proxy_efectivo, ""):
                try:
                    _reset_stream()
                    code, escrito = _stream_powershell(proxy_srv=cand_srv)
                    if 200 <= code < 300:
                        return escrito
                except InterruptedError:
                    raise
                except Exception as e:
                    logger.debug("Descarga PowerShell (%s) falló: %s", cand_srv or "default", e)
                if not cand_srv:
                    break

        _reset_stream()
        code, escrito = _stream_urllib()
        if code == 407:
            raise ProxyAuthError("HTTP 407: El proxy corporativo rechazó las credenciales de Windows.")
        if not (200 <= code < 300):
            raise ConnectionError(f"El servidor respondió HTTP {code} al descargar {url}")
        return escrito

    # Modo 3: Proxy Manual
    if modo == MODO_MANUAL:
        host, port, user, pwd, auth_mode = _normalizar_proxy_manual(cfg)
        proxy_server = f"{host}:{port}" if host else ""
        err_offline: Exception | None = None
        last_code = 0

        if not (auth_mode == AUTH_WINDOWS_SSPI and not user):
            try:
                _reset_stream()
                code, escrito = _stream_urllib()
                if 200 <= code < 300:
                    return escrito
                last_code = code
                if code != 407 and code != 0:
                    raise ConnectionError(f"El servidor respondió HTTP {code} al descargar {url}")
            except InterruptedError:
                raise
            except ProxyAuthError as e:
                last_code = 407
                err_offline = e
            except OfflineError as e:
                err_offline = e
        else:
            last_code = 407

        try:
            from PySide6.QtWidgets import QApplication
            if QApplication.instance() is not None and host:
                _reset_stream()
                code, escrito = _descargar_stream_via_qnam(
                    url, archivo_out, hdrs, cfg, on_chunk=on_chunk, should_cancel=should_cancel, timeout_sec=timeout_sec
                )
                if 200 <= code < 300:
                    return escrito
                if code not in (0, 407):
                    raise ConnectionError(f"El servidor respondió HTTP {code} al descargar {url}")
        except InterruptedError:
            raise
        except ConnectionError:
            raise
        except Exception as e:
            logger.debug("Descarga QNAM manual falló (%s); probando fallback nativo...", e)

        if sys.platform == "win32" and host:
            eff_user = "" if auth_mode == AUTH_WINDOWS_SSPI else user
            eff_pass = "" if auth_mode == AUTH_WINDOWS_SSPI else pwd
            try:
                _reset_stream()
                code, body, _ = _ejecutar_via_winhttp_com(
                    url,
                    "GET",
                    None,
                    hdrs,
                    proxy_server=proxy_server,
                    proxy_user=eff_user,
                    proxy_pass=eff_pass,
                    timeout_sec=max(timeout_sec, 120),
                )
                if 200 <= code < 300:
                    archivo_out.write(body)
                    if on_chunk:
                        on_chunk(len(body), len(body))
                    return len(body)
            except Exception as e:
                logger.debug("Descarga WinHTTP manual falló: %s", e)

            try:
                _reset_stream()
                code, escrito = _stream_powershell(proxy_srv=proxy_server, p_user=eff_user, p_pass=eff_pass)
                if 200 <= code < 300:
                    return escrito
            except InterruptedError:
                raise
            except Exception as e:
                logger.debug("Descarga PowerShell manual falló: %s", e)

        if last_code == 407 or isinstance(err_offline, ProxyAuthError):
            raise ProxyAuthError("HTTP 407: Autenticación de proxy requerida. Verifique usuario y contraseña.")
        if err_offline:
            raise err_offline
        raise ConnectionError(f"No se pudo descargar {url} (HTTP {last_code or 'sin respuesta'})")

    raise ValueError(f"Modo de proxy desconocido: {modo}")


def construir_headers_supabase(api_key: str, extra: dict[str, str] | None = None) -> dict[str, str]:
    """Construye cabeceras HTTP compatibles con Publishable/Secret keys (sb_*) y JWT legacy."""
    key = (api_key or "").strip().strip('"').strip("'")
    headers: dict[str, str] = {"apikey": key}
    # Las Publishable y Secret keys modernas no deben enviarse como Authorization: Bearer
    # porque el API Gateway las rechaza ('Secret API key required' / 'Invalid Compact JWS').
    if key and not (key.startswith("sb_publishable_") or key.startswith("sb_secret_")):
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
