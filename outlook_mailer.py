"""Envío de informes usando el Outlook clásico instalado en Windows."""

import uuid

OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_PERMANENT_DELETE = 2


def enviar_informe_outlook(destinatarios, asunto, cuerpo):
    if not destinatarios:
        raise RuntimeError("No hay destinatarios activos configurados para informes.")
    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError("Falta la dependencia pywin32 para enviar mediante Outlook.") from error

    outlook = win32com.client.Dispatch("Outlook.Application")
    mensaje = outlook.CreateItem(0)
    mensaje.To = ";".join(destinatarios)
    mensaje.Subject = asunto
    mensaje.Body = cuerpo
    mensaje.Send()


def enviar_sin_rastro(destinatarios, asunto, cuerpo):
    """Envía un correo y elimina permanentemente la copia de Enviados/Eliminados.

    Usa un token UUID en UserProperties para identificar el mensaje recién
    enviado y borrarlo con ``olPermanentDelete``.  Si no se encuentra en
    Enviados tras el envío, se emite un warning pero no se levanta excepción
    (el correo ya fue entregado al destinatario).
    """
    if not destinatarios:
        raise RuntimeError("No hay destinatarios configurados.")
    try:
        import win32com.client
    except ImportError as error:
        raise RuntimeError("Falta la dependencia pywin32 para enviar mediante Outlook.") from error

    token = uuid.uuid4().hex[:12]

    outlook = win32com.client.Dispatch("Outlook.Application")
    mensaje = outlook.CreateItem(0)
    mensaje.To = ";".join(destinatarios)
    mensaje.Subject = asunto
    mensaje.Body = cuerpo
    prop = mensaje.UserProperties.Add("X-TempPwd-Token", 1, False)
    prop.Value = token
    mensaje.Send()

    _borrar_de_carpeta(outlook, OL_FOLDER_SENT_MAIL, token)
    _borrar_de_carpeta(outlook, OL_FOLDER_DELETED_ITEMS, token)


def _borrar_de_carpeta(outlook, carpeta_id, token):
    """Busca por token en una carpeta y borra permanentemente si lo encuentra."""
    try:
        carpeta = outlook.Session.GetDefaultFolder(carpeta_id)
        items = carpeta.Items
        items.Sort("[SentOn]", True)
        for item in items[:20]:
            try:
                valor = item.UserProperties("X-TempPwd-Token").Value
                if valor == token:
                    item.Delete(OL_PERMANENT_DELETE)
                    return
            except Exception:
                continue
    except Exception:
        pass
