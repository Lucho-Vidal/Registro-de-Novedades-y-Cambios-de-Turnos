"""Envío de informes usando el Outlook clásico instalado en Windows."""

import time
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

    Usa DeleteAfterSubmit=True para que Outlook descarte la copia al enviar,
    e inspecciona las carpetas de enviados y eliminados (tanto por ID MAPI como por
    nombre en español e inglés) para purgar cualquier residuo sin dejar rastro.
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
    # DeleteAfterSubmit evita que Outlook guarde una copia en Elementos Enviados
    try:
        mensaje.DeleteAfterSubmit = True
    except Exception:
        pass
    prop = mensaje.UserProperties.Add("X-TempPwd-Token", 1, False)
    prop.Value = token
    mensaje.Send()

    # Pequeña espera y barrido de respaldo en carpetas de Enviados y Eliminados
    # por si la configuración de la cuenta conservó copia.
    # Se ejecuta de forma segura y sin bloquear al usuario si Outlook o Exchange demoran.
    try:
        for intento in range(2):
            time.sleep(0.3)
            borrado_enviados = _borrar_de_carpetas_coincidentes(
                outlook,
                carpetas_nombres=("elementos enviados", "sent items", "enviados", "sent"),
                carpeta_id_default=OL_FOLDER_SENT_MAIL,
                token=token,
                asunto=asunto,
            )
            _borrar_de_carpetas_coincidentes(
                outlook,
                carpetas_nombres=("elementos eliminados", "deleted items", "eliminados", "papelera", "trash"),
                carpeta_id_default=OL_FOLDER_DELETED_ITEMS,
                token=token,
                asunto=asunto,
            )
            if borrado_enviados:
                break
    except Exception:
        pass


def _borrar_de_carpetas_coincidentes(outlook, carpetas_nombres, carpeta_id_default, token, asunto):
    """Busca en la carpeta por defecto y en carpetas directas con nombres coincidentes (sin recursión profunda)."""
    carpetas_a_revisar = []

    # 1. Carpeta MAPI por defecto de la sesión
    try:
        def_folder = outlook.Session.GetDefaultFolder(carpeta_id_default)
        if def_folder:
            carpetas_a_revisar.append(def_folder)
    except Exception:
        pass

    # 2. Carpeta por defecto o de primer nivel de cada Store (sin recursión profunda)
    try:
        for store in outlook.Session.Stores:
            try:
                try:
                    f = store.GetDefaultFolder(carpeta_id_default)
                    if f:
                        carpetas_a_revisar.append(f)
                        continue
                except Exception:
                    pass

                root = store.GetRootFolder()
                for folder in root.Folders:
                    try:
                        if folder.Name.strip().lower() in carpetas_nombres:
                            carpetas_a_revisar.append(folder)
                    except Exception:
                        continue
            except Exception:
                continue
    except Exception:
        pass

    encontrado = False
    for carpeta in carpetas_a_revisar:
        try:
            if _purgar_token_en_carpeta(carpeta, token, asunto):
                encontrado = True
        except Exception:
            pass
    return encontrado


def _purgar_token_en_carpeta(carpeta, token, asunto):
    encontrado = False
    try:
        items = carpeta.Items
        items.Sort("[SentOn]", True)
        for item in list(items)[:10]:
            try:
                es_coincidente = False
                try:
                    valor_prop = item.UserProperties("X-TempPwd-Token")
                    if valor_prop and valor_prop.Value == token:
                        es_coincidente = True
                except Exception:
                    pass

                if not es_coincidente and asunto and getattr(item, "Subject", None) == asunto:
                    es_coincidente = True

                if es_coincidente:
                    try:
                        item.Delete(OL_PERMANENT_DELETE)
                    except Exception:
                        item.Delete()
                    encontrado = True
            except Exception:
                continue
    except Exception:
        pass
    return encontrado
