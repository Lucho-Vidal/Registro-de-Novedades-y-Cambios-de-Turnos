"""Servicio de correo electrónico mediante Outlook clásico (MAPI COM).

Permite el envío ordinario de informes de novedades y el envío seguro
'sin rastro' para recuperación de contraseñas de acceso.
"""

import time
import uuid
from typing import List, Optional

from src.core.logger import get_logger

logger = get_logger("MailService")

OL_FOLDER_SENT_MAIL = 5
OL_FOLDER_DELETED_ITEMS = 3
OL_PERMANENT_DELETE = 2


class MailService:
    """Cliente MAPI COM para Microsoft Outlook."""

    @staticmethod
    def enviar_informe_outlook(
        destinatarios: List[str],
        asunto: str,
        cuerpo: str,
        adjuntos: Optional[List[str]] = None,
    ) -> None:
        """Envía un correo mediante la instancia local de Outlook."""
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

        if adjuntos:
            for adj in adjuntos:
                mensaje.Attachments.Add(adj)

        mensaje.Send()
        logger.info(f"Informe enviado exitosamente vía Outlook a {len(destinatarios)} destinatarios.")

    @staticmethod
    def enviar_sin_rastro(
        destinatarios: List[str],
        asunto: str,
        cuerpo: str,
    ) -> None:
        """Envía un correo y elimina permanentemente la copia de Enviados y Eliminados.

        Usa DeleteAfterSubmit=True para que Outlook descarte la copia al enviar,
        y barre las carpetas de enviados y eliminados con un identificador único
        para purgar cualquier residuo sin dejar rastro de contraseñas temporales.
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

        try:
            mensaje.DeleteAfterSubmit = True
        except Exception:
            pass

        try:
            prop = mensaje.UserProperties.Add("X-TempPwd-Token", 1, False)
            prop.Value = token
        except Exception:
            pass

        mensaje.Send()

        # Limpieza residual
        try:
            time.sleep(0.3)
            namespace = outlook.GetNamespace("MAPI")
            carpetas_a_revisar = []

            for folder_id in (OL_FOLDER_SENT_MAIL, OL_FOLDER_DELETED_ITEMS):
                try:
                    carpetas_a_revisar.append(namespace.GetDefaultFolder(folder_id))
                except Exception:
                    pass

            for folder in carpetas_a_revisar:
                try:
                    items = folder.Items
                    total = items.Count
                    for i in range(total, max(0, total - 15), -1):
                        try:
                            item = items.Item(i)
                            prop = item.UserProperties.Find("X-TempPwd-Token")
                            if prop and prop.Value == token:
                                item.Delete()
                        except Exception:
                            continue
                except Exception:
                    pass
        except Exception as e:
            logger.warning(f"Error menor en barrido sin rastro de Outlook: {e}")
