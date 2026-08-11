"""Envío de informes usando el Outlook clásico instalado en Windows."""


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
