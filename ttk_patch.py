"""Parche de compatibilidad y robustez para ttkbootstrap.

Corrige el error 'bad window path name' que ocurre cuando ttkbootstrap intenta
actualizar los estilos de menús desplegables (popdowns) en Comboboxes que ya han
sido destruidos (por ejemplo al cerrar diálogos o al cambiar de usuario).
"""

from ttkbootstrap.publisher import Publisher, Channel
from ttkbootstrap.style import StyleBuilderTTK

_PATCH_APLICADO = False


def aplicar_patch_ttkbootstrap():
    """Aplica parches sobre Publisher y StyleBuilderTTK de ttkbootstrap."""
    global _PATCH_APLICADO
    if _PATCH_APLICADO:
        return

    # 1. Parche sobre Publisher.publish_message para desuscribir widgets muertos
    original_publish = Publisher.publish_message

    def safe_publish_message(channel, *args):
        subs = Publisher.get_subscribers(channel)
        muertos = []
        for sub in subs:
            try:
                sub.func(*args)
            except Exception:
                muertos.append(sub.name)
        for nombre in muertos:
            Publisher.unsubscribe(nombre)

    Publisher.publish_message = safe_publish_message

    # 2. Parche sobre StyleBuilderTTK.update_combobox_popdown_style para verificar existencia
    original_update_popdown = StyleBuilderTTK.update_combobox_popdown_style

    def safe_update_combobox_popdown_style(self, widget):
        try:
            if hasattr(widget, "winfo_exists") and not widget.winfo_exists():
                return
            original_update_popdown(self, widget)
        except Exception:
            pass

    StyleBuilderTTK.update_combobox_popdown_style = safe_update_combobox_popdown_style

    _PATCH_APLICADO = True
