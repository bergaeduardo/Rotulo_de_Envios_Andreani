import win32print
import logging

logger = logging.getLogger(__name__)

def set_default_printer(printer_name):
    """
    Establece la impresora predeterminada del sistema.

    Args:
        printer_name: Nombre de la impresora a establecer como predeterminada.
    """
    try:
        win32print.SetDefaultPrinter(printer_name)
        logger.info(f"Impresora predeterminada establecida a: '{printer_name}'")
    except Exception as e:
        logger.error(f"Error al establecer la impresora predeterminada '{printer_name}': {e}")
        raise

def get_default_printer():
    """
    Obtiene el nombre de la impresora predeterminada actual.

    Returns:
        El nombre de la impresora predeterminada.
    """
    try:
        default_printer = win32print.GetDefaultPrinter()
        logger.info(f"Impresora predeterminada actual obtenida: '{default_printer}'")
        return default_printer
    except Exception as e:
        logger.error(f"Error al obtener la impresora predeterminada: {e}")
        return None

if __name__ == '__main__':
    # Ejemplo de uso
    logging.basicConfig(level=logging.INFO)

    current_default_printer = get_default_printer()
    if current_default_printer:
        print(f"Impresora predeterminada actual: {current_default_printer}")

    label_printer_name = "\\\\192.168.0.97\\ZDesigner GC420t (EPL)" #  Reemplazar con el nombre de tu impresora de etiquetas para pruebas
    try:
        set_default_printer(label_printer_name)
        print(f"Impresora predeterminada cambiada a: {label_printer_name}")
    except Exception as e:
        print(f"Error al cambiar la impresora: {e}")

    #  Reestablecer a la impresora predeterminada anterior (opcional para pruebas)
    if current_default_printer:
        try:
            set_default_printer(current_default_printer)
            print(f"Impresora predeterminada reestablecida a: {current_default_printer}")
        except Exception as e:
            print(f"Error al reestablecer la impresora: {e}")
