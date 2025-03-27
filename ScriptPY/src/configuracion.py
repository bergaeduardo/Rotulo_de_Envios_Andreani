import configparser
import os
import logging

logger = logging.getLogger(__name__)

class Configuracion:
    def __init__(self, config_file_path='config/config.ini'):
        self.config = configparser.ConfigParser()
        self.config_file_path = config_file_path
        self.load_config()

    def load_config(self):
        """Carga la configuración desde el archivo INI."""
        try:
            config_dir = os.path.dirname(self.config_file_path)
            # print ('config_dir: ',self.config_file_path)
            if not os.path.exists(config_dir):
                os.makedirs(config_dir)

            if not os.path.exists(self.config_file_path):
                # Si el archivo no existe, crea uno con secciones vacías
                self._create_default_config()
                logger.warning(f"Archivo de configuración no encontrado en '{self.config_file_path}'. Se ha creado un archivo de configuración predeterminado. Por favor, revíselo y configúrelo.")


            self.config.read(self.config_file_path)
            logger.info(f"Configuración cargada desde '{self.config_file_path}'")

        except Exception as e:
            logger.error(f"Error al cargar la configuración desde '{self.config_file_path}': {e}")
            raise

    def _create_default_config(self):
        """Crea un archivo de configuración predeterminado con secciones vacías."""
        self.config['Andreani'] = {
            'username': 'your_username',
            'password': 'your_password',
            'base_url': 'https://andreanionline.com'
        }
        self.config['Excel'] = {
            'input_file_path': 'X:\\\\Logistica\\\\1. Logistica Interno\\\\Historiales_de_Logistica\\\\ImportacionMasivaEncomiendaDomicilio.xlsx',
            'template_file_path': 'X:\\\\Logistica\\\\1. Logistica Interno\\\\Historiales_de_Logistica\\\\ImportacionAndreani.xlsx',
            'output_file_path_pruebas': 'X:\\\\Logistica\\\\1. Logistica Interno\\\\Historiales_de_Logistica\\\\ImportacionMasiva_Pruebas.xlsx'
        }
        self.config['Printer'] = {
            'label_printer_name': '\\\\\\\\192.168.0.97\\\\\\\\ZDesigner GC420t (EPL)'
        }
        self.config['Logging'] = {
            'log_file_path': 'logs/andreani_automation.log',
            'log_level': 'INFO'
        }
        with open(self.config_file_path, 'w') as configfile:
            self.config.write(configfile)

    def get_config(self):
        """Retorna el objeto de configuración."""
        return self.config

    def get_setting(self, section, key):
        """
        Obtiene un valor de configuración específico.

        Args:
            section: Sección del archivo de configuración.
            key: Clave dentro de la sección.

        Returns:
            El valor de configuración o None si no se encuentra.
        """
        try:
            return self.config.get(section, key)
        except (configparser.NoSectionError, configparser.NoOptionError):
            logger.warning(f"Configuración no encontrada: Sección='{section}', Clave='{key}'")
            return None

if __name__ == '__main__':
    # Ejemplo de uso
    logging.basicConfig(level=logging.INFO)

    configuracion_manager = Configuracion()
    config = configuracion_manager.get_config()

    # Imprimir la configuración para verificar
    print("Configuración cargada:")
    for section in config.sections():
        print(f"[{section}]")
        for key, value in config.items(section):
            print(f"  {key} = {value}")
        print()

    # Ejemplo de obtener un valor específico
    username = configuracion_manager.get_setting('Andreani', 'username')
    print(f"Username de Andreani: {username}")

    log_level_str = configuracion_manager.get_setting('Logging', 'log_level')
    print(f"Nivel de log (string): {log_level_str}")
