from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
import autoit  # Manteniendo la dependencia de autoit
import json
from bs4 import BeautifulSoup
import pyodbc
import os

CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
logger = logging.getLogger(__name__)

class AndreaniAutomator:
    def __init__(self, config):
        self.config = config
        self.browser = None # Inicializar browser como None
        self.core_dir = CORE_DIR

    def start_browser(self):
        """Inicia el navegador Chrome con las opciones configuradas."""
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--log-level=3')
            self.browser = webdriver.Chrome(options=options) # No se necesita executable_path en WebContainer
            self.browser.maximize_window()
            logger.info("Navegador Chrome iniciado exitosamente.")
        except Exception as e:
            logger.error(f"Error al iniciar el navegador Chrome: {e}")
            raise

    def close_browser(self):
        """Cierra el navegador."""
        if self.browser:
            try:
                self.browser.quit()
                logger.info("Navegador cerrado.")
            except Exception as e:
                logger.error(f"Error al cerrar el navegador: {e}")

    def login_andreani(self):
        """Inicia sesión en Andreani Online."""
        try:
            self.browser.get(self.config['Andreani']['base_url'] + '/login')
            WebDriverWait(self.browser, 200).until(
                EC.presence_of_element_located((By.ID, 'main'))
            )
            search_button = self.browser.find_element(By.ID, 'loginButton')
            search_button.click()

            WebDriverWait(self.browser, 200).until(
                EC.presence_of_element_located((By.ID, 'username'))
            )
            search_textinput = self.browser.find_element(By.ID, 'username')
            search_textinput.send_keys(self.config['Andreani']['username'])
            time.sleep(2)
            search_button = self.browser.find_element(By.ID, 'continueButton')
            search_button.click()
            time.sleep(5)
            element = WebDriverWait(self.browser, 200).until(EC.presence_of_element_located((By.ID,'password')))
            search_textinput = self.browser.find_element(By.ID, 'password')
            search_textinput.send_keys(self.config['Andreani']['password'])
            time.sleep(2)
            search_button = self.browser.find_element(By.ID, 'continueButton')
            search_button.click()
            logger.info("Inicio de sesión en Andreani exitoso.")

        except Exception as e:
            logger.error(f"Error al iniciar sesión en Andreani: {e}")
            raise

    def navigate_to_massive_upload(self):
        """Navega a la página de carga masiva de envíos."""
        try:
            self.browser.get(self.config['Andreani']['base_url'] + '/hacer-un-envio')
            WebDriverWait(self.browser, 200).until(
                EC.visibility_of_element_located((By.ID, 'tabs-modo-de-carga'))
            )
            time.sleep(3)
            search_button = self.browser.find_element(By.ID, '1').click()
            WebDriverWait(self.browser, 200).until(
                EC.visibility_of_element_located((By.ID, 'btn-continuar-masiva-1'))
            )
            time.sleep(3)
            # seleccionar lista desplegable de que enviar
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j_tipo_de_destino_y_envio_chosen"]').click()
            time.sleep(2)
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j_tipo_de_destino_y_envio_chosen"]/div/ul/li[2]').click()
            time.sleep(2)
            # seleccionar lista desplegable de tipo de envio
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j_modo_de_entrega_chosen"]').click()
            time.sleep(2)
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j_modo_de_entrega_chosen"]/div/ul/li[3]').click()
            time.sleep(5)
            # seleccionar siguiente
            search_button = self.browser.find_element(By.ID, 'btn-continuar-masiva-1')
            search_button.click()
            time.sleep(5)
            logger.info("Navegación a página de carga masiva completada.")

        except Exception as e:
            logger.error(f"Error al navegar a la página de carga masiva: {e}")
            raise

    def upload_excel_file(self, excel_file_path):
        """Carga el archivo Excel en la página de carga masiva."""
        try:
            WebDriverWait(self.browser, 200).until(
                EC.visibility_of_element_located((By.ID, 'form-masiva-paso-2'))
            )
            # Cargar archivo
            time.sleep(5)
            elem = self.browser.find_element(By.XPATH, "//input[@type='file']")
            elem.send_keys(excel_file_path)
            time.sleep(10)
            # seleccionar boton siguiente
            search_button = self.browser.find_element(By.ID, 'btn-continuar-masiva-2')
            search_button.click()
            logger.info(f"Archivo Excel '{excel_file_path}' cargado exitosamente.")

        except Exception as e:
            logger.error(f"Error al cargar el archivo Excel: {e}")
            raise

    def confirm_massive_upload(self):
        """Confirma la carga masiva y procesa los envíos."""
        try:
            time.sleep(5)
            # Selecciona sucursal de origen
            WebDriverWait(self.browser, 200).until(
                EC.visibility_of_element_located((By.ID,'btn-confirmar-masiva'))
            )
            search_button = self.browser.find_element(By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[1]/form/div[1]/input').click()
            search_button = self.browser.find_element(By.ID, 'j_direccion_origen_domicilio_chosen').click()
            time.sleep(5)
            search_button = self.browser.find_element(By.XPATH,'/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[1]/form/div[2]/div[2]/div[1]/div/ul/li[2]').click()
            time.sleep(2)
            # Presiona Confirmar
            search_button = self.browser.find_element(By.ID, 'btn-confirmar-masiva')
            search_button.click()
            time.sleep(10)
            WebDriverWait(self.browser, 500).until(
                EC.visibility_of_element_located((By.XPATH, "//*[@class='botones-resumen'][@style='display: block;']"))
            )
            # self.browser.execute_script("window.scrollBy(0, window.innerHeight)")  # Desplaza la pagina al final

            # selecciona el check de responsabilidad
            WebDriverWait(self.browser, 500).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="j-checkbox-confirmacion"]'))
            )
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j-checkbox-confirmacion"]').click()
            time.sleep(5)
            # selecciona el boton de confirmar
            search_button = self.browser.find_element(By.PARTIAL_LINK_TEXT, 'Confirmar')
            search_button.click()
            time.sleep(5)
            logger.info("Carga masiva confirmada y envíos procesados.")

        except Exception as e:
            logger.error(f"Error al confirmar la carga masiva: {e}")
            raise

    def extract_operation_numbers(self):
        """Extrae los números de operación de los mensajes de la página."""
        operation_numbers = []
        try:
            WebDriverWait(self.browser, 200).until(
                EC.visibility_of_element_located((By.CLASS_NAME, 'message'))
            )
            elements = self.browser.find_elements(By.XPATH, "//div[@class='message']/h4")

            for element in elements:
                message_text = element.text
                parts = message_text.split(" ")
                # La variable parts[] contiene el numero de operacion (posición 3)
                operation_numbers.append(parts[3])
            logger.info(f"Números de operación extraídos: {operation_numbers}")
            return operation_numbers

        except Exception as e:
            logger.error(f"Error al extraer los números de operación: {e}")
            return operation_numbers # Retornar la lista vacía o parcial en caso de error

    def navigate_to_envio_management(self):
        """Navega a la página de gestión de envíos."""
        try:
            self.browser.get(self.config['Andreani']['base_url'] + '/envio')
            time.sleep(5) # Considerar WebDriverWait si es necesario esperar a que cargue la página
            logger.info("Navegación a la página de gestión de envíos completada.")
        except Exception as e:
            logger.error(f"Error al navegar a la página de gestión de envíos: {e}")
            raise

    def search_operation_number(self, operation_number):
        """Busca un número de operación en la página de gestión de envíos."""
        try:
            WebDriverWait(self.browser, 200).until(
                EC.presence_of_element_located((By.ID, 'filtroGeneral'))
            )
            search_textinput = self.browser.find_element(By.ID, 'filtroGeneral')
            search_textinput.send_keys(operation_number)
            time.sleep(5)
            # Presiona el boton Buscar
            search_textinput = self.browser.find_element(By.XPATH,'//*[@id="historial"]/div/div/div/form/div[1]/div[1]/button[1]')
            search_textinput.click()
            time.sleep(10)
            # Mostrar 100 registros (opcional, si es necesario asegurar que se muestren todos los resultados)
            search_button = self.browser.find_element(By.ID, 'j-cantidad').click()
            search_button = self.browser.find_element(By.XPATH, '//*[@id="j-cantidad"]/option[4]').click()
            time.sleep(10)
            logger.info(f"Búsqueda del número de operación '{operation_number}' completada.")

        except Exception as e:
            logger.error(f"Error al buscar el número de operación '{operation_number}': {e}")
            raise

    def print_labels_for_operation(self, operation_number):
        path_docJson = self.core_dir + '/Documents/tabla_envios.json'
        try:
            # Web scraping para extraer datos de la tabla e imprimir en JSON
            html_content = self.browser.page_source
            soup = BeautifulSoup(html_content, 'html.parser')
            table = soup.find('table', class_='table table-striped table-hover grilla-envios grilla-generica')

            if table is None:
                logger.error("No se encontró la tabla en la página.")
                return

            # Agregar corchetes a los nombres de las columnas
            headers = [f"[{th.text.strip()}]" for th in table.find('thead').find_all('th')][1:]  # Ignora la primera columna de checkbox
            data = []
            for row in table.find('tbody').find_all('tr'):
                row_data = [td.text.strip() for td in row.find_all('td')][1:]  # Ignora la primera columna de checkbox
                if row_data:  # Asegurarse de que la fila no esté vacía
                    data.append(dict(zip(headers, row_data)))

            if not data:
                logger.error("No se encontraron datos en la tabla.")
                return

            try:
                # Guardar los datos con los nombres de columnas entre corchetes
                with open(path_docJson, 'w+', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                logger.info("Datos de la tabla guardados en 'tabla_envios.json'")
                print("Datos de la tabla guardados en 'tabla_envios.json'")  # Mensaje para la consola
            except Exception as json_error:
                logger.error(f"Error al escribir el archivo JSON: {json_error}")
                print(f"Error al escribir el archivo JSON: {json_error}")  # Mensaje para la consola

            # Presiona el check seleccionar todos
            self.browser.execute_script("window.scrollBy(0,-4000)")  # Desplazamiento de la pagina hacia arriba
            search_textinput = self.browser.find_element(By.XPATH,
                                                        '//*[@id="historial"]/div/div/div/form/table/thead/tr/th[1]/input')
            search_textinput.click()
            # Presiona el boton Imprimir
            time.sleep(3)
            search_textinput = self.browser.find_element(By.ID, 'imprimirConstanciasSeleccionadas')
            search_textinput.click()
            time.sleep(10)
            # Guardar el valor de la nueva ventana
            ventana_nueva = self.browser.window_handles[1]  # Asumiendo que la nueva ventana es la segunda
            ventana_actual = self.browser.window_handles[0]
            time.sleep(5)
            self.browser.switch_to.window(ventana_nueva)
            time.sleep(5)
            autoit.send("^p")  # Mantenemos autoit para simular la impresión
            time.sleep(15)
            # autoit.send("{ENTER}")
            autoit.send("{ESC}")
            print(f'Se imprimió la operacion {operation_number}')  # print para mostrar en la consola de WebContainer
            logger.info(f"Impresión de etiquetas para la operación '{operation_number}' simulada con AutoIt.")
            time.sleep(15)
            self.browser.switch_to.window(ventana_actual)

        except Exception as e:
            logger.error(f"Error al imprimir etiquetas para la operación '{operation_number}': {e}")
            raise

    def import_json_to_sql_server(self):
        """Importa datos desde tabla_envios.json a SQL Server."""
        # json_file_path = 'ScriptPY/src/tabla_envios.json' #  Ruta al archivo JSON
        json_file_path = self.core_dir + '/Documents/tabla_envios.json'
        config_sql = self.config['SQLServer'] #  Obtener la sección de configuración de SQL Server
        table_name = config_sql['table_name'] #  Nombre de la tabla desde la configuración

        try:
            # Leer datos desde JSON
            with open(json_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            if not data:
                logger.warning(f"No hay datos en el archivo JSON: '{json_file_path}'.")
                print(f"Advertencia: No hay datos en el archivo JSON: '{json_file_path}'.") #  Mensaje para la consola
                return

            conn_str = (
                f'DRIVER={{ODBC Driver 17 for SQL Server}};' #  Asegúrate de tener el driver correcto instalado
                f'SERVER={config_sql["server"]};'
                f'PORT={config_sql["port"]};'
                f'DATABASE={config_sql["database"]};'
                f'UID={config_sql["username"]};'
                f'PWD={config_sql["password"]}'
            )
            cnxn = pyodbc.connect(conn_str) # Comentado para WebContainer
            cursor = cnxn.cursor() # Comentado para WebContainer

            # Crear la tabla si no existe
            cursor.execute(f"""
            IF NOT EXISTS (
                SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '{table_name}'
            )
            CREATE TABLE {table_name} (
                [Acciones] NVARCHAR(MAX),
                [N° Seguimiento] NVARCHAR(50),
                [Canal] NVARCHAR(MAX),
                [N° Interno] NVARCHAR(50),
                [N° Operación] NVARCHAR(50),
                [Fecha alta] DATETIME,
                [Fecha admisión] DATETIME,
                [Estado envío] NVARCHAR(50),
                [Fecha estado envío] DATETIME,
                [Usuario] NVARCHAR(100),
                [Centro de costos] NVARCHAR(MAX),
                [Sector] NVARCHAR(MAX),
                [Estado pedido] NVARCHAR(50),
                [Servicio] NVARCHAR(MAX),
                [Insumo] NVARCHAR(MAX),
                [Remitente] NVARCHAR(MAX),
                [Email remitente] NVARCHAR(100),
                [Teléfono remitente] NVARCHAR(20),
                [Dirección remitente] NVARCHAR(MAX),
                [C.P. remitente] NVARCHAR(20),
                [Localidad remitente] NVARCHAR(MAX),
                [Provincia remitente] NVARCHAR(MAX),
                [Origen] NVARCHAR(MAX),
                [Destinatario] NVARCHAR(100),
                [Email destinatario] NVARCHAR(100),
                [Teléfono destinatario] NVARCHAR(20),
                [Dirección destino] NVARCHAR(MAX),
                [C.P. destino] NVARCHAR(20),
                [Localidad destino] NVARCHAR(MAX),
                [Provincia destino] NVARCHAR(MAX),
                [Nombre artículo] NVARCHAR(MAX),
                [Volumen] NVARCHAR(MAX),
                [Peso] NVARCHAR(MAX),
                [Valor declarado] NVARCHAR(MAX),
                [Fecha de retiro] DATETIME,
                [Fecha de entrega] DATETIME,
                [Sucursal de retiro] NVARCHAR(MAX),
                [Info. adicional] NVARCHAR(MAX),
                [Referencia] NVARCHAR(MAX),
                [Sistema origen] NVARCHAR(MAX),
                [Bultos] NVARCHAR(MAX),
                [Número de bulto] NVARCHAR(MAX),
                [Cupón] NVARCHAR(MAX),
                [Tarifa S/IVA] NVARCHAR(MAX),
                [Descuento S/IVA] NVARCHAR(MAX),
                [Seguro S/IVA] NVARCHAR(MAX),
                [Valor de envío S/IVA] NVARCHAR(MAX),
                [Total S/IVA] NVARCHAR(MAX),
                [Detalles] NVARCHAR(MAX)
            );
            """)

            # Insertar datos validando duplicados
            for record in data:
                unique_field = record.get("N° Seguimiento")
                cursor.execute(f"SELECT COUNT(*) FROM {table_name} WHERE [N° Seguimiento] = ?", (unique_field,))
                count = cursor.fetchone()[0]

                if count == 0:
                    columns = ', '.join(record.keys())
                    placeholders = ', '.join(['?'] * len(record))
                    insert_query = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
                    cursor.execute(insert_query, tuple(record.values()))
                else:
                    logging.info(f"Registro duplicado detectado: {unique_field}")

            cnxn.commit()
            cnxn.close()

            logging.info(f"Datos importados a la tabla '{table_name}' correctamente.")


        except FileNotFoundError:
            logger.error(f"Archivo JSON no encontrado: {json_file_path}")
            print(f"Error: Archivo JSON no encontrado: {json_file_path}") #  Mensaje para la consola
        except json.JSONDecodeError:
            logger.error(f"Error al decodificar JSON desde el archivo: {json_file_path}")
            print(f"Error al decodificar JSON desde el archivo: {json_file_path}") #  Mensaje para la consola
        except Exception as e: #  Capturar otras excepciones que puedan ocurrir
            logger.error(f"Error al importar datos a SQL Server: {e}")
            print(f"Error al importar datos a SQL Server: {e}") #  Mensaje para la consola


if __name__ == '__main__':
    # Ejemplo de uso (necesita un archivo config.ini de prueba)
    import configparser
    config = configparser.ConfigParser()
    config.read('config/config.ini') #  Ajustar la ruta si es necesario

    logging.basicConfig(level=logging.INFO)
    automator = AndreaniAutomator(config)
    # automator.import_json_to_sql_server()
    try:
        automator.start_browser()
        automator.login_andreani()
        # automator.navigate_to_massive_upload() #  Comentar para probar solo la parte de carga masiva
        # excel_template_path = config['Excel']['template_file_path'] # Asegúrate de que esta ruta sea correcta para pruebas
        # automator.upload_excel_file(excel_template_path) # Comentar para probar solo la parte de carga masiva
        # automator.confirm_massive_upload() #  Comentar para probar solo la parte de carga masiva
        # operation_numbers = automator.extract_operation_numbers() # Comentar para probar solo la parte de carga masiva
        operation_numbers = ['19069936'] #  Para pruebas de impresión, usar un número de operación conocido
        if operation_numbers:
            for op_num in operation_numbers:
                automator.navigate_to_envio_management()
                time.sleep(5)
                automator.search_operation_number(op_num)
                automator.print_labels_for_operation(op_num)
                

    except Exception as e:
        logger.error(f"Error durante la automatización: {e}")
    finally:
        automator.close_browser()
    print("Proceso de automatización finalizado.") # print para mostrar en la consola de WebContainer

    try:
        automator.import_json_to_sql_server()
    except Exception as e:
        logger.error(f"Error al importar datos a SQL Server: {e}")
    print("Proceso de importación a SQL Server finalizado.")
