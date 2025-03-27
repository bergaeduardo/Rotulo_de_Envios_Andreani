from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
import autoit  # Manteniendo la dependencia de autoit

logger = logging.getLogger(__name__)

class AndreaniAutomator:
    def __init__(self, config):
        self.config = config
        self.browser = None # Inicializar browser como None

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
            time.sleep(2)
            element = WebDriverWait(self.browser, 200).until(EC.presence_of_element_located((By.ID,'password')))
            search_textinput = self.browser.find_element(By.ID, 'password')
            search_textinput.send_keys('Elarge,123')
            time.sleep(5)
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
        """Imprime las etiquetas para un número de operación."""
        try:
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
            ventana_nueva = self.browser.window_handles[1] # Asumiendo que la nueva ventana es la segunda
            ventana_actual = self.browser.window_handles[0]
            time.sleep(5)
            self.browser.switch_to.window(ventana_nueva)
            time.sleep(5)
            autoit.send("^p") #  Mantenemos autoit para simular la impresión
            time.sleep(15)
            autoit.send("{ENTER}")
            print(f'Se imprimió la operacion {operation_number}') #  print para mostrar en la consola de WebContainer
            logger.info(f"Impresión de etiquetas para la operación '{operation_number}' simulada con AutoIt.")
            time.sleep(15)
            self.browser.switch_to.window(ventana_actual) # Volver a la ventana principal

        except Exception as e:
            logger.error(f"Error al imprimir etiquetas para la operación '{operation_number}': {e}")
            raise


if __name__ == '__main__':
    # Ejemplo de uso (necesita un archivo config.ini de prueba)
    import configparser
    config = configparser.ConfigParser()
    config.read('../config/config.ini') #  Ajustar la ruta si es necesario

    logging.basicConfig(level=logging.INFO)
    automator = AndreaniAutomator(config)

    try:
        automator.start_browser()
        automator.login_andreani()
        # automator.navigate_to_massive_upload() #  Comentar para probar solo la parte de impresión
        # excel_template_path = config['Excel']['template_file_path'] # Asegúrate de que esta ruta sea correcta para pruebas
        # automator.upload_excel_file(excel_template_path) # Comentar para probar solo la parte de impresión
        # automator.confirm_massive_upload() #  Comentar para probar solo la parte de impresión
        # operation_numbers = automator.extract_operation_numbers() # Comentar para probar solo la parte de impresión
        operation_numbers = ['1234567'] #  Para pruebas de impresión, usar un número de operación conocido
        if operation_numbers:
            for op_num in operation_numbers:
                automator.navigate_to_envio_management()
                automator.search_operation_number(op_num)
                automator.print_labels_for_operation(op_num)

    except Exception as e:
        logger.error(f"Error durante la automatización: {e}")
    finally:
        automator.close_browser()
    print("Proceso de automatización finalizado.") # print para mostrar en la consola de WebContainer
