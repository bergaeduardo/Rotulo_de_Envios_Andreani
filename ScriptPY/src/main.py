# -*- coding: utf-8 -*-
import os
# # from config import ROOT_DIR
# CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# # ROOT_DIR = os.path.join(CORE_DIR, 'ScriptPY')
# print("CORE_DIR: ",CORE_DIR)
# # print("ROOT_DIR: ",ROOT_DIR)

# # print(ROOT_DIR)
# SRC_DIR = os.path.join(CORE_DIR, 'src')
# print("SRC_DIR: ",SRC_DIR)
import sys
# # print("RUTA: ",sys.path)
CORE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# ROOT_DIR = os.path.join(CORE_DIR, 'src')
# print("CORE_DIR: ",CORE_DIR)
# # print("ROOT_DIR: ",ROOT_DIR)
# sys.path.insert(0, CORE_DIR)
# print("RUTA: ",sys.path)
import time
import math
import logging
import configparser

# from src import excel_manager
# from src import andreani_automation
# from src import printer_manager
# from src.configuracion import Configuracion
# from src.logger_config import setup_logging

import excel_manager
import andreani_automation
import printer_manager
import configuracion
import logger_config

print()
def normalize_excel_data(config, logger):
    """
    Normaliza los datos en el archivo Excel de entrada (ImportacionMasivaEncomiendaDomicilio.xlsx).

    Args:
        config: Objeto de configuración.
        logger: Objeto logger.
    """
    input_file_path = config.get('Excel', 'input_file_path')
    input_file_path = CORE_DIR + input_file_path
    output_file_path_pruebas = config.get('Excel', 'output_file_path_pruebas')
    output_file_path_pruebas = CORE_DIR + output_file_path_pruebas

    try:
        workbook_origin = excel_manager.load_excel_workbook(input_file_path)
        sheet = workbook_origin.active

        i = 3
        while sheet.cell(row=i, column=1).value is not None:
            aux = sheet.cell(row=i, column=7).value
            if aux: # Verificar que aux no sea None antes de hacer split
                lista_nom = aux.split(" ")
                aux01 = len(lista_nom)
                apellido = lista_nom[aux01 - 1]
                nom_cli = ' '.join(lista_nom[:aux01 - 1]) #  Unir todos los nombres excepto el último (apellido)
                sheet.cell(row=i, column=7).value = nom_cli.strip() #  Eliminar espacios al inicio y final
                sheet.cell(row=i, column=8).value = apellido.strip() # Eliminar espacios al inicio y final

            aux_direc = sheet.cell(row=i, column=14).value
            cod_direc = 9999 # Valor por defecto si no se encuentra un código válido
            if aux_direc: # Verificar que aux_direc no sea None antes de procesar
                lista_dir = [int(s) for s in aux_direc.split() if s.isdigit()]
                for var01 in range(len(lista_dir) - 1, -1, -1): #  Iterar en reversa para encontrar el código correcto
                    if len(str(lista_dir[var01])) > 1: #  Considerar códigos numéricos de más de un dígito
                        cod_direc = lista_dir[var01]
                        break
                    elif var01 == 0 and lista_dir: # Si solo hay un número y es de un dígito, tomarlo
                        cod_direc = lista_dir[0]

            sheet.cell(row=i, column=15).value = str(cod_direc)
            i += 1

        workbook_origin.save(output_file_path_pruebas) # Guardar en el archivo de pruebas
        logger.info(f"Datos normalizados y guardados en: {output_file_path_pruebas}")

    except FileNotFoundError:
        logger.error(f"Error: Archivo de entrada no encontrado: {input_file_path}")
        raise
    except Exception as e:
        logger.error(f"Error durante la normalización de datos: {e}")
        raise


def main():
    # Cargar configuración e inicializar logger
    configuracion_manager = configuracion.Configuracion()
    config = configuracion_manager.get_config()

    log_level_str = config.get('Logging', 'log_level').upper() # Obtener nivel de log como string desde config
    log_level = getattr(logging, log_level_str, logging.INFO) # Convertir string a nivel de logging, INFO por defecto si no es válido
    logger = logger_config.setup_logging(log_level=log_level, log_file=config.get('Logging', 'log_file_path')) #  Usar ruta de log desde config

    logger.info("Inicio del script Imprimir_Etiquetas_Andreani.py")
    t_0 = time.time()  # Grabar el tiempo de inicio
    reg_maximos = 50 #  TODO: Mover a archivo de configuración si es necesario

    #  Inicializar PrinterManager y AndreaniAutomator
    printer_manager_instance = printer_manager
    andreani_automator = andreani_automation.AndreaniAutomator(config)

    operation_numbers_list = [] # Lista para almacenar todos los números de operación generados

    # 1. Cambiar a impresora de etiquetas
    logger.info("1. Estableciendo impresora de etiquetas...")
    label_printer_name = config.get('Printer', 'label_printer_name')
    print(f"Impresora de etiquetas establecida: {label_printer_name}")
    current_printer = printer_manager.get_default_printer() # Obtener impresora predeterminada actual
    printer_manager.set_default_printer(label_printer_name)
    print(f"Impresora de etiquetas establecida: {label_printer_name}")
    logger.info(f'Impresora de etiquetas establecida: {label_printer_name}')
    try:

        # 2. Normalizar datos en Excel
        logger.info("2. Normalizando datos en Excel...")
        normalize_excel_data(config, logger)
        output_file_path_pruebas = config.get('Excel', 'output_file_path_pruebas')
        output_file_path_pruebas = CORE_DIR + output_file_path_pruebas

        # 3. Calcular cantidad de archivos a procesar
        logger.info("3. Calculando cantidad de archivos a procesar...")
        workbook_pruebas = excel_manager.load_excel_workbook(output_file_path_pruebas)
        sheet_pruebas = workbook_pruebas.active
        num_rows = 0
        while sheet_pruebas.cell(row=num_rows + 3, column=1).value is not None: #  Contar filas con datos, asumiendo que la data empieza en la fila 3
            num_rows += 1
        cant_arch = math.ceil(num_rows / reg_maximos) if num_rows > 0 else 1 # Asegurar al menos 1 archivo si hay datos
        logger.info(f'Se procesarán {num_rows} registros en {cant_arch} archivos.')

        # 4. Automatización principal (bucle por cada archivo/lote)
        logger.info("4. Iniciando automática de impresión de etiquetas...")
        excel_template_path = config.get('Excel', 'template_file_path')
        excel_template_path = CORE_DIR + excel_template_path

        andreani_automator.start_browser() # Iniciar el navegador al principio del proceso
        logger.info("Iniciar el navegador...")
        andreani_automator.login_andreani() # Iniciar sesión en Andreani al principio del proceso
        logger.info("Iniciar sesión en Andreani...")

        for i in range(cant_arch):
            logger.info(f'Procesando archivo {i+1} de {cant_arch}...')
            # a. Limpiar plantilla Excel
            excel_manager.clear_cells(excel_template_path)

            # b. Copiar rango de registros a plantilla
            workbook_template = excel_manager.load_excel_workbook(excel_template_path)
            template_sheet = workbook_template.active
            workbook_origen_lotes = excel_manager.load_excel_workbook(output_file_path_pruebas) #  Cargar desde el archivo de pruebas normalizado
            sheet_origen_lotes = workbook_origen_lotes.active

            start_row_origin = 3 + (reg_maximos * i)
            end_row_origin = min(202 + (reg_maximos * i), sheet_origen_lotes.max_row) #  Ajustar fila final para no exceder los datos
            if start_row_origin <= end_row_origin: #  Verificar que haya datos para copiar en este lote
                selected_range = excel_manager.copy_range(1, start_row_origin, 29, end_row_origin, sheet_origen_lotes)
                excel_manager.paste_range(1, 3, 29, 2 + len(selected_range), template_sheet, selected_range) # Ajustar fila final pegado
                workbook_template.save(excel_template_path)
                excel_manager.save_excel_xlsx_format(excel_template_path) # Guardar en formato .xlsx

                # c. Automatizar carga en Andreani y obtener números de operación
                if i == 0: #  Navegar a carga masiva y seleccionar opciones solo la primera vez
                    andreani_automator.navigate_to_massive_upload()

                andreani_automator.upload_excel_file(excel_template_path)
                andreani_automator.confirm_massive_upload()
                operation_numbers_batch = andreani_automator.extract_operation_numbers()
                operation_numbers_list.extend(operation_numbers_batch) #  Añadir a la lista global

                logger.info(f'Números de operación del lote {i+1}: {operation_numbers_batch}')
            else:
                logger.warning(f'No hay datos para procesar en el lote {i+1}. Rango de filas vacío.')


        logger.info(f'Números de operación generados: {operation_numbers_list}')

        # 5. Imprimir etiquetas (bucle por cada número de operación)
        logger.info("5. Iniciando impresión de etiquetas...")
        if operation_numbers_list: #  Verificar si hay números de operación para imprimir
            logger.info('Iniciando impresión de etiquetas...')
            for op_num in operation_numbers_list:
                try:
                    andreani_automator.navigate_to_envio_management()
                    andreani_automator.search_operation_number(op_num)
                    andreani_automator.print_labels_for_operation(op_num)
                except Exception as print_err:
                    logger.error(f"Error al imprimir etiqueta para la operación '{op_num}': {print_err}")
        else:
            logger.warning('No se generaron números de operación, no se imprimirán etiquetas.')


    except Exception as main_exception:
        logger.critical(f"Error principal durante la ejecución del script: {main_exception}")
    finally:
        # 6. Finalización y limpieza
        logger.info("6. Finalización y limpieza...")
        if andreani_automator.browser:
            andreani_automator.close_browser() # Asegurar cerrar el navegador incluso si hay errores

        # 7. Reestablecer impresora predeterminada
        logger.info("7.Reestableciendo impresora predeterminada...")
        if current_printer: #  Verificar si se obtuvo la impresora predeterminada al inicio
            try:
                printer_manager.set_default_printer(current_printer)
                logger.info(f'(info)Impresora predeterminada reestablecida: {current_printer}')
            except Exception as e_printer_reset:
                logger.error(f"(error)Error al reestablecer la impresora predeterminada: {e_printer_reset}")
        else:
            logger.warning('No se pudo reestablecer la impresora predeterminada porque no se obtuvo al inicio.')


        t_1 = time.time()
        t_total = t_1 - t_0
        logger.info('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')
        logger.info('>>>>  El proceso finalizo  <<<<')
        logger.info('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')
        logger.info(f'Tiempo total de ejecución: {t_total:.2f} segundos') # Mostrar tiempo con 2 decimales
        print('Los Numero de Operacion generados son: \n') # print para mostrar en la consola de WebContainer
        print('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')
        for op_num in operation_numbers_list:
            print(op_num) # print para mostrar en la consola de WebContainer
            logger.info(f'Número de operación generado: {op_num}')
        print('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')


if __name__ == "__main__":
    main()
