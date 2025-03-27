import logging
import logging.config
import os

def setup_logging(log_level=logging.INFO, log_file='andreani_automation.log'):
    """
    Configura el sistema de logging.

    Args:
        log_level: Nivel de logging (DEBUG, INFO, WARNING, ERROR, CRITICAL).
        log_file: Nombre del archivo de log.
    """
    log_dir = 'logs'
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
    # log_filepath = os.path.join(log_dir, log_file)
    log_filepath = log_file

    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filepath, mode='w', encoding='utf-8'), # Modo 'w' para sobreescribir el log en cada ejecución
            logging.StreamHandler()  # Salida a consola
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info(f'Logging configurado al nivel: {logging.getLevelName(log_level)}')
    logger.info(f'Guardando logs en: {log_filepath}')
    return logger

if __name__ == '__main__':
    # Ejemplo de uso del logger
    logger = setup_logging(log_level=logging.DEBUG)
    logger.debug('Este es un mensaje de debug')
    logger.info('Este es un mensaje de info')
    logger.warning('Este es un mensaje de warning')
    logger.error('Este es un mensaje de error')
    logger.critical('Este es un mensaje crítico')
