import win32serviceutil
import logging


def restart1cService(self):
    logging.info('restart')
    try:
        win32serviceutil.RestartService(
            'Агент сервера 1С:Предприятия 8.3 (x86-64)',
            machine=self.Srvr)
    except Exception as err:
        logging.error(f'{err}')
