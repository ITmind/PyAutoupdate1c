import win32com.client
import logging


def checkModify(self):
    conStr = ';'.join(
        [f'Srvr="{self.Srvr}"', f'Ref="{self.Ref}"',
         f'Usr="{self.Usr}"', f'Pwd="{self.Pwd}"'])+';'
    logging.debug(conStr)
    result = False
    try:
        _v8 = self.__com.Connect(conStr)
        result = _v8.ConfigurationChanged()
    except Exception as err:
        logging.error(f'{self.Ref}: {err.excepinfo[2]}')

    return result
