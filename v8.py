import win32com.client
import win32serviceutil
import subprocess
import logging


class V83:
    __permissionCode = '123'
    __1cpath = '"C:\\Program Files (x86)\\1cv8\\8.3.10.2580\\bin\\1cv8.exe"'

    def __init__(self, serverName):
        self.__serverName = serverName
        self.__com = win32com.client.Dispatch("v83.COMConnector")

    def connectWorkingProcess(self, port=1560):
        serverUrl = "tcp://" + self.__serverName + ":" + str(port)
        self.__workingProcces = self.__com.ConnectWorkingProcess(serverUrl)

    def checkModify(self, baseName, user, password):
        conStr = ';'.join(
            [f'Srvr="{self.__serverName}"', f'Ref="{baseName}"',
             f'Usr="{user}"', f'Pwd="{password}"'])+';'
        logging.debug(conStr)
        result = False
        try:
            _v8 = self.__com.Connect(conStr)
            result = _v8.ConfigurationChanged()
        except Exception as err:
            logging.error(f'{err}')

        return result

    def authentication(self, authDict):
        if not isinstance(authDict, dict):
            logging.warning(
                'authentication(authDict). authDict not dictionary')
            return

        for key, val in authDict.items():
            self.__workingProcces.AddAuthentication(key, val)

    def __setSessionDenided(self, baseNames, value):

        for base in self.__workingProcces.GetInfoBases():
            if isinstance(baseNames, list):
                if base.Name not in baseNames:
                    continue
            elif isinstance(baseNames, str):
                if base.Name != baseNames:
                    continue
            else:
                logging.warning('bad baseName type')
                return

            base.ScheduledJobsDenied = value
            base.PermissionCode = self.__permissionCode
            base.SessionsDenied = value
            self.__workingProcces.UpdateInfoBase(base)

    def enableSessionDenied(self, baseNames):
        self.__setSessionDenided(baseNames, True)

    def disableSessionDenied(self, baseNames):
        self.__setSessionDenided(baseNames, False)

    def restart1cService():
        logging.info('restart')
        try:
            win32serviceutil.RestartService(
                'Агент сервера 1С:Предприятия 8.3 (x86-64)', machine=self.__serverName)
        except Exception as err:
            logging.error(f'{err}')

    def updateConfig(self, baseName, user, password):
        logging.info("start update")
        cmdstr = f'{self.__1cpath} CONFIG /UpdateDBCfg /S"{self.__serverName}/{baseName}" /N"{user}" /P"{password}" /WA- /UC123'
        subprocess.run(cmdstr, check=True)
        logging.info("update comlite")
