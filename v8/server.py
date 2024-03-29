import win32com.client
import logging
import time


# вызывает функцию нескольо раз при исключении
def attempts(numAttemps):
    def decorator(func):
        def wrapper(*args, **kwargs):
            res = False
            for i in range(0, numAttemps):
                try:
                    func(*args, **kwargs)
                    res = True
                    break
                except Exception as err:
                    logging.error(f'{err}')
                    time.sleep(10)
            return res
        return wrapper
    return decorator


class V83:
    __permissionCode = '123'
    
    def __init__(self, base):
        self.Srvr = base['Srvr']
        self.Ref = base['Ref']
        self.Usr = base['Usr']
        self.Pwd = base['Pwd']
        self.__com = win32com.client.Dispatch("v83.COMConnector")

    def DisconectAllSessions(self):
        logging.debug(f"Отключение всех сессий для {self.Ref}")
        self.connectWorkingProcess()
        time.sleep(1)
        for workproc in self.__workingProceses:
            InfoBases = workproc.GetInfoBases()
            for InfoBase in InfoBases:
                if InfoBase.Name.lower() != self.Ref.lower():
                    continue

                Connections = workproc.GetInfoBaseConnections(InfoBase)
                for connection in Connections:
                    if connection.AppID.lower() == "comconsole":
                        continue
                    workproc.Disconnect(connection)

                time.sleep(10)
                Connections = workproc.GetInfoBaseConnections(InfoBase)
                for connection in Connections:
                    if connection.AppID.lower() == "comconsole":
                        continue
                    logging.error(
                        f"В базе {self.Ref} не получилось отключить пользователя {connection.ConnID} {connection.AppID}")
                    return False
                continue
        return True

    @attempts(10)
    def connectWorkingProcess(self, port=1560):
        self.__agent = self.__com.ConnectAgent(self.Srvr)
        self.__workingProceses = []
        for claster in self.__agent.GetClusters():
            self.__agent.Authenticate(claster, "", "")
            WorkingProcesses = self.__agent.GetWorkingProcesses(claster)
            for WorkingProcess in WorkingProcesses:
                if WorkingProcess.Running != 1:
                    continue
                ConnectToWorkProcess = self.__com.ConnectWorkingProcess(
                    f"tcp://{WorkingProcess.HostName}:{WorkingProcess.MainPort}")
                ConnectToWorkProcess.AddAuthentication(self.Usr, self.Pwd)
                self.__workingProceses.append(ConnectToWorkProcess)

    
    def setSessionDenided(self, value):
        self.connectWorkingProcess()
        time.sleep(1)  # без задержки частые обрывы
        for workproc in self.__workingProceses:
            InfoBases = workproc.GetInfoBases()
            for base in InfoBases:
                if base.Name.lower() != self.Ref.lower():
                    continue

                logging.info(f'session for {self.Ref} is {value}')
                base.ScheduledJobsDenied = value
                base.PermissionCode = self.__permissionCode
                base.SessionsDenied = value
                workproc.UpdateInfoBase(base)
                return True
        return False