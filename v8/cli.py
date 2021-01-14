import logging
import subprocess


def updateConfig(self):
    logging.info("start update")
    cmdstr = f'{self.__1cpath} CONFIG /UpdateDBCfg /S"{self.Srvr}/{self.Ref}" /N"{self.Usr}" /P"{self.Pwd}" /WA- /UC{self.__permissionCode}'
    try:
        output = subprocess.run(cmdstr, capture_output=True, check=True)
        logging.debug(
            "Вывод комманды обноления конфиуграции {output.stdout}")
    except subprocess.CalledProcessError as err:
        logging.error(err.stderr)

    logging.info("update comlite")


__1cpath = '"C:\\Program Files (x86)\\1cv8\\8.3.10.2580\\bin\\1cv8.exe"'
