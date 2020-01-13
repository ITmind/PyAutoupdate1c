#!python3-32
import pyuac
from v8 import V83
import os
import sys
import logging


def main():
    logSetup()
    logging.info('start PyAutoupdate1c')
    auth = {}
    passFile = os.path.abspath(os.path.join(
        os.path.dirname(__file__), 'pass.txt'))
    logging.debug(passFile)
    with open(passFile, encoding='utf-8') as file:
        for line in file:
            line = line.rstrip('\n')
            key, val = line.split(':')
            auth[key] = val

    baseNames = ['buhcopy', 'buh']

    _1c = V83('c2-it-s-1c')
    _1c.connectWorkingProcess()
    _1c.authentication(auth)
    for baseName in baseNames:

        if _1c.checkModify(baseName, 'Кулик И.А.', auth['Кулик И.А.']):
            logging.info(f'{baseName} modify')
            _1c.enableSessionDenied(baseName)
            _1c.restart1cService()
            _1c.updateConfig(baseName, 'Кулик И.А.', auth['Кулик И.А.'])            
            _1c.disableSessionDenied(baseName)
        else:
            logging.info(f'{baseName} not modify')

    logging.info('end')


def logSetup():
    logFile = os.path.abspath(os.path.join(
        os.path.dirname(__file__), 'log.log'))
    logging.basicConfig(format='%(asctime)s [%(levelname)s]: %(message)s',
                        handlers=[logging.FileHandler(
                            logFile, encoding='utf-8'), logging.StreamHandler()],
                        level=logging.DEBUG)


def runAsAdmin(func):
    if not pyuac.isUserAdmin():
        rc = pyuac.runAsAdmin()
    else:
        try:
            func()
            input('ENTER')
        except Exception as err:
            logging.error(f'{err}')

    return rc


if __name__ == "__main__":
    # main()
    res = runAsAdmin(main)
    sys.exit(res)
