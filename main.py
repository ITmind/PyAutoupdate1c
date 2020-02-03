#!python3-32
import pyuac
from V8py.v8 import V83
import os
import sys
import json
import logging
import time
from io import StringIO
from smtplib import SMTP
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

_stringStream = StringIO()


def main():

    config = readConfig()
    if config is None:
        logging.warning('config.json created. restart app')
        return

    logging.info('start PyAutoupdate1c')

    _1c = V83(config['server'], config['1cpass'])
    for baseName in config['bases']:
        if _1c.checkModify(baseName, 'Администратор'):
            logging.info(f'{baseName} modify')
            if _1c.enableSessionDenied(baseName):
                _1c.restart1cService()
                _1c.updateConfig(baseName, 'Администратор')
                time.sleep(20)
                _1c.disableSessionDenied(baseName)
        else:
            logging.info(f'{baseName} not modify')

    logging.info('end')

    message = _stringStream.getvalue()
    msg = MIMEMultipart()
    msg['From'] = config['mail']['from']
    msg['To'] = config['mail']['to']
    msg['Subject'] = "PyAutoupdate1c"
    msg.attach(MIMEText(message, 'plain'))

    with SMTP(f"{config['mail']['smtpserver']}: {config['mail']['smtpport']}") as smtp:
        smtp.starttls()
        smtp.login(msg['From'], config['mail']['pass'])
        smtp.sendmail(msg['From'], msg['To'], msg.as_string())


def logSetup():
    logFile = os.path.abspath(os.path.join(
        os.path.dirname(__file__), 'log.log'))
    logging.basicConfig(format='%(asctime)s [%(levelname)s]: %(message)s',
                        handlers=[logging.FileHandler(
                            logFile, encoding='utf-8'),
                            logging.StreamHandler(),
                            logging.StreamHandler(_stringStream)],
                        level=logging.DEBUG)


def readConfig():
    configFile = os.path.abspath(os.path.join(
        os.path.dirname(__file__), 'config.json'))
    logging.debug(configFile)
    if not os.path.exists(configFile):
        createDefConfig()
        return None

    with open(configFile, encoding='utf-8') as file:
        config = json.load(file)

    return config


def createDefConfig():
    configFile = os.path.abspath(os.path.join(
        os.path.dirname(__file__), 'config.json'))
    logging.debug(configFile)

    config = {}
    config['server'] = '127.0.0.1'
    config['bases'] = ['ut11', 'bp']
    config['1cpass'] = {'user1': 'pass1', 'user2': 'pass2'}
    config['mail'] = {'smtpserver': 'smtp.gmail.com',
                      'smtpport': '587',
                      'pass': 'password'}

    with open(configFile, 'w', encoding='utf-8') as file:
        json.dump(config, file)


def runAsAdmin(func):
    if not pyuac.isUserAdmin():
        rc = pyuac.runAsAdmin()
        logging.info('start as admin')
    else:
        try:
            logging.info('admin privelegies')
            func()
            # input('ENTER')
        except Exception as err:
            logging.error(f'{err}')

    return rc


if __name__ == "__main__":
    logSetup()
    logging.debug('start')
    # main()
    res = runAsAdmin(main)
    sys.exit(res)
