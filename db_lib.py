import configparser
import logging
import os

import mysql
from mysql.connector import Error
from mysql.connector import errorcode
from mysql.connector.locales.eng import client_error

import getAppPath
from loadconfig import getconfig


class MySQLDataBase:
    @staticmethod
    def connect(db_config):
        configFolder, breakDownConfig = getconfig()
        logging.getLogger("db_lib").disabled = True
        config = configparser.ConfigParser()
        app_path = getAppPath.getapppath()
        ini = os.path.join(configFolder, 'config\\database.ini')
        config.read(ini)
        db_user = config['breakdown_db']['USER']
        db_password = config['breakdown_db']['PASSWORD']
        db_host = config['breakdown_db']['HOST']
        db_name = config['breakdown_db']['DATABASE']
        db_auth = config['breakdown_db']['AUTH_PLUGIN']
        try:
            conn = mysql.connector.connect(user=db_user, password=db_password, host=db_host,
                                           database=db_name, auth_plugin=db_auth)
            return conn
        except Error as err:
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                print("Something is wrong with your user name or password!")
            elif err.errno == errorcode.ER_BAD_DB_ERROR:
                print("Database does not exist!")
            else:
                print(err)
        else:
            conn.close()
            print("Conncetion is closed.")