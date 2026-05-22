import os
from configparser import ConfigParser

import getAppPath
from loadconfig import getconfig


def read_db_config():
    app_path = getAppPath.getapppath()
    configFolder, breakDownConfig = getconfig()
    filename = os.path.join(configFolder, 'config\\database.ini')
    section = 'breakdown_db'
    # create parser and read ini configuration file
    parser = ConfigParser()
    parser.read(filename)
    db = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            db[item[0]] = item[1]
    else:
        raise Exception(
            '{0} not found in the {1} file'.format(
                section, filename))

    return db


def read_table_config(process_name):
    app_path = getAppPath.getapppath()
    configFolder, breakDownConfig = getconfig()
    filename = os.path.join(configFolder, 'config\\database.ini')
    section = 'table_names'
    parser = ConfigParser()
    parser.optionxform = str
    parser.read(filename)
    table_names = {}
    if parser.has_section(section):
        items = parser.items(section)
        for item in items:
            table_names[item[0]] = item[1]
    else:
        raise Exception(
            '{0} not found in the {1} file'.format(
                section, filename))
    table_name = table_names[process_name]
    return table_name
