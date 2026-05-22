import os
import sys

import yaml


def getconfig():
    if getattr(sys, 'frozen', False):
        app_path = os.path.dirname(sys.executable)
    elif __file__:
        app_path = os.path.dirname(__file__)
    # configPath = os.path.join(app_path, "config\\breakDown.yaml")
    startupConfigPath = os.path.join(app_path, "startupConfig.yaml")
    try:
        with open(startupConfigPath, 'r') as file:
            basicConfig = yaml.safe_load(file)
            configFolder = basicConfig['CONFIG']['BreakDown']
    except FileNotFoundError:
        print("ERROR: [001]: Startup config Yaml not found in tool path")
        exit()
    configPath = os.path.join(configFolder, "config\\breakDown.yaml")
    try:
        with open(configPath, 'r') as file:
            configDetails = yaml.safe_load(file)
            return configFolder, configDetails
    except FileNotFoundError:
        print("ERROR: [001]: Breakdown config Yaml not found in configPath")
        exit()

getconfig()