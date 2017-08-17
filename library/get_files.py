################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################

# Read the list of files from a folder and return a the list with the full path of each file

import os

def getFiles(folder):
    # print os.listdir(folder)

    files = os.listdir(folder)
    return files
