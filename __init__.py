################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################

import argparse
import sys
import os

library = os.getcwd()+'/Library/'


sys.path.insert(0, library)
from get_files import *

sys.path.insert(0, library)
from xls_reader import *

sys.path.insert(0, library)
from xlsx_reader import *

sys.path.insert(0, library)
from excel_writer import *

sys.path.insert(0, library)
from xls_writter import *

db_sap = os.getcwd()+'/resources/db_sap.xlsx'

# Get the param folder to proccess
parser = argparse.ArgumentParser(description='Syntax Example: python ofertas_producto.py  /Users/user/BOMS_Listos')
parser.add_argument('folder', metavar='[folder]', help='Example: /Users/user/BOMS_Listos')
args = parser.parse_args()


# Get the files within folder
folder = args.folder


# Print message
logo = """
################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################
"""
print logo

files = getFiles(folder)


#Iterate files
new_boms = []
for f in files:
    # Concatenate the folder with file to create full path filename
    if folder[len(folder)-1] == "/":
        filename = folder+f
    else:
        filename = folder+"/"+f
    # Read the file
    if '~' in filename:
        continue

    if '.xls' in filename:
        bom = readxls(filename)
    else:
        continue




    for row in bom:
        row.append("")
        row.append("")

    # Get the list of SAP codes
    sap = readxlsx(db_sap)

    # Iterate the bom
    # print bom

    flag = False
    for row in bom:
        for col in range(len(row)):
            try:
                if 'Line Number' in row[col]:
                    flag = True
                    # Add the headers for the two extra columns
                    row[10] = 'Codigo SAP'
                    row[11] = 'Descrip Corta'
                if 'Valid through:' in row[col]:
                    flag = False
            except:
                pass
        if flag:
            # Iterate the sap db
            for line in sap:
                # Check if the sku exist in sap db
                if row[1] in line or row[1]+'=' in line or row[1].replace('=','') in line or row[1].replace('+','') in line or row[1]+'+' in line:
                    # Add the values of sap db to the bom
                    row[10] = line[4]
                    row[11] = line[3]


    # print bom
    new_boms.append(xlswritter(bom, filename))
print "The following BOMS where created"
for elem in new_boms:
    print elem
    # Create the new BOM with SAP codes
