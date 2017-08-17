################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################

# Write a xlsx file with the values passed in a dictionary

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import Color, Fill
from openpyxl.cell import Cell

def writexlsx(bom, filename):
    newfilename = filename.split('.')[0]+'_codigos_sap.xlsx'
    wb = Workbook()
    ws = wb.active
    for i in range(len(bom)):
        # if i == 25:
        #     c = Cell(ws,column = "A",row=i)
        #     c.font = Font(bold=True)
        #     print bom[i]
        ws.append(bom[i])
    wb.save(newfilename)
    return newfilename
