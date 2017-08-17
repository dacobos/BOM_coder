################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################

# Write a xls file with the values passed in a dictionary

import xlrd
import xlwt
import xlutils.copy
from xlutils.styles import Styles
from xls_reader import getIds

def xlswritter(bom, filename):

    Ids = getIds(filename)
    start_of_bom_id = Ids[0]
    end_of_bom_id = Ids[1]
    total_price_id  = Ids[2]


    # Copy the workbook to create a new one passing the formatting
    newfilename = filename.split('.')[0]+'_codigos_sap.xls'
    originwb = xlrd.open_workbook(filename, formatting_info=True)
    styles = Styles(originwb)
    rs = originwb.sheet_by_index(0)
    destinationwb = xlutils.copy.copy(originwb)
    xl_sheet = destinationwb.get_sheet(0)


    # Write the

    # for i,cell in enumerate(rs.col(8)):
    #     if not i:
    #         continue
    #     print i
    #     # xl_sheet.write(row,column,value)
    #     xl_sheet.write(i,7,22)

    # header_style = styles[rs.cell(start_of_bom_id,1)]
    # content_style = styles[rs.cell(start_of_bom_id+1,1)]

    # style = xlwt.XFStyle()
    # # bold
    # font = xlwt.Font()
    # font.bold = True
    # style.font = font
    #
    # # background color
    # pattern = xlwt.Pattern()
    # pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    # pattern.pattern_fore_colour = xlwt.Style.colour_map['pale_blue']
    # style.pattern = pattern

    header_style = xlwt.easyxf('pattern: pattern solid, fore_colour gray40;'
                              'font: colour black, bold True;')

    # content_style = xlwt.easyxf('pattern: pattern solid, fore_colour light_blue;'
    #                           'font: colour black, bold ff;')

    content_style = xlwt.easyxf('font: bold off, color black;\
                     borders: top_color gray25, bottom_color gray25, right_color gray25, left_color gray25,\
                              left thin, right thin, top thin, bottom thin;\
                     pattern: pattern solid, fore_color white;')

    for i in range(start_of_bom_id, end_of_bom_id-2):
        xl_sheet.write(i, 7, bom[i][7],content_style)

    for i in range(start_of_bom_id, end_of_bom_id-2):
        xl_sheet.write(i, 9, bom[i][9],content_style)

    xl_sheet.write(end_of_bom_id, 9, bom[end_of_bom_id][9])

    xl_sheet.write(start_of_bom_id-1, 10, 'Codigo SAP',header_style)
    xl_sheet.write(start_of_bom_id-1, 11, 'Descrip Corta',header_style)

    for i in range(start_of_bom_id, end_of_bom_id-2):
        xl_sheet.write(i, 10, bom[i][10],content_style)

    for i in range(start_of_bom_id, end_of_bom_id-2):
        xl_sheet.write(i, 11, bom[i][11],content_style)

    destinationwb.save(newfilename)
    return newfilename
