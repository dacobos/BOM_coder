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
    newfilename = filename.split('.')[0]+'_oferta_cisco.xls'
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

def createSkeleton(ps,filename):
    newfilename = filename.split('.')[0]+'_oferta_claro.xls'
    originwb = xlrd.open_workbook(ps, formatting_info=True)
    styles = Styles(originwb)
    rs = originwb.sheet_by_index(0)
    destinationwb = xlutils.copy.copy(originwb)
    xl_sheet = destinationwb.get_sheet(0)
    destinationwb.save(newfilename)
    return newfilename

def writePriceSheet(bom, siteList, ps, filename):

    content_normal = xlwt.easyxf('font: bold off, color black;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;\
                     pattern: pattern solid, fore_color white;\
                     align: horiz center;')

    header_normal = xlwt.easyxf('font: bold True, color black;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thin, right thin, top thin, bottom thin;\
                     pattern: pattern solid, fore_color white;\
                     align: horiz center;')

    content_strong = xlwt.easyxf('font: bold True, color black;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thick, right thick, top thick, bottom thick;\
                     pattern: pattern solid, fore_color white;\
                     align: horiz center;')

    subTotal_stong = xlwt.easyxf('font: bold True, color black;\
                     borders: top_color black, top thin;\
                     pattern: pattern solid, fore_color white;')

    granTotal_strong = xlwt.easyxf('font: bold True, color black;\
                     pattern: pattern solid, fore_color yellow;')

    sites_red = xlwt.easyxf('font: bold off, color white, height 280;\
                     pattern: pattern solid, fore_color red;\
                     align: horiz center, vert centre')

    gray_background = xlwt.easyxf('font: bold off, color white;\
                     pattern: pattern solid, fore_color gray40;\
                     align: horiz center')

    title_strong = xlwt.easyxf('font: bold on, color black, height 280;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thick, right thick, top thick, bottom thick;\
                     pattern: pattern solid, fore_color ice_blue;')

    fix_border =   xlwt.easyxf('font: bold True, color black;\
                     pattern: pattern solid, fore_color white;\
                     borders: top_color black, bottom_color black, right_color black, left_color black,\
                              left thick, right thick, top thick, bottom thick;\
                     align: horiz center, vert centre;')


    # Intial values
    site_id = []
    granTotal = 0.0
    subTotals = []
    linesPerSite = []


    # Create skeleton
    newfilename = filename.split('.')[0]+'_oferta_claro.xls'
    originwb = xlrd.open_workbook(ps, formatting_info=True)
    styles = Styles(originwb)
    rs = originwb.sheet_by_index(0)
    destinationwb = xlutils.copy.copy(originwb)
    xl_sheet = destinationwb.get_sheet(0)

    # Get original BOM start and end
    Ids = getIds(filename)
    start_of_bom_id = Ids[0]
    end_of_bom_id = Ids[1]
    total_price_id  = Ids[2]


    xl_sheet.write_merge(1,3, 4, 9, "Country (Project)", fix_border)
    xl_sheet.write_merge(10,10, 2, 4, "Cisco Systems", fix_border)
    xl_sheet.row(1).height_mismatch = True
    xl_sheet.row(1).height = 15*20
    xl_sheet.row(2).height_mismatch = True
    xl_sheet.row(2).height = 15*20
    xl_sheet.row(3).height_mismatch = True
    xl_sheet.row(3).height = 15*20
    xl_sheet.row(10).height_mismatch = True
    xl_sheet.row(10).height = 33*20

    # Get each site ID
    for i in range(start_of_bom_id, end_of_bom_id):
        for site in siteList:
            if site in bom[i]:
                site_id.append(i)

    # Get variables with site, siteID and endID
    for i in range(len(siteList)):
        site = siteList[i]
        siteID = site_id[i]
        try:
            endID = site_id[i+1]
        except IndexError:
            endID = end_of_bom_id

    # Define offset for first Site is static
        if i == 0:
            offset = 13
            # linesPerSite.append(offset)

    # For others sites just increase the value leaving two spaces
        else:
            offset += 2

        # Write the site name
        xl_sheet.row(offset).height_mismatch = True
        xl_sheet.row(offset).height = 19*19
        xl_sheet.write_merge(offset,offset, 4, 8, site, sites_red)
        offset += 1
        xl_sheet.write_merge(offset,offset, 4, 8, '', gray_background)
        # Add 5 spaces to offset the template
        offset += 2
        # Write part of the template
        xl_sheet.row(offset).height_mismatch = True
        xl_sheet.row(offset).height = 20*20
        xl_sheet.write_merge(offset, offset, 1, 4, 'Hardware y/o Garantias', title_strong)
        offset +=2
        # Write the headers
        headers = ['Item','# de Parte del Fabricante','Codigo Sinergia','Descripcion Corta',
        'Cantidad','Precio Unitario','Descuento','Precio Unitario con descuento','Total']
        for i in range(len(headers)):
            xl_sheet.write(offset, i+1, headers[i], header_normal)
        offset += 1
        # Clear subTotal and Item index
        subTotal = 0.0
        item = 1
        # Iterate within the range of items of each site
        for i in range(siteID+1,endID):
            # print bom[i]
            # Write each item
            if bom[i][5] == 0:
                continue

            xl_sheet.write(offset, 1, item, content_normal)
            xl_sheet.write(offset, 2, bom[i][1], content_normal)
            xl_sheet.write(offset, 3, bom[i][10], content_normal)
            xl_sheet.write(offset, 4, bom[i][11], content_normal)
            xl_sheet.write(offset, 5, bom[i][6], content_normal)
            xl_sheet.write(offset, 6, bom[i][5], content_normal)
            xl_sheet.write(offset, 7, bom[i][8], content_normal)
            xl_sheet.write(offset, 8, bom[i][7], content_normal)
            xl_sheet.write(offset, 9, bom[i][9], content_normal)
            offset += 1
            item += 1
            try:
                subTotal += bom[i][9]
            except:
                pass
        # Write Total for each site
        xl_sheet.write(offset, 8, 'Total: '+site, subTotal_stong)
        xl_sheet.write(offset, 9, subTotal, subTotal_stong)

        # Add the total to a tuple for resumen later
        subTotals.append((site, subTotal))
        # print "Sub Total:" + site +" "+ str(subTotal)
        # Add the Total to the grand total for resumen later
        granTotal += subTotal


    # Write Resumen Table
    offset = 6
    for site in subTotals:
        xl_sheet.write(offset, 13, 'Total: '+site[0], content_strong)
        xl_sheet.write(offset, 14, site[1], content_strong)
        offset +=1

    # Write Grand Total
    offset +=1

    xl_sheet.write(offset, 13, 'Total de la solucion', granTotal_strong)
    xl_sheet.write(offset, 14, granTotal, granTotal_strong)

    # print "Grand Total:" + str(granTotal)
    # Save the pricesheet to the newfile
    destinationwb.save(newfilename)
    return newfilename
