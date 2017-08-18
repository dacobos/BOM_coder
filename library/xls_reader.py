################################  MODULE  INFO  ################################
# Author: David  Cobos
# Cisco Systems Solutions Integrations Architect
# Mail: cdcobos1999@gmail.com  / dacobos@cisco.com
################################  MODULE  INFO  ################################

# Read a xls file and return a dictionary with all the values

# from openpyxl import Workbook
# from openpyxl import load_workbook

import xlrd

def getIds(filename):
        # Get the id of start and end
        book = xlrd.open_workbook(filename)
        # get the first worksheet
        xl_sheet = book.sheet_by_index(0)
        # print xl_sheet
        num_rows = xl_sheet.nrows

        for i in range(0, num_rows):
            values = xl_sheet.row_values(i)
            try:
                if 'Line Number' in values[0]:
                    start_of_bom_id = i+1
                    # print i
                if 'Valid through:' in values[0]:
                    end_of_bom_id = i
                    # print i
                if 'Total Price:' in values[7]:
                    total_price_id = i
                    # print i
            except:
                pass
        result = [start_of_bom_id, end_of_bom_id, total_price_id]
        return result

def readbom(filename):
    book = xlrd.open_workbook(filename)
    # get the first worksheet
    xl_sheet = book.sheet_by_index(0)
    # print xl_sheet
    num_rows = xl_sheet.nrows
    data = []
    product_total = 0.0
    total_price_id = 0

    Ids = getIds(filename)
    start_of_bom_id = Ids[0]
    end_of_bom_id = Ids[1]
    total_price_id  = Ids[2]

    for i in range(0, num_rows):
        values = xl_sheet.row_values(i)
        try:
            values[7] = round(values[5]-(values[5]*values[8])/100,2)
            values[9] = round(values[7]*values[6],2)
        except:
            pass
        data.append(values)

    for i in range(start_of_bom_id, end_of_bom_id):
        try:
            product_total += data[i][9]
        except:
            pass

                # values[8]
    data[end_of_bom_id][9] = product_total
    data[total_price_id][9] = product_total

    return data


def getSitelist(bom, filename):
    Ids = getIds(filename)
    start_of_bom_id = Ids[0]
    end_of_bom_id = Ids[1]
    total_price_id  = Ids[2]

    siteList = []
    for i in range(start_of_bom_id, end_of_bom_id):
        if bom[i][1] != '' and bom[i][2] == '' and bom[i][3] == '' and bom[i][4] == '' and bom[i][5] == '':
            siteList.append(bom[i][1])
    return siteList
