# -*- coding: utf-8 -*-

from bs4 import BeautifulSoup

def xml_to_sheet_array(path):
    with open(path, 'r', encoding = 'utf-8') as data_file:
    #file = open(path, 'r', encoding = 'utf-8').read()
        soup = BeautifulSoup(data_file, 'xml')

    #  Return an array with 3 layers: sheet, row and cell:
    #  NOTE: if xml content contains the string 'Worksheet', 'Row' or 'Cell', you know that......
    
    workbook = []
    for sheet in soup.findAll('Worksheet'):
        sheet_index = []
        for row in sheet.findAll('Row'):
            row_index = []
            for cell in row.findAll('Cell'):
                if cell.Data:
                    row_index.append(cell.Data.text)
            sheet_index.append(row_index)
        workbook.append(sheet_index)
    
    return workbook
