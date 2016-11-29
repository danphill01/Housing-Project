# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 14:00:12 2016

@author: Daniel
"""

import xlrd
from xlrd.sheet import ctype_text
import copy
from pandas import DataFrame
import pandas as pd

def getcell(sheet, row, col):
    cell_obj = sheet.cell(row, col)
    dtype = ctype_text.get(cell_obj.ctype,'unknown')
    if (dtype == 'text'):
        cell = cell_obj.value.strip()
        if cell == '-':
            cell = 0
            dtype = 'number'
    else:
        cell = cell_obj.value
    return cell, dtype

def clearlist(l):
    while len(l) > 0:
        l.pop()


inputfile = "C:\\Users\\Daniel\\Hack Oregon\\Housing\\affordable-housing-oregon-inventory.xls"
#headers = ['ID','NAME','TOTAL','AFFORDABLE','ADDRESS','CITY','ZIP','COUNTY','ADR','ALF','CC','CMI','DD','DV','ELD','EO/RO','FAM','FW','HIV','HOM','PD']

workbook = xlrd.open_workbook(inputfile)
sheet_names = workbook.sheet_names()
x1_sheet = workbook.sheet_by_index(0)
if x1_sheet is None:
    print ("sheet 0 does not exist")
else:
    rowtotal = x1_sheet.nrows
    coltotal = x1_sheet.ncols
    startReport = False
    lol = []
    rowlist = []
    header = []
    ident = []
    cell, dtype = getcell(x1_sheet,0,1)
    if dtype != 'empty':
        if (dtype == 'text'):
            for j in range(1,coltotal):
                cell, dtype = getcell(x1_sheet,0,j)
                if cell != '':
                    rowlist.append(cell)
            header=copy.copy(rowlist)
            clearlist(rowlist)    
    for i in range(1, rowtotal): #skip headers
        cell, dtype = getcell(x1_sheet,i,1)
        if dtype != 'empty':
            if (dtype == 'text'):
                cell, dtype = getcell(x1_sheet,i,0)
                if cell != '':
                    ident.append(cell)
                for j in range(1,coltotal):
                    cell, dtype = getcell(x1_sheet,i,j)
                    if cell != '':
                        rowlist.append(cell)
                lol.append(copy.copy(rowlist))
                clearlist(rowlist)

frame = DataFrame(lol,columns=header,index=ident)
