from turtle import end_fill
from types import NoneType
import numpy as np
import openpyxl
#from openpyxl import load_workbook
import pandas as pd
import os
import datetime
from time import strftime

rows        = []
row         = []
filenames   = []
timestamps  = []
part_list   = []

# PARTS TO BE EVALUATED:

# 380LX152M250A052
# B0192-B
# TM8050H-8W
# MX1A-21NW
# MX1A-11NW
# B0203B
# STM32L011K4T6
# LM3488MM/NOPB

critical_parts = ['380LX152M250A052', 'B0192-BL', 'TM8050H-8W', 'MX1A-21NW', 'MX1A-11NW', 'B0203-BL', 'STM32L011K4T6', 'LM3488MM/NOPB']

def xlsx_name():
    directory = './Supply/4Front/Inventory Files/'
    global filenames
    global timestamps
    for filename in os.listdir(directory):
        f = os.path.join(directory, filename)
        d = os.path.getmtime(f)
        filenames.append(str(f[32:]))
        timestamps.append(str(d))

    for j in range(len(filenames)):
        index = j
        try:
            if timestamps[j + 1] > timestamps[j]:
                index = j + 1
        except:
            index = j
    return filenames[index]

def part_data(mfg_part_num, maxrow, sheet):
    for k in range(1,maxrow + 1):
        cell = sheet.cell(row = k, column = 9)
        if(str(cell.value) == mfg_part_num):
            part_no = str(mfg_part_num)
            cell = sheet.cell(row = k, column = 4)
            while cell.value == None:
                k = k - 1
                cell = sheet.cell(row = k, column = 4)
            qty_per = str(cell.value)
            cell = sheet.cell(row = k, column = 6)
            qty_on_hand = str(cell.value)
            return[part_no, qty_per, qty_on_hand]

def open_orders(mfg_part_num, maxrow, sheet):
    qty = []
    due = []
    for k in range(1,maxrow + 1):
        cell = sheet.cell(row = k, column = 9)
        if(str(cell.value) == mfg_part_num):
            qty.append(sheet.cell(row = k, column = 7).value)
            due.append(sheet.cell(row = k, column = 8).value)
    return [qty, due]

def excel_read(datafile):
    datafile = './Supply/4Front/Inventory Files/' + datafile
    wb = openpyxl.load_workbook(datafile) 
    sheet = wb.active 
    maxrow = sheet.max_row
    return [wb, sheet, maxrow]

def get_parts(maxrow, sheet):
    part_list = []
    for k in range(3, maxrow + 1):
        cell = sheet.cell(row = k, column = 9)
        if str(cell.value) != 'None' or cell.value != None:
            part_list.append(cell.value)
    return(eliminate_duplicates(part_list))

def get_sheets(part_list):
    invalid_characters = ('\\', '/', '*', '?', ':', '[',']')        #  \ , / , * , ? , : , [ , ]
    sheet_name = []
    for part in part_list:
        for k in invalid_characters:
            if k in str(part):
                sheet_name.append(str(part).replace(k,' '))
                break
            else:
                sheet_name.append(str(part))
                break
    return(sheet_name)

def eliminate_duplicates(dataset):
    final_dataset = []
    for k in range(1,len(dataset)):
        if (dataset[k-1] != dataset[k]):
            final_dataset.append(dataset[k-1])
    return final_dataset

file_4Front                             = xlsx_name()
[wbk_4Front, sht_4Front, maxrow_4Front] = excel_read(file_4Front)
part_list                               = get_parts(maxrow_4Front, sht_4Front)
[qty, due]                              = open_orders(part_list[16], maxrow_4Front, sht_4Front)
wbk                                     = openpyxl.load_workbook('./Supply/4Front/Summary.xlsx')
wbksheet                                = wbk.active 
sheet_names                             = get_sheets(part_list)


#for part in part_list:
#    print(part)
for sheet in sheet_names:
    print(sheet)


wbk.save('./Supply/4Front/Summary.xlsx')
wbk.close()

"""
for k in parts: # critical components
    [part_no, qty_per, qty_on_hand] = part_data(k, maxrow, sheet)
    print(part_no)
    print(qty_per)
    print(qty_on_hand)
    """