
# _*_ coding:utf-8 _*_

import xlrd, json
from pymongo import MongoClient


json_file = 'readme.json'
bts_info_file = 'LST BTS.xlsx'

site_area = []

wb = xlrd.open_workbook(bts_info_file)
sheet = wb.sheet_by_index(0)

for i in range(1, sheet.nrows):
    site = {
        '_id' : int(sheet.cell_value(i,0)),
        'name' : sheet.cell_value(i,1),
        'area' : sheet.cell_value(i,3),
        'netType' : 'HW3',
        'alarm_pool' : {},
        'reason_pool' : {},
    }
    print(int(sheet.cell_value(i,0)), sheet.cell_value(i,1), sheet.cell_value(i,3))
    if i > 10 : break

#with open(json_file, "r") as fp :
#    d  = json.load(fp)