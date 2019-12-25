# *_*coding:utf-8 *_*
# __author__ = 'GonzaloWang'
import json
import xlwt
import os
from xlutils.copy import copy
import xlrd

xls = xlwt.Workbook()
sht1 = xls.add_sheet('Sheet1')
sht2 = xls.add_sheet('sheet2')
sht1.write(0, 0, 'PLAN_ID')
sht1.write(0, 1, 'BOX_ID')
sht1.write(0, 2, 'METER_NO')
sht1.write(0, 3, 'SERIAL_NO')
sht1.write(0, 4, 'MADE_NO')
sht1.write(0, 5, 'POS_ROWS')
sht1.write(0, 6, 'POS_COLS')
sht1.write(0, 7, 'CONS_NO')
sht1.write(0, 8, 'CONS_NAME')
sht1.write(0, 9, 'ASSET_NO')
sht2.write(0, 0, 'BOX_ID')
sht2.write(0, 1, 'RTK_POINT')
sht2.write(0, 2, 'INST_LOC')
sht2.write(0, 3, 'ASSET_NO')
sht2.write(0, 4, 'BOX_ROWS')
sht2.write(0, 5, 'BOX_COLS')
sht2.write(0, 6, 'VOL_LEVEL')
sht2.write(0, 7, 'ORG_NAME')
xls.save('mydata.xls')
rb = xlrd.open_workbook('mydata.xls')
wb = copy(rb)
sht1 = wb.get_sheet(0)
sht2 = wb.get_sheet(1)
n = 1
k = 1
path = os.getcwd() + r'\\003txtfile'
for i in os.listdir(path):
    with open(path + r'\\' + i, "r", encoding='utf-8') as f:
        dic = f.read()
        a = json.loads(dic)
        b = a["RS_MEASBOX_REF_LIST"]
        c = a['RS_MEASBOX_LIST']
        for it in b:
            # print(i)
            # print(i['PLAN_ID'], i['PLAN_ID']+'-'+i['BOX_ID'], i['METER_NO'], i['SERIAL_NO'], i['MADE_NO'],
            # i['POS_ROWS'], i['POS_COLS'], i['CONS_NO'], i['CONS_NAME'], i['ASSET_NO'])
            sht1.write(n, 0, it['PLAN_ID'])
            sht1.write(n, 1, it['PLAN_ID'] + '-' + it['BOX_ID'])
            sht1.write(n, 2, it['METER_NO'])
            sht1.write(n, 3, it['SERIAL_NO'])
            sht1.write(n, 4, it['MADE_NO'])
            sht1.write(n, 5, it['POS_ROWS'])
            sht1.write(n, 6, it['POS_COLS'])
            sht1.write(n, 7, it['CONS_NO'])
            sht1.write(n, 8, it['CONS_NAME'])
            sht1.write(n, 9, it['ASSET_NO'])
            n += 1
        for d in c:
            sht2.write(k, 0, d['PLAN_ID'] + '-' + d['BOX_ID'])
            sht2.write(k, 1, d['RTK_POINT'])
            sht2.write(k, 2, d['INST_LOC'])
            sht2.write(k, 3, d['ASSET_NO'])
            sht2.write(k, 4, d['BOX_ROWS'])
            sht2.write(k, 5, d['BOX_COLS'])
            sht2.write(k, 6, d['VOL_LEVEL'])
            sht2.write(k, 7, d['ORG_NAME'])
            k += 1
wb.save('mydata.xls')
