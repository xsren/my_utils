# encoding: utf-8
"""
@author: xsren 
@contact: bestrenxs@gmail.com
@site: xsren.me

@version: 1.0
@license: Apache Licence
@file: pack_multiple_excel_to_one.py
@time: 27/07/2017 5:22 PM

将当前目录下的excel合并成一个新的excel
"""
import os
import time

from openpyxl import Workbook
from openpyxl import load_workbook


def run():
    wb = Workbook()
    ws_w = wb.active
    fpath = '.'

    i = 1
    for fname in os.listdir(fpath):
        if fname.endswith(".xlsx"):
            wb_r = load_workbook(filename=fname)
            ws_r = wb_r.worksheets[0]
            j = 1
            print fname
            for row in ws_r.rows:
                if i == 1:
                    if j == 1:
                        ws_w.title = ws_r.title
                    n_row = [r.value for r in row]
                    ws_w.append(n_row)
                else:
                    if j >= 4:
                        n_row = [r.value for r in row]
                        ws_w.append(n_row)
                j += 1
            i += 1

    today = time.strftime('%Y%m%d_%H:%M:%S', time.localtime(time.time()))
    wb.save('new_file_%s.xlsx' % today)


if __name__ == '__main__':
    run()
