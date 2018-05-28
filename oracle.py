#!/usr/bin/python
#coding=utf-8

import cx_Oracle as oracle
import xlrd as reader
import os

file_path = "D:\\Excel"

def get_base(cursor):
    sql = "select COD_TR from TR_NAM_CH"
    cursor.execute(sql)
    data=cursor.fetchall()
    return data

def get_name():
    f_list = os.listdir(file_path)
    if len(f_list) != 1 :
        print("Please confirm number of files under the directory!")
    return f_list[0]

if __name__=='__main__':
    ipaddr=""
    username=""
    password=""
    oracle_port=""
    oracle_service=""
    try:
        db = oracle.connect(username+"/"+password+"@"+ipaddr+":"+oracle_port+"/"+oracle_service)
    except Exception as e:
        print(e)
    else:
        cursor = db.cursor()
        base = oraclesql(cursor)
        for i in base:
            print(i)
        cursor.close()
        db.close()

    xls_name = get_name()
    xls_data = reader.open_workbook(file_path+"\\"+xls_name)
    sheetname = data.sheet_names()
    sheet = xls_data.sheet_by_name('BoEing交易量按渠道统计')
    row_num = sheet.nrows
    col_num = sheet.ncols
    for i in range(row_num):
        row = sheet.row_values(i)
        if row(i)[0][0:6] in base :


