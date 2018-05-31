import cx_Oracle as oracle
import xlrd as reader
import os
import time
import numpy as np
from pandas import Series,DataFrame
import matplotlib.pyplot as plt
import io
import sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf-8')

cur_time = time.strftime("%Y%m%d",time.localtime())
file_path = "D:\\Excel"
table_name= "TR_STAT_CH"
insert_query = "INSERT INTO " + table_name + " (COD_TR,NAM_TR,SUM_TR,DATE_TR) VALUES ('%s','%s','%d','%s')"
count_gap = [10,10000,50000,100000,200000,400000,600000,800000,1000000]
count_list = [0,0,0,0,0,0,0,0,0,0]
horizon = ['0-10','10-1W','1W-5W','5W-10W','10W-20W','20W-40W','40W-60W','60W-80W','80W-100W','100W~']

def get_base(cursor):
    sql = "select COD_TR from TR_NAM_CH"
    cursor.execute(sql)
    data=cursor.fetchall()
    return data

def get_sheetname():
    f_list = os.listdir(file_path)
    if len(f_list) != 1 :
        print("Please confirm number of files under the directory!")
    return f_list[0]

def count_num(total):
    for i in range(len(count_gap)-1):
        if total < count_gap[i]:
            count_list[i] += 1
            return

    count_list[len(count_gap)] += 1

    return

def draw_bar_graph():
    df = DataFrame(count_list,columns=['volume of transaction'],index=horizon)
    df.plot(kind='bar')
    plt.show()

def type_transfer(base_sql):
    data = []
    for i in range(len(base_sql)):
        data.append(list(base_sql[i])[0])
    return data

if __name__=='__main__':
    ipaddr="10.230.171.180"
    username="avadbo"
    password="avadbop"
    oracle_port="1521"
    oracle_service="ora12c"
    try:
        connection = oracle.connect(username+"/"+password+"@"+ipaddr+":"+oracle_port+"/"+oracle_service,encoding="UTF-8",nencoding="UTF-8")
    except Exception as e:
        print(e)
    else:
        cursor = connection.cursor()
        base_sql = get_base(cursor)
        base = type_transfer(base_sql)
        #print(len(base))
        #print(isinstance(base,list))

    xls_name = get_sheetname()
    xls_data = reader.open_workbook(file_path+"\\"+xls_name)
    #sheetname = data.sheet_names()
    sheet = xls_data.sheet_by_name('BoEing交易量按渠道通道统计')
    row_num = sheet.nrows
    col_num = sheet.ncols
    for i in range(3,row_num):
        row = sheet.row_values(i)
        if row[0] in base :
            cod_tr_value = row[0]
            nam_tr_value = row[1]
            num_tr_value = int(row[4])
            date_tr_value = cur_time
            cursor.execute(insert_query%(cod_tr_value,nam_tr_value,num_tr_value,date_tr_value))
            count_num(num_tr_value)

    connection.commit()
    cursor.close()
    connection.close()

    draw_bar_graph()
