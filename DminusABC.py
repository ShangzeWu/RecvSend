# -*- coding: UTF-8 -*-
from openpyxl import *
import os
import time
from datetime import datetime, timedelta
import pandas as pd
import xlrd

format_pattern = '%Y-%m-%d %H:%M:%S'
cur_time = datetime.now()
# 将 'cur_time' 类型时间通过格式化模式转换为 'str' 时间
cur_time = cur_time.strftime(format_pattern)

def find_new_file(dir):
    #查找目录下最新的文件
    file_lists = os.listdir(dir)
    file_lists.sort(key=lambda fn: os.path.getmtime(dir + "/" + fn)
    if not os.path.isdir(dir + "/" + fn) else 0)
#    print('最新的文件为： ' + file_lists[-1])
    file = os.path.join(dir, file_lists[-1])
#    print('完整路径：', file)
    return file_lists[-1]   #返回文件的名字，不包含路径
'''
def setNull(ws,n):
    cols_num = ws.max_column
    #print(cols_num)
    #print(" *********** ")
    for x in range(1,cols_num+1):
        ws.cell(row=n,column=x,value="")  #清空一行数据
'''

path =  "/var/www/html/RecvSend/"
#print(path)
dir_A = path+'/uploadA/' #用来读取A文件 的 路径
dir_B = path+'/uploadB/' #用来读取B文件 的 路径
dir_C = path+'/uploadC/' #用来读取C文件 的 路径
dir_D = path+'/uploadD/' #用来读取D文件 的 路径

dir_save_D= "/var/www/html/RecvSend/resultD/"  #输出 D文件 的保存路径

file_name_A = find_new_file(dir_A)
file_name_B = find_new_file(dir_B)
file_name_C = find_new_file(dir_C)
file_name_D = find_new_file(dir_D)

#业务逻辑
#加载ABC表的第一列
wb1 = load_workbook(dir_A+file_name_A) #A表
ws1 = wb1[wb1.sheetnames[0]]           #A表第一页
wb2 = load_workbook(dir_B+file_name_B) #B表
ws2 = wb2[wb2.sheetnames[0]]           #B表第一页
wb3 = load_workbook(dir_C+file_name_C) #C表
ws3 = wb3[wb3.sheetnames[0]]           #C表第一页

Allrow1 = ws1.max_row
#Allcol1 = ws1.max_column
Allrow2 = ws2.max_row
Allrow3 = ws3.max_row

list_number = ['9999999999999']

for n in range(2, Allrow1+1):   #读取A表的单号序列,存入列表
    value_number = ws1.cell(n,1).value
    if value_number == None:
        continue
    else:
        value_number = str(value_number)
        list_number.append(value_number)
wb1.close()        
#wb1.save(dir_A+file_name_A)

for m in range(2, Allrow2+1):   #读取B表的单号序列
    value_number = ws2.cell(m,1).value
    if value_number == None:
        continue
    else:
        value_number = str(value_number)
        list_number.append(value_number)

wb2.close() 
#wb2.save(dir_B+file_name_B)
        
for o in range(2, Allrow3+1):    #读取C表的单号序列
    value_number = ws3.cell(o,1).value
    if value_number == None:
        continue
    else:
        value_number = str(value_number)
        list_number.append(value_number)
wb3.close() 
#使用pandas剔除杂项
df = pd.read_excel(dir_D+file_name_D)
#print(df['寄件网点'])
#print((df['寄件网点']!='江苏省市场部五十七部') & (df['寄件网点']!='江苏盐城公司'))
#删除其他网点
df = df.drop(df[(df['寄件网点']!='江苏省市场部五十七部') & (df['寄件网点']!='江苏盐城公司') & (df['寄件网点']!='江苏盐城宝龙公司') & (df['寄件网点']!='江苏盐城龙冈公司') & (df['寄件网点']!='江苏盐城亭湖公司') & (df['寄件网点']!='江苏盐城万达公司') & (df['寄件网点']!='江苏盐城吾悦公司') & (df['寄件网点']!='江苏盐城盐都公司') & (df['寄件网点']!='江苏盐城盐南高新公司') & (df['寄件网点']!='江苏盐城招商公司')  ].index)
#删除空行
df = df.dropna(axis=0, how='all', thresh=None, subset=None, inplace=False)
# print(df['运单编号'].dtype) #int64

for str1 in list_number:
    df = df.drop(df[ int(str1) == df['运单编号'] ].index)

writer = pd.ExcelWriter(dir_D+file_name_D)
#df为需要保存的DataFrame
df.to_excel(writer,index = False ,encoding='utf-8',sheet_name='Sheet1')
writer.save()



#删除D表中ABC的重复项
'''
for y in range(2,Allrow4+1):
    if ws4.cell(y,1).value == None:
        continue
    else:
        value_numberD = ws4.cell(y,1).value
        value_numberD = str(value_numberD)
        for str1 in list_number:
            if str1 == value_numberD:
                setNull(ws4,y) #清除整行内容
                #ws4.cell(row=y,column=1,value="")  #清空单号
                break
                
wb4.save(dir_save_D+cur_time+'.xlsx')
'''
#wb2.save(dir_namelist+file_name_list)
