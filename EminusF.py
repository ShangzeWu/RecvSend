
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

path =  "/var/www/html/RecvSend/"
#print(path)
dir_E = path+'/uploadE/' #用来读取E文件 的 路径

dir_save_E = "/var/www/html/RecvSend/resultE/"  #输出 E文件 的保存路径

file_name_E = find_new_file(dir_E)

#业务逻辑

df = pd.read_excel(dir_E+file_name_E)
#print(df['寄件网点'])
#print((df['寄件网点']!='江苏省市场部五十七部') & (df['寄件网点']!='江苏盐城公司'))
df = df.drop(df[(df['寄件网点']!='江苏省市场部五十七部') & (df['寄件网点']!='江苏盐城公司') & (df['寄件网点']!='江苏盐城宝龙公司') & (df['寄件网点']!='江苏盐城龙冈公司') & (df['寄件网点']!='江苏盐城亭湖公司') & (df['寄件网点']!='江苏盐城万达公司') & (df['寄件网点']!='江苏盐城吾悦公司') & (df['寄件网点']!='江苏盐城盐都公司') & (df['寄件网点']!='江苏盐城盐南高新公司') & (df['寄件网点']!='江苏盐城招商公司')  ].index)
#删除空行
df = df.dropna(axis=0, how='all', thresh=None, subset=None, inplace=False)
# print(df['运单编号'].dtype) #int64

for num1 in list_number:
    num1 = int(num1)
    #print(type(num1))
    #print(df['运单编号'].dtype)
    df = df.drop(df[ df['运单编号'] == num1 ].index)

#df = df.drop(df[ df['运单编号'] == 777069457504657].index)
writer = pd.ExcelWriter(path+'/resultD/'+file_name_D)
#df为需要保存的DataFrame
df.to_excel(writer,index = False ,encoding='utf-8',sheet_name='Sheet1')
writer.save()
