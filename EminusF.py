
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
df1 = pd.DataFrame(columns=['运单号','包号','扫描网点','扫描类型','扫描时间','上传时间','上一站','下一站','收/派件员','扫描员','寄件网点','车辆任务号','重量','物品类别','快件类型','设备类型','问题件类型','问题件说明','合作品牌','寄件客户','长','宽','高']) 
#print(df[(df['扫描类型']=='网点收件') | (df['扫描类型']=='业务员收件')])
df1 = df1.append(df[((df['扫描类型']=='网点收件') | (df['扫描类型']=='业务员收件')) &((df['扫描网点']=='江苏省市场部五十七部')|(df['扫描网点']=='江苏盐城公司')|(df['扫描网点']=='江苏盐城宝龙公司')|(df['扫描网点']=='江苏盐城亭湖公司')|(df['扫描网点']=='江苏盐城万达公司')|(df['扫描网点']=='江苏盐城吾悦公司')|(df['扫描网点']=='江苏盐城龙冈公司')|(df['扫描网点']=='江苏盐城盐都公司')|(df['扫描网点']=='江苏盐城盐南高新公司')|(df['扫描网点']=='江苏盐城招商公司'))])
#print(df1)
df1 = df1.drop_duplicates(subset='运单号', keep='first', inplace=False)
df = df.drop_duplicates(subset='运单号', keep='first', inplace=False)

#df1 = df1.reset_index(drop=True)

List = df1['运单号'].values.tolist()
#print(List)
for X in List:
    df = df.drop(df[df['运单号']==X].index)

writer = pd.ExcelWriter(path+'/resultE/Changed'+file_name_E)
#df为需要保存的DataFrame
df.to_excel(writer,index = False ,encoding='utf-8',sheet_name='Sheet1')
writer.save()
