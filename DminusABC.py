# -*- coding: UTF-8 -*-
from openpyxl import *
import os
import time
from datetime import datetime, timedelta

format_pattern = '%Y-%m-%d %H:%M:%S'
cur_time = datetime.now()
# 将 'cur_time' 类型时间通过格式化模式转换为 'str' 时间
cur_time = cur_time.strftime(format_pattern)

def find_new_file(dir):
    '''查找目录下最新的文件'''
    file_lists = os.listdir(dir)
    file_lists.sort(key=lambda fn: os.path.getmtime(dir + "/" + fn)
    if not os.path.isdir(dir + "/" + fn) else 0)
#    print('最新的文件为： ' + file_lists[-1])
    file = os.path.join(dir, file_lists[-1])
#    print('完整路径：', file)
    return file_lists[-1]   #返回文件的名字，不包含路径

path =  "/var/www/html/RecvSend/"
#print(path)
dir_A = path+'/uploadA/' #用来读取A文件 的 路径
dir_B = path+'/uploadB/' #用来读取B文件 的 路径
dir_C = path+'/uploadC/' #用来读取C文件 的 路径
dir_D = path+'/uploadD/' #用来读取D文件 的 路径

#dir_namelist = path+'/namelist/' #用来读取人员名单表 的 路径
#dir_template = path+'/template/template.xlsx' #用来读取输出模版表格 的 路径
dir_save_D= "/var/www/html/RecvSend/resultD/"  #输出 D文件 的保存路径

file_name_A = find_new_file(dir_A)
file_name_B = find_new_file(dir_B)
file_name_C = find_new_file(dir_C)
file_name_D = find_new_file(dir_D)

#file_name_list = find_new_file(dir_namelist)

#业务逻辑
wb1 = load_workbook(dir_A+file_name_A) #A表
ws1 = wb1[wb1.sheetnames[0]]           #A表第一页
wb2 = load_workbook(dir_B+file_name_B) #B表
ws2 = wb2[wb2.sheetnames[0]]           #B表第一页
wb3 = load_workbook(dir_C+file_name_C) #C表
ws3 = wb3[wb3.sheetnames[0]]           #C表第一页
wb4 = load_workbook(dir_D+file_name_D) #D表
ws4 = wb4[wb4.sheetnames[0]]           #D表第一页

# wb2 = load_workbook(dir_namelist+file_name_list)      #namelist表
# ws2 = wb2[wb2.sheetnames[0]]           #namelist表第一页
# wb3 = load_workbook(dir_template)      #模板表

#读取名单表，写入数组
Allrow1 = ws1.max_row
Allcol1 = ws1.max_column
Allrow2 = ws2.max_row
Allcol2 = ws2.max_column
Allrow3 = ws3.max_row
Allcol3 = ws3.max_column
Allrow4 = ws4.max_row
print(Allrow4)
#Allcol2 = ws2.max_column
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
#wb3.save(dir_C+file_name_C)

#print(list_name)
print(len(list_number))
print(list_number)
#删除D表中的进港件记录
for x in range(2,Allrow4+1):
    if ws4.cell(x,11).value == None:
        ws4.cell(row=x,column=1,value="")  #清空单号
        #print(x)
        #x-=1
    else:
        sendpoint = ws4.cell(x,11).value
        sendpoint = str(sendpoint)
        if sendpoint != '江苏省市场部五十七部':# or sendpoint != '江苏盐城宝龙公司' or sendpoint != '江苏盐城公司' or sendpoint != '江苏盐城龙冈公司' or sendpoint != '江苏盐城亭湖公司' or sendpoint != '江苏盐城万达公司' or sendpoint != '江苏盐城吾悦公司' or sendpoint != '江苏盐城盐都公司' or sendpoint != '江苏盐城盐南高新公司' or sendpoint != '江苏盐城招商公司':
            ws4.cell(row=x,column=1,value="")  #清空单号
            #print(sendpoint)
            #print(x)
            #x-=1
            #print(x)
            #print(" *********** ")
    
Allrow4 = ws4.max_row
print(Allrow4)
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
                ws4.delete_rows(y)
                y-=1
                break
'''
#合并三个表
'''Allrow1 = ws1.max_row
Allcol1 = ws1.max_column

#print(Allrow1)
index_C_col=1
while index_C_col<=Allcol1:
        if ws1.cell(1,index_C_col).value=="派件员":
            index_C_row = 2
            while index_C_row<=Allrow1:
                name_C = ws1.cell(index_C_row,index_C_col).value
                flag = True
                if name_C == None:  #排除C表的空行
                    index_C_row += 1
                    continue
                else:
                    name_C = str(name_C)
                    if name_C in list_name:
                        index_C_row += 1
                        continue
                    else:
                        ws1.cell(row = index_C_row, column = index_C_col, value ='')
                        index_C_row+=1
        index_C_col=index_C_col+1
'''
wb4.save(dir_save_D+cur_time+'.xlsx')
#wb2.save(dir_namelist+file_name_list)
