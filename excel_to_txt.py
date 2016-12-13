# ! c:/python27/python
# coding=utf-8

import os
import sys
import xlrd
import xlwt
import re
import codecs
'''
#全局编码方式为utf-8
default_encoding="utf-8"
if(default_encoding!=sys.getdefaultencoding()):
    reload(sys)
    sys.setdefaultencoding(default_encoding)
'''
'''	
#输入文件地址
excelfile_path=raw_input("please input excel file path,for example:c:/icctxt/fm91cp.xlsx \n")
while os.path.exists(excelfile_path)!=1:
	excelfile_path=raw_input("please input txt file path,for example:c:/icctxt/fm91cp1109.txt \n")
'''
excelfile_path=r"c:/icctxt/fm31cp.xlsx"
#读取文件
excelfile=xlrd.open_workbook(excelfile_path)
excel_sheets=excelfile.sheet_names()   #获取Excel文件sheets
sheet_num=len(excel_sheets)            #获取Excel文件sheet个数
#打开icctxt文件
'''
icctxtfile_path=raw_input("please input icctxt file folder,for example: c:/iccicctxt.xlsx \n")
(icctxt_dir,icctxt_name)=os.path.split(icctxtfile_path)
while os.path.exists(icctxtfile_path)==1:
	next_requst=raw_input("the file exits,please confirm do you want to replace the file.(y/n) \n")
	if next_requst=="y":
		os.remove(icctxtfile_path)
		icctxt=open(icctxtfile_path,"w+")
		break
	else:
		icctxt=open(icctxtfile_path,"w+")
'''
icctxtfile_path=r"c:/icctxt/test.txt"
icctxt=codecs.open(icctxtfile_path,"a+","utf-8")

for num in range(sheet_num):
	sheet_list="sheet"+str(num)
	sheet_list=excelfile.sheets()[num]
	num_rows=sheet_list.nrows  #行
	num_cols=sheet_list.ncols   #列
	for i in range(1,num_rows):
		for j in range(num_cols):
			if j==num_cols-1:
				content_end="END"+"\n"
			else:
				content_end=""
			content= sheet_list.cell_value(0,j)+"="+sheet_list.cell_value(i,j).lstrip(" ")+"\n"+content_end
			icctxt.write(content)
