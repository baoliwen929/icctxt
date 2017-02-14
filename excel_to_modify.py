# ! c:/python27/python
# coding=utf-8

import os
import sys
import xlrd
import xlwt
import re
import codecs
import msvcrt
print """
#######################################################################################################
#   this program can generate an txt file which can import to CP with iccdrvr command from excel file.#
#   the excel file must respect certain rules:                                                        #
#   the first colum must be COMPND,the second is TAGNAME,the third COMPND:TAGNAME.                    #
#   the Remaining colums have no limit,but you must ensure the paraments is right.                    #
#   the txt file use modify command, so it don't have "TYPE" and "COMPND"                             #
#   when you input the path of file, you must type entire path and name,also suffix.                  #
#   any quetion or suggestion please cantact me with e_mail : liwen.bao@schneider-electric.com.       #
#######################################################################################################
"""
default_encoding="utf-8"
if(default_encoding!=sys.getdefaultencoding()):
    reload(sys)
    sys.setdefaultencoding(default_encoding)

#输入文件地址
excelfile_path=raw_input("please input excel file path,for example:c:/icctxt/fm91cp.xls \n")
while os.path.exists(excelfile_path)!=1:
	print "file path is error, please check again"
	excelfile_path=raw_input("please input excel file path,for example:c:/icctxt/fm91cp.xls \n")

#读取文件
try:
	excelfile=xlrd.open_workbook(excelfile_path)
except IOError:
	print "open file error,please check the path and name."
else:
	print "open file success"	
excel_sheets=excelfile.sheet_names()   #获取Excel文件sheets
sheet_num=len(excel_sheets)            #获取Excel文件sheet个数
#打开icctxt文件

icctxtfile_path=raw_input("please input icctxt file folder,for example: c:/icctxt/fam91cp.txt \n")
(icctxt_dir,icctxt_name)=os.path.split(icctxtfile_path)
while os.path.exists(icctxtfile_path)==1:
	next_requst=raw_input("the file exits,please confirm do you want to replace the file.(y/n) \n")
	if next_requst=="y":
		os.remove(icctxtfile_path)
		try:
			icctxt=open(icctxtfile_path,"w+")
		except IOError:
			print "open file error,please check the path and name."
		else:
			print "open file success"
		break
	else:
		icctxtfile_path=raw_input("please input icctxt file folder,for example: c:/icctxt/fm91cp.txt \n")
if os.path.exists(icctxt_dir):
	icctxt=codecs.open(icctxtfile_path,"a+","utf-8")
else:
	os.mkdir(icctxt_dir)
	icctxt=codecs.open(icctxtfile_path,"a+","utf-8")

for num in range(sheet_num):
	sheet_list="sheet"+str(num)
	sheet_list=excelfile.sheets()[num]
	num_rows=sheet_list.nrows  #行
	num_cols=sheet_list.ncols   #列
	for i in range(1,num_rows):
		for j in range(2,num_cols):
			if j==num_cols-1:
				content_end="END"+"\n"
			else:
				content_end=""
			if type(sheet_list.cell_value(i,j))==float:
				value=str(int(sheet_list.cell_value(i,j)))
			else:
				value=sheet_list.cell_value(i,j)
			content=str(sheet_list.cell_value(0,j)).strip()+"="+value.strip()+"\n"+content_end
			name_add=content.replace("NAME=","MODIFY ")
			name_no_type=re.sub("TYPE=.*\n","",name_add)
			icctxt.write(name_no_type)
print "press any key to exit"
print ord(msvcrt.getch())