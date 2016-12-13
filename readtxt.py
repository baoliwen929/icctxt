#! c:\python27\python
# coding=utf-8
from __future__ import division
import re
import os
import xlrd
import xlwt
import sys
#更改全局默认编码方式为UTF-8
default_encoding="utf-8"
if(default_encoding!=sys.getdefaultencoding()):
    reload(sys)
    sys.setdefaultencoding(default_encoding)

#输入文件地址
txtfile_path=raw_input("please input txt file path,for example:c:/icctxt/fm91cp1109.txt \n")
while os.path.exists(txtfile_path)!=1:
	txtfile_path=raw_input("please input txt file path,for example:c:/icctxt/fm91cp1109.txt \n")
	

#文件读取
txtfile=open(txtfile_path,"r") 
lines=txtfile.readlines()                      #读取文件，将文件读取为一个列表
length=len(lines)                              #获取文件行数
#print length

#创建EXCEL文件
workbook=xlwt.Workbook(encoding="utf-8")
AIN=workbook.add_sheet("AIN")
CIN=workbook.add_sheet("CIN")
COUT=workbook.add_sheet("COUT")
AOUT=workbook.add_sheet("AOUT")
AOUTR=workbook.add_sheet("AOUTR")
CINR=workbook.add_sheet("CINR")
RIN=workbook.add_sheet("RIN")
COUTR=workbook.add_sheet("COUTR")

#style=xlwt.XFStyle()
#front=xlwt.Font()
#front.name='Simsun'
#style.front=front

#判断有多少个可用块
i=0
AIN_num=0
CIN_num=0
COUT_num=0
AOUT_num=0
AOUTR_num=0
CINR_num=0
RIN_num=0
COUTR_num=0
j=0
while (j<length):    #遍历txt文件
	if ("TYPE" in lines[j]):
		if ("AIN" in lines[j]):
			AIN_num=AIN_num+1
		elif ("CIN" in lines[j]):
			CIN_num=CIN_num+1
		elif ("COUT" in lines[j]):
			COUT_num=COUT_num+1
		elif ("AOUT" in lines[j]):
			AOUT_num=AOUT_num+1
		elif ("RIN" in lines[j]):
			RIN_num=RIN_num+1
		elif ("CINR" in lines[j]):
			CINR_num=CINR_num+1		
		elif ("AOUTR" in lines[j]):
			AOUTR_num=AOUTR_num+1			
		elif ("COUTR" in lines[j]):
			COUTR_num=COUTR_num+1		
		else:
			pass
	else:
		pass
	j=j+1
#print AIN_num
#print CIN_num
#print COUT_num
#print AOUT_num

AIN_row=1
CIN_row=1
COUT_row=1
AOUT_row=1
CINR_row=1
COUTR_row=1
AOUTR_row=1
RIN_row=1

for i in range(0,length):
	
	if ("TYPE" in lines[i]):
		if ("AIN" in lines[i]):        #获取标志行type=
			AIN_list=[]                 #定义一个列表()
			AIN_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+10],lines[i+11],lines[i+12], \
			lines[i+30],lines[i+31],lines[i+38],lines[i+40],lines[i+41],lines[i+42],lines[i+43],lines[i+47], \
			lines[i+48],lines[i+49],lines[i+50],lines[i+51]]) #将获取的符合条件的行赋值到新列表
			AIN_list_no_equal=[]           #定义一个没有等号的空列表
			
			for AIN_list_element in AIN_list:  
				AIN_list_element=AIN_list_element.split("=")[1]  #删除等号及前面数据
				AIN_list_element=AIN_list_element.split("\n")[0]  #删除换行符\n
				AIN_list_no_equal.append(AIN_list_element)   #将没有等号的新数据放入定义列表
			AIN_col=0
			#print AIN_list_no_equal
			while AIN_col<len(AIN_list_no_equal):                 #循环将数据放入EXCEl中
				AIN.write(AIN_row,AIN_col,AIN_list_no_equal[AIN_col])
				AIN_col=AIN_col+1
			del AIN_list
			AIN_row=AIN_row+1			
	
		elif (("CIN" in lines[i]) and ("CINR" not in lines[i])):  #CIN
			CIN_list=[]                 
			CIN_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+10],lines[i+11], \
			lines[i+19],lines[i+22],lines[i+23]]) 
			CIN_list_no_equal=[]           
			for CIN_list_element in CIN_list:  
				CIN_list_element=CIN_list_element.split("=")[1]  
				CIN_list_element=CIN_list_element.split("\n")[0]  
				CIN_list_no_equal.append(CIN_list_element)  
			CIN_col=0
			#print CIN_list_no_equal
			while CIN_col<len(CIN_list_no_equal):                 
				CIN.write(CIN_row,CIN_col,CIN_list_no_equal[CIN_col])
				CIN_col=CIN_col+1
			del CIN_list
			CIN_row=CIN_row+1	
		
		elif(("COUT" in lines[i]) and ("COUTR" not in lines[i])):   #COUT 
			COUT_list=[]                 
			COUT_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+13], \
			lines[i+19],lines[i+20]]) 
			COUT_list_no_equal=[]           
			for COUT_list_element in COUT_list:  
				COUT_list_element=COUT_list_element.split("=")[1]  
				COUT_list_element=COUT_list_element.split("\n")[0]  
				COUT_list_no_equal.append(COUT_list_element)  
			COUT_col=0
			#print COUT_list_no_equal
			while COUT_col<len(COUT_list_no_equal):                 
				COUT.write(COUT_row,COUT_col,COUT_list_no_equal[COUT_col])
				COUT_col=COUT_col+1
			del COUT_list
			COUT_row=COUT_row+1	
		elif (("AOUT" in lines[i]) and ("AOUTR" not in lines[i])):     #AOUT 
			AOUT_list=[]                 
			AOUT_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+11], \
			lines[i+43],lines[i+44]]) 
			AOUT_list_no_equal=[]           
			for AOUT_list_element in AOUT_list:  
				AOUT_list_element=AOUT_list_element.split("=")[1]  
				AOUT_list_element=AOUT_list_element.split("\n")[0]  
				AOUT_list_no_equal.append(AOUT_list_element)  
			AOUT_col=0
			#print AOUT_list_no_equal
			while AOUT_col<len(AOUT_list_no_equal):                 
				AOUT.write(AOUT_row,AOUT_col,AOUT_list_no_equal[AOUT_col])
				AOUT_col=AOUT_col+1
			del AOUT_list
			AOUT_row=AOUT_row+1		
		elif ("AOUTR" in lines[i]):     #AOUTR 
			AOUTR_list=[]                 
			AOUTR_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+10],lines[i+12], \
			lines[i+44],lines[i+45]]) 
			AOUTR_list_no_equal=[]           
			for AOUTR_list_element in AOUTR_list:  
				AOUTR_list_element=AOUTR_list_element.split("=")[1]  
				AOUTR_list_element=AOUTR_list_element.split("\n")[0]  
				AOUTR_list_no_equal.append(AOUTR_list_element)  
			AOUTR_col=0
			#print AOUTR_list_no_equal
			while AOUTR_col<len(AOUTR_list_no_equal):                 
				AOUTR.write(AOUTR_row,AOUTR_col,AOUTR_list_no_equal[AOUTR_col])
				AOUTR_col=AOUTR_col+1
			del AOUTR_list
			AOUTR_row=AOUTR_row+1	
		elif ("RIN" in lines[i]):     #RIN 
			RIN_list=[]                 
			RIN_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+10],lines[i+11]]) 
			RIN_list_no_equal=[]           
			for RIN_list_element in RIN_list:  
				RIN_list_element=RIN_list_element.split("=")[1]  
				RIN_list_element=RIN_list_element.split("\n")[0]  
				RIN_list_no_equal.append(RIN_list_element)  
			RIN_col=0
			#print RIN_list_no_equal
			while RIN_col<len(RIN_list_no_equal):                 
				RIN.write(RIN_row,RIN_col,RIN_list_no_equal[RIN_col])
				RIN_col=RIN_col+1
			del RIN_list
			RIN_row=RIN_row+1
		elif ("CINR" in lines[i]):     #CINR 
			CINR_list=[]                 
			CINR_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+10],lines[i+11],lines[i+12], \
			lines[i+23],lines[i+24]]) 
			CINR_list_no_equal=[]           
			for CINR_list_element in CINR_list:  
				CINR_list_element=CINR_list_element.split("=")[1]  
				CINR_list_element=CINR_list_element.split("\n")[0]  
				CINR_list_no_equal.append(CINR_list_element)  
			CINR_col=0
			#pCINRt CINR_list_no_equal
			while CINR_col<len(CINR_list_no_equal):                 
				CINR.write(CINR_row,CINR_col,CINR_list_no_equal[CINR_col])
				CINR_col=CINR_col+1
			del CINR_list
			CINR_row=CINR_row+1
		elif ("COUTR" in lines[i]):     #COUTR 
			COUTR_list=[]                 
			COUTR_list.extend([lines[i-1],lines[i],lines[i+1],lines[i+2],lines[i+5], \
			lines[i+6],lines[i+7],lines[i+8],lines[i+9],lines[i+18],lines[i+20]]) 
			COUTR_list_no_equal=[]           
			for COUTR_list_element in COUTR_list:  
				COUTR_list_element=COUTR_list_element.split("=")[1]  
				COUTR_list_element=COUTR_list_element.split("\n")[0]  
				COUTR_list_no_equal.append(COUTR_list_element)  
			COUTR_col=0
			#pCOUTRt COUTR_list_no_equal
			while COUTR_col<len(COUTR_list_no_equal):                 
				COUTR.write(COUTR_row,COUTR_col,COUTR_list_no_equal[COUTR_col])
				COUTR_col=COUTR_col+1
			del COUTR_list
			COUTR_row=COUTR_row+1				
		else:
			pass
	else:
		pass
	#进度条显示		
	sys.stdout.write(str(int((i/(length-1))*100))+'% ||'+'->'+"\r")
	sys.stdout.flush()
print

#AIN title 内容	
AIN_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","PNT_NO","SCI","HSCO1","LSCO1","DELTO1", \
"E01","BAO","BAT","HLOP","HAL","HAT","LAL","LAT","HHAOPT","HHALIM","HHATXT","LLALIM","LLATXT"]
AIN_title_num=0
while AIN_title_num<len(AIN_title):
	AIN.write(0,AIN_title_num,AIN_title[AIN_title_num])
	AIN_title_num=AIN_title_num+1

#CIN title 内容	
CIN_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","PNT_NO","ANM","NM0","NM1","IVO","SAO", \
"BAO","BAT"]
CIN_title_num=0
while CIN_title_num<len(CIN_title):
	CIN.write(0,CIN_title_num,CIN_title[CIN_title_num])
	CIN_title_num=CIN_title_num+1

#COUT title 内容	
COUT_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","PNT_NO","IN","INVCO","BAO","BAT"]
COUT_title_num=0
while COUT_title_num<len(COUT_title):
	COUT.write(0,COUT_title_num,COUT_title[COUT_title_num])
	COUT_title_num=COUT_title_num+1

#AOUT title 内容	
AOUT_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","PNT_NO","SCO","ATC","MEAS","BAO","BAT"]
AOUT_title_num=0
while AOUT_title_num<len(AOUT_title):
	AOUT.write(0,AOUT_title_num,AOUT_title[AOUT_title_num])
	AOUT_title_num=AOUT_title_num+1	
#AOUTR title 内容	
AOUTR_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","IOMIDR","PNT_NO","SCO","ATC","MEAS","BAO","BAT"]
AOUTR_title_num=0
while AOUTR_title_num<len(AOUTR_title):
	AOUTR.write(0,AOUTR_title_num,AOUTR_title[AOUTR_title_num])
	AOUTR_title_num=AOUTR_title_num+1	

#COUTR title 内容	
COUTR_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","IOMIDR","PNT_NO","IN","BAO","BAT"]
COUTR_title_num=0
while COUTR_title_num<len(COUTR_title):
	COUTR.write(0,COUTR_title_num,COUTR_title[COUTR_title_num])
	COUTR_title_num=COUTR_title_num+1
#CINR title 内容	
CINR_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","IOMIDR","PNT_NO","ANM","NM0","NM1","IVO","SAO", \
"BAO","BAT"]
CINR_title_num=0
while CINR_title_num<len(CINR_title):
	CINR.write(0,CINR_title_num,CINR_title[CINR_title_num])
	CINR_title_num=CINR_title_num+1
#RIN title 内容	
RIN_title=["NAME","TYPE","DESCRIP","PERIOD","IOMOPT","IOM_ID","PNT_NO","SCI","HSCI1","LSCI1"]
RIN_title_num=0
while RIN_title_num<len(RIN_title):
	RIN.write(0,RIN_title_num,RIN_title[RIN_title_num])
	RIN_title_num=RIN_title_num+1
	
txtfile.close()
#输出文件目录
excelfile_path=raw_input("please input excel file folder,for example: c:/iccexcel.xlsx \n")
(excel_dir,excel_name)=os.path.split(excelfile_path)
while os.path.exists(excelfile_path)==1:
	next_requst=raw_input("the file exits,please confirm do you want to replace the file.(y/n) \n")
	if next_requst=="y":
		os.remove(excelfile_path)
		workbook.save(excelfile_path)
		break
	else:
		excelfile_path=raw_input("please input excel file folder,for example: c:/iccexcel.xlsx \n")
		(excel_dir,excel_name)=os.path.split(excelfile_path)
os.mkdir(excel_dir)
workbook.save(excelfile_path)




























