# coding=utf-8#for import modlesimport re import osimport xlrdimport xlwtimport sysimport msvcrt#change all enviroment to UTF-8default_encoding="utf-8"if(default_encoding!=sys.getdefaultencoding()):    reload(sys)    sys.setdefaultencoding(default_encoding)	class AIN:	def _init_(self,NAME,TYPE,DESCRP,PERIOD,PHASE,LOOPID, \
	IOMOPT,IOM_ID,PNT_NO,SCI,HSCO1,LSCO1,DELTO1,EO1,OSV,EXTBLK, \
	MA,INITMA,BADOPT,LASTGV,INHOPT,INHIB,INHALM,MANALM,MTRF,FLOP, \
	FTIM,XREFOP,XREFIN,KSCALE,BSCALE,BAO,BAT,BAP,BAG,ORAO,ORAT,ORAP, \
	ORAG,HLOP,ANM,HAL,HAT,LAL,LAT,HLDB,HLPR,HLGP,HHAOPT,HHALIM,HHATXT, \
	LLALIM,LLATXT,HHAPRI,HHAGRP,PROPT,MEAS,AMRTIN,NASTDB,NASOPT):
		self.NAME=NAME
		self.TYPE=TYPE
		self.DESCRP=DESCRP
		self.PERIOD=PERIOD
		self.PHASE=PHASE
		self.LOOPID=LOOPID
		self.IOMOPT=IOMOPT
		self.IOM_ID=IOM_ID
		self.PNT_NO=PNT_NO
		self.SCI=SCI
		self.HSCO1=HSCO1
		self.LSCO1=LSCO1
		self.DELTO1=DELTO1
		self.EO1=EO1
		self.OSV=OSV
		self.EXTBLK=EXTBLK
		self.MA=MA
		self.INITMA=INITMA
		self.BADOPT=BADOPT
		self.LASTGV=LASTGV
		self.INHOPT=INHOPT
		self.INHIB=INHIB
		self.INHALM=INHALM
		self.MANALM=MANALM
		self.MTRF=MTRF
		self.FLOP=FLOP
		self.FTIM=FTIM
		self.XREFOP=XREFOP
		self.XREFIN=XREFIN
		self.KSCALE=KSCALE
		self.BSCALE=BSCALE
		self.BAO=BAO
		self.BAT=BAT
		self.BAP=BAP
		self.BAG=BAG
		self.ORAO=ORAO
		self.ORAT=ORAT
		self.ORAP=ORAP
		self.ORAG=ORAG
		self.HLOP=HLOP
		self.ANM=ANM
		self.HAL=HAL
		self.HAT=HAT
		self.LAL=LAL
		self.LAT=LAT
		self.HLDB=HLDB
		self.HLPR=HLPR
		self.HLGP=HLGP
		self.HHAOPT=HHAOPT
		self.HHALIM=HHALIM
		self.HHATXT=HHATXT
		self.LLALIM=LLALIM
		self.LLATXT=LLATXT
		self.HHAPRI=HHAPRI
		self.HHAGRP=HHAGRP
		self.PROPT=PROPT
		self.MEAS=MEAS
		self.AMRTIN=AMRTIN
		self.NASTDB=NASTDB
		self.NASOPT=NASOPT

#input txt file path
txtfile_path=raw_input("please input txt file path,for example:c:/icctxt/fm91cp1109.txt \n")
while os.path.exists(txtfile_path)!=1:
	txtfile_path=raw_input("please input txt file path,for example:c:/icctxt/fm91cp1109.txt \n")


#reading txt file
try:
	txtfile=open(txtfile_path,"r")
except IOError:
	print "open file error,please check the path and name."
else:
	print "open file success"
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

