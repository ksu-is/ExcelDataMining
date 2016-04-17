import os
import xlrd
import xlwt
import sys
import os.path
import types

#Get The Company Names From Excel 
companyNamesXlsxParh = "/Users/subvin/Desktop/allCompanys.xlsx"
companyNamesData = xlrd.open_workbook(companyNamesXlsxParh,encoding_override = "utf-8")
table = companyNamesData.sheets()[0]
companyNameData = table.col_values(0)

#Delete The Frist Item ,because The First Item Is Not The Company Name
del companyNameData[0]
#print Company Names 
#print companayNameData


# My Computer did't Installed Office excel , So Write The Following to gain The Indicators
rootdir2 = "/Users/subvin/Desktop/pagate24.xlsx"
book = xlrd.open_workbook(rootdir2,encoding_override = "utf-8")
table = book.sheets()[0]
titles = table.row_values(0)
del titles[0]
del titles[2]
del titles[2]
del titles[2]
del titles[2]
# Get The Indicator Of Each Company


# Write Excel Data Test  Begin
#book = xlwt.Workbook(encoding = 'utf-8',style_compression = 0)
#sheet = book.add_sheet('companary',cell_overwrite_ok = True)
#sheet.write(0,0,colNameData[2])
#book.save('/Users/wangyunfeng/Desktop/companarynames.xlsx')
#Write Excel Date Test End

rootdir = "/Users/subvin/Desktop/fi"
for i in xrange(0,len(companyNameData)):     # 
	book = xlwt.Workbook(encoding = 'utf-8',style_compression = 0)
	companyName = companyNameData[i]
	for parent,dirnames,filenames in os.walk(rootdir):
		sheetNum = 0
		for dirname in dirnames:
			sheetNum = sheetNum + 1
			subDocuPath = os.path.join(parent,dirname)
			sheet = book.add_sheet(dirname,cell_overwrite_ok = True)
			for x in xrange(0,len(titles)):
				sheet.write(0,x + 1,titles[x])
			for secondParent,secondDirname,xlsFileNames in os.walk(subDocuPath):
				docuNum = 0
				for xlsFile in xlsFileNames:
					if xlsFile == '.DS_Store':
						continue
					#Open The Excel Package Data
					xlsPath = os.path.join(secondParent,xlsFile)
					xlsData = xlrd.open_workbook(xlsPath)
					table = xlsData.sheets()[0]
					firstCol = table.col_values(0)

					row = 0
					# Traverse All The Company Names , And Choose Get The Row Data
					for z in xrange(0,len(firstCol)):
						indexName = firstCol[z]
						row = row + 1
						rowData = table.row_values(z) 
						# If Get The Same Name In Package excel 
						if indexName == companyName:
							# If Four Indicators are Not Null ,then add The Indicator massage To company excel
							if table.cell(z,1).value != '' and table.cell(z,2).value != '' and (type(rowData[2]) is types.FloatType) and table.cell(z,7).value!= '' and table.cell(z,8).value!='' :
								docuNum = docuNum + 1;
								sheet.write(docuNum,0,xlsFile)
								sheet.write(docuNum,1,table.cell(z,1).value)
								sheet.write(docuNum,2,table.cell(z,2).value)
								sheet.write(docuNum,3,table.cell(z,7).value)
								sheet.write(docuNum,4,table.cell(z,8).value)
								
							break
	savepath = '/Users/subvin/Desktop/CompanyInfo/%s.xlsx'%companyName
	book.save(savepath) 

#print row