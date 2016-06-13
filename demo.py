# -*-coding: utf-8 -*-
'''
operation of excel by using python and win32 com
Install by:
	https://sourceforge.net/projects/pywin32/files/pywin32/
Docs:
    1.https://github.com/pythonexcels/examples
    2.http://pythonexcels.com/python-excel-mini-cookbook/

DCOM can not see excel or word:
    1. https://blogs.technet.microsoft.com/the_microsoft_excel_support_team_blog/2012/11/12/microsoft-excel-or-microsoft-word-does-not-appear-in-dcom-configuration-snap-in/
    2. http://www.win7qjb.com/win7XiTongAnZhuang/729.html
    3. http://bbs.csdn.net/topics/350263923
'''
import sys
reload(sys)
sys.setdefaultencoding('utf8')
import win32com.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')

workbook = excel.Workbooks.Open(r'C:\Template.xls')
sheet = workbook.Sheets('Sheet')

print sheet.Range(sheet.Cells(4, 'A'), sheet.Cells(5, 'C'))

sheet.Range('D4:O5').Value = 'xxxxxxxxxxxxxxxxxx'
sheet.Range('D4:O5').VerticalAlignment = win32.constants.xlCenter
sheet.Range('D4:O5').HorizontalAlignment = win32.constants.xlCenter

workbook.SaveAs(r'C:\new2.xls')
excel.Application.Quit()
