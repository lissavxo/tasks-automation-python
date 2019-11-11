import pywinauto
from pywinauto.application import Application
# import xlrd 


print("start")


program_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
file_path    = r"C:\Users\lissa.oliveira\Documents\devDir\excel-table-example\planilha_teste.xlsx"

app = Application().start(r'{} "{}"'.format(program_path, file_path))

screen = app.window(title='planilha_teste - Excel')
print(screen)
# wb = xlrd.open_workbook(file_path) 
# sheet = wb.sheet_by_index(0) 
  
# # For row 0 and column 0 
# sheet.cell_value(0, 0) 
  
# for i in range(sheet.ncols): 
#     print(type(sheet.cell_value(0, i))) 