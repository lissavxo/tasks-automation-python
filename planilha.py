import pywinauto
from pywinauto.application import Application
import xlrd 


print("start")


program_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
file_path    = r"C:\Users\lissa.oliveira\Documents\devDir\excel-table-example\planilha-teste.xlsx"

app = Application(backend="uia").start(r'{} "{}"'.format(program_path, file_path))

stylesheet =  pywinauto.timings.wait_until_passes(5,0.5, lambda: Application.connect(title_re="planilha-teste - Excel"))
print(stylesheet)


# app.stylesheet.print_control_identifiers()
# table.print_control_identifiers()


# screen.pywinauto.mouse.click(button='left', coords=(2979,265))
# # screen.pywinauto.mouse.scroll(coords=(3045,370), wheel_dist=1)

# # screen.Pane.wait
# # print(screen)

# #read that https://jpereiran.github.io/articles/2019/06/14/Excel-automation-with-pywin32.html

# wb = xlrd.open_workbook(file_path) 
# sheet = wb.sheet_by_index(0) 
  
# # For row 0 and column 0 
# sheet.cell_value(0, 0) 
  
# for i in range(sheet.ncols): 
#     print(type(sheet.cell_value(0, i))) 





