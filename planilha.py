import pywinauto
from pywinauto.application import Application
from pywinauto import Desktop
import time



print("start")


program_path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
file_path    = r"C:\Users\lissa.oliveira\Documents\devDir\excel-table-example\planilha-teste.xlsx"

app = Application(backend="uia").start(r'%s "%s"'%(program_path, file_path))

main_dialog =  pywinauto.timings.wait_until_passes(30,0.5, lambda: app.connect(best_match="planilha-teste"))

interface_dialog = main_dialog.window(best_match="planilha-teste",class_name="XLMAIN")


print("be prepared")


#interface_dialog.Pane6.Toolbar.print_control_identifiers()

#caminho para botao guia de arquivo
##interface_dialog.Pane6.Toolbar.Button8.click()



# escrevendo nas celulas 


grid_dialog = interface_dialog.TabControl2

#grid_dialog.print_control_identifiers()

#interface_dialog.TabControl2.print_control_identifiers()

grid_dialog.DataItem22.type_keys(r"it")
grid_dialog.DataItem22.type_keys("{ENTER}")
grid_dialog.DataItem23.type_keys(r"works")
grid_dialog.DataItem23.type_keys("{ENTER}")




