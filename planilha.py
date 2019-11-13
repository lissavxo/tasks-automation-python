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
time.sleep(3)
print("be prepared")
time.sleep(2)

#interface_dialog.Pane6.Toolbar.print_control_identifiers()

#caminho para botao guia de arquivo
##interface_dialog.Pane6.Toolbar.Button8.click()



