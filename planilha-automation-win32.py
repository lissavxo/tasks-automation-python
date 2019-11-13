'''
helpers:
http://raaviblog.com/python-2-7-read-and-write-excel-file-with-win32com/
https://jpereiran.github.io/articles/2019/06/14/Excel-automation-with-pywin32.html

'''

import win32com.client
import sys,io


excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
# Select a file and open it
file = r"C:\Users\lissa.oliveira\Documents\devDir\excel-table-example\planilha-teste.xlsx"
workbook = excel.Workbooks.Open(file)
wb_data = excel.Workbooks.Open(file) 

readData = wb_data.Worksheets('Contas a Pagar')
allData = readData.UsedRange
for data in allData:
    if data != 'None':
        print(data)
#print ("Data on selected sheet : ",allData)




##oe1=wb_data.Worksheets("Contas a Pagar").Range("A2")
  
# Get the answers to the Q1A and write them into the summary file
# mission=wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("C6")
# vision =wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("C7")
# print("Question 1A")
# print("Mission:",mission)
# print("Vision:" ,vision)
# print()

# Get the answers to the Q1B and write them into the summary file
#oe1=wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("C14")
# ju1=wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("D14")
# oe2=wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("C15")
# ju2=wb_data.Worksheets("1ayb_MisiónyVisiónFutura").Range("D15")
# print("Question 1B")
# print("Table:",oe1)
# print("OEN2:",oe2, "- JUSTIF:",ju2)
# print()
    
# Close the file without saving
wb_data.Close(True)

# Wait before closing it
_ = input("Press enter to close Excel")
excel.Quit()
 