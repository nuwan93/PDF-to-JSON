import os, os.path
import win32com.client


xl=win32com.client.Dispatch("Excel.Application")
xl.Workbooks.Open(Filename=r"C:\Users\Nuwan\Documents\excelsheet.xlsm", ReadOnly=1)


directory = r"E:\Google Drive - AssetOwl\Test PCR data - rentfindinspector"

for root, subdirectories, files in os.walk(directory):
    
    has_converted = False
    has_word = False
    
    for file in files:
        if file.endswith(".docx"):
            has_word = True
        elif file.endswith(".xlsx"):
            has_converted = True
    
    if has_word and not has_converted:
        xl.Application.Run("excelsheet.xlsm!Module1.PDF_To_Excel", root, root)
        print("Converted folder : ", root)

xl.Application.Quit() 
del xl