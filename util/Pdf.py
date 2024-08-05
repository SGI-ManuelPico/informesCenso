import win32com.client as win32
import os 

class Pdf:

    def excelPdf(self, excel_path, pdf_path):
        excel = win32.Dispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.abspath(excel_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        wb.Close()
        excel.Quit()