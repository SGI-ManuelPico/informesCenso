import os
import win32com.client

class Pdf:
    def excelPdf(self, excel_path, pdf_path):
        """
        Convierte un archivo Excel a PDF ajustando la configuración de página.
        """
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # No mostrar la ventana de Excel

        try:
            # Abrir el archivo
            workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
            sheet = workbook.ActiveSheet

            # Ajustar configuración de página
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1  
            sheet.PageSetup.FitToPagesTall = False
            sheet.PageSetup.Orientation = 2  
            sheet.PageSetup.LeftMargin = 0.5
            sheet.PageSetup.RightMargin = 0.5
            sheet.PageSetup.TopMargin = 0.5
            sheet.PageSetup.BottomMargin = 0.5

            # Exportar a PDF
            workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  
            print(f"Archivo convertido a PDF correctamente: {pdf_path}")

        except Exception as e:
            print(f"Error al exportar a PDF: {e}")
        finally:
            workbook.Close(False)
            excel.Quit()
