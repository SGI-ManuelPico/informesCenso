import os
import win32com.client

class Pdf:
    def excelPdf(self, excel_path, pdf_path):
        """
        Convierte un archivo Excel a PDF de forma silenciosa.
        """
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
        except:
            pass

        excel.DisplayAlerts = False

        workbook = None
        try:
            # Abrimos el workbook
            workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
            
            sheet = workbook.ActiveSheet

            # 2) Configuración de página
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False
            sheet.PageSetup.Orientation = 2 
            sheet.PageSetup.LeftMargin = 0.5
            sheet.PageSetup.RightMargin = 0.5
            sheet.PageSetup.TopMargin = 0.5
            sheet.PageSetup.BottomMargin = 0.5

            workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            print(f"[OK] PDF generado: {pdf_path}")

        except Exception as e:
            print(f"[ERROR] Al exportar a PDF: {e}")

        finally:
            if workbook is not None:
                # Decirle a Excel que no hay cambios pendientes de guardar
                workbook.Close(SaveChanges=False)
            excel.Quit()
