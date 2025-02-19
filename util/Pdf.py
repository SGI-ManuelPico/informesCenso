import os
import win32com.client

class Pdf:
    def excelPdf(self, excel_path, pdf_path):
        """
        Convierte un archivo Excel a PDF de forma silenciosa,
        sin alterar los anchos ni altos de la plantilla.
        """
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
        except:
            pass

        excel.DisplayAlerts = False
        workbook = None
        try:
            # 1) Abre el workbook
            workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
            sheet = workbook.ActiveSheet

            # 2) Asegúrate de no establecer área de impresión limitada
            sheet.PageSetup.PrintArea = ""

            # 3) Ajuste de imágenes (Placement) si fuera necesario
            #    (solo si sabes que quieres cambiar su modo de anclaje)
            for shape in sheet.Shapes:
                shape.Placement = 2  # 1=xlMove, 2=xlMoveAndSize, 3=xlFreeFloating

            # 4) Configuración de página (zoom o fit)
            #    *Ojo: FitToPagesWide=1 reduce el ancho a 1 página, que a veces “encoge” el layout
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = False
            sheet.PageSetup.Orientation = 2  # 2 = Landscape
            sheet.PageSetup.LeftMargin = 0.5
            sheet.PageSetup.RightMargin = 0.5
            sheet.PageSetup.TopMargin = 0.5
            sheet.PageSetup.BottomMargin = 0.5

            # 5) Exportar a PDF
            workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            print(f"[OK] PDF generado: {pdf_path}")

        except Exception as e:
            print(f"[ERROR] Al exportar a PDF: {e}")

        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
            excel.Quit()
