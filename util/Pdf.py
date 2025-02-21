import os
import win32com.client

class Pdf:
    def excelPdf(self, excel_path, pdf_path, sheet_name=None):
        """
        Convierte un archivo Excel a PDF de forma silenciosa.
        - Fuerza todo en UNA SOLA PÁGINA (FitToPagesWide=1, FitToPagesTall=1).
        - Orientación horizontal (Landscape).
        - Centra el contenido en la página.
        - Papel A4.
        
        Si la hoja es más pequeña que la página, Excel NO hará "scale up", 
        simplemente quedará centrada y con espacio en blanco. 
        Si la hoja es grande, Excel la reducirá para que quepa en 1 sola página.
        
        Parámetros:
        -----------
        excel_path : str
            Ruta al archivo Excel de entrada.
        pdf_path   : str
            Ruta donde se generará el PDF de salida.
        sheet_name : str, opcional
            Nombre de la hoja que se desea exportar. 
            Si no se especifica, se usa la hoja activa.
        """
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
        except:
            pass

        excel.DisplayAlerts = False
        workbook = None
        try:
            # 1) Abrir el archivo Excel
            workbook = excel.Workbooks.Open(os.path.abspath(excel_path))

            # 2) Seleccionar la hoja deseada (si se especifica)
            if sheet_name:
                try:
                    sheet = workbook.Sheets(sheet_name)
                    sheet.Select()  # Activa la hoja
                except Exception as e:
                    print(f"[WARNING] No se pudo seleccionar la hoja '{sheet_name}': {e}")
                    sheet = workbook.ActiveSheet
            else:
                sheet = workbook.ActiveSheet

            # 3) Limpiar el área de impresión, si la hubiera
            sheet.PageSetup.PrintArea = ""

            # 4) Ajuste de imágenes (opcional: mover y redimensionar con celdas)
            for shape in sheet.Shapes:
                shape.Placement = 2  # 2 = xlMoveAndSize

            # 5) Configuración de página
            #    - Zoom = False => desactiva el zoom manual
            #    - FitToPagesWide=1 => todo el ancho en 1 página
            #    - FitToPagesTall=1 => todo el alto en 1 página (no se cortará)
            sheet.PageSetup.Zoom = False
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1

            # Centrar horizontal y verticalmente
            sheet.PageSetup.CenterHorizontally = True
            sheet.PageSetup.CenterVertically   = True

            # Orientación apaisada
            sheet.PageSetup.Orientation = 2  # 2 = Landscape

            # Tamaño de papel A4 (PaperSize=9). Para Carta usar 2, etc.
            sheet.PageSetup.PaperSize = 9

            # Márgenes (en pulgadas)
            sheet.PageSetup.LeftMargin   = 0.5
            sheet.PageSetup.RightMargin  = 0.5
            sheet.PageSetup.TopMargin    = 0.5
            sheet.PageSetup.BottomMargin = 0.5

            # 6) Exportar a PDF
            workbook.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            print(f"[OK] PDF generado: {pdf_path}")

        except Exception as e:
            print(f"[ERROR] Al exportar a PDF: {e}")

        finally:
            if workbook is not None:
                # Cerrar el libro sin guardar cambios
                workbook.Close(SaveChanges=False)
            excel.Quit()
