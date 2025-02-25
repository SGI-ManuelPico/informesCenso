import os
import win32com.client

class Pdf:
    def excelPdf(
        self,
        excel_path,
        pdf_path,
        sheet_name=None,
        orientation=2,          # 2 = Landscape, 1 = Portrait
        paper_size=9,           # 9 = A4, 2 = Carta, 8 = A3, etc.
        fit_to_pages_wide=1,    # 1 => ajusta todo el ancho en 1 pág;  False/0 => no ajusta
        fit_to_pages_tall=1,    # 1 => ajusta todo el alto en 1 pág;   False/0 => varias págs de alto
        zoom=None,              # None => usa FitToPages, un número => escala
        center_horizontally=True,
        center_vertically=True,
        left_margin=0.5,
        right_margin=0.5,
        top_margin=0.5,
        bottom_margin=0.5,
        clear_print_area=True
    ):
        """
        Convierte un archivo Excel a PDF, con parámetros configurables.

        Parámetros:
        -----------
        excel_path : str
            Ruta al archivo Excel de entrada.
        pdf_path   : str
            Ruta donde se generará el PDF de salida.
        sheet_name : str, opcional
            Nombre de la hoja que se desea exportar.
            Si no se especifica, se usa la hoja activa.

        orientation : int
            1 = Portrait, 2 = Landscape (default 2).
        paper_size  : int
            9 = A4 (default), 2 = Letter, 8 = A3, etc.
            (Ver enumeraciones xlPaperSize en la documentación de Excel.)

        fit_to_pages_wide : int | bool
            Cuántas páginas de ancho se permite. 1 => 1 sola página de ancho,
            False o 0 => no ajusta el ancho.  (default=1)
        fit_to_pages_tall : int | bool
            Cuántas páginas de alto se permite. 1 => 1 sola página de alto,
            False o 0 => no ajusta el alto.   (default=1)

        zoom : int | None
            Si se especifica un número (ej. 100, 80, 120), se usará ese zoom fijo
            en vez de FitToPages. (default=None)

        center_horizontally : bool
            Centra horizontalmente la impresión. (default=True)
        center_vertically   : bool
            Centra verticalmente la impresión. (default=True)

        left_margin, right_margin, top_margin, bottom_margin : float
            Márgenes en pulgadas. (default=0.5 cada uno)

        clear_print_area : bool
            Si True, limpia el PrintArea antes de exportar. (default=True)
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

            # 3) (Opcional) Limpiar el área de impresión
            if clear_print_area:
                sheet.PageSetup.PrintArea = ""

            # 4) Ajuste de imágenes (si hay shapes)
            for shape in sheet.Shapes:
                shape.Placement = 2  # 2 = xlMoveAndSize

            # 5) Configuración de página

            # Márgenes
            sheet.PageSetup.LeftMargin   = left_margin
            sheet.PageSetup.RightMargin  = right_margin
            sheet.PageSetup.TopMargin    = top_margin
            sheet.PageSetup.BottomMargin = bottom_margin

            # Orientación (1 = Portrait, 2 = Landscape)
            sheet.PageSetup.Orientation = orientation

            # Tamaño de papel
            sheet.PageSetup.PaperSize = paper_size

            # Centrado
            sheet.PageSetup.CenterHorizontally = center_horizontally
            sheet.PageSetup.CenterVertically   = center_vertically

            # Si se especifica un zoom fijo, lo usamos.
            # En este caso, desactivamos FitToPages.
            if zoom is not None:
                sheet.PageSetup.Zoom = zoom
                sheet.PageSetup.FitToPagesWide = False
                sheet.PageSetup.FitToPagesTall = False
            else:
                # Sin zoom => usar FitToPages
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = fit_to_pages_wide
                sheet.PageSetup.FitToPagesTall = fit_to_pages_tall

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
