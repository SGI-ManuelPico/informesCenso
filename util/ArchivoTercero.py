from persistence.informeTercero import InformeTercero
import os
from openpyxl import load_workbook
from .Pdf import Pdf

class ArchivoTercero:
    def crearArchivoTercero(self):
                
        archivoInicial = InformeTercero().lecturaArchivoTercero()
        rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 3 COMERCIAL - Aprobado.xlsx"
        direc_guardado = os.getcwd() + "\\Formatos Finales"
        if not os.path.exists(direc_guardado):
            os.makedirs(direc_guardado)
        pdf = Pdf()
        for index, row in archivoInicial.iterrows():
            wb = load_workbook(rutaArchivoFormato)
            ws = wb.active
        
            InformeTercero().crearArchivoTercero(ws, row)
        

            output_path = f"{direc_guardado}" + "\\" + f"formularioTerceroLleno_{index + 1}.xlsx"
            wb.save(output_path)

            # Convertir a PDF
            pdf_path = output_path.replace('.xlsx', '.pdf')
            pdf.excelPdf(output_path, pdf_path)