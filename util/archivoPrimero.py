from persistence.informePrimero import InformePrimero
import os
from openpyxl import load_workbook
import os
from .Pdf import Pdf
class ArchivoPrimero:
    def crearArchivoPrimero(self):
                
        archivoInicial = InformePrimero().lecturaArchivoPrimero()
        rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx"
        direc_guardado = os.getcwd() + "\\Formatos Finales"
        if not os.path.exists(direc_guardado):
            os.makedirs(direc_guardado)
        pdf = Pdf()
        for index, row in archivoInicial.iterrows():
            wb = load_workbook(rutaArchivoFormato)
            ws = wb.active
        
            InformePrimero().crearArchivoPrimero(ws, row)
        

            output_path = f"{direc_guardado}" + "\\" + f"formularioPrimeroLleno_{index + 1}.xlsx"
            wb.save(output_path)

            # Convertir a PDF
            pdf_path = output_path.replace('.xlsx', '.pdf')
            pdf.excelPdf(output_path, pdf_path)
