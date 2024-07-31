from persistence.informeNoveno import InformeNoveno
import os
from openpyxl import load_workbook
import os

class ArchivoNoveno:
    def crearArchivoNoveno(self):
                
        archivoInicial = InformeNoveno().lecturaArchivoNoveno()
        rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 9 MINERÍA - Aprobado.xlsx"
        direc_guardado = os.getcwd() + "\\Formatos Finales"
        if not os.path.exists(direc_guardado):
            os.makedirs(direc_guardado)
        for index, row in archivoInicial.iterrows():
            wb = load_workbook(rutaArchivoFormato)
            ws = wb.active
        
            InformeNoveno().crearArchivoNoveno(ws, row)
        

            output_path = f"{direc_guardado}" + "\\" + f"formularioNovenoLleno_{index + 1}.xlsx"
            wb.save(output_path)
