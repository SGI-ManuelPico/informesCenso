from persistence.informeSeptimo import InformeSeptimo
from openpyxl import load_workbook
import os

class ArchivoSeptimo:
    def crearArchivoSeptimo(self):
                
        archivoInicial = InformeSeptimo().lecturaArchivoSeptimo()
        rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 8 AGROINDUSTRIA - Aprobado.xlsx"
        direc_guardado = os.getcwd() + "\\Formatos Finales"
        if not os.path.exists(direc_guardado):
            os.makedirs(direc_guardado)
        for index, row in archivoInicial.iterrows():
            wb = load_workbook(rutaArchivoFormato)
            ws = wb.active
        
            InformeSeptimo().crearArchivoSeptimo(ws, row)
        

            output_path = f"{direc_guardado}" + "\\" + f"formularioSéptimoLleno_{index + 1}.xlsx"
            wb.save(output_path)
