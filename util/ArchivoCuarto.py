from persistence.informeCuarto import InformeCuarto
import os
from openpyxl import load_workbook
import os

class ArchivoCuarto:
    def crearArchivoCuarto(self):
                
        archivoInicial = InformeCuarto().lecturaArchivoCuarto()
        rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 8 AGROINDUSTRIA - Aprobado.xlsx"
        direc_guardado = os.getcwd() + "\\Formatos Finales"
        if not os.path.exists(direc_guardado):
            os.makedirs(direc_guardado)
        for index, row in archivoInicial.iterrows():
            wb = load_workbook(rutaArchivoFormato)
            ws = wb.active
        
            InformeCuarto().crearArchivoCuarto(ws, row)
        

            output_path = f"{direc_guardado}" + "\\" + f"formularioCuartoLleno_{index + 1}.xlsx"
            wb.save(output_path)