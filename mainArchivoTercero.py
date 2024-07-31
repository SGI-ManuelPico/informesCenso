from persistence.informeTercero import InformeTercero
import os
from openpyxl import load_workbook

def main():
            
    archivoInicial = InformeTercero().lecturaArchivoTercero()
    rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 8 AGROINDUSTRIA - Aprobado.xlsx"
    direc_guardado = os.getcwd() + "\\Formatos Finales"
    if not os.path.exists(direc_guardado):
        os.makedirs(direc_guardado)
    for index, row in archivoInicial.iterrows():
        wb = load_workbook(rutaArchivoFormato)
        ws = wb.active
    
        InformeTercero().crearArchivoTercero(ws, row)
    

        output_path = f"{direc_guardado}" + "\\" + f"formularioTerceroLleno_{index + 1}.xlsx"
        wb.save(output_path)

if __name__== "__main__":
    main()