from persistence.informeOctavo import InformeOctavo
import os
from openpyxl import load_workbook

def main():
            
    archivoInicial = InformeOctavo().lecturaArchivoOctavo()
    rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 8 AGROINDUSTRIA - Aprobado.xlsx"
    direc_guardado = os.getcwd() + "\\Formatos Finales"
    if not os.path.exists(direc_guardado):
        os.makedirs(direc_guardado)
    for index, row in archivoInicial.iterrows():
        wb = load_workbook(rutaArchivoFormato)
        ws = wb.active
    
        InformeOctavo().crearArchivoOctavo(ws, row)
    

        output_path = f"{direc_guardado}" + "\\" + f"formularioOctavoLleno_{index + 1}.xlsx"
        wb.save(output_path)

if __name__== "__main__":
    main()