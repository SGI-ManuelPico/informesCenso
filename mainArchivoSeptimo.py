from persistence.informeSeptimo import InformeSeptimo
import os
from openpyxl import load_workbook

def main():
            
    archivoInicial = InformeSeptimo().lecturaArchivoSeptimo()
    rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 8 AGROINDUSTRIA - Aprobado.xlsx"
    direc_guardado = os.getcwd() + "\\Formatos Finales"
    if not os.path.exists(direc_guardado):
        os.makedirs(direc_guardado)
    archivoInicial.to_excel("hola.xlsx")
    for index, row in archivoInicial.iterrows():
        wb = load_workbook(rutaArchivoFormato)
        ws = wb.active
    
        InformeSeptimo().crearArchivoSeptimo(ws, row)
    

        output_path = f"{direc_guardado}" + "\\" + f"form_lleno_{index + 1}.xlsx"
        wb.save(output_path)

if __name__== "__main__":
    main()