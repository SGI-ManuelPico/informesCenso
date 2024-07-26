import pandas as pd
import os

class InformeTercero:
    def __init__(self):
        pass

    def lecturaArchivoTercero(self):
        """
        Lectura del tercer archivo del censo económico.
        """

        # Lectura de rutas
        rutaArchivoInicial = os.getcwd() + "\\censos\\Censo Económico Maute.xlsm"
        archivoInicial = pd.ExcelFile(rutaArchivoInicial)
        archivoInicial = archivoInicial.parse(sheet_name="FORMATO 3. COMERCIAL", header=None)

        # Ajustes preliminares al archivo inicial.
        archivoInicial = archivoInicial.drop(columns=[0,1,2,3]).transpose()
        archivoInicial.columns = archivoInicial.iloc[0].str.lstrip()
        archivoInicial = archivoInicial.drop(archivoInicial.index[0])
        archivoInicial.columns = pd.io.common.dedup_names(archivoInicial.columns, is_potential_multiindex=False)

        archivoInicial = archivoInicial.reset_index(drop = True)

        return archivoInicial
    
    def crearArchivoTercero(self, hoja, fila):
        """
        Creación del tercer archivo del censo económico.
        """
        
        hoja['AI1'] = fila["Encuesta No."]
        hoja['AI2'] = fila["Fecha"]###############################
        hoja['AE3'] = fila["Encuestador"]
        hoja['F6'] = fila["Nombre"]
        hoja['AC6'] = fila["Empresa"]
        hoja['AP6'] = fila["Cargo"]
        hoja[''] = fila[""] ####### ASOCIACIÓN SI NO 
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]
        hoja[''] = fila[""]



rutaArchivoFormato = os.getcwd() + "\\censos\\FORMATO 3 COMERCIAL - Aprobado.xlsx"

a = InformeTercero().lecturaArchivoTercero()
print(a)