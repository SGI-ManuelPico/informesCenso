import pandas as pd
import os, re


class InformePrimero:
    def lecturaArchivoPrimero(self):
        rutaArchivoInicial = os.getcwd() + "\\Censo Económico Maute.xlsm"
        xl = pd.ExcelFile(rutaArchivoInicial)
        df = xl.parse(sheet_name='FORMATO 1. IDENTIFICACIÓN', header=None)
        df = df.drop(columns=[0, 1])
        df_T = df.transpose()
        df_T.columns = df_T.iloc[0]
        df_T = df_T.drop(df_T.index[0])
        df_T.columns = df_T.columns.str.strip()
        df_T.columns = pd.io.common.dedup_names(df_T.columns, is_potential_multiindex=False)
        df_enc = df_T.reset_index(drop = True)

        return df_enc


    def crearArchivoPrimero(self, ws, df_fila):
        ws['Y1'] = df_fila['Encuesta No.']
        if pd.notna(df_fila['Fecha(DD/MM/AAAA)']):
            fecha_str = str(df_fila['Fecha(DD/MM/AAAA)'])
            if '/' in fecha_str:
                ws['X2'] = re.findall('\d+',fecha_str.split("/")[2])[0]
                ws['Z2'] = fecha_str.split('/')[1]
                ws['AD2'] = fecha_str.split('/')[0]
            elif '-' in fecha_str:
                ws['X2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                ws['Z2'] = fecha_str.split('-')[1]
                ws['AD2'] = fecha_str.split('-')[0]
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')     
        ws['W3'] = df_fila['Encuestador']
        ws['A8'] = df_fila['Departamento']
        ws['N8'] = df_fila['Municipio']
        ws['U8'] = df_fila['Vereda/Centro Poblado']
        ws['A10'] = str(df_fila['Coordenada Norte']) + ',' + str(df_fila['Coordenada Este'])

        if df_fila['Permite Entrevista'] == 'Si':
            ws['W10'] = 'X'

            ws['G14'] = df_fila['Nombre del establecimiento']
            ws['D15'] = df_fila['Dirección']
            ws['U15'] = df_fila['Teléfono de contacto']
            ws['G16'] = df_fila['Actividad económica principal']
            ws['W16'] = df_fila['¿En qué año inició la actividad?']
            ws['D17'] = df_fila['Propietario']
            ws['Q17'] = df_fila['Procedencia']
            ws['Y17'] = df_fila['Lugar de Residencia']
            ws['D18'] = df_fila['Administrador']
            ws['Q18'] = df_fila['Procedencia.1']
            ws['Z18'] = df_fila['Lugar de Residencia.1']

            actividad = df_fila['Este establecimiento desarrolla su actividad como']
            if actividad == 'Persona Natural':
                ws['G23'] = 'X'
            elif actividad == 'Sociedad de hecho':
                ws['N23'] = 'X'
            elif actividad == 'Empresa unipersonal':
                ws['G24'] = 'X'
            elif actividad == 'Sociedad Comercial':
                ws['N24'] = 'X'
            elif actividad == 'Cooperativa':
                ws['G25'] = 'X'
            elif actividad == 'Predio':
                ws['G26'] = 'X'
            elif actividad == 'Otro tipo de sociedad':
                ws['N25'] = 'X'
                ws['K26'] = df_fila['¿Cuál?']


            actividad2 = df_fila['Tipo de actividad']
            
            if pd.notna(actividad2):
                if actividad2 == 'Agrícola':
                    ws['G30'] = 'X'
                elif actividad2 == 'Pecuaria':
                    ws['G31'] = 'X'
                elif actividad2 == 'Agroindustrial':
                    ws['G32'] = 'X'
                elif actividad2 == 'Servicios':
                    ws['G33'] = 'X'
                elif actividad2 == 'Servicios prestados al sector de hidrocarburos':
                    ws['G34'] = 'X'
                elif actividad2 == 'Comercial':
                    ws['M31'] = 'X'
                elif actividad2 == 'Manufactura':
                    ws['M32'] = 'X'
                elif actividad2 == 'Transporte':
                    ws['M33'] = 'X'
                elif actividad2 == 'Minería':
                    ws['M34'] = 'X'
                else:
                    print('No hay coincidencias en actividad2')
            else:
                print('Tipo de actividad es NaN')


            actividad3 = df_fila['Tenencia de la propiedad']
            print(f"Tenencia de la propiedad: {actividad3}")
            if pd.notna(actividad3):
                if actividad3 == 'Propia':
                    ws['E40'] = 'X'
                elif actividad3 == 'Administrada':
                    ws['E41'] = 'X'
                elif actividad3 == 'Arrendada*  Responder 16':
                    ws['E42'] = 'X'
                    ws['J41'] = df_fila['Canon de arrendamiento']
                elif actividad3 == 'Otra':
                    ws['E44'] = 'X'
                    ws['C45'] = df_fila['¿Cuál?.1']
                else:
                    print(f'No hay coincidencias en actividad3: {actividad3}')
            else:
                print('Tenencia de la propiedad es NaN')
            

            ws['A48'] = df_fila['¿De que actividad proviene la mayor parte de ingresos obtenidos en la unidad económica?']
        

            actividad4 = df_fila['Frecuencia con la que recibe ingresos por actividad']
            print(f"Frecuencia con la que recibe ingresos por actividad: {actividad4}")
            if pd.notna(actividad4):
                if actividad4 == 'Diario':
                    ws['E53'] = 'X'
                elif actividad4 == 'Semanal':
                    ws['E54'] = 'X'
                elif actividad4 == 'Quincenal':
                    ws['E55'] = 'X'
                elif actividad4 == 'Mensual':
                    ws['M53'] = 'X'
                elif actividad4 == 'Semestral':
                    ws['M54'] = 'X'
                elif actividad4 == 'Otra':
                    ws['M55'] = 'X'
                    ws['K56'] = df_fila['¿Cuál?.2']
                else:
                    print(f'No hay coincidencias en actividad4: {actividad4}')
            else:
                print('Frecuencia con la que recibe ingresos por actividad es NaN')
            
            actividad5 = df_fila['¿Cuál es la cantidad de ingresos recibidos por la actividad?']
            if pd.notna(actividad5):
                if actividad5 == 'Inferiores a $600.000':
                    ws['Y23'] = 'X'
                elif actividad5 == 'Entre $601.000 y $1.500.000':
                    ws['Y24'] = 'X'
                elif actividad5 == 'Entre $1.501.000 y $3.000.000':
                    ws['Y25'] = 'X'
                elif actividad5 == 'Superior a $ 3.000.000':
                    ws['Y26'] = 'X'
                elif actividad5 == 'Otro':
                    ws['AD24'] = 'X'
                    ws['AA25'] = "$  " +  df_fila['¿Cuál?.3']
                else:
                    print(f'No hay coincidencias en actividad5: {actividad5}')
            else:
                print('¿Cuál es la cantidad de ingresos recibidos por la actividad? es NaN')

            ws['P30'] = df_fila['¿En qué horario desempeña la actividad?']

        
            actividad6 = df_fila['¿Tiene registro de cámara y comercio?']
            print(f"¿Tiene registro de cámara y comercio?: {actividad6}")  
            if pd.notna(actividad6):
                if actividad6 == 'Si':
                    ws['T33'] = 'X'
                elif actividad6 == 'No':
                    ws['W33'] = 'X'
                else:
                    print(f'No hay coincidencias en actividad6: {actividad6}')
            else:
                print('¿Tiene registro de cámara y comercio? es NaN')

            
            actividad7 = df_fila['¿En qué lugares comercializa?']
            print(f"¿En qué lugares comercializa?: {actividad7}") 
            if pd.notna(actividad7):
                if actividad7 == 'Directamente en el sitio':
                    ws['X39'] = 'X'
                elif actividad7 == 'Empresa':
                    ws['X40'] = 'X'
                elif actividad7 == 'Mercado':
                    ws['X41'] = 'X'
                elif actividad7 == 'Centro de acopio':
                    ws['X42'] = 'X'
                elif actividad7 == 'Otro':
                    ws['X44'] = 'X'
                    ws['S45'] = df_fila['¿Dónde?']
                else:
                    print(f'No hay coincidencias en actividad7: {actividad7}')
            else:
                print('¿En qué lugares comercializa? es NaN')

            
            actividad8 = df_fila['¿Compra productos o insumos en la vereda?']
            print(f"¿Compra productos o insumos en la vereda?: {actividad8}")  
            if pd.notna(actividad8):
                if actividad8 == 'Si':
                    ws['T48'] = 'X'
                elif actividad8 == 'No':
                    ws['W48'] = 'X'
                else:
                    print(f'No hay coincidencias en actividad8: {actividad8}')
            else:
                print('¿Compra productos o insumos en la vereda? es NaN')

            actividad9 = df_fila['¿Comercializa productos o insumos en otras veredas?']
            print(f"¿Comercializa productos o insumos en otras veredas?: {actividad9}")  
            if pd.notna(actividad9):
                if actividad9 == 'Si':
                    ws['U53'] = 'X'
                    ws['U55'] = df_fila['¿En qué lugares?']
                elif actividad9 == 'No':
                    ws['Y53'] = 'X'
                else:
                    print(f'No hay coincidencias en actividad9: {actividad9}')
            else:
                print('¿Comercializa productos o insumos en otras veredas? es NaN')

            # For 'Estrato'
            actividad10 = df_fila['Estrato']
            print(f"Estrato: {actividad10}")  
            if pd.notna(actividad10):
                if actividad10 == 1:
                    ws['R59'] = 'X'
                elif actividad10 == 2:
                    ws['T59'] = 'X'
                elif actividad10 == 3:
                    ws['V59'] = 'X'
                else:
                    print(f'No hay coincidencias en actividad10: {actividad10}')
            else:
                print('Estrato es NaN')
            
            ws['Y59'] = df_fila['¿Cuánto pagó el último mes por concepto de servicios públicos?']

        elif df_fila['Permite Entrevista'] == 'No':
            ws['Z10'] = 'X'