import pandas as pd
import os, re

class InformeNoveno:
    def valorCol(self, base_name, index, df_fila):
        if index == 0:
            return df_fila.get(base_name, '')
        else:
            return df_fila.get(f'{base_name}.{index}', '')

    def lecturaArchivoNoveno(self):
        rutaArchivoInicial = os.getcwd() + "\\Censo Económico Maute.xlsm"
        xl = pd.ExcelFile(rutaArchivoInicial)
        df9 = xl.parse(sheet_name='FORMATO 9. MINERIA', header=None)
        df9 = df9.drop(columns=[0, 1, 2, 3])
        df9_T = df9.transpose()
        df9_T.columns = df9_T.iloc[0]
        df9_T = df9_T.drop(df9_T.index[0])
        df9_T.columns = df9_T.columns.str.strip()
        df9_T.columns = pd.io.common.dedup_names(df9_T.columns, is_potential_multiindex=False)
        df_enc9 = df9_T.reset_index(drop = True)

        return df_enc9

    def crearArchivoNoveno(self, ws, df_fila):

        ws['AN1'] = df_fila['Encuesta No.']

        if pd.notna(df_fila['Fecha(DD/MM/AAAA)']):
            fecha_str = str(df_fila['Fecha(DD/MM/AAAA)'])
            if '/' in fecha_str:
                ws['AL2'] = re.findall('\d+',fecha_str.split("/")[2])[0]
                ws['AN2'] = fecha_str.split('/')[1]
                ws['AS2'] = fecha_str.split('/')[0]
            elif '-' in fecha_str:
                ws['AL2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                ws['AN2'] = fecha_str.split('-')[1]
                ws['AS2'] = fecha_str.split('-')[0]
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')  
        
        ws['AN3'] = df_fila['Encuestador']
        
        # A. Identificación del entrevistado
        ws['F7'] = df_fila['Nombre']
        ws['X7'] = df_fila['Empresa']
        ws['AN7'] = df_fila['Cargo']

        # ¿Pertenece a alguna asociación?
        asociacion = df_fila['¿Pertenece a alguna asociación?']
        if pd.notna(asociacion):
            if asociacion == 'Sí':
                ws['AA8'] = 'X'
                ws['AN8'] = df_fila['¿Cuál?']
            elif asociacion == 'No':
                ws['AC8'] = 'X'
        else:
            print('Campo vacío')

        # B. INFORMACIÓN DE LA ACTIVIDAD
        # Minerales extraídos o explotados y/o transformados
        for i in range(6): 
            row_idx = 15 + i
            ws[f"K{row_idx}"] = InformeNoveno().valorCol('Unidad de medida', i, df_fila)
            ws[f"M{row_idx}"] = InformeNoveno().valorCol('Cantidad día', 2*i, df_fila)
            ws[f"O{row_idx}"] = InformeNoveno().valorCol('Cantidad Mes', 2*i, df_fila)
            ws[f"R{row_idx}"] = InformeNoveno().valorCol('Cantidad Año', 2*i, df_fila)
            ws[f"X{row_idx}"] = InformeNoveno().valorCol('Costo de producción/por unidad de medida', i, df_fila)
            ws[f"AE{row_idx}"] = InformeNoveno().valorCol('Cantidad día', 2*i+1, df_fila)
            ws[f"AG{row_idx}"] = InformeNoveno().valorCol('Cantidad Mes', 2*i+1, df_fila)
            ws[f"AK{row_idx}"] = InformeNoveno().valorCol('Cantidad Año', 2*i+1, df_fila)
            ws[f"AN{row_idx}"] = InformeNoveno().valorCol('Valor total de ventas según frecuencia', i, df_fila)


        tiene_calculo = df_fila['¿Tiene un cálculo aproximado del tiempo que puede seguir explotando el mineral?']
        if pd.notna(tiene_calculo):
            if tiene_calculo == 'Sí':
                ws['O22'] = 'X'
                ws['AE22'] = df_fila['¿Cuánto']
            elif tiene_calculo == 'No':
                ws['Q22'] = 'X'
        else:
            print('Campo vacío')

        
        objeto_explotacion = df_fila['¿Cuál es el objeto de la explotación?']
        if pd.notna(objeto_explotacion):
            if objeto_explotacion == 'Extracción de minerales': 
                ws['T23'] = 'X'
            elif objeto_explotacion == 'Transformación de minerales':
                ws['AE23'] = 'X'
            elif objeto_explotacion == 'Extracción y transformación de minerales':
                ws['AT23'] = 'X'
        else:
            print('Campo vacío')

        recurso_hidrico = df_fila['¿De dónde se abastece del recurso hídrico?']
        if pd.notna(recurso_hidrico):
            if recurso_hidrico == 'Aljibe': 
                ws['V24'] = 'X'
            elif recurso_hidrico == 'Acueducto veredal':
                ws['AD24'] = 'X'
            elif recurso_hidrico == 'Otro':
                ws['AK24'] = 'X'
                ws['AP24'] = df_fila['¿Cuál?.1']
        else:
            print('Campo vacío: ¿Cuál es el objeto de la explotación?')   

        ws['V25'] = df_fila['Forma de extracción']
        ws['AP25'] = df_fila['Cantidad estimada (m3)']

        # Tipo de energía que utiliza
        energia = df_fila['¿Qué tipo de energía utiliza?']
        if pd.notna(energia):
            if energia == 'Energía eléctrica':
                ws['AA26'] = 'X'
            elif energia == 'Energía solar':
                ws['AI26'] = 'X'
            elif energia == 'Otro':
                if pd.notna(df_fila['¿Cuál?.2']):
                    ws['AQ26'] = df_fila['¿Cuál?.2']
        else:
            print('Campo vacio')

        energia_coccion = df_fila['¿De dónde proviene la energía que utiliza para la cocción de alimentos?']
        if pd.notna(energia_coccion):
            if energia_coccion == 'Energía eléctrica':
                ws['AA27'] = 'X'
            elif energia_coccion == 'Leña':
                ws['AE27'] = 'X'
            elif energia_coccion == 'Gas':
                ws['AK27'] = 'X'
            elif energia_coccion == 'Otro':
                if pd.notna(df_fila['¿Cuál?.3']):
                    ws['AQ27'] = df_fila['¿Cuál?.3']
        else:
            print('Campo vacío')

        alcantarillado = df_fila['¿Cuenta con servicio de alcantarillado?']
        if pd.notna(alcantarillado):
            if alcantarillado == 'Sí':
                ws['AA28'] = 'X'
                ws['AA28'] = df_fila['¿Cuál?.4']
            elif alcantarillado == 'No':
                ws['AC28'] = 'X'

        ws['A30'] = df_fila['¿Cómo se maneja el agua residual y sólidos?']

        ws['Z31'] = df_fila['Sitio de venta']
        

        if pd.notna(df_fila['Hidrocarburos']):
            ws['Q33'] = 'X'
            ws['AC33'] = df_fila['¿Cuál? (Empresa y proyecto)']
        
        if pd.notna(df_fila['Otros']):
            ws['Q34'] = 'X'
            ws['AC34'] = df_fila['¿Cuáles?']


        ws['O36'] = df_fila['Valor total de la maquinaria']

        if pd.notna(df_fila['¿Cuenta con permisos ambientales?']):
            if df_fila['¿Cuenta con permisos ambientales?'] == 'Sí':
                ws['Y37'] = 'X'
            elif df_fila['¿Cuenta con permisos ambientales?'] == 'No':
                ws['AA37'] = 'X'
        else:
            print('Campo vacío')   
        
        # Frecuencia de demanda de servicios

        ws['AB37'] = df_fila['Lugar de procedencia de la maquinaria']
        
        ws['L40'] = df_fila['Seguridad']

        ws['L41'] = df_fila['Mano de obra calificada']

        ws['L42'] = df_fila['Mano de obra no calificada']

        ws['L43'] = df_fila['Transporte']

        ws['AJ40'] = df_fila['Alojamiento']

        ws['AJ41'] = df_fila['Alimentación']

        ws['AF42'] = df_fila['Otro, ¿Cuál?']


        continuar = df_fila['¿Piensa continuar con la actividad?']
        if pd.notna(continuar):
            if continuar == 'Si':
                ws['M45'] = 'X'
            elif continuar == 'No':
                ws['O45'] = 'X'
        else:
            print('Campo vacio')

        produccion = df_fila['¿Piensa ampliar la producción?']
        if pd.notna(produccion):
            if produccion == 'Si':
                ws['AA45'] = 'X'
            elif produccion == 'No':
                ws['AC45'] = 'X'
        else:
            print('Campo vacio')


        permanecer = df_fila['¿Piensa permanecer con la misma producción?']
        if pd.notna(permanecer):
            if permanecer == 'Si':
                ws['AQ45'] = 'X'
            elif permanecer == 'No':
                ws['AS45'] = 'X'
        else:
            print('Campo vacio')

        # C. INFORMACIÓN LABORAL

        for i in range(10):
            prefijo_persona = 51 + i
            ws[f'E{prefijo_persona}'] = InformeNoveno().valorCol('Cargo', i, df_fila)
            ws[f'K{prefijo_persona}'] = InformeNoveno().valorCol('Edad (años)', i, df_fila)
            ws[f'L{prefijo_persona}'] = InformeNoveno().valorCol('Duración jornada (horas)', i, df_fila)

            manoObra = InformeNoveno().valorCol('Tipo de mano de obra', i, df_fila)
            if pd.notna(manoObra):
                if manoObra == 'Familiar':
                    ws[f'B{prefijo_persona}'] == 'X'
                elif manoObra == 'Contratado':
                    ws[f'D{prefijo_persona}20'] == 'X'

            # Genero
            genero = InformeNoveno().valorCol('Género', i,df_fila)
            if pd.notna(genero):
                if genero == 'Masculino':
                    ws[f'J{prefijo_persona}20'] == 'X'
                elif genero ==  'Femenino':
                    ws[f'H{prefijo_persona}'] == 'X'

            # Escolaridad 
            escolaridad = InformeNoveno().valorCol('Escolaridad', i, df_fila)
            if pd.notna(escolaridad):
                if escolaridad:
                    if escolaridad == 'Primaria':
                        ws[f'N{prefijo_persona}'] = 'X'
                    elif escolaridad == 'Bachillerato':
                        ws[f'P{prefijo_persona}'] = 'X'
                    elif escolaridad == 'Técnico':
                        ws[f'R{prefijo_persona}'] = 'X'
                    elif escolaridad == 'Pregrado':
                        ws[f'T{prefijo_persona}'] = 'X'
                    elif escolaridad == 'Posgrado':
                        ws[f'V{prefijo_persona}'] = 'X'
            else:
                print(f'Campo vacio')

            # Contrato 
            contrato = InformeNoveno().valorCol('Contrato', i,df_fila)
            if contrato:
                if contrato == 'Tem.':
                    ws[f'AA{prefijo_persona}'] = InformeNoveno().valorCol('Contrato', i, df_fila)
                elif contrato == 'Fij':
                    ws[f'AC{prefijo_persona}'] = 'X'
            else:
                print(f'Campo vacio')

            # Pago de seguridad social 
            pago_seguridad = InformeNoveno().valorCol('Pago de seguridad social', i, df_fila)
            if pago_seguridad:
                if pago_seguridad == 'Si':
                    ws[f'AE{prefijo_persona}'] = 'X'
                elif pago_seguridad == 'No':
                    ws[f'AG{prefijo_persona}'] = 'X'
            else:
                print(f'Campo vacio')

            # Remuneración 
            remuneracion = InformeNoveno().valorCol('Remuneración', i, df_fila)
            if remuneracion:
                if remuneracion == 'Inferiores a $900.000':
                    ws[f'AS{prefijo_persona}'] = 'X'
                elif remuneracion == '$900.000 a $1.800.000':
                    ws[f'AT{prefijo_persona}'] = 'X'
                elif remuneracion == '$1.801.000 a $2.700.000':
                    ws[f'AU{prefijo_persona}'] = 'X'
                elif remuneracion == 'Superiores a $2.701.000':
                    ws[f'AV{prefijo_persona}'] = 'X'
            else:
                print(f'Campo vacio')

            # Información adicional
        
            ws[f'AH{prefijo_persona}'] = InformeNoveno().valorCol('Procedencia', i, df_fila)
            ws[f'AI{prefijo_persona}'] = InformeNoveno().valorCol('Residencia', i, df_fila)
            ws[f'AL{prefijo_persona}'] = InformeNoveno().valorCol('Tiempo trabajado', i, df_fila)
            ws[f'AM{prefijo_persona}'] = InformeNoveno().valorCol('# Personas núcleo familiar', i, df_fila)
            ws[f'AO{prefijo_persona}'] = InformeNoveno().valorCol('Personas a cargo', i, df_fila)
            ws[f'AP{prefijo_persona}'] = InformeNoveno().valorCol('Lugar de residencia familiar', i,df_fila)