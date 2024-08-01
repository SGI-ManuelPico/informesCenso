import pandas as pd
import os, re

class InformeQuinto:
    def lecturaArchivoQuinto(self):
        rutaArchivoInicial = os.getcwd() + "\\Censo Económico Maute.xlsm"
        xl = pd.ExcelFile(rutaArchivoInicial)
        df5 = xl.parse(sheet_name='FORMATO 5. SERVICIOS PRESTADOS', header=None)
        df5 = df5.drop(columns=[0, 1, 2, 3])
        df5_T = df5.transpose()
        df5_T.columns = df5_T.iloc[0]
        df5_T = df5_T.drop(df5_T.index[0])
        df5_T.columns = df5_T.columns.str.strip()
        df5_T.columns = pd.io.common.dedup_names(df5_T.columns, is_potential_multiindex=False)
        df_enc5 = df5_T.reset_index(drop = True)

        return df_enc5


    def crearArchivoQuinto(self, ws, df_fila):
        # A. IDENTIFICACIÓN ENTREVISTADO
        ws['AO1'] = df_fila['Encuesta No.']

        if pd.notna(df_fila['Fecha(DD/MM/AAAA)']):
            fecha_str = str(df_fila['Fecha(DD/MM/AAAA)'])
            if '/' in fecha_str:
                ws['AN2'] = re.findall('\d+',fecha_str.split("/")[2])[0]
                ws['AQ2'] = fecha_str.split('/')[1]
                ws['AU2'] = fecha_str.split('/')[0]
            elif '-' in fecha_str:
                ws['AN2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                ws['AQ2'] = fecha_str.split('-')[1]
                ws['AU2'] = fecha_str.split('-')[0]
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')  
        
        ws['AO3'] = df_fila['Encuestador']


        # A. IDENTIFICACIÓN ENTREVISTADO

        ws['F7'] = df_fila['Nombre']
        ws['Y7'] = df_fila['Empresa']
        ws['AO7'] = df_fila['Cargo']

        asociacion = df_fila['¿Pertenece a alguna asociación?']
        if pd.notna(asociacion):
            if asociacion == 'Si':
                ws['AB8'] = 'X'
                ws['AO8'] = df_fila['Otro, ¿Cuál?']
            elif asociacion == 'No':
                ws['AD8'] = 'X'
        else:
            print(f'Campo vacío')


        # B. DESCRIPCIÓN DE LA ACTIVIDAD

        ws['A12'] = df_fila['Con cuántos empleados cuenta su empresa actualmente']

        partip_capital = df_fila['Participación de capital\nSi (Responder 3)']
        if pd.notna(partip_capital):
            if partip_capital == 'Si':
                ws['P15'] = 'X'
                ws['K16'] = df_fila['País del cual proviene el capital']

            elif partip_capital == 'No':
                ws['R15'] = 'X'
            
        else:
            print('Campo vacío')

        # 4. Tipo de servicio

        tipo_servicio = df_fila['Tipo de servicio']
        if pd.notna(tipo_servicio):
            if tipo_servicio == 'Servicios relacionados con perforación de pozos':
                ws['R19'] = 'X'
            elif tipo_servicio == 'Venta y alquiler de equipos y herramientas':
                ws['R20'] = 'X'
            elif tipo_servicio == 'Transporte de maquinaria o combustibles':
                ws['R21'] = 'X'
            elif tipo_servicio == 'Otros servicios':
                ws['R22'] = 'X'
                ws['G23'] = df_fila['Otros, ¿Cuáles?']
        else:
            print('Campo vacio')

        # 5. ¿Cuál es el principal servicio demandado por el sector de hidrocarburos?

        ws['A25'] = df_fila['¿Cuál es el principal servicio demandado por el sector de hidrocarburos?']

        # 6. Frecuencia con la que son demandados los servicios

        actividad_frecuencia = df_fila['Frecuencia con la que son demandados los servicios por parte del sector de hidrocarburos']
        if pd.notna(actividad_frecuencia):
            if actividad_frecuencia == 'Mensual':
                ws['F29'] = 'X'
            elif actividad_frecuencia == 'Bimestral':
                ws['F30'] = 'X'
            elif actividad_frecuencia == 'Trimestral':
                ws['F31'] = 'X'
            elif actividad_frecuencia == 'Semestral':
                ws['K28'] = 'X'
            elif actividad_frecuencia == 'Semanal':
                ws['M29'] = 'X'
            elif actividad_frecuencia == 'Anual':
                ws['M30'] = 'X'            
            elif actividad_frecuencia == 'Contrato permanente':
                ws['V30'] = 'X'
            elif actividad_frecuencia == 'Otro':
                ws['M31'] = 'X'
                ws['P32'] = df_fila['Otro, ¿Cuál?.1']
        else: 
            print('Campo vacío')

        # 7. En promedio cuántos servicios son contratados al mes por el sector de hidrocarburos
        ws['C36'] = df_fila['Servicio 1']
        ws['N36'] = df_fila['Cantidad']

        ws['C37'] = df_fila['Servicio 2']
        ws['N37'] = df_fila['Cantidad.1']

        ws['C38'] = df_fila['Servicio 3']
        ws['N38'] = df_fila['Cantidad.2']

        ws['C39'] = df_fila['Servicio 4']
        ws['N39'] = df_fila['Cantidad.3']


        # 8. Ingreso promedio mensual

        ws['N40'] = df_fila['Ingreso promedio mensual ($)']


        # 9. ¿El pago por sus servicios es oportuno?
        pago_oportuno = df_fila['¿El pago por sus servicios es oportuno?']
        if pd.notna(pago_oportuno):
            if pago_oportuno == 'Si':
                ws['P42'] = 'X'
            elif pago_oportuno == 'No':
                ws['R42'] = 'X'
        else:
            print('Campo vacío')

        # 10. ¿Cuál es el precio de los servicios ofertados para el sector de hidrocarburos?
        
        ws['AA13'] = df_fila['Servicio 1.1']
        ws['AJ13'] = df_fila['Cantidad (según frecuencia)']
        ws['AR13'] = df_fila['Precio']

        ws['AA14'] = df_fila['Servicio 2.1']
        ws['AJ14'] = df_fila['Cantidad (según frecuencia).1']
        ws['AR14'] = df_fila['Precio.1']


        ws['AA15'] = df_fila['Servicio 3.1']
        ws['AJ15'] = df_fila['Cantidad (según frecuencia).2']
        ws['AR15'] = df_fila['Precio.2']


        ws['AA16'] = df_fila['Servicio 4.1']
        ws['AJ16'] = df_fila['Cantidad (según frecuencia).3']
        ws['AR16'] = df_fila['Precio.3']



        # 11. ¿Cuál es el costo aproximado de los servicios ofertados según la frecuencia de venta?
        ws['AB20'] = df_fila['Servicio 1.2']
        ws['AQ20'] = df_fila['Costo']

        ws['AB21'] = df_fila['Servicio 2.2']
        ws['AQ21'] = df_fila['Costo.1']

        ws['AB22'] = df_fila['Servicio 3.2']
        ws['AQ22'] = df_fila['Costo.2']
        
        ws['AB23'] = df_fila['Servicio 4.2']
        ws['AQ23'] = df_fila['Costo.3']

        # 12. ¿Cuál es el costo o valor aproximado de los insumos o equipos utilizados para prestar los servicios en el último mes?
        ws['AD26'] = df_fila['¿Cuál es el costo o valor aproximado de los insumos o equipos utilizados para prestar los servicios en el último mes?']

        # 13. ¿De dónde provienen los insumos y/o maquinaria y/o equipos que emplea?
        ws['Z29'] = df_fila['¿De dónde provienen los insumos y/o maquinaria y/o equipos que emplea?']

        # 14. Vende principalmente en:
        lugar_venta = df_fila['Vende principalmente en:']
        if pd.notna(lugar_venta):
            if lugar_venta == 'Sitio':
                ws['AH32'] = 'X'
            elif lugar_venta == 'Vereda':
                ws['AH33'] = 'X'
            elif lugar_venta == 'Casco Urbano':
                ws['AH35'] = 'X'
            elif lugar_venta == 'Otros municipios y/o veredas':
                ws['AS33'] = 'X'
                ws['AO35'] = df_fila['¿Cuáles?.1']
        else:
            print('Campo vacío')

        # 15. ¿Posee algún permiso ambiental?
        permiso_ambiental = df_fila['¿Posee algún permiso ambiental?']
        if pd.notna(permiso_ambiental):
            if permiso_ambiental == 'No':
                ws['AM37'] = 'X'
            elif permiso_ambiental == 'Si':
                ws['AK37'] = 'X'
                ws['AR37'] = df_fila['¿Cuál?']
        else:
            print('Campo vacío')

        # 16. Mantenimiento de la actividad
        continuar = df_fila['¿Piensa continuar con la actividad?']
        if pd.notna(permiso_ambiental):
            if continuar == 'Si':
                ws['AQ39'] = 'X'
            elif continuar == 'No':
                ws['AS39'] = 'X'
        else:
            print('Campo vacio')

        produccion = df_fila['¿Piensa ampliar la producción?']
        if pd.notna(permiso_ambiental):
            if produccion == 'Si':
                ws['AQ40'] = 'X'
            elif produccion == 'No':
                ws['AS40'] = 'X'
        else:
            print('Campo vacio')


        permanecer = df_fila['¿Piensa permanecer con la misma producción?']
        if pd.notna(permiso_ambiental):
            if permanecer == 'Si':
                ws['AQ41'] = 'X'
            elif permanecer == 'No':
                ws['AS41'] = 'X'
        else:
            print('Campo vacio')

        finalizar = df_fila['¿Piensa finalizar el servicio?']
        if pd.notna(permiso_ambiental):
            if finalizar == 'Si':
                ws['AQ42'] = 'X'
            elif finalizar == 'No':
                ws['AS42'] = 'X'
        else:
            print('Campo vacio')
        
        
        # C. INFORMACIÓN LABORAL

        # Mano de obra contratada
        
        ws['I48'] = df_fila['#']

        genero = df_fila['Género']
        if pd.notna(genero):
            if genero == 'Masculino':
                ws['M48'] = 'X'
            elif genero == 'Femenino':
                ws['K48'] = 'X'
        else:
            print('Campo vacio')

        contrato = df_fila['Contrato']
        if pd.notna(genero):
            if contrato == 'Termino Fijo':
                ws['O48'] = df_fila['¿Cuánto?']
            elif contrato == 'Indefinido':
                ws['S48'] = 'X'
        else:
            print('Campo vacio')

        ws['W48'] = df_fila['Jornal y turno laboral']

        escolaridad = df_fila['Escolaridad']
        if pd.notna(df_fila['Escolaridad']):
            if df_fila['Escolaridad'] == 'Primaria':
                ws['AB48'] = 'X'
            elif df_fila['Escolaridad'] == 'Bachillerato':
                ws['AC48'] = 'X'
            elif df_fila['Escolaridad'] == 'Técnico o tecnológico ':
                ws['AD48'] = 'X'
            elif df_fila['Escolaridad'] == 'Profesional':
                ws['AE48'] = 'X'
            elif df_fila['Escolaridad'] == 'Posgrado':
                ws['AG48'] = 'X'
        else:
            print('Campo vacio')

        procedencia = df_fila['Procedencia']
        if pd.notna(procedencia):
            if procedencia == 'Vereda':
                ws['AJ48'] = 'X'
            elif procedencia == 'Municipio':
                ws['AM48'] = 'X'
            elif procedencia == 'Otro':
                ws['AO48'] = 'X'
        else:
            print('Campo vacio')

        residencia = df_fila['Residencia']
        if pd.notna(escolaridad):
            if residencia == 'Vereda':
                ws['AQ48'] = 'X'
            elif residencia == 'Municipio':
                ws['AS48'] = 'X'
            elif residencia == 'Otro':
                ws['AU48'] = 'X'
        else:
            print('Campo vacio')
    

        # Mano de obra no calificada

        ws['I49'] = df_fila['#.1']

        genero = df_fila['Género.1']
        if pd.notna(genero):
            if genero == 'Masculino':
                ws['M49'] = 'X'
            elif genero == 'Femenino':
                ws['K49'] = 'X'
        else:
            print('Campo vacio')

        contrato = df_fila['Contrato.1']
        if pd.notna(genero):
            if contrato == 'Termino Fijo':
                ws['O49'] = df_fila['¿Cuánto?.1']
            elif contrato == 'Indefinido':
                ws['S49'] = 'X'
        else:
            print('Campo vacio')

        ws['W49'] = df_fila['Jornal y turno laboral.1']

        escolaridad = df_fila['Escolaridad.1']
        if pd.notna(escolaridad):
            if escolaridad == 'Primaria':
                ws['AB49'] = 'X'
            elif escolaridad == 'Bachillerato':
                ws['AC49'] = 'X'
            elif escolaridad == 'Técnico o tecnológico ':
                ws['AD49'] = 'X'
            elif escolaridad == 'Profesional':
                ws['AE49'] = 'X'
            elif escolaridad == 'Posgrado':
                ws['AG49'] = 'X'
        else:
            print('Campo vacio')

        procedencia = df_fila['Procedencia.1']
        if pd.notna(escolaridad):
            if procedencia == 'Vereda':
                ws['AJ49'] = 'X'
            elif procedencia == 'Municipio':
                ws['AM49'] = 'X'
            elif procedencia == 'Otro':
                ws['AO49'] = 'X'
        else:
            print('Campo vacio')

        residencia = df_fila['Residencia.1']
        if pd.notna(escolaridad):
            if residencia == 'Vereda':
                ws['AQ49'] = 'X'
            elif residencia == 'Municipio':
                ws['AS49'] = 'X'
            elif residencia == 'Otro':
                ws['AU49'] = 'X'
        else:
            print('Campo vacio')



        # Empleados administrativos y contables


        ws['I50'] = df_fila['#.2']

        genero = df_fila['Género.2']
        if pd.notna(genero):
            if genero == 'Masculino':
                ws['M50'] = 'X'
            elif genero == 'Femenino':
                ws['K50'] = 'X'
        else:
            print('Campo vacio')

        contrato = df_fila['Contrato.2']
        if pd.notna(genero):
            if contrato == 'Termino Fijo':
                ws['O50'] = df_fila['¿Cuánto?.2']
            elif contrato == 'Indefinido':
                ws['S50'] = 'X'
        else:
            print('Campo vacio')

        ws['W50'] = df_fila['Jornal y turno laboral.2']

        escolaridad = df_fila['Escolaridad.2']
        if pd.notna(escolaridad):
            if escolaridad == 'Primaria':
                ws['AB50'] = 'X'
            elif escolaridad == 'Bachillerato':
                ws['AC50'] = 'X'
            elif escolaridad == 'Técnico o tecnológico ':
                ws['AD50'] = 'X'
            elif escolaridad == 'Profesional':
                ws['AE50'] = 'X'
            elif escolaridad == 'Posgrado':
                ws['AG50'] = 'X'
        else:
            print('Campo vacio')

        procedencia = df_fila['Procedencia.2']
        if pd.notna(escolaridad):
            if procedencia == 'Vereda':
                ws['AJ50'] = 'X'
            elif procedencia == 'Municipio':
                ws['AM50'] = 'X'
            elif procedencia == 'Otro':
                ws['AO50'] = 'X'
        else:
            print('Campo vacio')

        residencia = df_fila['Residencia.2']
        if pd.notna(escolaridad):
            if residencia == 'Vereda':
                ws['AQ50'] = 'X'
            elif residencia == 'Municipio':
                ws['AS50'] = 'X'
            elif residencia == 'Otro':
                ws['AU50'] = 'X'
        else:
            print('Campo vacio')

        # Gerentes y directivos

        ws['I52'] = df_fila['#.3']

        genero = df_fila['Género.3']
        if pd.notna(genero):
            if genero == 'Masculino':
                ws['M52'] = 'X'
            elif genero == 'Femenino':
                ws['K52'] = 'X'
        else:
            print('Campo vacio')

        contrato = df_fila['Contrato.3']
        if pd.notna(genero):
            if contrato == 'Termino Fijo':
                ws['O52'] = df_fila['¿Cuánto?.3']
            elif contrato == 'Indefinido':
                ws['S52'] = 'X'
        else:
            print('Campo vacio')

        ws['W52'] = df_fila['Jornal y turno laboral.3']

        escolaridad = df_fila['Escolaridad.3']
        if pd.notna(escolaridad):
            if escolaridad == 'Primaria':
                ws['AB52'] = 'X'
            elif escolaridad == 'Bachillerato':
                ws['AC52'] = 'X'
            elif escolaridad == 'Técnico o tecnológico ':
                ws['AD52'] = 'X'
            elif escolaridad == 'Profesional':
                ws['AE52'] = 'X'
            elif escolaridad == 'Posgrado':
                ws['AG52'] = 'X'
        else:
            print('Campo vacio')

        procedencia = df_fila['Procedencia.3']
        if pd.notna(escolaridad):
            if procedencia == 'Vereda':
                ws['AJ52'] = 'X'
            elif procedencia == 'Municipio':
                ws['AM52'] = 'X'
            elif procedencia == 'Otro':
                ws['AO52'] = 'X'
        else:
            print('Campo vacio')

        residencia = df_fila['Residencia.3']
        if pd.notna(escolaridad):
            if residencia == 'Vereda':
                ws['AQ52'] = 'X'
            elif residencia == 'Municipio':
                ws['AS52'] = 'X'
            elif residencia == 'Otro':
                ws['AU52'] = 'X'
        else:
            print('Campo vacio')
        
        # 28. ¿Contrata servicios profesionales?
        if df_fila['Contrata servicios profesionales * Sí (Responder 30 y 31)'] == 'Si':
            ws['L56'] = 'X'

            # ¿Qué tipo de servicios? 
            servicios = df_fila['¿Qué tipo de servicios?']
            if pd.notna(servicios):
                if servicios == 'Contaduría':
                    ws['AK56'] = 'X'
                elif servicios == 'Consultoría':
                    ws['AT56'] = 'X'
                elif servicios == 'Asesoría legal':
                    ws['AK57'] = 'X'
                elif servicios == 'Otros':
                    ws['AT57'] = 'X'
                    ws['AE58'] = df_fila['¿Cuál?.1']
            else:
                print('Campo vacio')

            # 30. Con qué frecuencia contrata servicios profesionales

            ws['A59'] = df_fila['Servicio 1.3']

            frecuencia_servicios1 = df_fila['Frecuencia']
            if pd.notna(frecuencia_servicios1):
                if frecuencia_servicios1 == 'Mensual':
                    ws['F59'] = 'X'
                elif frecuencia_servicios1 == 'Semestral':
                    ws['J59'] = 'X'
                elif frecuencia_servicios1 == 'Trimestral':
                    ws['P59'] = 'X'
                elif frecuencia_servicios1 == 'Anual':
                    ws['U59'] = 'X'

            ws['A60'] = df_fila['Servicio 2.3']

            frecuencia_servicios2 = df_fila['Frecuencia.1']
            if pd.notna(frecuencia_servicios2):
                if frecuencia_servicios2 == 'Mensual':
                    ws['F60'] = 'X'
                elif frecuencia_servicios2 == 'Semestral':
                    ws['J60'] = 'X'
                elif frecuencia_servicios2 == 'Trimestral':
                    ws['P60'] = 'X'
                elif frecuencia_servicios2 == 'Anual':
                    ws['U60'] = 'X'
            else:
                print('Campo vacio')

            ws['A61'] = df_fila['Servicio 3.3']

            frecuencia_servicios3 = df_fila['Frecuencia.2']
            if pd.notna(frecuencia_servicios3):
                if frecuencia_servicios3 == 'Mensual':
                    ws['F61'] = 'X'
                elif frecuencia_servicios3 == 'Semestral':
                    ws['J61'] = 'X'
                elif frecuencia_servicios3 == 'Trimestral':
                    ws['P61'] = 'X'
                elif frecuencia_servicios3 == 'Anual':
                    ws['U61'] = 'X'
            else:
                print('Campo vacio')

            ws['A62'] = df_fila['Servicio 4.3']

            frecuencia_servicios4 = df_fila['Frecuencia.3']
            if pd.notna(frecuencia_servicios4):
                if frecuencia_servicios4 == 'Mensual':
                    ws['F62'] = 'X'
                elif frecuencia_servicios4 == 'Semestral':
                    ws['J62'] = 'X'
                elif frecuencia_servicios4 == 'Trimestral':
                    ws['P62'] = 'X'
                elif frecuencia_servicios4 == 'Anual':
                    ws['U62'] = 'X'
            else:
                print('Campo vacio')

            ws['A63'] = df_fila['Servicio 5.3']
            frecuencia_servicios5 = df_fila['Frecuencia.4']
            if pd.notna(frecuencia_servicios5):
                if frecuencia_servicios5 == 'Mensual':
                    ws['F63'] = 'X'
                elif frecuencia_servicios5 == 'Semestral':
                    ws['J63'] = 'X'
                elif frecuencia_servicios5 == 'Trimestral':
                    ws['P63'] = 'X'
                elif frecuencia_servicios5 == 'Anual':
                    ws['U63'] = 'X'       
            else:
                print('Campo vacio')

            # 31. ¿Cuál es el monto pagado por estos servicios durante el último semestre?
            ws['AE62'] = df_fila['¿Cuál es el monto pagado por estos servicios durante el último semestre?']


        elif df_fila['Contrata servicios profesionales * Sí (Responder 30 y 31)'] == 'No':
            ws['N56'] = 'X'
        


        # D. REMUNERACIONES
        ws['Z65'] = df_fila['Salarios pagados a la mano de obra calificada']
        ws['Z66'] = df_fila['Salarios pagados a la mano de obra no calificada']
        ws['Z67'] = df_fila['Salarios pagados a empleados y administrativos']
        ws['Z68'] = df_fila['Salarios pagados a gerentes y directivos']
        ws['Z69'] = df_fila['Total remuneraciones']

