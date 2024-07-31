import pandas as pd
import os

class InformeSegundo:
    def valorCol(self, base_name, index, df_fila):
        if index == 0:
            return df_fila.get(base_name, '')
        else:
            return df_fila.get(f'{base_name}.{index}', '')

    def determinar_mayoria_unidad(self, unidades):
        """
        Esta función determina la unidad mayoritaria en una lista de unidades.
        """
        return max(set(unidades), key=unidades.count)

    def lecturaArchivoSegundo(self):
        rutaArchivoInicial = os.getcwd() + "\\Censo Económico Maute.xlsm"
        xl = pd.ExcelFile(rutaArchivoInicial)
        df = xl.parse(sheet_name='FORMATO 2. AGROPECUARIO', header=None)
        df = df.drop(columns=[0, 1, 2, 3])
        df_T = df.transpose()
        df_T.columns = df_T.iloc[0]
        df_T = df_T.drop(df_T.index[0])
        df_T.columns = df_T.columns.str.strip()     
        df_T.columns = pd.io.common.dedup_names(df_T.columns, is_potential_multiindex=False)
        df_enc = df_T.reset_index(drop = True)

        return df_enc

    def crearArchivoSegundo(self, ws, df_fila):
        # A. IDENTIFICACIÓN ENTREVISTADO
        ws['AO1'] = df_fila['Encuesta No.']

        if pd.notna(df_fila['Fecha']):
            fecha_str = str(df_fila['Fecha'])
            if '/' in fecha_str:
                ws['AL2'] = fecha_str.split('/')[0]
                ws['AO2'] = fecha_str.split('/')[1]
                ws['AU2'] = fecha_str.split('/')[2]
            elif '-' in fecha_str:
                ws['AM2'] = fecha_str.split('-')[0]
                ws['AP2'] = fecha_str.split('-')[1]
                ws['AU2'] = fecha_str.split('-')[2]
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')  
        
        ws['AL3'] = df_fila['Encuestador']

        ws['E7'] = df_fila['Nombre']
        ws['AC7'] = df_fila['Empresa']
        ws['AQ7'] = df_fila['Cargo']


        pertenece_asociacion = df_fila['¿Pertenece a alguna asociación?']
        if pd.notna(pertenece_asociacion):
            if pertenece_asociacion == 'Sí':
                ws['M8'] = 'X'
                ws['W8'] = df_fila['¿Cuál?']
            elif pertenece_asociacion == 'No':
                ws['AO8'] = 'X'
        else:
            print("Campo vacío")

        # B. INFORMACIÓN E IDENTIFICACIÓN GENERAL DEL PREDIO

        # Área Total
        ws['G11'] = df_fila['Área Total (Ha)']

        # Tipo de uso
        ws['G14'] = df_fila['Cultivos (Ha)']
        ws['G15'] = df_fila['Pastos (Ha)']
        ws['G16'] = df_fila['Bosque natural (Ha)']
        ws['G17'] = df_fila['Rastrojo (Ha)']
        ws['G18'] = df_fila['Bosque plantado (Ha)']
        ws['AB14'] = df_fila['Tierras erosionadas (Ha)']
        ws['AB15'] = df_fila['Lagos y lagunas (m2)']
        ws['AB16'] = df_fila['Reservorios (m2)']
        ws['AB17'] = df_fila['Construcciones (m2)']
        ws['Y18'] = df_fila['Otros']

        # Precio de arrendamiento por Ha
        ws['AN13'] = df_fila['Precio de arrendamiento por Ha']

        # Observaciones
        ws['AM16'] = df_fila['Observaciones']

        # ACTIVIDADES AGROPECUARIAS

        # Producto
        ws['F23'] = df_fila['Producto']

        # Columna inicial para cada cultivo
        cultivo_col_start = ['V', 'AF', 'AN']

        # Listas para almacenar las unidades de cada categoría
        unidades_area_cultivada = []
        unidades_establecimiento = []
        unidades_mantenimiento = []
        unidades_cosecha = []
        unidades_precio = []
        unidades_autoconsumo = []

        for i in range(3):  # Hay tres cultivos
            cultivo_prefix = cultivo_col_start[i]

            # Área cultivada
            ws[f'{cultivo_prefix}24'] = InformeSegundo().valorCol('Área cultivada', i, df_fila)
            
            # Unidad de Área cultivada
            unidad_area_cultivada = InformeSegundo().valorCol('Unidad', 6 * i, df_fila)
            unidades_area_cultivada.append(unidad_area_cultivada)
            
            # Número de cosechas por año
            ws[f'{cultivo_prefix}25'] = InformeSegundo().valorCol('No. de cosechas por año', i, df_fila)
            
            # Costos de establecimiento
            ws[f'{cultivo_prefix}26'] = InformeSegundo().valorCol('Costos de establecimiento', i, df_fila)
            
            # Unidad de Costos de establecimiento
            unidad_establecimiento = InformeSegundo().valorCol('Unidad', 6 * i + 1, df_fila)
            unidades_establecimiento.append(unidad_establecimiento)
            
            # Costos de mantenimiento
            ws[f'{cultivo_prefix}27'] = InformeSegundo().valorCol('Costos de mantenimiento', i, df_fila)
            
            # Unidad de Costos de mantenimiento
            unidad_mantenimiento = InformeSegundo().valorCol('Unidad', 6 * i + 2, df_fila)
            unidades_mantenimiento.append(unidad_mantenimiento)
            
            # Costos de cosecha
            ws[f'{cultivo_prefix}28'] = InformeSegundo().valorCol('Costos de cosecha', i, df_fila)
            
            # Unidad de Costos de cosecha
            unidad_cosecha = InformeSegundo().valorCol('Unidad', 6 * i + 3, df_fila)
            unidades_cosecha.append(unidad_cosecha)
            
            # Volumen de producción
            ws[f'{cultivo_prefix}29'] = InformeSegundo().valorCol('Volumen de producción', i, df_fila)
            
            # Precio de venta
            ws[f'{cultivo_prefix}30'] = InformeSegundo().valorCol('Precio de venta', i, df_fila)
            
            # Unidad de Precio de venta
            unidad_precio = InformeSegundo().valorCol('Unidad', 6 * i + 4, df_fila)
            unidades_precio.append(unidad_precio)

            # Autoconsumo
            ws[f'{cultivo_prefix}31'] = InformeSegundo().valorCol('Autoconsumo', i, df_fila)
            
            # Unidad de Autoconsumo
            unidad_autoconsumo = InformeSegundo().valorCol('Unidad', 6 * i + 5, df_fila)
            unidades_autoconsumo.append(unidad_autoconsumo)
        
        # Determinar las unidades mayoritarias
        unidad_area_cultivada_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_area_cultivada)
        unidad_establecimiento_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_establecimiento)
        unidad_mantenimiento_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_mantenimiento)
        unidad_cosecha_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_cosecha)
        unidad_precio_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_precio)
        unidad_autoconsumo_mayoria = InformeSegundo().determinar_mayoria_unidad(unidades_autoconsumo)

        # Rellenar la unidad mayoritaria en el campo correspondiente
        if unidad_area_cultivada_mayoria == 'Ha':
            ws['Q24'] = 'X'
        elif unidad_area_cultivada_mayoria == 'm2':
            ws['T24'] = 'X'
        
        if unidad_establecimiento_mayoria == 'Ha':
            ws['Q26'] = 'X'
        elif unidad_establecimiento_mayoria == 'm2':
            ws['T26'] = 'X'

        if unidad_mantenimiento_mayoria == 'Ha':
            ws['Q27'] = 'X'
        elif unidad_mantenimiento_mayoria == 'm2':
            ws['T27'] = 'X'

        if unidad_cosecha_mayoria == 'Ha':
            ws['Q28'] = 'X'
        elif unidad_cosecha_mayoria == 'm2':
            ws['T28'] = 'X'

        if unidad_precio_mayoria == 'Unidad':
            ws['N30'] == 'X'
        elif unidad_precio_mayoria == 'Kg':
            ws['Q30'] = 'X'
        elif unidad_precio_mayoria == 'Ton':
            ws['T30'] = 'X'

        if unidad_autoconsumo_mayoria == 'Carga':
            ws['N31'] = 'X'
        elif unidad_autoconsumo_mayoria == 'Kg':
            ws['P31'] = 'X'
        elif unidad_autoconsumo_mayoria == '@':
            ws['R31'] = 'X'
        elif unidad_autoconsumo_mayoria == 'Ton':
            ws['U31'] = 'X' 

        ws['L32'] = df_fila['¿Destino final del producto?']
        

        continuidad = df_fila['Sobre la actividad, piensa: Continuidad']
        if pd.notna(continuidad):
            if continuidad == 'Continuar con la actividad':
                ws['G36'] == 'X'
                ws['AP36'] == 'X'
            elif continuidad == 'Finalizar la actividad':
                ws['I36'] == 'X'
                ws['AN36'] == 'X'  
        else:
            print('Campo vacío') 

        produccion = df_fila['Sobre la actividad, piensa: Producción']
        if pd.notna(produccion):
            if produccion == 'Ampliar la producción':
                ws['Q36'] == 'X'
                ws['AF36'] == 'X'               
            elif produccion == 'Permanecer con la misma producción':
                ws['S36'] == 'X'
                ws['AD36'] == 'X' 
            elif produccion == 'Ninguna de las anteriores':
                ws['AF36'] == 'X'
                ws['S36'] == 'X'
        else:
            print('Campo vacío') 

        ws['AS36'] == df_fila['¿Por qué?']

    # Llenar raza y # de cabezas
        for i in range(3):
            # Leche o Cría
            ws[f'N{41 + i}'] = InformeSegundo().valorCol('Raza', i, df_fila)
            ws[f'AB{41 + i}'] = InformeSegundo().valorCol('# de cabezas', i, df_fila)

            # Carne
            ws[f'AJ{41 + i}'] = InformeSegundo().valorCol('Raza', i + 3, df_fila)
            ws[f'AT{41 + i}'] = InformeSegundo().valorCol('# de cabezas', i + 3, df_fila)

        # Número de terneros    
        ws['Q44'] = InformeSegundo().valorCol('Número de reses en producción', df_fila)
        ws['AK44'] = InformeSegundo().valorCol('Número de reses en producción.1', df_fila)

        ws['Q45'] = InformeSegundo().valorCol('Número de terneros', df_fila)
        ws['AK45'] = InformeSegundo().valorCol('Número de terneros.1', df_fila)  

        # Número de novillos
        ws['Q46'] = InformeSegundo().valorCol('Número de novillos', df_fila)
        ws['AK46'] = InformeSegundo().valorCol('Número de novillos.1', df_fila)

        # Número de novillas
        ws['Q47'] = InformeSegundo().valorCol('Número de novillas', df_fila)
        ws['AK47'] = InformeSegundo().valorCol('Número de novillas.1', df_fila)

        # Número de reproductores
        ws['Q48'] = InformeSegundo().valorCol('Número de reproductores', df_fila)
        ws['AK48'] = InformeSegundo().valorCol('Número de reproductores.1', df_fila)

        # Número de partos al año
        ws['Q49'] = InformeSegundo().valorCol('Número de partos al año', df_fila)
        ws['AK49'] = InformeSegundo().valorCol('Número de partos al año.1', df_fila)

        # Número de vacas para ordeño
        ws['Q50'] = InformeSegundo().valorCol('Número de vacas para ordeño', df_fila)
        ws['AK50'] = InformeSegundo().valorCol('Número de vacas para ordeño.1', df_fila)

        # Tiempo de venta después destetado
        ws['Q51'] = InformeSegundo().valorCol('Tiempo de venta después destetado', df_fila)
        ws['AK51'] = InformeSegundo().valorCol('Tiempo de venta después destetado.1', df_fila)

        # Peso promedio para la venta en Kg
        ws['Q52'] = InformeSegundo().valorCol('Peso promedio para la venta en Kg', df_fila)
        ws['AK52'] = InformeSegundo().valorCol('Peso promedio para la venta en Kg.1', df_fila)

        # Precio promedio

        ws['N53'] = df_fila['Precio promedio: Litro']
        ws['X53'] = df_fila['Precio promedio: Botella']
        ws['AH53'] = df_fila['Precio promedio: Kg']
        ws['AQ53'] = df_fila['Precio promedio: Cabeza']

        # ¿Cada cuánto produce? 

        frecuencia_leche = df_fila['¿Cada cuánto produce?']
        if pd.notna(frecuencia_leche):
            if frecuencia_leche == 'Diario':
                ws['P54'] = 'X'
            elif frecuencia_leche == 'Semanal':
                ws['W54'] = 'X'
            elif frecuencia_leche == 'Otro':
                ws['AD54'] = 'X'

        # Frecuencia de producción para Carne
        frecuencia_carne = df_fila['¿Cada cuánto produce?.1']
        if pd.notna(frecuencia_carne):
            if frecuencia_carne == 'Mensual':
                ws['AJ54'] = 'X'
            elif frecuencia_carne == 'Trimestral':
                ws['AO54'] = 'X'
            elif frecuencia_carne == 'Otro':
                ws['AU54'] = 'X'

        # Litros y Botella para Leche o Cría
        promedio_produccion_litro = df_fila['Promedio de producción Litro']
        if pd.notna(promedio_produccion_litro):
            ws['M55'] = promedio_produccion_litro

        promedio_produccion_botella = df_fila['Promedio de producción Botella']
        if pd.notna(promedio_produccion_botella):
            ws['W55'] = promedio_produccion_botella

        # Kg y Cabezas para Carne
        promedio_produccion_kg = df_fila['Promedio de producción Kg']
        if pd.notna(promedio_produccion_kg):
            ws['AG55'] = promedio_produccion_kg

        promedio_produccion_cabezas = df_fila['Promedio de producción Cabezas']
        if pd.notna(promedio_produccion_cabezas):
            ws['AP55'] = promedio_produccion_cabezas   

        if pd.notna(df_fila['¿Destino final del producto?.1']) & pd.notna(df_fila['¿Destino final del producto?.2']):
            ws['L56'] = 'Leche o cría: ' + df_fila['¿Destino final del producto?.1'] + ' Carne: ' + df_fila['¿Destino final del producto?.2']

        continuidad1 = df_fila['Sobre la actividad, piensa: Continuidad.1']
        if pd.notna(continuidad1):
            if continuidad1 == 'Continuar con la actividad':
                ws['G60'] == 'X'
                ws['AP60'] == 'X'
            elif continuidad1 == 'Finalizar la actividad':
                ws['I60'] == 'X'
                ws['AN60'] == 'X'  
        else:
            print('Campo vacío') 

        produccion1 = df_fila['Sobre la actividad, piensa: Producción.1']
        if pd.notna(produccion1):
            if produccion1 == 'Ampliar la producción':
                ws['Q60'] == 'X'
                ws['AF60'] == 'X'               
            elif produccion1 == 'Permanecer con la misma producción':
                ws['S60'] == 'X'
                ws['AD60'] == 'X'  
            elif produccion1 == 'Ninguna de las anteriores':
                ws['S60'] == 'X'
                ws['AF60'] == 'X'
        else:
            print('Campo vacío') 

        ws['AS60'] == df_fila['¿Por qué?.1']


        ws['T64'] =df_fila['Raza']
        ws['T65'] =df_fila['# Hembras']
        ws['T66'] =df_fila['# Machos']
        ws['T67'] =df_fila['Tiene Marranos para la venta']
        ws['T68'] =df_fila['Peso promedio para la venta por animal (Kg)']
        ws['T69'] =df_fila['# Promedio de animales vendidos por año']
        ws['T70'] =df_fila['Cantidad empleada para autoconsumo (Kg)']
        ws['T71'] =df_fila['Precio unitario de venta']
        # if df_fila['Unidad'] == "Animal":
        #     ws['Q71'] = 'X'
        # elif df_fila['Unidad'] == "Kg":
        #     ws['S71'] = 'X'
        ws['T72'] =df_fila['Costo aproximado de producción']
        # if df_fila['Unidad.1'] == "Animal":
        #     ws['Q71'] = 'X'
        # elif df_fila['Unidad.1'] == "Kg":
        #     ws['S71'] = 'X'


        ws['AF64'] =df_fila['Raza.1']
        ws['AF65'] =df_fila['# Hembras.1']
        ws['AF66'] =df_fila['# Machos.1']
        ws['AE67'] =df_fila['Tiene Marranos para la venta.1']
        ws['AF68'] =df_fila['Peso promedio para la venta por animal (Kg).1']
        ws['AF69'] =df_fila['# Promedio de animales vendidos por año.1']
        ws['AF70'] =df_fila['Cantidad empleada para autoconsumo (Kg).1']
        ws['AF71'] =df_fila['Precio unitario de venta.1']
        ws['AF72'] =df_fila['Costo aproximado de producción.1']


        ws['AO64'] =df_fila['Raza.2']
        ws['AO65'] =df_fila['# Hembras.2']
        ws['AO66'] =df_fila['# Machos.2']
        ws['AM67'] =df_fila['Tiene Marranos para la venta.2']
        ws['AO68'] =df_fila['Peso promedio para la venta por animal (Kg).2']
        ws['AO69'] =df_fila['# Promedio de animales vendidos por año.2']
        ws['AO70'] =df_fila['Cantidad empleada para autoconsumo (Kg).2']
        ws['AO71'] =df_fila['Precio unitario de venta.2']
        ws['AO72'] =df_fila['Costo aproximado de producción.2']


        continuidad2 = df_fila['Sobre la actividad, piensa: Continuidad.1']
        if pd.notna(continuidad2):
            if continuidad2 == 'Continuar con la actividad':
                ws['G77'] == 'X'
                ws['AP77'] == 'X'
            elif continuidad2 == 'Finalizar la actividad':
                ws['I77'] == 'X'
                ws['AN77'] == 'X'  
        else:
            print('Campo vacío') 

        produccion2 = df_fila['Sobre la actividad, piensa: Producción.1']
        if pd.notna(produccion2):
            if produccion2 == 'Ampliar la producción':
                ws['Q77'] == 'X'
                ws['AF77'] == 'X'               
            elif produccion2 == 'Permanecer con la misma producción':
                ws['S77'] == 'X'
                ws['AD77'] == 'X'  
            elif produccion2 == 'Ninguna de las anteriores':
                ws['S77'] == 'X'
                ws['AF77'] == 'X'
        else:
            print('Campo vacío') 

        ws['AS77'] == df_fila['¿Por qué?.2']


        if df_fila['Tipo de explotación'] == "Cría":
            ws['R80'] = 'X'
        elif df_fila['Tipo de explotación'] == "Engorde":
            ws['AC80'] = 'X'
        elif df_fila['Tipo de explotación'] == "Ponedoras":
            ws['AK80'] = 'X'
        elif df_fila['Tipo de explotación'] == "Gallina campesina":
            ws['AU80'] = 'X'

        ws['N82'] = df_fila['# Animales']
        ws['N83'] = df_fila['Producción mensual (Aves)']
        ws['N84'] = df_fila['Unidades vendidas al mes (Aves)']
        ws['O86'] = df_fila['Costo por animal']

        ws['AB82'] = df_fila['# Animales.1']
        ws['AB83'] = df_fila['Producción mensual (Aves).1']
        ws['AB84'] = df_fila['Unidades vendidas al mes (Aves).1']
        ws['AC86'] = df_fila['Costo por animal.1']

        ws['AM82'] = df_fila['# Animales.2']
        ws['AM83'] = df_fila['Producción mensual (Aves).2']
        ws['AM84'] = df_fila['Unidades vendidas al mes (Aves).2']
        ws['AN86'] = df_fila['Costo por animal.2']




        continuidad3 = df_fila['Sobre la actividad, piensa: Continuidad.1']
        if pd.notna(continuidad2):
            if continuidad3 == 'Continuar con la actividad':
                ws['G92'] == 'X'
                ws['AP92'] == 'X'
            elif continuidad3 == 'Finalizar la actividad':
                ws['I92'] == 'X'
                ws['AN92'] == 'X'  
        else:
            print('Campo vacío') 

        produccion3 = df_fila['Sobre la actividad, piensa: Producción.1']
        if pd.notna(produccion3):
            if produccion3 == 'Ampliar la producción':
                ws['Q92'] == 'X'
                ws['AF92'] == 'X'               
            elif produccion3 == 'Permanecer con la misma producción':
                ws['S92'] == 'X'
                ws['AD92'] == 'X'  
            elif produccion3 == 'Ninguna de las anteriores':
                ws['S92'] == 'X'
                ws['AF92'] == 'X'
        else:
            print('Campo vacío') 

        ws['AS92'] == df_fila['¿Por qué']

        

        

