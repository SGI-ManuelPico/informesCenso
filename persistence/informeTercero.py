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
        archivoInicial.columns = archivoInicial.iloc[0].str.strip()
        archivoInicial = archivoInicial.drop(archivoInicial.index[0])
        archivoInicial.columns = pd.io.common.dedup_names(archivoInicial.columns, is_potential_multiindex=False)

        archivoInicial = archivoInicial.reset_index(drop = True)

        return archivoInicial
    
    def crearArchivoTercero(self, hoja, fila):
        """
        Creación del tercer archivo del censo económico.
        """


        hoja['AI1'] = fila["Encuesta No."]

        if pd.notna(fila['Fecha']):
            fecha_str = str(fila['Fecha'])
            if '/' in fecha_str:
                hoja['AK2'] = fecha_str.split('/')[0]
                hoja['AN2'] = fecha_str.split('/')[1]
                hoja['AS2'] = fecha_str.split('/')[2]
            elif '-' in fecha_str:
                hoja['AK2'] = fecha_str.split('-')[0]
                hoja['AN2'] = fecha_str.split('-')[1]
                hoja['AS2'] = fecha_str.split('-')[2] ####### LLENAR FECHA EN ESPACIOS VACÍOS Y NO SOBRE EL SÍMBOLO.
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')

        hoja['AK3'] = fila["Encuestador"]
        hoja['F6'] = fila["Nombre"]
        hoja['AC6'] = fila["Empresa"]
        hoja['AP6'] = fila["Cargo"]

        if fila["¿Pertenece a alguna asociación?"] == 'Si':
            hoja['S7'] = 'X'

        elif fila["¿Pertenece a alguna asociación?"] == 'No':
            hoja['Y7'] = 'X'

        hoja['AF7'] = fila["Otro, ¿Cuál?"]

        actividad = fila['¿Qué tipo de productos comercializa?']
        if actividad == 'Agrícola':
            hoja['L11'] = 'X'
        elif actividad == 'Pecuario':
            hoja['L12'] = 'X'
        elif actividad == 'Víveres':
            hoja['L13'] = 'X'
        elif actividad == 'Agua en botella/bolsa':
            hoja['L14'] = 'X'
        elif actividad == 'Licores':
            hoja['U11'] = 'X'
        elif actividad == 'Productos Naturales: Animal':
            hoja['U13'] = 'X'
        elif actividad == 'Productos Naturales: Vegetal':
            hoja['Y13'] = 'X'
        hoja['T14'] = fila['Otro, ¿Cuáles?']

        hoja['C16'] = fila["¿Cuál es el principal producto que comercializa?"]

        actividad2 = fila['¿Con qué frecuencia compra los productos que comercializa?']
        if actividad2 == 'Diario':
            hoja['L20'] = 'X'
        elif actividad2 == 'Semanal':
            hoja['L21'] = 'X'
        elif actividad2 == 'Quincenal':
            hoja['L22'] = 'X'
        elif actividad2 == 'Mensual':
            hoja['L23'] = 'X'
        elif actividad2 == 'Semestral':
            hoja['V22'] = actividad2
        elif actividad2 == 'Otra':
            hoja['V22'] = fila["Otro"]

        hoja['C28'] = fila["Producto"]
        hoja['I28'] = fila["Cantidad"]
        hoja['M28'] = fila["Unidad de medida"]
        hoja['S28'] = fila["Valor"]
        hoja['C29'] = fila["Producto.1"]
        hoja['I29'] = fila["Cantidad.1"]
        hoja['M29'] = fila["Unidad de medida.1"]
        hoja['S29'] = fila["Valor.1"]
        hoja['C30'] = fila["Producto.2"]
        hoja['I30'] = fila["Cantidad.2"]
        hoja['M30'] = fila["Unidad de medida.2"]
        hoja['S30'] = fila["Valor.2"]
        hoja['C31'] = fila["Producto.3"]
        hoja['I31'] = fila["Cantidad.3"]
        hoja['M31'] = fila["Unidad de medida.3"]
        hoja['S31'] = fila["Valor.3"]

        hoja['D34'] = fila["Producto.4"]
        hoja['R34'] = fila["Precio"]
        hoja['D35'] = fila["Producto.5"]
        hoja['R35'] = fila["Precio.1"]
        hoja['D36'] = fila["Producto.6"]
        hoja['R36'] = fila["Precio.2"]

        actividad3 = fila["Señale el tipo de emplazamiento"]
        if actividad3 == 'Local':
            hoja['Q39'] = 'X'
        elif actividad3 == 'Puesto Fijo':
            hoja['Q40'] = 'X'
        elif actividad3 == 'Vivienda económica':
            hoja['Q41'] = 'X'
        elif actividad3 == 'Venta ambulante':
            hoja['Q42'] = 'X'

        actividad4 = fila["¿Cuál fue el valor promedio de ventas en el último mes?"]
        if actividad4 == 'Inferiores a $600.000':
            hoja['AN10'] = 'X'
        elif actividad4 == 'Entre $ 601.000 y $ 1.500.000':
            hoja['AN11'] = 'X'
        elif actividad4 == 'Entre $ 1.501.000 y $ 3.000.000':
            hoja['AN12'] = 'X'
        elif actividad4 == 'Superior a $ 3.000.000':
            hoja['AN13'] = 'X'

        actividad5 = fila["Vende principalmente en:"]
        if actividad5 == 'Sitio':
            hoja['AH16'] = 'X'
        elif actividad5 == 'Vereda':
            hoja['AH17'] = 'X'
        elif actividad5 == 'Casco Urbano':
            hoja['AH18'] = 'X'
        elif actividad5 == 'Otros Municipios y/o Veredas':
            hoja['AN16'] = 'X'
            hoja['AO18'] = fila["¿Cuáles?"]

        if fila["¿Lleva libros contables del establecimiento?"] == 'Si':
            hoja['AO22'] = 'X'

        elif fila["¿Lleva libros contables del establecimiento?"] == 'No':
            hoja['AQ22'] = 'X'

        p1 = fila["Producto 1"]
        p2 = fila["Prodcuto 2"]
        p3 = fila["Producto 3"]
        p4 = fila["Prodcuto 4"]
        p5 = fila["Producto 5"]

        hoja['AE26'] = p1 + " " + p2 + " " + p3 + " " + p4 + " " + p5

        if fila["Hidrocarburos"] != "":
            hoja['AH29'] = 'X'
            hoja['AM29'] = fila['Hidrocarburos']
        if fila["Otro.1"] != "":
            hoja['AH30'] = 'X'
            hoja['AM30'] = fila['Otro.1']

        actividad5 = fila['¿Con qué frecuencia compra productos en otras veredas y/o municipios?']
        if actividad5 == 'Diario':
            hoja['AJ33'] = 'X'
        elif actividad5 == 'Semanal':
            hoja['AJ34'] = 'X'
        elif actividad5 == 'Quincenal':
            hoja['AJ35'] = 'X'
        elif actividad5 == 'Mensual':
            hoja['AJ36'] = 'X'
        elif actividad2 == 'Otro':
            hoja['AQ33'] = 'X'
            hoja['AP34'] = fila["Otro, ¿Cuál?.1"]

        if fila["Sobre la actividad, piensa: Continuidad"] == "Continuar con la actividad":
            hoja['AN39'] = 'X'
        elif fila["Sobre la actividad, piensa: Continuidad"] == "Finalizar la actividad":
            hoja['AP39'] = 'X'
        if fila["Sobre la actividad, piensa: Producción"] == "Ampliar la producción":
            hoja['AP40'] = 'X'
            hoja['AN41'] = 'X'
        elif fila["Sobre la actividad, piensa: Producción"] == "Permanecer con la misma producción":
            hoja['AN40'] = 'X'
            hoja['AP41'] = 'X'

        actividad6 = fila['¿De dónde se abastece del recurso hídrico?']
        if actividad6 == 'Aljibe':
            hoja['I45'] = 'X'
        elif actividad6 == 'Acueducto Veredal':
            hoja['I46'] = 'X'
        elif actividad6 == 'Otro':
            hoja['I47'] = 'X'
            hoja['R47'] = fila['Otro, ¿Cuál?.2']
    
        hoja['U45'] = fila["Forma de extracción"]
        hoja['P46'] = fila["Cantidad estimada (m3)"]

        if fila["¿Cuenta con servicio de alcantarillado?"] == "Si":
            hoja['AP44'] = 'X'
        elif fila["¿Cuenta con servicio de alcantarillado?"] == "No":
            hoja['AR44'] = 'X'

        hoja['AF45'] = fila["¿Cuál?"]

        actividad7 = fila['¿Qué tipo de energía utiliza?']
        if actividad7 == 'Energía Eléctrica':
            hoja['AG47'] = 'X'
        elif actividad7 == 'Energía Solar':
            hoja['AN47'] = 'X'
        elif actividad7 == 'Otro':
            hoja['AS47'] = fila['Otro, ¿Cuál?.3']

        actividad8 = fila['¿De dónde proviene la energía que utiliza para la cocción de alimentos?']
        if actividad8 == 'Energía elétrica':
            hoja['AG48'] = 'X'
        elif actividad8 == 'Leña':
            hoja['AM48'] = 'X'
        elif actividad8 == 'Gas':
            hoja['AO48'] = 'X'
        elif actividad8 == 'Otro':
            hoja['AS48'] = fila['Otro, ¿Cuál?.4']

        if fila["Contrata algún tipo de mano de obra"] == "Si":
            hoja['Z50'] = 'X'

            #### Persona 1 ####

            if fila["Tipo de mano de obra"] == "Familiar":
                hoja['B54'] = 'X'
            elif fila["Tipo de mano de obra"] == "Contratado":
                hoja['D54'] = 'X'

            hoja['E54'] = fila["Cargo.1"]

            if fila["Género"] == "Masculino":
                hoja['M54'] = 'X'
            elif fila["Género"] == "Femenino":
                hoja['K54'] = 'X'

            hoja['N54'] = fila["Edad (años)"]
            hoja['Q54'] = fila["Duración jornada (horas)"]

            actividad9 = fila['Escolaridad']
            if actividad9 == 'Primaria':
                hoja['S54'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U54'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W54'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y54'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA54'] = 'X'

            if fila['Contrato'] == 'Tem.':
                hoja['AC54'] = 'X'
            elif fila['Contrato'] == 'Fij':
                hoja['AG54'] = 'X'
            
            if fila['Pago de seguridad'] == 'Si':
                hoja['AH54'] = 'X'
                hoja['AJ54'] = ''
            elif fila['Pago de seguridad'] == 'No':
                hoja['AH54'] = ''
                hoja['AJ54'] = 'X'

            hoja['AL54'] = fila["Procedencia"]
            hoja['AM54'] = fila["Residencia"]
            hoja['AN54'] = fila["Tiempo trabajado"]
            hoja['AO54'] = fila["# Personas núcleo familiar"]
            hoja['AP54'] = fila["Personas a cargo"]
            hoja['AQ54'] = fila["Lugar de residencia familiar"]
        
            actividad10 = fila['Remuneración']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT54'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU54'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV54'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW54'] = 'X'

            ##### Persona 2 #####

            if fila["Tipo de mano de obra.1"] == "Familiar":
                hoja['B55'] = 'X'
            elif fila["Tipo de mano de obra.1"] == "Contratado":
                hoja['D55'] = 'X'

            hoja['E55'] = fila["Cargo.2"]

            if fila["Género.1"] == "Masculino":
                hoja['M55'] = 'X'
            elif fila["Género.1"] == "Femenino":
                hoja['K55'] = 'X'

            hoja['N55'] = fila["Edad (años).1"]
            hoja['Q55'] = fila["Duración jornada (horas).1"]

            actividad9 = fila['Escolaridad.1']
            if actividad9 == 'Primaria':
                hoja['S55'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U55'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W55'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y55'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA55'] = 'X'

            if fila['Contrato.1'] == 'Tem.':
                hoja['AC55'] = 'X'
            elif fila['Contrato.1'] == 'Fij':
                hoja['AG55'] = 'X'
            
            if fila['Pago de seguridad.1'] == 'Si':
                hoja['AH55'] = 'X'
                hoja['AJ55'] = ''
            elif fila['Pago de seguridad.1'] == 'No':
                hoja['AH55'] = ''
                hoja['AJ55'] = 'X'

            hoja['AL55'] = fila["Procedencia.1"]
            hoja['AM55'] = fila["Residencia.1"]
            hoja['AN55'] = fila["Tiempo trabajado.1"]
            hoja['AO55'] = fila["# Personas núcleo familiar.1"]
            hoja['AP55'] = fila["Personas a cargo.1"]
            hoja['AQ55'] = fila["Lugar de residencia familiar.1"]
        
            actividad10 = fila['Remuneración.1']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT55'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU55'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV55'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW55'] = 'X'

            ##### Persona 3 #####

            if fila["Tipo de mano de obra.2"] == "Familiar":
                hoja['B56'] = 'X'
            elif fila["Tipo de mano de obra.2"] == "Contratado":
                hoja['D56'] = 'X'

            hoja['E56'] = fila["Cargo.3"]

            if fila["Género.2"] == "Masculino":
                hoja['M56'] = 'X'
            elif fila["Género.2"] == "Femenino":
                hoja['K56'] = 'X'

            hoja['N56'] = fila["Edad (años).2"]
            hoja['Q56'] = fila["Duración jornada (horas).2"]

            actividad9 = fila['Escolaridad.2']
            if actividad9 == 'Primaria':
                hoja['S56'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U56'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W56'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y56'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA56'] = 'X'

            if fila['Contrato.2'] == 'Tem.':
                hoja['AC56'] = 'X'
            elif fila['Contrato.2'] == 'Fij':
                hoja['AG56'] = 'X'
            
            if fila['Pago de seguridad.2'] == 'Si':
                hoja['AH56'] = 'X'
                hoja['AJ56'] = ''
            elif fila['Pago de seguridad.2'] == 'No':
                hoja['AH56'] = ''
                hoja['AJ56'] = 'X'

            hoja['AL56'] = fila["Procedencia.2"]
            hoja['AM56'] = fila["Residencia.2"]
            hoja['AN56'] = fila["Tiempo trabajado.2"]
            hoja['AO56'] = fila["# Personas núcleo familiar.2"]
            hoja['AP56'] = fila["Personas a cargo.2"]
            hoja['AQ56'] = fila["Lugar de residencia familiar.2"]
        
            actividad10 = fila['Remuneración.2']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT56'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU56'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV56'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW56'] = 'X'

            ##### Persona 4 #####

            if fila["Tipo de mano de obra.3"] == "Familiar":
                hoja['B57'] = 'X'
            elif fila["Tipo de mano de obra.3"] == "Contratado":
                hoja['D57'] = 'X'

            hoja['E57'] = fila["Cargo.4"]

            if fila["Género.3"] == "Masculino":
                hoja['M57'] = 'X'
            elif fila["Género.3"] == "Femenino":
                hoja['K57'] = 'X'

            hoja['N57'] = fila["Edad (años).3"]
            hoja['Q57'] = fila["Duración jornada (horas).3"]

            actividad9 = fila['Escolaridad.3']
            if actividad9 == 'Primaria':
                hoja['S57'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U57'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W57'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y57'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA57'] = 'X'

            if fila['Contrato.3'] == 'Tem.':
                hoja['AC57'] = 'X'
            elif fila['Contrato.3'] == 'Fij':
                hoja['AG57'] = 'X'
            
            if fila['Pago de seguridad.3'] == 'Si':
                hoja['AH57'] = 'X'
                hoja['AJ57'] = ''
            elif fila['Pago de seguridad.3'] == 'No':
                hoja['AH57'] = ''
                hoja['AJ57'] = 'X'

            hoja['AL57'] = fila["Procedencia.3"]
            hoja['AM57'] = fila["Residencia.3"]
            hoja['AN57'] = fila["Tiempo trabajado.3"]
            hoja['AO57'] = fila["# Personas núcleo familiar.3"]
            hoja['AP57'] = fila["Personas a cargo.3"]
            hoja['AQ57'] = fila["Lugar de residencia familiar.3"]
        
            actividad10 = fila['Remuneración.3']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT57'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU57'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV57'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW57'] = 'X'

            ##### Persona 5 #####

            if fila["Tipo de mano de obra.4"] == "Familiar":
                hoja['B58'] = 'X'
            elif fila["Tipo de mano de obra.4"] == "Contratado":
                hoja['D58'] = 'X'

            hoja['E58'] = fila["Cargo.5"]

            if fila["Género.4"] == "Masculino":
                hoja['M58'] = 'X'
            elif fila["Género.4"] == "Femenino":
                hoja['K58'] = 'X'

            hoja['N58'] = fila["Edad (años).4"]
            hoja['Q58'] = fila["Duración jornada (horas).4"]

            actividad9 = fila['Escolaridad.4']
            if actividad9 == 'Primaria':
                hoja['S58'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U58'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W58'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y58'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA58'] = 'X'

            if fila['Contrato.4'] == 'Tem.':
                hoja['AC58'] = 'X'
            elif fila['Contrato.4'] == 'Fij':
                hoja['AG58'] = 'X'
            
            if fila['Pago de seguridad.4'] == 'Si':
                hoja['AH58'] = 'X'
                hoja['AJ58'] = ''
            elif fila['Pago de seguridad.4'] == 'No':
                hoja['AH58'] = ''
                hoja['AJ58'] = 'X'

            hoja['AL58'] = fila["Procedencia.4"]
            hoja['AM58'] = fila["Residencia.4"]
            hoja['AN58'] = fila["Tiempo trabajado.4"]
            hoja['AO58'] = fila["# Personas núcleo familiar.4"]
            hoja['AP58'] = fila["Personas a cargo.4"]
            hoja['AQ58'] = fila["Lugar de residencia familiar.4"]
        
            actividad10 = fila['Remuneración.4']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT58'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU58'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV58'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW58'] = 'X'

            ##### Persona 6 #####

            if fila["Tipo de mano de obra.5"] == "Familiar":
                hoja['B59'] = 'X'
            elif fila["Tipo de mano de obra.5"] == "Contratado":
                hoja['D59'] = 'X'

            hoja['E59'] = fila["Cargo.6"]

            if fila["Género.5"] == "Masculino":
                hoja['M59'] = 'X'
            elif fila["Género.5"] == "Femenino":
                hoja['K59'] = 'X'

            hoja['N59'] = fila["Edad (años).5"]
            hoja['Q59'] = fila["Duración jornada (horas).5"]

            actividad9 = fila['Escolaridad.5']
            if actividad9 == 'Primaria':
                hoja['S59'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U59'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W59'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y59'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA59'] = 'X'

            if fila['Contrato.5'] == 'Tem.':
                hoja['AC59'] = 'X'
            elif fila['Contrato.5'] == 'Fij':
                hoja['AG59'] = 'X'
            
            if fila['Pago de seguridad.5'] == 'Si':
                hoja['AH59'] = 'X'
                hoja['AJ59'] = ''
            elif fila['Pago de seguridad.5'] == 'No':
                hoja['AH59'] = ''
                hoja['AJ59'] = 'X'

            hoja['AL59'] = fila["Procedencia.5"]
            hoja['AM59'] = fila["Residencia.5"]
            hoja['AN59'] = fila["Tiempo trabajado.5"]
            hoja['AO59'] = fila["# Personas núcleo familiar.5"]
            hoja['AP59'] = fila["Personas a cargo.5"]
            hoja['AQ59'] = fila["Lugar de residencia familiar.5"]
        
            actividad10 = fila['Remuneración.5']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT59'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU59'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV59'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW59'] = 'X'

            ##### Persona 7 #####

            if fila["Tipo de mano de obra.6"] == "Familiar":
                hoja['B60'] = 'X'
            elif fila["Tipo de mano de obra.6"] == "Contratado":
                hoja['D60'] = 'X'

            hoja['E60'] = fila["Cargo.7"]

            if fila["Género.6"] == "Masculino":
                hoja['M60'] = 'X'
            elif fila["Género.6"] == "Femenino":
                hoja['K60'] = 'X'

            hoja['N60'] = fila["Edad (años).6"]
            hoja['Q60'] = fila["Duración jornada (horas).6"]

            actividad9 = fila['Escolaridad.6']
            if actividad9 == 'Primaria':
                hoja['S60'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U60'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W60'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y60'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA60'] = 'X'

            if fila['Contrato.6'] == 'Tem.':
                hoja['AC60'] = 'X'
            elif fila['Contrato.6'] == 'Fij':
                hoja['AG60'] = 'X'
            
            if fila['Pago de seguridad.6'] == 'Si':
                hoja['AH60'] = 'X'
                hoja['AJ60'] = ''
            elif fila['Pago de seguridad.6'] == 'No':
                hoja['AH60'] = ''
                hoja['AJ60'] = 'X'

            hoja['AL60'] = fila["Procedencia.6"]
            hoja['AM60'] = fila["Residencia.6"]
            hoja['AN60'] = fila["Tiempo trabajado.6"]
            hoja['AO60'] = fila["# Personas núcleo familiar.6"]
            hoja['AP60'] = fila["Personas a cargo.6"]
            hoja['AQ60'] = fila["Lugar de residencia familiar.6"]
        
            actividad10 = fila['Remuneración.6']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT60'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU60'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV60'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW60'] = 'X'

            ##### Persona 8 #####

            if fila["Tipo de mano de obra.7"] == "Familiar":
                hoja['B61'] = 'X'
            elif fila["Tipo de mano de obra.7"] == "Contratado":
                hoja['D61'] = 'X'

            hoja['E61'] = fila["Cargo.8"]

            if fila["Género.7"] == "Masculino":
                hoja['M61'] = 'X'
            elif fila["Género.7"] == "Femenino":
                hoja['K61'] = 'X'

            hoja['N61'] = fila["Edad (años).7"]
            hoja['Q61'] = fila["Duración jornada (horas).7"]

            actividad9 = fila['Escolaridad.7']
            if actividad9 == 'Primaria':
                hoja['S61'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U61'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W61'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y61'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA61'] = 'X'

            if fila['Contrato.7'] == 'Tem.':
                hoja['AC61'] = 'X'
            elif fila['Contrato.7'] == 'Fij':
                hoja['AG61'] = 'X'
            
            if fila['Pago de seguridad.7'] == 'Si':
                hoja['AH61'] = 'X'
                hoja['AJ61'] = ''
            elif fila['Pago de seguridad.7'] == 'No':
                hoja['AH61'] = ''
                hoja['AJ61'] = 'X'

            hoja['AL61'] = fila["Procedencia.7"]
            hoja['AM61'] = fila["Residencia.7"]
            hoja['AN61'] = fila["Tiempo trabajado.7"]
            hoja['AO61'] = fila["# Personas núcleo familiar.7"]
            hoja['AP61'] = fila["Personas a cargo.7"]
            hoja['AQ61'] = fila["Lugar de residencia familiar.7"]
        
            actividad10 = fila['Remuneración.7']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT61'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU61'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV61'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW61'] = 'X'

            ##### Persona 9 #####

            if fila["Tipo de mano de obra.8"] == "Familiar":
                hoja['B62'] = 'X'
            elif fila["Tipo de mano de obra.8"] == "Contratado":
                hoja['D62'] = 'X'

            hoja['E62'] = fila["Cargo.9"]

            if fila["Género.8"] == "Masculino":
                hoja['M62'] = 'X'
            elif fila["Género.8"] == "Femenino":
                hoja['K62'] = 'X'

            hoja['N62'] = fila["Edad (años).8"]
            hoja['Q62'] = fila["Duración jornada (horas).8"]

            actividad9 = fila['Escolaridad.8']
            if actividad9 == 'Primaria':
                hoja['S62'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U62'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W62'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y62'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA62'] = 'X'

            if fila['Contrato.8'] == 'Tem.':
                hoja['AC62'] = 'X'
            elif fila['Contrato.8'] == 'Fij':
                hoja['AG62'] = 'X'
            
            if fila['Pago de seguridad.8'] == 'Si':
                hoja['AH62'] = 'X'
                hoja['AJ62'] = ''
            elif fila['Pago de seguridad.8'] == 'No':
                hoja['AH62'] = ''
                hoja['AJ62'] = 'X'

            hoja['AL62'] = fila["Procedencia.8"]
            hoja['AM62'] = fila["Residencia.8"]
            hoja['AN62'] = fila["Tiempo trabajado.8"]
            hoja['AO62'] = fila["# Personas núcleo familiar.8"]
            hoja['AP62'] = fila["Personas a cargo.8"]
            hoja['AQ62'] = fila["Lugar de residencia familiar.8"]
        
            actividad10 = fila['Remuneración.8']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT62'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU62'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV62'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW62'] = 'X'

            ##### Persona 10 #####

            if fila["Tipo de mano de obra.9"] == "Familiar":
                hoja['B63'] = 'X'
            elif fila["Tipo de mano de obra.9"] == "Contratado":
                hoja['D63'] = 'X'

            hoja['E63'] = fila["Cargo.90"]

            if fila["Género.9"] == "Masculino":
                hoja['M63'] = 'X'
            elif fila["Género.9"] == "Femenino":
                hoja['K63'] = 'X'

            hoja['N63'] = fila["Edad (años).9"]
            hoja['Q63'] = fila["Duración jornada (horas).9"]

            actividad9 = fila['Escolaridad.9']
            if actividad9 == 'Primaria':
                hoja['S63'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['U63'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['W63'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['Y63'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['AA63'] = 'X'

            if fila['Contrato.9'] == 'Tem.':
                hoja['AC63'] = 'X'
            elif fila['Contrato.9'] == 'Fij':
                hoja['AG63'] = 'X'
            
            if fila['Pago de seguridad.9'] == 'Si':
                hoja['AH63'] = 'X'
                hoja['AJ63'] = ''
            elif fila['Pago de seguridad.9'] == 'No':
                hoja['AH63'] = ''
                hoja['AJ63'] = 'X'

            hoja['AL63'] = fila["Procedencia.9"]
            hoja['AM63'] = fila["Residencia.9"]
            hoja['AN63'] = fila["Tiempo trabajado.9"]
            hoja['AO63'] = fila["# Personas núcleo familiar.9"]
            hoja['AP63'] = fila["Personas a cargo.9"]
            hoja['AQ63'] = fila["Lugar de residencia familiar.9"]
        
            actividad10 = fila['Remuneración.9']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AT63'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AU63'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AV63'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AW63'] = 'X'


        elif fila["Contrata algún tipo de mano de obra"] == "No":
            hoja['AF50'] = 'X'
