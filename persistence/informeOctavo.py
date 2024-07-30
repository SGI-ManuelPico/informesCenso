import pandas as pd
import os

class InformeOctavo:
    def __init__(self):
        pass

    def lecturaArchivoOctavo(self):
        """
        Lectura del tercer archivo del censo económico.
        """

        # Lectura de rutas
        rutaArchivoInicial = os.getcwd() + "\\censos\\Censo Económico Maute.xlsm"
        archivoInicial = pd.ExcelFile(rutaArchivoInicial)
        archivoInicial = archivoInicial.parse(sheet_name="FORMATO 8. AGROINDUSTRIA", header=None)

        # Ajustes preliminares al archivo inicial.
        archivoInicial = archivoInicial.drop(columns=[0,1,2,3]).transpose()
        archivoInicial.columns = archivoInicial.iloc[0].str.lstrip()
        archivoInicial.columns = archivoInicial.iloc[0].str.rstrip()
        archivoInicial = archivoInicial.drop(archivoInicial.index[0])
        archivoInicial.columns = pd.io.common.dedup_names(archivoInicial.columns, is_potential_multiindex=False)

        archivoInicial = archivoInicial.reset_index(drop = True)

        return archivoInicial
    
    def crearArchivoOctavo(self, hoja, fila):
        """
        Creación del tercer archivo del censo económico.
        """


        hoja['AO1'] = fila["Encuesta No."]

        if pd.notna(fila['Fecha']):
            fecha_str = str(fila['Fecha'])
            if '/' in fecha_str:
                hoja['AN2'] = fecha_str.split('/')[0]
                hoja['AQ2'] = fecha_str.split('/')[1]
                hoja['AU2'] = fecha_str.split('/')[2]
            elif '-' in fecha_str:
                hoja['AN2'] = fecha_str.split('-')[0]
                hoja['AQ2'] = fecha_str.split('-')[1]
                hoja['AU2'] = fecha_str.split('-')[2] ####### LLENAR FECHA EN ESPACIOS VACÍOS Y NO SOBRE EL SÍMBOLO.
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')

        hoja['AO3'] = fila["Encuestador"]
        hoja['F7'] = fila["Nombre"]
        hoja['Y7'] = fila["Empresa"]
        hoja['AO7'] = fila["Cargo"]

        if fila["¿Pertenece a alguna asociación?"] == 'Si':
            hoja['AB8'] = 'X'
        elif fila["¿Pertenece a alguna asociación?"] == 'No':
            hoja['AD8'] = 'X'

        hoja['AO8'] = fila["Otro, ¿Cuál?"]

        if fila["¿Tiene registro industrial y/o permiso ambiental?"] == 'Si':
            hoja['I12'] = 'X'
        elif fila["¿Pertenece a alguna asociación?"] == 'No':
            hoja['K12'] = 'X'
            hoja['G13'] = fila['¿Cuál?.1']


        actividad = fila['Tipo de Cultivo']
        if actividad == 'Caucho':
            hoja['J15'] = 'X'
        elif actividad == 'Palma africana':
            hoja['J16'] = 'X'
        elif actividad == 'Acacia':
            hoja['R15'] = 'X'
        elif actividad == 'Otros':
            hoja['P16'] = fila['Otro, ¿Cuál?']

        if fila["Vende principalmente en:"] =='Sitio':
            hoja['I28'] = 'X'
        if fila["Vende principalmente en:"] =='Vereda':
            hoja['I20'] = 'X'
        if fila["Vende principalmente en:"] =='Casco Urbano':
            hoja['I30'] = 'X'
        if fila["Vende principalmente en:"] =='Otros Municipios y/o Veredas':
            hoja['T28'] = 'X'
            hoja['P30'] = fila['Otro, ¿Cuáles?.1']

        hoja['B19'] = fila["¿Con cuántos empleados cuenta la planta?"]

        actividad2 = fila["Vende principalmente en:"]
        if actividad2 == 'Sitio':
            hoja['R22'] = 'X'
        elif actividad2 == 'Vereda':
            hoja['R23'] = 'X'
        elif actividad2 == 'Casco Urbano':
            hoja['R24'] = 'X'
        elif actividad2 == 'Otros Municipios y/o Veredas':
            hoja['R25'] = 'X'
            hoja['G26'] = fila["¿Cuáles?"]
        
        actividad3 = fila["La planta obtiene el producto de:"]
        if actividad3 == 'Plantaciones propias':
            hoja['AQ12'] = 'X'
        elif actividad2 == 'Venta del producto por parte de particulares':
            hoja['AQ13'] = 'X'
        elif actividad3 == 'Otro':
            hoja['AQ14'] = 'X'
            hoja['AD15'] = fila["Otro, ¿Cuál?.1"]

        hoja['AB19'] = fila["¿Cuál es el precio de venta?"]
        hoja['AM19'] = fila["Unidad"]

        if fila["Sobre la actividad, piensa: Continuidad"] == "Continuar con la actividad":
            hoja['AQ23'] = 'X'
        elif fila["Sobre la actividad, piensa: Continuidad"] == "Finalizar la actividad":
            hoja['AR23'] = 'X'
        if fila["Sobre la actividad, piensa: Producción"] == "Ampliar la producción":
            hoja['AQ24'] = 'X'
            hoja['AR25'] = 'X'
        elif fila["Sobre la actividad, piensa: Producción"] == "Permanecer con la misma producción":
            hoja['AQ25'] = 'X'
            hoja['AR24'] = 'X'

        if fila['Hidrocaarburos'] != "":
            hoja['R28'] = 'X'
            hoja['AD28'] = fila['Hidrocaarburos']
        elif fila['Plantas de procesamiento'] != "":
            hoja['R29'] = 'X'
            hoja['AD29'] = fila['Plantas de procesamiento']
        elif fila['Distrubuidores regionales'] != "":
            hoja['R30'] = 'X'
        elif fila['Otro, ¿Cuál?.2'] != "":
            hoja['AD30'] = fila['Otro, ¿Cuál?.2']

        hoja['AK33'] = fila[' ¿Cuál es el área total cultivada? (Ha)']

        if fila['Unidad.1'] == "m2":
            hoja['V34'] = 'X'
        elif fila['Unidad.1'] == "Ha":
            hoja['Z34'] = 'X'
        elif fila['Unidad.1'] == "Cosecha":
            hoja['AC34'] = 'X'

        hoja['AL34'] = fila['Costo aproximado de establecimiento']

        if fila['Unidad.2'] == "m2":
            hoja['V35'] = 'X'
        elif fila['Unidad.2'] == "Ha":
            hoja['Z35'] = 'X'
        elif fila['Unidad.2'] == "Cosecha":
            hoja['AC35'] = 'X'

        hoja['AL35'] = fila['Costo aproximado de mantenimiento']

        if fila['Unidad.3'] == "m2":
            hoja['V36'] = 'X'
        elif fila['Unidad.3'] == "Ha":
            hoja['Z36'] = 'X'
        elif fila['Unidad.3'] == "Cosecha":
            hoja['AC36'] = 'X'

        hoja['AL36'] = fila['Costo aproximado de cosecha']

        hoja['AK37'] = fila["""Duración de cada ciclo de producción
(Indicar la unidad)"""]

        if fila['Unidad.4'] == "Tn":
            hoja['V38'] = 'X'
        elif fila['Unidad.4'] == "Lt":
            hoja['Z38'] = 'X'
        elif fila['Unidad.4'] == "Gn":
            hoja['AD38'] = 'X'

        hoja['AK38'] = fila['Volumen de producción por Ha']

        if fila['Unidad.5'] == "Tb":
            hoja['V39'] = 'X'
        elif fila['Unidad.5'] == "Kg":
            hoja['Z39'] = 'X'
        elif fila['Unidad.5'] == "Carga":
            hoja['AD39'] = 'X'

        hoja['AL39'] = fila['Precio de venta del producto']

        fila['Costo aproximado de establecimiento']

        ##### ABASTECIMIENTO DE INSUMOS #####

        ## INSUMO 1 - AGUA
        hoja['N44'] = fila["Precio compra"]
        hoja['Y44'] = fila["Cantidad (m3)"]
        hoja['AG44'] = fila["Frecuencia de abastecimiento"]
        hoja['AO44'] = fila["Procedencia de los insumos"]

        ## INSUMO 2 - COMBUSTIBLE
        hoja['N45'] = fila["Precio compra.1"]
        hoja['Y45'] = fila["Cantidad (Gn)"]
        hoja['AG45'] = fila["Frecuencia de abastecimiento.1"]
        hoja['AO45'] = fila["Procedencia de los insumos.1"]

        ## INSUMO 3
        hoja['N46'] = fila["Precio compra.2"]
        hoja['Y46'] = fila["Cantidad"]
        hoja['AG46'] = fila["Frecuencia de abastecimiento.2"]
        hoja['AO46'] = fila["Procedencia de los insumos.2"]

        ## INSUMO 4
        hoja['N47'] = fila["Precio compra.2"]
        hoja['Y47'] = fila["Cantidad.1"]
        hoja['AG47'] = fila["Frecuencia de abastecimiento.2"]
        hoja['AO47'] = fila["Procedencia de los insumos.2"]


        actividad6 = fila['¿De dónde se abastece del recurso hídrico?']
        if actividad6 == 'Aljibe':
            hoja['W48'] = 'X'
        elif actividad6 == 'Acueducto Veredal':
            hoja['AE48'] = 'X'
        elif actividad6 == 'Otro':
            hoja['AL48'] = 'X'
            hoja['AQ48'] = fila['¿Cuál?.2']
    
        hoja['W49'] = fila["Forma de extracción"]

        actividad7 = fila['¿Qué tipo de energía utiliza?']
        if actividad7 == 'Energía Eléctrica':
            hoja['AA50'] = 'X'
        elif actividad7 == 'Energía Solar':
            hoja['AJ50'] = 'X'
        elif actividad7 == 'Otro':
            hoja['AP50'] = fila['¿Cuál?.3']

        if fila["¿Cuenta con servicio de alcantarillado?"] == "Si":
            hoja['AB40'] = 'X'
        elif fila["¿Cuenta con servicio de alcantarillado?"] == "No":
            hoja['AD40'] = 'X'
        hoja['AO40'] = fila['¿Cuál?']

        if fila['¿Cuenta con servicio de alcantarillado?'] == "Si":
            hoja['AB51'] = 'X'
        elif fila['¿Cuenta con servicio de alcantarillado?'] == "No":
            hoja['AD51'] = 'X'

        hoja['AO51'] = fila['¿Cuál?.5']
        hoja['Y52'] = fila["¿Cuál es el manejo de aguas residuales y solidos?"]
        hoja['Y53'] = fila["¿Cuál es el gasto aproximado de suministros en el proceso durante un mes?"]

        ##### ¿Qué tipo de equipos o maquinaría utiliza? #####

        ## EQUIPO 1
        hoja['B56'] = fila["Equipo/maquinaria 1"]
        hoja['N56'] = fila["Precio al que lo compró"]
        hoja['X56'] = fila["Cantidad que posee la unidad económica"]
        hoja['AF56'] = fila["Vida útil"]
        hoja['AO56'] = fila["Procedencia"]

        ## EQUIPO 2
        hoja['B57'] = fila["Equipo/maquinaria 2"]
        hoja['N57'] = fila["Precio al que lo compró.1"]
        hoja['X57'] = fila["Cantidad que posee la unidad económica.1"]
        hoja['AF57'] = fila["Vida útil.1"]
        hoja['AO57'] = fila["Procedencia.1"]

        ## EQUIPO 3
        hoja['B58'] = fila["Equipo/maquinaria 3"]
        hoja['N58'] = fila["Precio al que lo compró.1"]
        hoja['X58'] = fila["Cantidad que posee la unidad económica.1"]
        hoja['AF58'] = fila["Vida útil.1"]
        hoja['AO58'] = fila["Procedencia.1"]

        ## EQUIPO 4
        hoja['B59'] = fila["Equipo/maquinaria 4"]
        hoja['N59'] = fila["Precio al que lo compró.1"]
        hoja['X59'] = fila["Cantidad que posee la unidad económica.1"]
        hoja['AF59'] = fila["Vida útil.1"]
        hoja['AO59'] = fila["Procedencia.1"]

        ## EQUIPO 5
        hoja['B60'] = fila["Equipo/maquinaria 5"]
        hoja['N60'] = fila["Precio al que lo compró.1"]
        hoja['X60'] = fila["Cantidad que posee la unidad económica.1"]
        hoja['AF60'] = fila["Vida útil.1"]
        hoja['AO60'] = fila["Procedencia.1"]

        hoja['W49'] = fila["¿Cuál fue el monto total gastado en insumos del último mes?"]

        ##### INFORMACIÓN LABORAL #####

        #### MANO DE OBRA CALIFICADA ####
        hoja['I67'] = fila["#"]

        if fila["Género"] == "Femenino":
            hoja['K66'] = 'X'
        elif fila["Género"] == "Masculino":
            hoja['M66'] = 'X'

        if fila["Contrato"] == "Termino Fijo":
            hoja['O66'] = 'X'
        elif fila["Contrato"] == "Indefinido":
            hoja['S66'] = 'X'

        #hoja['XB53'] = fila["¿Cuánto?"]
        hoja['W66'] = fila["Jornal y turno laboral"]

        if fila["Escolaridad"] == "Primaria":
            hoja['AB66'] = 'X'
        elif fila["Escolaridad"] == "Bachillerato":
            hoja['AC66'] = 'X'
        elif fila["Escolaridad"] == "Técnico o tecnológico":
            hoja['AD66'] = 'X'
        elif fila["Escolaridad"] == "Profesional":
            hoja['AE66'] = 'X'
        elif fila["Escolaridad"] == "Posgrado":
            hoja['AG66'] = 'X'
        
        if fila["Procedencia.3"] == "Vereda":
            hoja['AJ66'] = 'X'
        elif fila["Procedencia.3"] == "Municipio":
            hoja['AM66'] = 'X'
        elif fila["Procedencia.3"] == "Otro":
            hoja['AO66'] = 'X'
        
        if fila["Residencia"] == "Vereda":
            hoja['AQ66'] = 'X'
        elif fila["Residencia"] == "Municipio":
            hoja['AS66'] = 'X'
        elif fila["Residencia"] == "Otro":
            hoja['AU66'] = 'X'
        

        #### MANO DE OBRA NO CALIFICADA ####
        hoja['I67'] = fila["#.1"]

        if fila["Género.1"] == "Femenino":
            hoja['K67'] = 'X'
        elif fila["Género.1"] == "Masculino":
            hoja['M67'] = 'X'

        if fila["Contrato.1"] == "Termino Fijo":
            hoja['O67'] = 'X'
        elif fila["Contrato.1"] == "Indefinido":
            hoja['S67'] = 'X'

        #hoja['XB53'] = fila["¿Cuánto?.1"]
        hoja['W67'] = fila["Jornal y turno laboral.1"]

        if fila["Escolaridad.1"] == "Primaria":
            hoja['AB67'] = 'X'
        elif fila["Escolaridad.1"] == "Bachillerato":
            hoja['AC67'] = 'X'
        elif fila["Escolaridad.1"] == "Técnico o tecnológico":
            hoja['AD67'] = 'X'
        elif fila["Escolaridad.1"] == "Profesional":
            hoja['AE67'] = 'X'
        elif fila["Escolaridad.1"] == "Posgrado":
            hoja['AG67'] = 'X'
        
        if fila["Procedencia.4"] == "Vereda":
            hoja['AJ67'] = 'X'
        elif fila["Procedencia.4"] == "Municipio":
            hoja['AM67'] = 'X'
        elif fila["Procedencia.4"] == "Otro":
            hoja['AO67'] = 'X'
        
        if fila["Residencia.1"] == "Vereda":
            hoja['AQ67'] = 'X'
        elif fila["Residencia.1"] == "Municipio":
            hoja['AS67'] = 'X'
        elif fila["Residencia.1"] == "Otro":
            hoja['AU67'] = 'X'
        

        #### EMPLEADOS ADMINISTRATIVOS Y CONTABLES ####
        hoja['I68'] = fila["#.2"]

        if fila["Género.2"] == "Femenino":
            hoja['K68'] = 'X'
        elif fila["Género.2"] == "Masculino":
            hoja['M68'] = 'X'

        if fila["Contrato.2"] == "Termino Fijo":
            hoja['O68'] = 'X'
        elif fila["Contrato.2"] == "Indefinido":
            hoja['S68'] = 'X'

        #hoja['XB53'] = fila["¿Cuánto?.2"]
        hoja['W68'] = fila["Jornal y turno laboral.2"]

        if fila["Escolaridad.2"] == "Primaria":
            hoja['AB68'] = 'X'
        elif fila["Escolaridad.2"] == "Bachillerato":
            hoja['AC68'] = 'X'
        elif fila["Escolaridad.2"] == "Técnico o tecnológico":
            hoja['AD68'] = 'X'
        elif fila["Escolaridad.2"] == "Profesional":
            hoja['AE68'] = 'X'
        elif fila["Escolaridad.2"] == "Posgrado":
            hoja['AG68'] = 'X'
        
        if fila["Procedencia.5"] == "Vereda":
            hoja['AJ68'] = 'X'
        elif fila["Procedencia.5"] == "Municipio":
            hoja['AM68'] = 'X'
        elif fila["Procedencia.5"] == "Otro":
            hoja['AO68'] = 'X'
        
        if fila["Residencia.2"] == "Vereda":
            hoja['AQ68'] = 'X'
        elif fila["Residencia.2"] == "Municipio":
            hoja['AS68'] = 'X'
        elif fila["Residencia.2"] == "Otro":
            hoja['AU68'] = 'X'
        
        #### GERENTES Y DIRECTIVOS ####
        hoja['I69'] = fila["#.3"]

        if fila["Género.3"] == "Femenino":
            hoja['K69'] = 'X'
        elif fila["Género.3"] == "Masculino":
            hoja['M69'] = 'X'
        
        if fila["Contrato.3"] == "Termino Fijo":
            hoja['O69'] = 'X'
        elif fila["Contrato.3"] == "Indefinido":
            hoja['S69'] = 'X'
        
        # hoja['XB53'] = fila["¿Cuánto?.3"]
        hoja['W69'] = fila["Jornal y turno laboral.3"]

        if fila["Escolaridad.3"] == "Primaria":
            hoja['AB69'] = 'X'
        elif fila["Escolaridad.3"] == "Bachillerato":
            hoja['AC69'] = 'X'
        elif fila["Escolaridad.3"] == "Técnico o tecnológico":
            hoja['AD69'] = 'X'
        elif fila["Escolaridad.3"] == "Profesional":
            hoja['AE69'] = 'X'
        elif fila["Escolaridad.3"] == "Posgrado":
            hoja['AG69'] = 'X'
        
        if fila["Procedencia.6"] == "Vereda":
            hoja['AJ69'] = 'X'
        elif fila["Procedencia.6"] == "Municipio":
            hoja['AM69'] = 'X'
        elif fila["Procedencia.6"] == "Otro":
            hoja['AO69'] = 'X'
        
        if fila["Residencia.4"] == "Vereda":
            hoja['AQ69'] = 'X'
        elif fila["Residencia.4"] == "Municipio":
            hoja['AS69'] = 'X'
        elif fila["Residencia.4"] == "Otro":
            hoja['AU69'] = 'X'
        
        ##### OBRA O LABOR #####
        ## Obra o labor 1
        hoja['A74'] = fila['Tipo de obra o labor 1']
        hoja['K74'] = fila['Frecuencia de contratación/año']
        hoja['R74'] = fila['Duración en Jornales del contrato']
        hoja['AA74'] = fila['Valor del jornal']
        hoja['AG74'] = fila['Cantidad de jornaleros empleados por contrato']
        hoja['AO74'] = fila['Residencia de los jornaleros']

        ## Obra o labor 2
        hoja['A75'] = fila['Tipo de obra o labor 2']
        hoja['K75'] = fila['Frecuencia de contratación/año.1']
        hoja['R75'] = fila['Duración en Jornales del contrato.1']
        hoja['AA75'] = fila['Valor del jornal.1']
        hoja['AG75'] = fila['Cantidad de jornaleros empleados por contrato.1']
        hoja['AO75'] = fila['Residencia de los jornaleros.1']

        ## Obra o labor 3
        hoja['A76'] = fila['Tipo de obra o labor 3']
        hoja['K76'] = fila['Frecuencia de contratación/año.2']
        hoja['R76'] = fila['Duración en Jornales del contrato.2']
        hoja['AA76'] = fila['Valor del jornal.2']
        hoja['AG76'] = fila['Cantidad de jornaleros empleados por contrato.2']
        hoja['AO76'] = fila['Residencia de los jornaleros.2']


        if fila['Contrata servicios profesionales * Sí (Responder 30 y 31)'] == "Si":
            hoja['L78'] = 'X'
        elif fila['Contrata servicios profesionales * Sí (Responder 30 y 31)'] == "No":
            hoja['N78'] = 'X'

        if fila['¿Qué tipo de servicios?'] == "Contaduría":
            hoja['AK78'] = 'X'
        elif fila['¿Qué tipo de servicios?'] == "Consultoría":
            hoja['AT78'] = 'X'
        elif fila['¿Qué tipo de servicios?'] == "Asesoría legal":
            hoja['AK79'] = 'X'
        elif fila['¿Qué tipo de servicios?'] == "Otros":
            hoja['AT79'] = 'X'
            hoja['AE80'] = fila['Otros, ¿Cuáles?.1']
        

        ##### SERVICIOS #####
        hoja['A81'] = fila['Servicio 1']
        if fila['Frecuencia'] == "Mensual":
            hoja['F81'] = 'X'
        elif fila['Frecuencia'] == "Semestral":
            hoja['J81'] = 'X'
        elif fila['Frecuencia'] == "Trimestral":
            hoja['P81'] = 'X'
        elif fila['Frecuencia'] == "Anual":
            hoja['U81'] = 'X'

        hoja['A82'] = fila['Servicio 2']
        if fila['Frecuencia.1'] == "Mensual":
            hoja['F82'] = 'X'
        elif fila['Frecuencia.1'] == "Semestral":
            hoja['J82'] = 'X'
        elif fila['Frecuencia.1'] == "Trimestral":
            hoja['P82'] = 'X'
        elif fila['Frecuencia.1'] == "Anual":
            hoja['U82'] = 'X'
        
        hoja['A83'] = fila['Servicio 3']
        if fila['Frecuencia.2'] == "Mensual":
            hoja['F83'] = 'X'
        elif fila['Frecuencia.2'] == "Semestral":
            hoja['J83'] = 'X'
        elif fila['Frecuencia.2'] == "Trimestral":
            hoja['P83'] = 'X'
        elif fila['Frecuencia.2'] == "Anual":
            hoja['U83'] = 'X'

        
        hoja['A84'] = fila['Servicio 4']
        if fila['Frecuencia.3'] == "Mensual":
            hoja['F84'] = 'X'
        elif fila['Frecuencia.3'] == "Semestral":
            hoja['J84'] = 'X'
        elif fila['Frecuencia.3'] == "Trimestral":
            hoja['P84'] = 'X'
        elif fila['Frecuencia.3'] == "Anual":
            hoja['U84'] = 'X'

        
        hoja['A85'] = fila['Servicio 5']
        if fila['Frecuencia.4'] == "Mensual":
            hoja['F85'] = 'X'
        elif fila['Frecuencia.4'] == "Semestral":
            hoja['J85'] = 'X'
        elif fila['Frecuencia.4'] == "Trimestral":
            hoja['P85'] = 'X'
        elif fila['Frecuencia.4'] == "Anual":
            hoja['U85'] = 'X'


        hoja['AE84'] = fila['¿Cuál es el monto pagado por estos servicios durante el último semestre?']
        hoja['Z87'] = fila['Salarios pagados a la mano de obra calificada']
        hoja['Z88'] = fila['Salarios pagados a la mano de obra no calificada']
        hoja['Z89'] = fila['Salarios pagados a empleados y administrativos']
        hoja['Z90'] = fila[' Salarios pagados a gerentes y directivos']
        hoja['Z91'] = fila[' Total remuneraciones']
