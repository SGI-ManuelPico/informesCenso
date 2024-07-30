import pandas as pd
import os, re

class InformeCuarto:
    def __init__(self):
        pass

    def lecturaArchivoCuarto(self):
        """
        Lectura del tercer archivo del censo económico.
        """

        # Lectura de rutas
        rutaArchivoInicial = os.getcwd() + "\\censos\\Censo Económico Maute.xlsm"
        archivoInicial = pd.ExcelFile(rutaArchivoInicial)
        archivoInicial = archivoInicial.parse(sheet_name="FORMATO 4. SERVICIOS", header=None)

        # Ajustes preliminares al archivo inicial.
        archivoInicial = archivoInicial.drop(columns=[0,1,2,3]).transpose()
        archivoInicial.columns = archivoInicial.iloc[0].str.strip()
        archivoInicial = archivoInicial.drop(archivoInicial.index[0])
        archivoInicial.columns = pd.io.common.dedup_names(archivoInicial.columns, is_potential_multiindex=False)

        archivoInicial = archivoInicial.reset_index(drop = True)

        return archivoInicial
    
    def crearArchivoCuarto(self, hoja, fila):
        """
        Creación del tercer archivo del censo económico.
        """


        hoja['AO1'] = fila["Encuesta No."]

        if pd.notna(fila['Fecha']):
            fecha_str = str(fila['Fecha'])
            if '/' in fecha_str:
                hoja['AN2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                hoja['AQ2'] = fecha_str.split('/')[1]
                hoja['AT2'] = fecha_str.split('/')[0]
            elif '-' in fecha_str:
                hoja['AN2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                hoja['AQ2'] = fecha_str.split('-')[1]
                hoja['AT2'] = fecha_str.split('-')[0] ####### LLENAR FECHA EN ESPACIOS VACÍOS Y NO SOBRE EL SÍMBOLO.
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')

        hoja['AN3'] = fila["Encuestador"]
        hoja['F7'] = fila["Nombre"]
        hoja['AB7'] = fila["Empresa"]
        hoja['AR7'] = fila["Cargo"]

        if fila["¿Pertenece a alguna asociación?"] == 'Si':
            hoja['AB8'] = 'X'
        elif fila["¿Pertenece a alguna asociación?"] == 'No':
            hoja['AD8'] = 'X'

        hoja['AO8'] = fila["Otro, ¿Cuál?"]

        actividad = fila['¿Qué tipo de servicios oferta?']
        if actividad == 'Restaurante, cafetería':
            hoja['P13'] = 'X'
        elif actividad == 'Bar y centro nocturno':
            hoja['P14'] = 'X'
        elif actividad == 'Servicios sexuales':
            hoja['P15'] = 'X'
        elif actividad == 'Estación de servicios (combustible, montallantas, servicios)':
            hoja['P17'] = 'X'
        elif actividad == 'Servicio de giros y/o financieros':
            hoja['P19'] = 'X'
        elif actividad == 'Hospedaje (diligenciar título E)':
            hoja['P22'] = 'X'

            ################
            #### Hoja E ####
            ################

            if fila['¿Qué tipo de hospedaje oferta?'] == 'Hotel':
                hoja['F60'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Motel':
                hoja['F61'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Apartahotel':
                hoja['F62'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Pensión':
                hoja['F63'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Cabaña':
                hoja['M60'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Finca ecoturística':
                hoja['M61'] = 'X'
            elif fila['¿Qué tipo de hospedaje oferta?'] == 'Otro':
                hoja['M62'] = 'X'
                hoja['J24'] = fila['Otro, ¿Cuál?.4']
            
            hoja['R60'] = fila["¿Qué capacidad de alojamiento tiene?"]
            hoja['AE60'] = fila["Semanalmente, ¿Cuántas personas en promedio hospeda?"]
            hoja['AO60'] = fila["¿Cuáles son los principales sitios de procedencia de los huéspedes?"]


        elif actividad == 'Educacion':
            hoja['P22'] = 'X'
        elif actividad == 'Otros':
            hoja['P23'] = 'X'
            hoja['J24'] = fila['Otro, ¿Cuáles?']

        if fila["Vende principalmente en:"] =='Sitio':
            hoja['I28'] = 'X'
        if fila["Vende principalmente en:"] =='Vereda':
            hoja['I29'] = 'X'
        if fila["Vende principalmente en:"] =='Casco Urbano':
            hoja['I30'] = 'X'
        if fila["Vende principalmente en:"] =='Otros Municipios y/o Veredas':
            hoja['T28'] = 'X'
            hoja['P30'] = fila['Otro, ¿Cuáles?.1']


        hoja['AB13'] = fila["Servicio 1"]
        hoja['AN13'] = fila["Precio"]
        hoja['AB14'] = fila["Servicio 2"]
        hoja['AN14'] = fila["Precio.1"]
        hoja['AB15'] = fila["Servicio 3"]
        hoja['AN151'] = fila["Precio.2"]
        hoja['AB16'] = fila["Servicio 4"]
        hoja['AN16'] = fila["Precio.3"]

        hoja['AC19'] = fila["Servicio 1.1"]
        hoja['AC20'] = fila["Servicio 2.1"]
        hoja['AO19'] = fila["Servicio 3"]
        hoja['AO20'] = fila["Servicio 4"]

        actividad2 = fila['Frecuencia con que vende los servicios:']
        if actividad2 == 'Diario':
            hoja['AI23'] = 'X'
        elif actividad2 == 'Semanal':
            hoja['AI25'] = 'X'
        elif actividad2 == 'Quincenal':
            hoja['AI24'] = 'X'
        elif actividad2 == 'Mensual':
            hoja['AR24'] = 'X'


        if str(fila['Hidrocarburos']) != "nan":
            hoja['AI28'] = 'X'
            hoja['AO28'] = fila['Hidrocarburos']
        elif str(fila['Vereda']) != "nan":
            hoja['AI29'] = 'X'
            hoja['AO29'] = fila['Vereda']
        elif str(fila['Finca/Propiet.']) != "nan":
            hoja['AI30'] = 'X'
            hoja['AO30'] = fila['Finca/Propiet.']

        if fila["Sobre la actividad, piensa: Continuidad"] == "Continuar con la actividad":
            hoja['L33'] = 'X'
            hoja['AG34'] = 'X'
        elif fila["Sobre la actividad, piensa: Continuidad"] == "Finalizar la actividad":
            hoja['N33'] = 'X'
            hoja['AE34'] = 'X'
        if fila["Sobre la actividad, piensa: Producción"] == "Ampliar la producción":
            hoja['L34'] = 'X'
            hoja['AG33'] = 'X'
        elif fila["Sobre la actividad, piensa: Producción"] == "Permanecer con la misma producción":
            hoja['N34'] = 'X'
            hoja['AE33'] = 'X'

        actividad6 = fila['¿De dónde se abastece del recurso hídrico?']
        if actividad6 == 'Aljibe':
            hoja['W36'] = 'X'
        elif actividad6 == 'Acueducto Veredal':
            hoja['AE36'] = 'X'
        elif actividad6 == 'Otro':
            hoja['AL36'] = 'X'
            hoja['AQ36'] = fila['Otro, ¿Cuál?.1']
    
        hoja['W37'] = fila["Forma de extracción"]
        hoja['AQ37'] = fila["Cantidad estimada (agregar m3)"]


        actividad7 = fila['¿Qué tipo de energía utiliza?']
        if actividad7 == 'Energía Eléctrica':
            hoja['AA38'] = 'X'
        elif actividad7 == 'Energía Solar':
            hoja['AJ38'] = 'X'
        elif actividad7 == 'Otro':
            hoja['AP38'] = fila['Otro, ¿Cuál?.2']

        actividad8 = fila['¿De dónde proviene la energía que utiliza para la cocción de alimentos?']
        if actividad8 == 'Energía elétrica':
            hoja['AA39'] = 'X'
        elif actividad8 == 'Leña':
            hoja['AF39'] = 'X'
        elif actividad8 == 'Gas':
            hoja['AL39'] = 'X'
        elif actividad8 == 'Otro':
            hoja['AQ39'] = fila['Otro, ¿Cuál?.3']

        if fila["¿Cuenta con servicio de alcantarillado?"] == "Si":
            hoja['AB40'] = 'X'
        elif fila["¿Cuenta con servicio de alcantarillado?"] == "No":
            hoja['AD40'] = 'X'
        hoja['AO40'] = fila['¿Cuál?']


        ##### ABASTECIMIENTO DE INSUMOS #####

        ## SERVICIO 1
        hoja['B44'] = fila["Servicio 1.2"]
        hoja['J44'] = fila["Insumo/Materia prima"]
        hoja['S44'] = fila["Precio compra"]
        hoja['AB44'] = fila["Cantidad"]
        hoja['AI44'] = fila["Frecuencia de compra"]
        hoja['AQ44'] = fila["Procedencia"]

        ## SERVICIO 2
        hoja['B45'] = fila["Servicio 2.2"]
        hoja['J45'] = fila["Insumo/Materia prima.1"]
        hoja['S45'] = fila["Precio compra.1"]
        hoja['AB45'] = fila["Cantidad.1"]
        hoja['AI45'] = fila["Frecuencia de compra.1"]
        hoja['AQ45'] = fila["Procedencia.1"]

        ## SERVICIO 3
        hoja['B46'] = fila["Servicio 3.1"]
        hoja['J46'] = fila["Insumo/Materia prima.2"]
        hoja['S46'] = fila["Precio compra.2"]
        hoja['AB46'] = fila["Cantidad.2"]
        hoja['AI46'] = fila["Frecuencia de compra.2"]
        hoja['AQ46'] = fila["Procedencia.2"]

        ## SERVICIO 4
        hoja['B47'] = fila["Servicio 4.1"]
        hoja['J47'] = fila["Insumo/Materia prima.3"]
        hoja['S47'] = fila["Precio compra.3"]
        hoja['AB47'] = fila["Cantidad.3"]
        hoja['AI47'] = fila["Frecuencia de compra.3"]
        hoja['AQ47'] = fila["Procedencia.3"]

        ## SERVICIO 5
        hoja['B48'] = fila["Servicio 5"]
        hoja['J48'] = fila["Insumo/Materia prima.4"]
        hoja['S48'] = fila["Precio compra.4"]
        hoja['AB48'] = fila["Cantidad.4"]
        hoja['AI48'] = fila["Frecuencia de compra.4"]
        hoja['AQ48'] = fila["Procedencia.4"]

        hoja['W49'] = fila["¿Cuál fue el monto total gastado en insumos del último mes?"]

        ##### EQUIPOS Y MAQUINARIA #####

        ## EQUIPO 1
        hoja['B53'] = fila["Equipo/maquinaria"]
        hoja['N53'] = fila["Precio compra"]
        hoja['X53'] = fila["Cantidad que posee la unidad económica"]
        hoja['AF53'] = fila["Vida útil"]
        hoja['AO53'] = fila["Procedencia.5"]

        ## EQUIPO 2
        hoja['B54'] = fila["Equipo/maquinaria.1"]
        hoja['N54'] = fila["Precio compra.1"]
        hoja['X54'] = fila["Cantidad que posee la unidad económica.1"]
        hoja['AF54'] = fila["Vida útil.1"]
        hoja['AO54'] = fila["Procedencia.6"]

        ## EQUIPO 3
        hoja['B55'] = fila["Equipo/maquinaria.2"]
        hoja['N55'] = fila["Precio compra.2"]
        hoja['X55'] = fila["Cantidad que posee la unidad económica.2"]
        hoja['AF55'] = fila["Vida útil.2"]
        hoja['AO55'] = fila["Procedencia.7"]

        ## EQUIPO 4
        hoja['B56'] = fila["Equipo/maquinaria.3"]
        hoja['N56'] = fila["Precio compra.3"]
        hoja['X56'] = fila["Cantidad que posee la unidad económica.3"]
        hoja['AF56'] = fila["Vida útil.3"]
        hoja['AO56'] = fila["Procedencia.8"]

        if fila["Contrata algún tipo de mano de obra"] == "Si":
            hoja['AC64'] = 'X'

            #### Persona 1 ####

            if fila["Tipo de mano de obra"] == "Familiar":
                hoja['B69'] = 'X'
            elif fila["Tipo de mano de obra"] == "Contratado":
                hoja['D69'] = 'X'

            hoja['E69'] = fila["Cargo.1"]

            if fila["Género"] == "Masculino":
                hoja['J69'] = 'X'
            elif fila["Género"] == "Femenino":
                hoja['H69'] = 'X'

            hoja['K69'] = fila["Edad (años)"]
            hoja['L69'] = fila["Duración jornada (horas)"]

            actividad9 = fila['Escolaridad']
            if actividad9 == 'Primaria':
                hoja['N69'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q69'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S69'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U69'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W69'] = 'X'

            if fila['Contrato'] == 'Tem.':
                hoja['Y69'] = 'X'
            elif fila['Contrato'] == 'Fij':
                hoja['AC69'] = 'X'
            
            if fila['Pago de seguridad'] == 'Si':
                hoja['AE69'] = 'X'
            elif fila['Pago de seguridad'] == 'No':
                hoja['AG69'] = 'X'

            hoja['AH69'] = fila["Procedencia.9"]
            hoja['AI69'] = fila["Residencia"]
            hoja['AL69'] = fila["Tiempo trabajado"]
            hoja['AM69'] = fila["# Personas núcleo familiar"]
            hoja['AO69'] = fila["Personas a cargo"]
            hoja['AP69'] = fila["Lugar de residencia familiar"]
        
            actividad10 = fila['Remuneración']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR69'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS69'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT69'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU69'] = 'X'

            ##### Persona 2 #####

            if fila["Tipo de mano de obra.1"] == "Familiar":
                hoja['B70'] = 'X'
            elif fila["Tipo de mano de obra.1"] == "Contratado":
                hoja['D70'] = 'X'

            hoja['E70'] = fila["Cargo.2"]

            if fila["Género.1"] == "Masculino":
                hoja['J70'] = 'X'
            elif fila["Género.1"] == "Femenino":
                hoja['H70'] = 'X'

            hoja['K70'] = fila["Edad (años).1"]
            hoja['L70'] = fila["Duración jornada (horas).1"]

            actividad9 = fila['Escolaridad.1']
            if actividad9 == 'Primaria':
                hoja['N70'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q70'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S70'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U70'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W70'] = 'X'

            if fila['Contrato.1'] == 'Tem.':
                hoja['Y70'] = 'X'
            elif fila['Contrato.1'] == 'Fij':
                hoja['AC70'] = 'X'
            
            if fila['Pago de seguridad.1'] == 'Si':
                hoja['AE70'] = 'X'
            elif fila['Pago de seguridad.1'] == 'No':
                hoja['Ag70'] = 'X'

            hoja['AH70'] = fila["Procedencia.10"]
            hoja['AI70'] = fila["Residencia.1"]
            hoja['AL70'] = fila["Tiempo trabajado.1"]
            hoja['AM70'] = fila["# Personas núcleo familiar.1"]
            hoja['AO70'] = fila["Personas a cargo.1"]
            hoja['AP70'] = fila["Lugar de residencia familiar.1"]
        
            actividad10 = fila['Remuneración.1']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR70'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS70'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT70'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU70'] = 'X'

            ##### Persona 3 #####

            if fila["Tipo de mano de obra.2"] == "Familiar":
                hoja['B71'] = 'X'
            elif fila["Tipo de mano de obra.2"] == "Contratado":
                hoja['D71'] = 'X'

            hoja['E71'] = fila["Cargo.3"]

            if fila["Género.2"] == "Masculino":
                hoja['J71'] = 'X'
            elif fila["Género.2"] == "Femenino":
                hoja['H71'] = 'X'

            hoja['K71'] = fila["Edad (años).2"]
            hoja['L71'] = fila["Duración jornada (horas).2"]

            actividad9 = fila['Escolaridad.2']
            if actividad9 == 'Primaria':
                hoja['N71'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q71'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S71'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U71'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W71'] = 'X'

            if fila['Contrato.2'] == 'Tem.':
                hoja['Y71'] = 'X'
            elif fila['Contrato.2'] == 'Fij':
                hoja['AC71'] = 'X'
            
            if fila['Pago de seguridad.2'] == 'Si':
                hoja['AE71'] = 'X'
            elif fila['Pago de seguridad.2'] == 'No':
                hoja['AG71'] = 'X'

            hoja['AH71'] = fila["Procedencia.11"]
            hoja['AI71'] = fila["Residencia.2"]
            hoja['AL71'] = fila["Tiempo trabajado.2"]
            hoja['AM71'] = fila["# Personas núcleo familiar.2"]
            hoja['AO71'] = fila["Personas a cargo.2"]
            hoja['AP71'] = fila["Lugar de residencia familiar.2"]
        
            actividad10 = fila['Remuneración.2']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR71'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS71'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT71'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU71'] = 'X'

            ##### Persona 4 #####

            if fila["Tipo de mano de obra.3"] == "Familiar":
                hoja['B72'] = 'X'
            elif fila["Tipo de mano de obra.3"] == "Contratado":
                hoja['D72'] = 'X'

            hoja['E72'] = fila["Cargo.4"]

            if fila["Género.3"] == "Masculino":
                hoja['J72'] = 'X'
            elif fila["Género.3"] == "Femenino":
                hoja['H72'] = 'X'

            hoja['K72'] = fila["Edad (años).3"]
            hoja['L72'] = fila["Duración jornada (horas).3"]

            actividad9 = fila['Escolaridad.3']
            if actividad9 == 'Primaria':
                hoja['N72'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q72'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S72'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U72'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W72'] = 'X'

            if fila['Contrato.3'] == 'Tem.':
                hoja['Y72'] = 'X'
            elif fila['Contrato.3'] == 'Fij':
                hoja['AC72'] = 'X'
            
            if fila['Pago de seguridad.3'] == 'Si':
                hoja['AE72'] = 'X'
            elif fila['Pago de seguridad.3'] == 'No':
                hoja['AG72'] = 'X'

            hoja['AH72'] = fila["Procedencia.12"]
            hoja['AI72'] = fila["Residencia.3"]
            hoja['AL72'] = fila["Tiempo trabajado.3"]
            hoja['AM72'] = fila["# Personas núcleo familiar.3"]
            hoja['AO72'] = fila["Personas a cargo.3"]
            hoja['AP72'] = fila["Lugar de residencia familiar.3"]
        
            actividad10 = fila['Remuneración.3']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR72'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS72'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT72'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU72'] = 'X'

            ##### Persona 5 #####

            if fila["Tipo de mano de obra.4"] == "Familiar":
                hoja['B73'] = 'X'
            elif fila["Tipo de mano de obra.4"] == "Contratado":
                hoja['D73'] = 'X'

            hoja['E73'] = fila["Cargo.5"]

            if fila["Género.4"] == "Masculino":
                hoja['J73'] = 'X'
            elif fila["Género.4"] == "Femenino":
                hoja['H73'] = 'X'

            hoja['K73'] = fila["Edad (años).4"]
            hoja['L73'] = fila["Duración jornada (horas).4"]

            actividad9 = fila['Escolaridad.4']
            if actividad9 == 'Primaria':
                hoja['N73'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q73'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S73'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U73'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W73'] = 'X'

            if fila['Contrato.4'] == 'Tem.':
                hoja['Y73'] = 'X'
            elif fila['Contrato.4'] == 'Fij':
                hoja['AC73'] = 'X'
            
            if fila['Pago de seguridad.4'] == 'Si':
                hoja['AE73'] = 'X'
            elif fila['Pago de seguridad.4'] == 'No':
                hoja['AG73'] = 'X'

            hoja['AH73'] = fila["Procedencia.13"]
            hoja['AI73'] = fila["Residencia.4"]
            hoja['AL73'] = fila["Tiempo trabajado.4"]
            hoja['AM73'] = fila["# Personas núcleo familiar.4"]
            hoja['AO73'] = fila["Personas a cargo.4"]
            hoja['AP73'] = fila["Lugar de residencia familiar.4"]
        
            actividad10 = fila['Remuneración.4']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR73'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS73'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT73'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU73'] = 'X'

            ##### Persona 6 #####

            if fila["Tipo de mano de obra.5"] == "Familiar":
                hoja['B74'] = 'X'
            elif fila["Tipo de mano de obra.5"] == "Contratado":
                hoja['D74'] = 'X'

            hoja['E74'] = fila["Cargo.6"]

            if fila["Género.5"] == "Masculino":
                hoja['j74'] = 'X'
            elif fila["Género.5"] == "Femenino":
                hoja['H74'] = 'X'

            hoja['K74'] = fila["Edad (años).5"]
            hoja['L74'] = fila["Duración jornada (horas).5"]

            actividad9 = fila['Escolaridad.5']
            if actividad9 == 'Primaria':
                hoja['N74'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q74'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S74'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U74'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W74'] = 'X'

            if fila['Contrato.5'] == 'Tem.':
                hoja['Y74'] = 'X'
            elif fila['Contrato.5'] == 'Fij':
                hoja['AC74'] = 'X'
            
            if fila['Pago de seguridad.5'] == 'Si':
                hoja['AE74'] = 'X'
            elif fila['Pago de seguridad.5'] == 'No':
                hoja['AG74'] = 'X'

            hoja['AH74'] = fila["Procedencia.14"]
            hoja['AI74'] = fila["Residencia.5"]
            hoja['AL74'] = fila["Tiempo trabajado.5"]
            hoja['AM74'] = fila["# Personas núcleo familiar.5"]
            hoja['AO74'] = fila["Personas a cargo.5"]
            hoja['AP74'] = fila["Lugar de residencia familiar.5"]
        
            actividad10 = fila['Remuneración.5']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR74'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS74'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT74'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU74'] = 'X'

            ##### Persona 7 #####

            if fila["Tipo de mano de obra.6"] == "Familiar":
                hoja['B75'] = 'X'
            elif fila["Tipo de mano de obra.6"] == "Contratado":
                hoja['D75'] = 'X'

            hoja['E75'] = fila["Cargo.7"]

            if fila["Género.6"] == "Masculino":
                hoja['J75'] = 'X'
            elif fila["Género.6"] == "Femenino":
                hoja['H75'] = 'X'

            hoja['K75'] = fila["Edad (años).6"]
            hoja['L75'] = fila["Duración jornada (horas).6"]

            actividad9 = fila['Escolaridad.6']
            if actividad9 == 'Primaria':
                hoja['N75'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q75'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S75'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U75'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W75'] = 'X'

            if fila['Contrato.6'] == 'Tem.':
                hoja['Y75'] = 'X'
            elif fila['Contrato.6'] == 'Fij':
                hoja['AC75'] = 'X'
            
            if fila['Pago de seguridad.6'] == 'Si':
                hoja['AE75'] = 'X'
            elif fila['Pago de seguridad.6'] == 'No':
                hoja['AG75'] = 'X'

            hoja['AH75'] = fila["Procedencia.15"]
            hoja['AI75'] = fila["Residencia.6"]
            hoja['AL75'] = fila["Tiempo trabajado.6"]
            hoja['AM75'] = fila["# Personas núcleo familiar.6"]
            hoja['AO75'] = fila["Personas a cargo.6"]
            hoja['AP75'] = fila["Lugar de residencia familiar.6"]
        
            actividad10 = fila['Remuneración.6']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR75'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS75'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT75'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU75'] = 'X'

            ##### Persona 8 #####

            if fila["Tipo de mano de obra.7"] == "Familiar":
                hoja['B76'] = 'X'
            elif fila["Tipo de mano de obra.7"] == "Contratado":
                hoja['D76'] = 'X'

            hoja['E76'] = fila["Cargo.8"]

            if fila["Género.7"] == "Masculino":
                hoja['J76'] = 'X'
            elif fila["Género.7"] == "Femenino":
                hoja['H76'] = 'X'

            hoja['K76'] = fila["Edad (años).7"]
            hoja['L76'] = fila["Duración jornada (horas).7"]

            actividad9 = fila['Escolaridad.7']
            if actividad9 == 'Primaria':
                hoja['N76'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q76'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S76'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U76'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W76'] = 'X'

            if fila['Contrato.7'] == 'Tem.':
                hoja['Y76'] = 'X'
            elif fila['Contrato.7'] == 'Fij':
                hoja['AC76'] = 'X'
            
            if fila['Pago de seguridad.7'] == 'Si':
                hoja['AE76'] = 'X'
            elif fila['Pago de seguridad.7'] == 'No':
                hoja['AG76'] = 'X'

            hoja['AH76'] = fila["Procedencia.16"]
            hoja['AI76'] = fila["Residencia.7"]
            hoja['AL76'] = fila["Tiempo trabajado.7"]
            hoja['AM76'] = fila["# Personas núcleo familiar.7"]
            hoja['AO76'] = fila["Personas a cargo.7"]
            hoja['AP76'] = fila["Lugar de residencia familiar.7"]
        
            actividad10 = fila['Remuneración.7']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR76'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS76'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT76'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU76'] = 'X'

            ##### Persona 9 #####

            if fila["Tipo de mano de obra.8"] == "Familiar":
                hoja['B77'] = 'X'
            elif fila["Tipo de mano de obra.8"] == "Contratado":
                hoja['D77'] = 'X'

            hoja['E77'] = fila["Cargo.9"]

            if fila["Género.8"] == "Masculino":
                hoja['J77'] = 'X'
            elif fila["Género.8"] == "Femenino":
                hoja['H77'] = 'X'

            hoja['K77'] = fila["Edad (años).8"]
            hoja['L77'] = fila["Duración jornada (horas).8"]

            actividad9 = fila['Escolaridad.8']
            if actividad9 == 'Primaria':
                hoja['N77'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q77'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S77'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U77'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W77'] = 'X'

            if fila['Contrato.8'] == 'Tem.':
                hoja['Y77'] = 'X'
            elif fila['Contrato.8'] == 'Fij':
                hoja['AC77'] = 'X'
            
            if fila['Pago de seguridad.8'] == 'Si':
                hoja['AE77'] = 'X'
            elif fila['Pago de seguridad.8'] == 'No':
                hoja['AG77'] = 'X'

            hoja['AH77'] = fila["Procedencia.17"]
            hoja['AI77'] = fila["Residencia.8"]
            hoja['AL77'] = fila["Tiempo trabajado.8"]
            hoja['AM77'] = fila["# Personas núcleo familiar.8"]
            hoja['AO77'] = fila["Personas a cargo.8"]
            hoja['AP77'] = fila["Lugar de residencia familiar.8"]
        
            actividad10 = fila['Remuneración.8']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR77'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS77'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT77'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU77'] = 'X'

            ##### Persona 10 #####

            if fila["Tipo de mano de obra.9"] == "Familiar":
                hoja['B78'] = 'X'
            elif fila["Tipo de mano de obra.9"] == "Contratado":
                hoja['D78'] = 'X'

            hoja['E78'] = fila["Cargo.9"]

            if fila["Género.9"] == "Masculino":
                hoja['J78'] = 'X'
            elif fila["Género.9"] == "Femenino":
                hoja['H78'] = 'X'

            hoja['K78'] = fila["Edad (años).9"]
            hoja['L78'] = fila["Duración jornada (horas).9"]

            actividad9 = fila['Escolaridad.9']
            if actividad9 == 'Primaria':
                hoja['N78'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q78'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S78'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U78'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W78'] = 'X'

            if fila['Contrato.9'] == 'Tem.':
                hoja['Y78'] = 'X'
            elif fila['Contrato.9'] == 'Fij':
                hoja['AC78'] = 'X'
            
            if fila['Pago de seguridad.9'] == 'Si':
                hoja['AE78'] = 'X'
            elif fila['Pago de seguridad.9'] == 'No':
                hoja['AG78'] = 'X'

            hoja['AH78'] = fila["Procedencia.18"]
            hoja['AI78'] = fila["Residencia.9"]
            hoja['AL78'] = fila["Tiempo trabajado.9"]
            hoja['AM78'] = fila["# Personas núcleo familiar.9"]
            hoja['AO78'] = fila["Personas a cargo.9"]
            hoja['AP78'] = fila["Lugar de residencia familiar.9"]
        
            actividad10 = fila['Remuneración.9']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AR78'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AS78'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AT78'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AU78'] = 'X'


        elif fila["Contrata algún tipo de mano de obra"] == "No":
            hoja['AE64'] = 'X'

