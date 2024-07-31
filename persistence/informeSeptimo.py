import pandas as pd
import os, re

class InformeSeptimo:
    def __init__(self):
        pass

    def lecturaArchivoSeptimo(self):
        """
        Lectura del tercer archivo del censo económico.
        """

        # Lectura de rutas
        rutaArchivoInicial = os.getcwd() + "\\Censo Económico Maute.xlsm"
        archivoInicial = pd.ExcelFile(rutaArchivoInicial)
        archivoInicial = archivoInicial.parse(sheet_name="FORMATO 7. TRANSPORTE", header=None)

        # Ajustes preliminares al archivo inicial.
        archivoInicial = archivoInicial.drop(columns=[0,1,2,3]).transpose()
        archivoInicial.columns = archivoInicial.iloc[0].str.strip()
        archivoInicial = archivoInicial.drop(archivoInicial.index[0])
        archivoInicial.columns = pd.io.common.dedup_names(archivoInicial.columns, is_potential_multiindex=False)
        archivoInicial = archivoInicial.reset_index(drop = True)

        return archivoInicial
    
    def crearArchivoSeptimo(self, hoja, fila):
        """
        Creación del tercer archivo del censo económico.
        """


        hoja['AQ1'] = fila["Encuesta No."]

        if pd.notna(fila['Fecha']):
            fecha_str = str(fila['Fecha'])
            if '/' in fecha_str:
                hoja['AO2'] = re.findall('\d+',fecha_str.split("/")[2])[0]
                hoja['AR2'] = fecha_str.split('/')[1]
                hoja['AU2'] = fecha_str.split('/')[0]
            elif '-' in fecha_str:
                hoja['AO2'] = re.findall('\d+',fecha_str.split("-")[2])[0]
                hoja['AR2'] = fecha_str.split('-')[1]
                hoja['AU2'] = fecha_str.split('-')[0] ####### LLENAR FECHA EN ESPACIOS VACÍOS Y NO SOBRE EL SÍMBOLO.
            else:
                print(f'Formato de fecha inesperado: {fecha_str}')
        else:
            print('Campo de fecha vacío')

        hoja['AP3'] = fila["Encuestador"]
        hoja['F7'] = fila["Nombre"]
        hoja['Z7'] = fila["Empresa"]
        hoja['AQ7'] = fila["Cargo"]

        if fila["¿Pertenece a alguna asociación?"] == 'Si':
            hoja['AA8'] = 'X'

        elif fila["¿Pertenece a alguna asociación?"] == 'No':
            hoja['AC8'] = 'X'

        hoja['AO8'] = fila["Otro, ¿Cuál?"]

        actividad = fila['¿Qué tipo de servicio de transporte ofrece?']
        if actividad == 'Transporte público':
            hoja['T12'] = 'X'
        elif actividad == 'Servicios especiales o transporte de pasajeros':
            hoja['T13'] = 'X'
        elif actividad == 'Transporte de carga o maquinaria':
            hoja['T14'] = 'X'


        actividad2 = fila['Presta los servicios de transporte como']
        if actividad2 == 'Particular':
            hoja['F16'] = 'X'
        elif actividad2 == 'Afiliado':
            hoja['F17'] = 'X'
        elif actividad2 == 'Cooperativa':
            hoja['M16'] = actividad2
        elif actividad2 == 'Empresa':
            hoja['M17'] = 'X'
        elif actividad2 == 'Otro':
            hoja['N17'] = fila["Otro, ¿Cuál?"]


        if fila["¿Presta los servicios al sector de hidrocarburos?"] == 'Si':
            hoja['AM12'] = 'X'

        elif fila["¿Presta los servicios al sector de hidrocarburos?"] == 'No':
            hoja['AO12'] = 'X'

        if fila["¿El pago por parte del prestador de hidrocarburos es oportuno?"] == 'Si':
            hoja['AM15'] = 'X'

        elif fila["¿El pago por parte del prestador de hidrocarburos es oportuno?"] == 'No':
            hoja['AO15'] = 'X'

        hoja['AI16'] = fila['Observaciones']


        if fila["Sobre la actividad, piensa: Continuidad"] == "Continuar con la actividad":
            hoja['L19'] = 'X'
            hoja['AU20'] = 'X'
        elif fila["Sobre la actividad, piensa: Continuidad"] == "Finalizar la actividad":
            hoja['N19'] = 'X'
            hoja['AS20'] = 'X'
        if fila["Sobre la actividad, piensa: Producción"] == "Ampliar la producción":
            hoja['AB19'] = 'X'
            hoja['AU19'] = 'X'
        elif fila["Sobre la actividad, piensa: Producción"] == "Permanecer con la misma producción":
            hoja['AD19'] = 'X'
            hoja['AS19'] = 'X'
        elif fila["Sobre la actividad, piensa: Producción"] == "Ninguna de las anteriores":
            hoja['AD19'] = 'X'
            hoja['AU19'] = 'X'

        if fila["Se encuentra afiliado a alguna empresa y/o cooperativa de transporte:"] == 'Si':
            hoja['H24'] = 'X'

        elif fila["Se encuentra afiliado a alguna empresa y/o cooperativa de transporte:"] == 'No':
            hoja['J24'] = 'X'

        hoja['D25'] = fila['¿Cuál?.1']
        hoja['R25'] = fila['Número de Contacto']

        if fila["¿Es propietario del vehículo?"] == 'Si':
            hoja['H28'] = 'X'

        elif fila["¿Es propietario del vehículo?"] == 'No':
            hoja['J28'] = 'X'

        hoja['A31'] = fila['¿Cuál es el porcentaje pagado a la cooperativa por cada servicio? (%)']
        hoja['A33'] = fila['Durante la última semana cuántos Km recorrió']
        hoja['AP22'] = fila['Hace cuánto presta servicios de transporte']
        
        if fila['¿Cuál es el costo aproximado de mantenimiento y gastos del vehículo en una semana?'] == "Entre $100.000 - $200.000":
            hoja['AS26'] = 'X'
        elif fila['¿Cuál es el costo aproximado de mantenimiento y gastos del vehículo en una semana?'] == "Entre $201.000 - $400.000":
            hoja['AS27'] = 'X'
        elif fila['¿Cuál es el costo aproximado de mantenimiento y gastos del vehículo en una semana?'] == "Entre $401.000 - $600.000":
            hoja['AS28'] = 'X'
        elif fila['¿Cuál es el costo aproximado de mantenimiento y gastos del vehículo en una semana?'] == "Mayor a $600.000":
            hoja['AS29'] = 'X'

        if fila["¿El estado de las vías le genera sobre costos?"] == 'Si':
            hoja['AN32'] = 'X'

        elif fila["¿El estado de las vías le genera sobre costos?"] == 'No':
            hoja['AP32'] = 'X'

        hoja['AL33'] = fila['Costos Incurridos']
        hoja['C37'] = fila['Cuántos afiliados tiene la cooperativa y/o empresa de transporte']
        hoja['D40'] = fila['Costo 1']
        hoja['Q40'] = fila['Valor']
        hoja['D41'] = fila['Costo 2']
        hoja['Q41'] = fila['Valor.1']
        hoja['D42'] = fila['Costo 3']
        hoja['Q42'] = fila['Valor.2']
        hoja['AC37'] = fila['¿Cuál es el porcentaje cobrado a los afiliados por cada servicio prestado? (%)']

        if fila["Emplea directamente algún tipo de mano de obra (si la respuesta es SI, diligenciar el título G)"] == 'Si':
            hoja['AN40'] = 'X'
            #### Persona 1 ####
            if fila["Tipo de mano de obra"] == "Familiar":
                hoja['B64'] = 'X'
            elif fila["Tipo de mano de obra"] == "Contratado":
                hoja['D64'] = 'X'

            hoja['E64'] = fila["Cargo.1"]

            if fila["Género"] == "Masculino":
                hoja['J64'] = 'X'
            elif fila["Género"] == "Femenino":
                hoja['H64'] = 'X'

            hoja['K64'] = fila["Edad (años)"]
            hoja['L64'] = fila["Duración jornada (horas)"]

            actividad9 = fila['Escolaridad']
            if actividad9 == 'Primaria':
                hoja['N64'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q64'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S64'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U64'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W64'] = 'X'

            if fila['Contrato'] == 'Tem.':
                hoja['AC64'] = 'X'
            elif fila['Contrato'] == 'Fij':
                hoja['AE64'] = 'X'
            
            if fila['Pago de seguridad'] == 'Si':
                hoja['AG64'] = 'X'
            elif fila['Pago de seguridad'] == 'No':
                hoja['AI64'] = 'X'

            hoja['AJ64'] = fila["Procedencia.4"]
            hoja['AK64'] = fila["Residencia"]
            hoja['AN64'] = fila["Tiempo trabajado"]
            hoja['AO64'] = fila["# Personas núcleo familiar"]
            hoja['AQ64'] = fila["Personas a cargo"]
            hoja['AS64'] = fila["Lugar de residencia familiar"]
        
            actividad10 = fila['Remuneración']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU64'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV64'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW64'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX64'] = 'X'

            ##### Persona 2 #####

            if fila["Tipo de mano de obra.1"] == "Familiar":
                hoja['B65'] = 'X'
            elif fila["Tipo de mano de obra.1"] == "Contratado":
                hoja['D65'] = 'X'

            hoja['E65'] = fila["Cargo.2"]

            if fila["Género.1"] == "Masculino":
                hoja['J65'] = 'X'
            elif fila["Género.1"] == "Femenino":
                hoja['H65'] = 'X'

            hoja['K65'] = fila["Edad (años).1"]
            hoja['L65'] = fila["Duración jornada (horas).1"]

            actividad9 = fila['Escolaridad.1']
            if actividad9 == 'Primaria':
                hoja['N65'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q65'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S65'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U65'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W65'] = 'X'

            if fila['Contrato.1'] == 'Tem.':
                hoja['AC65'] = 'X'
            elif fila['Contrato.1'] == 'Fij':
                hoja['AE65'] = 'X'
            
            if fila['Pago de seguridad.1'] == 'Si':
                hoja['AG65'] = 'X'
            elif fila['Pago de seguridad.1'] == 'No':
                hoja['AI65'] = 'X'

            hoja['AJ65'] = fila["Procedencia.5"]
            hoja['AK65'] = fila["Residencia.1"]
            hoja['AN65'] = fila["Tiempo trabajado.1"]
            hoja['AO65'] = fila["# Personas núcleo familiar.1"]
            hoja['AQ65'] = fila["Personas a cargo.1"]
            hoja['AS65'] = fila["Lugar de residencia familiar.1"]
        
            actividad10 = fila['Remuneración.1']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU65'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV65'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW65'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX65'] = 'X'

            ##### Persona 3 #####

            if fila["Tipo de mano de obra.2"] == "Familiar":
                hoja['B66'] = 'X'
            elif fila["Tipo de mano de obra.2"] == "Contratado":
                hoja['D66'] = 'X'

            hoja['E66'] = fila["Cargo.3"]

            if fila["Género.2"] == "Masculino":
                hoja['J66'] = 'X'
            elif fila["Género.2"] == "Femenino":
                hoja['H66'] = 'X'

            hoja['K66'] = fila["Edad (años).2"]
            hoja['L66'] = fila["Duración jornada (horas).2"]

            actividad9 = fila['Escolaridad.2']
            if actividad9 == 'Primaria':
                hoja['N66'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q66'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S66'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U66'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W66'] = 'X'

            if fila['Contrato.2'] == 'Tem.':
                hoja['AC66'] = 'X'
            elif fila['Contrato.2'] == 'Fij':
                hoja['AE66'] = 'X'
            
            if fila['Pago de seguridad.2'] == 'Si':
                hoja['AG66'] = 'X'
            elif fila['Pago de seguridad.2'] == 'No':
                hoja['AI66'] = 'X'

            hoja['AJ66'] = fila["Procedencia.6"]
            hoja['AK66'] = fila["Residencia.2"]
            hoja['AN66'] = fila["Tiempo trabajado.2"]
            hoja['AO66'] = fila["# Personas núcleo familiar.2"]
            hoja['AQ66'] = fila["Personas a cargo.2"]
            hoja['AS66'] = fila["Lugar de residencia familiar.2"]
        
            actividad10 = fila['Remuneración.2']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU66'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV66'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW66'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX66'] = 'X'

            ##### Persona 4 #####

            if fila["Tipo de mano de obra.3"] == "Familiar":
                hoja['B67'] = 'X'
            elif fila["Tipo de mano de obra.3"] == "Contratado":
                hoja['D67'] = 'X'

            hoja['E67'] = fila["Cargo.4"]

            if fila["Género.3"] == "Masculino":
                hoja['J67'] = 'X'
            elif fila["Género.3"] == "Femenino":
                hoja['H67'] = 'X'

            hoja['K67'] = fila["Edad (años).3"]
            hoja['L67'] = fila["Duración jornada (horas).3"]

            actividad9 = fila['Escolaridad.3']
            if actividad9 == 'Primaria':
                hoja['N67'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q67'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S67'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U67'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W67'] = 'X'

            if fila['Contrato.3'] == 'Tem.':
                hoja['AC67'] = 'X'
            elif fila['Contrato.3'] == 'Fij':
                hoja['AE67'] = 'X'
            
            if fila['Pago de seguridad.3'] == 'Si':
                hoja['AG67'] = 'X'
            elif fila['Pago de seguridad.3'] == 'No':
                hoja['AI67'] = 'X'

            hoja['AJ67'] = fila["Procedencia.7"]
            hoja['AK67'] = fila["Residencia.3"]
            hoja['AN67'] = fila["Tiempo trabajado.3"]
            hoja['AO67'] = fila["# Personas núcleo familiar.3"]
            hoja['AQ67'] = fila["Personas a cargo.3"]
            hoja['AS67'] = fila["Lugar de residencia familiar.3"]
        
            actividad10 = fila['Remuneración.3']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU67'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV67'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW67'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX67'] = 'X'

            ##### Persona 5 #####

            if fila["Tipo de mano de obra.4"] == "Familiar":
                hoja['B68'] = 'X'
            elif fila["Tipo de mano de obra.4"] == "Contratado":
                hoja['D68'] = 'X'

            hoja['E68'] = fila["Cargo.5"]

            if fila["Género.4"] == "Masculino":
                hoja['J68'] = 'X'
            elif fila["Género.4"] == "Femenino":
                hoja['H68'] = 'X'

            hoja['K68'] = fila["Edad (años).4"]
            hoja['L68'] = fila["Duración jornada (horas).4"]

            actividad9 = fila['Escolaridad.4']
            if actividad9 == 'Primaria':
                hoja['N68'] = 'X'
            elif actividad9 == 'Bachillerato':
                hoja['Q68'] = 'X'
            elif actividad9 == 'Técnico o tecnológico':
                hoja['S68'] = 'X'
            elif actividad9 == 'Profesional':
                hoja['U68'] = 'X'
            elif actividad9 == 'Posgrado':
                hoja['W68'] = 'X'

            if fila['Contrato.4'] == 'Tem.':
                hoja['AC68'] = 'X'
            elif fila['Contrato.4'] == 'Fij':
                hoja['AE68'] = 'X'
            
            if fila['Pago de seguridad.4'] == 'Si':
                hoja['AG68'] = 'X'
            elif fila['Pago de seguridad.4'] == 'No':
                hoja['AI68'] = 'X'

            hoja['AJ68'] = fila["Procedencia.8"]
            hoja['AK68'] = fila["Residencia.4"]
            hoja['AN68'] = fila["Tiempo trabajado.4"]
            hoja['AO68'] = fila["# Personas núcleo familiar.4"]
            hoja['AQ68'] = fila["Personas a cargo.4"]
            hoja['AS68'] = fila["Lugar de residencia familiar.4"]
        
            actividad10 = fila['Remuneración.4']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU68'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV68'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW68'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX68'] = 'X'

            ##### Persona 6 #####

            if fila["Tipo de mano de obra.5"] == "Familiar":
                hoja['B69'] = 'X'
            elif fila["Tipo de mano de obra.5"] == "Contratado":
                hoja['D69'] = 'X'

            hoja['E69'] = fila["Cargo.6"]

            if fila["Género.5"] == "Masculino":
                hoja['j69'] = 'X'
            elif fila["Género.5"] == "Femenino":
                hoja['H69'] = 'X'

            hoja['K69'] = fila["Edad (años).5"]
            hoja['L69'] = fila["Duración jornada (horas).5"]

            actividad9 = fila['Escolaridad.5']
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

            if fila['Contrato.5'] == 'Tem.':
                hoja['AC69'] = 'X'
            elif fila['Contrato.5'] == 'Fij':
                hoja['AE69'] = 'X'
            
            if fila['Pago de seguridad.5'] == 'Si':
                hoja['AG69'] = 'X'
            elif fila['Pago de seguridad.5'] == 'No':
                hoja['AI69'] = 'X'

            hoja['AJ69'] = fila["Procedencia.9"]
            hoja['AK69'] = fila["Residencia.5"]
            hoja['AN69'] = fila["Tiempo trabajado.5"]
            hoja['AO69'] = fila["# Personas núcleo familiar.5"]
            hoja['AQ69'] = fila["Personas a cargo.5"]
            hoja['AS69'] = fila["Lugar de residencia familiar.5"]
        
            actividad10 = fila['Remuneración.5']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU69'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV69'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW69'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX69'] = 'X'

            ##### Persona 7 #####

            if fila["Tipo de mano de obra.6"] == "Familiar":
                hoja['B70'] = 'X'
            elif fila["Tipo de mano de obra.6"] == "Contratado":
                hoja['D70'] = 'X'

            hoja['E70'] = fila["Cargo.7"]

            if fila["Género.6"] == "Masculino":
                hoja['J70'] = 'X'
            elif fila["Género.6"] == "Femenino":
                hoja['H70'] = 'X'

            hoja['K70'] = fila["Edad (años).6"]
            hoja['L70'] = fila["Duración jornada (horas).6"]

            actividad9 = fila['Escolaridad.6']
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

            if fila['Contrato.6'] == 'Tem.':
                hoja['AC70'] = 'X'
            elif fila['Contrato.6'] == 'Fij':
                hoja['AE70'] = 'X'
            
            if fila['Pago de seguridad.6'] == 'Si':
                hoja['AG70'] = 'X'
            elif fila['Pago de seguridad.6'] == 'No':
                hoja['AI70'] = 'X'

            hoja['AJ70'] = fila["Procedencia.10"]
            hoja['AK70'] = fila["Residencia.6"]
            hoja['AN70'] = fila["Tiempo trabajado.6"]
            hoja['AO70'] = fila["# Personas núcleo familiar.6"]
            hoja['AQ70'] = fila["Personas a cargo.6"]
            hoja['AS70'] = fila["Lugar de residencia familiar.6"]
        
            actividad10 = fila['Remuneración.6']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU70'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV70'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW70'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX70'] = 'X'

            ##### Persona 8 #####

            if fila["Tipo de mano de obra.7"] == "Familiar":
                hoja['B71'] = 'X'
            elif fila["Tipo de mano de obra.7"] == "Contratado":
                hoja['D71'] = 'X'

            hoja['E71'] = fila["Cargo.8"]

            if fila["Género.7"] == "Masculino":
                hoja['J71'] = 'X'
            elif fila["Género.7"] == "Femenino":
                hoja['H71'] = 'X'

            hoja['K71'] = fila["Edad (años).7"]
            hoja['L71'] = fila["Duración jornada (horas).7"]

            actividad9 = fila['Escolaridad.7']
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

            if fila['Contrato.7'] == 'Tem.':
                hoja['AC71'] = 'X'
            elif fila['Contrato.7'] == 'Fij':
                hoja['AE71'] = 'X'
            
            if fila['Pago de seguridad.7'] == 'Si':
                hoja['AG71'] = 'X'
            elif fila['Pago de seguridad.7'] == 'No':
                hoja['AI71'] = 'X'

            hoja['AJ71'] = fila["Procedencia.11"]
            hoja['AK71'] = fila["Residencia.7"]
            hoja['AN71'] = fila["Tiempo trabajado.7"]
            hoja['AO71'] = fila["# Personas núcleo familiar.7"]
            hoja['AQ71'] = fila["Personas a cargo.7"]
            hoja['AS71'] = fila["Lugar de residencia familiar.7"]
        
            actividad10 = fila['Remuneración.7']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU71'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV71'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW71'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX71'] = 'X'

            ##### Persona 9 #####

            if fila["Tipo de mano de obra.8"] == "Familiar":
                hoja['B72'] = 'X'
            elif fila["Tipo de mano de obra.8"] == "Contratado":
                hoja['D72'] = 'X'

            hoja['E72'] = fila["Cargo.9"]

            if fila["Género.8"] == "Masculino":
                hoja['J72'] = 'X'
            elif fila["Género.8"] == "Femenino":
                hoja['H72'] = 'X'

            hoja['K72'] = fila["Edad (años).8"]
            hoja['L72'] = fila["Duración jornada (horas).8"]

            actividad9 = fila['Escolaridad.8']
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

            if fila['Contrato.8'] == 'Tem.':
                hoja['AC72'] = 'X'
            elif fila['Contrato.8'] == 'Fij':
                hoja['AE72'] = 'X'
            
            if fila['Pago de seguridad.8'] == 'Si':
                hoja['AG72'] = 'X'
            elif fila['Pago de seguridad.8'] == 'No':
                hoja['AI72'] = 'X'

            hoja['AJ72'] = fila["Procedencia.12"]
            hoja['AK72'] = fila["Residencia.8"]
            hoja['AN72'] = fila["Tiempo trabajado.8"]
            hoja['AO72'] = fila["# Personas núcleo familiar.8"]
            hoja['AQ72'] = fila["Personas a cargo.8"]
            hoja['AS72'] = fila["Lugar de residencia familiar.8"]
        
            actividad10 = fila['Remuneración.8']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU72'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV72'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW72'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX72'] = 'X'

            ##### Persona 10 #####

            if fila["Tipo de mano de obra.9"] == "Familiar":
                hoja['B73'] = 'X'
            elif fila["Tipo de mano de obra.9"] == "Contratado":
                hoja['D73'] = 'X'

            hoja['E73'] = fila["Cargo.9"]

            if fila["Género.9"] == "Masculino":
                hoja['J73'] = 'X'
            elif fila["Género.9"] == "Femenino":
                hoja['H73'] = 'X'

            hoja['K73'] = fila["Edad (años).9"]
            hoja['L73'] = fila["Duración jornada (horas).9"]

            actividad9 = fila['Escolaridad.9']
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

            if fila['Contrato.9'] == 'Tem.':
                hoja['AC73'] = 'X'
            elif fila['Contrato.9'] == 'Fij':
                hoja['AE73'] = 'X'
            
            if fila['Pago de seguridad.9'] == 'Si':
                hoja['AG73'] = 'X'
            elif fila['Pago de seguridad.9'] == 'No':
                hoja['AI73'] = 'X'

            hoja['AJ73'] = fila["Procedencia.13"]
            hoja['AK73'] = fila["Residencia.9"]
            hoja['AN73'] = fila["Tiempo trabajado.9"]
            hoja['AO73'] = fila["# Personas núcleo familiar.9"]
            hoja['AQ73'] = fila["Personas a cargo.9"]
            hoja['AS73'] = fila["Lugar de residencia familiar.9"]
        
            actividad10 = fila['Remuneración.9']
            if actividad10 == 'Inferiores a $900.000':
                hoja['AU73'] = 'X'
            elif actividad10 == '$901.000- a $1.800.000':
                hoja['AV73'] = 'X'
            elif actividad10 == '$1.801.000 - $2.700.000':
                hoja['AW73'] = 'X'
            elif actividad10 == 'Superiores s a $2.701.000':
                hoja['AX73'] = 'X'


        elif fila["Emplea directamente algún tipo de mano de obra (si la respuesta es SI, diligenciar el título G)"] == 'No':
            hoja['AP40'] = 'X'

        ##### INFORMACIÓN TRANSPORTE DE PASAJEROS #####

        hoja['N47'] = fila['Tarifa única ($)']
        hoja['V47'] = fila['Servicio contratado por día ($)']
        hoja['AE47'] = fila['Servicio contratado por semana ($)']
        hoja['AL47'] = fila['Servicio contratado por mes ($)']
        hoja['AT47'] = fila['Servicio contratado por km recorrido ($)']

        hoja['M48'] = fila['Tarifa única ($).1']
        hoja['U48'] = fila['Servicio contratado por día ($).1']
        hoja['AD48'] = fila['Servicio contratado por semana ($).1']
        hoja['AK48'] = fila['Servicio contratado por mes ($).1']
        hoja['AS48'] = fila['Servicio contratado por km recorrido ($).1']

        hoja['M50'] = fila['Tarifa única ($).2']
        hoja['U50'] = fila['Servicio contratado por día ($).2']
        hoja['AD50'] = fila['Servicio contratado por semana ($).2']
        hoja['AK50'] = fila['Servicio contratado por mes ($).2']
        hoja['AS50'] = fila['Servicio contratado por km recorrido ($).2']

        hoja['AC52'] = fila['¿Cuál es el destino más frecuente?']


        ##### INFORMACIÓN TRANSPORTE DE PASAJEROS 2 #####

        hoja['A56'] = fila['Elemento transportado']
        hoja['I56'] = fila['Cantidad mensual']
        hoja['M56'] = fila['Procedencia']
        hoja['T56'] = fila['Destino']
        hoja['AA56'] = fila['Frecuencia de movilización']
        hoja['AG56'] = fila['Costo promedio del flete o tarifa']
        hoja['AM56'] = fila['Promedio mensual e ingreso']
        hoja['AT56'] = fila['Medio de transporte']

        hoja['A57'] = fila['Elemento transportado.1']
        hoja['I57'] = fila['Cantidad mensual.1']
        hoja['M57'] = fila['Procedencia.1']
        hoja['T57'] = fila['Destino.1']
        hoja['AA57'] = fila['Frecuencia de movilización.1']
        hoja['AG57'] = fila['Costo promedio del flete o tarifa.1']
        hoja['AM57'] = fila['Promedio mensual e ingreso.1']
        hoja['AT57'] = fila['Medio de transporte.1']

        hoja['A58'] = fila['Elemento transportado.2']
        hoja['I58'] = fila['Cantidad mensual.2']
        hoja['M58'] = fila['Procedencia.2']
        hoja['T58'] = fila['Destino.2']
        hoja['AA58'] = fila['Frecuencia de movilización.2']
        hoja['AG58'] = fila['Costo promedio del flete o tarifa.2']
        hoja['AM58'] = fila['Promedio mensual e ingreso.2']
        hoja['AT58'] = fila['Medio de transporte.2']
    
        hoja['A59'] = fila['Elemento transportado.3']
        hoja['I59'] = fila['Cantidad mensual.3']
        hoja['M59'] = fila['Procedencia.3']
        hoja['T59'] = fila['Destino.3']
        hoja['AA59'] = fila['Frecuencia de movilización.3']
        hoja['AG59'] = fila['Costo promedio del flete o tarifa.3']
        hoja['AM59'] = fila['Promedio mensual e ingreso.3']
        hoja['AT59'] = fila['Medio de transporte.3']


