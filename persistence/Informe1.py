import pandas as pd
import re
from openpyxl import load_workbook
from datetime import datetime

plantilla = r'censos\FORMATO 1 IDENTIFICACIÓN - Aprobado.xlsx'
wb = load_workbook(plantilla)
ws = wb.active

datos = pd.read_excel(r'censos\Encuesta 1 Identificación.xlsx')
mapeo_columnas = {
    'data-num_encuesta': 'Encuesta No.',
    'data-fecha': 'Fecha(DD/MM/AAAA)',
    'data-encuestador': 'Encuestador',
    'data-departamento': 'Departamento',
    'data-municipio': 'Municipio',
    'data-vereda_centro_poblado': 'Vereda/Centro Poblado',
    'data-permite_entrevista': 'Permite Entrevista',
    'data-coordenadas': 'Coordenada Norte',
    'data-coordenadas-altitude': 'Coordenada Este',
    'data-nombre_establecimiento': 'Nombre del establecimiento',
    'data-direccion': 'Dirección',
    'data-telefono_contacto': 'Teléfono de contacto',
    'data-actividad_economica': 'Actividad económica principal',
    'data-inicio_actividad': '¿En qué año inició la actividad?',
    'data-propietario': 'Propietario',
    'data-procedencia_propietario': 'Procedencia',
    'data-lugar_residencia': 'Lugar de Residencia',
    'data-administrador': 'Administrador',
    'data-procedencia_administrador': 'Procedencia.1',
    'data-lugar_residencia_admin': 'Lugar de Residencia.1',
    'data-actividad_tipo': 'Este establecimiento desarrolla su actividad como',
    'data-tipo_actividad': 'Tipo de actividad',
    'data-producto_principal': '¿Cuál es el principal producto o servicio que oferta?',
    'data-tenencia_propiedad': 'Tenencia de la propiedad',
    'data-tenencia_propiedad_other': '¿Cuál?.1',
    'data-canon_arrendamiento': 'Canon de arrendamiento',
    'data-actividad_ingresos': '¿De que actividad proviene la mayor parte de ingresos obtenidos en la unidad económica?',
    'data-frecuencia_ingresos': 'Frecuencia con la que recibe ingresos por actividad',
    'data-ingresos': '¿Cuál es la cantidad de ingresos recibidos por la actividad?',
    'data-ingresos_other': '¿Cuál?.3',
    'data-horario_actividad': '¿En qué horario desempeña la actividad?',
    'data-tiene_registro': '¿Tiene registro de cámara y comercio?',
    'data-lugares_comercializa': '¿En qué lugares comercializa?',
    'data-lugares_comercializa_other': '¿Dónde?',
    'data-compra_vereda': '¿Compra productos o insumos en la vereda?',
    'data-comercializa_otra_vereda': '¿Comercializa productos o insumos en otras veredas?',
    'data-donde_comercializa': '¿En qué lugares?',
    'data-estrato': 'Estrato',
    'data-servicios_publicos': '¿Cuánto pagó el último mes por concepto de servicios públicos?'
}

    # Renombrar las columnas en la fila para usar los nombres esperados en la plantilla
datos = datos.rename(columns=mapeo_columnas)
def crearArchivoPrimero(ws, df_fila):
    # Diccionario de mapeo de columnas ODK -> Plantilla
    # Llenar campos básicos
    ws['Y1'] = df_fila['Encuesta No.']
    if pd.notna(df_fila['Fecha(DD/MM/AAAA)']):
        fecha_valor = df_fila['Fecha(DD/MM/AAAA)']
        
        # Extraer día, mes y año
        ws['X2'] = fecha_valor.day    # Día
        ws['Z2'] = fecha_valor.month  # Mes
        ws['AD2'] = fecha_valor.year   # Año
    else:
        print("Campo de fecha vacío.")


    ws['W3'] = df_fila['Encuestador']
    ws['A8'] = df_fila['Departamento']
    ws['N8'] = df_fila['Municipio']
    ws['U8'] = df_fila['Vereda/Centro Poblado']
    ws['A10'] = f"{df_fila['Coordenada Norte']}, {df_fila['Coordenada Este']}"

    # Campo "Permite entrevista"
    if df_fila['Permite Entrevista'] == 'yes':
        ws['W10'] = 'X'

        # Sección B: Identificación
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

        # Sección C: Descripción
        actividad = df_fila['Este establecimiento desarrolla su actividad como']
        mapeo_actividad = {
            'natural': 'G23',
            'sociedad_hecho': 'N23',
            'unipersonal': 'G24',
            'sociedad_comercial': 'N24',
            'cooperativa': 'G25',
            'predio': 'G26',
            'other': 'N25'
        }
        if actividad in mapeo_actividad:
            ws[mapeo_actividad[actividad]] = 'X'
            if actividad == 'other':
                ws['K26'] = df_fila['¿Cuál?.1']

        # Tipo de actividad
        actividad2 = df_fila['Tipo de actividad']
        mapeo_tipo_actividad = {
            'agricola': 'G30',
            'pecuaria': 'G31',
            'agroindustrial': 'G32',
            'servicios': 'G33',
            'comercial': 'M31',
            'manufactura': 'M32',
            'transporte': 'M33'
        }
        if actividad2 in mapeo_tipo_actividad:
            ws[mapeo_tipo_actividad[actividad2]] = 'X'

        # Tenencia de la propiedad
        tenencia = df_fila['Tenencia de la propiedad']
        mapeo_tenencia = {
            'Propia': 'E40',
            'Administrada': 'E41',
            'Arrendada*  Responder 16': 'E42',
            'Otra': 'E44'
        }
        if tenencia in mapeo_tenencia:
            ws[mapeo_tenencia[tenencia]] = 'X'
            if tenencia == 'Arrendada*  Responder 16':
                ws['J41'] = df_fila['Canon de arrendamiento']
            if tenencia == 'Otra':
                ws['C45'] = df_fila['¿Cuál?.1']

        # Frecuencia de ingresos
        frecuencia = df_fila['Frecuencia con la que recibe ingresos por actividad']
        mapeo_frecuencia = {
            'Diario': 'E53',
            'Semanal': 'E54',
            'Quincenal': 'E55',
            'Mensual': 'M53',
            'Semestral': 'M54',
            'Otra': 'M55'
        }
        if frecuencia in mapeo_frecuencia:
            ws[mapeo_frecuencia[frecuencia]] = 'X'
            if frecuencia == 'Otra':
                ws['K56'] = df_fila['¿Cuál?.2']

        # Estrato
        estrato = df_fila['Estrato']
        mapeo_estrato = {
            1: 'R59',
            2: 'T59',
            3: 'V59'
        }
        if estrato in mapeo_estrato:
            ws[mapeo_estrato[estrato]] = 'X'

        ws['Y59'] = df_fila['¿Cuánto pagó el último mes por concepto de servicios públicos?']

    elif df_fila['Permite Entrevista'] == 'no':
        ws['Z10'] = 'X'
    

if __name__ == '__main__':
    for index, fila in datos.iterrows():
        print(f"Llenando datos para la fila {index + 1}...")
        crearArchivoPrimero(ws, fila)  # Llamar a la función para llenar la plantilla

        # Guardar la plantilla con un nombre único para cada registro
        archivo_salida = f"plantilla_ejemlo_{index + 1}.xlsx"
        wb.save(archivo_salida)
        print(f"Archivo guardado: {archivo_salida}")