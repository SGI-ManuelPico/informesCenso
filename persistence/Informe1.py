import pandas as pd
import re
from openpyxl import load_workbook
from datetime import datetime


def llenarInforme1(ws, df_fila):
    ws['Y1'] = df_fila['data-num_encuesta']
    if pd.notna(df_fila['data-fecha']):  # Verificar que no sea NaT
        fecha_valor = df_fila['data-fecha']
        ws['X2'] = str(fecha_valor.day)
        ws['Z2'] = str(fecha_valor.month)
        ws['AD2'] = str(fecha_valor.year)
    else:
        print("Campo de fecha vacío o inválido.")

    ws['W3'] = df_fila['data-encuestador']
    ws['A8'] = df_fila['data-departamento']
    ws['N8'] = df_fila['data-municipio']
    ws['U8'] = df_fila['data-vereda_centro_poblado']
    ws['A10'] = f"{df_fila['data-coordenadas']}, {df_fila['data-coordenadas-altitude']}"

    if df_fila['data-permite_entrevista'] == 'yes':
        ws['W10'] = 'X'

        ws['G14'] = df_fila['data-nombre_establecimiento']
        ws['D15'] = df_fila['data-direccion']
        ws['U15'] = df_fila['data-telefono_contacto']
        ws['G16'] = df_fila['data-actividad_economica']
        ws['W16'] = df_fila['data-inicio_actividad']
        ws['D17'] = df_fila['data-propietario']
        ws['Q17'] = df_fila['data-procedencia_propietario']
        ws['Y17'] = df_fila['data-lugar_residencia']
        ws['D18'] = df_fila['data-administrador']
        ws['Q18'] = df_fila['data-procedencia_administrador']
        ws['Z18'] = df_fila['data-lugar_residencia_admin']

        actividad = df_fila['data-actividad_tipo']
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
                ws['K26'] = df_fila['data-tenencia_propiedad_other']

        actividad2 = df_fila['data-tipo_actividad']
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

        ws['B37'] = df_fila['data-producto_principal']

        tenencia = df_fila['data-tenencia_propiedad']
        mapeo_tenencia = {
            'propia': 'E40',
            'administrada': 'E41',
            'arrendada': 'E42',
            'other': 'E44'
        }
        if tenencia in mapeo_tenencia:
            ws[mapeo_tenencia[tenencia]] = 'X'
            if tenencia == 'arrendada':
                ws['J41'] = df_fila['data-canon_arrendamiento']
            if tenencia == 'other':
                ws['C45'] = df_fila['data-tenencia_propiedad_other']

        ws['A48'] = df_fila['data-actividad_ingresos']

        ingresos = df_fila['data-ingresos']
        mapeo_ingresos = {
            'inferior_600k': 'Y23',
            'entre_601k_1500k' : 'Y24',
            'entre_1501k_3000k' : 'Y25',
            'superior_3000k' : 'Y26',
            'other' : 'AD24'
        }
        if ingresos in mapeo_ingresos:
            ws[mapeo_ingresos[ingresos]] = 'X'
            if ingresos == 'other':
                ws['AA25'] = df_fila['data-ingresos_other']
        
        ws['P30'] = df_fila['data-horario_actividad']

        if df_fila['data-tiene_registro'] == 'yes':
            ws['T33'] = 'X'
        else:
            ws['W33'] = 'X'

        lugares_comercializa = df_fila['data-lugares_comercializa']
        mapeo_lugares_comercializa = {
            'sitio': 'X39',
            'empresa': 'X40',
            'mercado': 'X41',
            'acopio': 'X43',
            'other': 'X44'
        }
        if lugares_comercializa in mapeo_lugares_comercializa:
            ws[mapeo_lugares_comercializa[lugares_comercializa]] = 'X'
            if lugares_comercializa == 'other':
                ws['K50'] = df_fila['data-lugares_comercializa_other']


        frecuencia = df_fila['data-frecuencia_ingresos']
        mapeo_frecuencia = {
            'diario': 'E53',
            'semanal': 'E54',
            'quincenal': 'E55',
            'mensual': 'M53',
            'semestral': 'M54',
            'other': 'M55'
        }
        if frecuencia in mapeo_frecuencia:
            ws[mapeo_frecuencia[frecuencia]] = 'X'
            if frecuencia == 'other':
                ws['K56'] = df_fila['data-ingresos_other']


        if df_fila['data-compra_vereda'] == 'yes':
            ws['T48'] = 'X'
        else:
            ws['W48'] = 'X'


        if df_fila['data-comercializa_otra_vereda'] == 'yes':
            ws['U53'] = 'X'
            ws['U55'] = df_fila['data-donde_comercializa']
        else:
            ws['Y53'] = 'X'

        estrato = int(df_fila['data-estrato'])
        mapeo_estrato = {
            1: 'R59',
            2: 'T59',
            3: 'V59'
        }
        if estrato in mapeo_estrato:
            ws[mapeo_estrato[estrato]] = 'X'

        ws['Y59'] = df_fila['data-servicios_publicos']

    elif df_fila['data-permite_entrevista'] == 'no':
        ws['Z10'] = 'X'
